var _ = require ("underscore"); 
var Backbone = require ("backbone");
var JSZip = require ("jszip");
var xml2js = require ("xml2js");
var VError = require ("verror");
var fs = require ("fs");
var async = require ("async");
var Chart = Backbone.Model.extend ({
	/*
		Read XML file from xlsx as object
	*/
	read: function (opts, cb) {
		var me = this;
		var t = me.zip.file (opts.file).asText ();
		var parser = new xml2js.Parser ({explicitArray: false});
		parser.parseString (t, function (err, o) {
			if (err) {
				return new VError (err, "getXML");
			}
			cb (err, o);
	    });
	},
	/*
		Build XML from object and write to zip
	*/
	write: function (opts) {
		var me = this;
		var builder = new xml2js.Builder ();
		var xml = builder.buildObject (opts.object);
		me.zip.file (opts.file, new Buffer (xml), {base64: true});
	},
	/*
		Get column name
	*/
	getColName: function (n) {
		var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		n --;
		if (n < 26) {
			return abc [n];
		} else {
			return abc [(n / 26 - 1) | 0] + abc [n % 26];
		}
	},
	/*
		Get shared string index
	*/
	getStr: function (s) {
		var me = this;
		if (!me.str.hasOwnProperty (s)) {
			throw new VError ("getStr: Unknown string: " + s);
		}
		return me.str [s];
	},
	/*
		Write table
	*/
	writeTable: function (cb) {
		var me = this;
		me.read ({file: "xl/worksheets/sheet2.xml"}, function (err, o) {
			if (err) {
				return cb (new VError (err, "writeTable"));
			}
			o.worksheet.dimension.$.ref = "A1:" + me.getColName (me.titles.length + 1) + (me.fields.length + 1);
			var rows = [{
				$: {
					r: 1,
					spans: "1:" + (me.titles.length + 1)
				},
				c: _.map (me.titles, function (t, x) {
					return {
						$: {
							r: me.getColName (x + 2) + 1,
							t: "s"
						},
						v: me.getStr (t)
					}
				})
			}];
			_.each (me.fields, function (f, y) {
				var r = {
					$: {
						r: y + 2,
						spans: "1:" + (me.titles.length + 1)
					}
				};
				var c = [{
					$: {
						r: "A" + (y + 2),
						t: "s"
					},
					v: me.getStr (f)
				}];
				_.each (me.titles, function (t, x) {
					c.push ({
						$: {
							r: me.getColName (x + 2) + (y + 2)
						},
						v: me.data [t][f]
					});
				});
				r.c = c;
				rows.push (r);
			});
			o.worksheet.sheetData.row = rows;
			me.write ({file: "xl/worksheets/sheet2.xml", object: o});
			cb ();
		});
	},
	/*
		Write strings
	*/
	writeStrings: function (cb) {
		var me = this;
		me.read ({file: "xl/sharedStrings.xml"}, function (err, o) {
			if (err) {
				return cb (new VError (err, "writeStrings"));
			}
			o.sst.$.count = me.titles.length + me.fields.length;
			o.sst.$.uniqueCount = o.sst.$.count;
			var si = [];
			_.each (me.titles, function (t) {
				si.push ({t: t});
			});
			_.each (me.fields, function (t) {
				si.push ({t: t});
			});
			me.str = {};
			_.each (si, function (o, i) {
				me.str [o.t] = i;
			});
			o.sst.si = si;
			me.write ({file: "xl/sharedStrings.xml", object: o});
			cb ();
		});
	},
	/*
		Remove unused charts
	*/
	removeUnusedCharts: function (o) {
		var me = this;
		if (me.tplName != "charts") {
			return;
		};
		var axId = [];
		function addId (o) {
			_.each (o ["c:axId"], function (o) {
				axId.push (o.$.val);
			});
		};
		_.each (["line", "radar", "area", "scatter", "pie"], function (chart) {
			if (!me.charts [chart]) {
				delete o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + chart + "Chart"];
			} else {
				addId (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + chart + "Chart"]);
			};
		});
		if (!me.charts ["column"] && !me.charts ["bar"]) {
			delete o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"];
		} else
		if (me.charts ["column"] && !me.charts ["bar"]) {
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"] = o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][0];
			addId (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]);
		} else
		if (!me.charts ["column"] && me.charts ["bar"]) {
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"] = o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][1];
			addId (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]);
		} else {
			addId (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][0]);
			addId (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][1]);
		};

		var catAx = [];
		_.each (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"], function (o) {
			if (axId.indexOf (o ["c:axId"].$.val) > -1) {
				catAx.push (o);
			};
		});
		o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"] = catAx;

		var valAx = [];
		_.each (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"], function (o) {
			if (axId.indexOf (o ["c:axId"].$.val) > -1) {
				valAx.push (o);
			};
		});
		o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"] = valAx;
	},
	/*
		Write chart
	*/
	writeChart: function (cb) {
		var me = this;
		var chart;
		me.read ({file: "xl/charts/chart1.xml"}, function (err, o) {
			if (err) {
				return cb (new VError (err, "writeChart"));
			}
			var ser = {};
			_.each (me.titles, function (t, i) {
				var chart = me.data [t].chart || me.chart;
				var r = {
					"c:idx": {
						$: {
							val: i
						}
					},
					"c:order": {
						$: {
							val: i
						}
					},
					"c:tx": {
						"c:strRef": {
							"c:f": "Table!$" + me.getColName (i + 2) + "$1",
							"c:strCache": {
								"c:ptCount": {
									$: {
										val: 1
									}
								},
								"c:pt": {
									$: {
										idx: 0
									},
									"c:v": t
								}
							}
						}
					},
					"c:cat": {
						"c:strRef": {
							"c:f": "Table!$A$2:$A$" + (me.fields.length + 1),
							"c:strCache": {
								"c:ptCount": {
									$: {
										val: me.fields.length
									}
								},
								"c:pt": _.map (me.fields, function (f, j) {
									return {
										$: {
											idx: j
										},
										"c:v": f
									};
								})
							}
						}
					},
					"c:val": {
						"c:numRef": {
							"c:f": "Table!$" + me.getColName (i + 2) + "$2:$" + me.getColName (i + 2) + "$" + (me.fields.length + 1),
							"c:numCache": {
								"c:formatCode": "General",
								"c:ptCount": {
									$: {
										val: me.fields.length
									}
								},
								"c:pt": _.map (me.fields, function (f, j) {
									return {
										$: {
											idx: j
										},
										"c:v": me.data [t][f]
									};
								})
							}
						}
					}
				};
				if (chart == "scatter") {
					r ["c:xVal"] = r ["c:cat"];
					delete r ["c:cat"];
					r ["c:yVal"] = r ["c:val"];
					delete r ["c:val"];
					r ["c:spPr"] = {
						"a:ln": {
							$: {
								w: 28575
							},
							"a:noFill": ""
						}
					};
				};
				ser [chart] = ser [chart] || [];
				ser [chart].push (r);
			});
/*
			var tag = chart == "column" ? "bar" : chart;
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][0]["c:ser"] = ser;
*/
			_.each (ser, function (ser, chart) {
				if (chart == "column") {
					if (me.tplName == "charts") {
						o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][0]["c:ser"] = ser;
					} else {
						o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:ser"] = ser;
					};
				} else
				if (chart == "bar") {
					if (me.tplName == "charts") {
						o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][1]["c:ser"] = ser;
					} else {
						o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:ser"] = ser;
					};
				} else {
					o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + chart + "Chart"]["c:ser"] = ser;
				};
			});
			me.removeUnusedCharts (o);
/*
			if (me.showVal) {
				o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + tag + "Chart"]["c:dLbls"]["c:showVal"] = {
					$: {
						val: "1"
					}
				};
			};
*/
			if (me.chartTitle) {
				me.writeTitle (o, me.chartTitle);
			};
/*
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"] = o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][0];
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"] = o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"][0];
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"] = o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"][0];
			delete o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:lineChart"];
			delete o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:areaChart"];
*/
			me.write ({file: "xl/charts/chart1.xml", object: o});
			cb ();
		});
	},
	/*
		Chart title
	*/
	writeTitle: function (chart, title) {
		var me = this;
		chart ["c:chartSpace"]["c:chart"]["c:title"] = {
			"c:tx": {
				"c:rich": {
					"a:bodyPr": {},
					"a:lstStyle": {},
					"a:p": {
						"a:pPr": {
							"a:defRPr": {}
						},
						"a:r": {
							"a:rPr": {
								$: {
									lang: "ru-RU"
								}
							},
							"a:t": title
						}
					}
				}
			},
			"c:layout": {},
			"c:overlay": {
				$: {
					val: "0"
				}
			}
		};
		chart ["c:chartSpace"]["c:chart"]["c:autoTitleDeleted"] = {
			$: {
				val: "0"
			}
		};
	},
	/*
		Set template name
	*/
	setTemplateName: function () {
		var me = this;
		var charts = {};
		_.each (me.data, function (o) {
			charts [o.chart || me.chart] = true;
		});
		me.charts = charts;
		if (charts ["radar"]) {
			me.tplName = "radar";
			return;
		};
		if (charts ["scatter"]) {
			me.tplName = "scatter";
			return;
		};
		if (charts ["pie"]) {
			me.tplName = "pie";
			return;
		};
		if (_.keys (charts).length == 1) {
			me.tplName = _.keys (charts) [0];
			return;
		};
		me.tplName = "charts";
	},
	/*
		Generate XLSX with chart
		chart: column, bar, line, radar, area, scatter, pie
		titles: []
		fields: []
		data: {title: {field: value, ...}, ...}
	*/
	generate: function (opts, cb) {
		var me = this;
		opts.type = opts.type || "nodebuffer";
		_.extend (me, opts);
		async.series ([
			function (cb) {
				me.zip = new JSZip ();
				me.setTemplateName ();
				fs.readFile (__dirname + "/../template/" + me.tplName + ".xlsx", function (err, data) {
					if (err) {
						return cb (err);
					};
					me.zip.load (data);
					cb ();
				});
			},
			function (cb) {
				me.writeStrings (cb);
			},
			function (cb) {
				_.each (me.titles, function (t) {
					me.data [t] = me.data [t] || {};
					_.each (me.fields, function (f) {
						me.data [t][f] = me.data [t][f] || 0;
					});
				});
				me.writeTable (cb);
			},
			function (cb) {
				me.writeChart (cb);
			}
		], function (err) {
			if (err) {
				return cb (new VError (err, "build"));
			}
			var result = me.zip.generate ({type: me.type});
			cb (null, result);
		});
	}
});
module.exports = Chart;
