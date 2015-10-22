var _ = require ("underscore");
var Backbone = require ("backbone");
var JSZip = require ("jszip");
var xml2js = require ("xml2js");
var VError = require ("verror");
var fs = require ("fs");
var async = require ("async");
var XLSXChart = Backbone.Model.extend ({
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
			return abc [n / 26 | 0] + abc [n % 26];
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
		Write chart
	*/
	writeChart: function (cb) {
		var me = this;
		me.read ({file: "xl/charts/chart1.xml"}, function (err, o) {
			if (err) {
				return cb (new VError (err, "writeChart"));
			}
			var ser = [];
			_.each (me.titles, function (t, i) {
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
				if (me.chart == "scatter") {
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
				}
				ser.push (r);
			});
			var tag = me.chart == "column" ? "bar" : me.chart;
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + tag + "Chart"]["c:ser"] = ser;
			me.write ({file: "xl/charts/chart1.xml", object: o});
			cb ();
		});
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
				fs.readFile (__dirname + "/template/" + me.chart + ".xlsx", function (err, data) {
					if (err) {
						return cb (err);
					}
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
	},
	writeFile: function (opts, cb) {
		var me = this;
		me.generate (_.defaults ({type: "base64"}, opts), function (err, result) {
			if (err) {
				return cb (new VError (err, "writeFile"));
			}
			fs.writeFile (opts.file, result, "base64", function (err) {
				cb (err ? new VError (err, "writeFile") : null);
			});
		});
	}
});
module.exports = XLSXChart;
