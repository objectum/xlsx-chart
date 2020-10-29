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
		Write mult table
	*/
	writeMultTable: function (row, cb) {
		let me = this;
		
		me.read ({file: "xl/worksheets/sheet2.xml"}, function (err, o) {
			if (err) {
				return cb (new VError (err, "writeMultTable"));
			}
			o.worksheet.dimension.$.ref = `A${row}:${me.getColName (me.titles.length + 1)}${me.fields.length + row + 1}`;
			
			let rows = [{
				$: {
					r: row,
					spans: "1:2"
				},
				c: {
					$: {
						r: `A${row}`,
						t: "s"
					},
					v: me.getStr (me.chartTitle)
				}
			}, {
				$: {
					r: row + 1,
					spans: "1:" + (me.titles.length + 1)
				},
				c: _.map (me.titles, function (t, x) {
					return {
						$: {
							r: `${me.getColName (x + 2)}${row + 1}`,
							t: "s"
						},
						v: me.getStr (t)
					}
				})
			}];
			_.each (me.fields, function (f, y) {
				let r = {
					$: {
						r: y + 2 + row,
						spans: "1:" + (me.titles.length + 1)
					}
				};
				let c = [{
					$: {
						r: "A" + (y + 2 + row),
						t: "s"
					},
					v: me.getStr (f)
				}];
				_.each (me.titles, function (t, x) {
					c.push ({
						$: {
							r: me.getColName (x + 2) + (y + 2 + row)
						},
						v: me.data [t][f]
					});
				});
				r.c = c;
				rows.push (r);
			});
			if (row == 1) {
				o.worksheet.sheetData.row = rows;
			} else {
				o.worksheet.sheetData.row = [...o.worksheet.sheetData.row, ...rows];
			}
			me.write ({file: "xl/worksheets/sheet2.xml", object: o});
			cb ();
		});
	},
	/*
		Write strings
	*/
	writeStrings: function (cb) {
		let me = this;
		
		me.read ({file: "xl/sharedStrings.xml"}, function (err, o) {
			if (err) {
				return cb (new VError (err, "writeStrings"));
			}
			o.sst.$.count = me.titles.length + me.fields.length;
			o.sst.$.uniqueCount = o.sst.$.count;
			
			let si = [];
			
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
/*
	writeChart: function (chartN, cb) {
		var me = this;

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

			if (me.chartTitle) {
				me.writeTitle (o, me.chartTitle);
			};
			me.write ({file: `xl/charts/chart${chartN}.xml`, object: o});
			cb ();
		});
	},
*/
	writeChart: function (chartN, row, cb) {
		var me = this;
		
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
							"c:f": "Table!$" + me.getColName (i + 2) + "$" + row,
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
							"c:f": "Table!$A$" + (row + 1) + ":$A$" + (me.fields.length + row),
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
							"c:f": "Table!$" + me.getColName (i + 2) + "$" + (row + 1) + ":$" + me.getColName (i + 2) + "$" + (me.fields.length + row),
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
			
			if (me.chartTitle) {
				me.writeTitle (o, me.chartTitle);
			};
			me.write ({file: `xl/charts/chart${chartN}.xml`, object: o});
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
				let path = me.templatePath ? me.templatePath : (__dirname + "/../template/" + me.tplName + ".xlsx");
				fs.readFile(path, function (err, data) {
					if (err) {
						console.error(`Template ${path} not read: ${err}`);
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
						me.data [t][f] = me.data [t][f] || (me.deleteEmptyCells ? '' : 0); //deleteEmptyCells - don't display missing values as 0
					});
				});
				me.writeTable (cb);
			},
			function (cb) {
				me.writeChart (1, 1, cb);
			}
		], function (err) {
			if (err) {
				return cb (new VError (err, "build"));
			}
			var result = me.zip.generate ({type: me.type});
			cb (null, result);
		});
	},
	generateMult: function (opts, cb) {
		let me = this;
		
		opts.type = opts.type || "nodebuffer";
		_.extend (me, opts);
		
		async.series ([
			function (cb) {
				me.zip = new JSZip ();
				
				let path = me.templatePath || (__dirname + "/../template/mult.xlsx");
				
				fs.readFile (path, function (err, data) {
					if (err) {
						console.error(`Template ${path} not read: ${err}`);
						return cb (err);
					}
					me.zip.load (data);
					cb ();
				});
			},
			function (cb) {
				me.titles = [];
				me.fields = [];
				
				me.charts.forEach (chart => {
					me.titles = [...me.titles, ...chart.titles];
					me.fields = [...me.fields, ...chart.fields];
					me.titles.push (chart.chartTitle || "");
				});
				me.writeStrings (cb);
			},
			function (cb) {
				let row = 1;
				
				async.eachSeries (me.charts, (chart, cb) => {
					["chart", "titles", "fields", "data", "chartTitle"].forEach (a => me [a] = chart [a]);
					
					_.each (me.titles, function (t) {
						me.data [t] = me.data [t] || {};
						
						_.each (me.fields, function (f) {
							me.data [t][f] = me.data [t][f] || 0;
						});
					});
					me.writeMultTable (row, () => {
						row += 3 + me.fields.length;
						cb ();
					});
				}, cb)
			},
			function (cb) {
				let n = 0;
				let row = 2;
				
				async.eachSeries (me.charts, (chart, cb) => {
					["chart", "titles", "fields", "data", "chartTitle"].forEach (a => me [a] = chart [a]);
					
					async.series ([
						function (cb) {
							me.writeChart (++ n, row, cb);
						},
						function (cb) {
							row += 3 + me.fields.length;
							
							if (n == 1) {
								return cb ();
							}
							me.read ({file: "[Content_Types].xml"}, function (err, o) {
								if (err) {
									return cb (new VError (err, "generateMult"));
								}
								o ["Types"]["Override"].push ({
									"$": {
										"ContentType": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
										"PartName": `/xl/charts/chart${n}.xml`
									}
								});
								me.write ({file: `[Content_Types].xml`, object: o});
								cb ();
							});
						},
						function (cb) {
							if (n == 1) {
								return cb ();
							}
							me.read ({file: "xl/drawings/_rels/drawing1.xml.rels"}, function (err, o) {
								if (err) {
									return cb (new VError (err, "generateMult"));
								}
								if (n == 2) {
									o ["Relationships"]["Relationship"] = [o ["Relationships"]["Relationship"]];
								}
								o ["Relationships"]["Relationship"].push ({
									"$": {
										"Id": `rId${n}`,
										"Target": `../charts/chart${n}.xml`,
										"Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
									}
								});
								me.write ({file: `xl/drawings/_rels/drawing1.xml.rels`, object: o});
								cb ();
							});
						},
						function (cb) {
							if (n == 1) {
								return cb ();
							}
							me.read ({file: "xl/drawings/drawing1.xml"}, function (err, o) {
								if (err) {
									return cb (new VError (err, "generateMult"));
								}
								if (n == 2) {
									o ["xdr:wsDr"]["xdr:twoCellAnchor"] = [o ["xdr:wsDr"]["xdr:twoCellAnchor"]];
								}
								o ["xdr:wsDr"]["xdr:twoCellAnchor"].push ({
									"xdr:from": {
										"xdr:col": 0,
										"xdr:colOff": 0,
										"xdr:row": (n - 1) * 20,
										"xdr:rowOff": 0
									},
									"xdr:to": {
										"xdr:col": 10,
										"xdr:colOff": 0,
										"xdr:row": n * 20,
										"xdr:rowOff": 0
									},
									"xdr:graphicFrame": {
										"$": {
											"macro": ""
										},
										"xdr:nvGraphicFramePr": {
											"xdr:cNvPr": {
												"$": {
													"id": `${n + 1}`,
													"name": `Диаграмма ${n}`
												}
											},
											"xdr:cNvGraphicFramePr": {}
										},
										"xdr:xfrm": {
											"a:off": {
												"$": {
													"x": "0",
													"y": "0"
												}
											},
											"a:ext": {
												"$": {
													"cx": "0",
													"cy": "0"
												}
											}
										},
										"a:graphic": {
											"a:graphicData": {
												"$": {
													"uri": "http://schemas.openxmlformats.org/drawingml/2006/chart"
												},
												"c:chart": {
													"$": {
														"r:id": `rId${n}`,
														"xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
														"xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
													}
												}
											}
										}
									},
									"xdr:clientData": {}
								});
								me.write ({file: `xl/drawings/drawing1.xml`, object: o});
								cb ();
							});
						}
					], cb);
				}, cb);
			}
		], function (err) {
			if (err) {
				return cb (new VError (err, "build"));
			}
			let result = me.zip.generate ({type: me.type});
			
			cb (null, result);
		});
	}
});
module.exports = Chart;
