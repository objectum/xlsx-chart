var _ = require ("underscore"); 
var Backbone = require ("backbone");
var JSZip = require ("jszip");
var xml2js = require ("xml2js");
var VError = require ("verror");
var fs = require ("fs");
var async = require ("async");
var BaseChart = require ("./base");
var Chart = BaseChart.extend ({
	/*
		Write table
	*/
	writeTable: function (cb) {
		var me = this;
		me.read ({file: "xl/worksheets/sheet2.xml"}, function (err, o) {
			if (err) {
				return cb (new VError (err, "writeTable"));
			};
			o.worksheet.dimension.$.ref = "A1:" + me.getColName (me.fields.length + 1) + "3";
			var rows = [];
			// groups
			var cells = [], col = 2, mergeCells = [];
			_.each (me.groups, function (g, i) {
				cells.push ({
					$: {
						r: me.getColName (col) + "1",
						s: "1",
						t: "s"
					},
					v: me.getStr (g.name)
				});
				for (var j = 1; j < g.length; j ++) {
					cells.push ({
						$: {
							r: me.getColName (col + j) + "1",
							s: "1"
						}
					});
				};
				mergeCells.push ({
					$: {
						ref: me.getColName (col) + "1:" + me.getColName (col + g.length - 1) + "1"
					}
				});
				col += g.length;
			});
			rows.push ({
				$: {
					r: "1",
					spans: "1:" + (me.fields.length + 1),
					"x14ac:dyDescent": "0.2"
				},
				c: cells
			});
			o.worksheet.mergeCells = {
				mergeCell: mergeCells
			};
			// fields
			cells = [];
			_.each (me.fields, function (f, i) {
				cells.push ({
					$: {
						r: me.getColName (i + 2) + "2",
						t: "s"
					},
					v: me.getStr (f)
				});
			});
			rows.push ({
				$: {
					r: "2",
					spans: "1:" + (me.fields.length + 1),
					"x14ac:dyDescent": "0.2"
				},
				c: cells
			});
			// value
			cells = [{
				$: {
					r: "A3",
					t: "s"
				},
				v: me.getStr (me.title)
			}];
			_.each (me.fields, function (f, i) {
				cells.push ({
					$: {
						r: me.getColName (i + 2) + "3"
					},
					v: me.value [f]
				});
			});
			rows.push ({
				$: {
					r: "3",
					spans: "1:" + (me.fields.length + 1),
					"x14ac:dyDescent": "0.2"
				},
				c: cells
			});
			// avg
			cells = [{
				$: {
					r: "A4",
					t: "s"
				},
				v: me.getStr (me.avgTitle)
			}];
			_.each (me.fields, function (f, i) {
				cells.push ({
					$: {
						r: me.getColName (i + 2) + "4"
					},
					v: me.avg
				});
			});
			rows.push ({
				$: {
					r: "4",
					spans: "1:" + (me.fields.length + 1),
					"x14ac:dyDescent": "0.2"
				},
				c: cells
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
			};
			o.sst.$.count = 2 + me.groups.length + me.fields.length;
			o.sst.$.uniqueCount = o.sst.$.count;
			var si = [{t: me.title}, {t: me.avgTitle}];
			_.each (me.groups, function (g) {
				si.push ({t: g.name});
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
			};
			var groupCol = 0;
			var ser = [{
				"c:idx": {
					$: {
						val: 0
					}
				},
				"c:order": {
					$: {
						val: 0
					}
				},
				"c:tx": {
					"c:strRef": {
						"c:f": "Table!$A$3",
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
								"c:v": me.title
							}
						}
					}
				},
				"c:spPr": {
					"a:solidFill": {
						"a:schemeClr": {
							$: {
								val: "accent1"
							}
						}
					},
					"a:ln": {
						"a:noFill": {}
					},
					"a:effectLst": {}
				},
				"c:invertIfNegative": {
					$: {
						val: "0"
					}
				},
				"c:cat": {
					"c:multiLvlStrRef": {
						"c:f": "Table!$B$1:$" + me.getColName (me.fields.length + 1) + "$2",
						"c:multiLvlStrCache": {
							"c:ptCount": {
								$: {
									val: me.fields.length
								}
							},
							"c:lvl": [{
								"c:pt": _.map (me.fields, function (f, j) {
									return {
										$: {
											idx: j
										},
										"c:v": f
									};
								})
							}, {
								"c:pt": _.map (me.groups, function (g, j) {
									var o = {
										$: {
											idx: groupCol
										},
										"c:v": g.name
									};
									groupCol += g.length;
									return o;
								})
							}]
						}
					}
				},
				"c:val": {
					"c:numRef": {
						"c:f": "Table!$B$3:$" + me.getColName (me.fields.length + 1) + "$3",
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
									"c:v": me.value [f]
								};
							})
						}
					}
				}
			}];
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:ser"] = ser;
			if (me.chartTitle) {
				me.writeTitle (o, me.chartTitle);
			};
			if (me.showVal) {
				o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:dLbls"]["c:showVal"] = {
					$: {
						val: "1"
					}
				};
			};
			me.write ({file: "xl/charts/chart1.xml", object: o});
			cb ();
		});
	},
	/*
		Write avg
	*/
	writeAvg: function (cb) {
		var me = this;
		me.read ({file: "xl/charts/chart1.xml"}, function (err, o) {
			if (err) {
				return cb (new VError (err, "writeAvg"));
			};
			var groupCol = 0;
			var ser = [{
				"c:idx": {
					$: {
						val: 1
					}
				},
				"c:order": {
					$: {
						val: 1
					}
				},
				"c:tx": {
					"c:strRef": {
						"c:f": "Table!$A$4",
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
								"c:v": me.avgTitle
							}
						}
					}
				},
				"c:spPr": {
					"a:ln": {
						$: {
							w: "28575",
							cap: "rnd"
						},
						"a:solidFill": {
							"a:schemeClr": {
								$: {
									val: "accent2"
								}
							}
						},
						"a:round": {}
					},
					"a:effectLst": {}
				},
				"c:marker": {
					"c:symbol": {
						$: {
							val: "none"
						}
					}
				},
				"c:cat": {
					"c:multiLvlStrRef": {
						"c:f": "Table!$B$1:$" + me.getColName (me.fields.length + 1) + "$2",
						"c:multiLvlStrCache": {
							"c:ptCount": {
								$: {
									val: me.fields.length
								}
							},
							"c:lvl": [{
								"c:pt": _.map (me.fields, function (f, j) {
									return {
										$: {
											idx: j
										},
										"c:v": f
									};
								})
							}, {
								"c:pt": _.map (me.groups, function (g, j) {
									var o = {
										$: {
											idx: groupCol
										},
										"c:v": g.name
									};
									groupCol += g.length;
									return o;
								})
							}]
						}
					}
				},
				"c:val": {
					"c:numRef": {
						"c:f": "Table!$B$4:$" + me.getColName (me.fields.length + 1) + "$4",
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
									"c:v": me.avg
								};
							})
						}
					}
				}
			}];
			if (me.showValAvg) {
				var l = {
					"c:dLbl": [],
					"c:spPr": {
						"a:noFill": {},
						"a:ln": {
							"a:noFill": {}
						},
						"a:effectLst": {}
					},
					"c:dLblPos": {
						$: {
							val: "r"
						}
					},
					"c:showLegendKey": {
						$: {
							val: "0"
						}
					},
					"c:showVal": {
						$: {
							val: "1"
						}
					},
					"c:showCatName": {
						$: {
							val: "0"
						}
					},
					"c:showSerName": {
						$: {
							val: "0"
						}
					},
					"c:showPercent": {
						$: {
							val: "0"
						}
					},
					"c:showBubbleSize": {
						$: {
							val: "0"
						}
					},
					"c:showLeaderLines": {
						$: {
							val: "0"
						}
					},
					"c:extLst": {
						"c:ext": {
							$: {
								"xmlns:c15": "http://schemas.microsoft.com/office/drawing/2012/chart",
								"uri": "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"
							},
							"c15:layout": {},
							"c15:showLeaderLines": {
								$: {
									val: "1"
								}
							}
						}
					}
				};
				for (var i = 0; i < me.fields.length - 1; i ++) {
					l ["c:dLbl"].push ({
						"c:idx": {
							$: {
								val: String (i)
							}
						},
						"c:delete": {
							$: {
								val: "1"
							}
						},
						"c:extLst": {
							"c:ext": {
								$: {
									"xmlns:c15": "http://schemas.microsoft.com/office/drawing/2012/chart",
									"uri": "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"
								},
								"c15:layout": {}
							}
						}
					});
				};
				ser [0]["c:dLbls"] = l;
				ser [0]["c:smooth"] = {
					$: {
						val: "1"
					}
				};
			};
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:lineChart"]["c:ser"] = ser;
			me.write ({file: "xl/charts/chart1.xml", object: o});
			cb ();
		});
	},
	/*
		Generate
	*/
	generate: function (opts, cb) {
		var me = this;
		opts.type = opts.type || "nodebuffer";
		_.extend (me, opts);
		async.series ([
			function (cb) {
				me.zip = new JSZip ();
				fs.readFile (__dirname + "/../template/columnGroupAvg.xlsx", function (err, data) {
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
				_.each (me.fields, function (f) {
					me.value [f] = me.value [f] || 0;
				});
				me.writeTable (cb);
			},
			function (cb) {
				me.writeChart (cb);
			},
			function (cb) {
				me.writeAvg (cb);
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
