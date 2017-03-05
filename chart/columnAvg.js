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
			}
			o.worksheet.dimension.$.ref = "A1:C" + (me.fields.length + 1);
			var rows = [{
				$: {
					r: 1,
					spans: "1:3"
				},
				c: [{
					$: {
						r: "B1",
						t: "s"
					},
					v: me.getStr (me.title)
				}, {
					$: {
						r: "C1",
						t: "s"
					},
					v: me.getStr (me.avgTitle)
				}]
			}];
			_.each (me.fields, function (f, y) {
				var r = {
					$: {
						r: y + 2,
						spans: "1:3"
					}
				};
				var c = [{
					$: {
						r: "A" + (y + 2),
						t: "s"
					},
					v: me.getStr (f)
				}];
				c.push ({
					$: {
						r: "B" + (y + 2)
					},
					v: me.value [f]
				});
				c.push ({
					$: {
						r: "C" + (y + 2)
					},
					v: me.avg
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
			o.sst.$.count = 2 + me.fields.length;
			o.sst.$.uniqueCount = o.sst.$.count;
			var si = [{t: me.title}, {t: me.avgTitle}];
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
						"c:f": "Table!$B$1",
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
						"c:f": "Table!$B$2:$B$" + (me.fields.length + 1),
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
						"c:f": "Table!$C$1",
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
				"c:marker": {
					"c:symbol": {
						$: {
							val: "none"
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
						"c:f": "Table!$C$2:$C$" + (me.fields.length + 1),
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
				fs.readFile (__dirname + "/../template/columnAvg.xlsx", function (err, data) {
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
