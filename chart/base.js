var _ = require ("underscore"); 
var Backbone = require ("backbone");
var JSZip = require ("jszip");
var xml2js = require ("xml2js");
var VError = require ("verror");
var fs = require ("fs");
var async = require ("async");

const CHART_TAG_BY_CHART_NAME = {
	bar: "c:barChart",
	column: "c:barChart",
	line: "c:lineChart",
	radar: "c:radarChart",
	area: "c:areaChart",
	scatter: "c:scatterChart",
	pie: "c:pieChart",
	doughnut: "c:doughnutChart",
};

const CHART_GROUPING_BY_CHART_NAME = {
	bar: "clustered",
	column: "clustered",
	line: "standard",
	radar: undefined, // radar, scatter, pie and doughnut charts should not have "c:grouping" tag, otherwise rest of xml is ignored
	area: "standard",
	scatter: undefined,
	pie: undefined,
	doughnut: undefined,
};

const CHART_TYPES = ["bar", "column", "line", "radar", "area", "scatter", "pie", "doughnut"];

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
		me.zip.file (opts.file, Buffer.from (xml), {base64: true});
	},
	/*
		Get column name
	*/
	getColName: function (n) {
		var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		n--;
		if (n < 26) {
			return abc[n];
		} else {
			return abc[(n / 26 - 1) | 0] + abc[n % 26];
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
		return me.str[s];
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
						v: me.data[t][f]
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
						v: me.data[t][f]
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
				me.str[o.t] = i;
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

		// axis ids that are used in chart layers
		var axId = [];
		function addId (o) {
			if (!o ["c:axId"]) {
				return;
			}
			_.each (o ["c:axId"], function (o) {
				axId.push (o.$.val);
			});
		};

		// remove chart layers that are not used in chart
		_.each (CHART_TYPES, function (chart) {
			if (chart == "column") {
				chart = "bar";
			}

			if (!o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + chart + "Chart"]) {
				return;
			}

			if (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + chart + "Chart"].length) {
				_.each (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + chart + "Chart"], addId);
			} else {
				addId (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + chart + "Chart"]);
			}
		});

		if (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"]) {
			var catAx = [];
			// if category axes is single - it's an object, otherwise it's an array of objects
			if (!o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"].length) {
				o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"] = [o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"]];
			}

			// leave only category axes that are used in chart
			_.each (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"], function (o) {
				if (axId.indexOf (o ["c:axId"].$.val) > -1) {
					catAx.push (o);
				};
			});
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:catAx"] = catAx;
		}

		if (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"]) {
			var valAx = [];
			// if value axes is single - it's an object, otherwise it's an array of objects
			if (!o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"].length) {
				o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"] = [o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"]];
			}

			// leave only value axes that are used in chart
			_.each (o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"], function (o) {
				if (axId.indexOf (o ["c:axId"].$.val) > -1) {
					valAx.push (o);
				};
			});
			o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:valAx"] = valAx;
		}
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
			var seriesByChart = {};
			const chartOpts = me.charts[chartN - 1];
			_.each (me.titles, function (t, i) {
				var chart = me.data[t].chart || me.chart;
				var grouping = me.data[t].grouping || me.grouping || CHART_GROUPING_BY_CHART_NAME[chart];

				var customColorsPoints = {
					"c:dPt": [],
				};
				var customColorsSeries = {};

				if (chartOpts.customColors) {
					const customColors = chartOpts.customColors;

					if (customColors.points) {
						customColorsPoints["c:dPt"] = chartOpts.fields.map (function (field, i) {
							const color = _.chain (customColors).get ("points").get (t).get (field, null).value ();

							if (!color) {
								return null;
							}
							if (color === "noFill") {
								return {
									"c:idx": {
										$: {
											val: i,
										},
									},
									"c:spPr": {
										"a:noFill": ""
									}
								}
							}
							let fillColor = color;
							let lineColor = color;
							if (typeof color === "object") {
								fillColor = color.fill;
								lineColor = color.line;
							}
							return {
								"c:idx": {
									$: {
										val: i,
									},
								},
								"c:spPr": {
									"a:solidFill": {
										"a:srgbClr": {
											$: {
												val: fillColor,
											},
										},
									},
									"a:ln": {
										"a:solidFill": {
											"a:srgbClr": {
												$: {
													val: lineColor,
												},
											},
										},
									},
								},
							}
						}).filter (Boolean);
					}

					if (customColors.series && customColors.series[t]) {
						let fillColor = customColors.series[t];
						let lineColor = customColors.series[t];
						let markerColor = customColors.series[t];
						if (typeof customColors.series === "object") {
							fillColor = customColors.series[t].fill;
							lineColor = customColors.series[t].line;
							markerColor = customColors.series[t].marker;
						}
						customColorsSeries["c:spPr"] = {
							"a:solidFill": {
								"a:srgbClr": {
									$: {
										val: fillColor,
									},
								},
							},
							"a:ln": {
								"a:solidFill": {
									"a:srgbClr": {
										$: {
											val: lineColor,
										},
									},
								},
							},
							"c:marker": {
								"c:spPr": {
									"a:solidFill": {
										"a:srgbClr": {
											$: {
												val: markerColor,
											},
										},
									},
								},
							},
						};
					}
				}

				var ser = {
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
					...customColorsPoints,
					...customColorsSeries,
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
										"c:v": me.data[t][f]
									};
								})
							}
						}
					}
				};
				if (chart == "scatter") {
					ser["c:xVal"] = ser["c:cat"];
					delete ser["c:cat"];
					ser["c:yVal"] = ser["c:val"];
					delete ser["c:val"];
					ser["c:spPr"] = {
						"a:ln": {
							$: {
								w: 28575
							},
							"a:noFill": ""
						}
					};
				};
				const seriesKey = `${chart}\r\r${grouping}`;
				seriesByChart[seriesKey] = seriesByChart[seriesKey] || [];
				seriesByChart[seriesKey].push (ser);
			});

			const templateChartName = "c:" + CHART_TYPES.find ((chartType) => me.chartTemplate ["c:chartSpace"]["c:chart"]["c:plotArea"][`c:${chartType}Chart`]) + "Chart";

			// remove template barChart from the XML object;
			o ["c:chartSpace"]["c:chart"]["c:plotArea"][templateChartName] = [];

			_.each (seriesByChart, function (ser, chart) {

				var [chart, grouping] = chart.split ("\r\r");

				const chartTagName = CHART_TAG_BY_CHART_NAME[chart];

				if (!chartTagName) {
					return cb (new VError (new Error (`Chart type '${chart}' is not supported`), "writeChart"));
				}

				o ["c:chartSpace"]["c:chart"]["c:plotArea"][chartTagName] = o ["c:chartSpace"]["c:chart"]["c:plotArea"][chartTagName] || [];
				// minimal chart config
				let newChart = {};

				if (chart == "column" || chart == "bar") {
					// clone barChart from template
					newChart = _.clone (me.chartTemplate ["c:chartSpace"]["c:chart"]["c:plotArea"][templateChartName]);
					newChart["c:barDir"] = {
						$: {
							val: chart.substr (0, 3),
						},
					};
					
				} else if (chart == "line" || chart == "area" || chart == "radar" || chart == "scatter") {
					newChart = _.clone (me.chartTemplate ["c:chartSpace"]["c:chart"]["c:plotArea"][templateChartName]);
					delete newChart["c:barDir"];
				} else {
					newChart["c:varyColors"] = {
						$: {
							val: 1,
						},
					};
					
					newChart["c:ser"] = ser;
					if (chartOpts.firstSliceAng) {
						newChart["c:firstSliceAng"] = {
							$: {
								val: chartOpts.firstSliceAng,
							},
						};
					}
					if (chartOpts.holeSize) {
						newChart["c:holeSize"] = {
							$: {
								val: chartOpts.holeSize,
							},
						};
					}
					// if (chartOpts.showLabels) {
					// 	newChart["c:dLbls"] = {
					// 		"c:showLegendKey": {
					// 			$: {
					// 				val: 1,
					// 			},
					// 		},
					// 		"c:showVal": {
					// 			$: {
					// 				val: 0,
					// 			},
					// 		},
					// 		"c:showCatName": {
					// 			$: {
					// 				val: 0,
					// 			},
					// 		},
					// 		"c:showSerName": {
					// 			$: {
					// 				val: 0,
					// 			},
					// 		},
					// 		"c:showPercent": {
					// 			$: {
					// 				val: 0,
					// 			},
					// 		},
					// 		"c:showBubbleSize": {
					// 			$: {
					// 				val: 0,
					// 			},
					// 		},
					// 		"c:showLeaderLines": {
					// 			$: {
					// 				val: 1,
					// 			},
					// 		},
					// 	};
					// }
				};

				newChart["c:ser"] = ser;

				if (grouping != "undefined") {
					newChart["c:grouping"] = {
						$: {
							val: grouping || CHART_GROUPING_BY_CHART_NAME[chart],
						},
					};
					if (grouping == "stacked") {
						newChart["c:overlap"] = {
							$: {
								val: 100, // usually stacked expected to be seen with overlap 100%
							},
						};
					}
				}

				o ["c:chartSpace"]["c:chart"]["c:plotArea"][chartTagName].push (newChart);

				if (chartOpts.legendPos === undefined || chartOpts.legendPos) {
					o ["c:chartSpace"]["c:chart"]["c:legend"] = {
						"c:legendPos": {
							$: {
								val: chartOpts.legendPos || "r",
							},
						},
					};
				} else if (chartOpts.legendPos === null) {
					delete o ["c:chartSpace"]["c:chart"]["c:legend"];
				}

				if (chartOpts.manualLayout && chartOpts.manualLayout.plotArea) {
					o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:layout"] = o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:layout"] || {};
					o ["c:chartSpace"]["c:chart"]["c:plotArea"]["c:layout"]["c:manualLayout"] = {
						"c:xMode": {
							$: {
								val: "edge",
							},
						},
						"c:yMode": {
							$: {
								val: "edge",
							},
						},
						"c:x": {
							$: {
								val: chartOpts.manualLayout.plotArea.x,
							},
						},
						"c:y": {
							$: {
								val: chartOpts.manualLayout.plotArea.y,
							},
						},
						"c:w": {
							$: {
								val: chartOpts.manualLayout.plotArea.w,
							},
						},
						"c:h": {
							$: {
								val: chartOpts.manualLayout.plotArea.h,
							},
						},
					};
				}
			});
			me.removeUnusedCharts (o);

			if (me.chartTitle) {
				me.writeTitle (o, me.chartTitle, chartOpts);
			};
			me.write ({file: `xl/charts/chart${chartN}.xml`, object: o});
			cb ();
		});
	},
	/*
		Chart title
	*/
	writeTitle: function (o, title, chartOpts = {}) {
		var me = this;
		const layout = {};

		if (chartOpts.manualLayout && chartOpts.manualLayout.title) {
			layout["c:manualLayout"] = {
				"c:xMode": {
					$: {
						val: "edge",
					},
				},
				"c:yMode": {
					$: {
						val: "edge",
					},
				},
				"c:x": {
					$: {
						val: chartOpts.manualLayout.title.x,
					},
				},
				"c:y": {
					$: {
						val: chartOpts.manualLayout.title.y,
					},
				},
			};
		}

		o ["c:chartSpace"]["c:chart"]["c:title"] = {
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
			"c:layout": layout,
			"c:overlay": {
				$: {
					val: "0"
				}
			}
		};
		o ["c:chartSpace"]["c:chart"]["c:autoTitleDeleted"] = {
			$: {
				val: "0"
			}
		};
	},
	/*
		Set template name
	*/
	setTemplateName: function (opts) {
		var me = this;
		var chartTypes = {
			...getChartTypes (me),
		};

		_.each (opts.charts, function (chart) {
			chartTypes = {
				...chartTypes,
				...getChartTypes (chart),
			};
		});
		me.charts = opts.charts || [opts];

		me.chartTypes = chartTypes;

		if (chartTypes ["radar"]) {
			me.tplName = "radar";
			return;
		};
		if (chartTypes["scatter"]) {
			me.tplName = "scatter";
			return;
		};
		if (chartTypes["pie"]) {
			me.tplName = "pie";
			return;
		};
		if (_.keys (chartTypes).length == 1) {
			me.tplName = _.keys (chartTypes)[0];
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
				me.setTemplateName (opts);
				let path = me.templatePath ? me.templatePath : (__dirname + "/../template/" + me.tplName + ".xlsx");
				fs.readFile (path, function (err, data) {
					if (err) {
						console.error (`Template ${path} not read: ${err}`);
						return cb (err);
					};
					me.zip.load (data);
					cb ();
				});
			},
			function (cb) {
				// save template chart so it could be easily removed before adding charts according to config
				me.read ({file: "xl/charts/chart1.xml"}, function (err, o) {
					if (err) {
						return cb (new VError (err, "writeChart"));
					}
					me.chartTemplate = o;
					cb ();
				});
			},
			function (cb) {
				me.writeStrings (cb);
			},
			function (cb) {
				_.each (me.titles, function (t) {
					me.data[t] = me.data[t] || {};
					_.each (me.fields, function (f) {
						me.data[t][f] = me.data[t][f] || (me.deleteEmptyCells ? "" : 0); //deleteEmptyCells - don't display missing values as 0
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

		me.setTemplateName (opts);
		
		async.series ([
			function (cb) {
				me.zip = new JSZip ();

				let path = me.templatePath || (__dirname + "/../template/mult.xlsx");

				fs.readFile (path, function (err, data) {
					if (err) {
						console.error (`Template ${path} not read: ${err}`);
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
					["chart", "titles", "fields", "data", "chartTitle"].forEach (a => me[a] = chart[a]);

					_.each (me.titles, function (t) {
						me.data[t] = me.data[t] || {};

						_.each (me.fields, function (f) {
							me.data[t][f] = me.data[t][f] || 0;
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
					["chart", "titles", "fields", "data", "chartTitle"].forEach (a => me[a] = chart[a]);

					const position = Object.assign ({
						fromColumn: 0,
						fromColumnOffset: 0,
						fromRow: n * 20,
						fromRowOffset: 0,
						toColumn: 10,
						toColumnOffset: 0,
						toRow: (n + 1) * 20,
						toRowOffset: 0,
					},
					chart.position
					);

					async.series ([
						function (cb) {
							// save template chart so it could be easily removed before adding charts according to config
							me.read ({file: "xl/charts/chart1.xml"}, function (err, o) {
								if (err) {
									return cb (new VError (err, "writeChart"));
								}
								me.chartTemplate = o;
								cb ();
							});
						},
						function (cb) {
							me.writeChart (++n, row, cb);
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
								me.write ({file: "[Content_Types].xml", object: o});
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
								me.write ({file: "xl/drawings/_rels/drawing1.xml.rels", object: o});
								cb ();
							});
						},
						function (cb) {
							me.read ({file: "xl/drawings/drawing1.xml"}, function (err, o) {
								if (err) {
									return cb (new VError (err, "generateMult"));
								}
								if (n == 1) {
									o ["xdr:wsDr"]["xdr:twoCellAnchor"]["xdr:from"]["xdr:col"] = position.fromColumn;
									o ["xdr:wsDr"]["xdr:twoCellAnchor"]["xdr:from"]["xdr:colOff"] = position.fromColumnOffset;
									o ["xdr:wsDr"]["xdr:twoCellAnchor"]["xdr:from"]["xdr:row"] = position.fromRow;
									o ["xdr:wsDr"]["xdr:twoCellAnchor"]["xdr:from"]["xdr:rowOff"] = position.fromRowOffset;
									o ["xdr:wsDr"]["xdr:twoCellAnchor"]["xdr:to"]["xdr:col"] = position.toColumn;
									o ["xdr:wsDr"]["xdr:twoCellAnchor"]["xdr:to"]["xdr:colOff"] = position.toColumnOffset;
									o ["xdr:wsDr"]["xdr:twoCellAnchor"]["xdr:to"]["xdr:row"] = position.toRow;
									o ["xdr:wsDr"]["xdr:twoCellAnchor"]["xdr:to"]["xdr:rowOff"] = position.toRowOffset;
									me.write ({file: "xl/drawings/drawing1.xml", object: o});
									return cb ();
								}
								if (n == 2) {
									o ["xdr:wsDr"]["xdr:twoCellAnchor"] = [o ["xdr:wsDr"]["xdr:twoCellAnchor"]];
								}
								o ["xdr:wsDr"]["xdr:twoCellAnchor"].push ({
									"xdr:from": {
										"xdr:col": position.fromColumn,
										"xdr:colOff": position.fromColumnOffset,
										"xdr:row": position.fromRow,
										"xdr:rowOff": position.fromRowOffset
									},
									"xdr:to": {
										"xdr:col": position.toColumn,
										"xdr:colOff": position.toColumnOffset,
										"xdr:row": position.toRow,
										"xdr:rowOff": position.toRowOffset
									},
									"xdr:graphicFrame": {
										"$": {
											"macro": ""
										},
										"xdr:nvGraphicFramePr": {
											"xdr:cNvPr": {
												"$": {
													"id": `${n + 1}`,
													"name": `Diagram ${n}`
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
								me.write ({file: "xl/drawings/drawing1.xml", object: o});
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

/**
 *
 * @param {Object} chart
 * @returns {Object.<string,Boolean>}
 */
const getChartTypes = (chart) => {
	const chartTypes = {};
	if (chart.chart) {
		chartTypes[chart.chart] = true;
	}
	_.each (chart.data, function (series) {
		if (series.chart) {
			chartTypes[series.chart] = true;
		}
	});
	return chartTypes;
}
