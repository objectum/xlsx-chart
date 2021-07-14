# xlsx-chart
Node.js excel chart builder

## Quick start

Install
```bash
npm install xlsx-chart
```

Generate and write chart to file
```js
var XLSXChart = require ("xlsx-chart");
var xlsxChart = new XLSXChart ();
var opts = {
	file: "chart.xlsx",
	chart: "column",
	titles: [
		"Title 1",
		"Title 2",
		"Title 3"
	],
	fields: [
		"Field 1",
		"Field 2",
		"Field 3",
		"Field 4"
	],
	data: {
		"Title 1": {
			"Field 1": 5,
			"Field 2": 10,
			"Field 3": 15,
			"Field 4": 20 
		},
		"Title 2": {
			"Field 1": 10,
			"Field 2": 5,
			"Field 3": 20,
			"Field 4": 15
		},
		"Title 3": {
			"Field 1": 20,
			"Field 2": 15,
			"Field 3": 10,
			"Field 4": 5
		}
	}
};
xlsxChart.writeFile (opts, function (err) {
  console.log ("File: ", opts.file);
});

```

Generate and download chart data
```js
xlsxChart.generate (opts, function (err, data) {
	res.set ({
	  "Content-Type": "application/vnd.ms-excel",
	  "Content-Disposition": "attachment; filename=chart.xlsx",
	  "Content-Length": data.length
	});
	res.status (200).send (data);
});

```

## Chart types

column, bar, line, area, radar, scatter, pie

## Mixing

You can mix column, bar, line, area.

## Custom template

Default templates: xlsx-chart/template/*.xlsx
```js
var opts = {
	file: "chart.xlsx",
	chart: "column",
	templatePath: __dirname + "/myColumn.xlsx",
	...
};
xlsxChart.writeFile (opts, function (err) {
  console.log ("File: ", opts.file);
});

```

## Multiple charts (one type)

Only column chart. For other types use custom template.

```js
let fs = require ("fs");
let XLSXChart = require ("xlsx-chart");
let xlsxChart = new XLSXChart ();
let opts = {
	charts: [{
		chart: "column",
		titles: [
			"Title 1",
			"Title 2",
			"Title 3"
		],
		fields: [
			"Field 1",
			"Field 2",
			"Field 3",
			"Field 4"
		],
		data: {
			"Title 1": {
				"Field 1": 5,
				"Field 2": 10,
				"Field 3": 15,
				"Field 4": 20
			},
			"Title 2": {
				"Field 1": 10,
				"Field 2": 5,
				"Field 3": 20,
				"Field 4": 15
			},
			"Title 3": {
				"Field 1": 20,
				"Field 2": 15,
				"Field 3": 10,
				"Field 4": 5
			}
		},
		chartTitle: "Title 1"
	}, {
		chart: "column",
		titles: [
			"Title 1",
			"Title 2",
			"Title 3"
		],
		fields: [
			"Field 1",
			"Field 2",
			"Field 3",
			"Field 4"
		],
		data: {
			"Title 1": {
				"Field 1": 5,
				"Field 2": 10,
				"Field 3": 15,
				"Field 4": 20
			},
			"Title 2": {
				"Field 1": 10,
				"Field 2": 5,
				"Field 3": 20,
				"Field 4": 15
			},
			"Title 3": {
				"Field 1": 20,
				"Field 2": 15,
				"Field 3": 10,
				"Field 4": 5
			}
		},
		chartTitle: "Title 2"
	}, {
		chart: "column",
		titles: [
			"Title 1",
			"Title 2",
		],
		fields: [
			"Field 1",
			"Field 2",
			"Field 3",
		],
		data: {
			"Title 1": {
				"Field 1": 15,
				"Field 2": 30,
				"Field 3": 45,
			},
			"Title 2": {
				"Field 1": 5,
				"Field 2": 2,
				"Field 3": 10
			}
		},
		chartTitle: "Title 3"
	}]
};
xlsxChart.generate (opts, function (err, data) {
	fs.writeFileSync ("chart.xlsx", data);
});
```

## Available chart options
```js
var opts = {
	file: "chart.xlsx", // exported file
	type: "nodebuffer", // optional: used by JSZip library
	charts: [
		{
			chart: "column", // pie, doughnut, line, area, bar
			titles: [
				"title1", // list of chart titles
				"title2", // list of chart titles
			],
			fields: [
				"field1", // list of chart fields
				"field2",
			],
			data: {
				"title1": {
					"field1": 123, // structured data
					"field2": 321,
				},
				"title2": {
					"field1": 456,
					"field2": 654,
				},
			},
			chartTitle: "Title",
			position: { // optional: chart position
				fromColumn: 0, 			// chart top left x coordinate in columns
				fromColumnOffset: 0, 	// chart top left x coordinate in pixels
				fromRow: n * 20, 		// chart top left y coordinate in columns
				fromRowOffset: 0, 		// chart top left y coordinate in pixels
				toColumn: 10, 			// chart bottom right x coordinate in columns
				toColumnOffset: 0, 		// chart bottom right x coordinate in pixels
				toRow: (n + 1) * 20, 	// chart bottom right y coordinate in columns
				toRowOffset: 0, 		// chart bottom right y coordinate in pixels
			},
			customColors: { // optional: chart colors
				points: {
					"title1": {
						"field1": "FF0000", // colors in the same structure as data; not all points required
						"field1": {
							fill: "FF0000", // separate colors for fill and border could be set
							line: "FF0000",
						},
					},
				},
				series: {
					"title1": "FF0000", // whole series color could be also set
					"title2": {
						fill: "FF0000", // separate colors for fill and border could be set
						line: "FF0000",
						markerColor: "FF0000", // different marker color also could be set for line chart
					},
				},
			},
			manualLayout: { // optional blocks position on the chart canvas
				plotArea: { // chart position
					x: 0.5, // x and y of top left corner in proportion value 0..1
					y: 0.5,
					w: 0.5, // width and height of block point in proportion value 0..1
					h: 0.5,
				},
				title: {
					x: 0.5, // x and y of top left corner in proportion value 0..1
					y: 0.5,
				},
			},
			firstSliceAng: 270, // optional: first slice angle for pie and doughnut chart in degrees
			holeSize: 50, // optional: hole size of doughnut chart in percent
			legendPos: "r", // optional: 'l', 't', 'r', 'b', null - legend position; null to hide; right is default
			deleteEmptyCells: false, // optional: allow to will empty values with 0 or leave as gaps
		},
	],
};

xlsxChart.writeFile (opts, function (err) {
  console.log ("File: ", opts.file);
});

```

## Examples

<a href="examples/column.js">column.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/column.png)  
<a href="examples/bar.js">bar.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/bar.png)  
<a href="examples/line.js">line.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/line.png)  
<a href="examples/area.js">area.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/area.png)  
<a href="examples/radar.js">radar.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/radar.png)  
<a href="examples/scatter.js">scatter.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/scatter.png)  
<a href="examples/pie.js">pie.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/pie.png)  
<a href="examples/columnLine.js">columnLine.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/columnLine.png)  
<a href="examples/mix.js">mix.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/mix.png)  

## Author

**Dmitriy Samortsev**

+ http://github.com/objectum


## Copyright and license

MIT
