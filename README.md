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

## Examples

<a href="https://github.com/objectum/xlsx-chart/master/examples/column.js">column.js</a>  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/column.png)  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/bar.png)  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/line.png)  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/area.png)  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/radar.png)  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/scatter.png)  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/pie.png)  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/columnLine.png)  
![alt tag](https://raw.github.com/objectum/xlsx-chart/master/examples/mix.png)  

## Author

**Dmitriy Samortsev**

+ http://github.com/objectum


## Copyright and license

MIT
