var fs = require ("fs");
var XLSXChart = require ("./../chart");
var xlsxChart = new XLSXChart ();
var opts = {
	titles: [
		"Title 1",
		"Title 2",
		"Title 3",
		"Title 4"
	],
	fields: [
		"Field 1",
		"Field 2",
		"Field 3",
		"Field 4"
	],
	data: {
		"Title 1": {
			"chart": "column",
			"Field 1": 5,
			"Field 2": 10,
			"Field 3": 15,
			"Field 4": 20
		},
		"Title 2": {
			"chart": "bar",
			"Field 1": 10,
			"Field 2": 5,
			"Field 3": 20,
			"Field 4": 15
		},
		"Title 3": {
			"chart": "line",
			"Field 1": 20,
			"Field 2": 15,
			"Field 3": 10,
			"Field 4": 5
		},
		"Title 4": {
			"chart": "area",
			"Field 1": 10,
			"Field 2": 5,
			"Field 3": 20,
			"Field 4": 15
		}
	},
	chartTitle: "Column, bar, line and area chart"
};
xlsxChart.generate (opts, function (err, data) {
	if (err) {
		console.error (err);
	} else {
		fs.writeFileSync ("mix.xlsx", data);
		console.log ("mix.xlsx created.");
	};
});
