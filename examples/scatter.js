var fs = require ("fs");
var XLSXChart = require ("./../chart");
var xlsxChart = new XLSXChart ();
var opts = {
	chart: "scatter",
	titles: [
		"Price"
	],
	fields: [
		"Apple",
		"Blackberry",
		"Strawberry",
		"Cowberry"
	],
	data: {
		"Price": {
			"Apple": 10,
			"Blackberry": 5,
			"Strawberry": 15,
			"Cowberry": 20
		}
	},
	chartTitle: "Scatter chart"
};
xlsxChart.generate (opts, function (err, data) {
	if (err) {
		console.error (err);
	} else {
		fs.writeFileSync ("scatter.xlsx", data);
		console.log ("scatter.xlsx created.");
	};
});
