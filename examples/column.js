var fs = require ("fs");
var XLSXChart = require ("./../chart");
var xlsxChart = new XLSXChart ();
var opts = {
	chart: "column",
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
	chartTitle: "Column chart"
};
xlsxChart.generate (opts, function (err, data) {
	if (err) {
		console.error (err);
	} else {
		fs.writeFileSync ("column.xlsx", data);
		console.log ("column.xlsx created.");
	};
});
