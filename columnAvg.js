var fs = require ("fs");
var XLSXChart = require ("./chart");
var xlsxChart = new XLSXChart ();
var opts = {
	chart: "columnAvg",
	title: "Title",
	fields: [
		"Field 1",
		"Field 2",
		"Field 3",
		"Field 4"
	],
	value: {
		"Field 1": 5,
		"Field 2": 10,
		"Field 3": 15,
		"Field 4": 20 
	},
	avg: 12,
	avgTitle: "Среднее",
	showVal: true,
	showValAvg: true
};
xlsxChart.generate (opts, function (err, data) {
	console.log (err);
	fs.writeFileSync ("chart.xlsx", data);
});
