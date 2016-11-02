var fs = require ("fs");
var XLSXChart = require ("./chart");
var xlsxChart = new XLSXChart ();
var opts = {
	chart: "columnGroup",
	title: "Значение",
	groups: [{
		name: "Group1", length: 3
	}, {
		name: "Group2", length: 2
	}],
	fields: [
		"Field 1",
		"Field 2",
		"Field 3",
		"Field 4",
		"Field 5"
	],
	value: {
		"Field 1": 5,
		"Field 2": 10,
		"Field 3": 15,
		"Field 4": 20,
		"Field 5": 15
	}
};
xlsxChart.generate (opts, function (err, data) {
	console.log (err);
	fs.writeFileSync ("chart.xlsx", data);
});
