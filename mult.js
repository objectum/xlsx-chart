let fs = require ("fs");
let XLSXChart = require ("./chart");
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
	console.log (err);
	fs.writeFileSync ("chart.xlsx", data);
});
