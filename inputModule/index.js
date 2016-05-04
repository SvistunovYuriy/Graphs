var excelParser = require('xlsx');

var inputFile = excelParser.readFile("input.xlsx");

var worksheet = inputFile.Sheets["Лист1"];

var finishRow = 2404,
	startRow = 4;

var inputMapper = {
	"Ampl spektr zond signala":{
		"20": "A",
		"30": "B",
		"40": "C",
		"50": "D",
		"60": "E",
		"80": "F",
		"100": "G"
	},
	"Fazov spektr zond signala":{
		"20": "H",
		"30": "I",
		"40": "J",
		"50": "K",
		"60": "L",
		"80": "M",
		"100": "N"
	}
}

var result = {
	getColumnValues: function(columnHeader){
		var res = [];
		for(var i = startRow; i < finishRow; i++){
			res.push(worksheet[columnHeader + i].v);
		}
		return res;
	},
	getAmplSpektr: function(number){
		return this.getColumnValues(inputMapper["Ampl spektr zond signala"][number]);
	},
	getFazSpektr: function(number){
		return this.getColumnValues(inputMapper["Fazov spektr zond signala"][number]);

	}
}

module.export = result;