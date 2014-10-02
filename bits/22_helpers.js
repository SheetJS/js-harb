function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

var sheet_to_workbook = function (sheet) { return {SheetNames: ['Sheet1'], Sheets: {Sheet1: sheet}};	};

