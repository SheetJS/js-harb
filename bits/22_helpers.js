function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function numdate(v) {
	var date = ssf.parse_date_code(v);
	var val = new Date();
	val.setUTCFullYear(date.y);
	val.setUTCMonth(date.m-1);
	val.setUTCDate(date.d);
	val.setUTCHours(date.H);
	val.setUTCMinutes(date.M);
	val.setUTCSeconds(date.S);
	return val;
}

var sheet_to_workbook = function (sheet) { return {SheetNames: ['Sheet1'], Sheets: {Sheet1: sheet}};	};

