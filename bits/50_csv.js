var bpopts = {
	dynamicTyping: true,
	keepEmptyRows: true
};
var csv_to_aoa = function (str) {
	var data = babyParse.parse(str, bpopts).data;
	for(var R=0; R != data.length; ++R) for(var C=0; C != data[R].length; ++C) {
		var d = data[R][C];
		if(d === '') data[R][C] = null;
		else if(d === 'TRUE') data[R][C] = true;
		else if(d === 'FALSE') data[R][C] = false;
	}
	return data;
};

var csv_to_sheet = function (str) { return aoa_to_sheet(csv_to_aoa(str)); };

var csv_to_workbook = function (str) { return sheet_to_workbook(csv_to_sheet(str)); };

