var bpopts = {
	dynamicTyping: true,
	keepEmptyRows: true
};
function csv_to_aoa(str/*:string*/, opts)/*:AOA*/ {
	var data = babyParse.parse(str, bpopts).data;
	for(var R=0; R != data.length; ++R) for(var C=0; C != data[R].length; ++C) {
		var d = data[R][C];
		if(d === '') data[R][C] = null;
		else if(d === 'TRUE') data[R][C] = true;
		else if(d === 'FALSE') data[R][C] = false;
	}
	return data;
}

function csv_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(csv_to_aoa(str, opts), opts); }

function csv_to_workbook(str/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(csv_to_sheet(str, opts), opts); }

