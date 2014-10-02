/* TODO: find an actual specification */
var socialcalc_to_aoa = function(str) {
	var records = str.split('\n'), R = -1, C = -1, ri = 0, arr = [];
	for (; ri !== records.length; ++ri) {
		var record = records[ri].trim().split(":");
		if(record[0] !== 'cell') continue;
		var addr = decode_cell(record[1]);
		if(arr.length <= addr.r) for(R = arr.length; R <= addr.r; ++R) if(!arr[R]) arr[R] = [];
		R = addr.r; C = addr.c;
		switch(record[2]) {
			case 't': arr[R][C] = record[3]; break;
			case 'v': arr[R][C] = +record[3]; break;
			case 'vtc': case 'vtf':
				switch(record[3]) {
					case 'nl': arr[R][C] = +record[4] ? true : false; break;
					default: arr[R][C] = +record[4]; break;
				} break;
		}
	}
	return arr;
};

var socialcalc_to_sheet = function(str) { return aoa_to_sheet(socialcalc_to_aoa(str)); };

var socialcalc_to_workbook = function (str) { return sheet_to_workbook(socialcalc_to_sheet(str)); };

