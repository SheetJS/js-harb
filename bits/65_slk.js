/* TODO: find an actual specification */
var sylk_to_aoa = function(str) {
	var records = str.split(/[\n\r]+/), R = -1, C = -1, ri = 0, rj = 0, arr = [];
	for (; ri !== records.length; ++ri) {
		var record = records[ri].trim().split(";");
		var RT = record[0], val;
		if(RT !== 'C' && RT !== 'F') continue;
		for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
			case 'Y':
				R = parseInt(record[rj].substr(1))-1; C = 0;
				for(var j = arr.length; j <= R; ++j) arr[j] = [];
				break;
			case 'X': C = parseInt(record[rj].substr(1))-1; break;
			case 'K':
				val = record[rj].substr(1);
				if(val.charAt(0) === '"') val = val.substr(1,val.length - 2);
				else if(val === 'TRUE') val = true;
				else if(val === 'FALSE') val = false;
				else if(+val === +val) val = +val;
				arr[R][C] = val;
				break;
		}
	}
	return arr;
};

var sylk_to_sheet = function(str) { return aoa_to_sheet(sylk_to_aoa(str)); };

var sylk_to_workbook = function (str) { return sheet_to_workbook(sylk_to_sheet(str)); };

