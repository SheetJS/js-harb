/* TODO: find an actual specification */
var sylk_to_aoa = function(str) {
	var records = str.split(/[\n\r]+/), R = -1, C = -1, ri = 0, rj = 0, arr = [];
	var formats = [];
	var next_cell_format = null;
	for (; ri !== records.length; ++ri) {
		var record = records[ri].trim().split(";");
		var RT = record[0], val;
		if(RT === 'P') for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
			case 'P':
				formats.push(record[rj].substr(1));
				break;
		}
		else if(RT !== 'C' && RT !== 'F') continue;
		else for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
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
				else if(+val === +val) {
					val = +val;
					if(next_cell_format !== null && next_cell_format.match(/[ymdhmsYMDHMS]/)) val = numdate(val);
				}
				arr[R][C] = val;
				next_cell_format = null;
				break;
			case 'P':
				if(RT !== 'F') break;
				next_cell_format = formats[parseInt(record[rj].substr(1))];
		}
	}
	return arr;
};

var sylk_to_sheet = function(str) { return aoa_to_sheet(sylk_to_aoa(str)); };

var sylk_to_workbook = function (str) { return sheet_to_workbook(sylk_to_sheet(str)); };

