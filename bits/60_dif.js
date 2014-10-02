var dif_to_aoa = function (str) {
	var records = str.split('\n'), R = -1, C = -1, ri = 0, arr = [];
	for (; ri !== records.length; ++ri) {
		if (records[ri].trim() === 'BOT') { arr[++R] = []; C = 0; continue; }
		if (R < 0) continue;
		var metadata = records[ri].trim().split(",");
		var type = metadata[0], value = metadata[1];
		++ri;
		var data = records[ri].trim();
		switch (+type) {
			case -1:
				if (data === 'BOT') { arr[++R] = []; C = 0; continue; }
				else if (data === 'EOD') break;
				else throw new Error("Unrecognized DIF special command " + data);
				break; /* technically unreachable */
			case 0:
				if(data === 'TRUE') arr[R][C++] = true;
				else if(data === 'FALSE') arr[R][C++] = false;
				else if(+value == +value) arr[R][C++] = +value;
				else {
					var d = new Date(value);
					if(d.getDate() === d.getDate()) arr[R][C++] = d;
				}
				break;
			case 1:
				data = data.substr(1,data.length-2);
				arr[R][C++] = data !== '' ? data : null;
				break;
		}
		if (data === 'EOD') break;
	}
	return arr;
};

var dif_to_sheet = function (str) { return aoa_to_sheet(dif_to_aoa(str)); };

var dif_to_workbook = function (str) { return sheet_to_workbook(dif_to_sheet(str)); };

