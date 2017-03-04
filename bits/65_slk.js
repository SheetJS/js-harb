/* TODO: find an actual specification */
function sylk_to_aoa(str/*:string*/, opts)/*:AOA*/ {
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
}

function sylk_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(sylk_to_aoa(str, opts), opts); }

function sylk_to_workbook(str/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(sylk_to_sheet(str, opts), opts); }

function write_ws_cell_sylk(cell/*:Cell*/, ws/*:Worksheet*/, R/*:number*/, C/*:number*/, opts)/*:string*/ {
	var o = "C;Y" + (R+1) + ";X" + (C+1) + ";K";
	switch(cell.t) {
		case 'n': o += cell.v; break;
		case 'b': o += cell.v ? "TRUE" : "FALSE"; break;
		case 'e': o += cell.w || cell.v; break;
		case 'd': o += '"' + (cell.w || cell.v) + '"'; break;
		case 's': o += '"' + cell.v.replace(/"/g,"") + '"'; break;
	}
	return o;
}

function sheet_to_sylk(ws/*:Worksheet*/, opts/*:?any*/)/*:string*/ {
	var preamble/*:Array<string>*/ = ["ID;PWXL;N;E"], o/*:Array<string>*/ = [];
	preamble.push("P;PGeneral");
	var r = decode_range(ws['!ref']), cell/*:Cell*/;
	for(var R = r.s.r; R <= r.e.r; ++R) {
		for(var C = r.s.c; C <= r.e.c; ++C) {
			var coord = encode_cell({r:R,c:C});
			if(!(cell = ws[coord]) || cell.v == null) continue;
			o.push(write_ws_cell_sylk(cell, ws, R, C, opts));
		}
	}
	preamble.push("F;P0;DG0G8;M255");
	var RS = "\r\n";
	return preamble.join(RS) + RS + o.join(RS) + RS + "E" + RS;
}
