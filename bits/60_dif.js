function dif_to_aoa(str/*:string*/, opts)/*:AOA*/ {
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
				else if (data !== 'EOD') throw new Error("Unrecognized DIF special command " + data);
				break;
			case 0:
				if(data === 'TRUE') arr[R][C] = true;
				else if(data === 'FALSE') arr[R][C] = false;
				else if(+value == +value) arr[R][C] = +value;
				else if(!isNaN(new Date(value).getDate())) arr[R][C] = new Date(value);
				else arr[R][C] = value;
				++C; break;
			case 1:
				data = data.substr(1,data.length-2);
				arr[R][C++] = data !== '' ? data : null;
				break;
		}
		if (data === 'EOD') break;
	}
	return arr;
}

function dif_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(dif_to_aoa(str, opts), opts); }

function dif_to_workbook(str/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(dif_to_sheet(str, opts), opts); }



var sheet_to_dif = (function() {
	var push_field = function pf(o/*:Array<string>*/, topic/*:string*/, v/*:number*/, n/*:number*/, s/*:string*/) {
		o.push(topic);
		o.push(v + "," + n);
		o.push('"' + s.replace(/"/g,'""') + '"');
	};
	var push_value = function po(o/*:Array<string>*/, type/*:number*/, v/*:number*/, s/*:string*/) {
		o.push(type + "," + v);
		o.push(type == 1 ? '"' + s.replace(/"/g,'""') + '"' : s);
	};
	return function sheet_to_dif(ws/*:Worksheet*/, opts/*:?any*/)/*:string*/ {
		var o/*:Array<string>*/ = [];
		var r = decode_range(ws['!ref']), cell/*:Cell*/;
		push_field(o, "TABLE", 0, 1, "sheetjs");
		push_field(o, "VECTORS", 0, r.e.r - r.s.r + 1,"");
		push_field(o, "TUPLES", 0, r.e.c - r.s.c + 1,"");
		push_field(o, "DATA", 0, 0,"");
		for(var R = r.s.r; R <= r.e.r; ++R) {
			push_value(o, -1, 0, "BOT");
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var coord = encode_cell({r:R,c:C});
				if(!(cell = ws[coord]) || cell.v == null) { push_value(o, 1, 0, ""); continue;}
				switch(cell.t) {
					case 'n': push_value(o, 0, (cell.w || cell.v), "V"); break;
					case 'b': push_value(o, 0, cell.v ? 1 : 0, cell.v ? "TRUE" : "FALSE"); break;
					case 's': push_value(o, 1, 0, cell.v); break;
					default: push_value(o, 1, 0, "");
				}
			}
		}
		push_value(o, -1, 0, "EOD");
		var RS = "\r\n";
		var oo = o.join(RS);
		//while((oo.length & 0x7F) != 0) oo += "\0";
		return oo;
	};
})();
