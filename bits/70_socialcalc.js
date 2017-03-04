/* TODO: find an actual specification */
function scdecode(s/*:string*/)/*:string*/ {
	return s.replace(/\\b/g, "\\").replace(/\\c/g, ":").replace(/\\n/g,"\n");
}

function socialcalc_to_aoa(str/*:string*/, opts)/*:AOA*/ {
	var records = str.split('\n'), R = -1, C = -1, ri = 0, arr = [];
	for (; ri !== records.length; ++ri) {
		var record = records[ri].trim().split(":");
		if(record[0] !== 'cell') continue;
		var addr = decode_cell(record[1]);
		if(arr.length <= addr.r) for(R = arr.length; R <= addr.r; ++R) if(!arr[R]) arr[R] = [];
		R = addr.r; C = addr.c;
		switch(record[2]) {
			case 't': arr[R][C] = scdecode(record[3]); break;
			case 'v': arr[R][C] = +record[3]; break;
			case 'vtc': case 'vtf':
				switch(record[3]) {
					case 'nl': arr[R][C] = +record[4] ? true : false; break;
					default: arr[R][C] = +record[4]; break;
				} break;
		}
	}
	return arr;
}

function socialcalc_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(socialcalc_to_aoa(str, opts), opts); }

function socialcalc_to_workbook(str/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(socialcalc_to_sheet(str, opts), opts); }


/* originally from http://git.io/xlsx2socialcalc */
/* xlsx2socialcalc.js (C) 2014-present SheetJS -- http://sheetjs.com */
var sheet_to_socialcalc = (function() {
	var header = [
		"socialcalc:version:1.5",
		"MIME-Version: 1.0",
		"Content-Type: multipart/mixed; boundary=SocialCalcSpreadsheetControlSave"
	].join("\n");

	var sep = [
		"--SocialCalcSpreadsheetControlSave",
		"Content-type: text/plain; charset=UTF-8",
		""
	].join("\n");

	/* TODO: the other parts */
	var meta = [
		"# SocialCalc Spreadsheet Control Save",
		"part:sheet"
	].join("\n");

	var end = "--SocialCalcSpreadsheetControlSave--";

	var scencode = function(s) { return s.replace(/\\/g, "\\b").replace(/:/g, "\\c").replace(/\n/g,"\\n"); };

	var scsave = function scsave(ws) {
		if(!ws || !ws['!ref']) return "";
		var o = [], oo = [], cell, coord;
		var r = decode_range(ws['!ref']);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			for(var C = r.s.c; C <= r.e.c; ++C) {
				coord = encode_cell({r:R,c:C});
				if(!(cell = ws[coord]) || cell.v == null) continue;
				oo = ["cell", coord, 't'];
				switch(cell.t) {
					case 's': case 'str': oo.push(scencode(cell.v)); break;
					case 'n':
						if(cell.f) {
							oo[2] = 'vtf';
							oo.push('n');
							oo.push(cell.v);
							oo.push(scencode(cell.f));
						}
						else {
							oo[2] = 'v';
							oo.push(cell.v);
						} break;
					case 'b':
						if(cell.f) {
							oo[2] = 'vtf';
							oo.push('nl');
							oo.push(cell.v ? 1 : 0);
							oo.push(scencode(cell.f));
						} else {
							oo[2] = 'vtc';
							oo.push('nl');
							oo.push(cell.v ? 1 : 0);
							oo.push(cell.v ? 'TRUE' : 'FALSE');
						} break;
				}
				o.push(oo.join(":"));
			}
		}
		o.push("sheet:c:" + (r.e.c - r.s.c + 1) + ":r:" + (r.e.r - r.s.r + 1) + ":tvf:1");
		o.push("valueformat:1:text-wiki");
		o.push("copiedfrom:" + ws['!ref']);
		return o.join("\n");
	};

	return function socialcalcify(ws/*:Worksheet*/, opts/*:?any*/)/*:string*/ {
		return [header, sep, meta, sep, scsave(ws), end].join("\n");
		// return ["version:1.5", scsave(ws)].join("\n"); // clipboard form
	};
})();
