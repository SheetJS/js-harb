/* harb.js (C) 2014-present  SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint eqnull:true, funcscope:true */
var HARB = {};
(function make_harb (HARB) {
HARB.version = '0.1.1';
/*::
type AOA = Array<Array<any> >;
type CellAddrSpec = CellAddress | string; 
*/
/*::
	var babyParse = require('babyparse');
	var cptable = require('./dist/cpharb');
	var SSF = require('ssf');
	var fs = require('fs');
*/
if (typeof exports !== 'undefined') {
	if (typeof module !== 'undefined' && module.exports) {
		babyParse = require('babyparse');
		cptable = require('./dist/cpharb');
		SSF = require('ssf');
		fs = require('fs');
	}
}

function datenum(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
	var epoch = v.getTime();
	if(date1904) epoch += 1462*24*60*60*1000;
	return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}

function numdate(v/*:number*/)/*:Date*/ {
	var date = SSF.parse_date_code(v);
	var val = new Date();
	if(date == null) throw new Error("Bad Date Code: " + v);
	val.setUTCDate(date.d);
	val.setUTCMonth(date.m-1);
	val.setUTCFullYear(date.y);
	val.setUTCHours(date.H);
	val.setUTCMinutes(date.M);
	val.setUTCSeconds(date.S);
	return val;
}

function sheet_to_workbook(sheet/*:Worksheet*/, opts)/*:Workbook*/ {
	var n = opts && opts.sheet ? opts.sheet : "Sheet1";
	var sheets = {}; sheets[n] = sheet;
	return { SheetNames: [n], Sheets: sheets };
}

function aoa_to_sheet(data/*:AOA*/, opts)/*:Worksheet*/ {
	var o = opts || {};
	var ws/*:Worksheet*/ = ({}/*:any*/);
	var range/*:Range*/ = ({s: {c:10000000, r:10000000}, e: {c:0, r:0 }}/*:any*/);
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(data[R][C] == null) continue;
			var cell/*:Cell*/ = ({v: data[R][C] }/*:any*/);
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell_ref = encode_cell(({c:C,r:R}/*:any*/));
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n';
				cell.z = o.dateNF || SSF._table[14];
				cell.v = datenum(cell.v);
				cell.w = SSF.format(cell.z, cell.v);
			}
			else cell.t = 's';
			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = encode_range(range);
	return ws;
}

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

function set_text_arr(data/*:string*/, arr/*:AOA*/, R/*:number*/, C/*:number*/) {
	if(data === 'TRUE') arr[R][C] = true;
	else if(data === 'FALSE') arr[R][C] = false;
	else if(+data == +data) arr[R][C] = +data;
	else if(data !== "") arr[R][C] = data;
}

function prn_to_aoa(f/*:string*/, opts)/*:AOA*/ {
	var arr/*:AOA*/ = ([]/*:any*/);
	if(!f || f.length === 0) return arr;
	var lines = f.split(/[\r\n]/);
	var L = lines.length - 1;
	while(L >= 0 && lines[L].length === 0) --L;
	var start = 10, idx = 0;
	var R = 0;
	for(; R <= L; ++R) {
		idx = lines[R].indexOf(" ");
		if(idx == -1) idx = lines[R].length; else idx++;
		start = Math.max(start, idx);
	}
	for(R = 0; R <= L; ++R) {
		arr[R] = [];
		/* TODO: confirm that widths are always 10 */
		var C = 0;
		set_text_arr(lines[R].slice(0, start).trim(), arr, R, C);
		for(C = 1; C <= (lines[R].length - start)/10 + 1; ++C)
			set_text_arr(lines[R].slice(start+(C-1)*10,start+C*10).trim(),arr,R,C);
	}
	return arr;
}

function prn_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(prn_to_aoa(str, opts), opts); }

function prn_to_workbook(str/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(prn_to_sheet(str, opts), opts); }

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
function sc_to_aoa(str/*:string*/, opts)/*:AOA*/ {
	throw new Error('SC files unsupported');
}

function sc_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(sc_to_aoa(str, opts), opts); }

function sc_to_workbook(str/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(sc_to_sheet(str, opts), opts); }

/* Code Pages Supported by Visual FoxPro */
var dbf_codepage_map = {
	/*::[*/0x01/*::]*/:   437,
	/*::[*/0x02/*::]*/:   850,
	/*::[*/0x03/*::]*/:  1252,
	/*::[*/0x04/*::]*/: 10000,
	/*::[*/0x64/*::]*/:   852,
	/*::[*/0x65/*::]*/:   866,
	/*::[*/0x66/*::]*/:   865,
	/*::[*/0x67/*::]*/:   861,
	/*::[*/0x68/*::]*/:   895,
	/*::[*/0x69/*::]*/:   620,
	/*::[*/0x6A/*::]*/:   737,
	/*::[*/0x6B/*::]*/:   857,
	/*::[*/0x78/*::]*/:   950,
	/*::[*/0x79/*::]*/:   949,
	/*::[*/0x7A/*::]*/:   936,
	/*::[*/0x7B/*::]*/:   932,
	/*::[*/0x7C/*::]*/:   874,
	/*::[*/0x7D/*::]*/:  1255,
	/*::[*/0x7E/*::]*/:  1256,
	/*::[*/0x96/*::]*/: 10007,
	/*::[*/0x97/*::]*/: 10029,
	/*::[*/0x98/*::]*/: 10006,
	/*::[*/0xC8/*::]*/:  1250,
	/*::[*/0xC9/*::]*/:  1251,
	/*::[*/0xCA/*::]*/:  1254,
	/*::[*/0xCB/*::]*/:  1253,

	/* shapefile DBF extension */
	/*::[*/0x00/*::]*/: 20127,
	/*::[*/0x08/*::]*/:   865,
	/*::[*/0x09/*::]*/:   437,
	/*::[*/0x0A/*::]*/:   850,
	/*::[*/0x0B/*::]*/:   437,
	/*::[*/0x0D/*::]*/:   437,
	/*::[*/0x0E/*::]*/:   850,
	/*::[*/0x0F/*::]*/:   437,
	/*::[*/0x10/*::]*/:   850,
	/*::[*/0x11/*::]*/:   437,
	/*::[*/0x12/*::]*/:   850,
	/*::[*/0x13/*::]*/:   932,
	/*::[*/0x14/*::]*/:   850,
	/*::[*/0x15/*::]*/:   437,
	/*::[*/0x16/*::]*/:   850,
	/*::[*/0x17/*::]*/:   865,
	/*::[*/0x18/*::]*/:   437,
	/*::[*/0x19/*::]*/:   437,
	/*::[*/0x1A/*::]*/:   850,
	/*::[*/0x1B/*::]*/:   437,
	/*::[*/0x1C/*::]*/:   863,
	/*::[*/0x1D/*::]*/:   850,
	/*::[*/0x1F/*::]*/:   852,
	/*::[*/0x22/*::]*/:   852,
	/*::[*/0x23/*::]*/:   852,
	/*::[*/0x24/*::]*/:   860,
	/*::[*/0x25/*::]*/:   850,
	/*::[*/0x26/*::]*/:   866,
	/*::[*/0x37/*::]*/:   850,
	/*::[*/0x40/*::]*/:   852,
	/*::[*/0x4D/*::]*/:   936,
	/*::[*/0x4E/*::]*/:   949,
	/*::[*/0x4F/*::]*/:   950,
	/*::[*/0x50/*::]*/:   874,
	/*::[*/0x57/*::]*/:  1252,
	/*::[*/0x58/*::]*/:  1252,
	/*::[*/0x59/*::]*/:  1252,

	/*::[*/0xFF/*::]*/: 16969
};

/* TODO: find an actual specification */
function dbf_to_aoa(buf, opts)/*:AOA*/ {
	var out/*:AOA*/ = [];
	/* TODO: browser based */
	if(typeof Buffer === 'undefined') throw new Error("Buffer support required for DBF");
	var d = Buffer.isBuffer(buf) ? buf : new Buffer(buf, 'binary'), l = 0;

	/* header */
	var ft = d[l++];
	var memo = false;
	var vfp = false;
	switch(ft) {
		case 0x03: break;
		case 0x30: vfp = true; memo = true; break;
		case 0x31: vfp = true; break;
		case 0x83: memo = true; break;
		case 0x8B: memo = true; break;
		case 0xF5: memo = true; break;
		default: process.exit(); throw new Error("DBF Unsupported Version: " + ft.toString(16));
	}
	var filedate = new Date(d[l] + 1900, d[l+1] - 1, d[l+2]); l+=3;
	var nrow = d.readUInt32LE(l); l+=4;
	var fpos = d.readUInt16LE(l); l+=2;
	var rlen = d.readUInt16LE(l); l+=2;
	l+=16;

	var flags = d[l++];
	//if(memo && ((flags & 0x02) === 0)) throw new Error("DBF Flags " + flags.toString(16) + " ft " + ft.toString(16))i;

	/* codepage present in FoxPro */
	var current_cp = 1252;
	if(d[l] !== 0) current_cp = dbf_codepage_map[d[l]];
	l+=1;

	l+=2;
	var fields = [], field = {};
	var hend = fpos - 10 - (vfp ? 264 : 0);
	while(l < hend) {
		field = {};
		field.name = d.slice(l, l+11).toString().replace(/[\u0000\r\n].*$/g,"");
		field.type = String.fromCharCode(d[l+11]);
		field.offset = d.readUInt32LE(l+12);
		field.len = d[l+16];
		field.dec = d[l+17];
		if(field.name.length) fields.push(field);
		l += 32;
		switch(field.type) {
			// case 'B': break; // Binary
			case 'C': break; // character
			case 'D': break; // date
			case 'F': break; // floating point
			// case 'G': break; // General
			case 'I': break; // long
			case 'L': break; // boolean
			case 'M': break; // memo
			case 'N': break; // number
			// case 'O': break; // double
			// case 'P': break; // Picture
			case 'T': break; // datetime
			case 'Y': break; // currency
			case '0': break; // null ?
			case '+': break; // autoincrement
			case '@': break; // timestamp
			default: throw new Error('Unknown Field Type: ' + field.type);
		}
	}
	if(d[l] !== 0x0D) l = fpos-1;
	if(d[l++] !== 0x0D) throw new Error("DBF Terminator not found " + l + " " + d[l]);
	l = fpos;
	/* data */
	var R = 0, C = 0;
	out[0] = [];
	for(C = 0; C != fields.length; ++C) out[0][C] = fields[C].name;
	while(nrow-- > 0) {
		if(d[l] === 0x2A) { l+=rlen; continue; }
		++l;
		out[++R] = []; C = 0;
		for(C = 0; C != fields.length; ++C) {
			var dd = d.slice(l, l+fields[C].len); l+=fields[C].len;
			var s = cptable.utils.decode(current_cp, dd);//dd.toString();
			switch(fields[C].type) {
				case 'C':
					out[R][C] = cptable.utils.decode(current_cp, dd);
					out[R][C] = out[R][C].trim();
					break;
				case 'D':
					if(s.length === 8) out[R][C] = new Date(+s.substr(0,4), +s.substr(4,2)-1, +s.substr(6,2));
					else out[R][C] = s;
					break;
				case 'F': out[R][C] = parseFloat(s.trim()); break;
				case 'I': out[R][C] = dd.readInt32LE(0); break;
				case 'L': switch(s.toUpperCase()) {
					case 'Y': case 'T': out[R][C] = true; break;
					case 'N': case 'F': out[R][C] = false; break;
					case ' ': case '?': out[R][C] = false; break; /* NOTE: technically unitialized */
					default: throw new Error("DBF Unrecognized L:|" + s + "|");
					} break;
				case 'M': /* TODO: handle memo files */
					if(!memo) throw new Error("DBF Unexpected MEMO for type " + ft.toString(16));
					out[R][C] = "##MEMO##" + dd.readUInt32LE(0);
					break;
				case 'N': out[R][C] = +s.replace(/\u0000/g,"").trim(); break;
				case 'T':
					var day = dd.readUInt32LE(0), ms = dd.readUInt32LE(4);
					throw new Error(day + " | " + ms);
					//out[R][C] = new Date(); // FIXME!!!
					//break;
				case 'Y': out[R][C] = dd.readInt32LE(0)/1e4; break;
				case '0':
					if(fields[C].name === '_NullFlags') break;
					/* falls through */
				default: throw new Error("DBF Unsupported data type " + fields[C].type);
			}
		}
	}
	if(l < d.length && d[l++] != 0x1A) throw new Error("DBF EOF Marker missing " + (l-1) + " of " + d.length + " " + d[l-1].toString(16));
	return out;
}

function dbf_to_sheet(buf, opts)/*:Worksheet*/ {
	var o = opts || {};
	if(!o.dateNF) o.dateNF = "yyyymmdd";
	return aoa_to_sheet(dbf_to_aoa(buf, o), o);
}

function dbf_to_workbook(buf, opts)/*:Workbook*/ { return sheet_to_workbook(dbf_to_sheet(buf, opts), opts); }

function read_str(f/*:string*/, opts/*:any*/, hint/*:?any*/)/*:Workbook*/ {
	if(hint === 0xFFFE) return csv_to_workbook(f);
	if(f.substr(0, 5) === 'TABLE' && f.substr(0, 12).indexOf('0,1') > -1) return dif_to_workbook(f);
	else if(f.substr(0,2) === 'ID') return sylk_to_workbook(f);
	else if(f.charCodeAt(2) <= 12 && f.charCodeAt(3) <= 31) return dbf_to_workbook(new Buffer(f, 'binary'));
	else if(f.substr(0,19) === 'socialcalc:version:') return socialcalc_to_workbook(f);
	else if(f.substr(0,61) == '# This data file was generated by the Spreadsheet Calculator.') return sc_to_workbook(f);
	else if(f.indexOf(',') != -1 || f.indexOf('\t') != -1) return csv_to_workbook(f);
	else return prn_to_workbook(f);
}

function read_buf(f/*:RawData*/, opts/*:?ParseOpts*/, hint/*:?any*/)/*:Workbook*/ {
	if(f[2] <= 12 && f[3] <= 31) return dbf_to_workbook(f);
	return read_str(f.toString('binary'), opts);
}

function read(f/*:RawData*/, opts/*:?ParseOpts*/, hint/*:?any*/)/*:Workbook*/ {
	if(Buffer.isBuffer(f)) return read_buf(f, opts, hint);
	return read_str(f, opts, hint);
}

var readFile = function (f/*:string*/, o/*:?ParseOpts*/)/*:Workbook*/ {
	var b/*:Buffer*/ = fs.readFileSync(f);
/*
	if(b.length === 0) return null;
*/
	if(((b[0]<<8)|b[1])==0xFFFE) return read_str(cptable.utils.decode(1200, b.slice(2)), o, 0xFFFE);
	return read_buf(b, o);
};
function decode_row(rowstr/*:string*/)/*:number*/ { return parseInt(unfix_row(rowstr),10) - 1; }
function encode_row(row/*:number*/)/*:string*/ { return "" + (row + 1); }
function fix_row(cstr/*:string*/)/*:string*/ { return cstr.replace(/([A-Z]|^)(\d+)$/,"$1$$$2"); }
function unfix_row(cstr/*:string*/)/*:string*/ { return cstr.replace(/\$(\d+)$/,"$1"); }

function decode_col(colstr/*:string*/)/*:number*/ { var c = unfix_col(colstr), d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function encode_col(col/*:number*/)/*:string*/ { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }
function fix_col(cstr/*:string*/)/*:string*/ { return cstr.replace(/^([A-Z])/,"$$$1"); }
function unfix_col(cstr/*:string*/)/*:string*/ { return cstr.replace(/^\$([A-Z])/,"$1"); }

function split_cell(cstr/*:string*/)/*:Array<string>*/ { return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/,"$1,$2").split(","); }
function decode_cell(cstr/*:string*/)/*:CellAddress*/ { var splt = split_cell(cstr); return ({ c:decode_col(splt[0]), r:decode_row(splt[1]) }/*:any*/); }
function encode_cell(cell/*:CellAddress*/)/*:string*/ { return encode_col(cell.c) + encode_row(cell.r); }
function decode_range(range/*:string*/)/*:Range*/ { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
/*# if only one arg, it is assumed to be a Range.  If 2 args, both are cell addresses */
function encode_range(cs/*:CellAddrSpec|Range*/,ce/*:?CellAddrSpec*/)/*:string*/ {
	if(typeof ce === 'undefined' || typeof ce === 'number') {
/*:: if(!(cs instanceof Range)) throw "unreachable"; */
		return encode_range(cs.s, cs.e);
	}
/*:: if((cs instanceof Range)) throw "unreachable"; */
	if(typeof cs !== 'string') cs = encode_cell((cs/*:any*/));
	if(typeof ce !== 'string') ce = encode_cell((ce/*:any*/));
/*:: if(typeof cs !== 'string') throw "unreachable"; */
/*:: if(typeof ce !== 'string') throw "unreachable"; */
	return cs == ce ? cs : cs + ":" + ce;
}

var utils = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: decode_col,
	decode_row: decode_row,
	decode_cell: decode_cell,
	decode_range: decode_range,
	sheet_to_dif: sheet_to_dif,
	sheet_to_sylk: sheet_to_sylk,
	sheet_to_socialcalc: sheet_to_socialcalc
};
HARB.read = read;
HARB.readFile = readFile;
HARB.utils = utils;
})(typeof exports !== 'undefined' ? exports : HARB);
