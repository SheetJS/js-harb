/* harb.js (C) 2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint eqnull:true, funcscope:true */
var HARB = {};
(function make_harb (HARB) {
HARB.version = '0.0.2';
if (typeof exports !== 'undefined') {
	if (typeof module !== 'undefined' && module.exports) {
		babyParse = require('babyparse');
		cptable = require('./dist/cpexcel');
		ssf = require('ssf');
		fs = require('fs');
	}
}

function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

var sheet_to_workbook = function (sheet) { return {SheetNames: ['Sheet1'], Sheets: {Sheet1: sheet}};	};

function aoa_to_sheet(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = encode_cell({c:C,r:R});
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = ssf._table[14];
				cell.v = datenum(cell.v);
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
var csv_to_aoa = function (str) {
	var data = babyParse.parse(str, bpopts).data;
	for(var R=0; R != data.length; ++R) for(var C=0; C != data[R].length; ++C) {
		var d = data[R][C];
		if(d === '') data[R][C] = null;
		else if(d === 'TRUE') data[R][C] = true;
		else if(d === 'FALSE') data[R][C] = false;
		else if(typeof d === 'string') {
			var dt = new Date(d);
			if(dt.getDate() === dt.getDate()) data[R][C] = dt;
		}
	}
	return data;
};

var csv_to_sheet = function (str) { return aoa_to_sheet(csv_to_aoa(str)); };

var csv_to_workbook = function (str) { return sheet_to_workbook(csv_to_sheet(str)); };

var prn_to_sheet = function(f) { throw new Error('PRN files unsupported'); };

var prn_to_workbook = function (str) { return sheet_to_workbook(prn_to_sheet(str)); };

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

var sylk_to_sheet = function(f) { throw new Error('SYLK files unsupported'); };

var sylk_to_workbook = function (str) { return sheet_to_workbook(sylk_to_sheet(str)); };

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

var sc_to_sheet = function(f) { throw new Error('SC files unsupported'); };

var sc_to_workbook = function (str) { return sheet_to_workbook(sc_to_sheet(str)); };

var read = function (f, opts) {
	if(f.substr(0, 5) === 'TABLE' && f.substr(0, 12).indexOf('0,1') > -1) return dif_to_workbook(f);
	else if(f.substr(0,2) === 'ID') return sylk_to_workbook(f);
	else if(f.substr(0,19) === 'socialcalc:version:') return socialcalc_to_workbook(f);
	else if(f.substr(0,61) == '# This data file was generated by the Spreadsheet Calculator.') return sc_to_workbook(f);
	else if(f.indexOf(',') != -1 || f.indexOf('\t') != -1) return csv_to_workbook(f);
	else return prn_to_workbook(f);
};

var readFile = function (f, o) {
	var b = fs.readFileSync(f);
	if(((b[0]<<8)|b[1])==0xFFFE) return read(cptable.utils.decode(1200, b.slice(2)), o);
	return read(b.toString(), o);
};
function decode_row(rowstr) { return parseInt(rowstr,10) - 1; }
function encode_row(row) { return "" + (row + 1); }

function decode_col(colstr) { var c = colstr, d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function encode_col(col) { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }

function split_cell(cstr) { return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/,"$1,$2").split(","); }
function decode_cell(cstr) { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
function encode_cell(cell) { return encode_col(cell.c) + encode_row(cell.r); }
function encode_range(cs,ce) {
	if(typeof ce === 'undefined' || typeof ce === 'number') return encode_range(cs.s, cs.e);
	if(typeof cs !== 'string') cs = encode_cell(cs); if(typeof ce !== 'string') ce = encode_cell(ce);
	return cs == ce ? cs : cs + ":" + ce;
}

var utils = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: encode_col,
	decode_row: encode_row,
	decode_cell: encode_cell
};
HARB.read = read;
HARB.readFile = readFile;
HARB.utils = utils;
})(typeof exports !== 'undefined' ? exports : HARB);
