/* vim: set ts=2: */
var X;
var modp = './';
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){X=require(modp);});});

var ex = [".csv",".tsv",".txt",".prn",".dif",".slk",".sc",".dbf",".socialcalc"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});

var parsefile = function(p) { return function() { if(fs.existsSync(p)) X.readFile(p); } };
var test_dbf_file = "./test_files/libreoffice/calc/dbf/numeric-field-with-zero-by-excel.dbf";

describe('write.*', function() {
	it.skip('should parse csv', parsefile('./test_files/write.csv'));
	it('should parse tsv', parsefile('./test_files/write.txt'));
	it('should parse dif', parsefile('./test_files/write.dif'));
	it('should parse slk', parsefile('./test_files/write.slk'));
	it('should parse dbf', parsefile(test_dbf_file));
	it('should parse prn', parsefile('./test_files/write.prn'));
	it.skip('should parse socialcalc', parsefile('./test_files/write.socialcalc'));
	it.skip('should parse sc', parsefile('./test_files/write.sc'));
});

describe('dif', function() {
	it('should roundtrip write.dif', function() {
		var wb1 = X.readFile('./test_files/write.dif');
		var wb2 = X.read(X.utils.sheet_to_dif(wb1.Sheets.Sheet1));
		var ws1 = wb1.Sheets.Sheet1, ws2 = wb2.Sheets.Sheet1;
		assert.equal(ws1['!ref'], ws2['!ref']);
		"A1;B1;C1;A2;B2;D2;A3;B3;C3;D3;A4;C4".split(";").forEach(function(s) {
			assert.equal(ws1[s].v, ws2[s].v);
			assert.equal(ws1[s].t, ws2[s].t);
		});
	});
});

describe('socialcalc', function() {
	it('should roundtrip bad characters', function() {
		var wb = X.read('c:\>dir,cls');
		var sc = X.utils.sheet_to_socialcalc(wb.Sheets.Sheet1);
		var wb2 = X.read(sc);
		assert.equal(wb.Sheets.Sheet1.A1.v, wb2.Sheets.Sheet1.A1.v);
	});
	it('should roundtrip data', function() {
		var wb = {
			SheetNames: ["Sheet1"],
			Sheets: {
				Sheet1: {
					A1: {t:"n",v:1},
					A2: {t:"n",v:1,f:"A1+1"},
					B1: {t:"b",v:true},
					B2: {t:"b",v:true,f:"B1"},
					C1: {t:"s",v:"sheetjs"},
					C2: {t:"s",v:"sheetjs",f:"C1"},
					'!ref': "A1:C2"
				}
			}
		};
		var sc = X.utils.sheet_to_socialcalc(wb.Sheets.Sheet1);
		var wb2 = X.read(sc);
		assert.equal(wb.Sheets.Sheet1.A1.v, wb2.Sheets.Sheet1.A1.v);
		assert.equal(wb.Sheets.Sheet1.A2.v, wb2.Sheets.Sheet1.A2.v);
		assert.equal(wb.Sheets.Sheet1.B1.v, wb2.Sheets.Sheet1.B1.v);
		assert.equal(wb.Sheets.Sheet1.B2.v, wb2.Sheets.Sheet1.B2.v);
		assert.equal(wb.Sheets.Sheet1.C1.v, wb2.Sheets.Sheet1.C1.v);
		assert.equal(wb.Sheets.Sheet1.C2.v, wb2.Sheets.Sheet1.C2.v);
	});
});
