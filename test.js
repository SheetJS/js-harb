/* vim: set ts=2: */
var X;
var modp = './';
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){X=require(modp);});});

var ex = [".csv",".tsv",".txt",".prn",".dif",".slk",".sc",".dbf",".socialcalc"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});

var parsefile = function(p) { return function() { X.readFile(p); } };

describe('write.*', function() {
	it.skip('should parse csv', parsefile('./test_files/write.csv'));
	it('should parse tsv', parsefile('./test_files/write.txt'));
	it('should parse dif', parsefile('./test_files/write.dif'));
	it('should parse slk', parsefile('./test_files/write.slk'));
	it.skip('should parse dbf', parsefile('./test_files/write.dbf'));
	it('should parse prn', parsefile('./test_files/write.prn'));
	it.skip('should parse socialcalc', parsefile('./test_files/write.socialcalc'));
	it.skip('should parse sc', parsefile('./test_files/write.sc'));
});
