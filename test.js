/* vim: set ts=2: */
var X;
var modp = './';
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){X=require(modp);});});

var ex = [".csv",".tsv",".txt",".prn",".dif",".slk",".sc",".socialcalc"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});

describe('write.*', function() {
	it.skip('should parse csv', function() { X.readFile('./test_files/write.csv'); });
	it('should parse tsv', function() { X.readFile('./test_files/write.txt'); });
	it('should parse dif', function() { X.readFile('./test_files/write.dif'); });
	it('should parse slk', function() { X.readFile('./test_files/write.slk'); });
	it('should fail on prn', function() { assert.throws(function() { X.readFile('./test_files/write.prn'); }); });
	it.skip('should parse socialcalc', function() { X.readFile('./test_files/write.socialcalc'); });
	it.skip('should fail on sc', function() { assert.throws(function() { X.readFile('./test_files/write.sc'); }); });
});
