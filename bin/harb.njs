#!/usr/bin/env node

var HARB = require('../');

console.log(HARB.readFile(process.argv[2]).Sheets.Sheet1);
