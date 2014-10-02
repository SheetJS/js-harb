# harb 

Host of Archaic Representations of Books (parsing).  Designed to provide support
for [j](https://github.com/SheetJS/j).  Pure-JS cleanroom implementation.

Currently supported formats:

- DIF (Data Interchange Format)
- CSV/TSV/TXT
- SYLK (Symbolic Link)
- SocialCalc

Planned but not currently implemented:

- SC (Spreadsheet Calculator)
- PRN (Space-Delimited Format)

## Installation

In [nodejs](https://www.npmjs.org/package/harb):

    npm install harb 

## Usage

This module houses provides support for [j](https://www.npmjs.org/package/j). See
[the xlsx module](http://git.io/xlsx) for usage information, as they use the same
interface and style. 

## Test Files

Test files are housed in [another repo](https://github.com/SheetJS/test_files).

Running `make init` will refresh the `test_files` submodule and get the files.

## Contributing

Due to the precarious nature of the Open Specifications Promise, it is very
important to ensure code is cleanroom.  Consult CONTRIBUTING.md

## XLSX Support

XLSX is available in [js-xlsx](http://git.io/xlsx).

## XLS Support

XLS is available in [js-xls](http://git.io/xls).

## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 license are reserved by the Original Author.

It is the opinion of the Original Author that this code conforms to the terms of
the Microsoft Open Specifications Promise, falling under the same terms as
OpenOffice (which is governed by the Apache License v2).  Given the vagaries of
the promise, the Original Author makes no legal claim that in fact end users are
protected from future actions.  It is highly recommended that, for commercial
uses, you consult a lawyer before proceeding.

## Badges

[![Build Status](https://travis-ci.org/SheetJS/js-harb.svg?branch=master)](https://travis-ci.org/SheetJS/js-harb)

[![Coverage Status](http://img.shields.io/coveralls/SheetJS/js-harb/master.svg)](https://coveralls.io/r/SheetJS/js-harb?branch=master)

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/ed5bb2c4c4346a474fef270f847f3f78 "githalytics.com")](http://githalytics.com/SheetJS/js-xlsx)
