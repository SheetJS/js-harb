#!/bin/bash
# make_cp.sh -- make codepage script for js-harb 
# Copyright (C) 2016-present  SheetJS
# usage: make_cp.sh [path_to_codepage_repo]
# writes output file to <pwd>/misc/cpharb.js
P=`pwd`
cd ${1:-../js-codepage}
if [ ! -e make.sh ]; then make; fi
bash make.sh "$P"/misc/harb.csv cpharb.js cptable
cat cpharb.js cputils.js > "$P"/misc/cpharb.js
