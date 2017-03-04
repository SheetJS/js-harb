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

