function set_text_arr(data/*:string*/, arr/*:Array<Array<any> >*/, R/*:number*/, C/*:number*/) {
	if(data === 'TRUE') arr[R][C] = true;
	else if(data === 'FALSE') arr[R][C] = false;
	else if(+data == +data) arr[R][C] = +data;
	else if(data !== "") arr[R][C] = data;
}

var prn_to_aoa = function(f) {
	var arr = [];
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
};

var prn_to_sheet = function (str) { return aoa_to_sheet(prn_to_aoa(str)); };

var prn_to_workbook = function (str) { return sheet_to_workbook(prn_to_sheet(str)); };

