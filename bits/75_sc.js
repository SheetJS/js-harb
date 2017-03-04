function sc_to_aoa(str/*:string*/, opts)/*:AOA*/ {
	throw new Error('SC files unsupported');
}

function sc_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(sc_to_aoa(str, opts), opts); }

function sc_to_workbook(str/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(sc_to_sheet(str, opts), opts); }

