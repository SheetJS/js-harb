var sc_to_sheet = function(f) { throw new Error('SC files unsupported'); };

var sc_to_workbook = function (str) { return sheet_to_workbook(sc_to_sheet(str)); };

