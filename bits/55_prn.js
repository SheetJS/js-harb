var prn_to_sheet = function(f) { throw new Error('PRN files unsupported'); };

var prn_to_workbook = function (str) { return sheet_to_workbook(prn_to_sheet(str)); };

