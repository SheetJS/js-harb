var sylk_to_sheet = function(f) { throw new Error('SYLK files unsupported'); };

var sylk_to_workbook = function (str) { return sheet_to_workbook(sylk_to_sheet(str)); };

