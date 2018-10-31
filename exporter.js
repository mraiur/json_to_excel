var xlsx            = require('xlsx');

var exporter = {
    singleWorkbook: function(config, rows, filename ){
        var sheetName = config.sheetName || "Sheet 1";
        var letterDataMap = config.letterDataMap;

        function Workbook() {
            if (!(this instanceof Workbook)) return new Workbook();
            this.SheetNames = [];
            this.Sheets = {};
        }
        var wb = new Workbook();

        var ws = {};
        var range = {s: {c: 0, r: 0}, e: {c: 1, r: 1}};

        var col = 0;
        for (var pos in config.headers) {

            if (range.e.c < col) range.e.c = col + 2;

            ws[pos] = {
                v: config.headers[pos],
                t: 's'
            };
            col++;
        }

        for( var i = 0; i < rows.length; i++)
        {
            if (range.e.r < i) range.e.r = i + 2;

            for( var letter in letterDataMap )
            {
                var cell_ref = letter + (i + 2);
                ws[cell_ref] = {
                    v: '' + rows[i][ letterDataMap[letter] ],
                    t: 's'
                };
            }
        }

        if (range.s.c < 10000000) ws['!ref'] = xlsx.utils.encode_range(range);

        wb.SheetNames.push(sheetName);
        wb.Sheets[sheetName] = ws;

        xlsx.writeFile(wb, filename, {
            encoding: 'utf8'
        });
    }
};

module.exports = exporter;

