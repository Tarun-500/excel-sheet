document.addEventListener("DOMContentLoaded", function () {
    if (typeof GC !== 'undefined' && GC.Spread.Sheets) {
        var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), { sheetCount: 1 });
        var sheet = spread.getActiveSheet();

        // Sample data to fill in the sheet
        sheet.setValue(0, 0, "Name");
        sheet.setValue(0, 1, "Age");
        sheet.setValue(1, 0, "John Doe");
        sheet.setValue(1, 1, 28);
        sheet.setValue(2, 0, "Jane Smith");
        sheet.setValue(2, 1, 34);

        // Set column width
        sheet.setColumnWidth(0, 150);
        sheet.setColumnWidth(1, 100);

        // Style header
        var headerStyle = new GC.Spread.Sheets.Style();
        headerStyle.backColor = "#4CAF50";
        headerStyle.foreColor = "#FFFFFF";
        headerStyle.font = "bold 14px Arial";
        headerStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
        headerStyle.vAlign = GC.Spread.Sheets.VerticalAlign.center;
        headerStyle.borderBottom = new GC.Spread.Sheets.LineBorder("#000000", GC.Spread.Sheets.LineStyle.thin);

        sheet.setStyle(0, 0, headerStyle);
        sheet.setStyle(0, 1, headerStyle);

        // Apply styles to cells
        var cellStyle = new GC.Spread.Sheets.Style();
        cellStyle.font = "12px Arial";
        cellStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.left;
        cellStyle.vAlign = GC.Spread.Sheets.VerticalAlign.center;

        for (var row = 1; row <= 2; row++) {
            for (var col = 0; col <= 1; col++) {
                sheet.setStyle(row, col, cellStyle);
            }
        }
    } else {
        console.error('GC.Spread.Sheets is not defined. Ensure the SpreadJS library is loaded correctly.');
    }
});
