function checkSheetStats() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const stats = sheets.map(s => {
        const name = s.getName();
        const lastRow = s.getLastRow();
        const lastCol = s.getLastColumn();
        let sample = [];
        if (lastRow > 0 && lastCol > 0) {
            sample = s.getRange(1, 1, Math.min(lastRow, 2), lastCol).getValues();
        }
        return { name, lastRow, lastCol, sample };
    });
    Logger.log(JSON.stringify(stats, null, 2));
}
