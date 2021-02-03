import fs from 'fs';
import FileExport from './FileExport';

// Creating titles for columns
const columns: string[] = ['A', 'B', 'C', 'D', 'E', 'F'];

// Initializing and returning Worksheet, Workbook and instance of FileExport Class
const [ws, wb, wbFile] = new FileExport('xlsx').init('Test 1', columns, 20);

// Adding a row to the new Worksheet. Note that this method comes directly from the
// Excel.js 'Workbook' class (Worksheet methods)
ws.addRow([1, 2, 3, 4, 5, 6]);

// Cannot use 'intit' method more than once, it should give an error into the console
// and return 'null'
const shouldBeNull = wbFile.init('Test 2', ['T1', 'T2']);
console.log('SHOULD BE NULL: ', shouldBeNull);

// Creating a new Worksheet with a name and column titles
const ws2 = wbFile.addWs('Test 2', ['T1', 'T2'], 50);
// Adding a row to the second Worksheet
ws2.addRow([9, 8]);

// Getting all Worksheets from the current Workbook
wb.eachSheet((ws: { name: string }, id: number) =>
	console.log('Worksheet: ', ws.name, '\nID: ', id)
);

// Exporting the current Workbook as the specified format when creating a FileExport instance
(async () => {
	const stream = fs.createWriteStream('test.xlsx');
	const fileName = await wbFile.exportWb('SomeFile', stream);
	console.log('FILE NAME:', fileName);
})();
