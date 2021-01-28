// Using Excel.js lib: https://github.com/exceljs/exceljs
const Workbook = require('exceljs').Workbook;

// This 'ExportFile' class should be able to use all Excel.js 'Workbook' class methods
export default class ExportFile extends Workbook {
	constructor(fileType: 'xlsx' | 'csv' = 'xlsx') {
		super();
		this.fileType = fileType;
	}

	// Initialization method
	init(wsName: string, columns: string[]): Iterator<object[]> | any {
		// Checking if the method has been called for this instance
		if (this.wb instanceof Workbook) {
			console.error(
				'ERROR! Cannot call the "init()" method more than once!\nIf you want to create a new Worksheet, please use "addWs()" method.'
			);
			return null;
		}
		this.wb = new Workbook();
		// Returning the Worksheet, Workbook and the instance of this class
		return [this.addWs(wsName, columns), this.wb, this];
	}

	// Creating Workshet for the current Workbook
	addWs(name: string, columns: string[]): object {
		const ws = this.wb.addWorksheet(name);
		// Creating the column headers
		ws.columns = columns.map(c => ({
			header: c.toString(),
			key: c.toString().toLowerCase(),
		}));
		// Returning the newly created Worksheet
		return ws;
	}

	// Exporting current Workbook in the format specified into the contructor
	async exportWb(fileName: string): Promise<string> {
		const finalName = `${fileName}.${this.fileType}`;
		// Change the 'writeFile' method to 'write' in order to write in a stream.
		await this.wb[this.fileType].writeFile(finalName);
		// Returning the name of the exported file with the extension
		return finalName;
	}
}
