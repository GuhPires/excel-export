// Using Excel.js lib: https://github.com/exceljs/exceljs
const Workbook = require('exceljs').Workbook;

// This 'FileExport' class should be able to use all Excel.js 'Workbook' class methods
export default class FileExport extends Workbook {
	constructor(fileType: 'xlsx' | 'csv' = 'xlsx') {
		super();
		this.fileType = fileType;
	}

	// Initialization method
	init(
		wsName: string,
		columns: string[],
		sizes: number | number[] = 15
	): Iterator<object[]> | any {
		// Checking if the method has been called for this instance
		if (this.wb instanceof Workbook) {
			console.error(
				'ERROR! Cannot call the "init()" method more than once!\nIf you want to create a new Worksheet, please use "addWs()" method.'
			);
			return null;
		}
		this.wb = new Workbook();
		// Returning the Worksheet, Workbook and the instance of this class
		return [this.addWs(wsName, columns, sizes), this.wb, this];
	}

	// Creating Workshet for the current Workbook
	addWs(
		name: string,
		columns: string[],
		sizes: number | number[] = 15
	): object {
		const ws = this.wb.addWorksheet(name);
		// Creating the column headers
		ws.columns = columns.map((c: string, i: number) => ({
			header: c,
			key: c.toLowerCase(),
			width: Array.isArray(sizes) ? sizes[i] : sizes,
		}));
		// Returning the newly created Worksheet
		return ws;
	}

	// Exporting current Workbook in the format specified into the contructor
	async exportWb(fileName: string, stream: object): Promise<string> {
		const finalName = `${fileName}.${this.fileType}`;
		// Change the 'write' method to 'writeFile' in order to write into a file instead of a stream.
		await this.wb[this.fileType].write(stream);
		// Returning the name of the exported file with the extension
		return finalName;
	}
}
