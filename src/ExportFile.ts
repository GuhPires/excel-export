// Using Excel.js lib: https://github.com/exceljs/exceljs
const Workbook = require('exceljs').Workbook;

// This 'ExportFile' class should be able to use all Excel.js 'Workbook' class methods
export default class ExportFile extends Workbook {
	constructor(fileType: 'xlsx' | 'csv' = 'xlsx') {
		super();
		this.fileType = fileType;
	}

	init(wsName: string, columns: string[]): Iterator<object[]> | any {
		if (this.wb instanceof Workbook) {
			console.error(
				'ERROR! Cannot call the "init()" method more than once!\nIf you want to create a new Worksheet, please use "addWs()" method.'
			);
			return null;
		}
		this.wb = new Workbook();
		return [this.addWs(wsName, columns), this.wb, this];
	}

	addWs(name: string, columns: string[]): object {
		const ws = this.wb.addWorksheet(name);
		ws.columns = columns.map(c => ({
			header: c.toString(),
			key: c.toString().toLowerCase(),
		}));
		return ws;
	}

	async exportWb(fileName: string): Promise<string> {
		const finalName = `${fileName}.${this.fileType}`;
		// Change the 'writeFile' to 'write' in order to write in a stream.
		await this.wb[this.fileType].writeFile(finalName);
		return finalName;
	}
}
