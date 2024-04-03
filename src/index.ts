import Excel, { Workbook, Worksheet } from 'exceljs';
import students from './students.json';





class ExcelReportGenerator {
    private workbook: Workbook;
    private worksheet: Worksheet;

    constructor() {
        this.workbook = new Excel.Workbook();
        this.worksheet = this.workbook.addWorksheet('Sheet 1');
    }

    private applyHeaderStyles(): void {
        const headerStyle: Partial<Excel.Style> = {
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF34a853' },
                bgColor: { argb: "FFFFFFFF" }
            },
            border: {
                top: { style: 'thin', color: { argb: 'FF34a853' } },
                left: { style: 'thin', color: { argb: 'FF34a853' } },
                bottom: { style: 'thin', color: { argb: 'FF34a853' } },
                right: { style: 'thin', color: { argb: 'FF34a853' } },
            },
            alignment: {
                vertical: 'middle',
                horizontal: 'center',
            },
            font: {
                color: { argb: "FFFFFFFF" },
                bold: true
            }
        };

        // Apply styles for the merged headers
        const headers = ['A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'E1:E2', 'F1:H1'];
        headers.forEach(headerRange => {
            this.worksheet.getCell(headerRange.split(':')[0]).style = headerStyle;
        });

        // Main headers' values
        const values = ['ID (student)', 'Name', 'Surname', 'Phone', 'Age', 'Collected Points'];
        values.forEach((value, index) => {
            // For merged cells, set value to the first cell of the range
            this.worksheet.getCell(headers[index].split(':')[0]).value = value;
        });

        // Sub-headers for "Collected Points"
        const subHeaders = ['F2', 'G2', 'H2'];
        const subHeaderValues = ['Task 1', 'Task 2', 'Overall'];
        subHeaders.forEach((header, index) => {
            this.worksheet.getCell(header).style = headerStyle;
            this.worksheet.getCell(header).value = subHeaderValues[index];
        });
    }

    private mergeCells(): void {
        // Merging cells
        this.worksheet.mergeCells('A1:A2');
        this.worksheet.mergeCells('B1:B2');
        this.worksheet.mergeCells('C1:C2');
        this.worksheet.mergeCells('D1:D2');
        this.worksheet.mergeCells('E1:E2');
        this.worksheet.mergeCells('F1:H1');
    }

    private applyGlobalStyles(): void {
        // Apply global alignment
        this.worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.alignment = { horizontal: 'center',vertical: 'middle' };
            });
        });
    }

    private configureColumns(): void {
        this.worksheet.columns = [
            { key: 'id', width: 20 },
            { key: 'name', width: 32 },
            { key: 'surname', width: 10 },
            { key: 'phone', width: 15 },
            { key: 'age', width: 10 },
            { key: 'collected_points_task_1', width: 12 },
            { key: 'collected_points_task_2', width: 12 },
            { key: 'collected_points_overall', width: 12 },
        ];
    }

    public async createExcelFile(filename: string): Promise<void> {
        this.mergeCells();
        this.applyHeaderStyles();
        this.configureColumns();



        // students.forEach(
        //     (value,i) => {
        //         console.log(value);

        //         this.worksheet.addRow(value)
        //     }
        // )


        this.worksheet.addRows(students);






        this.applyGlobalStyles();

        await this.workbook.xlsx.writeFile(filename);
        console.log(`${filename} has been created.`);
    }
}

// Usage
const generator = new ExcelReportGenerator();
generator.createExcelFile('StudentsPoints.xlsx').catch(console.error);
