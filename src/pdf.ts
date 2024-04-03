

// import PDFDocument from 'pdfkit';
const PDFDocument = require('pdfkit') as any;
import fs from 'fs';
import students from './students.json';

function drawTable(doc: typeof PDFDocument, data: string[][], startX: number, startY: number, rowHeight: number, colWidths: number[], headerBorderColor: string, backgroundColor: string, textColor: string = "#000000") {
    let currentY = startY;
    const space = 1; // Space between cells

    data.forEach((row, rowIndex) => {
        // Check if adding another row exceeds the page height
        if (currentY + rowHeight > doc.page.height - doc.page.margins.bottom) {
            doc.addPage(); // Add a new page
            currentY = doc.page.margins.top; // Reset Y position for the new page
        }

        let currentX = startX;
        row.forEach((text, colIndex) => {
            const cellWidth = colWidths[colIndex] - space;
            const cellHeight = rowHeight - space;

            doc.rect(currentX, currentY, cellWidth, cellHeight)
                .fill(backgroundColor)
                .strokeColor(headerBorderColor)
                .stroke();

            doc.fillColor(textColor)
                .text(text, currentX + 2, currentY + 5, {
                    width: cellWidth - 4,
                    align: 'center',
                });

            currentX += colWidths[colIndex];
        });

        currentY += rowHeight;
    });
}

function createPDF() {
    const doc = new PDFDocument();
    doc.pipe(fs.createWriteStream('table.pdf'));

    // Table data and configuration
    const headers = [["Student Information", "Collected Points"]];
    const headers2 = [["ID", "Name", "Surname", "Phone", "Age", "Task 1", "Task 2", "Overall"]];

    // Define column widths and row height
    console.log(doc.page.width);

    const colWidths =  [30 + 130 + 130 + 119 + 35, 130];
    const colWidths2 = [30 , 130 , 130 , 119 , 35, 40, 40, 50];
    const rowHeight = 20;


    drawTable(doc, headers, (doc.page.width - 564) / 2, 30, rowHeight, colWidths, '#ffffff',"#34a853","#ffffff"); // Example with black ('#ffffff') as the header border color
    drawTable(doc, headers2, (doc.page.width - 564) / 2, 50.5, rowHeight, colWidths2, '#ffffff', "#34a853","#ffffff"); // Repeating for consistency

    // Then draw your data rows without specifying a header border color (or adjust as needed)



    const valuesList: string[][] = students.map(obj => Object.values(obj).map(value => value.toString()));

    console.log(valuesList);
    drawTable(doc, valuesList, (doc.page.width - 564) / 2, 71, rowHeight, colWidths2, '#ffffff',"#ffffff"); // The border color parameter won't affect data rows in current setup

    doc.end();
    console.log('PDF created successfully.');
}

createPDF();
