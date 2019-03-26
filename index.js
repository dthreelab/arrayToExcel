const Excel = require('exceljs');

exports.ExcelOutput = (JSONToExcel, {
  fontSize = 12,
  fontColor = '#000',
  imageColumns = [],
  imageFolderAddress = 'images',
  imageURL,
  imageAspectRatio = 1, // = width/height;
  excelFileName = `${Date.now()}${Math.random()}`,
}) => {
  // Create a new instance of a Workbook class
  const workbook = new Excel.Workbook();

  const excelHeaders = [];
  const excelHeadersMeta = [];
  const imageColumnExtension = '-image';

  const cellWidth = 3;

  workbook.views = [{
    x: 0,
    y: 0,
    width: 10000,
    height: 20000,
    firstSheet: 0,
    activeTab: 1,
    visibility: 'visible',
  }];

  const worksheet = workbook.addWorksheet('My Sheet');
  // const worksheet = workbook.getWorksheet('My Sheet');

  if (typeof JSONToExcel !== 'undefined' && typeof JSONToExcel[0] !== 'undefined') {
    // adding the headers to worksheet.
    Object.keys(JSONToExcel[0]).forEach((column) => {
      if (!excelHeadersMeta.includes(column)) {
        const header = {
          header: column,
          key: column,
          width: cellWidth * column.length,
        };
        excelHeadersMeta.push(column);
        excelHeaders.push(header);
      }
    });
    imageColumns.forEach((imageColumn) => {
      const column = `${imageColumn}${imageColumnExtension}`;
      if (!excelHeadersMeta.includes(column)) {
        const header = {
          header: column,
          key: column,
          width: cellWidth * imageColumn.length,
        };
        excelHeadersMeta.push(column);
        excelHeaders.push(header);
      }
    });
    worksheet.columns = excelHeaders;

    let rowNo = 0;
    // adding rows to the excel sheet.
    JSONToExcel.forEach((excelRow) => {
      const row = excelRow;
      rowNo += 1;
      worksheet.addRow(row);
      // fetching image details from the rows.
      imageColumns.forEach((imageCell) => {
        if (typeof row[imageCell] !== 'undefined' && row[imageCell] != null) {
          try {
            // console.log(`${imageFolderAddress}/${row[imageCell]}`);

            const imageId = workbook.addImage({
              filename: `${imageFolderAddress}/${row[imageCell]}`,
              extension: 'jpeg',
            });

            const colNo = excelHeadersMeta.indexOf(`${imageCell}${imageColumnExtension}`);

            // insert an image over part of B2:D6
            // console.log(colNo, rowNo, 'hola', `${imageCell}${imageColumnExtension}`);

            worksheet.addImage(imageId, {
              tl: {
                col: colNo,
                row: rowNo
              },
              br: {
                col: colNo + 1,
                row: rowNo + 1
              },
            });

            worksheet.getRow(rowNo + 1).height = 7.2 * cellWidth / imageAspectRatio * imageCell.length;
            // console.log('row height', rowNo, imageCell.length);
          } catch (error) {
            row[imageCell] = 'no image uploaded';
            console.log('error occured', error);
          }
        }
      });
      // worksheet.addRow(row);
    });
  }

  // write to a file
  workbook.xlsx.writeFile(`${imageFolderAddress}/${excelFileName}.xlsx`)
    .then(() => {
      // done
    });
  return `${excelFileName}.xlsx`;
};
