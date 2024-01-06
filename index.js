import * as fs from "node:fs";
import * as path from "path";
import * as xlsx from "xlsx";
import * as axios from "axios";

const downloadFolderName = "COMPROVANTES";

export async function processXlsxFiles(rootDirectory) {
  const files = getAllXlsxFiles(rootDirectory);

  for (const file of files) {
    const filePath = path.join(rootDirectory, file);
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet

    // Extract values from the specified columns
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
    });

    const columnIndex = sheetData[0]?.indexOf("Comprovante Transfeera");
    const favorecidoIndex = sheetData[0]?.indexOf("Favorecido");
    const nomeLoteIndex = sheetData[0]?.indexOf("Nome do lote");
    const valorIndex = sheetData[0]?.indexOf("Valor");

    if (
      columnIndex !== -1 &&
      favorecidoIndex !== -1 &&
      nomeLoteIndex !== -1 &&
      valorIndex !== -1
    ) {
      const downloadFolderPath = path.join(rootDirectory, downloadFolderName);
      createDirectoryIfNotExists(downloadFolderPath);

      console.log(`Downloading PDF files from file "${file}"...`);
      for (let i = 1; i < sheetData.length; i++) {
        const cellValue = sheetData[i][columnIndex];

        // Extract values for renaming the file
        const favorecido = sheetData[i][favorecidoIndex];
        const nomeLote = sheetData[i][nomeLoteIndex];
        const valor = sheetData[i][valorIndex];

        // Generate a filename based on the extracted values
        const filename = `PGM ${favorecido} ${nomeLote} (${valor}).pdf`;

        // Download the PDF file
        try {
          const response = await axios.get(cellValue, {
            responseType: "arraybuffer",
          });
          const pdfPath = path.join(downloadFolderPath, filename);

          // Save the PDF file
          fs.writeFileSync(pdfPath, Buffer.from(response.data));
          console.log(`Downloaded and saved: ${pdfPath}`);
        } catch (error) {
          console.error(
            `Error downloading PDF from ${cellValue}: ${error.message}`
          );
        }
      }
    } else {
      console.log(
        `Skipping file "${file}" - Missing one or more required columns.`
      );
    }
  }
}

function getAllXlsxFiles(directory) {
  let files = [];

  const entries = fs.readdirSync(directory, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(directory, entry.name);

    if (entry.isDirectory()) {
      // Recursively process subdirectories
      files = files.concat(getAllXlsxFiles(fullPath));
    } else if (entry.isFile() && path.extname(entry.name) === ".xlsx") {
      files.push(entry.name);
    }
  }

  return files;
}

function createDirectoryIfNotExists(directory) {
  if (!fs.existsSync(directory)) {
    fs.mkdirSync(directory);
  }
}

if (process.argv.length !== 3) {
  console.error("Usage: node app.js <directory>");
  process.exit(1);
}

const directoryPath = process.argv[2];

processXlsxFiles(directoryPath);
