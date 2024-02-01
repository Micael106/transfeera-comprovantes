const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const axios = require("axios");

const downloadFolderName = "COMPROVANTES";

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

if (
  !fs.existsSync(directoryPath) ||
  !fs.statSync(directoryPath).isDirectory()
) {
  console.error("Invalid directory path. Please provide a valid directory.");
  process.exit(1);
}

fs.readdir(directoryPath, async (err, files) => {
  if (err) {
    console.error(`Error reading directory: ${err.message}`);
    process.exit(1);
  }

  // Filter only .xlsx files
  const xlsxFiles = files.filter((file) => path.extname(file) === ".xlsx");

  if (xlsxFiles.length === 0) {
    console.log(`No .xlsx files found in directory "${directoryPath}".`);
  } else {
    const devolvidoRows = [];
    const downloadFolderPath = path.join(directoryPath, downloadFolderName);
    createDirectoryIfNotExists(downloadFolderPath);

    for (const file of xlsxFiles) {
      const filePath = path.join(directoryPath, file);
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet

      // Extract values from the specified columns
      const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1,
      });

      const columnIndex = sheetData[0].indexOf("Comprovante Transfeera");
      const favorecidoIndex = sheetData[0].indexOf("Favorecido");
      const nomeLoteIndex = sheetData[0].indexOf("Nome do lote");
      const valorIndex = sheetData[0].indexOf("Valor");
      const statusIndex = sheetData[0].indexOf("Status");
      const cpfCnpjIndex = sheetData[0]?.indexOf("CPF ou CNPJ");
      const emailIndex = sheetData[0]?.indexOf("Email");

      if (
        columnIndex !== -1 &&
        favorecidoIndex !== -1 &&
        nomeLoteIndex !== -1 &&
        valorIndex !== -1
      ) {
        console.log(`Downloading PDF files from file "${file}"...`);
        for (let i = 1; i < sheetData.length; i++) {
          const status = sheetData[i][statusIndex];
          const cellValue = sheetData[i][columnIndex];
          const favorecido = sheetData[i][favorecidoIndex];
          const nomeLote = sheetData[i][nomeLoteIndex];
          const valor = sheetData[i][valorIndex];
          const email = sheetData[i][emailIndex];
          const cpfCnpj = sheetData[i][cpfCnpjIndex];

          if (status !== "FINALIZADO") {
            devolvidoRows.push([
              favorecido,
              cpfCnpj,
              email,
              "",
              "",
              "",
              "",
              "",
              valor,
              "",
              "",
              nomeLote,
            ]);
            continue; // Skip because there is not comprovante
          }

          // Generate a filename based on the extracted values
          const filename = `PGM ${favorecido} ${nomeLote} (${valor}).pdf`;

          // Download the PDF file
          try {
            const response = await axios.get(cellValue, {
              responseType: "arraybuffer",
            });
            const pdfPath = path.join(
              downloadFolderPath,
              file.split(".")[0],
              filename
            );
            createDirectoryIfNotExists(
              path.join(downloadFolderPath, file.split(".")[0])
            );
            //const pdfPath = path.join(downloadFolderPath, filename);

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
          `One or more required columns not found in file "${file}".`
        );
      }
    }

    if (devolvidoRows.length > 0) {
      const outputWorkbook = xlsx.utils.book_new();
      const headerRow = [
        [
          "Mantenha sempre o cabeçalho original da planilha e esta linha, mantendo os titulos e a ordem dos campos",
        ],
        [
          "Nome ou Razão Social",
          "CPF ou CNPJ",
          "Email (opcional)",
          "Banco",
          "Agência",
          "Conta",
          "Dígito da conta",
          "Tipo de Conta (Corrente ou Poupança)",
          "Valor",
          "ID integração (opcional)",
          "Data de agendamento (opcional)",
          "Descrição Pix (opcional)",
        ],
        ...devolvidoRows,
      ];
      const outputSheet = xlsx.utils.aoa_to_sheet(headerRow, {
        "!merges": [{ s: { r: 0, c: 0 }, e: { r: 0, c: 11 } }],
      });
      xlsx.utils.book_append_sheet(outputWorkbook, outputSheet, "Devolvidos");
      const outputFilePath = path.join(
        downloadFolderPath,
        "PGM DEVOLVIDOS.xlsx"
      );
      xlsx.writeFile(outputWorkbook, outputFilePath);
      console.log(`New Excel file created pgm devolvidos: ${outputFilePath}`);
    } else {
      console.log("No rows with 'DEVOLVIDO' status found.");
    }
  }
});
