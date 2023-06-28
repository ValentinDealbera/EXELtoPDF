const express = require("express");
const readExcel = require("read-excel-file/node");
const { PDFDocument, rgb } = require('pdf-lib');
const fs = require('fs');
const multer = require('multer');
const path = require('path');
const archiver = require('archiver');

const server = express();
const upload = multer({ dest: 'uploads/' });

server.use(express.json());
server.use(express.urlencoded({ extended: true }));

async function fillPDFTemplate(templatePath, data) {
  const templateBytes = await fs.promises.readFile(templatePath);

  const pdfDoc = await PDFDocument.load(templateBytes);
  const pages = pdfDoc.getPages();
  const firstPage = pages[0];

  firstPage.drawText(data.nombre, { x: 135, y: 778, size: 11, color: rgb(0, 0, 0) });
  firstPage.drawText(data.apellido, { x: 135, y: 753, size: 11, color: rgb(0, 0, 0) });
  firstPage.drawText(("$ " + data.sueldoBruto), { x: 175, y: 728, size: 11, color: rgb(0, 0, 0) });
  firstPage.drawText((data.dni + ""), { x: 119, y: 705, size: 11, color: rgb(0, 0, 0) });
  firstPage.drawText(data.cuil, { x: 120, y: 680, size: 11, color: rgb(0, 0, 0) });
  firstPage.drawText(data.email, { x: 130, y: 655, size: 11, color: rgb(0, 0, 0) });

  const modifiedPdfBytes = await pdfDoc.save();
  const fileName = `output/${data.nombre}_${data.apellido}_${data.dni}.pdf`
  fs.writeFileSync(fileName, modifiedPdfBytes);
  console.log('PDF generated successfully!');
}
server.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// server.post("/convert", upload.single('PDF'), async (req, res) => {
//   const zipFileName = 'boletas.zip';
//   try {
//     const arr = [];
//     const files = req.file;
//     // const data = await readExcel("pdf example.xlsx");
//     const data = await readExcel(files.path, { schema: null });
//     let user = {};
//     for (let i in data) {
//       user = {
//         nombre: data[i][0],
//         apellido: data[i][1],
//         sueldoBruto: data[i][2],
//         dni: data[i][3],
//         cuil: data[i][4],
//         email: data[i][5],
//       };
//       fillPDFTemplate('Template.pdf', user);
//       const filePath = path.resolve(__dirname, 'output', `${user.nombre}_${user.apellido}_${user.dni}.pdf`);
//       arr.push(filePath);
//       for (let j in data[i]) {
//       }
//     }
    
//     const zip = new AdmZip();
//     arr.forEach((filePath) => {
//       zip.addLocalFile(filePath);
//     });
    
//     res.setHeader('Content-Disposition', `attachment; filename="${zipFileName}"`);
//     res.setHeader('Content-Type', 'application/zip');
//     const zipBuffer = zip.toBuffer();
//     console.log(zipBuffer);
//     res.status(200).send(zipBuffer);
//   } catch (error) {
//     console.log("An error occurred:", error);
//     res.sendStatus(500).end();
//   }
// });


server.post("/convert", upload.single('PDF'), async (req, res) => {
  const zipFileName = 'boletas.zip'; // Replace 'boletas' with the desired name
  try {
    const arr = [];
    const files = req.file;
    const data = await readExcel(files.path, { schema: null });
    let user = {};
    for (let i in data) {
      user = {
        nombre: data[i][0],
        apellido: data[i][1],
        sueldoBruto: data[i][2],
        dni: data[i][3],
        cuil: data[i][4],
        email: data[i][5],
      };
      fillPDFTemplate('Template.pdf', user);
      const filePath = path.resolve(__dirname, 'output', `${user.nombre}_${user.apellido}_${user.dni}.pdf`);
      arr.push(filePath);
      for (let j in data[i]) {
      }
    }

    const outputZipPath = path.resolve(__dirname, 'temp', zipFileName);
    const outputZipStream = fs.createWriteStream(outputZipPath);
    const archive = archiver('zip', {
      zlib: { level: 9 } // Compression level
    });

    outputZipStream.on('close', () => {
      console.log('Zip file created successfully');
      res.setHeader('Content-Disposition', `attachment; filename=${zipFileName}`);
      res.setHeader('Content-Type', 'application/zip');
      res.download(outputZipPath, zipFileName, (err) => {
        if (err) {
          console.error('An error occurred while sending the zip file:', err);
          res.sendStatus(500);
        }

        // Cleanup: Delete the temporary zip file
        fs.unlink(outputZipPath, (unlinkErr) => {
          if (unlinkErr) {
            console.error('An error occurred while deleting the temporary zip file:', unlinkErr);
          }
        });
      });
    });

    archive.on('error', (err) => {
      console.error('An error occurred while creating the zip file:', err);
      res.sendStatus(500);
    });

    archive.pipe(outputZipStream);

    arr.forEach((filePath) => {
      archive.file(filePath, { name: path.basename(filePath) });
    });

    archive.finalize();
  } catch (error) {
    console.log("An error occurred:", error);
    res.sendStatus(500).end();
  }
});


server.listen(3001, () => {
  console.log("Server up on port 3001");
});