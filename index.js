const express = require("express");
const readExcel = require("read-excel-file/node");
const { PDFDocument, rgb, degrees } = require('pdf-lib');
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

  firstPage.drawText(data.A.split(' ').slice(0, 3).join(' '), { x: 120, y: 75, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.A.split(' ').slice(0, 3).join(' '), { x: 120, y: 445, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText(data.B + '', { x: 93, y: 60, size: 6.5, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.B + '', { x: 93, y: 430, size: 6.5, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText(data.C.toLocaleDateString("es-AR", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  }), { x: 120, y: 220, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.C.toLocaleDateString("es-AR", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  }), { x: 120, y: 590, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText(data.D + '', { x: 160, y: 175, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.D + '', { x: 160, y: 545, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText(data.E + '', { x: 160, y: 210, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.E + '', { x: 160, y: 580, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText(data.F + '', { x: 132, y: 325, size: 7, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.F + '', { x: 132, y: 695, size: 7, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.F + '', { x: 158, y: 325, size: 7, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.F + '', { x: 158, y: 695, size: 7, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText(data.K + '', { x: 72, y: 230, size: 6.5, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.K + '', { x: 72, y: 600, size: 6.7, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  //---------------------------
  firstPage.drawText('100', { x: 190, y: 29, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('SUELDO MENSUAL', { x: 190, y: 50, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.J + ',00', { x: 190, y: 178, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.L + ',00', { x: 190, y: 215, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('100', { x: 190, y: 399, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('SUELDO MENSUAL', { x: 190, y: 420, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.J + ',00', { x: 190, y: 548, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.L + ',00', { x: 190, y: 585, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText('301', { x: 205, y: 29, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('JUBILACION 11%', { x: 205, y: 50, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.M + ',00', { x: 205, y: 325, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('301', { x: 205, y: 399, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('JUBILACION 11%', { x: 205, y: 420, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.M + ',00', { x: 205, y: 695, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText('302', { x: 220, y: 29, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('LEY 19032 3%', { x: 220, y: 50, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.N + ',00', { x: 220, y: 325, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('302', { x: 220, y: 399, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('LEY 19032 3%', { x: 220, y: 420, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.N + ',00', { x: 220, y: 695, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText('303', { x: 235, y: 29, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('OBRA SOCIAL Y ANSSAL', { x: 235, y: 50, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.O + ',00', { x: 235, y: 325, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('303', { x: 235, y: 399, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('OBRA SOCIAL Y ANSSAL', { x: 235, y: 420, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.O + ',00', { x: 235, y: 695, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText('305', { x: 250, y: 29, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('CUOTA SINDICAL 2,5%', { x: 250, y: 50, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.P + ',00', { x: 250, y: 325, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('305', { x: 250, y: 399, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('CUOTA SINDICAL 2,5%', { x: 250, y: 420, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.P + ',00', { x: 250, y: 695, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText('980', { x: 265, y: 29, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('REDONDEO', { x: 265, y: 50, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.S.toFixed(2) + '', { x: 265, y: 280, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('980', { x: 265, y: 399, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText('REDONDEO', { x: 265, y: 420, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.S.toFixed(2) + '', { x: 265, y: 650, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  //---------------------------
  firstPage.drawText(data.Q.toLocaleDateString("es-AR", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  }) + '', { x: 483, y: 160, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.Q.toLocaleDateString("es-AR", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  }) + '', { x: 483, y: 550, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText(data.R + ',00', { x: 483, y: 215, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.R + ',00', { x: 483, y: 585, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
   //---------------------------
   firstPage.drawText(data.S.toFixed(2) + '', { x: 483, y: 280, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
   firstPage.drawText(data.S.toFixed(2) + '', { x: 483, y: 650, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
   //---------------------------
   firstPage.drawText(data.T.toFixed(2) + '', { x: 483, y: 325, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
   firstPage.drawText(data.T.toFixed(2) + '', { x: 483, y: 695, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  //---------------------------
  firstPage.drawText(data.U.toFixed(2) + '', { x: 523, y: 325, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});
  firstPage.drawText(data.U.toFixed(2) + '', { x: 523, y: 695, size: 8, color: rgb(0, 0, 0), rotate: degrees(90)});

  const modifiedPdfBytes = await pdfDoc.save();
  const fileName = `output/${data.A}_${data.B}.pdf`
  fs.writeFileSync(fileName, modifiedPdfBytes);
  console.log('PDF generated successfully!');
}
server.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

server.post("/convert", upload.single('PDF'), async (req, res) => {
  try {
    const zipFileName = 'boletas.zip';
    const arr = [];
    const files = req.file;
    const data = await readExcel(files.path, { schema: null });
    const formatData = data.filter((_, index) => index >= 4);
    const dataToPrint = formatData.map((e) => {
      return e.filter((element) => element !== null);
    });
    const printData = dataToPrint.map((e) => {
      const obj = {};
      e.forEach((element, index) => {
        obj[String.fromCharCode(65 + index)] = element;
      });
      if (obj.A !== 'NOMBRE') fillPDFTemplate('Recibo en blanco - Antonio Sansone.pdf', obj)
      const filePath = path.resolve(__dirname, 'output', `${obj.A}_${obj.B}.pdf`);
      arr.push(filePath);
      return obj;
    });

    const createZipFile = () => {
      const outputZipPath = path.resolve(__dirname, 'zips', zipFileName);
      const outputZipStream = fs.createWriteStream(outputZipPath);
      const archive = archiver('zip', {
        zlib: { level: 9 }
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
    };

    setTimeout(createZipFile, 3000); // Delay of 3 seconds (3000 milliseconds)

  } catch (error) {
    console.log("An error occurred:", error);
    res.sendStatus(500).json({ error: error.message });
  }
});

server.listen(3001, () => {
  console.log("Server up on port 3001");
});