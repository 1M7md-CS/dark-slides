const { PDFDocument } = require('pdf-lib');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const { createCanvas } = require('canvas');
const fs = require('fs');

const theme = { r: 42, g: 37, b: 34 };

async function convertPdfToDark(inputPath, outputPath) {
    const data = new Uint8Array(fs.readFileSync(inputPath));
    const pdf = await pdfjsLib.getDocument({
        data,
        canvasFactory: {
            create: (w, h) => createCanvas(w, h)
        }
    }).promise;
    const newDoc = await PDFDocument.create();
    const totalPages = pdf.numPages;

    for (let i = 1; i <= totalPages; i++) {
        process.stdout.write(`\r  Page ${i}/${totalPages}`);
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 2 });
        const canvas = createCanvas(viewport.width, viewport.height);
        const ctx = canvas.getContext('2d');

        await page.render({ canvasContext: ctx, viewport }).promise;

        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const d = imageData.data;
        for (let j = 0; j < d.length; j += 4) {
            const brightness = 0.299 * d[j] + 0.587 * d[j + 1] + 0.114 * d[j + 2];
            const factor = 1 - (brightness / 255);
            d[j] = theme.r + (255 - theme.r) * factor;
            d[j + 1] = theme.g + (255 - theme.g) * factor;
            d[j + 2] = theme.b + (255 - theme.b) * factor;
        }
        ctx.putImageData(imageData, 0, 0);

        const pngBytes = canvas.toBuffer('image/png');
        const img = await newDoc.embedPng(pngBytes);
        const newPage = newDoc.addPage([viewport.width, viewport.height]);
        newPage.drawImage(img, { x: 0, y: 0, width: viewport.width, height: viewport.height });
    }

    process.stdout.write('\n');
    fs.writeFileSync(outputPath, await newDoc.save());
}

const [input, output] = process.argv.slice(2);
if (!input || !output) { console.error('Usage: node convert-dark.js <input.pdf> <output.pdf>'); process.exit(1); }
convertPdfToDark(input, output).then(() => console.log('  -> OK')).catch(e => { console.error('  -> ERROR:', e.message); process.exit(1); });
