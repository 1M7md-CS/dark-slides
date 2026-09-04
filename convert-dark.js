const { PDFDocument } = require('pdf-lib');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const { createCanvas } = require('canvas');
const fs = require('fs');

const theme = { r: 42, g: 37, b: 34 };
const SCALE = 2;

async function convertPdfToDark(inputPath, outputPath) {
    const data = new Uint8Array(fs.readFileSync(inputPath));

    const pdf = await pdfjsLib.getDocument({
        data,
        canvasFactory: {
            create: (width, height) => createCanvas(width, height)
        }
    }).promise;

    const newDoc = await PDFDocument.create();
    const totalPages = pdf.numPages;

    for (let i = 1; i <= totalPages; i++) {
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: SCALE });

        const canvas = createCanvas(
            Math.ceil(viewport.width),
            Math.ceil(viewport.height)
        );

        const ctx = canvas.getContext('2d', {
            alpha: false
        });

        await page.render({
            canvasContext: ctx,
            viewport
        }).promise;

        const imageData = ctx.getImageData(
            0,
            0,
            canvas.width,
            canvas.height
        );

        const pixels = imageData.data;
        const tr = theme.r;
        const tg = theme.g;
        const tb = theme.b;

        for (let j = 0; j < pixels.length; j += 4) {
            const r = pixels[j];
            const g = pixels[j + 1];
            const b = pixels[j + 2];

            const brightness =
                0.299 * r +
                0.587 * g +
                0.114 * b;

            const factor = 1 - brightness / 255;

            pixels[j] = tr + (255 - tr) * factor;
            pixels[j + 1] = tg + (255 - tg) * factor;
            pixels[j + 2] = tb + (255 - tb) * factor;
        }

        ctx.putImageData(imageData, 0, 0);

        const pngBytes = canvas.toBuffer('image/png');
        const image = await newDoc.embedPng(pngBytes);

        const newPage = newDoc.addPage([
            viewport.width,
            viewport.height
        ]);

        newPage.drawImage(image, {
            x: 0,
            y: 0,
            width: viewport.width,
            height: viewport.height
        });

        page.cleanup();
    }

    const pdfBytes = await newDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);
}

const [input, output] = process.argv.slice(2);

if (!input || !output) {
    process.exit(1);
}

convertPdfToDark(input, output).catch(() => {
    process.exit(1);
});