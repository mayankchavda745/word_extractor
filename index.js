const express = require('express');
const app = express();

const mammoth = require('mammoth');
const fs = require('fs');

const row = 5;
const column = 3;

app.get('/readfile', (req, res) => {
    // Provide the path to your DOCX file
    const filePath = __dirname + '/test.docx';

    // Read the content of the DOCX file
    fs.readFile(filePath, 'utf8', (err, data) => {
        if (err) {
            console.error(err);
            res.status(500).send('Error reading file');
            return;
        }

        // Use mammoth to convert the DOCX content to HTML
        mammoth.extractRawText({ path: filePath })
            .then(result => {
                const docxContent = result.value;

                // Split the DOCX content by newline character
                const docxContentArray = docxContent.split('\n');
                const finalArr = docxContentArray.filter(d => d.length);

                // Send the array as the response
                // let ans;
                // for(const i of )
                res.json(finalArr);
            })
            .catch(error => {
                console.error(error);
                res.status(500).send('Error extracting text from DOCX');
            });
    });
});

app.get('/wTh', (req, res) => {
    const inputFilePath = 'test.docx'; // Replace with your Word file path
    const outputFilePath = 'output.html'; // Replace with your desired HTML file path

    // Read the Word file content
    fs.readFile(inputFilePath, 'utf8', (err, data) => {
        if (err) {
            res.send(err);
            return;
        }

        // Convert Word to HTML using mammoth
        mammoth.extractRawText({ arrayBuffer: data })
            .then(result => {
                const htmlContent = `
                <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <title>Your Document Title</title>
                </head>
                <body>
                    <p>${result.value}</p>
                </body>
                </html>
            `;
                console.log(htmlContent);
                // Write the HTML content to the output file
                fs.writeFile(outputFilePath, htmlContent, 'utf8', (err) => {
                    if (err) {
                        console.log(err);
                        res.status(200).send(err);
                    } else {
                        console.log('success');
                        res.status(200).send(`Conversion successful. HTML file saved at ${outputFilePath}`);
                    }
                });
            })
            .catch(error => {
                res.status(400).send(error);
            });
    });
});
app.listen(3001, () => {
    console.log('server started on 3001');
});