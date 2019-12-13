const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const fs = require('fs');
const path = require('path');
const yamlFront = require('yaml-front-matter');

fs.readFile(
  '/home/lyd/APPS-SCRIPTS/190907-developpeur-react-nord/src/markdown-pages/cv.md',
  'utf8',
  (err, data) => {
    if (err) throw err;
    console.log(yamlFront.loadFront(data));

    //Load the docx file as a binary
    var content = fs.readFileSync(
      path.resolve(__dirname, 'input.docx'),
      'binary'
    );

    var zip = new PizZip(content);

    var doc = new Docxtemplater();
    doc.loadZip(zip);

    const { description, objectif } = yamlFront.loadFront(data);

    //set the templateVariables
    doc.setData({
      // first_name: 'John',
      // last_name: 'Doe',
      // phone: '0652455478',
      // description: 'New Website',
      description,
      objectif
    });

    try {
      // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
      doc.render();
    } catch (error) {
      var e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties
      };
      console.log(JSON.stringify({ error: e }));
      // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
      throw error;
    }

    var buf = doc.getZip().generate({ type: 'nodebuffer' });

    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);
  }
);
