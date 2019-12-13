const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const fs = require('fs');
const path = require('path');
const yamlFront = require('yaml-front-matter');

const dateFormat = require('dateformat');

function compare(a, b) {
  if (a.begin < b.begin) return 1;
  return -1;
}

fs.readFile(
  '/home/lyd/APPS-SCRIPTS/190907-developpeur-react-nord/src/markdown-pages/cv.md',
  'utf8',
  (err, data) => {
    if (err) throw err;

    //Load the docx file as a binary
    var content = fs.readFileSync(
      path.resolve(__dirname, 'input.docx'),
      'binary'
    );

    var zip = new PizZip(content);

    var doc = new Docxtemplater();
    doc.loadZip(zip);

    const { description, objectif, skills, experience } = yamlFront.loadFront(
      data
    );

    doc.setData({
      description,
      objectif,
      mainSkills: skills.main
        .map(skill => `${skill.title}: ${skill.rate}/10`)
        .join(', '),
      otherSkills: skills.other
        .map(skill => `${skill.title}: ${skill.rate}/10`)
        .join(', '),
      experience: experience
        .map(
          xp =>
            `${dateFormat(xp.begin, 'mm/yyyy')} à ${dateFormat(
              xp.end,
              'mm/yyyy'
            )} ${xp.job} ${xp.company}`
        )
        .join('; '),
      experiences: experience.sort(compare).map(xp => {
        return {
          begin: dateFormat(xp.begin, 'mm/yyyy'),
          end: dateFormat(xp.end, 'mm/yyyy'),
          company: xp.company,
          job: xp.job
        };
      })
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
