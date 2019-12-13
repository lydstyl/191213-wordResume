const fs = require('fs');
const path = require('path');
const yamlFront = require('yaml-front-matter');
const dateFormat = require('dateformat');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

function compare(a, b) {
  if (a.begin < b.begin) return 1;
  return -1;
}

function makeResume(markdownCV, input, output, anonyme = false) {
  fs.readFile(markdownCV, 'utf8', (err, data) => {
    if (err) throw err;

    const content = fs.readFileSync(path.resolve(__dirname, input), 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater();
    doc.loadZip(zip);

    const { description, objectif, skills, experience } = yamlFront.loadFront(
      data
    );

    doc.setData({
      description: anonyme
        ? description.replace('Gabriel Brun', '<nom anonyme>')
        : description,
      objectif,
      mainSkills: skills.main
        .map(skill => `${skill.title}: ${skill.rate}/10`)
        .join(', '),
      otherSkills: skills.other.map(skill => `${skill.title}`).join(', '),
      experiences: experience.sort(compare).map(xp => {
        return {
          begin: dateFormat(xp.begin, 'mm/yyyy'),
          end: dateFormat(xp.end, 'mm/yyyy'),
          company: anonyme ? '<entreprise anonyme>' : xp.company,
          job: xp.job,
          tasks: xp.body.split('\n').map(task => {
            return { task };
          })
        };
      })
    });

    try {
      doc.render(); // replace all occurences in document
    } catch (error) {
      const e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties
      };
      console.log(JSON.stringify({ error: e }));
      // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
      throw error;
    }

    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname, output), buf);
  });
}

const markdownCV =
  '/home/lyd/APPS-SCRIPTS/190907-developpeur-react-nord/src/markdown-pages/cv.md';

makeResume(markdownCV, 'cvTemplate.docx', 'CV_Gabriel_Brun.docx');

makeResume(markdownCV, 'anonymeTemplate.docx', 'CV_Anonyme.docx', true);
