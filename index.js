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

    //NORMAL RESUME
    const content = fs.readFileSync(
      path.resolve(__dirname, 'input.docx'),
      'binary'
    );

    const zip = new PizZip(content);

    const doc = new Docxtemplater();
    doc.loadZip(zip);

    let { description, objectif, skills, experience } = yamlFront.loadFront(
      data
    );

    doc.setData({
      description,
      objectif,
      mainSkills: skills.main
        .map(skill => `${skill.title}: ${skill.rate}/10`)
        .join(', '),
      otherSkills: skills.other.map(skill => `${skill.title}`).join(', '),
      experiences: experience.sort(compare).map(xp => {
        return {
          begin: dateFormat(xp.begin, 'mm/yyyy'),
          end: dateFormat(xp.end, 'mm/yyyy'),
          company: xp.company,
          job: xp.job,
          tasks: xp.body.split('\n').map(task => {
            return { task };
          })
        };
      })
    });

    //ANONYME RESUME
    const anonymeContent = fs.readFileSync(
      path.resolve(__dirname, 'anonymeTemplate.docx'),
      'binary'
    );

    const anonymeZip = new PizZip(anonymeContent);

    const anonymeDoc = new Docxtemplater();
    anonymeDoc.loadZip(anonymeZip);

    description = yamlFront.loadFront(data).description;
    objectif = yamlFront.loadFront(data).objectif;
    skills = yamlFront.loadFront(data).skills;
    experience = yamlFront.loadFront(data).experience;

    anonymeDoc.setData({
      description: description.replace('Gabriel Brun', 'X'),
      objectif,
      mainSkills: skills.main
        .map(skill => `${skill.title}: ${skill.rate}/10`)
        .join(', '),
      otherSkills: skills.other.map(skill => `${skill.title}`).join(', '),
      experiences: experience.sort(compare).map(xp => {
        return {
          begin: dateFormat(xp.begin, 'mm/yyyy'),
          end: dateFormat(xp.end, 'mm/yyyy'),
          company: 'Entreprise X',
          job: xp.job,
          tasks: xp.body.split('\n').map(task => {
            return { task };
          })
        };
      })
    });

    try {
      // replace all occurences in document
      doc.render();
      anonymeDoc.render();
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
    const buf2 = anonymeDoc.getZip().generate({ type: 'nodebuffer' });

    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname, 'CV_Gabriel_Brun.docx'), buf);
    fs.writeFileSync(path.resolve(__dirname, 'CV_Anonyme.docx'), buf2);
  }
);
