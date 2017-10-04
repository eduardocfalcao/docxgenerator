'use strict'

const generateDocx = require('generate-docx');
var loremIpsum = require('lorem-ipsum')
const filesTogenerate = 1

for (var i = 0; i < filesTogenerate; i++) {

    (function(j) {
        var body = loremIpsum({
            count: 95000 // Number of words, sentences, or paragraphs to generate. - 95000 -> 6MB file; 9500 -> 6KB File
                ,
            units: 'sentences' // Generate words, sentences, or paragraphs. 
                ,
            sentenceLowerBound: 5 // Minimum words per sentence. 
                ,
            sentenceUpperBound: 15 // Maximum words per sentence. 
                ,
            paragraphLowerBound: 3 // Minimum sentences per paragraph. 
                ,
            paragraphUpperBound: 7 // Maximum sentences per paragraph. 
                ,
            format: 'plain' // Plain text or html       
        });

        var fileName = 'testdoc' + j + '.docx';
        console.log('generating ' + fileName);

        var options = {
            template: {
                filePath: 'template/template.docx',
                data: {
                    'title': 'This is the title o file ' + fileName,
                    'description': 'Description is good of file' + fileName,
                    'body': body
                }
            },
            save: {
                filePath: 'C:\\files\\' + fileName
            }
        }

        generateDocx(options)
            .then(message => console.log(message))
            .catch(error => console.error(error));

    })(i);
}