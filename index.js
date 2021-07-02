const Doc = require('./Doc');
const doc = new Doc('myFiles');

//Clear Terminal
process.stdout.write('\x1Bc'); 

doc
	.askFileName()
	.then(() => doc.convertFile('docx to zip'))
	.then(() => doc.extractZippedFile())
	.then(() => doc.delete('delete-zip'))
	.then(() => doc.deleteSettingsXML())
	.then(() => doc.zipFoldersDoc())
	.then(() => doc.convertFile('zip to docx'))
	.then(() => doc.delete('delete-folder'))
	.then(() => console.log('Votre fichier Word a bien été débloqué.'));
