const extract = require('extract-zip')
const zipper = require('zip-local');

const fs = require('fs');
const readline = require('readline').createInterface({
	input: process.stdin,
	output: process.stdout,
});

class Doc {
	constructor(fileDirectory) {
		this.operatingSystem = process.platform;
		this.fileDirectory = `${__dirname}/${fileDirectory}`;
		this.fileName = null;
	}

	askFileName() {
		return new Promise((resolve) => {
			readline.question(
				"\nUndoc, a tool made by Camille Rakotoarisoa \n_______\n\n1) Avant tout, activez l'affichage des extensions de fichiers sur votre machine.\n2) Ensuite, placez votre fichier .docx dans le dossier 'myFiles/' (ne mettez ni espaces, ni caractères spéciaux dans le nom de celui-ci).\n3) Saisissez le nom de votre fichier Word sans saisir l'extension en .docx \n\n>>> ",
				(name) => {
					readline.close();
					this.fileName = name;
					resolve();
				}
			);
		});
	}

	convertFile(action) {
		if (action === 'zip to docx') {
			const oldName = `${this.fileDirectory}/${this.fileName}.zip`;
			const newName = `${this.fileDirectory}/${this.fileName}.docx`;
			return new Promise((resolve) => {
				fs.rename(oldName, newName, (err) => {
					if (err) throw err;
					resolve();
				});
			});
		} else if (action === 'docx to zip') {
			const oldName = `${this.fileDirectory}/${this.fileName}.docx`;
			const newName = `${this.fileDirectory}/${this.fileName}.zip`;
			return new Promise((resolve) => {
				fs.rename(oldName, newName, (err) => {
					if (err) throw err;
					resolve();
				});
			});
		}
	}

	extractZippedFile() {
		const fileToExtract = `${this.fileDirectory}/${this.fileName}.zip`;
		const sourceOfExtractedFile = `${this.fileDirectory}/${this.fileName}`;
		return new Promise(async (resolve) => {
			await extract(fileToExtract, { dir: sourceOfExtractedFile });
			resolve();
		});
	}

	delete(action) {
		if (action === 'delete-zip') {
			const pathOfFileToDelete = `${this.fileDirectory}/${this.fileName}.zip`;
			return new Promise((resolve) => {
				fs.unlinkSync(pathOfFileToDelete);
				resolve();
			});
		} else if (action === 'delete-folder') {
			const pathOfFolderToDelete = `${this.fileDirectory}/${this.fileName}`;
			return new Promise((resolve) => {
				fs.rmdirSync(pathOfFolderToDelete, {
					recursive: true,
				});
				resolve();
			});
		}
	}

	deleteSettingsXML() {
		const settingsFilePath = `${this.fileDirectory}/${this.fileName}/word/settings.xml`;
		new Promise((resolve) => {
			fs.unlinkSync(settingsFilePath);
			resolve();
		});
	}

	zipFoldersDoc() {
		const folder_source = `${this.fileDirectory}/${this.fileName}`;
		const zip_source = `${this.fileDirectory}/${this.fileName}.zip`;
		return new Promise((resolve) => {
			zipper.sync.zip(folder_source).compress().save(zip_source);
			resolve();
		});
	}
}

module.exports = Doc;
