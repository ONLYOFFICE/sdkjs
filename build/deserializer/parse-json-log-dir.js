/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

const fs = require('node:fs');

function run(inputDir = "logs", outputFile = "unique.txt", opt_lastFile = "") {
	let files = fs.readdirSync(inputDir);
	files = files.filter((file) => {
		return fs.statSync(`${inputDir}/${file}`).isFile() && file !== opt_lastFile;
	});
	console.log(`Read inputDir - files without ${opt_lastFile}: ` + JSON.stringify(files));
	let paths = files.map((file) => {
		return `${inputDir}/${file}`;
	});
	//sort files by modified date
	paths.sort((a, b) => {
		return fs.statSync(a).mtimeMs - fs.statSync(b).mtimeMs;
	});
	console.log('Sort by Modified date: ' + JSON.stringify(paths));
	//read files
	let lines = paths.map((path) => {
		let text = fs.readFileSync(path, {encoding: 'utf-8'});
		return text.split('\n');
	}).flat(1);
	console.log('All lines: ' + lines.length);
	//filter duplicates
	let unique = {}, linesUnique;
	let filterUnique = (line) => {
		let indexStart = line.indexOf("Error:");
		if (-1 !== indexStart) {
			let key = line;
			key = key.substring(indexStart);
			//replace host (diff .com .co)
			key = key.replace(/https?:\/\/.*?\//, '/');

			let indexUserAgent = key.indexOf("userAgent:");
			if (-1 !== indexUserAgent) {
				key = key.substring(0, indexUserAgent);
			} else {
				//for History.Add without Create_NewPoint
				let indexNewLine = key.indexOf("\\n");
				if (-1 !== indexNewLine) {
					indexNewLine = key.indexOf("\\n", indexNewLine + 1);
				}
				if (-1 === indexNewLine) {
					indexNewLine = key.indexOf("}");
				}
				if (-1 !== indexNewLine) {
					key = key.substring(0, indexNewLine);
				}
			}
			let version = "no-version";
			let sdkMatchRes = line.match(/\/([a-zA-z0-9\-\.]*)\/(sdkjs|web-apps)\//);
			if (sdkMatchRes && 1 < sdkMatchRes.length) {
				version = sdkMatchRes[1]
			}
			let uniqueVersion = unique[version];
			if (!uniqueVersion) {
				uniqueVersion = {};
				unique[version] = uniqueVersion;
			}
			if (!uniqueVersion.hasOwnProperty(key)) {
				uniqueVersion[key] = 1;
				return true;
			}
		}
		return false;
	};
	linesUnique = lines.filter(filterUnique);
	console.log('Found versions: ' + JSON.stringify(Object.keys(unique)));
	console.log('Unique lines: ' + linesUnique.length);
	if (opt_lastFile) {
		let text = fs.readFileSync(`${inputDir}/${opt_lastFile}`, {encoding: 'utf-8'});
		let lines = text.split('\n');
		linesUnique = lines.filter(filterUnique);
		console.log(`All lines in ${opt_lastFile}: ` + lines.length);
		console.log(`Found versions with ${opt_lastFile}: ` + JSON.stringify(Object.keys(unique)));
		console.log(`Unique lines in ${opt_lastFile}: ` + linesUnique.length);
	}
	//trim line
	linesUnique = linesUnique.map((line) => {
		let indexStart = line.indexOf('{\\"startTime\\":');
		if (-1 !== indexStart) {
			line = line.substring(indexStart);
		}
		let indexUserAgent = line.indexOf("userAgent:");
		if (-1 === indexUserAgent) {
			line = "History.Add without Create_NewPoint or other self triggered error:\n" + line;
		}
		return line;
	});

	let output = linesUnique.join('\n\n');
	output = output.replace(/\\n/g, '\n');

	fs.writeFileSync(outputFile, output, {encoding: 'utf-8'});
	console.log('Complete writeFileSync:' + outputFile);
}

run.apply(this, process.argv.slice(2));
