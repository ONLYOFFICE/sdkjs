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
// Input and output file with errors
const INPUTFILE = process.argv[2] ? `${__dirname}/${process.argv[2]}` : `${__dirname}/input.txt`;
const OUTPUTFILE = `${__dirname}/output.txt`;
// Old serialized props map
const OLD_PROPS_MAP_NAME = 'sdk-all.props.js.map';
// New serialized props map
const NEW_PROPS_MAP_WORD_NAME = 'word.props.js.map';
const NEW_PROPS_MAP_CELL_NAME = 'cell.props.js.map';
const NEW_PROPS_MAP_SLIDE_NAME = 'slide.props.js.map';

function readProps(file) {
	let props = fs.readFileSync(file, {encoding: 'utf-8'});
	let propsLines = props.split('\n');
	let propsMap = {};
	propsLines.forEach(function (line) {
		let lineElems = line.split(':');
		propsMap[lineElems[1]] = lineElems[0]
	});
	return propsMap;
}

async function run(inputFile = "unique.txt", outputFile = "deserialized.docx", mapsDir = "maps", docxTemplate = "deserialize-template.docx") {
	let sdkMaps = {};
	console.log('Read: ', inputFile);
	let isDocx = outputFile.endsWith('.docx');
	let propsReplaced = 0;
	let text = fs.readFileSync(inputFile, {encoding: 'utf-8'});
	let lines = text.split('\n');
	let replaced = lines.map((line) => {
		let sdkMatchRes = line.match(/\/([a-zA-z0-9\-\.]*)\/(sdkjs|web-apps)\/([a-zA-z0-9\-\.]*)\//);

		if (sdkMatchRes && 4 === sdkMatchRes.length) {
			let maps = sdkMaps[sdkMatchRes[0]];
			if (!maps) {
				let pathProps = `${mapsDir}/${sdkMatchRes[1]}/${sdkMatchRes[3]}.props.js.map`;
				let pathVars = `${mapsDir}/${sdkMatchRes[1]}/${sdkMatchRes[3]}.vars.js.map`;
				try {
					maps = {props: readProps(pathProps), vars: readProps(pathVars)};
					sdkMaps[sdkMatchRes[0]] = maps;
					console.log(`Read maps: ${pathProps} and ${pathVars}`);
				} catch (e){
				}
			}
			if (maps) {
				let propsReplacedOld = propsReplaced;
				line = line.replace(/(new )?([a-zA-z0-9$]*)(@http| \(http)/, (match, p1, p2, p3, offset) => {
					let props = (p1 || '.' !== line[offset-1]) ? maps.vars[p2] : maps.props[p2];
					if (props) {
						propsReplaced++;
						let toReplace;
						if (isDocx) {
							toReplace = `<b/>(${props})<b/>`;
						} else {
							toReplace = `(${props})`;
						}
						return `${p1 || ''}${p2}${toReplace}${p3}`;
					} else {
						return match;
					}
				});
			}
			if (isDocx) {
				line = line.replace(`/`+sdkMatchRes[3], `/<b/>${sdkMatchRes[3]}<b/>`)
			}
			//force 8 spaces indent
			line = ' '.repeat(8) + line.trim();
		}
		if (isDocx) {
			let toBold = ['changesError:', 'isDocumentLoadComplete:'];
			for (let i = 0; i < toBold.length; i++) {
				let index = line.indexOf(toBold[i]);
				if (-1 !== index) {
					line = line.substring(0, index) + '<b/>' + line.substring(index) + '<b/>';
				}
			}
			line = makeP(line);
		}
		return line;
	});
	console.log(`Number of replaced properties: ${propsReplaced}`);

	//write with utf8 BOM
	let output = replaced.join('\n');
	if (isDocx) {
		await writeDocx(docxTemplate, outputFile, output);
	} else {
		fs.writeFileSync(outputFile, '\ufeff' + output, {encoding: 'utf-8'});
	}
	console.log('Complete writeFileSync:' + outputFile);
}

function escapeXml(unsafe) {
	return unsafe.replace(/[<>&'"]/g, function (c) {
		switch (c) {
			case '<':
				return '&lt;';
			case '>':
				return '&gt;';
			case '&':
				return '&amp;';
			case '\'':
				return '&apos;';
			case '"':
				return '&quot;';
		}
	});
}

function makeP(val) {
	let res = '<w:p><w:pPr><w:pStyle w:val="1_634"/></w:pPr>';
	let vals = val.split('<b/>');
	for (let i = 0; i < vals.length; i++) {
		res += makeBoldRun(vals[i], i % 2 === 1);
	}
	res += '</w:p>';
	return res;
}

function makeBoldRun(val, isBold) {
	let bold = '';
	if (isBold) {
		switch (val) {
			case 'word':
				bold = `<w:rPr><w:b/><w:bCs/><w:color w:val="0000FF"/></w:rPr>`;
				break;
			case 'cell':
				bold = `<w:rPr><w:b/><w:bCs/><w:color w:val="7FBC00"/></w:rPr>`;
				break;
			case 'slide':
				bold = `<w:rPr><w:b/><w:bCs/><w:color w:val="FFA500"/></w:rPr>`;
				break;
			case 'apps':
				bold = `<w:rPr><w:b/><w:bCs/><w:color w:val="7030A0"/></w:rPr>`;
				break;
			default:
				bold = `<w:rPr><w:b/><w:bCs/><w:color w:val="FF0000"/></w:rPr>`;
		}
	}
	return `<w:r>${bold}<w:t xml:space="preserve">${escapeXml(val)}</w:t></w:r>`;
}

async function writeDocx(templateFile, outputFile, data) {
	const JSZip = require('jszip');
	let zipIn = fs.readFileSync(templateFile);
	let zip = await JSZip.loadAsync(zipIn);
	let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:body>'
	xml += data;
	xml += '<w:sectPr><w:footnotePr></w:footnotePr><w:endnotePr></w:endnotePr><w:type w:val="nextPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="709" w:footer="709" w:gutter="0"/><w:cols w:num="1" w:sep="0" w:space="708" w:equalWidth="1" ></w:cols><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
	zip.file("word/document.xml", xml);
	let zipOut = await zip.generateAsync({type: "nodebuffer", compression: "DEFLATE"});
	fs.writeFileSync(outputFile, zipOut);
}

run.apply(this, process.argv.slice(2)).catch(console.error);
