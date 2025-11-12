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

const fs = require('fs');
const {writeFile, mkdir, stat} = require('fs/promises');
const {S3Client, GetObjectCommand} = require("@aws-sdk/client-s3");
const fetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));

async function run(accessKeyId, secretAccessKey, inputFile = "unique.txt", mapsDir = "maps", region = "eu-west-1",
				   endpoint = "https://s3.eu-west-1.amazonaws.com", bucketName = "repo-doc-onlyoffice-com",
				   keyPrefix = "closure-maps/sdkjs/commercial/{version}") {
	if (!accessKeyId || !secretAccessKey) {
		console.error(`missing arguments.USAGE: download-maps.js [accessKeyId] [secretAccessKey] [inputFile=${inputFile}] [mapsDir=${mapsDir}] [region=${region}] [endpoint=${endpoint}] [bucketName=${bucketName}] [keyPrefix=${keyPrefix}]`);
		console.error(`inputFile must contain at least one line like: "/7.5.1-23/sdkjs/"`);
		return;
	}
	const configS3 = {
		region: region,
		endpoint: endpoint,
		credentials: {
			accessKeyId: accessKeyId,
			secretAccessKey: secretAccessKey
		}
	};

	let versionBuild = '';
	let versionsSet = new Set();
	console.log('Read: ', inputFile);
	let text = fs.readFileSync(inputFile, {encoding: 'utf-8'});
	let lines = text.split('\n');
	for (let line of lines) {
		let sdkMatchRes = line.match(/(http.*?)\/([a-zA-z0-9\-\.]*)\/sdkjs\//);
		if (!sdkMatchRes || 3 !== sdkMatchRes.length) {
			continue;
		}
		versionsSet.add(sdkMatchRes[2]);
		if (!versionBuild) {
			const apiUrl = `${sdkMatchRes[1]}/web-apps/apps/api/documents/api.js`;
			console.log(`Request for build version: ${apiUrl}`);
			const response = await fetch(`${sdkMatchRes[1]}/web-apps/apps/api/documents/api.js`);
			const apiContent = await response.text();
			const match = apiContent.match(/\(build:(\d+)\)/);
			if (match) {
				versionBuild = match[1];
				console.log(`Found build version: ${versionBuild}`);
			}
		}
	}
	let versions = Array.from(versionsSet);
	console.log('Found versions: ' + JSON.stringify(versions));
	const editors = ['word', 'cell', 'slide'];
	const maps = ['.props.js.map', '.vars.js.map', '-all.js.map', '-all-min.js.map'];

	const client = new S3Client(configS3);
	for (let version of versions) {
		await mkdir(`${mapsDir}/${version}`, {recursive: true});
		for (let editor of editors) {
			for (let map of maps) {
				let filePath = `${mapsDir}/${version}/${editor}${map}`;
				try {
					await stat(filePath);
					console.log('Skip file exists: ', filePath);
				} catch (err) {
					let versionS3 = version.replace(/-[a-zA-z0-9\-\.]*/g, '/');
					const key = keyPrefix.replace('{version}', versionS3 + versionBuild) + `/${editor}${map}`;
					console.log(`Download: ${mapsDir}/${version}/${editor}${map} from ${key}`);
					let body = "";
					try {
						const command = new GetObjectCommand({
							Bucket: bucketName,
							Key: key
						});
						const output = await client.send(command);
						body = output.Body;
					} catch (err) {
						console.log('Skip file exists due download error(maybe wrong versionBuild detection): ', err);
					}
					await writeFile(filePath, body);
				}
			}
		}
	}
	console.log('Download complete');
}

run.apply(this, process.argv.slice(2)).catch(console.error);
