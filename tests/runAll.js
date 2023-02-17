/*
 * (c) Copyright Ascensio System SIA 2010-2022
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

const path = require("path");

const allTests = [
	'cell/spreadsheet-calculation/FormulaTests.html',
	'cell/spreadsheet-calculation/PivotTests.html',
	'cell/spreadsheet-calculation/CopyPasteTests.html',
	'word/unit-tests/paragraphContentPos.html',
	'word/document-calculation/paragraph.html',
	'word/document-calculation/table/flowTablePosition.html',
	'word/document-calculation/textShaper/textShaper.html',
	'word/forms/forms.html',
	'word/forms/complexForm.html',
	'word/api/api.html',
	'word/api/textInput.html',
	'word/styles/paraPr.html',
	'word/plugins/pluginsApi.html',
	'word/merge-documents/mergeDocuments.html',

	'oform/xml/oformXml.html'
];

const {performance} = require('perf_hooks');

const {
  runQunitPuppeteer,
  printResultSummary,
  printFailedTests
} = require("node-qunit-puppeteer");

async function Run()
{
	let startTime = performance.now();
	let count  = 0;
	let failed = [];
	for (let nIndex = 0, nCount = allTests.length; nIndex < nCount; ++nIndex)
	{
		await runQunitPuppeteer({targetUrl : path.join(__dirname, allTests[nIndex])})
			.then(result =>
			{
				count++;
				printResultSummary(result, console);

				if (result.stats.failed > 0)
				{
					printFailedTests(result, console);
					failed.push(allTests[nIndex]);
				}
			})
			.catch(ex =>
			{
				count++;
				failed.push(allTests[nIndex]);
				console.error(ex);
			});
	}

	console.log("\nOverall Elapsed " + (Math.round(( ((performance.now() - startTime) / 1000) + Number.EPSILON) * 1000) / 1000) + "s");
	console.log("\n"+ (count - failed.length) + "/" + count + " modules successfully passed the tests");

	if (failed.length)
	{
		console.log("\nFAILED".red.bold);
		for (let nIndex = 0, nCount = failed.length; nIndex < nCount; ++nIndex)
		{
			console.log(failed[nIndex]);
		}
	}
	else
	{
		console.log("\nPASSED".green.bold);
	}
}

Run();

