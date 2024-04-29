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

var polygon1 = {
	inverted : false,
	regions : [
		[
			[50, 50],
			[150, 50],
			[150, 150],
			[50, 150]
		],
		[
			[50, 450],
			[50, 350],
			[100, 300],
			[150, 350],
			[400, 350],
			[400, 450]
		]
	]
};
var polygon2 = {
	inverted : false,
	regions : [
		[
			[100, 100],
			[200, 100],
			[200, 200],
			[100, 200]
		],
		[
			[100, 480],
			[70, 310],
			[130, 330],
			[150, 300],
			[300, 400]
		],
		[
			[25, 400],
			[225, 300],
			[475, 400],
			[225, 475]
		]
	]
};

function onchangemode(mode)
{
	draw(mode);
}

function fillRegion(polygon, ctx)
{
	for (let i = 0, countPolygons = polygon.regions.length; i < countPolygons; i++)
	{
		let region = polygon.regions[i];
		let countPoints = region.length;

		if (2 > countPoints)
			continue;

		ctx.moveTo(region[0][0], region[0][1]);

		for (let j = 1, countPoints = region.length; j < countPoints; j++)
		{
			ctx.lineTo(region[j][0], region[j][1]);
		}

		ctx.closePath();
	}
}

function draw(mode)
{
	let width = 500;
	let height = 500;

	let canvasSource = document.getElementById("sources");
	let ctxSource = canvasSource.getContext("2d");

	ctxSource.clearRect(0, 0, width, height);
	ctxSource.strokeStyle = "#000000";
	ctxSource.lineWidth = 1;
	ctxSource.strokeRect(0, 0, width, height);

	// 1
	ctxSource.beginPath();
	ctxSource.fillStyle = "rgba(255, 0, 0, 0.5)";
	fillRegion(polygon1, ctxSource);
	ctxSource.fill("evenodd");

	// 2
	ctxSource.beginPath();
	ctxSource.fillStyle = "rgba(0, 0, 255, 0.5)";
	fillRegion(polygon2, ctxSource);
	ctxSource.fill("evenodd");

	// clear path
	ctxSource.beginPath();

	let canvasResult = document.getElementById("result");
	let ctxResult = canvasResult.getContext("2d");

	ctxResult.clearRect(0, 0, width, height);
	ctxResult.strokeStyle = "#000000";
	ctxResult.lineWidth = 1;
	ctxResult.strokeRect(0, 0, width, height);

	let apply = AscGeometry.PolyBool[mode];

	let time1 = performance.now();
	let regionResult = apply(polygon1, polygon2);
	let time2 = performance.now();

	console.log("time: " + (time2 - time1));

	ctxResult.beginPath();
	fillRegion(regionResult, ctxResult);
	ctxResult.stroke();
	ctxResult.beginPath();
}

window.onload = function()
{
	let combo = document.getElementById("id_mode");
	combo.selectedIndex = 0;
	onchangemode(combo.value);
}
