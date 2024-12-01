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

"use strict";

function DrawTocPreview(api, canvas, index, pixW, scale)
{
	if (undefined === scale)
		scale = AscCommon.AscBrowser.retinaPixelRatio;

	let prop = null;
	if (0 === index)
	{
		prop = {
			OutlineStart : 1,
			OutlineEnd : 3,
			Hyperlink : false,
			StylesType : Asc.c_oAscTOCStylesType.Simple,
			RightTab : true,
			PageNumbers : true,
			TabLeader : Asc.c_oAscTabLeader.Dot,
			Pages : [2, 5, 15]
		};
	}
	else
	{
		prop = {
			OutlineStart : 1,
			OutlineEnd : 3,
			Hyperlink : true,
			StylesType : Asc.c_oAscTOCStylesType.Web,
			RightTab : true,
			PageNumbers : false,
			TabLeader : Asc.c_oAscTabLeader.None
		};
	}

	let mmW = pixW * g_dKoef_pix_to_mm;

	// content
	let oLogicDocument   = api.WordControl.m_oLogicDocument;
	let oDrawindDocument = api.WordControl.m_oDrawingDocument;
	let oStyles          = oLogicDocument.GetStyles();
	let oHeader          = new CHeaderFooter(oLogicDocument.HdrFtr, oLogicDocument, oDrawindDocument, AscCommon.hdrftr_Header);
	let oDocumentContent = oHeader.GetContent();

	let arrLevels         = [];
	let arrStylesToDelete = [];

	let wPxOffset = ((8 * scale) + 0.5) >> 0;
	let wMmOffset = wPxOffset * g_dKoef_pix_to_mm / scale;

	for (let nCurrentLevel = prop.OutlineStart; nCurrentLevel <= prop.OutlineEnd; ++nCurrentLevel)
	{
		let sName = "Heading " + nCurrentLevel;
		let nLvl  = nCurrentLevel - 1;

		let oStyle = new CStyle("", null, null, styletype_Paragraph, true);
		oStyle.CreateTOC(nLvl, prop.StylesType);

		oStyle.ParaPr.Spacing.Line = 1.2;
		oStyle.ParaPr.Spacing.LineRule = linerule_Auto;
		oStyle.ParaPr.Spacing.Before = 0;
		oStyle.ParaPr.Spacing.After = 0;
		oStyle.ParaPr.ContextualSpacing = true;

		oStyle.ParaPr.Ind.Left = 15 * (nCurrentLevel - 1) * g_dKoef_pt_to_mm;

		oStyle.TextPr.FontFamily = {Name: "Arial", Index: -1};
		oStyle.TextPr.FontSize = 10;

		oStyles.Add(oStyle);

		arrLevels[nLvl] = {
			Styles  : [sName],
			StyleId : oStyle.GetId()
		};

		arrStylesToDelete.push(oStyle.GetId());
	}

	for (let nCurrentLevel = prop.OutlineStart; nCurrentLevel <= prop.OutlineEnd; ++nCurrentLevel)
	{
		var sStyleId = arrLevels[nCurrentLevel - 1].StyleId;
		for (let nStyle = 0, nStylesCount = arrLevels[nCurrentLevel - 1].Styles.length; nStyle < nStylesCount; ++nStyle)
		{
			var sStyleName = AscCommon.translateManager.getValue(arrLevels[nCurrentLevel - 1].Styles[nStyle]);

			var oParagraph = new AscWord.Paragraph(oDocumentContent, false);
			oDocumentContent.AddToContent(nCurrentLevel - 1, oParagraph);
			oParagraph.SetParagraphStyleById(sStyleId);

			var oRun = new ParaRun(oParagraph, false);
			oParagraph.AddToContent(0, oRun);
			oRun.AddText(sStyleName);

			if (prop.PageNumbers)
			{
				if (prop.RightTab)
				{
					var oParaTabs = new CParaTabs();
					oParaTabs.Add(new CParaTab(tab_Right, mmW - 4 - wMmOffset, prop.TabLeader));
					oParagraph.SetParagraphTabs(oParaTabs);

					oRun.AddToContent(-1, new AscWord.CRunTab());
				}
				else
				{
					oRun.AddToContent(-1, new AscWord.CRunSpace());
				}

				oRun.AddText("" + prop.Pages[nCurrentLevel - 1]);
			}
		}
	}

	// удаляем последний параграф
	oDocumentContent.Content.splice(3, 1);
	oDocumentContent.Reset(1, 0, 1000, 10000);
	oDocumentContent.Recalculate_Page(0, true);

	for (let nStyle = 0, nStylesCount = arrStylesToDelete.length; nStyle < nStylesCount; ++nStyle)
	{
		oStyles.Remove(arrStylesToDelete[nStyle]);
	}

	let mmH = oDocumentContent.GetSummaryHeight() + (wMmOffset * 2);
	let pixH = mmH / g_dKoef_pix_to_mm;
	pixH = ((pixH + 3) >> 2) << 2;
	let pixW_natural = ((pixW * scale) + 0.5) >> 0;
	let pixH_natural = ((pixH * scale) + 0.5) >> 0;

	let graphics = null;
	if (canvas.style !== undefined)
	{
		canvas.style.width = pixW + "px";
		canvas.style.height = pixH + "px";

		canvas.width = pixW_natural;
		canvas.height = pixH_natural;

		let ctx = canvas.getContext('2d');

		ctx.fillStyle = "#FFFFFF";
		ctx.fillRect(0, 0, pixW_natural, pixH_natural);

		graphics = new AscCommon.CGraphics();
		graphics.init(ctx, pixW_natural, pixH_natural, mmW, mmH);
		graphics.m_oFontManager = AscCommon.g_fontManager;
		graphics.m_oCoordTransform.tx = graphics.m_oCoordTransform.ty = wPxOffset;
		graphics.transform(1, 0, 0, 1, 0, 0);
	}
	else
	{
		graphics = new CNativeGraphics(canvas);
		graphics.width  = pixW_natural;
		graphics.height = pixH_natural;
		graphics.create(null, pixW_natural, pixH_natural, mmW, mmH);
		graphics.CoordTransformOffset(wPxOffset, wPxOffset);
		graphics.transform(1, 0, 0, 1, 0, 0);
	}

	oDocumentContent.Draw(0, graphics);
}

function DrawCustomTocPreview(api, canvas, props, widthPx, heightPx, scale)
{
	let oLogicDocument   = api.WordControl.m_oLogicDocument;
	let oDrawingDocument = api.WordControl.m_oDrawingDocument;
	let oStyles          = oLogicDocument.GetStyles();

	let oHeader          = new CHeaderFooter(oLogicDocument.HdrFtr, oLogicDocument, oDrawingDocument, AscCommon.hdrftr_Header);
	let oDocumentContent = oHeader.GetContent();

	let nOutlineStart = props.get_OutlineStart();
	let nOutlineEnd   = props.get_OutlineEnd();
	let nStylesType   = props.get_StylesType();
	let isShowPageNum = props.get_ShowPageNumbers();
	let isRightTab    = props.get_RightAlignTab();
	let nTabLeader    = props.get_TabLeader();

	if (undefined === scale)
		scale = AscCommon.AscBrowser.retinaPixelRatio;

	let wPx = ((widthPx * scale) + 0.5) >> 0;
	let hPx = ((heightPx * scale) + 0.5) >> 0;
	let wMm = wPx * g_dKoef_pix_to_mm / scale;
	let hMm = hPx * g_dKoef_pix_to_mm / scale;

	let wPxOffset = ((8 * scale) + 0.5) >> 0;
	let wMmOffset = wPxOffset * g_dKoef_pix_to_mm / scale;

	if (undefined === nTabLeader || null === nTabLeader)
		nTabLeader = Asc.c_oAscTabLeader.Dot;

	let arrLevels         = [];
	let arrStylesToDelete = [];

	let nStyle, nStylesCount, nAddStyle, nAddStyleCount;
	let nLvl, sName, sStyleId, oStyle, isAddStyle;
	for (nStyle = 0, nStylesCount = props.get_StylesCount(); nStyle < nStylesCount; ++nStyle)
	{
		nLvl  = props.get_StyleLevel(nStyle) - 1;
		sName = props.get_StyleName(nStyle);

		if (!arrLevels[nLvl])
		{
			sStyleId = null;
			if (Asc.c_oAscTOCStylesType.Current === nStylesType)
			{
				sStyleId = oStyles.GetDefaultTOC(nLvl);
			}
			else
			{
				oStyle = new CStyle("", null, null, styletype_Paragraph, true);
				oStyle.CreateTOC(nLvl, nStylesType);
				sStyleId = oStyle.GetId();
				oStyles.Add(oStyle);
				arrStylesToDelete.push(oStyle.GetId());
			}
			arrLevels[nLvl] = {
				Styles  : [],
				StyleId : sStyleId
			};
		}

		isAddStyle = true;
		for (nAddStyle = 0, nAddStyleCount = arrLevels[nLvl].Styles.length; nAddStyle < nAddStyleCount; ++nAddStyle)
		{
			if (arrLevels[nLvl].Styles[nAddStyle] === sName)
			{
				isAddStyle = false;
				break;
			}
		}

		if (isAddStyle)
			arrLevels[nLvl].Styles.push(sName);
	}

	if (-1 !== nOutlineEnd && -1 !== nOutlineStart)
	{
		for (let _nLvl = nOutlineStart; _nLvl <= nOutlineEnd; ++_nLvl)
		{
			sName = "Heading " + _nLvl;
			nLvl  = _nLvl - 1;

			if (!arrLevels[nLvl])
			{
				sStyleId = null;
				if (Asc.c_oAscTOCStylesType.Current === nStylesType)
				{
					sStyleId = oStyles.GetDefaultTOC(nLvl);
				}
				else
				{
					oStyle = new CStyle("", null, null, styletype_Paragraph, true);
					oStyle.CreateTOC(nLvl, nStylesType);
					sStyleId = oStyle.GetId();
					oStyles.Add(oStyle);
					arrStylesToDelete.push(oStyle.GetId());
				}

				arrLevels[nLvl] = {
					Styles  : [],
					StyleId : sStyleId
				};
			}

			isAddStyle = true;
			for (nAddStyle = 0, nAddStyleCount = arrLevels[nLvl].Styles.length; nAddStyle < nAddStyleCount; ++nAddStyle)
			{
				if (arrLevels[nLvl].Styles[nAddStyle] === sName)
				{
					isAddStyle = false;
					break;
				}
			}

			if (isAddStyle)
				arrLevels[nLvl].Styles.push(sName);
		}
	}

	let oParaIndex = 0;
	let nPageIndex = 1;

	for (nLvl = 0; nLvl <= 8; ++nLvl)
	{
		if (!arrLevels[nLvl])
			continue;

		sStyleId = arrLevels[nLvl].StyleId;
		for (nStyle = 0, nStylesCount = arrLevels[nLvl].Styles.length; nStyle < nStylesCount; ++nStyle)
		{
			let sStyleName = AscCommon.translateManager.getValue(arrLevels[nLvl].Styles[nStyle]);

			let oParagraph = new AscWord.Paragraph(oDocumentContent, false);
			oDocumentContent.AddToContent(oParaIndex++, oParagraph);
			oParagraph.SetParagraphStyleById(sStyleId);

			let oRun = new ParaRun(oParagraph, false);
			oParagraph.AddToContent(0, oRun);
			oRun.AddText(sStyleName);

			if (isShowPageNum)
			{
				if (isRightTab)
				{
					var oParaTabs = new CParaTabs();
					oParaTabs.Add(new CParaTab(tab_Right, wMm - 2 - wMmOffset, nTabLeader));
					oParagraph.SetParagraphTabs(oParaTabs);

					oRun.AddToContent(-1, new AscWord.CRunTab());
				}
				else
				{
					oRun.AddToContent(-1, new AscWord.CRunSpace());
				}

				oRun.AddText("" + nPageIndex);

				nPageIndex += 2;
			}
		}
	}

	oDocumentContent.Reset(1, 0, 1000, 10000);
	oDocumentContent.Recalculate_Page(0, true);

	for (nStyle = 0, nStylesCount = arrStylesToDelete.length; nStyle < nStylesCount; ++nStyle)
	{
		oStyles.Remove(arrStylesToDelete[nStyle]);
	}

	var nContentHeight = oDocumentContent.GetSummaryHeight();
	var nContentHeightPx = (scale * nContentHeight / g_dKoef_pix_to_mm) >> 0;

	if (nContentHeightPx > hPx)
	{
		hPx = nContentHeightPx;
		hMm = nContentHeight;
	}

	let graphics = null;
	if (canvas.style !== undefined)
	{
		canvas.style.width = widthPx + "px";
		canvas.width = wPx;

		canvas.style.height = (((hPx / scale) + 0.5) >> 0) + "px";
		canvas.height = hPx;

		let ctx = canvas.getContext('2d');

		ctx.fillStyle = "#FFFFFF";
		ctx.fillRect(0, 0, wPx, hPx);

		graphics = new AscCommon.CGraphics();
		graphics.init(ctx, wPx, hPx, wMm, hMm);
		graphics.m_oFontManager = AscCommon.g_fontManager;
		graphics.m_oCoordTransform.tx = graphics.m_oCoordTransform.ty = wPxOffset;
		graphics.transform(1, 0, 0, 1, 0, 0);
	}
	else
	{
		graphics = new CNativeGraphics(canvas);
		graphics.width  = wPx;
		graphics.height = hPx;
		graphics.create(null, wPx, hPx, wMm, hMm);
		graphics.CoordTransformOffset(wPxOffset, wPxOffset);
		graphics.transform(1, 0, 0, 1, 0, 0);
	}

	oDocumentContent.Draw(0, graphics);
}
