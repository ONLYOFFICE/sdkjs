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

(function (window)
{

	const asc_PreviewBulletType = {
		text: 0,
		char: 1,
		image: 2,
		number: 3,
		multiLevel: 4
	}
	window["Asc"].asc_PreviewBulletType = window["Asc"]["asc_PreviewBulletType"] = asc_PreviewBulletType;
	asc_PreviewBulletType["text"] = asc_PreviewBulletType.text;
	asc_PreviewBulletType["char"] = asc_PreviewBulletType.char;
	asc_PreviewBulletType["image"] = asc_PreviewBulletType.image;
	asc_PreviewBulletType["number"] = asc_PreviewBulletType.number;
	asc_PreviewBulletType["multiLevel"] = asc_PreviewBulletType.multiLevel;
	function CBulletPreviewDrawerBase()
	{
		this.m_arrNumberingLvl = [];
		this.m_oApi = editor || Asc.editor || window["Asc"]["editor"];
		this.m_oLogicDocument = this.m_oApi.WordControl && this.m_oApi.WordControl.m_oLogicDocument;
		this.m_oDrawingDocument = this.m_oLogicDocument && this.m_oLogicDocument.DrawingDocument;
		this.m_oLang = this.m_oApi.asc_GetPossibleNumberingLanguage();

		this.m_oPrimaryTextColor = new AscCommonWord.CDocumentColor(0, 0, 0);
		// для словесного текста используем цвет контрастнее
		this.m_oSecondaryTextColor = new AscCommonWord.CDocumentColor(121, 121, 121);
		this.m_oSecondaryLineTextColor = new AscCommonWord.CDocumentColor(203, 203, 203);
		this.m_oBackgroundColor = new AscCommonWord.CDocumentColor(255, 255, 255);

		this.m_nAmountOfLvls = 9;

		this.m_bIsMobile = AscCommon.AscBrowser.isMobile;
	}

	CBulletPreviewDrawerBase.prototype.cleanTextPr = function (oTextPr)
	{
		oTextPr.VertAlign  = undefined;
		oTextPr.RStyle     = undefined;
		oTextPr.Position   = undefined; // Смещение по Y

		oTextPr.BoldCS     = undefined;
		oTextPr.ItalicCS   = undefined;
		oTextPr.FontSizeCS = undefined;
		oTextPr.CS         = undefined;
		oTextPr.RTL        = undefined;
		oTextPr.FontRef    = undefined;
		oTextPr.Shd        = undefined;
		oTextPr.Vanish     = undefined;
		oTextPr.Ligatures  = undefined;
		oTextPr.TextOutline    = undefined;
		oTextPr.TextFill       = undefined;
		oTextPr.PrChange       = undefined;
		oTextPr.ReviewInfo     = undefined;
	};
	CBulletPreviewDrawerBase.prototype.drawImageBulletsWithLine = function (oImageInfo, nX, nY, nLineHeight, oGraphics, oStyleTextOptions, oTextPr) {
		const oImage = oImageInfo.image;
		if (oImage)
		{
			const sFullImageSrc = oImage.src;
			const oSizes = AscCommon.getSourceImageSize(sFullImageSrc);
			const nImageHeight = oSizes.height;
			const nImageWidth = oSizes.width;
			const nAdaptImageHeight = nLineHeight;
			const nAdaptImageWidth = (nImageWidth * nAdaptImageHeight / (nImageHeight ? nImageHeight : 1));

			for (let i = 0; i < oImageInfo.amount; i += 1)
			{
				this.cleanParagraphField(oGraphics, nX * AscCommon.g_dKoef_pix_to_mm, (nY - nLineHeight) * AscCommon.g_dKoef_pix_to_mm, (nAdaptImageWidth + 2) * AscCommon.g_dKoef_pix_to_mm, (nLineHeight + (nLineHeight >> 1)) * AscCommon.g_dKoef_pix_to_mm);
				oGraphics.drawImage(sFullImageSrc, nX * AscCommon.g_dKoef_pix_to_mm, (nY - nAdaptImageHeight * (0.85)) * AscCommon.g_dKoef_pix_to_mm, nAdaptImageWidth * AscCommon.g_dKoef_pix_to_mm, nAdaptImageHeight * AscCommon.g_dKoef_pix_to_mm);
				nX += nAdaptImageWidth;
			}
			this.drawStyleText(oGraphics, oStyleTextOptions, nX, nY, nLineHeight, oTextPr);
		}
	};

	CBulletPreviewDrawerBase.prototype.getFirstLineIndent = function (oLvl, nCustomNumberPosition, nCustomIndentSize, nCustomStopTab)
	{
		const nSuff = oLvl.GetSuff();
		const nNumberPosition = (AscFormat.isRealNumber(nCustomNumberPosition) ? nCustomNumberPosition : oLvl.GetNumberPosition()) || 0;
		let nXPositionOfLine;
		if (nSuff === Asc.c_oAscNumberingSuff.Tab)
		{
			const nStopTab = AscFormat.isRealNumber(nCustomStopTab) || nCustomStopTab === null ? nCustomStopTab : oLvl.GetStopTab();
			const nIndentSize = (AscFormat.isRealNumber(nCustomIndentSize) ? nCustomIndentSize : oLvl.GetIndentSize()) || 0;
			if (AscFormat.isRealNumber(nStopTab))
			{
				nXPositionOfLine = Math.max(nStopTab, nNumberPosition);
			}
			else
			{
				nXPositionOfLine =  Math.max(nNumberPosition, nIndentSize);
			}

		}
		else
		{
			nXPositionOfLine = nNumberPosition;
		}

		return nXPositionOfLine;
	}

	CBulletPreviewDrawerBase.prototype.getFontSizeByLineHeight = function (nLineHeight)
	{
		return ((2 * nLineHeight * AscCommonExcel.sizePxinPt) >> 0) / 2;
	};

	CBulletPreviewDrawerBase.prototype.getLvlTextWidth = function (sText, oTextPr)
	{
		const oParagraph = this.getParagraphWithText(sText, oTextPr);


		oParagraph.Reset(0, 0, 1000, 1000, 0, 0, 1);
		oParagraph.Recalculate_Page(0);
		oParagraph.LineNumbersInfo = null;

		return oParagraph.Lines[0].Ranges[0].W * AscCommon.g_dKoef_mm_to_pix;
	};

	CBulletPreviewDrawerBase.prototype.getHeadingTextInformation = function (oLvl, nTextXPosition, nTextYPosition, oColor)
	{
		if (!this.m_oLogicDocument) return null;

		const oStyles = this.m_oLogicDocument.Get_Styles();
		const oStyle = oStyles.Get(oLvl.GetPStyle());
		if (oStyle)
		{
			const sName = oStyle.Get_Name();
			const sParagraphText = " " + AscCommon.translateManager.getValue(sName);
			return {addingText: sParagraphText, startPositionX: nTextXPosition, startPositionY: nTextYPosition, color: oColor.Copy()};
		}
		return null;
	};

	CBulletPreviewDrawerBase.prototype.convertAscToNumberingLvl = function (arrAscLvl)
	{
		const arrResult = [];
		for (let i = 0; i < arrAscLvl.length; i += 1)
		{
			let oLvl;
			if (arrAscLvl[i] instanceof  Asc.CAscNumberingLvl)
			{
				oLvl = new AscCommonWord.CNumberingLvl();
				oLvl.FillFromAscNumberingLvl(arrAscLvl[i]);
			}
			else
			{
				oLvl = arrAscLvl[i].Copy();
			}

			arrResult.push(oLvl);
		}
		return arrResult;
	}

	CBulletPreviewDrawerBase.prototype.getCanvas = function (sDivId)
	{
		if (!sDivId) return;
		const oDivElement = document.getElementById(sDivId);
		const nWidth_px = oDivElement.clientWidth;
		const nHeight_px = oDivElement.clientHeight;

		let oCanvas = oDivElement.firstChild;
		if (!oCanvas)
		{
			oCanvas = document.createElement('canvas');
			oCanvas.style.cssText = "padding:0;margin:0;user-select:none;width:100%;height:100%;";
			if (nWidth_px > 0 && nHeight_px > 0)
			{
				oDivElement.appendChild(oCanvas);
			}
		}

		oCanvas.width = AscCommon.AscBrowser.convertToRetinaValue(nWidth_px, true);
		oCanvas.height = AscCommon.AscBrowser.convertToRetinaValue(nHeight_px, true);
		return oCanvas;
	}

	CBulletPreviewDrawerBase.prototype.getGraphics = function (oCanvas)
	{
		if (!oCanvas) return;
		const nHeight_px = oCanvas.clientHeight;
		const nWidth_px = oCanvas.clientWidth;
		const nRetinaWidth = oCanvas.width;
		const nRetinaHeight = oCanvas.height;
		const oContext = oCanvas.getContext("2d");

		const oGraphics = new AscCommon.CGraphics();
		oGraphics.init(oContext,
			nRetinaWidth,
			nRetinaHeight,
			nWidth_px * AscCommon.g_dKoef_pix_to_mm,
			nHeight_px * AscCommon.g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;

		oGraphics.SetIntegerGrid(true);
		oGraphics.transform(1, 0, 0, 1, 0, 0);

		if (this.m_oApi && this.m_oApi.isDarkMode)
		{
			if(this.m_oApi.getEditorId() === AscCommon.c_oEditorId.Word)
			{
				oGraphics.setDarkMode();
			}
		}

		oGraphics.b_color1(this.m_oBackgroundColor.r, this.m_oBackgroundColor.g, this.m_oBackgroundColor.b, 255);
		oGraphics.rect(0, 0, nWidth_px * AscCommon.g_dKoef_pix_to_mm, nHeight_px * AscCommon.g_dKoef_pix_to_mm);
		oGraphics.df();

		return oGraphics;
	};

	CBulletPreviewDrawerBase.prototype.getParagraphWithText = function (sText, oTextPr)
	{
		const oShape = new AscFormat.CShape();
		oShape.createTextBody();
		const oParagraph = oShape.txBody.content.GetAllParagraphs()[0];
		oParagraph.MoveCursorToStartPos();

		oParagraph.Pr = new AscCommonWord.CParaPr();
		const oParaRun = new AscCommonWord.ParaRun(oParagraph);
		oParaRun.Set_Pr(oTextPr);
		oParaRun.AddText(sText);
		oParagraph.AddToContent(0, oParaRun);

		return oParagraph;
	}


	CBulletPreviewDrawerBase.prototype.drawTextWithLvlInformation = function(sText, oLvl, nX, nY, nLineHeight, oGraphics, oParagraphTextOptions)
	{
		const oTextPr = oLvl.GetTextPr().Copy();

		const nSuff = oLvl.GetSuff();
		const nAlign = oLvl.GetJc();
		oTextPr.FontSize = oTextPr.FontSizeCS = oTextPr.FontSize || this.getFontSizeByLineHeight(nLineHeight);

		let oParagraph = this.getParagraphWithText(sText, oTextPr);
		if (!oParagraph) return null;

		oParagraph.Reset(0, 0, 1000, 1000, 0, 0, 1);
		oParagraph.Recalculate_Page(0);
		oParagraph.LineNumbersInfo = null;

		const nNumberingTextWidth = oParagraph.Lines[0].Ranges[0].W * AscCommon.g_dKoef_mm_to_pix;
		const nBaseLineOffset = oParagraph.Lines[0].Y;
		const nYOffset = nY - ((nBaseLineOffset * AscCommon.g_dKoef_mm_to_pix) >> 0);
		let nXOffset = nX;

		if (nAlign === AscCommon.align_Right) {
			nXOffset -= nNumberingTextWidth;
		} else if (nAlign === AscCommon.align_Center) {
			nXOffset -= (nNumberingTextWidth >> 1);
		}

		let nBackTextWidth = 0;
		if (nNumberingTextWidth !== 0)
		{
			nBackTextWidth = nNumberingTextWidth + 4; // 4 - чтобы линия никогда не была 'совсем рядом'
			if (nSuff === Asc.c_oAscNumberingSuff.Space ||
				nSuff === Asc.c_oAscNumberingSuff.None)
			{
				nBackTextWidth += 4;
			}
		}

		this.cleanParagraphField(oGraphics, (nXOffset - 1) * AscCommon.g_dKoef_pix_to_mm, (nY - nLineHeight) * AscCommon.g_dKoef_pix_to_mm, (nBackTextWidth + 1) * AscCommon.g_dKoef_pix_to_mm, (nLineHeight + (nLineHeight >> 1)) * AscCommon.g_dKoef_pix_to_mm);
		this.drawParagraph(oGraphics, oParagraph, nXOffset, nYOffset);

		// рисуем текст вместо черты текста
		this.drawStyleText(oGraphics, oParagraphTextOptions, nXOffset + nBackTextWidth, nY, nLineHeight, oTextPr);
	};

	CBulletPreviewDrawerBase.prototype.drawStyleText = function (oGraphics, oParagraphTextOptions, nXEndPositionOfNumbering, nY, nLineHeight, oNumberingTextPr)
	{
		if (oParagraphTextOptions)
		{
			const sParagraphText = oParagraphTextOptions.addingText;
			const oHeadingTextPr = new AscCommonWord.CTextPr();
			oHeadingTextPr.RFonts.SetAll("Arial");
			oHeadingTextPr.FontSize = oHeadingTextPr.FontSizeCS = oNumberingTextPr.FontSize * 0.8;
			oHeadingTextPr.Color = oParagraphTextOptions.color.Copy();

			const oParagraph = this.getParagraphWithText(sParagraphText, oHeadingTextPr);
			if (!oParagraph) return;

			oParagraph.Reset(0, 0, 1000, 1000, 0, 0, 1);
			oParagraph.Recalculate_Page(0);
			oParagraph.LineNumbersInfo = null;

			const nParagraphTextWidth = oParagraph.Lines[0].Ranges[0].W * AscCommon.g_dKoef_mm_to_pix;
			const nBaseLineOffset = oParagraph.Lines[0].Y;

			const nYOffset = nY - ((nBaseLineOffset * AscCommon.g_dKoef_mm_to_pix) >> 0);
			const nTextXOffset = Math.max(nXEndPositionOfNumbering, oParagraphTextOptions.startPositionX);

			this.cleanParagraphField(oGraphics, nTextXOffset * AscCommon.g_dKoef_pix_to_mm, (nY - nLineHeight) * AscCommon.g_dKoef_pix_to_mm, (nParagraphTextWidth + 2) * AscCommon.g_dKoef_pix_to_mm, (nLineHeight + (nLineHeight >> 1)) * AscCommon.g_dKoef_pix_to_mm);
			this.drawParagraph(oGraphics, oParagraph, nTextXOffset, nYOffset);
		}
	}

	CBulletPreviewDrawerBase.prototype.cleanParagraphField = function (oGraphics, nX, nY, nWidth, nHeight)
	{
		oGraphics._s();
		oGraphics.b_color1(this.m_oBackgroundColor.r, this.m_oBackgroundColor.g, this.m_oBackgroundColor.b, 255);
		oGraphics.rect(nX, nY, nWidth, nHeight);
		oGraphics.df();
		oGraphics._e();
	};

	CBulletPreviewDrawerBase.prototype.drawParagraph = function (oGraphics, oParagraph, nXOffset, nYOffset)
	{
		const oApi = this.m_oApi;

		oGraphics._s();
		oGraphics.save();
		oGraphics.SetIntegerGrid(true);

		oGraphics.m_oCoordTransform.tx = AscCommon.AscBrowser.convertToRetinaValue(nXOffset, true);
		oGraphics.m_oCoordTransform.ty = AscCommon.AscBrowser.convertToRetinaValue(nYOffset, true);

		const bOldViewMode = oApi.isViewMode;
		const bOldMarks = oApi.ShowParaMarks;
		oApi.isViewMode = true;
		oApi.ShowParaMarks = false;

		oGraphics.transform(1, 0, 0, 1, 0, 0);
		oParagraph.Draw(0, oGraphics);

		oApi.isViewMode = bOldViewMode;
		oApi.ShowParaMarks = bOldMarks;

		oGraphics.m_oCoordTransform.tx = 0;
		oGraphics.m_oCoordTransform.ty = 0;
		oGraphics.transform(1, 0, 0, 1, 0, 0);

		oGraphics.restore();
	};

	CBulletPreviewDrawerBase.prototype.checkEachLvl = function (callback)
	{
		for (let i = 0; i < this.m_arrNumberingLvl.length; i += 1)
		{
			if (Array.isArray(this.m_arrNumberingLvl[i]))
			{
				for (let j = 0; j < this.m_arrNumberingLvl[i].length; j += 1)
				{
					callback(this.m_arrNumberingLvl[i][j], j, this.m_arrNumberingLvl[i]);
				}
			}
			else
			{
				callback(this.m_arrNumberingLvl[i], i, this.m_arrNumberingLvl);
			}
		}
	};

	CBulletPreviewDrawerBase.prototype.checkFonts = function (fCallback)
	{
		const oApi = this.m_oApi;
		const oFontsDict = {};
		const oThis = this;
		this.checkEachLvl(function (oLvl) {
			const sText = oLvl.GetSymbols();
			if (sText)
			{
				AscFonts.FontPickerByCharacter.checkTextLight(sText);
			}
			const oTextPr = oLvl.GetTextPr();
			oThis.cleanTextPr(oTextPr);
			if (oTextPr && oTextPr.RFonts)
			{
				if (oTextPr.RFonts.Ascii) oFontsDict[oTextPr.RFonts.Ascii.Name] = true;
				if (oTextPr.RFonts.EastAsia) oFontsDict[oTextPr.RFonts.EastAsia.Name] = true;
				if (oTextPr.RFonts.HAnsi) oFontsDict[oTextPr.RFonts.HAnsi.Name] = true;
				if (oTextPr.RFonts.CS) oFontsDict[oTextPr.RFonts.CS.Name] = true;
			}
		});

		const arrFonts = [];
		for (let sFamilyName in oFontsDict)
		{
			arrFonts.push(new AscFonts.CFont(AscFonts.g_fontApplication.GetFontInfoName(sFamilyName)));
		}
		AscFonts.FontPickerByCharacter.extendFonts(arrFonts);

		if (false === AscCommon.g_font_loader.CheckFontsNeedLoading(arrFonts))
		{
			return fCallback();
		}

		const oLoader = new AscCommon.CGlobalFontLoader();
		oLoader.put_Api(oApi);
		oLoader.LoadDocumentFonts2(arrFonts, Asc.c_oAscAsyncActionType.Information, fCallback);
	};

	CBulletPreviewDrawerBase.prototype.draw = function () {};

	CBulletPreviewDrawerBase.prototype.checkFontsAndDraw = function ()
	{
		const oThis = this;
		this.checkFonts(function ()
		{

			if (oThis.m_oLogicDocument && oThis.m_oLogicDocument.IsDocumentEditor())
			{
				const bIsOldTrackRevisions = oThis.m_oLogicDocument.GetLocalTrackRevisions();
				oThis.m_oLogicDocument.SetTrackRevisions(false);
				oThis.draw();
				oThis.m_oLogicDocument.SetTrackRevisions(bIsOldTrackRevisions);
			}
			else
			{
				oThis.draw();
			}
		});
	};


	function CBulletPreviewDrawer(arrLvlInfo, nType)
	{
		CBulletPreviewDrawerBase.call(this);
		this.m_nType = nType;
		this.m_nCountOfLines = 3;
		this.m_oApi = editor || Asc.editor || window["Asc"]["editor"];
		this.m_arrNumberingLvl = arrLvlInfo.map(function (oDrawingInfo) {return oDrawingInfo.arrLvls});
		this.m_arrNumberingInfo = arrLvlInfo;
		if (this.m_bIsMobile)
		{
			this.m_nSingleBulletNoneFontSizeCoefficient = 0.21;
			this.m_nLvlWithLinesNoneFontSizeCoefficient = 0.21;
			this.m_nSingleBulletFontSizeCoefficient = 6 / 17;
		}
		else
		{
			this.m_nSingleBulletFontSizeCoefficient = 0.6;
			this.m_nSingleBulletNoneFontSizeCoefficient = 0.225;
			this.m_nLvlWithLinesNoneFontSizeCoefficient = 0.1375;
		}

		this.m_nMultiLvlIndentCoefficient = 1 / AscCommon.AscBrowser.retinaPixelRatio;
	}
	CBulletPreviewDrawer.prototype = Object.create(CBulletPreviewDrawerBase.prototype);
	CBulletPreviewDrawer.prototype.constructor = CBulletPreviewDrawer;

	CBulletPreviewDrawer.prototype.drawSingleBullet = function (sDivId, arrLvls)
	{
		const oCanvas = this.getCanvas(sDivId);
		if (!oCanvas) return;

		const oGraphics = this.getGraphics(oCanvas);
		const oLvl = arrLvls[0];
		const nHeight_px = oCanvas.clientHeight;
		const nWidth_px = oCanvas.clientWidth;
		const drawingContent = oLvl.GetDrawingContent([oLvl], 0, undefined, this.m_oLang);
		if (typeof drawingContent !== "string")
		{
			const oImage = drawingContent.image;
			if (oImage)
			{
				const oFormatBullet = new AscFormat.CBullet();
				oFormatBullet.fillBulletImage(oImage.src);
				oFormatBullet.drawSquareImage(sDivId, 0.125);
			}
		}
		else
		{
			const nMaxFontSize = nHeight_px * this.m_nSingleBulletFontSizeCoefficient;
			// для буллетов решено не уменьшать их превью, как и в word
			//const oFitInformation = this.getInformationWithFitFontSize(oLvl, nWidth_px * AscCommon.g_dKoef_pix_to_mm, nHeight_px * AscCommon.g_dKoef_pix_to_mm, nMaxFontSize, nMaxFontSize);
			//const oFitTextPr = oFitInformation.textPr;
			const oTextPr = oLvl.GetTextPr();
			oTextPr.FontSize = oTextPr.FontSizeCS = nMaxFontSize;
			oLvl.SetJc(AscCommon.align_Left);
			const oCalculationPosition = this.getXYForCenterPosition(oLvl, nWidth_px, nHeight_px);
			const nX = oCalculationPosition.nX;
			const nY = oCalculationPosition.nY;
			const nLineHeight = oCalculationPosition.nLineHeight;
			const sText = drawingContent;
			this.drawTextWithLvlInformation(sText, oLvl, nX, nY, nLineHeight, oGraphics);
		}
	};

	CBulletPreviewDrawer.prototype.getInformationWithFitFontSize = function (oLvl, nMaxWidth, nMaxHeight, nMinFontSize, nMaxFontSize)
	{
		const sText = oLvl.GetDrawingContent([oLvl], 0, undefined, this.m_oLang);
		if (typeof sText !== "string") return;
		const oNewShape = new AscFormat.CShape();
		oNewShape.createTextBody();
		oNewShape.extX = nMaxWidth;
		oNewShape.extY = nMaxHeight;
		oNewShape.contentWidth = oNewShape.extX;
		oNewShape.setPaddings({Left: 0, Top: 0, Right: 0, Bottom: 0});

		const oParagraph = oNewShape.txBody.content.GetAllParagraphs()[0];
		oParagraph.MoveCursorToStartPos();

		oParagraph.Pr = new AscCommonWord.CParaPr();

		const oParaRun = new AscCommonWord.ParaRun(oParagraph);
		const oTextPr = oLvl.GetTextPr().Copy();
		oParaRun.Set_Pr(oTextPr);
		oParaRun.AddText(sText);
		oParagraph.AddToContent(0, oParaRun);

		oTextPr.FontSize = nMaxFontSize;

		oParagraph.TextPr.SetFontSize(oTextPr.FontSize);
		// TODO: add function after merge, add set new font size
		let nParagraphWidth = oParagraph.RecalculateMinMaxContentWidth().Max;
		if (nParagraphWidth > oNewShape.contentWidth) {
			const nNewFontSize = oNewShape.findFitFontSize(nMinFontSize, nMaxFontSize, true);
			if (nNewFontSize !== null)
			{
				oNewShape.setFontSizeInSmartArt(nNewFontSize);
			}
		}

		return oTextPr;
	};


	CBulletPreviewDrawer.prototype.getWidthHeightGlyphs = function (sText, oTextPr)
	{
		AscCommon.g_oTextMeasurer.SetTextPr(oTextPr);
		const oSumInformation = {Width: 0, rasterOffsetX: 0, Height: 0, Ascent: 0, rasterOffsetY: 0};
		let bFirstGlyphSymbol = false;
		let nRemoveRightOffset = 0;
		for (const oIterator = sText.getUnicodeIterator(); oIterator.check(); oIterator.next())
		{
			const nValue = oIterator.value();
			const nFontSlot = AscWord.GetFontSlotByTextPr(nValue, oTextPr);
			AscCommon.g_oTextMeasurer.SetFontSlot(nFontSlot, 1);
			const oInfo = AscCommon.g_oTextMeasurer.Measure2Code(nValue);

			// в ворде крайние пробелы в превью буллетов прижимаются к глифу, а не к ширине символа
			if (!bFirstGlyphSymbol)
			{
				if (oInfo.WidthG)
				{
					oSumInformation.rasterOffsetX = oInfo.rasterOffsetX;
					const nWidthWithRightOffset = oInfo.Width - oInfo.rasterOffsetX;
					oSumInformation.Width += nWidthWithRightOffset;
					nRemoveRightOffset = nWidthWithRightOffset - oInfo.WidthG;
					bFirstGlyphSymbol = true;
				}
				else
				{
					oSumInformation.Width += oInfo.Width;
				}

			}
			else
			{
				oSumInformation.Width += oInfo.Width;
				if (oInfo.WidthG)
				{
					nRemoveRightOffset = oInfo.Width - oInfo.rasterOffsetX - oInfo.WidthG;
				}
			}

			if (oSumInformation.Ascent < oInfo.Ascent)
			{
				oSumInformation.Ascent = oInfo.Ascent;
			}
			if (oSumInformation.Ascent - oSumInformation.Height > oInfo.Ascent - oInfo.Height)
			{
				oSumInformation.rasterOffsetY = oInfo.rasterOffsetY;
				oSumInformation.Height = oSumInformation.Ascent - (oInfo.Ascent - oInfo.Height);
			}
		}
		oSumInformation.Width -= nRemoveRightOffset;
		return oSumInformation;
	};

	CBulletPreviewDrawer.prototype.getXYForCenterPosition = function (oLvl, nWidth, nHeight)
	{
		// Здесь будем считать позицию отрисовки
		const sText = oLvl.GetDrawingContent([oLvl], 0, undefined, this.m_oLang);
		if (typeof sText !== 'string') return;
		const oTextPr = oLvl.GetTextPr().Copy();
		const oSumInformation = this.getWidthHeightGlyphs(sText, oTextPr);

		const nX = (nWidth >> 1) - Math.round((oSumInformation.Width / 2 + oSumInformation.rasterOffsetX) * AscCommon.g_dKoef_mm_to_pix);
		const nY = (nHeight >> 1) + Math.round((oSumInformation.Height / 2 + (oSumInformation.Ascent - oSumInformation.Height + oSumInformation.rasterOffsetY)) * AscCommon.g_dKoef_mm_to_pix);
		return {nX: nX, nY: nY, nLineHeight: oSumInformation.Height};
	};

	CBulletPreviewDrawer.prototype.drawSingleLvlWithLines = function (sDivId, arrLvls)
	{
		const nCountOfLines = this.m_nCountOfLines;
		const oCanvas = this.getCanvas(sDivId);
		if (!oCanvas) return;

		const oGraphics = this.getGraphics(oCanvas);
		oGraphics.p_color(this.m_oSecondaryLineTextColor.r, this.m_oSecondaryLineTextColor.g, this.m_oSecondaryLineTextColor.b, 255);

		const oLvl = arrLvls[0];
		oLvl.SetJc(AscCommon.align_Left);

		const oTextPr = oLvl.GetTextPr();

		const nWidth_px = oCanvas.clientWidth;
		const nHeight_px = oCanvas.clientHeight;
		const oContext = oCanvas.getContext("2d");
		oContext.beginPath();

		const nOffsetBase = 4;
		const nLineWidth = 2;
		// считаем расстояние между линиями
		const nLineDistance = Math.floor(((nHeight_px - (nOffsetBase << 2)) - nLineWidth * nCountOfLines) / nCountOfLines);
		// убираем погрешность в offset
		const nOffset = (nHeight_px - (nLineWidth * nCountOfLines + nLineDistance * nCountOfLines)) >> 1;

		const nTextBaseOffsetX = nOffset + Math.floor(2.25 * AscCommon.g_dKoef_mm_to_pix);

		let nY = nOffset + 11;
		for (let j = 0; j < nCountOfLines; j += 1)
		{
			const nYmm = Math.round(nY) * AscCommon.g_dKoef_pix_to_mm;
			const nTextBaseXmm = Math.round(nTextBaseOffsetX) * AscCommon.g_dKoef_pix_to_mm;
			const nWidthmm = Math.round((nWidth_px - nOffsetBase)) * AscCommon.g_dKoef_pix_to_mm;
			const nWidthLinemm = 2 * AscCommon.g_dKoef_pix_to_mm;

			oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nYmm, nTextBaseXmm, nWidthmm, nWidthLinemm);
			const nTextYx =  nTextBaseOffsetX - Math.floor(3.25 * AscCommon.g_dKoef_mm_to_pix);
			const nTextYy = nY + (nLineWidth * 2.5);
			const nLineHeight = nLineDistance - 4;
			oTextPr.FontSize = this.getFontSizeByLineHeight(nLineHeight);
			const drawingContent = oLvl.GetDrawingContent([oLvl], 0, j + 1, this.m_oLang);
			if (typeof drawingContent !== "string")
			{
				this.drawImageBulletsWithLine(drawingContent, nTextYx, nTextYy, nLineHeight, oGraphics);
			}
			else
			{
				this.drawTextWithLvlInformation(drawingContent, oLvl, nTextYx, nTextYy, nLineHeight, oGraphics);
			}
			nY += (nLineWidth + nLineDistance);
		}
		this.cleanParagraphField(oGraphics, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, 0, nWidth_px * AscCommon.g_dKoef_pix_to_mm, nHeight_px * AscCommon.g_dKoef_pix_to_mm);
	};

	CBulletPreviewDrawer.prototype.drawNoneTextPreview = function (sDivId, arrLvls, nFontSizeCoefficient)
	{
		const oCanvas = this.getCanvas(sDivId);
		if (!oCanvas) return;
		const oGraphics = this.getGraphics(oCanvas);

		const oLvl = arrLvls[0];
		const sText = oLvl.GetDrawingContent([oLvl], 0, undefined, this.m_oLang);
		if (typeof sText !== 'string') return;
		const nHeight_px = oCanvas.clientHeight;
		const nWidth_px = oCanvas.clientWidth;
		const nMaxFontSize = nWidth_px * nFontSizeCoefficient;

		const oFitTextPr = this.getInformationWithFitFontSize(oLvl, nWidth_px * AscCommon.g_dKoef_pix_to_mm, nHeight_px * AscCommon.g_dKoef_pix_to_mm, 5, nMaxFontSize);
		oLvl.SetTextPr(oFitTextPr);
		const oCalculationPosition = this.getXYForCenterPosition(oLvl, nWidth_px, nHeight_px);
		const nX = oCalculationPosition.nX;
		const nY = oCalculationPosition.nY;
		const nLineHeight = oCalculationPosition.nLineHeight;

		this.drawTextWithLvlInformation(sText, oLvl, nX, nY, nLineHeight, oGraphics);
	};

	CBulletPreviewDrawer.prototype.getHeadingTextInformation = function (oLvl, nTextXPosition, nTextYPosition)
	{
		return CBulletPreviewDrawerBase.prototype.getHeadingTextInformation.call(this, oLvl, nTextXPosition, nTextYPosition, this.m_oSecondaryTextColor.Copy());
	};

	CBulletPreviewDrawer.prototype.getMultiLvlAddedOffsetX = function (arrLvls)
	{
		let nMinNumberPosition = arrLvls[0].GetNumberPosition();
		let nMinTextIndent = arrLvls[0].GetIndentSize();
		for (let i = 1; i < arrLvls.length; i += 1)
		{
			const oLvl = arrLvls[i];
			const nCurrentNumberPosition = oLvl.GetNumberPosition();
			const nCurrentIndentSize = oLvl.GetIndentSize();
			if (nCurrentNumberPosition < nMinNumberPosition)
			{
				nMinNumberPosition = nCurrentNumberPosition;
			}
			if (nCurrentIndentSize < nMinTextIndent)
			{
				nMinTextIndent = nCurrentIndentSize;
			}
		}
		return -Math.min(nMinTextIndent, nMinNumberPosition);
	};

	CBulletPreviewDrawer.prototype.drawMultiLevelBullet = function (sDivId, arrLvls)
	{
		const nCountOfLines = this.m_nCountOfLines;
		const oCanvas = this.getCanvas(sDivId);
		if (!oCanvas) return;

		const oGraphics = this.getGraphics(oCanvas);
		oGraphics.p_color(this.m_oSecondaryLineTextColor.r, this.m_oSecondaryLineTextColor.g, this.m_oSecondaryLineTextColor.b, 255);

		const nHeight_px = oCanvas.clientHeight;
		const nWidth_px = oCanvas.clientWidth;

		const nOffsetBase = 4;
		const nLineWidth = 2;
		const nLineDistance = Math.floor(((nHeight_px - (nOffsetBase << 2)) - nLineWidth * nCountOfLines) / nCountOfLines);
		const nLineHeight = nLineDistance - 4;
		const nOffset = (nHeight_px - (nLineWidth * nCountOfLines + nLineDistance * nCountOfLines)) >> 1;
		let nY = nOffset + 11;
		const nCorrectAddedOffsetX = this.getMultiLvlAddedOffsetX(arrLvls);
		for (let i = 0; i < nCountOfLines; i += 1)
		{
			const oLvl = arrLvls[i];
			oLvl.SetJc(AscCommon.align_Left);
			const oTextPr = oLvl.GetTextPr();
			oTextPr.FontSize = this.getFontSizeByLineHeight(nLineHeight);
			const nNumberPosition = oLvl.GetNumberPosition() + nCorrectAddedOffsetX;
			const nFirstLineIndent = this.getFirstLineIndent(oLvl) + nCorrectAddedOffsetX;
			const nTextYx =  nOffsetBase + nNumberPosition * this.m_nMultiLvlIndentCoefficient;
			const nTextYy = nY + (nLineWidth * 2.5);
			const nXPositionOfLine = nOffsetBase + (nFirstLineIndent * this.m_nMultiLvlIndentCoefficient) << 0;

			oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nXPositionOfLine * AscCommon.g_dKoef_pix_to_mm, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm);

			const drawingContent = oLvl.GetDrawingContent(arrLvls, i, 1, this.m_oLang);
			const oParagraphTextOptions = this.getHeadingTextInformation(oLvl, nXPositionOfLine, nTextYy);
			if (typeof drawingContent !== 'string')
			{
				this.drawImageBulletsWithLine(drawingContent, nTextYx, nTextYy, nLineHeight, oGraphics, oParagraphTextOptions, oTextPr);
			}
			else
			{
				this.drawTextWithLvlInformation(drawingContent, oLvl, nTextYx, nTextYy, nLineHeight, oGraphics, oParagraphTextOptions);
			}
			nY += (nLineWidth + nLineDistance);
		}
		this.cleanParagraphField(oGraphics, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, 0, nWidth_px * AscCommon.g_dKoef_pix_to_mm, nHeight_px * AscCommon.g_dKoef_pix_to_mm);
	};

	CBulletPreviewDrawer.prototype.draw = function ()
	{
		AscFormat.ExecuteNoHistory(function () {
			for (let i = 0; i < this.m_arrNumberingInfo.length; i++)
			{
				const oDrawingInfo = this.m_arrNumberingInfo[i];
				const sId = oDrawingInfo.divId;
				const arrLvls = oDrawingInfo.arrLvls;
				
				if (!arrLvls || !arrLvls.length)
					continue;
				
				if (this.m_nType === 0)
				{
					if (oDrawingInfo.isRemoving)
					{
						this.drawNoneTextPreview(sId, arrLvls, this.m_nSingleBulletNoneFontSizeCoefficient);
					}
					else
					{
						this.drawSingleBullet(sId, arrLvls);
					}
				}
				else if (this.m_nType === 1)
				{
					if (oDrawingInfo.isRemoving)
					{
						this.drawNoneTextPreview(sId, arrLvls, this.m_nLvlWithLinesNoneFontSizeCoefficient);
					}
					else
					{
						this.drawSingleLvlWithLines(sId, arrLvls);
					}
				}
				else if (this.m_nType === 2)
				{
					if (oDrawingInfo.isRemoving)
					{
						this.drawNoneTextPreview(sId, arrLvls, this.m_nLvlWithLinesNoneFontSizeCoefficient);
					}
					else
					{
						this.drawMultiLevelBullet(sId, arrLvls);
					}
				}
			}
		}, this, []);


	};

	function CBulletPreviewDrawerChangeList(arrId, arrAscLvl) {
		CBulletPreviewDrawerBase.call(this);
		this.m_arrNumberingLvl = this.convertAscToNumberingLvl(arrAscLvl.Lvl);
		this.m_arrId = arrId;
	}
	CBulletPreviewDrawerChangeList.prototype = Object.create(CBulletPreviewDrawerBase.prototype);
	CBulletPreviewDrawerChangeList.prototype.constructor = CBulletPreviewDrawerChangeList;

	CBulletPreviewDrawerChangeList.prototype.getHeadingTextInformation = function (oLvl, nTextXPosition, nTextYPosition)
	{
		return CBulletPreviewDrawerBase.prototype.getHeadingTextInformation.call(this, oLvl, nTextXPosition, nTextYPosition, this.m_oSecondaryTextColor.Copy());
	};

	CBulletPreviewDrawerChangeList.prototype.getScaleCoefficientForMultiLevel = function (arrLvl, nWorkspaceWidth)
	{
		let nMaxNumberPosition = arrLvl[0].GetNumberPosition();
		for (let i = 1; i < arrLvl.length; i += 1)
		{
			const oLvl = arrLvl[i];
			const nNumberPosition = oLvl.GetNumberPosition();
			if (nMaxNumberPosition < nNumberPosition)
			{
				nMaxNumberPosition = nNumberPosition;
			}
		}

		const nNumberPositionScale = nWorkspaceWidth / (nMaxNumberPosition * AscCommon.g_dKoef_mm_to_pix);
		const nThresholdScaleCoefficient = 0.3 / AscCommon.AscBrowser.retinaPixelRatio;
		if (nNumberPositionScale < nThresholdScaleCoefficient)
		{
			return nNumberPositionScale;
		}
		return nThresholdScaleCoefficient;
	};
	CBulletPreviewDrawerChangeList.prototype.draw = function ()
	{
		AscFormat.ExecuteNoHistory(function ()
		{
			const nAmountOfPreview = Math.min(this.m_arrNumberingLvl.length, this.m_arrId.length);
			const nOffsetBase = 5;
			const nLineWidth = 2;

			// посчитаем нужные переменные для одного canvas
			let sDivId = this.m_arrId[0];
			let oCanvas = this.getCanvas(sDivId);
			const nHeight_px = oCanvas.clientHeight;
			const nWidth_px = oCanvas.clientWidth;

			const nY = (nHeight_px >> 1) - (nLineWidth >> 1);
			const nLineHeight = (nHeight_px >> 1);
			const nScaleCoefficient = this.getScaleCoefficientForMultiLevel(this.m_arrNumberingLvl, nWidth_px - nOffsetBase * 6);

			for (let i = 0; i < nAmountOfPreview; i += 1)
			{
				const oLvl = this.m_arrNumberingLvl[i];
				oLvl.Jc = AscCommon.align_Left;
				const drawingContent = oLvl.GetDrawingContent(this.m_arrNumberingLvl, i, 1, this.m_oLang);
				sDivId = this.m_arrId[i];
				oCanvas = this.getCanvas(sDivId);
				if (!oCanvas) return;
				const oGraphics = this.getGraphics(oCanvas);
				oGraphics.p_color(this.m_oSecondaryLineTextColor.r, this.m_oSecondaryLineTextColor.g, this.m_oSecondaryLineTextColor.b, 255);

				const oTextPr = oLvl.GetTextPr();
				oTextPr.FontSize = oTextPr.FontSizeCS = this.getFontSizeByLineHeight(nLineHeight);

				const nNumberPosition = oLvl.GetNumberPosition();
				const nXLinePosition = nOffsetBase + (this.getFirstLineIndent(oLvl) * AscCommon.g_dKoef_mm_to_pix * nScaleCoefficient) << 0;
				const nTextYx = nOffsetBase + nNumberPosition * AscCommon.g_dKoef_mm_to_pix * nScaleCoefficient;
				const nTextYy = nY + (nLineWidth << 1);
				oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nXLinePosition * AscCommon.g_dKoef_pix_to_mm, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm);
				const oParagraphTextOptions = this.getHeadingTextInformation(oLvl, nXLinePosition, nTextYy);
				if (typeof drawingContent === "string")
				{
					this.drawTextWithLvlInformation(drawingContent, oLvl, nTextYx, nTextYy, (nHeight_px >> 1), oGraphics, oParagraphTextOptions);
				}
				else
				{
					if (drawingContent.image)
					{
						this.drawImageBulletsWithLine(drawingContent, nTextYx, nTextYy, (nHeight_px >> 1), oGraphics, oParagraphTextOptions, oTextPr);
					}
				}

				this.cleanParagraphField(oGraphics, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, 0, nWidth_px * AscCommon.g_dKoef_pix_to_mm, nHeight_px * AscCommon.g_dKoef_pix_to_mm);
			}
		}, this, []);

	};

	function CBulletPreviewDrawerAdvancedOptions(sDivId, props, nLvl, bIsMultiLvlAdvanceOptions)
	{
		CBulletPreviewDrawerBase.call(this);
		this.m_sId = sDivId;
		this.m_arrNumberingLvl = this.convertAscToNumberingLvl(props.Lvl);
		this.m_nCurrentLvl = nLvl;
		this.m_bIsMultiLvl = bIsMultiLvlAdvanceOptions;
		this.m_oCanvas = this.getCanvas(this.m_sId);
		this.m_oGraphics = this.getGraphics(this.m_oCanvas);
		this.m_nScaleIndentsCoefficient = 0.55;
		this.m_arrCalcNumberingInfo = null;
		this.m_nCalcNumberingLvl = -1;
		this.initNumberingInfo();
	}
	CBulletPreviewDrawerAdvancedOptions.prototype = Object.create(CBulletPreviewDrawerBase.prototype);
	CBulletPreviewDrawerAdvancedOptions.prototype.constructor = CBulletPreviewDrawerAdvancedOptions;

	CBulletPreviewDrawerAdvancedOptions.prototype.addControlMultiLvl = function ()
	{
		if (!this.m_bIsMultiLvl || !this.m_oCanvas) return;
		const oThis = this;
		AscCommon.addMouseEvent(this.m_oCanvas, "down", function(e) {
			AscCommon.stopEvent(e);
			const nOffsetBase = 10;
			const nLineWidth = 4;
			const nHeight = oThis.m_oCanvas.clientHeight;
			const nLineDistance = Math.floor(((nHeight - (nOffsetBase << 1)) - nLineWidth * 10) / 9);
			const nOffset = (nHeight - (nLineWidth * 10 + nLineDistance * 9)) >> 1;
			const nCurrentLvl = oThis.m_nCurrentLvl;

			let nYPos = e.pageY;
			if (!AscFormat.isRealNumber(nYPos))
			{
				nYPos = e.clientY;
			}
			nYPos = (nYPos * AscCommon.AscBrowser.zoom);
			const oClientRect = AscCommon.UI.getBoundingClientRect(this);

			if (AscFormat.isRealNumber(oClientRect.y))
			{
				nYPos -= oClientRect.y;
			}
			else if (AscFormat.isRealNumber(oClientRect.top))
			{
				nYPos -= oClientRect.top;
			}

			let nChangedCurrentLvl = 8;
			let nY = nOffset + 2;
			for (let i = 0; i < oThis.m_arrNumberingLvl.length; i++)
			{
				nY += (nLineWidth + nLineDistance);
				if (i === nCurrentLvl)
				{
					nY += (nLineWidth + nLineDistance);
				}

				if (nYPos < (nY - ((nLineWidth + nLineDistance) >> 1)))
				{
					nChangedCurrentLvl = i;
					break;
				}
			}
			oThis.m_oApi.sendEvent("asc_onPreviewLevelChange", nChangedCurrentLvl);
		});
	};

	CBulletPreviewDrawerAdvancedOptions.prototype.getHeadingTextInformation = function (oLvl, nTextXPosition, nTextYPosition)
	{
		return CBulletPreviewDrawerBase.prototype.getHeadingTextInformation.call(this, oLvl, nTextXPosition, nTextYPosition, this.m_oPrimaryTextColor.Copy());
	};
	CBulletPreviewDrawerAdvancedOptions.prototype.getNumberingValue = function (nNumberIndex, nDrawingLvl, oLvl)
	{
		if (nDrawingLvl <= this.m_nCalcNumberingLvl && this.m_arrCalcNumberingInfo && AscFormat.isRealNumber(this.m_arrCalcNumberingInfo[nDrawingLvl]) && AscFormat.isRealNumber(this.m_nSourceStart))
		{
			const nCalcValue = this.m_arrCalcNumberingInfo[nDrawingLvl];
			return nCalcValue - this.m_nSourceStart + nNumberIndex;
		}
		return nNumberIndex;
	};
	CBulletPreviewDrawerAdvancedOptions.prototype.drawMultiLvlAdvancedOptions = function ()
	{
		const oCanvas = this.m_oCanvas;
		if (!oCanvas) return;

		const oGraphics = this.m_oGraphics;
		const nHeight_px = oCanvas.clientHeight;
		const nWidth_px = oCanvas.clientWidth;

		const nOffsetBase = 10;
		const nLineWidth = 4;
		// считаем расстояние между линиями
		const nLineDistance = Math.floor(((nHeight_px - (nOffsetBase << 1)) - nLineWidth * 10) / 9);
		// убираем погрешность в offset
		const nOffset = (nHeight_px - (nLineWidth * 10 + nLineDistance * 9)) >> 1;
		const nCurrentLvl = this.m_nCurrentLvl;

		oGraphics.p_color(this.m_oSecondaryLineTextColor.r, this.m_oSecondaryLineTextColor.g, this.m_oSecondaryLineTextColor.b, 255);

		let nY = nOffset + 2;
		for (let i = 0; i < this.m_arrNumberingLvl.length; i += 1)
		{
			const oLvl = this.m_arrNumberingLvl[i];
			const oTextPr = oLvl.GetTextPr();
			oTextPr.FontSize = this.getFontSizeByLineHeight(nLineDistance);
			const nNumberPosition = oLvl.GetNumberPosition();
			const nTextYx = (nOffsetBase + nNumberPosition * AscCommon.g_dKoef_mm_to_pix * this.m_nScaleIndentsCoefficient) >> 0;
			const nIndentSize = (nOffsetBase + oLvl.GetIndentSize() * AscCommon.g_dKoef_mm_to_pix * this.m_nScaleIndentsCoefficient) >> 0;
			const nTextYy = nY + nLineWidth;
			const nOffsetText = nOffsetBase + (this.getFirstLineIndent(oLvl) * AscCommon.g_dKoef_mm_to_pix * this.m_nScaleIndentsCoefficient) >> 0;

			if (i === nCurrentLvl)
			{
				oGraphics.p_color(this.m_oPrimaryTextColor.r, this.m_oPrimaryTextColor.g, this.m_oPrimaryTextColor.b, 255);

				oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nOffsetText * AscCommon.g_dKoef_pix_to_mm, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm);
				nY += (nLineWidth + nLineDistance);
				oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nIndentSize * AscCommon.g_dKoef_pix_to_mm, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm);

				oGraphics.p_color(this.m_oSecondaryLineTextColor.r, this.m_oSecondaryLineTextColor.g, this.m_oSecondaryLineTextColor.b, 255);
			}
			else
			{
				oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nOffsetText * AscCommon.g_dKoef_pix_to_mm, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm);
			}

			const oParagraphTextOptions = this.getHeadingTextInformation(oLvl, nOffsetText, nTextYy);
			let nNumberIndex = this.getNumberingValue(1, i, oLvl);
			const drawingContent = oLvl.GetDrawingContent(this.m_arrNumberingLvl, i, nNumberIndex, this.m_oLang);
			if (typeof drawingContent === "string")
			{
				this.drawTextWithLvlInformation(drawingContent, oLvl, nTextYx,  nTextYy, nLineDistance, oGraphics, oParagraphTextOptions);
			}
			else
			{
				if (drawingContent.image)
				{
					this.drawImageBulletsWithLine(drawingContent, nTextYx, nTextYy, nLineDistance, oGraphics, oParagraphTextOptions, oTextPr);
				}
			}

			nY += (nLineWidth + nLineDistance);
		}
		this.cleanParagraphField(oGraphics, (nWidth_px - nOffsetBase) * AscCommon.g_dKoef_pix_to_mm, 0, nWidth_px * AscCommon.g_dKoef_pix_to_mm, nHeight_px * AscCommon.g_dKoef_pix_to_mm);

	};
	CBulletPreviewDrawerAdvancedOptions.prototype.getScaleCoefficientForSingleLevel = function (nWorkspaceWidth)
	{
		const nCurrentLvl = this.m_nCurrentLvl;
		const oCurrentLvl = this.m_arrNumberingLvl[nCurrentLvl];

		const nNumberPosition = Math.round(oCurrentLvl.GetNumberPosition() * AscCommon.g_dKoef_mm_to_pix);
		const nIndentSize = Math.round(oCurrentLvl.GetIndentSize() * AscCommon.g_dKoef_mm_to_pix);
		const nTabSize = Math.round(oCurrentLvl.GetStopTab() * AscCommon.g_dKoef_mm_to_pix);

		const nNumberPositionScaleCoefficient = nWorkspaceWidth / nNumberPosition;
		const nIndentSizeScaleCoefficient = nWorkspaceWidth / nIndentSize;
		const nTabSizeScaleCoefficient = nWorkspaceWidth / nTabSize;
		const nScaleCoefficient = Math.min(nNumberPositionScaleCoefficient, nIndentSizeScaleCoefficient, nTabSizeScaleCoefficient);
		if (nScaleCoefficient < 1)
		{
			return nScaleCoefficient;
		}
		return 1;
	};
	CBulletPreviewDrawerAdvancedOptions.prototype.initNumberingInfo = function ()
	{
		const oParagraph = this.m_oLogicDocument.GetCurrentParagraph(true);
		if (!oParagraph)
			return;

		const oNumbering = oParagraph.Numbering;
		if (oNumbering)
		{
			this.m_arrCalcNumberingInfo = oNumbering.GetCalculatedNumInfo();
			this.m_nCalcNumberingLvl = oNumbering.GetCalculatedNumberingLvl();
			const oNum = this.m_oLogicDocument.GetNumbering().GetNum(oNumbering.GetCalculatedNumId());
			if (oNum)
			{
				const oLvl = oNum.GetLvl(this.m_nCalcNumberingLvl);
				this.m_nSourceStart = oLvl ? oLvl.GetStart() : null;
			}
		}
	};
	CBulletPreviewDrawerAdvancedOptions.prototype.drawSingleLvlAdvancedOptions = function ()
	{
		const oCanvas = this.m_oCanvas;
		if (!oCanvas) return;
		const oGraphics = this.m_oGraphics;
		const nHeight_px = oCanvas.clientHeight;
		const nWidth_px = oCanvas.clientWidth;

		const nOffsetBase = 10;
		const nLineWidth = 4;
		const nCurrentLvl = this.m_nCurrentLvl;
		const oCurrentLvl = this.m_arrNumberingLvl[nCurrentLvl];
		const oTextPr = oCurrentLvl.GetTextPr();

		const nLineDistance = (((nHeight_px - (nOffsetBase << 1)) - nLineWidth * 10) / 9) << 0;
		oTextPr.FontSize = oTextPr.FontSizeCS = this.getFontSizeByLineHeight(nLineDistance);
		let nMaxTextWidth = 0;
		for (let i = 0; i < 3; i += 1)
		{
			const drawingContent = oCurrentLvl.GetDrawingContent(this.m_arrNumberingLvl, nCurrentLvl, i + 1, this.m_oLang);
			if (typeof drawingContent === 'string')
			{
				const nTextWidth = this.getLvlTextWidth(drawingContent, oTextPr);
				if (nMaxTextWidth < nTextWidth)
				{
					nMaxTextWidth = nTextWidth;
				}
			}
			else
			{
				if (drawingContent.image)
				{
					const sFullImageSrc = drawingContent.image.src;
					const oSizes = AscCommon.getSourceImageSize(sFullImageSrc);
					const nImageHeight = oSizes.height;
					const nImageWidth = oSizes.width;
					nMaxTextWidth = (nImageWidth * nLineDistance / (nImageHeight ? nImageHeight : 1)) * drawingContent.amount;
					break;
				}
			}

		}
		nMaxTextWidth = nMaxTextWidth >> 0;
		const nOffset = (nHeight_px - (nLineWidth * 10 + nLineDistance * 9)) >> 1;

		oGraphics.p_color(this.m_oSecondaryLineTextColor.r, this.m_oSecondaryLineTextColor.g, this.m_oSecondaryLineTextColor.b, 255);

		let nY = nOffset + 2 + 2 * (nLineWidth + nLineDistance);

		const arrTextYy = [];
		arrTextYy.push(nY + nLineWidth); nY += 2 * (nLineWidth + nLineDistance);
		arrTextYy.push(nY + nLineWidth); nY += 2 * (nLineWidth + nLineDistance);
		arrTextYy.push(nY + nLineWidth);

		nY = nOffset + 2;
		const nRightOffset = nWidth_px - nOffsetBase;
		const nYDist = nLineWidth + nLineDistance;

		const nLeftOffset2 = nOffsetBase;
		const nRightOffset2 = nWidth_px - nOffsetBase;

		// Здесь получаем коэффициент, чтобы при открытии всегда видеть отступ текста
		const nScaleCoefficient = this.getScaleCoefficientForSingleLevel(nWidth_px - nOffsetBase * 5);
		let nNumberPosition = nOffsetBase + ((oCurrentLvl.GetNumberPosition() * AscCommon.g_dKoef_mm_to_pix * nScaleCoefficient) << 0);
		let nIndentSize = nOffsetBase + ((oCurrentLvl.GetIndentSize() * AscCommon.g_dKoef_mm_to_pix * nScaleCoefficient) << 0);
		const nRawTabSize = oCurrentLvl.GetStopTab();
		let nTabSize;
		if (AscFormat.isRealNumber(nRawTabSize))
		{
			nTabSize = nOffsetBase + ((nRawTabSize * AscCommon.g_dKoef_mm_to_pix * nScaleCoefficient) << 0);
		}

		oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nLeftOffset2 * AscCommon.g_dKoef_pix_to_mm, nRightOffset2 * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm); nY += nYDist;
		oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nLeftOffset2 * AscCommon.g_dKoef_pix_to_mm, nRightOffset2 * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm); nY += nYDist;

		oGraphics.p_color(this.m_oPrimaryTextColor.r, this.m_oPrimaryTextColor.g, this.m_oPrimaryTextColor.b, 255);
		let nTextYx = nNumberPosition;
		let nOffsetTextX;
		// если при прилегании к правому краю левый край текста упирается в оффсет, то линии текста должны двигаться вправо(это относится ко всем типам прилегания)
		if ((nTextYx - nMaxTextWidth) < nLeftOffset2)
		{
			nTextYx = nLeftOffset2 + nMaxTextWidth;
			nIndentSize += (nTextYx - nNumberPosition);
			nIndentSize = nIndentSize >> 0;

			nOffsetTextX = this.getFirstLineIndent(oCurrentLvl, nTextYx * AscCommon.g_dKoef_pix_to_mm, nIndentSize * AscCommon.g_dKoef_pix_to_mm, AscFormat.isRealNumber(nTabSize) ? (nTabSize + (nTextYx - nNumberPosition)) * AscCommon.g_dKoef_pix_to_mm : null);

			const nCurrentAlign = oCurrentLvl.Jc;
			oCurrentLvl.Jc = AscCommon.align_Left;
			// считаем позицию отдельно, чтобы нумерация по горизонтали начиналась с одного и того же места
			if (nCurrentAlign === AscCommon.align_Right)
			{
				nTextYx -= nMaxTextWidth;
			}
			else if (nCurrentAlign === AscCommon.align_Center)
			{
				nTextYx -= nMaxTextWidth >> 1;
			}
		}
		else
		{
			nOffsetTextX = this.getFirstLineIndent(oCurrentLvl, nTextYx * AscCommon.g_dKoef_pix_to_mm, nIndentSize * AscCommon.g_dKoef_pix_to_mm, AscFormat.isRealNumber(nTabSize) ? nTabSize * AscCommon.g_dKoef_pix_to_mm : null);
		}

		for (let i = 0; i < 3; i += 1)
		{
			oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nOffsetTextX, nRightOffset * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm); nY += nYDist;
			oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nIndentSize * AscCommon.g_dKoef_pix_to_mm, nRightOffset * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm); nY += nYDist;
		}

		oGraphics.p_color(this.m_oSecondaryLineTextColor.r, this.m_oSecondaryLineTextColor.g, this.m_oSecondaryLineTextColor.b, 255);

		oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nLeftOffset2 * AscCommon.g_dKoef_pix_to_mm, nRightOffset2 * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm); nY += nYDist;
		oGraphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, nY * AscCommon.g_dKoef_pix_to_mm, nLeftOffset2 * AscCommon.g_dKoef_pix_to_mm, nRightOffset2 * AscCommon.g_dKoef_pix_to_mm, nLineWidth * AscCommon.g_dKoef_pix_to_mm);

		for (let i = 0; i < arrTextYy.length; i += 1)
		{
			const nNumberIndex = this.getNumberingValue(i + 1, nCurrentLvl, oCurrentLvl);
			const drawingContent = oCurrentLvl.GetDrawingContent(this.m_arrNumberingLvl, nCurrentLvl, nNumberIndex, this.m_oLang);
			const nTextYy = arrTextYy[i];
			if (typeof drawingContent === "string")
			{
				this.drawTextWithLvlInformation(drawingContent, oCurrentLvl, nTextYx, nTextYy, nLineDistance, oGraphics);
			}
			else
			{
				if (drawingContent.image)
				{
					this.drawImageBulletsWithLine(drawingContent, nTextYx, nTextYy, nLineDistance, oGraphics);
				}
			}
		}
		this.cleanParagraphField(oGraphics, nRightOffset2 * AscCommon.g_dKoef_pix_to_mm, 0, nWidth_px * AscCommon.g_dKoef_pix_to_mm, nHeight_px * AscCommon.g_dKoef_pix_to_mm);
	};
	CBulletPreviewDrawerAdvancedOptions.prototype.draw = function ()
	{
		AscFormat.ExecuteNoHistory(function ()
		{
			if (this.m_bIsMultiLvl)
			{
				this.drawMultiLvlAdvancedOptions();
				this.addControlMultiLvl();
			}
			else
			{
				this.drawSingleLvlAdvancedOptions();
			}
		}, this, []);
	};

	window["AscCommon"] = window["AscCommon"] || {};

	window["AscCommon"].CBulletPreviewDrawer = window["AscCommon"]["CBulletPreviewDrawer"] = CBulletPreviewDrawer;
	window["AscCommon"].CBulletPreviewDrawerChangeList = window["AscCommon"]["CBulletPreviewDrawerChangeList"] = CBulletPreviewDrawerChangeList;
	window["AscCommon"].CBulletPreviewDrawerAdvancedOptions = window["AscCommon"]["CBulletPreviewDrawerAdvancedOptions"] = CBulletPreviewDrawerAdvancedOptions;
})(window);
