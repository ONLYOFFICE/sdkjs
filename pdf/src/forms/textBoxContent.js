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

(function(window)
{
	let STYLES = new AscWord.CStyles(false);
	STYLES.Default.TextPr.Merge({
		RFonts : {
			Ascii : {Name : AscPDF.DEFAULT_FIELD_FONT, Index : -1},
			EastAsia : {Name : AscPDF.DEFAULT_FIELD_FONT, Index : -1},
			HAnsi : {Name : AscPDF.DEFAULT_FIELD_FONT, Index : -1},
			CS : {Name : AscPDF.DEFAULT_FIELD_FONT, Index : -1}
		}
	});
	STYLES.Default.ParaPr.Merge({
		KeepLines : false,
		KeepNext : false,
		WidowControl : false,
		PageBreakBefore : false
	});
	
	/**
	 * Class for working with rich text
	 * @param parent - parent class in PDF structure
	 * @param {AscPDF.CPDFDoc} pdfDocument - reference to the main class
	 * @constructor
	 * @extends {AscWord.CDocumentContent}
	 */
	function CTextBoxContent(parent, pdfDocument, isFormatContent) {
		AscWord.CDocumentContent.call(this, null, pdfDocument ? pdfDocument.GetDrawingDocument() : undefined, 0, 0, 0, 0, false, false, false);
		
		this.Content[0].LogicDocument = pdfDocument;
		
		this.ParentPDF 			= parent;
		this.PdfDoc    			= pdfDocument;
		this.isFormatContent	= !!isFormatContent;
		
		this.transform = new AscCommon.CMatrix();
		this.SetUseXLimit(false);
		this.MoveCursorToStartPos();
	}
	
	CTextBoxContent.prototype = Object.create(AscWord.CDocumentContent.prototype);
	CTextBoxContent.prototype.constructor = CTextBoxContent;
	
	CTextBoxContent.prototype.GetLogicDocument = function() {
		return this.PdfDoc;
	};
	CTextBoxContent.prototype.SetAlign = function(alignType) {
		let _alignType = getInternalAlignByPdfType(alignType);
		
		this.SetApplyToAll(true);
		this.SetParagraphAlign(_alignType);
		this.SetApplyToAll(false);
		this.Content.forEach(function(para) {
			para.RecalcCompiledPr(true);
		});
	};
	CTextBoxContent.prototype.GetAlign = function() {
		let align = this.GetElement(0).GetParagraphAlign();
		
		switch (align) {
			case align_Left: return AscPDF.ALIGN_TYPE.left;
			case align_Center: return AscPDF.ALIGN_TYPE.center;
			case align_Right: return AscPDF.ALIGN_TYPE.right;
		}
		
		return AscPDF.ALIGN_TYPE.left;
	};
	CTextBoxContent.prototype.IsUseInDocument = function() {
		// TODO: Временно, потом надо будет запрашивать у родительского класса
		return true;
	};
	CTextBoxContent.prototype.OnContentReDraw = function() {
		// TODO: Реализовать
	};
	CTextBoxContent.prototype.GetStyles = function() {
		return STYLES;
	};
	CTextBoxContent.prototype.SetFont = function(fontName) {
		this.SetApplyToAll(true);
		let oParaTextPr = new AscWord.ParaTextPr();
		oParaTextPr.Value.RFonts.SetAll(fontName, -1);
		this.AddToParagraph(oParaTextPr);
		this.SetApplyToAll(false);
	};
	CTextBoxContent.prototype.SetFontSize = function(fontSize) {
		this.SetApplyToAll(true);
		this.AddToParagraph(new AscWord.ParaTextPr({FontSize : fontSize}));
		this.SetApplyToAll(false);
	};
	CTextBoxContent.prototype.SetBold = function(bBold) {
		this.SetApplyToAll(true);
		this.AddToParagraph(new AscWord.ParaTextPr({Bold : bBold}));
		this.SetApplyToAll(false);
	};
	CTextBoxContent.prototype.SetItalic = function(bItalic) {
		this.SetApplyToAll(true);
		this.AddToParagraph(new AscWord.ParaTextPr({Italic : bItalic}));
		this.SetApplyToAll(false);
	};
	CTextBoxContent.prototype.replaceAllText = function(value, bIgnoreCount) {
		let codePoints = typeof(value) === "string" ? value.codePointsArray() : value;
		
		let paragraph = this.GetElement(0);
		if (!paragraph || !paragraph.IsParagraph())
			return;
		
		let run = paragraph.GetElement(0);
		if (!run || !(run instanceof AscWord.CRun))
			return;
		
		paragraph.RemoveFromContent(1, paragraph.GetElementsCount() - 1);
		run.ClearContent();
		
		if (codePoints) {
			if (this.ParentPDF && this.ParentPDF.GetCharLimit && 0 !== this.ParentPDF.GetCharLimit()) {
				let oDoc        = this.ParentPDF.GetDocument();
				let isOnOpen    = oDoc.Viewer.IsOpenFormsInProgress;
				let nCharLimit	= this.ParentPDF.GetCharLimit();
				
				if (false == isOnOpen && bIgnoreCount !== true) {
					let nCharsCount = AscWord.GraphemesCounter.GetCount(codePoints, this.GetCalculatedTextPr());
					
					if (nCharsCount > nCharLimit)
						codePoints.length = nCharLimit;
				}
			}

			for (let index = 0, inRunIndex = 0, count = codePoints.length; index < count; ++index) {
				let runElement = AscPDF.codePointToRunElement(codePoints[index]);
				if (runElement)
					run.AddToContent(inRunIndex++, runElement, true);
			}
			this.MoveCursorToEndPos();
		}
	};
	CTextBoxContent.prototype.EnterText = function(value) {
		let oParentPDF = this.ParentPDF;
		let isAllowLineBreak = oParentPDF.IsForm() && oParentPDF.GetType() == AscPDF.FIELD_TYPES.text && oParentPDF.IsMultiline();

		if (undefined === value
			|| null === value
			|| (Array.isArray(value) && !value.length))
			return false;
		
		let codePoints = typeof(value) === "string" ? value.codePointsArray() : value;
		
		if (Array.isArray(codePoints)) {
			for (let index = 0, count = codePoints.length; index < count; ++index) {
				let codePoint = codePoints[index];
				addToParagraph.call(this, codePoint);
			}
		}
		else {
			addToParagraph.call(this, codePoints);
		}
		
		function addToParagraph(codePoint) {
			if ((10 === codePoint || 13 === codePoint)) {
				if (isAllowLineBreak) {
					this.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line));
				}
				else {
					this.AddToParagraph(new AscWord.CRunSpace(32));
				}
			}
			else if (9 === codePoint) {
				this.AddToParagraph(new AscWord.CRunSpace(32));
			}
			else if (AscCommon.IsSpace(codePoint)) {
				this.AddToParagraph(new AscWord.CRunSpace(codePoint));
			}
			else {
				this.AddToParagraph(new AscWord.CRunText(codePoint));
			}
		}

		return true;
	};
	CTextBoxContent.prototype.getAllText = function() {
		let paragraph = this.GetElement(0);
		if (!paragraph || !paragraph.IsParagraph())
			return "";
		
		paragraph.SetApplyToAll(true);
		let text = paragraph.GetSelectedText(true, {ParaSeparator: ""});
		paragraph.SetApplyToAll(false);
		return text;
	};
	CTextBoxContent.prototype.OnContentChange = function() {
		if (this.ParentPDF && this.ParentPDF.OnContentChange && this.isFormatContent == false)
			this.ParentPDF.OnContentChange();
	};
	CTextBoxContent.prototype.Get_ParentTextTransform = function() {
		return this.transform;
	};
	CTextBoxContent.prototype.Get_AbsolutePage = function() {
		return this.ParentPDF.GetPage();
	};
	CTextBoxContent.prototype.Get_ParentTextTransform = function() {};
	
	function getInternalAlignByPdfType(nPdfType) {
		let nInternalType = AscCommon.align_Left;
		switch (nPdfType) {
			case AscPDF.ALIGN_TYPE.left:
				nInternalType = AscCommon.align_Left;
				break;
			case AscPDF.ALIGN_TYPE.center:
				nInternalType = AscCommon.align_Center;
				break;
			case AscPDF.ALIGN_TYPE.right:
				nInternalType = AscCommon.align_Right;
				break;
		}

		return nInternalType;
	}

	function getPdfTypeAlignByInternal(nInternalType) {
		let nPdfType = AscPDF.ALIGN_TYPE.left;
		switch (nInternalType) {
			case AscCommon.align_Left:
				nPdfType = AscPDF.ALIGN_TYPE.left;
				break;
			case AscCommon.align_Center:
				nPdfType = AscPDF.ALIGN_TYPE.center;
				break;
			case AscCommon.align_Right:
				nPdfType = AscPDF.ALIGN_TYPE.right;
				break;
		}

		return nPdfType;
	}
	//--------------------------------------------------------export----------------------------------------------------
	window['AscPDF'] = window['AscPDF'] || {};

	window['AscPDF'].getInternalAlignByPdfType	= getInternalAlignByPdfType;
	window['AscPDF'].getPdfTypeAlignByInternal	= getPdfTypeAlignByInternal;
	window['AscPDF'].CTextBoxContent			= CTextBoxContent;
	
})(window);
