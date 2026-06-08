/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
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
	 * @param isFormatContent
	 * @constructor
	 * @extends {AscWord.CDocumentContent}
	 */
	function CTextBoxContent(parent, pdfDocument, isFormatContent) {
		AscWord.CDocumentContent.call(this, null, pdfDocument ? pdfDocument.GetDrawingDocument() : undefined, 0, 0, 0, 0, false, false, false);
		
		this.Content[0].LogicDocument = pdfDocument;
		
		this.ParentPDF 			= parent;
		this.isFormatContent	= !!isFormatContent;
		
		this.transform = new AscCommon.CMatrix();
		this.SetUseXLimit(false);
		this.MoveCursorToStartPos();
	}
	
	CTextBoxContent.prototype = Object.create(AscWord.CDocumentContent.prototype);
	CTextBoxContent.prototype.constructor = CTextBoxContent;
	
	Object.defineProperties(CTextBoxContent.prototype, {
		StartPage: {
			get: function () {
				return this.GetAbsolutePage();
			},
			set: function() {
				return;
			}
		}
	});

	CTextBoxContent.prototype.GetLogicDocument = function() {
		return this.ParentPDF && this.ParentPDF.GetDocument();
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
		
		if (!this.ParentPDF.IsRTL()) {
			switch (align) {
				case align_Left: return AscPDF.ALIGN_TYPE.left;
				case align_Center: return AscPDF.ALIGN_TYPE.center;
				case align_Right: return AscPDF.ALIGN_TYPE.right;
			}
		}
		else {
			switch (align) {
				case align_Left: return AscPDF.ALIGN_TYPE.right;
				case align_Center: return AscPDF.ALIGN_TYPE.center;
				case align_Right: return AscPDF.ALIGN_TYPE.left;
			}
		}
		
		return AscPDF.ALIGN_TYPE.left;
	};
	CTextBoxContent.prototype.IsUseInDocument = function() {
		// TODO: Temporary, later will need to request from parent class
		return true;
	};
	CTextBoxContent.prototype.OnContentReDraw = function() {
		// TODO: Implement
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
				let isOnOpen    = Asc.editor.getDocumentRenderer().IsOpenFormsInProgress;
				let nCharLimit	= this.ParentPDF.GetCharLimit();
				
				if (false == isOnOpen && bIgnoreCount !== true && Asc.editor.isDocumentLoadComplete) {
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
		return AscWord.CDocumentContent.prototype.GetText.call(this, {ParaSeparator: ""});
	};
	CTextBoxContent.prototype.OnContentChange = function() {
		if (this.ParentPDF && this.ParentPDF.OnContentChange && this.isFormatContent == false)
			this.ParentPDF.OnContentChange();
	};
	CTextBoxContent.prototype.GetAbsolutePage = function() {
		return this.ParentPDF.GetPage();
	};
	CTextBoxContent.prototype.Get_ParentTextTransform = function() {
		let parentTransform = this.ParentPDF ? this.ParentPDF.GetTextTransform() : null;
		if (this.transform && parentTransform)
		{
			let transform = new AscCommon.CMatrix();
			global_MatrixTransformer.MultiplyAppend(transform, this.transform);
			global_MatrixTransformer.MultiplyAppend(transform, parentTransform);
			return transform;
		}
		return this.transform || parentTransform;
	};
	CTextBoxContent.prototype.SetTransform = function(transform) {
		if (!transform || transform.IsIdentity())
			this.transform = null;
		else
			this.transform = transform;
	};
	CTextBoxContent.prototype.GetCalculatedTextPr = function(skipFontCalculator) {
		if (this.Content.length == 0) {
			return null;
		}

		return AscWord.CDocumentContent.prototype.GetCalculatedTextPr.call(this, skipFontCalculator);
	};
	
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
			case AscPDF.ALIGN_TYPE.justify:
				nInternalType = AscCommon.align_Justify;
				break;
		}

		return nInternalType;
	}

	function getPdfTypeAlignByInternal(nInternalType, isRTL) {
		let nPdfType = AscPDF.ALIGN_TYPE.left;

		if (!isRTL) {
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
		}
		else {
			switch (nInternalType) {
				case AscCommon.align_Left:
					nPdfType = AscPDF.ALIGN_TYPE.right;
					break;
				case AscCommon.align_Center:
					nPdfType = AscPDF.ALIGN_TYPE.center;
					break;
				case AscCommon.align_Right:
					nPdfType = AscPDF.ALIGN_TYPE.left;
					break;
			}
		}

		return nPdfType;
	}
	//--------------------------------------------------------export----------------------------------------------------
	window['AscPDF'] = window['AscPDF'] || {};

	window['AscPDF'].getInternalAlignByPdfType	= getInternalAlignByPdfType;
	window['AscPDF'].getPdfTypeAlignByInternal	= getPdfTypeAlignByInternal;
	window['AscPDF'].CTextBoxContent			= CTextBoxContent;
	
})(window);
