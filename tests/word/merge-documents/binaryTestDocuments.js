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

const mockEditor = AscTest.Editor;
mockEditor.pre_Paste = function (first, second, callback)
{
	callback();
};
AscCommon.sendImgUrls = function (oApi, arrImages, fCallback) {fCallback()};
AscCommon.ResetNewUrls = function () {};

let oMainComments;
mockEditor.sync_AddComment = function ()
{
	
};
mockEditor.sync_ChangeCommentData = function ()
{

};
mockEditor.sync_RemoveComment = function ()
{
	
};
AscCommonWord.CDocument.prototype.Document_UpdateRulersState = function () {};

let oCurrentTestDocument;
AscCommonWord.CDocument.prototype.getTestObject = function ()
{
	oCurrentTestDocument = this;
	const oContentObject = {type: 'document', content: []};
	this.Content.forEach(function (oItem)
	{
		if (oItem.getTestObject)
		{
			oItem.getTestObject(oContentObject.content);
		}
		else
		{
			oContentObject.content.push(oItem.constructor.name);
		}
	});
	if (this.SectPr)
	{
		const arrHdrFtr = this.SectPr.GetAllHdrFtrs();
		for (let i = 0; i < arrHdrFtr.length; i += 1)
		{
			arrHdrFtr[i].getTestObject(oContentObject.content);
		}

	}
	if (this.Footnotes)
	{
		this.Footnotes.getTestObject(oContentObject.content);
	}
	return oContentObject;
};
AscCommon.ParaComment.prototype.getTestObject = function (arrParentContent)
{
	let oComment = oMainComments.Get_ById(this.GetCommentId());
	if (oComment)
	{
		const oContentObject = {type: 'comment', text: oComment.Data.Get_Text(), quoteText: oComment.Data.Get_QuoteText(), arrAnswers: oComment.Data.m_aReplies.map((e) => e.Get_Text())};
		arrParentContent.push(oContentObject);
	}
};
AscCommonWord.CTextPr.prototype.getTestObject = function () {
	const oRes = {};
	if (this.Bold !== undefined) {
		oRes.Bold = this.Bold;
	}
	if (this.Italic !== undefined) {
		oRes.Italic = this.Italic;
	}
	if (this.Strikeout !== undefined) {
		oRes.Strikeout = this.Strikeout;
	}
	if (this.Underline !== undefined) {
		oRes.Underline = this.Underline;
	}
	if (this.FontSize !== undefined) {
		oRes.FontSize = this.FontSize;
	}
	if (this.VertAlign !== undefined) {
		oRes.VertAlign = this.VertAlign;
	}
	if (this.RStyle !== undefined) {
		oRes.RStyle = this.RStyle;
	}
	if (this.Spacing !== undefined) {
		oRes.Spacing = this.Spacing;
	}
	if (this.DStrikeout !== undefined) {
		oRes.DStrikeout = this.DStrikeout;
	}
	if (this.Caps !== undefined) {
		oRes.Caps = this.Caps;
	}
	if (this.SmallCaps !== undefined) {
		oRes.SmallCaps = this.SmallCaps;
	}
	if (this.Position !== undefined) {
		oRes.Position = this.Position;
	}
	if (this.BoldCS !== undefined) {
		oRes.BoldCS = this.BoldCS;
	}
	if (this.ItalicCS !== undefined) {
		oRes.ItalicCS = this.ItalicCS;
	}
	if (this.FontSizeCS !== undefined) {
		oRes.FontSizeCS = this.FontSizeCS;
	}
	if (this.CS !== undefined) {
		oRes.CS = this.CS;
	}
	if (this.RTL !== undefined) {
		oRes.RTL = this.RTL;
	}
	if (this.FontScale !== undefined) {
		oRes.FontScale = this.FontScale;
	}
	if (this.FontSizeOrig !== undefined) {
		oRes.FontSizeOrig = this.FontSizeOrig;
	}
	if (this.FontSizeCSOrig !== undefined) {
		oRes.FontSizeCSOrig = this.FontSizeCSOrig;
	}
	return oRes;
};
AscCommonWord.CTextPr.prototype.isEqualTestObject = function (oTestObject) {
	if (this.Bold !== oTestObject.Bold) {
		return false;
	}
	if (this.Italic !== oTestObject.Italic) {
		return false;
	}
	if (this.Strikeout !== oTestObject.Strikeout) {
		return false;
	}
	if (this.Underline !== oTestObject.Underline) {
		return false;
	}
	if (this.FontSize !== oTestObject.FontSize) {
		return false;
	}
	if (this.VertAlign !== oTestObject.VertAlign) {
		return false;
	}
	if (this.RStyle !== oTestObject.RStyle) {
		return false;
	}
	if (this.Spacing !== oTestObject.Spacing) {
		return false;
	}
	if (this.DStrikeout !== oTestObject.DStrikeout) {
		return false;
	}
	if (this.Caps !== oTestObject.Caps) {
		return false;
	}
	if (this.SmallCaps !== oTestObject.SmallCaps) {
		return false;
	}
	if (this.Position !== oTestObject.Position) {
		return false;
	}
	if (this.BoldCS !== oTestObject.BoldCS) {
		return false;
	}
	if (this.ItalicCS !== oTestObject.ItalicCS) {
		return false;
	}
	if (this.FontSizeCS !== oTestObject.FontSizeCS) {
		return false;
	}
	if (this.CS !== oTestObject.CS) {
		return false;
	}
	if (this.RTL !== oTestObject.RTL) {
		return false;
	}
	if (this.FontScale !== oTestObject.FontScale) {
		return false;
	}
	if (this.FontSizeOrig !== oTestObject.FontSizeOrig) {
		return false;
	}
	if (this.FontSizeCSOrig !== oTestObject.FontSizeCSOrig) {
		return false;
	}
	return true;
};

AscCommonWord.CHeaderFooter.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'headerfooter', content: []};
	arrParentContent.push(oContentObject);
	this.Content.getTestObject(oContentObject.content);
};

AscCommonWord.CTable.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'table', rows: []};
	arrParentContent.push(oContentObject);
	for (let i = 0; i < this.Content.length; i += 1)
	{
		const row = this.Content[i];
		row.getTestObject(oContentObject.rows);
	}
};

AscCommonWord.CTableRow.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'row', content: []};
	arrParentContent.push(oContentObject);
	for (let i = 0; i < this.Content.length; i += 1)
	{
		const cell = this.Content[i];
		cell.getTestObject(oContentObject.content);
	}
};

AscCommonWord.CTableCell.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'cell', content: []};
	arrParentContent.push(oContentObject);
	const oContent = this.GetContent();
	oContent.CheckRunContent(function (oRun)
	{
		oRun.getTestObject(oContentObject.content);
	});
};


AscCommonWord.CDocument.prototype.CreateStyles = function()
{
	this.Styles = getDocumentStyles(this);
};
CGlossaryDocument.prototype.CreateStyles = function()
{
	this.Styles = getDocumentGlossaryStyles();
};

ParaMath.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'paramath', content: []};
	arrParentContent.push(oContentObject)
	this.Root.getTestObject(oContentObject.content);
}
CMathContent.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'mathcontent', content: []};
	arrParentContent.push(oContentObject)
	for (var i = 0; i < this.Content.length; ++i)
	{
		if (para_Math_Run === this.Content[i].Type)
			this.Content[i].getTestObject(oContentObject.content);
	}
}
CMathBase.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'mathbase', content: []};
	arrParentContent.push(oContentObject)
	this.Content.forEach(function (oRun)
	{
		oRun.getTestObject(oContentObject.content);
	});
}
CDocumentSectionsInfo.prototype.getTestObject = function (arrParentContent)
{
	const arrHeaders = this.GetAllHdrFtrs();
	for (let index = 0, count = arrHeaders.length; index < count; ++index)
	{
		const oContentObject = {
			type   : 'documentsectioninfo',
			content: []
		}
		arrParentContent.push(oContentObject);
		arrHeaders[index].getTestObject(oContentObject.content);
	}
}
CDocumentContentBase.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'documentcontentbase', content: []};
	arrParentContent.push(oContentObject)
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].getTestObject(oContentObject.content);
	}
}
CDocumentContentElementBase.prototype.getTestObject = function ()
{

};
CEndnotesController.prototype.getTestObject = function (arrParentContent)
{
	for (var sId in this.Endnote)
	{
		const oEndnote = this.Endnote[sId];
		const oContentObject = {type: 'endnote', content: []};
		arrParentContent.push(oContentObject);
		oEndnote.checkTestObject(oContentObject.content)
	}
};
CFootnotesController.prototype.getTestObject = function (arrParentContent)
{
	for (var sId in this.Footnote)
	{
		const oFootnote = this.Footnote[sId];
		const oContentObject = {type: 'footnote', content: []};
		arrParentContent.push(oContentObject);
		for (let i = 0; i < oFootnote.Content.length; i += 1)
		{
			oFootnote.Content[i].getTestObject(oContentObject.content);
		}
	}
};
CParagraphContentBase.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'paragraphcontentbase', content: []};
	arrParentContent.push(oContentObject)
	for (let i = 0; i < this.Content.length; i += 1)
	{
		this.Content[i].getTestObject(oContentObject.content);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.getTestObject = function ()
{

};
CBlockLevelSdt.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'blocklvlsdt', content: []};
	arrParentContent.push(oContentObject)
	this.Content.getTestObject(oContentObject.content);
};
Paragraph.prototype.getTestObject = function (arrParentContent)
{
	const oContentObject = {type: 'paragraph', content: []};
	arrParentContent.push(oContentObject);
	for (let i = 0; i < this.Content.length; i += 1)
	{
		if (this.Content[i].getTestObject)
		{
			this.Content[i].getTestObject(oContentObject.content);
		}
	}
};
AscCommonWord.CParagraphBookmark.prototype.getTestObject = function (arrParentContent)
{
	const oStartBookmark = oCurrentTestDocument.BookmarksManager.GetBookmarkById(this.GetBookmarkId())[0];
	const oContentObject = {
		type : 'bookmark',
		name : oStartBookmark.GetBookmarkName(),
		start: this.IsStart()
	};
	arrParentContent.push(oContentObject)
}
ParaRun.prototype.getTestObject = function (oParentContent)
{
	if (!this.Content.length)
		return;
	
	let oReviewInfo = this.GetReviewInfo();
	const oPrevAdded = oReviewInfo && oReviewInfo.GetPrevAdded();
	let nMainReviewType = this.GetReviewType();
	let sMainUserName = oReviewInfo && oReviewInfo.GetUserName();
	let nMainDateTime = oReviewInfo && oReviewInfo.GetDateTime();

	let nAdditionalReviewType;
	let sAdditionalUserName;
	let nAdditionalDateTime;

	if (oPrevAdded)
	{
		nAdditionalReviewType = reviewtype_Add;
		sAdditionalUserName = oPrevAdded.GetUserName();
		nAdditionalDateTime = oPrevAdded.GetDateTime();
	}
	let oCurrentTextInfo = oParentContent[oParentContent.length - 1];
	const needCreateNewText = (oParentContent.length === 0 ||
		oCurrentTextInfo.mainReviewType !== nMainReviewType || oCurrentTextInfo.mainUserName !== sMainUserName || oCurrentTextInfo.mainDateTime !== nMainDateTime ||
		oCurrentTextInfo.additionalReviewType !== nAdditionalReviewType || oCurrentTextInfo.additionalUserName !== sAdditionalUserName || oCurrentTextInfo.additionalDateTime !== nAdditionalDateTime ||
	!this.Pr.isEqualTestObject(oCurrentTextInfo.textPr));
	if (needCreateNewText || this.IsParaEndRun())
	{
		oCurrentTextInfo = {
			mainReviewType      : nMainReviewType,
			mainDateTime        : nMainDateTime,
			mainUserName        : sMainUserName,
			additionalReviewType: nAdditionalReviewType,
			additionalDateTime  : nAdditionalDateTime,
			additionalUserName  : sAdditionalUserName,
			text                : '',
			textPr              : this.Pr.getTestObject()
		};
		oParentContent.push(oCurrentTextInfo);
	}
	this.Content.forEach(function (el)
	{
		oCurrentTextInfo.text += String.fromCharCode(el.Value)
	});
};

window['AscCommonWord']['CDocumentComparison'].prototype.getReviewInfo = function (sCustomReviewUserName, nCustomReviewDate)
{
	let oReviewInfo = new AscWord.ReviewInfo();
	oReviewInfo.Editor = this.api;
	oReviewInfo.UserId = "";
	oReviewInfo.MoveType = Asc.c_oAscRevisionsMove.NoMove;
	oReviewInfo.PrevType = -1;
	oReviewInfo.PrevInfo = null;
	oReviewInfo.UserName = sCustomReviewUserName || "Valdemar";
	oReviewInfo.DateTime = 3000000;
	if (AscFormat.isRealNumber(nCustomReviewDate))
	{
		oReviewInfo.DateTime = nCustomReviewDate;
	}
	return oReviewInfo;
};

let GLOBAL_DOC_STYLES = null;
let GLOBAL_GLOSSARY_STYLES = null;
function getDocumentStyles(oLogicDocument) {
	if(!GLOBAL_DOC_STYLES) {
		GLOBAL_DOC_STYLES = new CStyles();
	}
	GLOBAL_DOC_STYLES.Set_LogicDocument(oLogicDocument);
	return GLOBAL_DOC_STYLES;
}
function getDocumentGlossaryStyles() {
	if(!GLOBAL_GLOSSARY_STYLES) {
		GLOBAL_GLOSSARY_STYLES = new CStyles();
	}
	return GLOBAL_GLOSSARY_STYLES;
}
function readMainDocument(oMainDocumentInfo)
{
	const oDocument = new AscWord.CDocument(mockEditor.WordControl.m_oDrawingDocument, true, false);
	mockEditor.WordControl.m_oDrawingDocument.m_oLogicDocument = oDocument;
	mockEditor.WordControl.m_oLogicDocument = oDocument;
	oDocument.Styles = getDocumentStyles();
	oDocument.Styles.Set_LogicDocument(oDocument);
	oDocument.GlossaryDocument.Styles = getDocumentGlossaryStyles();
	oDocument.Api = mockEditor;
	oMainComments = oDocument.Comments;
	createTestDocument(oDocument, oMainDocumentInfo);
	return oDocument
}

function readRevisedDocument(oRevisedDocumentInfo)
{
	const oMainDocument = mockEditor.WordControl.m_oLogicDocument;
	const oRevisedDocument = new CDocument(mockEditor.WordControl.m_oDrawingDocument, true, false);
	mockEditor.WordControl.m_oDrawingDocument.m_oLogicDocument = oRevisedDocument;
	mockEditor.WordControl.m_oLogicDocument = oRevisedDocument;
	oRevisedDocument.Styles = getDocumentStyles();
	oRevisedDocument.Styles.Set_LogicDocument(oRevisedDocument);
	oRevisedDocument.GlossaryDocument.Styles = getDocumentGlossaryStyles();

	createTestDocument(oRevisedDocument, oRevisedDocumentInfo);

	mockEditor.WordControl.m_oDrawingDocument.m_oLogicDocument = oMainDocument;
	mockEditor.WordControl.m_oLogicDocument = oMainDocument;
	if (oMainDocument.History)
		oMainDocument.History.Set_LogicDocument(oMainDocument);

	if (oMainDocument.CollaborativeEditing)
		oMainDocument.CollaborativeEditing.m_oLogicDocument = oMainDocument;

	return oRevisedDocument;
}
function getComment(oDocument, oCommentData)
{
	const oData = new AscCommon.CCommentData();
	oData.Set_Text(oCommentData.text);
	oData.Set_QuoteText(oCommentData.quoteText);
	if (oCommentData.arrAnswers)
	{
		for (let i = 0; i < oCommentData.arrAnswers.length; i += 1)
		{
			const oAnswer = new AscCommon.CCommentData();
			oAnswer.Set_Text(oCommentData.arrAnswers[i]);
			oData.Add_Reply(oAnswer);
		}
	}

	return new AscCommon.CComment(oDocument.Comments, oData);
}
function createTestDocument(oDocument, arrParagraphsTextInfo)
{
	const mapParaComments = {};
	for (let i = 0; i < arrParagraphsTextInfo.length; i += 1)
	{
		const oParagraphTextInfo = arrParagraphsTextInfo[i];
		let oParagraph;
		if (i === 0)
		{
			oParagraph = oDocument.Content[0];
		}
		else
		{
			oParagraph = AscTest.CreateParagraph();
		}
		for (let j = 0; j < oParagraphTextInfo.length; j += 1)
		{
			let arrStartBookmarkInfo;
			let arrEndBookmarkInfo;
			if (oParagraphTextInfo[j].bookmark)
			{
				arrStartBookmarkInfo = oParagraphTextInfo[j].bookmark.start;
				arrEndBookmarkInfo = oParagraphTextInfo[j].bookmark.end;
			}
			let arrStartCommentsInfo;
			let arrEndCommentsInfo;
			if (oParagraphTextInfo[j].options.comments)
			{
				arrStartCommentsInfo = oParagraphTextInfo[j].options.comments.start;
				arrEndCommentsInfo = oParagraphTextInfo[j].options.comments.end;
			}
			if (arrStartCommentsInfo)
			{
				for (let k = 0; k < arrStartCommentsInfo.length; k += 1)
				{
					const oStartCommentInfo = arrStartCommentsInfo[k];
					let oComment
					const oParaComment = new AscCommon.ParaComment(!!oStartCommentInfo.start);
					if (!oParaComment.IsCommentStart())
					{
						oComment = getComment(oDocument, oStartCommentInfo.data);
						oParaComment.SetCommentId(oComment.GetId());
						const oStartParaComment = mapParaComments[oStartCommentInfo.id];
						oStartParaComment.SetCommentId(oComment.GetId());
						oComment.SetRangeMark(oParaComment);
						oComment.SetRangeMark(oStartParaComment);
					}
					else
					{
						mapParaComments[oStartCommentInfo.id] = oParaComment;
					}
					oParagraph.AddToContentToEnd(oParaComment);
					if (oComment)
					{
						oDocument.Comments.Add(oComment);
					}
				}
			}
			if (arrStartBookmarkInfo)
			{
				for (let k = 0; k < arrStartBookmarkInfo.length; k += 1)
				{
					const oStartBookmarkInfo = arrStartBookmarkInfo[k];
					const oBookmark = new AscCommonWord.CParagraphBookmark(!!oStartBookmarkInfo.name, oStartBookmarkInfo.id, oStartBookmarkInfo.name);
					oParagraph.AddToContentToEnd(oBookmark);
				}
			}

			const oParaRun = new AscWord.ParaRun();
			if (oParagraphTextInfo[j].text)
			{
				oParaRun.AddText(oParagraphTextInfo[j].text);
				oParaRun.SetReviewTypeWithInfo(oParagraphTextInfo[j].reviewType, oParagraphTextInfo[j].reviewInfo);
				oParagraph.AddToContentToEnd(oParaRun);
			}
			if (oParagraphTextInfo[j].options.textPr) {
				const oTextPr = new AscCommonWord.CTextPr();
				oTextPr.Set_FromObject(oParagraphTextInfo[j].options.textPr);
				oParaRun.SetPr(oTextPr);
			}
			if (arrEndBookmarkInfo)
			{
				for (let k = 0; k < arrEndBookmarkInfo.length; k += 1)
				{
					const oEndBookmarkInfo = arrEndBookmarkInfo[k];
					const oBookmark = new AscCommonWord.CParagraphBookmark(!!oEndBookmarkInfo.name, oEndBookmarkInfo.id, oEndBookmarkInfo.name);
					oParagraph.AddToContentToEnd(oBookmark);
				}
			}
			if (arrEndCommentsInfo)
			{
				for (let k = 0; k < arrEndCommentsInfo.length; k += 1)
				{
					const oEndCommentInfo = arrEndCommentsInfo[k];
					let oComment;
					const oParaComment = new AscCommon.ParaComment(!!oEndCommentInfo.start);
					if (!oParaComment.IsCommentStart())
					{
						oComment = getComment(oDocument, oEndCommentInfo.data);
						oParaComment.SetCommentId(oComment.GetId());
						const oStartParaComment = mapParaComments[oEndCommentInfo.id];
						oStartParaComment.SetCommentId(oComment.GetId());
						oComment.SetRangeMark(oParaComment);
						oComment.SetRangeMark(oStartParaComment);

					}
					else
					{
						mapParaComments[oEndCommentInfo.id] = oParaComment;
					}
					oParagraph.AddToContentToEnd(oParaComment);
					if (oComment)
					{
						oDocument.Comments.Add(oComment);
					}
				}
			}
			if (!oParagraphTextInfo[j].text && oParagraphTextInfo[j].reviewType && oParagraphTextInfo[j].reviewInfo)
			{
				oParagraph.GetParaEndRun().SetReviewTypeWithInfo(oParagraphTextInfo[j].reviewType, oParagraphTextInfo[j].reviewInfo, false);
			}
		}
		oDocument.Comments.CheckMarks();
		if (i !== 0)
		{
			oDocument.AddToContent(oDocument.Content.length, oParagraph);
		}
	}
	return oDocument;
}

function createParagraphInfo(sText, oMainReviewInfoOptions, oAdditionalReviewInfoOptions, oBookmarkInfo, oOptions)
{
	const oResult = {
		text      : sText,
		reviewType: reviewtype_Common,
		bookmark  : oBookmarkInfo,
		options   : oOptions || {},
	};
	let oMainReviewInfo;
	if (oMainReviewInfoOptions)
	{
		oMainReviewInfo = createReviewInfoFromOptions(oMainReviewInfoOptions);
		oResult.reviewType = oMainReviewInfoOptions.reviewType;
		if (oAdditionalReviewInfoOptions)
		{
			const oAdditionalReviewInfo = createReviewInfoFromOptions(oAdditionalReviewInfoOptions);
			oAdditionalReviewInfo.SavePrev(oAdditionalReviewInfoOptions.reviewType);
			oMainReviewInfo.PrevType = oAdditionalReviewInfo.PrevType;
			oMainReviewInfo.PrevInfo = oAdditionalReviewInfo.PrevInfo;
		}
	}
	else
	{
		oMainReviewInfo = createReviewInfoFromOptions();
	}
	oResult.reviewInfo = oMainReviewInfo;
	return oResult;
}

function createShapeInfo()
{

}

function createReviewInfoFromOptions(oOptions)
{
	oOptions = oOptions || {};
	const oReviewInfo = new AscWord.ReviewInfo();

	oReviewInfo.Editor = mockEditor;
	oReviewInfo.UserId = "";
	oReviewInfo.MoveType = Asc.c_oAscRevisionsMove.NoMove;
	oReviewInfo.PrevType = -1;
	oReviewInfo.PrevInfo = null;
	oReviewInfo.UserName = oOptions.userName || oReviewInfo.UserName;
	oReviewInfo.DateTime = oOptions.dateTime || oReviewInfo.DateTime;

	return oReviewInfo;
}

function CCreatingReviewInfo(sUserName, nReviewType, nDateTime)
{
	this.userName = sUserName;
	this.reviewType = nReviewType;
	this.dateTime = nDateTime;
}

function createFindingReviewInfo(nReviewType)
{
	return new CCreatingReviewInfo('Valdemar', nReviewType, 3000000);
}
//
// asc_docs_api.prototype.getTest = function()
// {
// 	function getPr(oPr) {
// 		const oRes = {};
// 		if (oPr.Bold !== undefined) {
// 			oRes.Bold = oPr.Bold;
// 		}
// 		if (oPr.Italic !== undefined) {
// 			oRes.Italic = oPr.Italic;
// 		}
// 		if (oPr.Strikeout !== undefined) {
// 			oRes.Strikeout = oPr.Strikeout;
// 		}
// 		if (oPr.Underline !== undefined) {
// 			oRes.Underline = oPr.Underline;
// 		}
// 		if (oPr.FontSize !== undefined) {
// 			oRes.FontSize = oPr.FontSize;
// 		}
// 		if (oPr.VertAlign !== undefined) {
// 			oRes.VertAlign = oPr.VertAlign;
// 		}
// 		if (oPr.RStyle !== undefined) {
// 			oRes.RStyle = oPr.RStyle;
// 		}
// 		if (oPr.Spacing !== undefined) {
// 			oRes.Spacing = oPr.Spacing;
// 		}
// 		if (oPr.DStrikeout !== undefined) {
// 			oRes.DStrikeout = oPr.DStrikeout;
// 		}
// 		if (oPr.Caps !== undefined) {
// 			oRes.Caps = oPr.Caps;
// 		}
// 		if (oPr.SmallCaps !== undefined) {
// 			oRes.SmallCaps = oPr.SmallCaps;
// 		}
// 		if (oPr.Position !== undefined) {
// 			oRes.Position = oPr.Position;
// 		}
// 		if (oPr.BoldCS !== undefined) {
// 			oRes.BoldCS = oPr.BoldCS;
// 		}
// 		if (oPr.ItalicCS !== undefined) {
// 			oRes.ItalicCS = oPr.ItalicCS;
// 		}
// 		if (oPr.FontSizeCS !== undefined) {
// 			oRes.FontSizeCS = oPr.FontSizeCS;
// 		}
// 		if (oPr.CS !== undefined) {
// 			oRes.CS = oPr.CS;
// 		}
// 		if (oPr.RTL !== undefined) {
// 			oRes.RTL = oPr.RTL;
// 		}
// 		if (oPr.FontScale !== undefined) {
// 			oRes.FontScale = oPr.FontScale;
// 		}
// 		if (oPr.FontSizeOrig !== undefined) {
// 			oRes.FontSizeOrig = oPr.FontSizeOrig;
// 		}
// 		if (oPr.FontSizeCSOrig !== undefined) {
// 			oRes.FontSizeCSOrig = oPr.FontSizeCSOrig;
// 		}
// 		return oRes;
// 	}
// 	function getReview(oRun) {
// 		const bPrevAdded = oRun.ReviewInfo && oRun.ReviewInfo.GetPrevAdded();
// 		const nMainReviewInfo = oRun.GetReviewType();
//
// 		const oRes = {};
// 		if (nMainReviewInfo !== reviewtype_Common) {
// 			oRes.mainReviewType = nMainReviewInfo;
// 		}
//
// 		if (bPrevAdded) {
// 			oRes.additionalReviewType = reviewtype_Add;
// 		}
// 		return oRes;
// 	}
//
// 	function compareReviewDiff(rev1, rev2) {
// 		if (rev1.mainReviewType !== rev2.mainReviewType) {
// 			return false;
// 		}
// 		if (rev1.additionalReviewType !== rev2.additionalReviewType) {
// 			return false;
// 		}
// 		return true;
// 	}
//
// 	function compareTextPrDiff(oPr1, oPr2) {
// 		if (oPr1.Bold !== oPr2.Bold) {
// 			return false;
// 		}
// 		if (oPr1.Italic !== oPr2.Italic) {
// 			return false;
// 		}
// 		if (oPr1.Strikeout !== oPr2.Strikeout) {
// 			return false;
// 		}
// 		if (oPr1.Underline !== oPr2.Underline) {
// 			return false;
// 		}
// 		if (oPr1.FontSize !== oPr2.FontSize) {
// 			return false;
// 		}
// 		if (oPr1.VertAlign !== oPr2.VertAlign) {
// 			return false;
// 		}
// 		if (oPr1.RStyle !== oPr2.RStyle) {
// 			return false;
// 		}
// 		if (oPr1.Spacing !== oPr2.Spacing) {
// 			return false;
// 		}
// 		if (oPr1.DStrikeout !== oPr2.DStrikeout) {
// 			return false;
// 		}
// 		if (oPr1.Caps !== oPr2.Caps) {
// 			return false;
// 		}
// 		if (oPr1.SmallCaps !== oPr2.SmallCaps) {
// 			return false;
// 		}
// 		if (oPr1.Position !== oPr2.Position) {
// 			return false;
// 		}
// 		if (oPr1.BoldCS !== oPr2.BoldCS) {
// 			return false;
// 		}
// 		if (oPr1.ItalicCS !== oPr2.ItalicCS) {
// 			return false;
// 		}
// 		if (oPr1.FontSizeCS !== oPr2.FontSizeCS) {
// 			return false;
// 		}
// 		if (oPr1.CS !== oPr2.CS) {
// 			return false;
// 		}
// 		if (oPr1.RTL !== oPr2.RTL) {
// 			return false;
// 		}
// 		if (oPr1.FontScale !== oPr2.FontScale) {
// 			return false;
// 		}
// 		if (oPr1.FontSizeOrig !== oPr2.FontSizeOrig) {
// 			return false;
// 		}
// 		if (oPr1.FontSizeCSOrig !== oPr2.FontSizeCSOrig) {
// 			return false;
// 		}
// 		return true;
// 	}
//
// 	const oLogicDocument = this.private_GetLogicDocument();
// 	const arrTests = [];
// 	for (let i = 0; i < oLogicDocument.Content.length; i += 1) {
// 		const oParagraph = oLogicDocument.Content[i];
// 		let arrBookmarks = [];
// 		let arrComments = [];
// 		let test;
// 		let paragraphTest = [];
// 		arrTests.push(paragraphTest);
// 		for (let j = 0; j < oParagraph.Content.length; j += 1) {
// 			const oRunElement = oParagraph.Content[j];
// 			if (oRunElement instanceof AscCommon.ParaComment) {
// 				const isStart = oRunElement.IsCommentStart();
// 				const comment = oLogicDocument.Comments.GetById(oRunElement.CommentId);
// 				const data = !isStart ? `, data:{text: "${comment.Data.m_sText}", quoteText: "${comment.Data.m_sQuoteText}", arrAnswers: ${comment.Data.m_aReplies.length ? `[${comment.Data.m_aReplies.map((e) => '"' + e.Get_Text() + '"')}]` : "null"}}`: "";
// 				arrComments.push(`{start: ${isStart}, id: ${comment.Id}${data}}`);
// 			} else if (oRunElement instanceof AscCommonWord.CParagraphBookmark) {
// 				arrBookmarks.push(`{id: ${oRunElement.GetBookmarkId()}${oRunElement.IsStart() ? ", name: " + '"' + oRunElement.GetBookmarkName() + '"' : ""}}`);
// 			} else if (oRunElement instanceof AscCommonWord.ParaRun && oRunElement.Content.length) {
// 				const prChange = getPr(oRunElement.Pr);
// 				const revChange = getReview(oRunElement);
// 				if (test) {
// 					const prChange = getPr(oRunElement.Pr);
// 					const revChange = getReview(oRunElement);
// 					if (!compareReviewDiff(revChange, test.revChange) || !compareTextPrDiff(prChange, test.prChange) || arrBookmarks.length || arrComments.length) {
// 						test = {
// 							bookmarks: arrBookmarks,
// 							comments: arrComments,
// 							text: oRunElement.GetText(),
// 							revChange: revChange,
// 							prChange: prChange
// 						};
// 						arrComments = [];
// 						arrBookmarks = [];
// 						paragraphTest.push(test);
// 					} else {
// 						test.text += oRunElement.GetText();
// 					}
// 				} else {
// 					test = {
// 						bookmarks: arrBookmarks,
// 						comments: arrComments,
// 						text: oRunElement.GetText(),
// 						revChange: revChange,
// 						prChange: prChange
// 					};
// 					paragraphTest.push(test);
// 					arrComments = [];
// 					arrBookmarks = [];
// 				}
// 			}
// 		}
// 		if (arrBookmarks.length || arrComments.length) {
// 			paragraphTest.push({bookmarks: arrBookmarks, comments: arrComments});
// 			arrComments = [];
// 			arrBookmarks = [];
// 		}
// 	}
//
//
// 	let result = [];
// 	for (let i = 0; i < arrTests.length; i++) {
// 		const paragraphTest = arrTests[i];
// 		const paragraphResult = [];
// 		for (let j = 0; j < paragraphTest.length; j++) {
// 			const runInfo = paragraphTest[j];
// 			const stext = runInfo.text ? '"' + runInfo.text + '"' : "undefined";
// 			let review = "undefined";
// 			let addReview = "undefined"
// 			if (runInfo.revChange) {
// 				if (runInfo.revChange.mainReviewType === reviewtype_Add) {
// 					review = "new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000)";
// 				} else if (runInfo.revChange.mainReviewType === reviewtype_Remove) {
// 					review = "new CCreatingReviewInfo('Mark Potato', reviewtype_Remove, 1000)";
// 				}
// 				if (runInfo.revChange.additionalReviewType === reviewtype_Add) {
// 					addReview = "new CCreatingReviewInfo('Mark Potato', reviewtype_Add, 1000)";
// 				}
// 			}
// 			let bookmarks = "undefined";
// 			if (runInfo.bookmarks && runInfo.bookmarks.length) {
// 				bookmarks = `{start: [${runInfo.bookmarks.join(", ")}]}`;
// 			}
// 			let comments = "";
// 			if (runInfo.comments && runInfo.comments.length) {
// 				comments = `comments: {start: [${runInfo.comments.join(", ")}]},`;
// 			}
// 			let textpr = "";
// 			if (runInfo.prChange) {
// 				for (let sName in runInfo.prChange) {
// 					if (!textpr) {
// 						textpr = "textPr: {"
// 					}
// 					textpr += `${sName}: ${runInfo.prChange[sName]},`
// 				}
// 				if (textpr) {
// 					textpr += "}"
// 				}
// 			}
// 			paragraphResult.push(`createParagraphInfo(${stext}, ${review}, ${addReview}, ${bookmarks}, {${comments}${textpr}})`)
// 		}
// 		result.push("[" + paragraphResult.join(", ") + "]")
// 	}
// 	console.log(result.join(", "))
// };
