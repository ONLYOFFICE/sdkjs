/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
mockEditor.pre_Paste = function (first, second, callback) {
    callback();
};
AscCommon.sendImgUrls = function (oApi, arrImages, fCallback) {fCallback()};
AscCommon.ResetNewUrls = function () {};

AscCommonWord.CDocument.prototype.getTestObject = function () {
    const oContentObject = {type: 'document', content: []};
    this.Content.forEach(function (oItem) {
        if (oItem.getTestObject) {
            oItem.getTestObject(oContentObject.content);
        } else {
            oContentObject.content.push(oItem.constructor.name);
        }
    });
    if (this.SectPr) {
        const arrHdrFtr = this.SectPr.GetAllHdrFtrs();
        for (let i = 0; i < arrHdrFtr.length; i += 1) {
            arrHdrFtr[i].getTestObject(oContentObject.content);
        }

    }
    if (this.Footnotes) {
        this.Footnotes.getTestObject(oContentObject.content);
    }
    return oContentObject;
};

AscCommonWord.CHeaderFooter.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'headerfooter', content: []};
    arrParentContent.push(oContentObject);
    this.Content.getTestObject(oContentObject.content);
};

AscCommonWord.CTable.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'table', rows: []};
    arrParentContent.push(oContentObject);
    for (let i = 0; i < this.Content.length; i += 1) {
        const row = this.Content[i];
        row.getTestObject(oContentObject.rows);
    }
}

AscCommonWord.CTableRow.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'row', content: []};
    arrParentContent.push(oContentObject);
    for (let i = 0; i < this.Content.length; i += 1) {
        const cell = this.Content[i];
        cell.getTestObject(oContentObject.content);
    }
}

AscCommonWord.CTableCell.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'cell', content: []};
    arrParentContent.push(oContentObject);
    const oContent = this.GetContent();
    oContent.CheckRunContent(function (oRun) {
        oRun.getTestObject(oContentObject.content);
    });
}


ParaMath.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'paramath', content: []};
    arrParentContent.push(oContentObject)
    this.Root.getTestObject(oContentObject.content);
}
CMathContent.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'mathcontent', content: []};
    arrParentContent.push(oContentObject)
    for (var i = 0; i < this.Content.length; ++i)
    {
        if (para_Math_Run === this.Content[i].Type)
            this.Content[i].getTestObject(oContentObject.content);
    }
}
CMathBase.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'mathbase', content: []};
    arrParentContent.push(oContentObject)
    this.Content.forEach(function (oRun) {
        oRun.getTestObject(oContentObject.content);
    });
}
CDocumentSectionsInfo.prototype.getTestObject = function (arrParentContent) {
    const arrHeaders = this.GetAllHdrFtrs();
    for (let index = 0, count = arrHeaders.length; index < count; ++index)
    {
        const oContentObject = {
            type: 'documentsectioninfo',
            content: []
        }
        arrParentContent.push(oContentObject);
        arrHeaders[index].getTestObject(oContentObject.content);
    }
}
CDocumentContentBase.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'documentcontentbase', content: []};
    arrParentContent.push(oContentObject)
    for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
    {
        this.Content[nIndex].getTestObject(oContentObject.content);
    }
}
CDocumentContentElementBase.prototype.getTestObject = function () {

};
CEndnotesController.prototype.getTestObject = function (arrParentContent) {
    for (var sId in this.Endnote) {
        const oEndnote = this.Endnote[sId];
        const oContentObject = {type: 'endnote', content: []};
        arrParentContent.push(oContentObject);
        oEndnote.checkTestObject(oContentObject.content)
    }
};
CFootnotesController.prototype.getTestObject = function (arrParentContent) {
    for (var sId in this.Footnote) {
        const oFootnote = this.Footnote[sId];
        const oContentObject = {type: 'footnote', content: []};
        arrParentContent.push(oContentObject);
        for (let i = 0; i < oFootnote.Content.length; i += 1) {
            oFootnote.Content[i].getTestObject(oContentObject.content);
        }
    }
};
CParagraphContentBase.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'paragraphcontentbase', content: []};
    arrParentContent.push(oContentObject)
    for (let i = 0; i < this.Content.length;i += 1) {
        this.Content[i].getTestObject(oContentObject.content);
    }
};
CParagraphContentWithParagraphLikeContent.prototype.getTestObject = function () {

};
CBlockLevelSdt.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'blocklvlsdt', content: []};
    arrParentContent.push(oContentObject)
    this.Content.getTestObject(oContentObject.content);
};
Paragraph.prototype.getTestObject = function (arrParentContent) {
    const oContentObject = {type: 'paragraph', content: []};
    arrParentContent.push(oContentObject);
    const oTestParagraphContent =
        this.CheckRunContent(function (oRun) {
            oRun.getTestObject(oContentObject.content);
        });
};
ParaRun.prototype.getTestObject = function (oParentContent) {
    if (this.Content.length === 0) return;
    const oReviewInfo = this.GetReviewInfo();
    const oPrevAdded = oReviewInfo.GetPrevAdded();
    let nMainReviewType = this.GetReviewType && this.GetReviewType();
    let sMainUserName = oReviewInfo.GetUserName();
    let nMainDateTime = oReviewInfo.GetDateTime();

    let nAdditionalReviewType;
    let sAdditionalUserName;
    let nAdditionalDateTime;

    if (oPrevAdded) {
        nAdditionalReviewType = reviewtype_Add;
        sAdditionalUserName = oPrevAdded.GetUserName();
        nAdditionalDateTime = oPrevAdded.GetDateTime();
    }
    let oCurrentTextInfo = oParentContent[oParentContent.length - 1];
    const needCreateNewText = (oParentContent.length === 0 ||
        oCurrentTextInfo.mainReviewType !== nMainReviewType || oCurrentTextInfo.mainUserName !== sMainUserName || oCurrentTextInfo.mainDateTime !== nMainDateTime ||
        oCurrentTextInfo.additionalReviewType !== nAdditionalReviewType || oCurrentTextInfo.additionalUserName !== sAdditionalUserName || oCurrentTextInfo.additionalDateTime !== nAdditionalDateTime);
    if (needCreateNewText || this.IsParaEndRun()) {
        oCurrentTextInfo = {
            mainReviewType: nMainReviewType,
            mainDateTime: nMainDateTime,
            mainUserName: sMainUserName,
            additionalReviewType: nAdditionalReviewType,
            additionalDateTime: nAdditionalDateTime,
            additionalUserName: sAdditionalUserName,
            text: ''
        };
        oParentContent.push(oCurrentTextInfo);
    }
    this.Content.forEach(function (el) {
        oCurrentTextInfo.text += String.fromCharCode(el.Value)
    });
};

window['AscCommonWord']['CDocumentComparison'].prototype.setReviewInfo = function(oReviewInfo, sCustomReviewUserName, nCustomReviewDate)
{
    oReviewInfo.Editor   = this.api;
    oReviewInfo.UserId   = "";
    oReviewInfo.MoveType = Asc.c_oAscRevisionsMove.NoMove;
    oReviewInfo.PrevType = -1;
    oReviewInfo.PrevInfo = null;
    oReviewInfo.UserName = sCustomReviewUserName || "Valdemar";
    oReviewInfo.DateTime = 3000000;
    if (AscFormat.isRealNumber(nCustomReviewDate)) {
        oReviewInfo.DateTime = nCustomReviewDate;
    }
};
function readMainDocument(oMainDocumentInfo) {
    const oDocument = new AscWord.CDocument(mockEditor.WordControl.m_oDrawingDocument, true);
    mockEditor.WordControl.m_oDrawingDocument.m_oLogicDocument = oDocument;
    mockEditor.WordControl.m_oLogicDocument = oDocument;
    oDocument.Api = mockEditor;
    createTestDocument(oDocument, oMainDocumentInfo);
    return oDocument
}

function readRevisedDocument(oRevisedDocumentInfo) {
    const oMainDocument = mockEditor.WordControl.m_oLogicDocument;
    const oRevisedDocument = new CDocument(mockEditor.WordControl.m_oDrawingDocument, true);
    mockEditor.WordControl.m_oDrawingDocument.m_oLogicDocument = oRevisedDocument;
    mockEditor.WordControl.m_oLogicDocument = oRevisedDocument;

    createTestDocument(oRevisedDocument, oRevisedDocumentInfo);
    
    mockEditor.WordControl.m_oDrawingDocument.m_oLogicDocument = oMainDocument;
    mockEditor.WordControl.m_oLogicDocument = oMainDocument;
    if (oMainDocument.History)
        oMainDocument.History.Set_LogicDocument(oMainDocument);
    
    if (oMainDocument.CollaborativeEditing)
        oMainDocument.CollaborativeEditing.m_oLogicDocument = oMainDocument;
    
    return oRevisedDocument;
}

function createTestDocument(oDocument, arrParagraphsTextInfo) {
    for (let i = 0; i < arrParagraphsTextInfo.length; i += 1) {
        const oParagraphTextInfo = arrParagraphsTextInfo[i];
        let oParagraph;
        if (i === 0) {
            oParagraph = oDocument.Content[0];
        } else {
            oParagraph = AscTest.CreateParagraph();
        }
        
        for (let j = 0; j < oParagraphTextInfo.length; j += 1) {
            const oParaRun = new AscWord.ParaRun();
            if (oParagraphTextInfo[j].text) {
                oParaRun.AddText(oParagraphTextInfo[j].text);
                oParaRun.SetReviewTypeWithInfo(oParagraphTextInfo[j].reviewType, oParagraphTextInfo[j].reviewInfo);
                oParagraph.AddToContentToEnd(oParaRun);
            } else {
                oParagraph.GetParaEndRun().SetReviewTypeWithInfo(oParagraphTextInfo[j].reviewType, oParagraphTextInfo[j].reviewInfo, false);
            }
        }
        if (i !== 0) {
            oDocument.AddToContent(oDocument.Content.length, oParagraph);
        }
    }
    return oDocument;
}

function createParagraphInfo(sText, oMainReviewInfoOptions, oAdditionalReviewInfoOptions) {
    const oResult = {
        text: sText,
        reviewType: reviewtype_Common
    };
    let oMainReviewInfo;
    if (oMainReviewInfoOptions) {
        oMainReviewInfo = createReviewInfoFromOptions(oMainReviewInfoOptions);
        oResult.reviewType = oMainReviewInfoOptions.reviewType;
        if (oAdditionalReviewInfoOptions) {
            const oAdditionalReviewInfo = createReviewInfoFromOptions(oAdditionalReviewInfoOptions);
            oAdditionalReviewInfo.SavePrev(oAdditionalReviewInfoOptions.reviewType);
            oMainReviewInfo.PrevType = oAdditionalReviewInfo.PrevType;
            oMainReviewInfo.PrevInfo = oAdditionalReviewInfo.PrevInfo;
        }
    } else {
        oMainReviewInfo = createReviewInfoFromOptions();
    }
    oResult.reviewInfo = oMainReviewInfo;
    return oResult;
}

function createShapeInfo() {
    
}

function createReviewInfoFromOptions(oOptions) {
    oOptions = oOptions || {};
    const oReviewInfo = new CReviewInfo();

    oReviewInfo.Editor = mockEditor;
    oReviewInfo.UserId   = "";
    oReviewInfo.MoveType = Asc.c_oAscRevisionsMove.NoMove;
    oReviewInfo.PrevType = -1;
    oReviewInfo.PrevInfo = null;
    oReviewInfo.UserName = oOptions.userName || oReviewInfo.UserName;
    oReviewInfo.DateTime = oOptions.dateTime || oReviewInfo.DateTime;

    return oReviewInfo;
}
