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

"use strict";

AscDFH.changesFactory[AscDFH.historyitem_Pdf_Comment_Data]		= CChangesPDFCommentData;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Ink_Points]		= CChangesPDFInkPoints;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Ink_FlipV]			= CChangesPDFInkFlipV;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Ink_FlipH]			= CChangesPDFInkFlipH;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Annot_Rect]		= CChangesPDFAnnotRect;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Annot_Contents]	= CChangesPDFAnnotContents;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Annot_Pos]			= CChangesPDFAnnotPos;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Annot_Page]		= CChangesPDFAnnotPage;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Annot_Replies]		= CChangesPDFAnnotReplies;

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFCommentData(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFCommentData.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFCommentData.prototype.constructor = CChangesPDFCommentData;
CChangesPDFCommentData.prototype.Type = AscDFH.historyitem_Pdf_Comment_Data;
CChangesPDFCommentData.prototype.private_SetValue = function(Value)
{
	let oComment = this.Class;
	oComment.EditCommentData(Value);
	let AscCommentData = oComment.GetAscCommentData();
	let CommentData = new AscCommon.CCommentData();
	CommentData.Read_FromAscCommentData(AscCommentData);

	editor.sync_ChangeCommentData(oComment.GetId(), CommentData);
};


/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFInkPoints(Class, Pos, Items, Color)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesPDFInkPoints.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFInkPoints.prototype.constructor = CChangesPDFInkPoints;
CChangesPDFInkPoints.prototype.Type = AscDFH.historyitem_Pdf_Ink_Points;
CChangesPDFInkPoints.prototype.Undo = function()
{
	this.Class.RemoveLastAddedPath();
};
CChangesPDFInkPoints.prototype.Redo = function()
{
	this.Class.AddPath(this.Items);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFInkFlipV(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFInkFlipV.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFInkFlipV.prototype.constructor = CChangesPDFInkFlipV;
CChangesPDFInkFlipV.prototype.Type = AscDFH.historyitem_Pdf_Ink_FlipV;
CChangesPDFInkFlipV.prototype.private_SetValue = function(Value)
{
	let oAnnot = this.Class;
	oAnnot.SetFlipV(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFInkFlipH(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFInkFlipH.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFInkFlipH.prototype.constructor = CChangesPDFInkFlipH;
CChangesPDFInkFlipH.prototype.Type = AscDFH.historyitem_Pdf_Ink_FlipH;
CChangesPDFInkFlipH.prototype.private_SetValue = function(Value)
{
	let oAnnot = this.Class;
	oAnnot.SetFlipH(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFAnnotRect(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFAnnotRect.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFAnnotRect.prototype.constructor = CChangesPDFAnnotRect;
CChangesPDFAnnotRect.prototype.Type = AscDFH.historyitem_Pdf_Annot_Rect;
CChangesPDFAnnotRect.prototype.private_SetValue = function(Value)
{
	let oAnnot = this.Class;
	oAnnot.SetRect(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFAnnotPos(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFAnnotPos.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFAnnotPos.prototype.constructor = CChangesPDFAnnotPos;
CChangesPDFAnnotPos.prototype.Type = AscDFH.historyitem_Pdf_Annot_Pos;
CChangesPDFAnnotPos.prototype.private_SetValue = function(Value)
{
	let oAnnot = this.Class;
	oAnnot.SetPosition(Value[0], Value[1], true);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFAnnotContents(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFAnnotContents.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFAnnotContents.prototype.constructor = CChangesPDFAnnotContents;
CChangesPDFAnnotContents.prototype.Type = AscDFH.historyitem_Pdf_Annot_Contents;
CChangesPDFAnnotContents.prototype.private_SetValue = function(Value)
{
	let oAnnot = this.Class;
	oAnnot.SetContents(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFAnnotReplies(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFAnnotReplies.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFAnnotReplies.prototype.constructor = CChangesPDFAnnotReplies;
CChangesPDFAnnotReplies.prototype.Type = AscDFH.historyitem_Pdf_Annot_Replies;
CChangesPDFAnnotReplies.prototype.private_SetValue = function(Value)
{
	let oAnnot = this.Class;
	oAnnot.SetReplies(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFAnnotPage(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFAnnotPage.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFAnnotPage.prototype.constructor = CChangesPDFAnnotPage;
CChangesPDFAnnotPage.prototype.Type = AscDFH.historyitem_Pdf_Annot_Page;
CChangesPDFAnnotPage.prototype.private_SetValue = function(Value)
{
	let oAnnot = this.Class;
	oAnnot.SetPage(Value, true);
};
