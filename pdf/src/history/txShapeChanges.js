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

AscDFH.changesFactory[AscDFH.historyitem_Pdf_TxShape_Rect]		= CChangesPDFTxShapeRect;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_TxShape_Page]		= CChangesPDFTxShapePage;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_TxShape_RC]		= CChangesPDFTxShapeRC;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_TxShape_Rot]		= CChangesPDFTxShapeRot;


/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFTxShapeRect(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTxShapeRect.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFTxShapeRect.prototype.constructor = CChangesPDFTxShapeRect;
CChangesPDFTxShapeRect.prototype.Type = AscDFH.historyitem_Pdf_Annot_Rect;
CChangesPDFTxShapeRect.prototype.private_SetValue = function(Value)
{
	let oTxShape = this.Class;
	oTxShape.SetRect(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFTxShapePage(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTxShapePage.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFTxShapePage.prototype.constructor = CChangesPDFTxShapePage;
CChangesPDFTxShapePage.prototype.Type = AscDFH.historyitem_Pdf_TxShape_Page;
CChangesPDFTxShapePage.prototype.private_SetValue = function(Value)
{
	let oTxShape = this.Class;
	oTxShape.SetPage(Value, true);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFTxShapeRC(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTxShapeRC.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFTxShapeRC.prototype.constructor = CChangesPDFTxShapeRC;
CChangesPDFTxShapeRC.prototype.Type = AscDFH.historyitem_Pdf_TxShape_RC;
CChangesPDFTxShapeRC.prototype.private_SetValue = function(Value)
{
	let oTxShape = this.Class;
	oTxShape.SetRichContents(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFTxShapeRot(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTxShapeRot.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFTxShapeRot.prototype.constructor = CChangesPDFTxShapeRot;
CChangesPDFTxShapeRot.prototype.Type = AscDFH.historyitem_Pdf_TxShape_RC;
CChangesPDFTxShapeRot.prototype.private_SetValue = function(Value)
{
	let oTxShape = this.Class;
	oTxShape.SetRot(Value);
};

