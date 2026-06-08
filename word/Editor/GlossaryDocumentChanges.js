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

AscDFH.changesFactory[AscDFH.historyitem_GlossaryDocument_AddDocPart] = CChangesGlossaryAddDocPart;

AscDFH.changesFactory[AscDFH.historyitem_DocPart_Name]        = CChangesDocPartName;
AscDFH.changesFactory[AscDFH.historyitem_DocPart_Style]       = CChangesDocPartStyle;
AscDFH.changesFactory[AscDFH.historyitem_DocPart_Types]       = CChangesDocPartTypes;
AscDFH.changesFactory[AscDFH.historyitem_DocPart_Description] = CChangesDocPartDescription;
AscDFH.changesFactory[AscDFH.historyitem_DocPart_GUID]        = CChangesDocPartGUID;
AscDFH.changesFactory[AscDFH.historyitem_DocPart_Category]    = CChangesDocPartCategory;
AscDFH.changesFactory[AscDFH.historyitem_DocPart_Behavior]    = CChangesDocPartBehavior;
//----------------------------------------------------------------------------------------------------------------------
// Карта зависимости изменений
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_GlossaryDocument_AddDocPart] = [AscDFH.historyitem_GlossaryDocument_AddDocPart];

AscDFH.changesRelationMap[AscDFH.historyitem_DocPart_Name]        = [AscDFH.historyitem_DocPart_Name];
AscDFH.changesRelationMap[AscDFH.historyitem_DocPart_Style]       = [AscDFH.historyitem_DocPart_Style];
AscDFH.changesRelationMap[AscDFH.historyitem_DocPart_Types]       = [AscDFH.historyitem_DocPart_Types];
AscDFH.changesRelationMap[AscDFH.historyitem_DocPart_Description] = [AscDFH.historyitem_DocPart_Description];
AscDFH.changesRelationMap[AscDFH.historyitem_DocPart_GUID]        = [AscDFH.historyitem_DocPart_GUID];
AscDFH.changesRelationMap[AscDFH.historyitem_DocPart_Category]    = [AscDFH.historyitem_DocPart_Category];
AscDFH.changesRelationMap[AscDFH.historyitem_DocPart_Behavior]    = [AscDFH.historyitem_DocPart_Behavior];
//----------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesGlossaryAddDocPart(Class, Id)
{
	AscDFH.CChangesBase.call(this, Class);
	this.Id = Id;
}
CChangesGlossaryAddDocPart.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesGlossaryAddDocPart.prototype.constructor = CChangesGlossaryAddDocPart;
CChangesGlossaryAddDocPart.prototype.Type = AscDFH.historyitem_GlossaryDocument_AddDocPart;
CChangesGlossaryAddDocPart.prototype.Undo = function()
{
	delete this.Class.DocParts[this.Id];
};
CChangesGlossaryAddDocPart.prototype.Redo = function()
{
	this.Class.DocParts[this.Id] = AscCommon.g_oTableId.Get_ById(this.Id);
};
CChangesGlossaryAddDocPart.prototype.WriteToBinary = function(Writer)
{
	// String : Id
	Writer.WriteString2(this.Id);
};
CChangesGlossaryAddDocPart.prototype.ReadFromBinary = function(Reader)
{
	// String : Id
	this.Id = Reader.GetString2();
};
CChangesGlossaryAddDocPart.prototype.CreateReverseChange = function()
{
	return null;
};
CChangesGlossaryAddDocPart.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type && this.Id === oChange.Id)
		return false;

	return true;
};


/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesDocPartName(Class, Old, New)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New);
}
CChangesDocPartName.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesDocPartName.prototype.constructor = CChangesDocPartName;
CChangesDocPartName.prototype.Type = AscDFH.historyitem_DocPart_Name;
CChangesDocPartName.prototype.private_SetValue = function(Value)
{
	this.Class.Pr.Name = Value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesDocPartStyle(Class, Old, New)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New);
}
CChangesDocPartStyle.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesDocPartStyle.prototype.constructor = CChangesDocPartStyle;
CChangesDocPartStyle.prototype.Type = AscDFH.historyitem_DocPart_Style;
CChangesDocPartStyle.prototype.private_SetValue = function(Value)
{
	this.Class.Pr.Style = Value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesDocPartTypes(Class, Old, New)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
}
CChangesDocPartTypes.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesDocPartTypes.prototype.constructor = CChangesDocPartTypes;
CChangesDocPartTypes.prototype.Type = AscDFH.historyitem_DocPart_Types;
CChangesDocPartTypes.prototype.private_SetValue = function(Value)
{
	this.Class.Pr.Types = Value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesDocPartDescription(Class, Old, New)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New);
}
CChangesDocPartDescription.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesDocPartDescription.prototype.constructor = CChangesDocPartDescription;
CChangesDocPartDescription.prototype.Type = AscDFH.historyitem_DocPart_Description;
CChangesDocPartDescription.prototype.private_SetValue = function(Value)
{
	this.Class.Pr.Description = Value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesDocPartGUID(Class, Old, New)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New);
}
CChangesDocPartGUID.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesDocPartGUID.prototype.constructor = CChangesDocPartGUID;
CChangesDocPartGUID.prototype.Type = AscDFH.historyitem_DocPart_GUID;
CChangesDocPartGUID.prototype.private_SetValue = function(Value)
{
	this.Class.Pr.GUID = Value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesDocPartCategory(Class, Old, New)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New);
}
CChangesDocPartCategory.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesDocPartCategory.prototype.constructor = CChangesDocPartCategory;
CChangesDocPartCategory.prototype.Type = AscDFH.historyitem_DocPart_Category;
CChangesDocPartCategory.prototype.private_CreateObject = function()
{
	return new CDocPartCategory();
};
CChangesDocPartCategory.prototype.private_SetValue = function(Value)
{
	this.Class.Pr.Category = Value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesDocPartBehavior(Class, Old, New)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
}
CChangesDocPartBehavior.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesDocPartBehavior.prototype.constructor = CChangesDocPartBehavior;
CChangesDocPartBehavior.prototype.Type = AscDFH.historyitem_DocPart_Behavior;
CChangesDocPartBehavior.prototype.private_SetValue = function(Value)
{
	this.Class.Pr.Behavior = Value;
};

