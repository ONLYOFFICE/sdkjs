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

AscDFH.changesFactory[AscDFH.historyitem_Document_AddItem]                         = CChangesDocumentAddItem;
AscDFH.changesFactory[AscDFH.historyitem_Document_RemoveItem]                      = CChangesDocumentRemoveItem;
AscDFH.changesFactory[AscDFH.historyitem_Document_DefaultTab]                      = CChangesDocumentDefaultTab;
AscDFH.changesFactory[AscDFH.historyitem_Document_EvenAndOddHeaders]               = CChangesDocumentEvenAndOddHeaders;
AscDFH.changesFactory[AscDFH.historyitem_Document_DefaultLanguage]                 = CChangesDocumentDefaultLanguage;
AscDFH.changesFactory[AscDFH.historyitem_Document_MathSettings]                    = CChangesDocumentMathSettings;
AscDFH.changesFactory[AscDFH.historyitem_Document_SdtGlobalSettings]               = CChangesDocumentSdtGlobalSettings;
AscDFH.changesFactory[AscDFH.historyitem_Document_Settings_GutterAtTop]            = CChangesDocumentSettingsGutterAtTop;
AscDFH.changesFactory[AscDFH.historyitem_Document_Settings_MirrorMargins]          = CChangesDocumentSettingsMirrorMargins;
AscDFH.changesFactory[AscDFH.historyitem_Document_SpecialFormsGlobalSettings]      = CChangesDocumentSpecialFormsGlobalSettings;
AscDFH.changesFactory[AscDFH.historyitem_Document_Settings_TrackRevisions]         = CChangesDocumentSettingsTrackRevisions;
AscDFH.changesFactory[AscDFH.historydescription_Document_DocumentProtection]       = CChangesDocumentProtection;
AscDFH.changesFactory[AscDFH.historyitem_Document_Settings_AutoHyphenation]        = CChangesDocumentSettingsAutoHyphenation;
AscDFH.changesFactory[AscDFH.historyitem_Document_Settings_ConsecutiveHyphenLimit] = CChangesDocumentSettingsConsecutiveHyphenLimit;
AscDFH.changesFactory[AscDFH.historyitem_Document_Settings_DoNotHyphenateCaps]     = CChangesDocumentSettingsDoNotHyphenateCaps;
AscDFH.changesFactory[AscDFH.historyitem_Document_Settings_HyphenationZone]        = CChangesDocumentSettingsHyphenationZone;
AscDFH.changesFactory[AscDFH.historyitem_Document_PageColor]                       = CChangesDocumentPageColor;

AscDFH.changesFactory[AscDFH.historyitem_Document_DisconnectEveryone]              = CChangesDocumentDisconnectEveryone;
//----------------------------------------------------------------------------------------------------------------------
// Карта зависимости изменений
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_Document_AddItem]                = [
	AscDFH.historyitem_Document_AddItem,
	AscDFH.historyitem_Document_RemoveItem
];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_RemoveItem]             = [
	AscDFH.historyitem_Document_AddItem,
	AscDFH.historyitem_Document_RemoveItem
];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_DefaultTab]                 = [AscDFH.historyitem_Document_DefaultTab];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_EvenAndOddHeaders]          = [AscDFH.historyitem_Document_EvenAndOddHeaders];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_DefaultLanguage]            = [AscDFH.historyitem_Document_DefaultLanguage];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_MathSettings]               = [AscDFH.historyitem_Document_MathSettings];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_SdtGlobalSettings]          = [AscDFH.historyitem_Document_SdtGlobalSettings];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_Settings_GutterAtTop]       = [AscDFH.historyitem_Document_Settings_GutterAtTop];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_Settings_MirrorMargins]     = [AscDFH.historyitem_Document_Settings_MirrorMargins];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_SpecialFormsGlobalSettings] = [AscDFH.historyitem_Document_SpecialFormsGlobalSettings];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_Settings_TrackRevisions]    = [AscDFH.historyitem_Document_Settings_TrackRevisions];
AscDFH.changesRelationMap[AscDFH.historyitem_Document_DisconnectEveryone]         = [AscDFH.historyitem_Document_DisconnectEveryone];
//----------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesDocumentAddItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesDocumentAddItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesDocumentAddItem.prototype.constructor = CChangesDocumentAddItem;
CChangesDocumentAddItem.prototype.Type = AscDFH.historyitem_Document_AddItem;
CChangesDocumentAddItem.prototype.Undo = function()
{
	var oDocument = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var Pos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
		var Elements = oDocument.Content.splice(Pos, 1);
		oDocument.private_RecalculateNumbering(Elements);
		oDocument.private_ReindexContent(Pos);
		if (oDocument.SectionsInfo)
		{
			oDocument.SectionsInfo.Update_OnRemove(Pos, 1);
		}

		oDocument.private_UpdateSelectionPosOnRemove(Pos, 1);

		if (Pos > 0)
		{
			if (Pos <= oDocument.Content.length - 1)
			{
				oDocument.Content[Pos - 1].Next = oDocument.Content[Pos];
				oDocument.Content[Pos].Prev     = oDocument.Content[Pos - 1];
			}
			else
			{
				oDocument.Content[Pos - 1].Next = null;
			}
		}
		else if (Pos <= oDocument.Content.length - 1)
		{
			oDocument.Content[Pos].Prev = null;
		}
	}
};
CChangesDocumentAddItem.prototype.Redo = function()
{
	var oDocument = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var Element = this.Items[nIndex];
		var Pos     = true !== this.UseArray ? this.Pos + nIndex : this.PosArray[nIndex];

		oDocument.Content.splice(Pos, 0, Element);
		oDocument.private_RecalculateNumbering([Element]);
		oDocument.private_ReindexContent(Pos);
		if (oDocument.SectionsInfo)
		{
			oDocument.SectionsInfo.Update_OnAdd(Pos, [Element]);
		}

		oDocument.private_UpdateSelectionPosOnAdd(Pos, 1);

		if (Pos > 0)
		{
			oDocument.Content[Pos - 1].Next = Element;
			Element.Prev                    = oDocument.Content[Pos - 1];
		}
		else
		{
			Element.Prev = null;
		}

		if (Pos < oDocument.Content.length - 1)
		{
			oDocument.Content[Pos + 1].Prev = Element;
			Element.Next                    = oDocument.Content[Pos + 1];
		}
		else
		{
			Element.Next = null;
		}

		Element.Parent = oDocument;
	}
};
CChangesDocumentAddItem.prototype.private_WriteItem = function(Writer, Item)
{
	Writer.WriteString2(Item.Get_Id());
};
CChangesDocumentAddItem.prototype.private_ReadItem = function(Reader)
{
	return AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesDocumentAddItem.prototype.Load = function(Color)
{
	var oDocument = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var Pos     = oDocument.m_oContentChanges.Check(AscCommon.contentchanges_Add, true !== this.UseArray ? this.Pos + nIndex : this.PosArray[nIndex]);
		var Element = this.Items[nIndex];

		Pos = Math.min(Pos, oDocument.Content.length);

		if (null != Element)
		{
			if (Pos > 0)
			{
				oDocument.Content[Pos - 1].Next = Element;
				Element.Prev                    = oDocument.Content[Pos - 1];
			}
			else
			{
				Element.Prev = null;
			}

			if (Pos <= oDocument.Content.length - 1)
			{
				oDocument.Content[Pos].Prev = Element;
				Element.Next                = oDocument.Content[Pos];
			}
			else
			{
				Element.Next = null;
			}

			Element.Parent = oDocument;

			oDocument.Content.splice(Pos, 0, Element);
			oDocument.private_RecalculateNumbering([Element]);
			if (oDocument.SectionsInfo)
			{
				oDocument.SectionsInfo.Update_OnAdd(Pos, [Element]);
			}
			oDocument.private_ReindexContent(Pos);

			AscCommon.CollaborativeEditing.Update_DocumentPositionsOnAdd(oDocument, Pos);
			oDocument.private_UpdateSelectionPosOnAdd(Pos, 1);

			if (Element.IsParagraph())
			{
				Element.RecalcCompiledPr(true);
				Element.UpdateDocumentOutline();
			}
		}
	}
};
CChangesDocumentAddItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_Document_AddItem === oChanges.Type || AscDFH.historyitem_Document_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesDocumentAddItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesDocumentRemoveItem);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesDocumentRemoveItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, false);
}
CChangesDocumentRemoveItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesDocumentRemoveItem.prototype.constructor = CChangesDocumentRemoveItem;
CChangesDocumentRemoveItem.prototype.Type = AscDFH.historyitem_Document_RemoveItem;
CChangesDocumentRemoveItem.prototype.Undo = function()
{
	var oDocument = this.Class;

	var Array_start = oDocument.Content.slice(0, this.Pos);
	var Array_end   = oDocument.Content.slice(this.Pos);

	oDocument.private_RecalculateNumbering(this.Items);
	oDocument.private_ReindexContent(this.Pos);
	oDocument.Content = Array_start.concat(this.Items, Array_end);

	if(oDocument.SectionsInfo)
	{
        oDocument.SectionsInfo.Update_OnAdd(this.Pos, this.Items);
	}

	oDocument.private_UpdateSelectionPosOnAdd(this.Pos, this.Items.length);

	var nStartIndex = Math.max(this.Pos - 1, 0);
	var nEndIndex   = Math.min(oDocument.Content.length - 1, this.Pos + this.Items.length + 1);
	for (var nIndex = nStartIndex; nIndex <= nEndIndex; ++nIndex)
	{
		var oElement = oDocument.Content[nIndex];
		if (nIndex > 0)
			oElement.Prev = oDocument.Content[nIndex - 1];
		else
			oElement.Prev = null;

		if (nIndex < oDocument.Content.length - 1)
			oElement.Next = oDocument.Content[nIndex + 1];
		else
			oElement.Next = null;

		oElement.Parent = oDocument;
	}
};
CChangesDocumentRemoveItem.prototype.Redo = function()
{
	var oDocument = this.Class;
	var Elements = oDocument.Content.splice(this.Pos, this.Items.length);
	oDocument.private_RecalculateNumbering(Elements);
	oDocument.private_ReindexContent(this.Pos);
	if(oDocument.SectionsInfo)
	{
        oDocument.SectionsInfo.Update_OnRemove(this.Pos, this.Items.length);
	}

	oDocument.private_UpdateSelectionPosOnRemove(this.Pos, this.Items.length);

	var Pos = this.Pos;
	if (Pos > 0)
	{
		if (Pos <= oDocument.Content.length - 1)
		{
			oDocument.Content[Pos - 1].Next = oDocument.Content[Pos];
			oDocument.Content[Pos].Prev     = oDocument.Content[Pos - 1];
		}
		else
		{
			oDocument.Content[Pos - 1].Next = null;
		}
	}
	else if (Pos <= oDocument.Content.length - 1)
	{
		oDocument.Content[Pos].Prev = null;
	}
};
CChangesDocumentRemoveItem.prototype.private_WriteItem = function(Writer, Item)
{
	Writer.WriteString2(Item.Get_Id());
};
CChangesDocumentRemoveItem.prototype.private_ReadItem = function(Reader)
{
	return AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesDocumentRemoveItem.prototype.Load = function(Color)
{
	var oDocument = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var Pos = oDocument.m_oContentChanges.Check(AscCommon.contentchanges_Remove, true !== this.UseArray ? this.Pos : this.PosArray[nIndex]);

		// действие совпало, не делаем его
		if (false === Pos)
			continue;

		var Elements = oDocument.Content.splice(Pos, 1);
		oDocument.private_RecalculateNumbering(Elements);
		AscCommon.CollaborativeEditing.Update_DocumentPositionsOnRemove(oDocument, Pos, 1);
		oDocument.private_UpdateSelectionPosOnRemove(Pos, 1);

		for (var nElementIndex = 0; nElementIndex < Elements.length; ++nElementIndex)
		{
			if (Elements[nElementIndex].IsParagraph())
			{
				Elements[nElementIndex].RecalcCompiledPr(true);
				Elements[nElementIndex].UpdateDocumentOutline();
			}
		}

		if (Pos > 0)
		{
			if (Pos <= oDocument.Content.length - 1)
			{
				oDocument.Content[Pos - 1].Next = oDocument.Content[Pos];
				oDocument.Content[Pos].Prev     = oDocument.Content[Pos - 1];
			}
			else
			{
				oDocument.Content[Pos - 1].Next = null;
			}
		}
		else if (Pos <= oDocument.Content.length - 1)
		{
			oDocument.Content[Pos].Prev = null;
		}

		if (0 <= Pos && Pos <= oDocument.Content.length - 1)
		{
			if (oDocument.SectionsInfo)
			{
				oDocument.SectionsInfo.Update_OnRemove(Pos, 1);
			}

			oDocument.private_ReindexContent(Pos);
		}
	}
};
CChangesDocumentRemoveItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_Document_AddItem === oChanges.Type || AscDFH.historyitem_Document_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesDocumentRemoveItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesDocumentAddItem);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesDocumentDefaultTab(Class, Old, New)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Old = Old;
	this.New = New;
}
CChangesDocumentDefaultTab.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesDocumentDefaultTab.prototype.constructor = CChangesDocumentDefaultTab;
CChangesDocumentDefaultTab.prototype.Type = AscDFH.historyitem_Document_DefaultTab;
CChangesDocumentDefaultTab.prototype.Undo = function()
{
	AscCommonWord.Default_Tab_Stop = this.Old;
};
CChangesDocumentDefaultTab.prototype.Redo = function()
{
	AscCommonWord.Default_Tab_Stop = this.New;
};
CChangesDocumentDefaultTab.prototype.WriteToBinary = function(Writer)
{
	// Double : New
	// Double : Old
	Writer.WriteDouble(this.New);
	Writer.WriteDouble(this.Old);
};
CChangesDocumentDefaultTab.prototype.ReadFromBinary = function(Reader)
{
	// Double : New
	// Double : Old
	this.New = Reader.GetDouble();
	this.Old = Reader.GetDouble();
};
CChangesDocumentDefaultTab.prototype.CreateReverseChange = function()
{
	return new CChangesDocumentDefaultTab(this.Class, this.New, this.Old);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesDocumentEvenAndOddHeaders(Class, Old, New)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Old = Old;
	this.New = New;
}
CChangesDocumentEvenAndOddHeaders.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesDocumentEvenAndOddHeaders.prototype.constructor = CChangesDocumentEvenAndOddHeaders;
CChangesDocumentEvenAndOddHeaders.prototype.Type = AscDFH.historyitem_Document_EvenAndOddHeaders;
CChangesDocumentEvenAndOddHeaders.prototype.Undo = function()
{
	EvenAndOddHeaders = this.Old;
};
CChangesDocumentEvenAndOddHeaders.prototype.Redo = function()
{
	EvenAndOddHeaders = this.New;
};
CChangesDocumentEvenAndOddHeaders.prototype.WriteToBinary = function(Writer)
{
	// Bool : New
	// Bool : Old
	Writer.WriteBool(this.New);
	Writer.WriteBool(this.Old);
};
CChangesDocumentEvenAndOddHeaders.prototype.ReadFromBinary = function(Reader)
{
	// Bool : New
	// Bool : Old
	this.New = Reader.GetBool();
	this.Old = Reader.GetBool();
};
CChangesDocumentEvenAndOddHeaders.prototype.CreateReverseChange = function()
{
	return new CChangesDocumentEvenAndOddHeaders(this.Class, this.New, this.Old);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesDocumentDefaultLanguage(Class, Old, New)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Old = Old;
	this.New = New;
}
CChangesDocumentDefaultLanguage.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesDocumentDefaultLanguage.prototype.constructor = CChangesDocumentDefaultLanguage;
CChangesDocumentDefaultLanguage.prototype.Type = AscDFH.historyitem_Document_DefaultLanguage;
CChangesDocumentDefaultLanguage.prototype.Undo = function()
{
	let oDocument = this.Class;
	oDocument.Styles.Default.TextPr.Lang.Val = this.Old;
	oDocument.RestartSpellCheck();
};
CChangesDocumentDefaultLanguage.prototype.Redo = function()
{
	let oDocument = this.Class;
	oDocument.Styles.Default.TextPr.Lang.Val = this.New;
	oDocument.RestartSpellCheck();
};
CChangesDocumentDefaultLanguage.prototype.WriteToBinary = function(Writer)
{
	// Long : New
	// Long : Old
	Writer.WriteLong(this.New);
	Writer.WriteLong(this.Old);
};
CChangesDocumentDefaultLanguage.prototype.ReadFromBinary = function(Reader)
{
	// Long : New
	// Long : Old
	this.New = Reader.GetLong();
	this.Old = Reader.GetLong();
};
CChangesDocumentDefaultLanguage.prototype.CreateReverseChange = function()
{
	return new CChangesDocumentDefaultLanguage(this.Class, this.New, this.Old);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesDocumentMathSettings(Class, Old, New)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Old = Old;
	this.New = New;
}
CChangesDocumentMathSettings.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesDocumentMathSettings.prototype.constructor = CChangesDocumentMathSettings;
CChangesDocumentMathSettings.prototype.Type = AscDFH.historyitem_Document_MathSettings;
CChangesDocumentMathSettings.prototype.Undo = function()
{
	var oDocument = this.Class;
	oDocument.Settings.MathSettings.SetPr(this.Old);
};
CChangesDocumentMathSettings.prototype.Redo = function()
{
	var oDocument = this.Class;
	oDocument.Settings.MathSettings.SetPr(this.New);
};
CChangesDocumentMathSettings.prototype.WriteToBinary = function(Writer)
{
	// Variable : New
	// Variable : Old
	this.New.Write_ToBinary(Writer);
	this.Old.Write_ToBinary(Writer);
};
CChangesDocumentMathSettings.prototype.ReadFromBinary = function(Reader)
{
	// Variable : New
	// Variable : Old
	this.New = new AscWord.MathSettings();
	this.New.Read_FromBinary(Reader);
	this.Old = new AscWord.MathSettings();
	this.Old.Read_FromBinary(Reader);
};
CChangesDocumentMathSettings.prototype.CreateReverseChange = function()
{
	return new CChangesDocumentMathSettings(this.Class, this.New, this.Old);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesDocumentSdtGlobalSettings(Class, Old, New)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class);

	this.Old = Old;
	this.New = New;
}
CChangesDocumentSdtGlobalSettings.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesDocumentSdtGlobalSettings.prototype.constructor = CChangesDocumentSdtGlobalSettings;
CChangesDocumentSdtGlobalSettings.prototype.Type = AscDFH.historyitem_Document_SdtGlobalSettings;
CChangesDocumentSdtGlobalSettings.prototype.private_SetValue = function(Value)
{
	this.Class.Settings.SdtSettings = Value;
	this.Class.OnChangeSdtGlobalSettings();
};
CChangesDocumentSdtGlobalSettings.prototype.private_CreateObject = function()
{
	return new AscWord.SdtGlobalSettings();
};
CChangesDocumentSdtGlobalSettings.prototype.private_IsCreateEmptyObject = function()
{
	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesDocumentSettingsGutterAtTop(Class, Old, New)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New);
}
CChangesDocumentSettingsGutterAtTop.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesDocumentSettingsGutterAtTop.prototype.constructor = CChangesDocumentSettingsGutterAtTop;
CChangesDocumentSettingsGutterAtTop.prototype.Type = AscDFH.historyitem_Document_Settings_GutterAtTop;
CChangesDocumentSettingsGutterAtTop.prototype.private_SetValue = function(Value)
{
	this.Class.Settings.GutterAtTop = Value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesDocumentSettingsMirrorMargins(Class, Old, New)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New);
}
CChangesDocumentSettingsMirrorMargins.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesDocumentSettingsMirrorMargins.prototype.constructor = CChangesDocumentSettingsMirrorMargins;
CChangesDocumentSettingsMirrorMargins.prototype.Type = AscDFH.historyitem_Document_Settings_MirrorMargins;
CChangesDocumentSettingsMirrorMargins.prototype.private_SetValue = function(Value)
{
	this.Class.Settings.MirrorMargins = Value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesDocumentSpecialFormsGlobalSettings(Class, Old, New)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class);

	this.Old = Old;
	this.New = New;
}
CChangesDocumentSpecialFormsGlobalSettings.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesDocumentSpecialFormsGlobalSettings.prototype.constructor = CChangesDocumentSpecialFormsGlobalSettings;
CChangesDocumentSpecialFormsGlobalSettings.prototype.Type = AscDFH.historyitem_Document_SpecialFormsGlobalSettings;
CChangesDocumentSpecialFormsGlobalSettings.prototype.private_SetValue = function(Value)
{
	this.Class.Settings.SpecialFormsSettings = Value;
	this.Class.OnChangeSpecialFormsGlobalSettings();
};
CChangesDocumentSpecialFormsGlobalSettings.prototype.private_CreateObject = function()
{
	return new AscWord.SpecialFormsGlobalSettings();
};
CChangesDocumentSpecialFormsGlobalSettings.prototype.private_IsCreateEmptyObject = function()
{
	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesDocumentSettingsTrackRevisions(Class, Old, New, sUserId)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Old    = Old;
	this.New    = New;
	this.UserId = sUserId;
}
CChangesDocumentSettingsTrackRevisions.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesDocumentSettingsTrackRevisions.prototype.constructor = CChangesDocumentSettingsTrackRevisions;
CChangesDocumentSettingsTrackRevisions.prototype.Type = AscDFH.historyitem_Document_Settings_TrackRevisions;
CChangesDocumentSettingsTrackRevisions.prototype.Undo = function()
{
	this.Class.Settings.TrackRevisions = this.Old;
	this.Class.private_OnTrackRevisionsChange();
};
CChangesDocumentSettingsTrackRevisions.prototype.Redo = function()
{
	this.Class.Settings.TrackRevisions = this.New;
	this.Class.private_OnTrackRevisionsChange();
};
CChangesDocumentSettingsTrackRevisions.prototype.Load = function()
{
	this.Class.Settings.TrackRevisions = this.New;
	this.Class.private_OnTrackRevisionsChange(this.UserId);
};
CChangesDocumentSettingsTrackRevisions.prototype.WriteToBinary = function(oWriter)
{
	// Long   : Flags
	// Bool   : New
	// Bool   : Old
	// String : UserId

	var nStartPos = oWriter.GetCurPosition();
	oWriter.Skip(4);
	var nFlags = 0;

	if (undefined !== this.Old)
	{
		oWriter.WriteBool(this.Old);
		nFlags |= 1;
	}

	if (undefined !== this.New)
	{
		oWriter.WriteBool(this.New);
		nFlags |= 2;
	}

	if (this.UserId)
	{
		oWriter.WriteString2(this.UserId);
		nFlags |= 4;
	}

	var nEndPos = oWriter.GetCurPosition();
	oWriter.Seek(nStartPos);
	oWriter.WriteLong(nFlags);
	oWriter.Seek(nEndPos);
};
CChangesDocumentSettingsTrackRevisions.prototype.ReadFromBinary = function(oReader)
{
	// Long   : Flags
	// Bool   : New
	// Bool   : Old
	// String : UserId

	var nFlags = oReader.GetLong();

	if (nFlags & 1)
		this.Old = oReader.GetBool();
	else
		this.Old = undefined;

	if (nFlags & 2)
		this.New = oReader.GetBool();
	else
		this.New = undefined;

	if (nFlags & 4)
		this.UserId = oReader.GetString2();
	else
		this.UserId = undefined;
};
CChangesDocumentSettingsTrackRevisions.prototype.CreateReverseChange = function()
{
	return new CChangesDocumentSettingsTrackRevisions(this.Class, this.New, this.Old, this.UserId);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesDocumentProtection(Class, Old, New, sUserId) {
	AscDFH.CChangesBase.call(this, Class, Old, New);
	if (Old && New) {
		this.OldAlgorithmName = Old.algorithmName;
		this.OldEdit = Old.edit;
		this.OldEnforcement = Old.enforcement;
		this.OldFormatting = Old.formatting;
		this.OldHashValue = Old.hashValue;
		this.OldSaltValue = Old.saltValue;
		this.OldSpinCount = Old.spinCount;
		this.OldAlgIdExt = Old.algIdExt;
		this.OldAlgIdExtSource = Old.algIdExtSource;
		this.OldCryptAlgorithmClass = Old.cryptAlgorithmClass;
		this.OldCryptAlgorithmSid = Old.cryptAlgorithmSid;
		this.OldCryptAlgorithmType = Old.cryptAlgorithmType;
		this.OldCryptProvider = Old.cryptProvider;
		this.OldCryptProviderType = Old.cryptProviderType;
		this.OldCryptProviderTypeExt = Old.cryptProviderTypeExt;
		this.OldCryptProviderTypeExtSource = Old.cryptProviderTypeExtSource;

		this.NewAlgorithmName = New.algorithmName === Old.algorithmName ? undefined : New.algorithmName;
		this.NewEdit = New.edit === Old.edit ? undefined : New.edit;
		this.NewEnforcement = New.enforcement === Old.enforcement ? undefined : New.enforcement;
		this.NewFormatting = New.formatting === Old.formatting ? undefined : New.formatting;
		this.NewHashValue = New.hashValue === Old.hashValue ? undefined : New.hashValue;
		this.NewSaltValue = New.saltValue === Old.saltValue ? undefined : New.saltValue;
		this.NewSpinCount = New.spinCount === Old.spinCount ? undefined : New.spinCount;
		this.NewAlgIdExt = New.algIdExt === Old.algIdExt ? undefined : New.algIdExt;
		this.NewAlgIdExtSource = New.algIdExtSource === Old.algIdExtSource ? undefined : New.algIdExtSource;
		this.NewCryptAlgorithmClass = New.cryptAlgorithmClass === Old.cryptAlgorithmClass ? undefined : New.cryptAlgorithmClass;
		this.NewCryptAlgorithmSid = New.cryptAlgorithmSid === Old.cryptAlgorithmSid ? undefined : New.cryptAlgorithmSid;
		this.NewCryptAlgorithmType = New.cryptAlgorithmType === Old.cryptAlgorithmType ? undefined : New.cryptAlgorithmType;
		this.NewCryptProvider = New.cryptProvider === Old.cryptProvider ? undefined : New.cryptProvider;
		this.NewCryptProviderType = New.cryptProviderType === Old.cryptProviderType ? undefined : New.cryptProviderType;
		this.NewCryptProviderTypeExt = New.cryptProviderTypeExt === Old.cryptProviderTypeExt ? undefined : New.cryptProviderTypeExt;
		this.NewCryptProviderTypeExtSource = New.cryptProviderTypeExtSource === Old.cryptProviderTypeExtSource ? undefined : New.cryptProviderTypeExtSource;
	} else {
		this.OldAlgorithmName = undefined;
		this.OldEdit = undefined;
		this.OldEnforcement = undefined;
		this.OldFormatting = undefined;
		this.OldHashValue = undefined;
		this.OldSaltValue = undefined;
		this.OldSpinCount = undefined;
		this.OldAlgIdExt = undefined;
		this.OldAlgIdExtSource = undefined;
		this.OldCryptAlgorithmClass = undefined;
		this.OldCryptAlgorithmSid = undefined;
		this.OldCryptAlgorithmType = undefined;
		this.OldCryptProvider = undefined;
		this.OldCryptProviderType = undefined;
		this.OldCryptProviderTypeExt = undefined;
		this.OldCryptProviderTypeExtSource = undefined;

		this.NewAlgorithmName = undefined;
		this.NewEdit = undefined;
		this.NewEnforcement = undefined;
		this.NewFormatting = undefined;
		this.NewHashValue = undefined;
		this.NewSaltValue = undefined;
		this.NewSpinCount = undefined;
		this.NewAlgIdExt = undefined;
		this.NewAlgIdExtSource = undefined;
		this.NewCryptAlgorithmClass = undefined;
		this.NewCryptAlgorithmSid = undefined;
		this.NewCryptAlgorithmType = undefined;
		this.NewCryptProvider = undefined;
		this.NewCryptProviderType = undefined;
		this.NewCryptProviderTypeExt = undefined;
		this.NewCryptProviderTypeExtSource = undefined;
	}
	this.UserId = sUserId;
}
CChangesDocumentProtection.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesDocumentProtection.prototype.constructor = CChangesDocumentProtection;
CChangesDocumentProtection.prototype.Type = AscDFH.historydescription_Document_DocumentProtection;
CChangesDocumentProtection.prototype.Undo = function () {
	if (!this.Class) {
		return;
	}

	this.Class.algorithmName = this.OldAlgorithmName;
	this.Class.edit = this.OldEdit;
	this.Class.enforcement = this.OldEnforcement;
	this.Class.formatting = this.OldFormatting;

	this.Class.hashValue = this.OldHashValue;
	this.Class.saltValue = this.OldSaltValue;
	this.Class.spinCount = this.OldSpinCount;
	this.Class.algIdExt = this.OldAlgIdExt;

	this.Class.algIdExtSource = this.OldAlgIdExtSource;
	this.Class.cryptAlgorithmClass = this.OldCryptAlgorithmClass;
	this.Class.cryptAlgorithmSid = this.OldCryptAlgorithmSid;
	this.Class.cryptAlgorithmType = this.OldCryptAlgorithmType;

	this.Class.cryptProvider = this.OldCryptProvider;
	this.Class.cryptProviderType = this.OldCryptProviderType;
	this.Class.cryptProviderTypeExt = this.OldCryptProviderTypeExt;
	this.Class.cryptProviderTypeExtSource = this.OldCryptProviderTypeExtSource;

	var api = Asc.editor || editor;
	if (api) {
		api.asc_OnProtectionUpdate();
	}
};
CChangesDocumentProtection.prototype.Redo = function (sUserId, isLoadChanges) {
	if (!this.Class) {
		return;
	}

	this.Class.algorithmName = this.NewAlgorithmName;
	this.Class.edit = this.NewEdit;
	this.Class.enforcement = this.NewEnforcement;
	this.Class.formatting = this.NewFormatting;
	this.Class.hashValue = this.NewHashValue;
	this.Class.saltValue = this.NewSaltValue;
	this.Class.spinCount = this.NewSpinCount;
	this.Class.formatting = this.NewFormatting;
	this.Class.algIdExt = this.NewAlgIdExt;

	this.Class.algIdExtSource = this.NewAlgIdExtSource;
	this.Class.cryptAlgorithmClass = this.NewCryptAlgorithmClass;
	this.Class.cryptAlgorithmSid = this.NewCryptAlgorithmSid;
	this.Class.cryptAlgorithmType = this.NewCryptAlgorithmType;
	this.Class.cryptProvider = this.NewCryptProvider;
	this.Class.cryptProviderType = this.NewCryptProviderType;
	this.Class.cryptProviderTypeExt = this.NewCryptProviderTypeExt;
	this.Class.cryptProviderTypeExtSource = this.NewCryptProviderTypeExtSource;

	var api = Asc.editor || editor;
	if (api) {
		let oDocument = api.private_GetLogicDocument();
		if (oDocument && oDocument.Settings) {
			var docProtection = oDocument.Settings && oDocument.Settings.DocumentProtection;
			if (!docProtection || this.Class !== docProtection) {
				oDocument.Settings.DocumentProtection = this.Class;
			}
		}

		if (!isLoadChanges) {
			api.asc_OnProtectionUpdate(sUserId);
		} else {
			if (oDocument && oDocument.Settings) {
				var _docProtection = oDocument.Settings && oDocument.Settings.DocumentProtection;
				if (_docProtection) {
					_docProtection.SetNeedUpdate(sUserId);
				}
			}
		}
	}
};
CChangesDocumentProtection.prototype.Load = function () {
	this.Redo(this.UserId, true);
};
CChangesDocumentProtection.prototype.WriteToBinary = function (Writer) {
	if (null != this.NewAlgorithmName) {
		Writer.WriteBool(true);
		Writer.WriteByte(this.NewAlgorithmName);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewEdit) {
		Writer.WriteBool(true);
		Writer.WriteByte(this.NewEdit);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewEnforcement) {
		Writer.WriteBool(true);
		Writer.WriteBool(this.NewEnforcement);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewFormatting) {
		Writer.WriteBool(true);
		Writer.WriteBool(this.NewFormatting);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewHashValue) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.NewHashValue);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewSaltValue) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.NewSaltValue);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewSpinCount) {
		Writer.WriteBool(true);
		Writer.WriteLong(this.NewSpinCount);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewAlgIdExt) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.NewAlgIdExt);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewAlgIdExt) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.NewAlgIdExt);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewAlgIdExtSource) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.NewAlgIdExtSource);
	} else {
		Writer.WriteBool(false);
	}

	if (null != this.NewCryptAlgorithmClass) {
		Writer.WriteBool(true);
		Writer.WriteByte(this.NewCryptAlgorithmClass);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewCryptAlgorithmSid) {
		Writer.WriteBool(true);
		Writer.WriteLong(this.NewCryptAlgorithmSid);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewCryptAlgorithmType) {
		Writer.WriteBool(true);
		Writer.WriteByte(this.NewCryptAlgorithmType);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewCryptProvider) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.NewCryptProvider);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewCryptProviderType) {
		Writer.WriteBool(true);
		Writer.WriteByte(this.NewCryptProviderType);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewCryptProviderTypeExt) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.NewCryptProviderTypeExt);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.NewCryptProviderTypeExtSource) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.NewCryptProviderTypeExtSource);
	} else {
		Writer.WriteBool(false);
	}
	if (null != this.UserId) {
		Writer.WriteBool(true);
		Writer.WriteString2(this.UserId);
	} else {
		Writer.WriteBool(false);
	}
};

CChangesDocumentProtection.prototype.ReadFromBinary = function (Reader) {
	if (Reader.GetBool()) {
		this.NewAlgorithmName = Reader.GetByte();
	}
	if (Reader.GetBool()) {
		this.NewEdit = Reader.GetByte();
	}
	if (Reader.GetBool()) {
		this.NewEnforcement = Reader.GetBool();
	}
	if (Reader.GetBool()) {
		this.NewFormatting = Reader.GetBool();
	}
	if (Reader.GetBool()) {
		this.NewHashValue = Reader.GetString2();
	}
	if (Reader.GetBool()) {
		this.NewSaltValue = Reader.GetString2();
	}
	if (Reader.GetBool()) {
		this.NewSpinCount = Reader.GetLong();
	}
	if (Reader.GetBool()) {
		this.NewAlgIdExt = Reader.GetString2();
	}
	if (Reader.GetBool()) {
		this.NewAlgIdExt = Reader.GetString2();
	}
	if (Reader.GetBool()) {
		this.NewAlgIdExtSource = Reader.GetString2();
	}
	if (Reader.GetBool()) {
		this.NewCryptAlgorithmClass = Reader.GetByte();
	}
	if (Reader.GetBool()) {
		this.NewCryptAlgorithmSid = Reader.GetLong();
	}
	if (Reader.GetBool()) {
		this.NewCryptAlgorithmType = Reader.GetByte();
	}
	if (Reader.GetBool()) {
		this.NewCryptProvider = Reader.GetString2();
	}
	if (Reader.GetBool()) {
		this.NewCryptProviderType = Reader.GetByte();
	}
	if (Reader.GetBool()) {
		this.NewCryptProviderTypeExt = Reader.GetString2();
	}
	if (Reader.GetBool()) {
		this.NewCryptProviderTypeExtSource = Reader.GetString2();
	}
	if (Reader.GetBool()) {
		this.UserId = Reader.GetString2();
	}
};
CChangesDocumentProtection.prototype.CreateReverseChange = function () {
	var ret = new CChangesDocumentProtection(this.Class);
	
	ret.OldAlgorithmName = this.NewAlgorithmName;
	ret.OldEdit = this.NewEdit;
	ret.OldEnforcement = this.NewEnforcement;
	ret.OldFormatting = this.NewFormatting;
	ret.OldHashValue = this.NewHashValue;
	ret.OldSaltValue = this.NewSaltValue;
	ret.OldSpinCount = this.NewSpinCount;
	ret.OldAlgIdExt = this.NewAlgIdExt;
	ret.OldAlgIdExtSource = this.NewAlgIdExtSource;
	ret.OldCryptAlgorithmClass = this.NewCryptAlgorithmClass;
	ret.OldCryptAlgorithmSid = this.NewCryptAlgorithmSid;
	ret.OldCryptAlgorithmType = this.NewCryptAlgorithmType;
	ret.OldCryptProvider = this.NewCryptProvider;
	ret.OldCryptProviderType = this.NewCryptProviderType;
	ret.OldCryptProviderTypeExt = this.NewCryptProviderTypeExt;
	ret.OldCryptProviderTypeExtSource = this.NewCryptProviderTypeExtSource;

	ret.NewAlgorithmName = this.OldAlgorithmName;
	ret.NewEdit = this.OldEdit;
	ret.NewEnforcement = this.OldEnforcement;
	ret.NewFormatting = this.OldFormatting;
	ret.NewHashValue = this.OldHashValue;
	ret.NewSaltValue = this.OldSaltValue;
	ret.NewSpinCount = this.OldSpinCount;
	ret.NewAlgIdExt = this.OldAlgIdExt;
	ret.NewAlgIdExtSource = this.OldAlgIdExtSource;
	ret.NewCryptAlgorithmClass = this.OldCryptAlgorithmClass;
	ret.NewCryptAlgorithmSid = this.OldCryptAlgorithmSid;
	ret.NewCryptAlgorithmType = this.OldCryptAlgorithmType;
	ret.NewCryptProvider = this.OldCryptProvider;
	ret.NewCryptProviderType = this.OldCryptProviderType;
	ret.NewCryptProviderTypeExt = this.OldCryptProviderTypeExt;
	ret.NewCryptProviderTypeExtSource = this.OldCryptProviderTypeExtSource;

	ret.UserId = this.UserId;
	
	return ret;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesDocumentSettingsAutoHyphenation(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesDocumentSettingsAutoHyphenation.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesDocumentSettingsAutoHyphenation.prototype.constructor = CChangesDocumentSettingsAutoHyphenation;
CChangesDocumentSettingsAutoHyphenation.prototype.Type = AscDFH.historyitem_Document_Settings_AutoHyphenation;
CChangesDocumentSettingsAutoHyphenation.prototype.private_SetValue = function(value)
{
	this.Class.Settings.autoHyphenation = value;
	this.Class.OnChangeAutoHyphenation();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesDocumentSettingsConsecutiveHyphenLimit(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesDocumentSettingsConsecutiveHyphenLimit.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesDocumentSettingsConsecutiveHyphenLimit.prototype.constructor = CChangesDocumentSettingsConsecutiveHyphenLimit;
CChangesDocumentSettingsConsecutiveHyphenLimit.prototype.Type = AscDFH.historyitem_Document_Settings_ConsecutiveHyphenLimit;
CChangesDocumentSettingsConsecutiveHyphenLimit.prototype.private_SetValue = function(value)
{
	this.Class.Settings.consecutiveHyphenLimit = value;
	this.Class.OnChangeAutoHyphenation();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesDocumentSettingsDoNotHyphenateCaps(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesDocumentSettingsDoNotHyphenateCaps.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesDocumentSettingsDoNotHyphenateCaps.prototype.constructor = CChangesDocumentSettingsDoNotHyphenateCaps;
CChangesDocumentSettingsDoNotHyphenateCaps.prototype.Type = AscDFH.historyitem_Document_Settings_DoNotHyphenateCaps;
CChangesDocumentSettingsDoNotHyphenateCaps.prototype.private_SetValue = function(value)
{
	this.Class.Settings.doNotHyphenateCaps = value;
	this.Class.OnChangeAutoHyphenation();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesDocumentSettingsHyphenationZone(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesDocumentSettingsHyphenationZone.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesDocumentSettingsHyphenationZone.prototype.constructor = CChangesDocumentSettingsHyphenationZone;
CChangesDocumentSettingsHyphenationZone.prototype.Type = AscDFH.historyitem_Document_Settings_HyphenationZone;
CChangesDocumentSettingsHyphenationZone.prototype.private_SetValue = function(value)
{
	this.Class.Settings.hyphenationZone = value;
	this.Class.OnChangeAutoHyphenation();
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesDocumentPageColor(Class, Old, New)
{
	AscDFH.CChangesBase.call(this, Class);
	
	this.Old = Old;
	this.New = New;
}
CChangesDocumentPageColor.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesDocumentPageColor.prototype.constructor = CChangesDocumentPageColor;
CChangesDocumentPageColor.prototype.Type = AscDFH.historyitem_Document_PageColor;
CChangesDocumentPageColor.prototype.Undo = function()
{
	this.Class.Background = this.Old;
};
CChangesDocumentPageColor.prototype.Redo = function()
{
	this.Class.Background = this.New;
};
CChangesDocumentPageColor.prototype.WriteToBinary = function(writer)
{
	this.New.writeToBinary(writer);
	this.Old.writeToBinary(writer);
};
CChangesDocumentPageColor.prototype.ReadFromBinary = function(reader)
{
	this.New = new AscWord.DocumentBackground();
	this.Old = new AscWord.DocumentBackground();
	
	this.New.readFromBinary(reader);
	this.Old.readFromBinary(reader);
};
CChangesDocumentPageColor.prototype.CreateReverseChange = function()
{
	return new CChangesDocumentPageColor(this.Class, this.New, this.Old);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesDocumentDisconnectEveryone(Class)
{
	AscDFH.CChangesBase.call(this, Class);
}
CChangesDocumentDisconnectEveryone.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesDocumentDisconnectEveryone.prototype.constructor = CChangesDocumentDisconnectEveryone;
CChangesDocumentDisconnectEveryone.prototype.Type = AscDFH.historyitem_Document_DisconnectEveryone;
CChangesDocumentDisconnectEveryone.prototype.Undo = function()
{
};
CChangesDocumentDisconnectEveryone.prototype.Redo = function()
{
	let logicDocument = this.Class;
	
	logicDocument.sendEvent("asc_onDisconnectEveryone");
	let editorApi = logicDocument.GetApi();
	
	editorApi.setViewModeDisconnect(true);
	editorApi.asc_coAuthoringDisconnect();
};
AscDFH.CChangesDocumentDisconnectEveryone = CChangesDocumentDisconnectEveryone;
