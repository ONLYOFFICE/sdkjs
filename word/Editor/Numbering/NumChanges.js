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

AscDFH.changesFactory[AscDFH.historyitem_Num_LvlOverrideChange] = CChangesNumLvlOverrideChange;
AscDFH.changesFactory[AscDFH.historyitem_Num_AbstractNum]       = CChangesNumAbstractNum;

//----------------------------------------------------------------------------------------------------------------------
// Карта зависимости изменений
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_Num_LvlOverrideChange] = [
	AscDFH.historyitem_Num_LvlOverrideChange
];

AscDFH.changesRelationMap[AscDFH.historyitem_Num_AbstractNum] = [
	AscDFH.historyitem_Num_AbstractNum
];
//----------------------------------------------------------------------------------------------------------------------

/**
 *
 * @param Class {CNum}
 * @param Old {?CLvlOverride}
 * @param New {?CLvlOverride}
 * @param nLvl {number} 0..8
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesNumLvlOverrideChange(Class, Old, New, nLvl)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New);

	this.Lvl = nLvl;
}
CChangesNumLvlOverrideChange.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesNumLvlOverrideChange.prototype.constructor = CChangesNumLvlOverrideChange;
CChangesNumLvlOverrideChange.prototype.Type = AscDFH.historyitem_Num_LvlOverrideChange;
CChangesNumLvlOverrideChange.prototype.WriteToBinary = function(oWriter)
{
	// Long         : Lvl index
	// Long         : Flags
	// CLvlOverride : New
	// CLvlOverride : Old

	oWriter.WriteLong(this.Lvl);

	var nFlags = 0;

	if (undefined === this.New)
		nFlags |= 1;

	if (undefined === this.Old)
		nFlags |= 2;

	oWriter.WriteLong(nFlags);

	if (undefined !== this.New && this.New.WriteToBinary)
		this.New.WriteToBinary(oWriter);

	if (undefined !== this.Old && this.Old.WriteToBinary)
		this.Old.WriteToBinary(oWriter);
};
CChangesNumLvlOverrideChange.prototype.ReadFromBinary = function(oReader)
{
	// Long         : Lvl index
	// Long         : Flags
	// CLvlOverride : New
	// CLvlOverride : Old

	this.Lvl = oReader.GetLong();

	var nFlags = oReader.GetLong();

	if (nFlags & 1)
	{
		this.New = undefined;
	}
	else
	{
		this.New = new CLvlOverride();
		this.New.ReadFromBinary(oReader);
	}

	if (nFlags & 2)
	{
		this.Old = undefined;
	}
	else
	{
		this.Old = new CLvlOverride();
		this.Old.ReadFromBinary(oReader);
	}
};
CChangesNumLvlOverrideChange.prototype.private_SetValue = function(Value)
{
	var oNum = this.Class;
	oNum.LvlOverride[this.Lvl] = Value;
	oNum.RecalculateRelatedParagraphs(this.Lvl);
};
CChangesNumLvlOverrideChange.prototype.Load = function(Color)
{
	var oNum = this.Class;
	oNum.LvlOverride[this.Lvl] = this.New;

	// Сразу нельзя запускать пересчет, т.к. возможно еще не все ссылки проставлены
	AscCommon.CollaborativeEditing.Add_EndActions(this.Class, {Lvl : this.Lvl});
};
CChangesNumLvlOverrideChange.prototype.CreateReverseChange = function()
{
	return new CChangesNumLvlOverrideChange(this.Class, this.New, this.Old, this.Lvl);
};
CChangesNumLvlOverrideChange.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type)
		return false;

	return true;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesNumAbstractNum(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesNumAbstractNum.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesNumAbstractNum.prototype.constructor = CChangesNumAbstractNum;
CChangesNumAbstractNum.prototype.Type = AscDFH.historyitem_Num_AbstractNum;
CChangesNumAbstractNum.prototype.private_SetValue = function(Value)
{
	this.Class.AbstractNumId = Value;
};
