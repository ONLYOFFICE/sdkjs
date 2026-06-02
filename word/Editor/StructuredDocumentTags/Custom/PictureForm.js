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
	/**
	 * Класс с настройками для формы с изображением
	 * @constructor
	 */
	function CSdtPictureFormPr()
	{
		this.ScaleFlag   = Asc.c_oAscPictureFormScaleFlag.Always;
		this.Proportions = true;
		this.Borders     = false;
		this.ShiftX      = 0.5; // 0..1
		this.ShiftY      = 0.5; // 0..1
		this.Signature   = false;
	}
	CSdtPictureFormPr.prototype.Copy = function()
	{
		var oFormPr = new CSdtPictureFormPr();

		oFormPr.ScaleFlag   = this.ScaleFlag;
		oFormPr.Proportions = this.Proportions;
		oFormPr.Borders     = this.Borders;
		oFormPr.ShiftX      = this.ShiftX;
		oFormPr.ShiftY      = this.ShiftY;
		oFormPr.Signature   = this.Signature;

		return oFormPr;
	};
	CSdtPictureFormPr.prototype.IsEqual = function(oOther)
	{
		return (oOther
			&& this.ScaleFlag === oOther.ScaleFlag
			&& this.Proportions === oOther.Proportions
			&& this.Borders === oOther.Borders
			&& Math.abs(this.ShiftX - oOther.ShiftX) < 0.001
			&& Math.abs(this.ShiftY - oOther.ShiftY) < 0.001
			&& this.Signature === oOther.Signature);
	};
	CSdtPictureFormPr.prototype.WriteToBinary = function(oWriter)
	{
		oWriter.WriteLong(this.ScaleFlag);
		oWriter.WriteBool(this.Proportions);
		oWriter.WriteBool(this.Borders);
		oWriter.WriteDouble(this.ShiftX);
		oWriter.WriteDouble(this.ShiftY);
		oWriter.WriteBool(this.Signature);
	};
	CSdtPictureFormPr.prototype.ReadFromBinary = function(oReader)
	{
		this.ScaleFlag   = oReader.GetLong();
		this.Proportions = oReader.GetBool();
		this.Borders     = oReader.GetBool();
		this.ShiftX      = oReader.GetDouble();
		this.ShiftY      = oReader.GetDouble();
		this.Signature   = oReader.GetBool();
	};
	CSdtPictureFormPr.prototype.Write_ToBinary = function(oWriter)
	{
		this.WriteToBinary(oWriter);
	};
	CSdtPictureFormPr.prototype.Read_FromBinary = function(oReader)
	{
		this.ReadFromBinary(oReader);
	};
	CSdtPictureFormPr.prototype.GetScaleFlag = function()
	{
		return this.ScaleFlag;
	};
	CSdtPictureFormPr.prototype.SetScaleFlag = function(nFlag)
	{
		this.ScaleFlag = nFlag;
	};
	CSdtPictureFormPr.prototype.IsConstantProportions = function()
	{
		return this.Proportions;
	};
	CSdtPictureFormPr.prototype.SetConstantProportions = function(isConstant)
	{
		this.Proportions = isConstant;
	};
	CSdtPictureFormPr.prototype.IsRespectBorders = function()
	{
		return this.Borders;
	};
	CSdtPictureFormPr.prototype.SetRespectBorders = function(isRespect)
	{
		this.Borders = isRespect;
	};
	CSdtPictureFormPr.prototype.SetShiftX = function(nX)
	{
		this.ShiftX = nX;
	};
	CSdtPictureFormPr.prototype.GetShiftX = function()
	{
		return this.ShiftX;
	};
	CSdtPictureFormPr.prototype.SetShiftY = function(nY)
	{
		this.ShiftY = nY;
	};
	CSdtPictureFormPr.prototype.GetShiftY = function()
	{
		return this.ShiftY;
	};
	CSdtPictureFormPr.prototype.IsSignature = function()
	{
		return this.Signature;
	};
	CSdtPictureFormPr.prototype.SetSignature = function(isSignature)
	{
		this.Signature = isSignature;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CSdtPictureFormPr = CSdtPictureFormPr;

	window['AscCommon'].CSdtPictureFormPr    = CSdtPictureFormPr;
	window['AscCommon']['CSdtPictureFormPr'] = CSdtPictureFormPr;

	CSdtPictureFormPr.prototype['get_ScaleFlag']           = CSdtPictureFormPr.prototype.GetScaleFlag;
	CSdtPictureFormPr.prototype['put_ScaleFlag']           = CSdtPictureFormPr.prototype.SetScaleFlag;
	CSdtPictureFormPr.prototype['get_ConstantProportions'] = CSdtPictureFormPr.prototype.IsConstantProportions;
	CSdtPictureFormPr.prototype['put_ConstantProportions'] = CSdtPictureFormPr.prototype.SetConstantProportions;
	CSdtPictureFormPr.prototype['get_RespectBorders']      = CSdtPictureFormPr.prototype.IsRespectBorders;
	CSdtPictureFormPr.prototype['put_RespectBorders']      = CSdtPictureFormPr.prototype.SetRespectBorders;
	CSdtPictureFormPr.prototype['put_ShiftX']              = CSdtPictureFormPr.prototype.SetShiftX;
	CSdtPictureFormPr.prototype['get_ShiftX']              = CSdtPictureFormPr.prototype.GetShiftX;
	CSdtPictureFormPr.prototype['put_ShiftY']              = CSdtPictureFormPr.prototype.SetShiftY;
	CSdtPictureFormPr.prototype['get_ShiftY']              = CSdtPictureFormPr.prototype.GetShiftY;

})(window);
