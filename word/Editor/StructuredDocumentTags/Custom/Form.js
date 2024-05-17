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
	/**
	 * Класс для общих настроек у формы
	 * @param sKey
	 * @param sLabel
	 * @param sHelpText
	 * @param isRequired
	 * @constructor
	 */
	function CSdtFormPr(sKey, sLabel, sHelpText, isRequired)
	{
		this.Key      = sKey;
		this.Label    = sLabel;
		this.HelpText = sHelpText;
		this.Required = isRequired;
		this.Fixed    = false;
		this.Border   = undefined;
		this.Shd      = undefined;
		this.Field    = undefined;
		this.RoleName = undefined; // Данное поле только для работы с интерфейсом, в бинарник его не пишем
	}
	CSdtFormPr.prototype.Copy = function()
	{
		var oFormPr = new CSdtFormPr();

		oFormPr.Key      = this.Key;
		oFormPr.Label    = this.Label;
		oFormPr.HelpText = this.HelpText;
		oFormPr.Required = this.Required;

		if (this.Border)
			oFormPr.Border = this.Border.Copy();

		if (this.Shd)
			oFormPr.Shd = this.Shd.Copy();
		
		// При простом копировании настроек нельзя делать копию fieldMaster, т.к. копирование делается вне какого-либо
		// действия. Поэтом, просто копируем здесь ссылку. Реальное копирование объекта должно происходить в момент
		// его вставки куда-либо
		if (this.Field)
			oFormPr.Field = this.Field;
		
		oFormPr.RoleName = this.RoleName;

		return oFormPr;
	};
	CSdtFormPr.prototype.IsEqual = function(oOther)
	{
		return (oOther
			&& this.Key === oOther.Key
			&& this.Label === oOther.Label
			&& this.HelpText === oOther.HelpText
			&& this.Required === oOther.Required
			&& IsEqualStyleObjects(this.Border, oOther.Border)
			&& IsEqualStyleObjects(this.Shd, oOther.Shd)
			&& this.Field === oOther.Field
			&& this.RoleName === oOther.RoleName);
	};
	CSdtFormPr.prototype.WriteToBinary = function(oWriter)
	{
		var nStartPos = oWriter.GetCurPosition();
		oWriter.Skip(4);
		var nFlags = 0;

		if (undefined !== this.Key)
		{
			oWriter.WriteString2(this.Key);
			nFlags |= 1;
		}

		if (undefined !== this.Label)
		{
			oWriter.WriteString2(this.Label);
			nFlags |= 2;
		}

		if (undefined !== this.HelpText)
		{
			oWriter.WriteString2(this.HelpText);
			nFlags |= 4;
		}

		if (undefined !== this.Required)
		{
			oWriter.WriteBool(this.Required);
			nFlags |= 8;
		}

		if (undefined !== this.Border)
		{
			this.Border.WriteToBinary(oWriter);
			nFlags |= 16;
		}

		if (undefined !== this.Shd)
		{
			this.Shd.WriteToBinary(oWriter)
			nFlags |= 32;
		}
		
		if (AscCommon.IsSupportOFormFeature() && this.Field)
		{
			oWriter.WriteString2(this.Field.GetId());
			nFlags |= 64;
		}

		var nEndPos = oWriter.GetCurPosition();
		oWriter.Seek(nStartPos);
		oWriter.WriteLong(nFlags);
		oWriter.Seek(nEndPos);
	};
	CSdtFormPr.prototype.ReadFromBinary = function(oReader)
	{
		var nFlags = oReader.GetLong();

		if (nFlags & 1)
			this.Key = oReader.GetString2();

		if (nFlags & 2)
			this.Label = oReader.GetString2();

		if (nFlags & 4)
			this.HelpText = oReader.GetString2();

		if (nFlags & 8)
			this.Required = oReader.GetBool();

		if (nFlags & 16)
		{
			this.Border = new CDocumentBorder();
			this.Border.ReadFromBinary(oReader);
		}

		if (nFlags & 32)
		{
			this.Shd = new CDocumentShd();
			this.Shd.ReadFromBinary(oReader);
		}
		
		if (AscCommon.IsSupportOFormFeature() && (nFlags & 64))
		{
			let fieldId = oReader.GetString2();
			this.Field = AscCommon.g_oTableId.GetById(fieldId);
		}
	};
	CSdtFormPr.prototype.Write_ToBinary = function(oWriter)
	{
		this.WriteToBinary(oWriter);
	};
	CSdtFormPr.prototype.Read_FromBinary = function(oReader)
	{
		this.ReadFromBinary(oReader);
	};
	CSdtFormPr.prototype.GetKey = function()
	{
		return this.Key;
	};
	CSdtFormPr.prototype.SetKey = function(sKey)
	{
		this.Key = sKey;
	};
	CSdtFormPr.prototype.GetLabel = function()
	{
		return this.Label;
	};
	CSdtFormPr.prototype.SetLabel = function(sLabel)
	{
		this.Label = sLabel;
	};
	CSdtFormPr.prototype.GetHelpText = function()
	{
		return this.HelpText;
	};
	CSdtFormPr.prototype.SetHelpText = function(sText)
	{
		this.HelpText = sText;
	};
	CSdtFormPr.prototype.GetRequired = function()
	{
		return this.Required;
	};
	CSdtFormPr.prototype.SetRequired = function(isRequired)
	{
		this.Required = isRequired;
	};
	CSdtFormPr.prototype.GetFixed = function()
	{
		return this.Fixed;
	};
	CSdtFormPr.prototype.SetFixed = function(isFixed)
	{
		this.Fixed = isFixed;
	};
	CSdtFormPr.prototype.GetBorder = function()
	{
		return this.Border;
	};
	CSdtFormPr.prototype.GetAscBorder = function()
	{
		if (!this.Border)
			return undefined;

		return (new Asc.asc_CTextBorder(this.Border));
	};
	CSdtFormPr.prototype.SetAscBorder = function(oAscBorder)
	{
		if (!oAscBorder)
		{
			this.Border = undefined;
		}
		else
		{
			this.Border = new CDocumentBorder();
			this.Border.Set_FromObject(oAscBorder);
		}
	};
	CSdtFormPr.prototype.GetShd = function()
	{
		return this.Shd;
	};
	CSdtFormPr.prototype.SetShd = function(shd)
	{
		if (!shd)
			this.Shd = undefined;
		else
			this.Shd = AscWord.CShd.FromObject(shd);
	};
	CSdtFormPr.prototype.GetAscShd = function()
	{
		if (!this.Shd)
			return undefined;

		return (new Asc.asc_CParagraphShd(this.Shd));
	};
	CSdtFormPr.prototype.SetAscShd = function(isShd, oAscColor)
	{
		if (!isShd || !oAscColor)
		{
			this.Shd = undefined;
		}
		else
		{
			var oUnifill        = new AscFormat.CUniFill();
			oUnifill.fill       = new AscFormat.CSolidFill();
			oUnifill.fill.color = AscFormat.CorrectUniColor(oAscColor, oUnifill.fill.color, 1);

			var oLogicDocument = editor.WordControl.m_oLogicDocument;
			if (oLogicDocument && oLogicDocument.IsDocumentEditor())
				oUnifill.check(oLogicDocument.GetTheme(), oLogicDocument.GetColorMap());

			this.Shd = new CDocumentShd();
			this.Shd.Set_FromObject({
				Value: Asc.c_oAscShd.Clear,
				Color: {
					r: oAscColor.asc_getR(),
					g: oAscColor.asc_getG(),
					b: oAscColor.asc_getB(),
					Auto: false
				},
				Fill: {
					r: oAscColor.asc_getR(),
					g: oAscColor.asc_getG(),
					b: oAscColor.asc_getB(),
					Auto: false
				},
				Unifill: oUnifill
			});
		}
	};
	CSdtFormPr.prototype.SetFieldMaster = function(fieldMaster)
	{
		if (!AscCommon.IsSupportOFormFeature())
			return;

		this.Field = null !== fieldMaster ? fieldMaster : undefined;
	};
	CSdtFormPr.prototype.GetFieldMaster = function()
	{
		if (!AscCommon.IsSupportOFormFeature())
			return null;

		return this.Field ? this.Field : null;
	};
	CSdtFormPr.prototype.GetAscRole = function()
	{
		// Приоритет у роли выставленной через интерфейс, т.к. выставление в класс формы еще могло не произойти
		// но название роли мы должны уже отдавать новое
		if (this.RoleName)
			return this.RoleName;
		
		if (this.Field && this.Field.getUserCount())
			return this.Field.getFirstUser().getRole();
		
		return AscCommon.translateManager.getValue("Anyone");
	};
	CSdtFormPr.prototype.SetAscRole = function(roleName)
	{
		return this.SetRole(roleName);
	};
	CSdtFormPr.prototype.GetRole = function()
	{
		if (!AscCommon.IsSupportOFormFeature())
			return undefined;
		
		return this.RoleName;
	};
	CSdtFormPr.prototype.SetRole = function(roleName)
	{
		if (!AscCommon.IsSupportOFormFeature())
			return;
		
		this.RoleName = null !== roleName ? roleName : undefined;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CSdtFormPr = CSdtFormPr;

	//---------------------------------------------interface export-----------------------------------------------------
	window['AscCommon'].CSdtFormPr    = CSdtFormPr;
	window['AscCommon']['CSdtFormPr'] = CSdtFormPr;
	
	CSdtFormPr.prototype['get_Key']      = CSdtFormPr.prototype.GetKey;
	CSdtFormPr.prototype['put_Key']      = CSdtFormPr.prototype.SetKey;
	CSdtFormPr.prototype['get_Label']    = CSdtFormPr.prototype.GetLabel;
	CSdtFormPr.prototype['put_Label']    = CSdtFormPr.prototype.SetLabel;
	CSdtFormPr.prototype['get_HelpText'] = CSdtFormPr.prototype.GetHelpText;
	CSdtFormPr.prototype['put_HelpText'] = CSdtFormPr.prototype.SetHelpText;
	CSdtFormPr.prototype['get_Required'] = CSdtFormPr.prototype.GetRequired;
	CSdtFormPr.prototype['put_Required'] = CSdtFormPr.prototype.SetRequired;
	CSdtFormPr.prototype['get_Fixed']    = CSdtFormPr.prototype.GetFixed;
	CSdtFormPr.prototype['put_Fixed']    = CSdtFormPr.prototype.SetFixed;
	CSdtFormPr.prototype['get_Border']   = CSdtFormPr.prototype.GetAscBorder;
	CSdtFormPr.prototype['put_Border']   = CSdtFormPr.prototype.SetAscBorder;
	CSdtFormPr.prototype['get_Shd']      = CSdtFormPr.prototype.GetAscShd;
	CSdtFormPr.prototype['put_Shd']      = CSdtFormPr.prototype.SetAscShd;
	CSdtFormPr.prototype['get_Role']     = CSdtFormPr.prototype.GetAscRole;
	CSdtFormPr.prototype['put_Role']     = CSdtFormPr.prototype.SetAscRole;
})(window);
