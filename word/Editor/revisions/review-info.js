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

(function()
{
	/**
	 * @param {undefined | reviewtype_Add | reviewtype_Common | reviewtype_Remove} reviewType
	 * @constructor
	 */
	function ReviewInfo(reviewType)
	{
		this.Editor = editor;
		
		this.Type     = undefined !== reviewType && null !== reviewType ? reviewType : reviewtype_Common;
		this.UserId   = "";
		this.UserName = "";
		this.DateTime = "";
		
		this.MoveType = Asc.c_oAscRevisionsMove.NoMove;
		
		this.PrevType = -1;
		this.PrevInfo = null;
	}
	
	ReviewInfo.createAdd = function()
	{
		let info  = new ReviewInfo();
		info.Type = reviewtype_Add;
		return info;
	};
	ReviewInfo.createRemove = function()
	{
		let info  = new ReviewInfo();
		info.Type = reviewtype_Remove;
		return info;
	};
	ReviewInfo.fromBinary = function(reader)
	{
		let info = new ReviewInfo();
		info.ReadFromBinary(reader);
		return info;
	};
	
	
	ReviewInfo.prototype.setType = function(type)
	{
		this.Type = type;
	};
	ReviewInfo.prototype.getType = function()
	{
		return this.Type;
	};
	ReviewInfo.prototype.Update  = function()
	{
		if (this.Editor && this.Editor.DocInfo)
		{
			this.UserId   = this.Editor.DocInfo.get_UserId();
			this.UserName = this.Editor.DocInfo.get_UserName();
			this.DateTime = (new Date()).getTime();
		}
	};
	ReviewInfo.prototype.Copy    = function()
	{
		var Info      = new ReviewInfo();
		Info.Type     = this.Type;
		Info.UserId   = this.UserId;
		Info.UserName = this.UserName;
		Info.DateTime = this.DateTime;
		Info.MoveType = this.MoveType;
		Info.PrevType = this.PrevType;
		Info.PrevInfo = this.PrevInfo ? this.PrevInfo.Copy() : null;
		return Info;
	};
	/**
	 * Получаем имя пользователя
	 * @returns {string}
	 */
	ReviewInfo.prototype.GetUserName = function()
	{
		return this.UserName;
	};
	/**
	 * Получаем дату-время изменения
	 * @returns {number}
	 */
	ReviewInfo.prototype.GetDateTime = function()
	{
		return this.DateTime;
	};
	ReviewInfo.prototype.Write_ToBinary  = function(oWriter)
	{
		oWriter.WriteLong(this.Type);
		oWriter.WriteString2(this.UserId);
		oWriter.WriteString2(this.UserName);
		oWriter.WriteString2(this.DateTime);
		oWriter.WriteLong(this.MoveType);
		
		if (-1 !== this.PrevType && null !== this.PrevInfo)
		{
			oWriter.WriteBool(true);
			oWriter.WriteLong(this.PrevType);
			this.PrevInfo.Write_ToBinary(oWriter);
		}
		else
		{
			oWriter.WriteBool(false);
		}
	};
	ReviewInfo.prototype.Read_FromBinary = function(oReader)
	{
		this.Type     = oReader.GetLong();
		this.UserId   = oReader.GetString2();
		this.UserName = oReader.GetString2();
		this.DateTime = parseInt(oReader.GetString2());
		this.MoveType = oReader.GetLong();
		
		if (oReader.GetBool())
		{
			this.PrevType = oReader.GetLong();
			this.PrevInfo = new ReviewInfo();
			this.PrevInfo.Read_FromBinary(oReader);
		}
		else
		{
			this.PrevType = -1;
			this.PrevInfo = null;
		}
	};
	ReviewInfo.prototype.toBinary = function(writer)
	{
		this.Write_ToBinary(writer);
	};
	ReviewInfo.prototype.Get_Color       = function()
	{
		if (!this.UserId && !this.UserName)
			return REVIEW_COLOR;
		
		return AscCommon.getUserColorById(this.UserId, this.UserName, true, false);
	};
	ReviewInfo.prototype.IsCurrentUser   = function()
	{
		if (this.Editor && this.Editor.DocInfo)
		{
			var UserId = this.Editor.DocInfo.get_UserId();
			return (UserId === this.UserId);
		}
		
		return true;
	};
	/**
	 * Получаем идентификатор пользователя
	 * @returns {string}
	 */
	ReviewInfo.prototype.GetUserId = function()
	{
		return this.UserId;
	};
	ReviewInfo.prototype.WriteToBinary  = function(oWriter)
	{
		this.Write_ToBinary(oWriter);
	};
	ReviewInfo.prototype.ReadFromBinary = function(oReader)
	{
		this.Read_FromBinary(oReader);
	};
	/**
	 * Сохраняем предыдущее действие (обычно это добавление, а новое - удаление)
	 * @param {number} nType
	 */
	ReviewInfo.prototype.SavePrev = function(nType)
	{
		this.PrevType = nType;
		this.PrevInfo = this.Copy();
	};
	ReviewInfo.prototype.SetPrevReviewTypeWithInfoRecursively = function(nType, oInfo)
	{
		var last = this;
		while (last.PrevInfo)
		{
			last = last.PrevInfo;
		}
		last.PrevType = nType;
		last.PrevInfo = oInfo;
	};
	/**
	 * Данная функция запрашивает было ли ранее произведено добавление
	 * @returns {?ReviewInfo}
	 */
	ReviewInfo.prototype.GetPrevAdded = function()
	{
		var nPrevType = this.PrevType;
		var oPrevInfo = this.PrevInfo;
		while (oPrevInfo)
		{
			if (reviewtype_Add === this.PrevType)
			{
				return oPrevInfo;
			}
			
			nPrevType = oPrevInfo.PrevType;
			oPrevInfo = oPrevInfo.PrevInfo;
		}
		
		return null;
	};
	/**
	 * Данная функция запрашивает было ли ранее произведено добавление текущим пользователем
	 * @returns {?ReviewInfo}
	 */
	ReviewInfo.prototype.IsPrevAddedByCurrentUser = function()
	{
		var oPrevInfo = this.GetPrevAdded();
		if (!oPrevInfo)
			return false;
		
		return oPrevInfo.IsCurrentUser();
	};
	ReviewInfo.prototype.GetColor = function()
	{
		return this.Get_Color();
	};
	/**
	 * Выставляем тип переноса
	 * @param {Asc.c_oAscRevisionsMove} nType
	 */
	ReviewInfo.prototype.SetMove = function(nType)
	{
		this.MoveType = nType;
	};
	/**
	 * Добавленный текст во время переноса?
	 * @returns {boolean}
	 */
	ReviewInfo.prototype.IsMovedTo = function()
	{
		return this.MoveType === Asc.c_oAscRevisionsMove.MoveTo;
	};
	/**
	 * Удаленный текст во время переноса?
	 * @returns {boolean}
	 */
	ReviewInfo.prototype.IsMovedFrom = function()
	{
		return this.MoveType === Asc.c_oAscRevisionsMove.MoveFrom;
	};
	/**
	 * Сравнение информации об изменениях
	 * @returns {boolean}
	 */
	ReviewInfo.prototype.IsEqual = function(oAnotherReviewInfo, bIsMergingDocuments)
	{
		let oThisReviewInfo    = this;
		let oCompareReviewInfo = oAnotherReviewInfo;
		let bEquals            = true;
		while (bEquals && oThisReviewInfo && oCompareReviewInfo)
		{
			bEquals = oThisReviewInfo.Type === oCompareReviewInfo.Type &&
				oThisReviewInfo.UserName === oCompareReviewInfo.UserName &&
				oThisReviewInfo.DateTime === oCompareReviewInfo.DateTime &&
				oThisReviewInfo.MoveType === oCompareReviewInfo.MoveType &&
				oThisReviewInfo.PrevType === oCompareReviewInfo.PrevType;
			
			if (!bIsMergingDocuments)
			{
				bEquals = bEquals && oThisReviewInfo.UserId === oCompareReviewInfo.UserId;
			}
			
			oThisReviewInfo    = oThisReviewInfo.PrevInfo;
			oCompareReviewInfo = oCompareReviewInfo.PrevInfo;
		}
		if (!oThisReviewInfo && oCompareReviewInfo || oThisReviewInfo && !oCompareReviewInfo)
		{
			return false;
		}
		return bEquals;
	};
	/**
	 * @returns {Asc.c_oAscRevisionsMove}
	 */
	ReviewInfo.prototype.GetMoveType = function()
	{
		return this.MoveType;
	};
	ReviewInfo.prototype.SetMoveType = function(type)
	{
		this.MoveType = type;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ReviewInfo = ReviewInfo;
	
})(window);
