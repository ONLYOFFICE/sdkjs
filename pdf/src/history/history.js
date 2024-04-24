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

(function (window) {
	
	/**
	 * Класс локальной истории изменений
	 * @param logicDocument
	 * @constructor
	 */
	function History(logicDocument)
	{
		AscCommon.CHistory.apply(this, arguments);
	}
	History.prototype = Object.create(AscCommon.CHistory.prototype);
	History.prototype.constructor = History;
	
	History.prototype.ClearAdditional = function() {
		// Ничего не делаем
	};
	History.prototype.GetLastPointSourceObjectsPdf = function() {
		if (this.Index !== -1) {
			return this.Points[this.Index].Additional.Pdf;
		}
	};
	History.prototype.SetSourceObjectsToPointPdf = function(aObj) {
		if (this.Index !== -1) {
			this.Points[this.Index].Additional.Pdf = aObj;
		}
	};
	History.prototype.GetPdfConvertTextPoint = function() {
		if (this.Index !== -1) {
			return !!this.Points[this.Index].Additional.PdfConvertText;
		}
	};
	History.prototype.SetPdfConvertTextPoint = function(oConvertInfo) {
		if (this.Index !== -1) {
			this.Points[this.Index].Additional.PdfConvertText = oConvertInfo;
		}
	};
	History.prototype.Add = function(_Class, Data) {
		if (!this.CanAddChanges())
			return;

		AscCommon.CHistory.prototype.Add.call(this, _Class, Data);
		if (_Class.Class && _Class.Class.SetNeedRecalc) {
			if (!this.Points[this.Index].Additional.Pdf) {
				this.Points[this.Index].Additional.Pdf = [];
			}
			
			if (false == this.Points[this.Index].Additional.Pdf.includes(_Class.Class)) {
				this.Points[this.Index].Additional.Pdf.push(_Class.Class);
			}
		}
	}	
	
	//----------------------------------------------------------export--------------------------------------------------
	window['AscPDF'] = window['AscPDF'] || {};
	window['AscPDF'].History = History;
})(window);
