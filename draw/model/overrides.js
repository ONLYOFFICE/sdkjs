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

// Import
var CShape = AscFormat.CShape;
var CGroupShape = AscFormat.CGroupShape;
var CTheme = AscFormat.CTheme;
var CreateSolidFillRGBA = AscFormat.CreateSolidFillRGBA;

/**
 * @memberOf CShape
 * @return {{layout: null, slide: null, theme: CTheme, master: null}}
 */
CShape.prototype.getParentObjects = function ()
{
	let oTheme = null;
	if (this.parent) {
		oTheme = this.parent.theme;
	} else {
		console.log("Parent was not set for shape/group. GenerateDefaultTheme is used. shape/group:", this);
		oTheme = AscFormat.GenerateDefaultTheme(null, null);
	}
	return {slide: null, layout: null, master: null, theme: oTheme};
};

/**
 * @memberOf CGroupShape
 * @type {function(): {layout: null, slide: null, theme: CTheme, master: null}}
 */
CGroupShape.prototype.getParentObjects = CShape.prototype.getParentObjects;


/**
 * Draw editor.
 * @memberof CShape
 */
CShape.prototype.recalculate = function ()
{
	if(this.bDeleted || !this.parent) {
		console.log("no recalculate for bDeleted or no parent");
		return;
	}

	if(this.parent.getObjectType() === AscDFH.historyitem_type_Notes){
		return;
	}

	// var check_slide_placeholder = !this.isPlaceholder() || (this.parent && (this.parent.getObjectType() === AscDFH.historyitem_type_Slide));
	let check_placeholder = !this.isPlaceholder() || (this.parent && this.parent.constructor.name === "CVisioDocument");
	AscFormat.ExecuteNoHistory(function(){

		var bRecalcShadow = this.recalcInfo.recalculateBrush ||
			this.recalcInfo.recalculatePen ||
			this.recalcInfo.recalculateTransform ||
			this.recalcInfo.recalculateGeometry ||
			this.recalcInfo.recalculateBounds;
		if (this.recalcInfo.recalculateBrush) {
			this.recalculateBrush();
			this.recalcInfo.recalculateBrush = false;
		}

		if (this.recalcInfo.recalculatePen) {
			this.recalculatePen();
			this.recalcInfo.recalculatePen = false;
		}
		if (this.recalcInfo.recalculateTransform) {
			this.recalculateTransform();
			this.recalculateSnapArrays();
			this.recalcInfo.recalculateTransform = false;
		}

		if (this.recalcInfo.recalculateGeometry) {
			this.recalculateGeometry();
			this.recalcInfo.recalculateGeometry = false;
		}

		if (this.recalcInfo.recalculateContent && check_placeholder) {
			this.recalcInfo.oContentMetrics = this.recalculateContent();
			this.recalcInfo.recalculateContent = false;
		}
		if (this.recalcInfo.recalculateContent2 && check_placeholder) {
			this.recalculateContent2();
			this.recalcInfo.recalculateContent2 = false;
		}

		if (this.recalcInfo.recalculateTransformText && check_placeholder) {
			this.recalculateTransformText();
			this.recalcInfo.recalculateTransformText = false;
		}
		if(this.recalcInfo.recalculateBounds)
		{
			this.recalculateBounds();
			this.recalcInfo.recalculateBounds = false;
		}
		if(bRecalcShadow)
		{
			this.recalculateShdw();
		}

		this.clearCropObject();
	}, this, []);
};

/**
 *
 * @param idx
 * @param unicolor
 * @param {Boolean} isConnectorShape
 * @memberOf CTheme
 * @return {CUniFill|*}
 */
CTheme.prototype.getFillStyle = function (idx, unicolor, isConnectorShape) {
	if (idx === 0 || idx === 1000) {
		return AscFormat.CreateNoFillUniFill();
	}
	var ret;
	let fmtScheme = isConnectorShape ?
		this.themeElements.themeExt.fmtConnectorScheme :
		this.themeElements.fmtScheme;
	if (idx >= 1 && idx <= 999) {
		if (fmtScheme.fillStyleLst[idx - 1]) {
			ret = fmtScheme.fillStyleLst[idx - 1].createDuplicate();
			if (ret) {
				ret.checkPhColor(unicolor, false);
				return ret;
			}
		}
	} else if (idx >= 1001) {
		if (fmtScheme.bgFillStyleLst[idx - 1001]) {
			ret = fmtScheme.bgFillStyleLst[idx - 1001].createDuplicate();
			if (ret) {
				ret.checkPhColor(unicolor, false);
				return ret;
			}
		}
	}
	console.log("getFillStyle has not found fill and returned transparent fill")
	return CreateSolidFillRGBA(0, 0, 0, 255);
};

/**
 *
 * @param idx
 * @param unicolor
 * @param {Boolean} isConnectorShape
 * @memberOf CTheme
 * @return {CLn|*}
 */
CTheme.prototype.getLnStyle = function (idx, unicolor, isConnectorShape) {
	if (idx === 0) {
		return AscFormat.CreateNoFillLine();
	}
	let fmtScheme = isConnectorShape ?
		this.themeElements.themeExt.fmtConnectorScheme :
		this.themeElements.fmtScheme;
	if (fmtScheme.lnStyleLst[idx - 1]) {
		var ret = fmtScheme.lnStyleLst[idx - 1].createDuplicate();
		if (ret.Fill) {
			ret.Fill.checkPhColor(unicolor, false);
		}
		return ret;
	}
	console.log("getLnStyle has not found lineStyle and returned new CLn()");
	return new AscFormat.CLn();
};
