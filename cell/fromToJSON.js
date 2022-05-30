/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
(function(window, undefined){

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

	function private_PtToMM(pt)
	{
		return 25.4 / 72.0 * pt;
	}
	function private_Twips2MM(twips)
	{
		return 25.4 / 72.0 / 20 * twips;
	}
	function private_GetDrawingDocument()
	{
		return editor.WordControl.m_oLogicDocument.DrawingDocument;
	}
	function private_EMU2MM(EMU)
	{
		return EMU / 36000.0;
	}
	function private_MM2EMU(MM)
	{
		return MM * 36000.0;
	}
	function private_GetLogicDocument()
	{
		return editor.WordControl.m_oLogicDocument;
	}
	function private_MM2Twips(mm)
	{
		return mm / (25.4 / 72.0 / 20);
	}
	/**
	 * Get the first Run in the array specified.
	 * @typeofeditors ["CDE"]
	 * @param {Array} firstPos - first doc pos of element
	 * @param {Array} secondPos - second doc pos of element
	 * @return {1 || 0 || - 1}
	 * If returns 1  -> first element placed before second
	 * If returns 0  -> first element placed like second
	 * If returns -1 -> first element placed after second
	 */
	function private_checkRelativePos(firstPos, secondPos)
	{
		for (var nPos = 0, nLen = Math.min(firstPos.length, secondPos.length); nPos < nLen; ++nPos)
		{
			if (!secondPos[nPos] || !firstPos[nPos] || firstPos[nPos].Class !== secondPos[nPos].Class)
				return 1;

			if (firstPos[nPos].Position < secondPos[nPos].Position)
				return 1;
			else if (firstPos[nPos].Position > secondPos[nPos].Position)
				return -1;
		}

		return 0;
	}
	function private_MM2Pt(mm)
	{
		return mm / (25.4 / 72.0);
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// End of private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    var WriterToJSON   = window['AscCommon'].WriterToJSON;
	var ReaderFromJSON = window['AscCommon'].ReaderFromJSON;
	
	WriterToJSON.prototype.SerWorksheet = function(oWorksheet)
	{
		var aCols = [];
		for (var nCol = 0; nCol < oWorksheet.aCols.length; nCol++)
			aCols.push(this.SerCol(oWorksheet.aCols[nCol]));

		var aDrawings = [];
		for (var nDrawing = 0; nDrawing < oWorksheet.Drawings.length; nDrawing++)
			aDrawings.push(this.SerDrawingExel(oWorksheet.Drawings[nDrawing]));

		var aHyperlinks = [];
		var aWorksheetLinks = oWorksheet.worksheet.hyperlinkManager.getAll();
		// for (var nHyperlink = 0; nHyperlink < aWorksheetLinks.length; nHyperlink++)
		// 	aHyperlinks.push(this.SerHyperlinkExel(aWorksheetLinks[nHyperlink]));

		return {
			autoFilter:            this.SerAutoFilter(oWorksheet.AutoFilter),
			cols:                  aCols,
			conditionalFormatting: this.SerCondFormatting(),
			dataValidations:       this.SerDataValidations(oWorksheet.dataValidations),
			drawings:              aDrawings,      
			headerFooter:          null, /// всегда лежит объект CHeaderFooterData
			hiperlinks:            aHyperlinks /// to do

		}
	};
	WriterToJSON.prototype.SerDrawingExel = function(oDrawingExel)
	{
		return {
			pos: {
				x: oDrawingExel.Pos.X,
				y: oDrawingExel.Pos.Y
			},

			ext:     oDrawingExel.ext,
			from:    this.SerCellObjectInfo(oDrawingExel.from),
			to:      this.SerCellObjectInfo(oDrawingExel.to),
			graphic: this.SerGraphicObject(oDrawingExel.graphicObject),
			editAs:  ToXML_ST_EditAs(oDrawingExel.editAs)
		}
	};
	WriterToJSON.prototype.SerCellObjectInfo = function(oCellObjectInfo)
	{
		return {
			col:    oCellObjectInfo.col,
			colOff: private_MM2EMU(oCellObjectInfo.colOff),
			row:    oCellObjectInfo.row,
			rowOff: private_MM2EMU(oCellObjectInfo.rowOff)
		}
	};
	WriterToJSON.prototype.Ser = function(o)
	{

	};
	WriterToJSON.prototype.Ser = function(o)
	{

	};
	WriterToJSON.prototype.SerDataValidations = function(oDataValidations)
	{
		var aElems = [];
		for (var nElem = 0; nElem < oDataValidations.elems.length; nElem++)
			aElems.push(this.SerDataValitation(oDataValidations.elems[nElem]));

		return {
			dataValidation: aElems,
			disablePrompts: oDataValidations.disablePrompts,
			xWindow:        oDataValidations.xWindow,
			yWindow:        oDataValidations.yWindow
		}
	};
	WriterToJSON.prototype.SerDataValitation = function(oDataValidation)
	{
		return {
			formula1: this.SerFormula(oDataValidation.formula1),
			formula2: this.SerFormula(oDataValidation.formula2),

			allowBlank:       oDataValidation.allowBlank,
			error:            oDataValidation.error,
			errorStyle:       AscCommonExcel.ToXml_ST_DataValidationErrorStyle(oDataValidation.errorStyle),
			errorTitle:       oDataValidation.errorTitle,
			imeMode:          AscCommonExcel.ToXml_ST_DataValidationImeMode(oDataValidation.imeMode),
			operator:         AscCommonExcel.ToXml_ST_DataValidationOperator(oDataValidation.operator),
			prompt:           oDataValidation.prompt,
			promptTitle:      oDataValidation.promptTitle,
			showDropDown:     oDataValidation.showDropDown,
			showErrorMessage: oDataValidation.showErrorMessage,
			showInputMessage: oDataValidation.showInputMessage,
			sqref:            getSqRefString(oDataValidation.ranges),
			type:             AscCommonExcel.ToXml_ST_DataValidationType(oDataValidation.type)
		}
	};
	WriterToJSON.prototype.SerFormula = function(oFormula)
	{
		return {
			formula: oFormula.text,
			type:    "formula"
		}
	};
	WriterToJSON.prototype.SerCondFormatting = function(oCondFormatting)
	{
		var aCondRules = [];
		for (var nRule = 0; nRule < oCondFormatting.aConditionalFormattingRules.length; nRule++)
			aCondRules.push(this.SerConditionalRule(oCondFormatting.aConditionalFormattingRules[nRule]));

		return {
			cfRule: aCondRules,
			pivot:  oCondFormatting.pivot,
			sqref:  getSqRefString(oCondFormatting.ranges)
		}
	};
	WriterToJSON.prototype.SerConditionalRule = function(oCondRule)
	{
		var aRules = [];
		for (var nRule = 0; nRule < oCondRule.aRuleElements.length; nRule++)
			aRules.push(this.SerRuleElement(oCondRule.aRuleElements[nRule]));

		return {
			rules:        aRules,
			aboveAverage: oCondRule.aboveAverage,
			bottom:       oCondRule.bottom,
			dxfId:        oCondRule.dxf,
			equalAverage: oCondRule.equalAverage,
			operator:     AscCommonExcel.ToXml_CFOperatorType(oCondRule.operator),
			percent:      oCondRule.percent,
			priority:     oCondRule.priority,
			rank:         oCondRule.rank,
			stdDev:       oCondRule.stdDev,
			stopIfTrue:   oCondRule.stopIfTrue,
			text:         oCondRule.text,
			timePeriod:   AscCommonExcel.ToXml_ST_TimePeriod(oCondRule.timePeriod),
			type:         AscCommonExcel.ToXml_ST_CfType(oCondRule.type),
		}
	};
	WriterToJSON.prototype.SerRuleElement = function(oRule)
	{
		if (oRule instanceof AscCommonExcel.CColorScale)
			return this.SerColorScale(oRule);
		if (oRule instanceof AscCommonExcel.CDataBar)
			return this.SerDataBar(oRule);
		if (oRule instanceof AscCommonExcel.CIconSet)
			return this.SerIconSet(oRule);
		if (oRule instanceof AscCommonExcel.CFormulaCF)
			return this.SerFormulaCF(oRule);
	};
	WriterToJSON.prototype.SerFormulaCF = function(oFormulaCf)
	{
		return {
			formula: oFormulaCf.Text,
			type:    "formulaCf"
		}
	};
	WriterToJSON.prototype.SerIconSet = function(oIconSet)
	{
		var aCFVO = [];
		for (var nElem = 0; nElem < oIconSet.aCFVOs.length; nElem++)
			aCFVO.push(this.SerCondFmtValObj(oIconSet.aCFVOs[nElem]));

		var aCFIS = [];
		for (nElem = 0; nElem < oIconSet.aIconSets.length; nElem++)
			aCFIS.push(this.SerCondFmtIconSet(oIconSet.aIconSets[nElem]));

		return {
			cfvo:      aCFVO,
			cfIcon:    aCFIS,
			iconSet:   AscCommonExcel.ToXml_ST_IconSetType(oIconSet.IconSet),
			percent:   oIconSet.Percent,
			reverse:   oIconSet.Reverse,
			showValue: oIconSet.ShowValue,
			type:      "iconSet"
		}
	};
	WriterToJSON.prototype.SerCondFmtIconSet = function(oCFIS)
	{
		return {
			iconSet: AscCommonExcel.ToXml_ST_IconSetType(oCFIS.IconSet),
			iconId:  oCFIS.IconId
		}
	};
	WriterToJSON.prototype.SerDataBar = function(oDataBar)
	{
		var aCFVO = [];
		for (var nElem = 0; nElem < oDataBar.aCFVOs.length; nElem++)
			aCFVO.push(this.SerCondFmtValObj(oDataBar.aCFVOs[nElem]));

		var sAxisPos = undefined;
		switch (oDataBar.AxisPosition)
		{
			case AscCommonExcel.EDataBarAxisPosition.automatic:
				sAxisPos = "auto";
				break;
			case AscCommonExcel.EDataBarAxisPosition.middle:
				sAxisPos = "middle";
				break;
			case AscCommonExcel.EDataBarAxisPosition.none:
				sAxisPos = "none";
				break;
		}

		var sDir = undefined;
		switch (oDataBar.Direction)
		{
			case AscCommonExcel.EDataBarDirection.context:
				sDir = "context";
				break;
			case AscCommonExcel.EDataBarDirection.leftToRight:
				sDir = "leftToRight";
				break;
			case AscCommonExcel.EDataBarDirection.rightToLeft:
				sDir = "rightToLeft";
				break;
		}

		return {
			cfvo:                aCFVO,
			color:               this.SerColorExel(oDataBar.Color),
			negativeColor:       this.SerColorExel(oDataBar.NegativeColor),
			borderColor:         this.SerColorExel(oDataBar.BorderColor),
			axisColor:           this.SerColorExel(oDataBar.AxisColor),
			negativeBorderColor: this.SerColorExel(oDataBar.NegativeBorderColor),
			
			maxLength: oDataBar.MaxLength,
			minLength: oDataBar.MinLength,
			showValue: oDataBar.ShowValue,
			axPos:     sAxisPos,
			dir:       sDir,
			gradient:  oDataBar.Gradient,
			negBarClrSameAsPositive:    oDataBar.NegativeBarColorSameAsPositive,
			negBarBdrClrSameAsPositive: oDataBar.NegativeBarBorderColorSameAsPositive,
			type:                       "dataBar"
		}
	};
	WriterToJSON.prototype.SerColorScale = function(oColorScale)
	{
		var aCFVO = [];
		for (var nElem = 0; nElem < oColorScale.aCFVOs.length; nElem++)
			aCFVO.push(this.SerCondFmtValObj(oColorScale.aCFVOs[nElem]));

		var aColors = [];
		for (var nColor = 0; nColor < oColorScale.aColors.length; nColor++)
			aColors.push(this.SerColorExel(oColorScale.aColors[nColor]));

		return {
			cfvo:  aCFVO,
			color: aColors,
			type:  "clrScale"
		}
	};
	WriterToJSON.prototype.SerCondFmtValObj = function(oCondFmtValObj)
	{
		return {
			gte:  oCondFmtValObj.Gte,
			type: AscCommonExcel.ToXml_ST_CfvoType(oCondFmtValObj.Type),
			val:  oCondFmtValObj.Val
		}
	};
	WriterToJSON.prototype.SerColorExel = function(oColor)
	{
		var res;
		if (oColor instanceof AscCommonExcel.ThemeColor)
			res = {
				rgb:   oColor.rgb,
				theme: oColor.theme,
				tint:  oColor.tint,
				type:  "themeClr"

			}
		else if (oColor instanceof AscCommonExcel.RgbColor)
			res = {
				rgb:   oColor.rgb,
				type:  "rgbClr"
			}

		return res;
	};
	WriterToJSON.prototype.SerCol = function(oCol)
	{
		return {
			bestFit:      oCol.BestFit,
			collapsed:    oCol.collapsed,
			customWidth:  oCol.CustomWidth,
			hidden:       oCol.hd,
			outlineLevel: oCol.outlineLevel,
			width:        oCol.width,
			xfs:          this.SerXFS(oCol.xfs)
		}
	};
	WriterToJSON.prototype.SerXFS = function(oXFS)
	{
		return {
			alignment: this.SerAlign(oXFS.align),
			pretection: {
				hidden: oXFS.hidden,
				locked: oXFS.locked
			},

			applyAlignment:  oXFS.ApplyAlignment,
			applyBorder:     oXFS.ApplyBorder,
			applyFill:       oXFS.ApplyFill,
			applyFont:       oXFS.ApplyFont,
			ApplyFont:       oXFS.ApplyNumberFormat,
			applyProtection: oXFS.applyProtection,
			borderId:        oXFS.borderid,
			fillId:          oXFS.fillid,
			fontId:          oXFS.fontid,
			numFmtId:        oXFS.numid,
			pivotButton:     oXFS.PivotButton,
			quotePrefix:     oXFS.QuotePrefix,
			xfId:            oXFS.XfId,
		}
	};
	WriterToJSON.prototype.SerAlign = function(oAlign)
	{
		return {
			horizontal:     AscCommonExcel.ToXml_ST_HorizontalAlignment(oAlign.hor),
			indent:         oAlign.indent,
			relativeIndent: oAlign.RelativeIndent,
			shrinkToFit:    oAlign.shrink,
			textRotation:   oAlign.angle,
			vertical:       AscCommonExcel.ToXml_ST_VerticalAlignment(oAlign.ver),
			wrapText:       oAlign.wrap
		}
	};
	WriterToJSON.prototype.SerAutoFilter = function(oAutoFilter)
	{
		var aFilterColumns = [];
		if (oAutoFilter.FilterColumns)
		{
			for (var nFilterCol = 0; nFilterCol < oAutoFilter.FilterColumns.length; nFilterCol)
				aFilterColumns.push(this.SerFilterColumn(oAutoFilter.FilterColumns[nFilterCol]));
		}
		
		return {
			filterColumn: aFilterColumns,
			sortState:    this.SerSortState(oAutoFilter.SortState),
			ref:          this.SerRef(oAutoFilter.Ref)
		}
	};
	WriterToJSON.prototype.SerFilterColumn = function(oFilterColumn)
	{
		if (!oFilterColumn)
			return oFilterColumn;

		return {
			colorFilter:   this.SerColorFilter(oFilterColumn.ColorFilter),
			customFilters: this.SerCustomFilters(oFilterColumn.CustomFiltersObj),
			dynamicFilter: this.SerDynamicFilter(oFilterColumn.DynamicFilter),
			filters:       this.SerFilters(oFilterColumn.Filters),
			top10:         this.SerTop10(oFilterColumn.Top10),
			colId:         oFilterColumn.ColId,
			showButton:    oFilterColumn.ShowButton
		}
	};
	WriterToJSON.prototype.SerColorFilter = function(oColorFilter)
	{
		if (!oColorFilter)
			return oColorFilter;

		return {
			cellColor: oColorFilter.CellColor,
			dxfId:     dxf
		}
	};
	WriterToJSON.prototype.SerCustomFilters = function(oCustomFilters)
	{
		if (!oCustomFilters)
			return oCustomFilters;

		var aFilters = [];
		if (oCustomFilters.CustomFilters)
		{
			for (var nFilter = 0; nFilter < oCustomFilters.CustomFilters.length; nFilter++)
				aFilters.push(this.SerCustomFilter(oCustomFilters.CustomFilters[nFilter]));
		}

		return {
			and:           oCustomFilters.And,
			customFilters: aFilters
		}
	};
	WriterToJSON.prototype.SerCustomFilter = function(oCustomFilter)
	{
		return {
			operator: AscCommonExcel.ToXml_ST_FilterOperator(oCustomFilter.Operator),
			val:      oCustomFilter.val
		}
	};
	WriterToJSON.prototype.SerDynamicFilter = function(oDynamicFilter)
	{
		if (!oDynamicFilter)
			return oDynamicFilter;

		return {
			type:      AscCommonExcel.ToXml_ST_DynamicFilterType(oDynamicFilter.Type),
			maxValIso: oDynamicFilter.MaxVal,
			valIso:    oDynamicFilter.Val
		}
	};
	WriterToJSON.prototype.SerFilters = function(oFilters)
	{
		if (!oFilters)
			return oFilters;

		var aDateGroupItem = [];
		for (var nDate = 0; nDate < oFilters.Dates.length; nDate++)
			aDateGroupItem.push(this.SerFilterDate(oFilters.Dates[nDate]));
			
		var aFilter = [];
		for (var nFilter = 0; nFilter < oFilters.Values.length; nFilter++)
			aFilter.push(oFilters.Values[nFilter]);

		return {
			dateGroupItem: aDateGroupItem,
			filter:        aFilter,
			blank:         oFilters.blank
		}
	};
	WriterToJSON.prototype.SerFilterDate = function(oFilterDate)
	{
		return {
			start:            oFilterDate.start,
			end:              oFilterDate.end,
			dateTimeGrouping: AscCommonExcel.ToXml_ST_DateTimeGrouping(oFilterDate.dateTimeGrouping)
		}
	};
	WriterToJSON.prototype.SerTop10 = function(oTop10)
	{
		if (!oTop10)
			return oTop10;

		return {
			filterVal: oTop10.FilterVal,
			percent:   oTop10.Percent,
			top:       oTop10.Top,
			val:       oTop10.Val
		}
	};
	WriterToJSON.prototype.SerSortState = function(oSortState)
	{
		if (!oSortState)
			return oSortState;

		var aSortConditions= [];
		if (oSortState.SortConditions)
		{
			for (var nSort = 0; nSort < oSortState.SortConditions.length; nSort++)
				aSortConditions.push(this.SerSortCondition(oSortState.SortConditions[nSort]));
		}

		return {
			sortCondition: aSortConditions,
			caseSensitive: oSortState.CaseSensitive,
			columnSort:    oSortState.ColumnSort,
			ref:           this.SerRef(oSortState.Ref),
			sortMethod:    oSortState.SortMethod /// ???
		}
	};
	WriterToJSON.prototype.SerSortCondition = function(oSortCondition)
	{
		var sSortBy = undefined;
		switch (oSortCondition.ConditionSortBy)
		{
			case Asc.ESortBy.sortbyCellColor:
				sSortBy = "cellColor";
				break;
			case Asc.ESortBy.sortbyFontColor:
				sSortBy = "fontColor";
				break;
			case Asc.ESortBy.sortbyIcon:
				sSortBy = "icon";
				break;
			case Asc.ESortBy.sortbyValue:
				sSortBy = "value";
				break;
		}
		return {
			sortBy: sSortBy,
			descending: oSortCondition.ConditionDescending,
			ref:        this.SerRef(oSortCondition.Ref)
		}
	};
	WriterToJSON.prototype.SerRef = function(oRef)
	{
		if (!oRef)
			return oRef;

		var sColumn1 = AscCommon.g_oCellAddressUtils.colnumToColstr(oRef.c1);
		var sColumn2 = AscCommon.g_oCellAddressUtils.colnumToColstr(oRef.c2);

		return sColumn1 + oRef.r1 + ":" + sColumn2 + oRef.r2;
	};
	WriterToJSON.prototype.Ser = function(o)
	{

	};
	WriterToJSON.prototype.Ser = function(o)
	{

	};
	WriterToJSON.prototype.Ser = function(o)
	{

	};

	function getSqRefString(ranges)
	{
		var refs = [];
		for (var i = 0; i < ranges.length; ++i) {
			refs.push(ranges[i].getName());
		}
		return refs.join(' ');
	}

	function ToXML_ST_EditAs(nType)
	{
		var sType = undefined;
		switch (nType)
		{
			case AscCommon.c_oAscCellAnchorType.cellanchorAbsolute:
				sType = "absolute";
				break;
			case AscCommon.c_oAscCellAnchorType.cellanchorOneCell:
				sType = "oneCell";
				break;
			case AscCommon.c_oAscCellAnchorType.cellanchorTwoCell:
				sType = "twoCell";
				break;
		}

		return sType;
	}
	function FromXML_ST_EditAse(sType)
	{
		var nType = undefined;
		switch (sType)
		{
			case "absolute":
				nType = AscCommon.c_oAscCellAnchorType.cellanchorAbsolute;
				break;
			case "oneCell":
				nType = AscCommon.c_oAscCellAnchorType.cellanchorOneCell;
				break;
			case "twoCell":
				nType = AscCommon.c_oAscCellAnchorType.cellanchorTwoCell;
				break;
		}

		return nType;
	}
	
    //----------------------------------------------------------export----------------------------------------------------
    window['AscCommon']       = window['AscCommon'] || {};
    window['AscFormat']       = window['AscFormat'] || {};
	
})(window);



