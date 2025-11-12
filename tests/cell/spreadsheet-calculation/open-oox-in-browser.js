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

$( function () {

	var g_oIdCounter = AscCommon.g_oIdCounter;
	var oParser, wb, ws, dif = 1e-9, sData = AscCommon.getEmpty(), tmp;
	if ( AscCommon.c_oSerFormat.Signature === sData.substring( 0, AscCommon.c_oSerFormat.Signature.length ) ) {
		wb = new AscCommonExcel.Workbook( new AscCommonExcel.asc_CHandlersList(), {wb:{getWorksheet:function(){}}}, true);

		AscCommon.g_oTableId.init();
		if ( this.User )
			g_oIdCounter.Set_UserId(this.User.asc_getId());

		AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wb);
		AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wb);
		AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wb);
		AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wb, false);
		AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wb, true);
		AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wb);
		AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wb);
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wb, -1);
		g_oIdCounter.Set_Load(false);

		var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
		oBinaryFileReader.Read( sData, wb );
		ws = wb.getWorksheet( wb.getActive() );
		AscCommonExcel.getFormulasInfo();
	}

	module( "open oox" );
	test( "Tables", function () {
		var res = '"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
			'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" id="1" xr:uid="{A86CAF0D-84D6-44BD-A6C6-FD6D74D5D380}" name="Table1" displayName="Table1" ref="B2:B3" insertRow="1" totalsRowShown="0"><autoFilter ref="B2:B3" xr:uid="{A86CAF0D-84D6-44BD-A6C6-FD6D74D5D380}"/><tableColumns count="1"><tableColumn id="1" xr3:uid="{6047EC3F-80CF-4339-AAC1-69DA7B2C5951}" name="Column1"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
		var oNewTable = ws.createTablePart();
		var reader = new StaxParser(res, oNewTable);
		oNewTable.fromXml(reader);

		var writer = new AscCommon.CMemory();
		oNewTable.toXml(writer);

		var test = "\"";
		for (var i = 0; i < writer.data.length; i++) {
			test += String.fromCharCode(writer.data[i]);
		}

		/*var test = "\"";
		for (var i = 0; i < writer.data.length; i++) {
			var sym = String.fromCharCode(writer.data[i]);
			if (sym === "<" && !start) {
				start = i;
			}
			if (sym === ">") {
				strictEqual(test, res.substring(start, test.length), 'Read');
				start = i;
				test = "";
			}
			test += String.fromCharCode(writer.data[i]);
		}*/

		strictEqual(test, res, 'Read');
	} );
} );
