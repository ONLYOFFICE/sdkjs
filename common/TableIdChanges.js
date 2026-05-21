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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
	function(window, undefined)
{
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBase}
	 */
	function CChangesTableIdAdd(Class, Id, NewClass)
	{
		AscDFH.CChangesBase.call(this, Class);

		this.Id       = Id;
		this.NewClass = NewClass;
	}

	CChangesTableIdAdd.prototype = Object.create(AscDFH.CChangesBase.prototype);
	CChangesTableIdAdd.prototype.constructor = CChangesTableIdAdd;
	CChangesTableIdAdd.prototype.Type = AscDFH.historyitem_TableId_Add;
	CChangesTableIdAdd.prototype.Undo = function()
	{
	};
	CChangesTableIdAdd.prototype.Redo = function()
	{
	};
	CChangesTableIdAdd.prototype.WriteToBinary = function(Writer)
	{
		// String   : Element Id
		// Variable : The element itself

		Writer.WriteString2(this.Id);
		this.NewClass.Write_ToBinary2(Writer);
	};
	CChangesTableIdAdd.prototype.ReadFromBinary = function(Reader)
	{
		// String   : Element Id
		// Variable : The element itself

		this.Id       = Reader.GetString2();
		this.NewClass = this.private_ReadClassFromBinary(Reader);
	};
	CChangesTableIdAdd.prototype.Load = function(Color)
	{
		if (!this.NewClass)
			return;
		
		this.Class.m_aPairs[this.Id] = this.NewClass;
	};
	CChangesTableIdAdd.prototype.RefreshRecalcData = function()
	{
	};
	CChangesTableIdAdd.prototype.private_ReadClassFromBinary = function(Reader)
	{
		var oTableId = this.Class;

		var ElementType = Reader.GetLong();

		oTableId.TurnOff();
		var Element = oTableId.GetClassFromFactory(ElementType);

		if (null !== Element)
			Element.Read_FromBinary2(Reader);

		oTableId.TurnOn();

		return Element;
	};
	CChangesTableIdAdd.prototype.CreateReverseChange = function()
	{
		return null;
	};
	window["AscCommon"].CChangesTableIdAdd = CChangesTableIdAdd;
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBase}
	 */
	function CChangesTableIdDescription(Class, FileCheckSum, FileSize, Description, ItemsCount, PointIndex, StartPoint, LastPoint, SumIndex, DeletedIndex)
	{
		AscDFH.CChangesBase.call(this, Class);

		this.FileCheckSum  = FileCheckSum;
		this.FileSize      = FileSize;
		this.Description   = Description;
		this.ItemsCount    = ItemsCount;
		this.PointIndex    = PointIndex;
		this.StartPoint    = StartPoint;
		this.LastPoint     = LastPoint;
		this.SumIndex      = SumIndex;
		this.DeletedIndex  = DeletedIndex;
		this.VersionString = AscCommon.g_cProductVersion + "." + AscCommon.g_cBuildNumber;
	}

	CChangesTableIdDescription.prototype = Object.create(AscDFH.CChangesBase.prototype);
	CChangesTableIdDescription.prototype.constructor = CChangesTableIdDescription;
	CChangesTableIdDescription.prototype.Type = AscDFH.historyitem_TableId_Description;
	CChangesTableIdDescription.prototype.Undo = function()
	{
	};
	CChangesTableIdDescription.prototype.Redo = function()
	{
	};
	CChangesTableIdDescription.prototype.WriteToBinary = function(Writer)
	{
		// Long : FileCheckSum
		// Long : FileSize
		// Long : Description
		// Long : ItemsCount
		// Long : PointIndex
		// Long : StartPoint
		// Long : LastPoint
		// Long : SumIndex
		// Long : DeletedIndex
		// String : SDK version

		Writer.WriteLong(this.FileCheckSum);
		Writer.WriteLong(this.FileSize);
		Writer.WriteLong(this.Description);
		Writer.WriteLong(this.ItemsCount);
		Writer.WriteLong(this.PointIndex);
		Writer.WriteLong(this.StartPoint);
		Writer.WriteLong(this.LastPoint);
		Writer.WriteLong(this.SumIndex);
		Writer.WriteLong(null === this.DeletedIndex ? -10 : this.DeletedIndex);
		Writer.WriteString2(this.VersionString);
	};
	CChangesTableIdDescription.prototype.ReadFromBinary = function(Reader)
	{
		// Long : FileCheckSum
		// Long : FileSize
		// Long : Description
		// Long : ItemsCount
		// Long : PointIndex
		// Long : StartPoint
		// Long : LastPoint
		// Long : SumIndex
		// Long : DeletedIndex
		// String : SDK version

		this.FileCheckSum  = Reader.GetLong();
		this.FileSize      = Reader.GetLong();
		this.Description   = Reader.GetLong();
		this.ItemsCount    = Reader.GetLong();
		this.PointIndex    = Reader.GetLong();
		this.StartPoint    = Reader.GetLong();
		this.LastPoint     = Reader.GetLong();
		this.SumIndex      = Reader.GetLong();
		this.DeletedIndex  = Reader.GetLong();
		this.VersionString = Reader.GetString2();
	};
	CChangesTableIdDescription.prototype.GetBinarySize = function()
	{
		return (9 * 4 + 4 + ((this.VersionString.length & 0x7FFFFFFF) * 2));
	};
	CChangesTableIdDescription.prototype.Load = function(Color)
	{
		// var CollaborativeEditing = AscCommon.CollaborativeEditing;
		// // CollaborativeEditing LOG
		// console.log("ItemsCount2  " + CollaborativeEditing.m_nErrorLog_PointChangesCount);
		// if (CollaborativeEditing.m_nErrorLog_PointChangesCount !== CollaborativeEditing.m_nErrorLog_SavedPCC)
		// 	console.log("========================= BAD number of changes in the Point =====================");
		// if (CollaborativeEditing.m_nErrorLog_CurPointIndex + 1 !== this.PointIndex && 0 !== this.PointIndex)
		// 	console.log("========================= BAD Point number =======================================");
		// var bBadSumIndex = false;
		// if (0 === this.PointIndex)
		// {
		// 	CollaborativeEditing.m_nErrorLog_SumIndex = 0;
		// }
		// else
		// {
		// 	// Because we don't count this change
		// 	CollaborativeEditing.m_nErrorLog_SumIndex += CollaborativeEditing.m_nErrorLog_SavedPCC + 1;
		// 	if (this.PointIndex === this.StartPoint)
		// 	{
		// 		if (CollaborativeEditing.m_nErrorLog_SumIndex !== this.SumIndex)
		// 			bBadSumIndex = true;
		//
		// 		console.log("SumIndex2    " + CollaborativeEditing.m_nErrorLog_SumIndex);
		// 		CollaborativeEditing.m_nErrorLog_SumIndex = this.SumIndex;
		// 	}
		// }
		//
		// console.log("----------------------------");
		// console.log("FileCheckSum " + this.FileCheckSum);
		// console.log("FileSize     " + this.FileSize);
		// console.log("Description  " + this.Description + " " + AscDFH.GetHistoryPointStringDescription(this.Description));
		// console.log("PointIndex   " + this.PointIndex);
		// console.log("StartPoint   " + this.StartPoint);
		// console.log("LastPoint    " + this.LastPoint);
		// console.log("ItemsCount   " + this.ItemsCount);
		// console.log("SumIndex     " + this.SumIndex);
		// console.log("SDK build    " + this.VersionString);
		// console.log("DeletedIndex " + (-10 === this.DeletedIndex ? null : this.DeletedIndex));
		// // -1 because we shouldn't count this change
		// CollaborativeEditing.m_nErrorLog_SavedPCC          = this.ItemsCount;
		// CollaborativeEditing.m_nErrorLog_PointChangesCount = -1;
		// CollaborativeEditing.m_nErrorLog_CurPointIndex     = this.PointIndex;
		// if (bBadSumIndex)
		// 	console.log("========================= BAD Sum index ==========================================");
	};
	CChangesTableIdDescription.prototype.RefreshRecalcData = function()
	{
	};
	CChangesTableIdDescription.prototype.CreateReverseChange = function()
	{
		return null;
	};
	CChangesTableIdDescription.prototype.IsDescriptionChange = function()
	{
		return true;
	};
	window["AscCommon"].CChangesTableIdDescription = CChangesTableIdDescription;
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBase}
	 */
	function CChangesCommonAddWaterMark(Class, Url)
	{
		AscDFH.CChangesBase.call(this, Class);

		this.Url = Url ? Url : "";
	}

	CChangesCommonAddWaterMark.prototype = Object.create(AscDFH.CChangesBase.prototype);
	CChangesCommonAddWaterMark.prototype.constructor = CChangesCommonAddWaterMark;
	CChangesCommonAddWaterMark.prototype.Type = AscDFH.historyitem_Common_AddWatermark;
	CChangesCommonAddWaterMark.prototype.Undo = function()
	{
	};
	CChangesCommonAddWaterMark.prototype.Redo = function()
	{
	};
	CChangesCommonAddWaterMark.prototype.WriteToBinary = function(Writer)
	{
		Writer.WriteString2(this.Url);
	};
	CChangesCommonAddWaterMark.prototype.ReadFromBinary = function(Reader)
	{
		this.Url = Reader.GetString2();
	};
	CChangesCommonAddWaterMark.prototype.Load = function(Color)
	{
		var sUrl = this.Url;
		if (editor && editor.WordControl && editor.WordControl.m_oLogicDocument)
		{
			var oLogicDocument = editor.WordControl.m_oLogicDocument;
			if (oLogicDocument instanceof AscCommonWord.CDocument)
			{
				var oParaDrawing    = oLogicDocument.DrawingObjects.getTrialImage(sUrl);
				var oFirstParagraph = oLogicDocument.Get_FirstParagraph();
				AscFormat.ExecuteNoHistory(function()
				{
					var oRun = new AscCommonWord.ParaRun();
					oRun.Content.splice(0, 0, oParaDrawing);
					oFirstParagraph.Content.splice(0, 0, oRun);
					oLogicDocument.DrawingObjects.addGraphicObject(oParaDrawing);
				}, this, []);
			}
			else if (oLogicDocument instanceof AscCommonSlide.CPresentation)
			{
				if (oLogicDocument.Slides[0])
				{
					var oDrawing            = oLogicDocument.Slides[0].graphicObjects.createWatermarkImage(sUrl);
					oDrawing.spPr.xfrm.offX = (oLogicDocument.GetWidthMM() - oDrawing.spPr.xfrm.extX) / 2;
					oDrawing.spPr.xfrm.offY = (oLogicDocument.GetHeightMM() - oDrawing.spPr.xfrm.extY) / 2;
					oDrawing.parent         = oLogicDocument.Slides[0];
					oLogicDocument.Slides[0].cSld.spTree.push(oDrawing);
				}
			}
		}
		else
		{
			var oWsModel = window["Asc"]["editor"].wbModel.aWorksheets[0];
			if (oWsModel)
			{
				var objectRender          = new AscFormat.DrawingObjects();
				var oNewDrawing           = objectRender.createDrawingObject(AscCommon.c_oAscCellAnchorType.cellanchorAbsolute);
				var oImage                = AscFormat.DrawingObjectsController.prototype.createWatermarkImage(sUrl);
				oNewDrawing.ext.cx        = oImage.spPr.xfrm.extX;
				oNewDrawing.ext.cy        = oImage.spPr.xfrm.extY;
				oNewDrawing.graphicObject = oImage;
				oWsModel.Drawings.push(oNewDrawing);
			}
		}
	};
	CChangesCommonAddWaterMark.prototype.RefreshRecalcData = function()
	{
	};
	CChangesCommonAddWaterMark.prototype.CreateReverseChange = function()
	{
		return null;
	};
	window["AscCommon"].CChangesCommonAddWaterMark = CChangesCommonAddWaterMark;
})(window);



AscDFH.changesFactory[AscDFH.historyitem_TableId_Add]         = AscCommon.CChangesTableIdAdd;
AscDFH.changesFactory[AscDFH.historyitem_TableId_Description] = AscCommon.CChangesTableIdDescription;

AscDFH.changesFactory[AscDFH.historyitem_Common_AddWatermark] = AscCommon.CChangesCommonAddWaterMark;


//----------------------------------------------------------------------------------------------------------------------
// Changes dependency map
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_TableId_Add]         = [AscDFH.historyitem_TableId_Add];
AscDFH.changesRelationMap[AscDFH.historyitem_TableId_Reset]       = [AscDFH.historyitem_TableId_Reset];
AscDFH.changesRelationMap[AscDFH.historyitem_TableId_Description] = [AscDFH.historyitem_TableId_Description];
AscDFH.changesRelationMap[AscDFH.historyitem_Common_AddWatermark] = [AscDFH.historyitem_Common_AddWatermark];
//----------------------------------------------------------------------------------------------------------------------
