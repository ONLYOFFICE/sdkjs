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

(function (window) {
	
	/**
	 * Local history of changes class
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
		// Do nothing
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
    History.prototype.CanAddChanges = function() {
        return !this.UndoRedoInProgress && AscCommon.CHistory.prototype.CanAddChanges.call(this);
    };
    History.prototype.StartNoHistoryMode = function() {
        this.TurnOff();
        AscCommon.g_oTableId.TurnOff();
    };
    History.prototype.EndNoHistoryMode = function() {
        this.TurnOn();
        if (this.IsOn()) {
            AscCommon.g_oTableId.TurnOn();
        }
    };
	History.prototype.Add = function(_Class, Data) {
		AscCommon.CHistory.prototype.Add.call(this, _Class, Data);
		
		if (!this.CanAddChanges())
			return;
		
		if (_Class.Class && _Class.Class.SetNeedRecalc) {
			if (!this.Points[this.Index].Additional.Pdf) {
				this.Points[this.Index].Additional.Pdf = [];
			}
			
			if (false == this.Points[this.Index].Additional.Pdf.includes(_Class.Class)) {
				this.Points[this.Index].Additional.Pdf.push(_Class.Class);
			}
		}
	};
	History.prototype.CheckUnionLastPoints = function()
    {
        // Don't unite points during Undo/Redo
        if (this.Index < this.Points.length - 1)
        	return false;

        // Don't unite history points if save occurred at the previous point
        if (this.Points.length < 2
            || (true !== this.Is_UserSaveMode() && null !== this.SavedIndex && this.SavedIndex >= this.Points.length - 2)
            || (true === this.Is_UserSaveMode() && null !== this.UserSavedIndex && this.UserSavedIndex >= this.Points.length - 2))
            return false;

        var Point1 = this.Points[this.Points.length - 2];
        var Point2 = this.Points[this.Points.length - 1];

        // union is prohibited
        if (Point1.forbidUnion || Point2.forbidUnion) {
            return false;
        }

        // Don't unite words with more than 63 elements
        if (Point1.Items.length > 63 && AscDFH.historydescription_Document_AddLetterUnion === Point1.Description)
            return false;

        var StartIndex1 = 0;
        var StartIndex2 = 0;
        if (Point1.Items.length > 0 && Point1.Items[0].Data && AscDFH.historyitem_TableId_Description === Point1.Items[0].Data.Type)
            StartIndex1 = 1;

        if (Point2.Items.length > 0 && Point2.Items[0].Data && AscDFH.historyitem_TableId_Description === Point2.Items[0].Data.Type)
            StartIndex2 = 1;

        var NewDescription;
        if ((AscDFH.historydescription_Document_CompositeInput === Point1.Description || AscDFH.historydescription_Document_CompositeInputReplace === Point1.Description)
            && AscDFH.historydescription_Document_CompositeInputReplace === Point2.Description)
        {
            // Do nothing. This case means these two points can be united
            NewDescription = Point1.Description;
        }
		else if (AscDFH.historydescription_Document_CompositeInput === Point1.Description
			|| AscDFH.historydescription_Document_CompositeInputReplace === Point1.Description
			|| AscDFH.historydescription_Document_CompositeInput === Point2.Description
			|| AscDFH.historydescription_Document_CompositeInputReplace === Point2.Description)
		{
			// Composite input is not allowed to unite with anything except composite input
			return false;
		}
		else if ((AscDFH.historydescription_Document_AddLetter === Point1.Description || AscDFH.historydescription_Document_AddLetterUnion === Point1.Description)
			&& (AscDFH.historydescription_Document_AddLetter === Point2.Description || AscDFH.historydescription_Document_AddLetterUnion === Point2.Description))
		{
            var PrevItem = null;
            var Class    = null;
            for (var Index = StartIndex1; Index < Point1.Items.length; Index++)
            {
                var Item = Point1.Items[Index];

                if (null === Class)
                    Class = Item.Class;
                else if (Class != Item.Class || "undefined" === typeof(Class.Check_HistoryUninon) || false === Class.Check_HistoryUninon(PrevItem.Data, Item.Data))
                    return;

                PrevItem = Item;
            }

            for (var Index = StartIndex2; Index < Point2.Items.length; Index++)
            {
                var Item = Point2.Items[Index];

                if (Class != Item.Class || "undefined" === typeof(Class.Check_HistoryUninon) || false === Class.Check_HistoryUninon(PrevItem.Data, Item.Data))
                    return;

                PrevItem = Item;
            }

            NewDescription = AscDFH.historydescription_Document_AddLetterUnion;
        }
		else
		{
			return false;
		}

        if (0 !== StartIndex1)
            Point1.Items.splice(0, 1);

        if (0 !== StartIndex2)
            Point2.Items.splice(0, 1);

		let aObjects = [];
		if (Point1.Additional.Pdf) {
			aObjects = aObjects.concat(Point1.Additional.Pdf);
		}
		if (Point2.Additional.Pdf) {
			aObjects = aObjects.concat(Point2.Additional.Pdf);
		}

        var NewPoint =
        {
            State      : Point1.State,
            Items      : Point1.Items.concat(Point2.Items),
            Time       : Point1.Time,
            Additional : {Pdf: aObjects},
            Description: NewDescription
        };

		if (null !== this.SavedIndex && this.SavedIndex >= this.Points.length - 2)
            this.Set_SavedIndex(this.Points.length - 3);

        this.Points.splice( this.Points.length - 2, 2, NewPoint );
        if ( this.Index >= this.Points.length )
        {
            var DiffIndex = -this.Index + (this.Points.length - 1);
            this.Index    += DiffIndex;
            this.RecIndex  = Math.max( -1, this.RecIndex + DiffIndex);
        }

        return true;
	};
    History.prototype.ForbidUnionPoint = function(nIndex) {
        if (!nIndex) {
            nIndex = this.Points.length - 1;
        }

        if (this.Points[nIndex]) {
            this.Points[nIndex].forbidUnion = true;
        }
    };
	History.prototype.private_IsContentChange = function(Class, Data) {
		if (Data.IsContentChange)
			return Data.IsContentChange();
		
		return AscCommon.CHistory.prototype.private_IsContentChange.call(this, Class, Data);
	};
	History.prototype.private_UpdateContentChangesOnUndo = function(Item)
	{
		if (!this.private_IsContentChange(Item.Class, Item.Data))
			return;
		
		let contentChanges = Item.Data.GetContentChangesClass();
		if (contentChanges)
			contentChanges.RemoveByHistoryItem(Item);
	};
		/**
	 * Check lock for the last several points
	 * @param pointCount
	 */
	History.prototype.checkLock = function(pointCount)
	{
		if (!pointCount || pointCount - 1 > this.Index)
			return;
		
		for (let pointIndex = 0; pointIndex < pointCount; ++pointIndex) {
			let point = this.Points[this.Index - pointIndex];
			
			for (let changeIndex = 0; changeIndex < point.Items.length; ++changeIndex) {
                let oClass = point.Items[changeIndex].Class;

				if (oClass.IsAnnot && oClass.IsAnnot() || oClass.IsForm && oClass.IsForm()) {
                    let oParentPage = oClass.GetParentPage();

                    let check_obj = {
                        "type":     AscPDF.AscLockTypeElemPDF.Object,
                        "pageId":   oParentPage ? oParentPage.GetId() : null,
                        "objId":    oClass.GetId(),
                        "guid":     oClass.GetId()
                    };

                    oClass.Lock.Check(check_obj);
                }
                // lock merge
                else if (oClass instanceof AscPDF.CPDFDoc && point.Items[changeIndex].Data instanceof CChangesPDFDocumentStartMergePages) {
                    const oLocker = oClass.locks.merge; 
                    let sId = oLocker.GetId();
                    let oCheckData = {
                        "type":     AscPDF.AscLockTypeElemPDF.Document,
                        "objId":    sId,
                        "guid":     sId
                    };
                        
                    oLocker.Lock.Check(oCheckData);
                }
 			}
		}
	};

    History.prototype.GetAllMergedPagesBase64 = function() {
        
    };
	//----------------------------------------------------------export--------------------------------------------------
	window['AscPDF'] = window['AscPDF'] || {};
	window['AscPDF'].History = History;
})(window);
