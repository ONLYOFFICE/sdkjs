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
    History.prototype.CanAddChanges = function() {
        return !this.UndoRedoInProgress && AscCommon.CHistory.prototype.CanAddChanges.call(this);
    };
    History.prototype.StartNoHistoryMode = function() {
        this.TurnOff();
        AscCommon.g_oTableId.TurnOff();
    };
    History.prototype.EndNoHistoryMode = function() {
        this.TurnOn();
        AscCommon.g_oTableId.TurnOn();
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
	};
	History.prototype.CheckUnionLastPoints = function()
    {
        // Не объединяем точки во время Undo/Redo
        if (this.Index < this.Points.length - 1)
        	return false;

        // Не объединяем точки истории, если на предыдущей точке произошло сохранение
        if (this.Points.length < 2
            || (true !== this.Is_UserSaveMode() && null !== this.SavedIndex && this.SavedIndex >= this.Points.length - 2)
            || (true === this.Is_UserSaveMode() && null !== this.UserSavedIndex && this.UserSavedIndex >= this.Points.length - 2))
            return false;

        var Point1 = this.Points[this.Points.length - 2];
        var Point2 = this.Points[this.Points.length - 1];

        // запрет на объединение
        if (Point1.forbidUnion || Point2.forbidUnion) {
            return false;
        }

        // Не объединяем слова больше 63 элементов
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
            // Ничего не делаем. Эта ветка означает, что эти две точки можно объединить
            NewDescription = Point1.Description;
        }
		else if (AscDFH.historydescription_Document_CompositeInput === Point1.Description
			|| AscDFH.historydescription_Document_CompositeInputReplace === Point1.Description
			|| AscDFH.historydescription_Document_CompositeInput === Point2.Description
			|| AscDFH.historydescription_Document_CompositeInputReplace === Point2.Description)
		{
			// Композитный ввод не разрешаем объединять ни с чем, кроме композитного ввода
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
	 * Проверяем лок для последних нескольких точек
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
 			}
		}
	};
	//----------------------------------------------------------export--------------------------------------------------
	window['AscPDF'] = window['AscPDF'] || {};
	window['AscPDF'].History = History;
})(window);
