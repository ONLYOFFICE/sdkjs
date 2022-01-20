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

(function (window, undefined) {

function CHistory(Document)
{
    this.Index      = -1;
    this.SavedIndex = null;        // Номер точки отката, на которой произошло последнее сохранение
    this.ForceSave  = false;       // Нужно сохранение, случается, когда у нас точка SavedIndex смещается из-за объединения точек, и мы делаем Undo
    this.RecIndex   = -1;          // Номер точки, на которой произошел последний пересчет
    this.Points     = [];          // Точки истории, в каждой хранится массив с изменениями после текущей точки
    this.Document   = Document;
    this.Api                  = null;
    this.CollaborativeEditing = null;
    this.CanNotAddChanges = false; // флаг для отслеживания ошибок добавления изменений без точки:Create_NewPoint->Add->Save_Changes->Add
	this.CollectChanges   = false;

	this.RecalculateData =
	{
		Inline       : {
			Pos     : -1,
			PageNum : 0
		},
		Flow         : [],
		HdrFtr       : [],
		Drawings     : {
			All       : false,
			Map       : {},
			ThemeInfo : null
		},
		Tables       : [],
		NumPr        : [],
		NotesEnd     : false,
		NotesEndPage : 0,
		LineNumbers  : false,
		Update       : true
	};

	this.TurnOffHistory = 0;
    this.MinorChanges   = false; // Данный параметр нужен, чтобы определить влияют ли добавленные изменения на пересчет

    this.BinaryWriter = new AscCommon.CMemory();

    this.FileCheckSum = 0;
    this.FileSize     = 0;

    // Параметры для специального сохранения для локальной версии редактора
    this.UserSaveMode   = false;
    this.UserSavedIndex = null;  // Номер точки, на которой произошло последнее сохранение пользователем (не автосохранение)

	this.StoredData = [];
}

CHistory.prototype =
{
    Set_LogicDocument : function(LogicDocument)
    {
        if (!LogicDocument)
            return;

        this.Document             = LogicDocument;
        this.Api                  = LogicDocument.Get_Api();
        this.CollaborativeEditing = LogicDocument.Get_CollaborativeEditing();
    },

    Is_UserSaveMode : function()
    {
        return this.UserSaveMode;
    },

    Update_FileDescription : function(oStream)
    {
        var pData = oStream.data;
        var nSize = oStream.size;

        this.FileCheckSum = AscCommon.g_oCRC32.Calculate_ByByteArray(pData, nSize);
        this.FileSize     = nSize;
    },

	Update_PointInfoItem : function(PointIndex, StartPoint, LastPoint, SumIndex, DeletedIndex)
	{
		var Point = this.Points[PointIndex];
		if (Point)
		{
			// Проверяем первое изменение. Если оно уже нужного типа, тогда мы его удаляем. Добавляем специфическое
			// первое изменение с описанием.
			var Class = AscCommon.g_oTableId;

			if (Point.Items.length > 0)
			{
				var FirstItem = Point.Items[0];
				if (FirstItem.Class === Class && AscDFH.historyitem_TableId_Description === FirstItem.Data.Type)
					Point.Items.splice(0, 1);
			}

			var Data = new AscCommon.CChangesTableIdDescription(Class,
				this.FileCheckSum,
				this.FileSize,
				Point.Description,
				Point.Items.length,
				PointIndex,
				StartPoint,
				LastPoint,
				SumIndex,
				DeletedIndex
			);

			var Binary_Pos = this.BinaryWriter.GetCurPosition();
			this.BinaryWriter.WriteString2(Class.Get_Id());
			this.BinaryWriter.WriteLong(Data.Type);
			Data.WriteToBinary(this.BinaryWriter);

			var Binary_Len = this.BinaryWriter.GetCurPosition() - Binary_Pos;

			var Item = {
				Class  : Class,
				Data   : Data,
				Binary : {
					Pos : Binary_Pos,
					Len : Binary_Len
				},

				NeedRecalc : false
			};

			Point.Items.splice(0, 0, Item);
		}
	},

    Is_Clear : function()
    {
        if ( this.Points.length <= 0 )
            return true;

        return false;
    },

    Clear : function()
    {
        this.Index          = -1;
        this.SavedIndex     = null;
        this.ForceSave      = false;
        this.UserSavedIndex = null;
        this.Points.length  = 0;
		this.TurnOffHistory = 0;
        this.private_ClearRecalcData();
    },

    Can_Undo : function()
    {
        if ( this.Index >= 0 )
            return true;

        return false;
    },

    Can_Redo : function()
    {
        if ( this.Points.length > 0 && this.Index < this.Points.length - 1 )
            return true;

        return false;
    },

    UndoLastPoint : function(nBottomIndex)
    {
        var oPoint = this.Points[this.Index];
        if(oPoint)
        {
            var aItems = oPoint.Items;
            var _bottomIndex;
            if(AscFormat.isRealNumber(nBottomIndex))
            {
                _bottomIndex = nBottomIndex - 1;
            }
            else
            {
                _bottomIndex = -1;
            }
            for (var i = aItems.length - 1; i > _bottomIndex; i--)
            {
                var oItem = aItems[i];
                oItem.Data.Undo();
            }
            oPoint.Items.length = _bottomIndex + 1;
            this.Document.SetSelectionState( oPoint.State );
        }
    },

    Undo : function(Options)
    {
    	var arrChanges = [];

        this.CheckUnionLastPoints();

        // Проверяем можно ли сделать Undo
        if (true !== this.Can_Undo())
            return null;

        // Запоминаем самое последнее состояние документа для Redo
        if ( this.Index === this.Points.length - 1 )
            this.LastState = this.Document.GetSelectionState();
        
        this.Document.RemoveSelection(true);

        var Point = null;
        if (undefined !== Options && null !== Options && true === Options.All)
        {
            while (this.Index >= 0)
            {
                Point = this.Points[this.Index--];

                // Откатываем все действия в обратном порядке (относительно их выполенения)
                for (var Index = Point.Items.length - 1; Index >= 0; Index--)
                {
                    var Item = Point.Items[Index];

					if (Item.Data)
					{
						Item.Data.Undo();
						arrChanges.push(Item.Data);
					}
                    this.private_UpdateContentChangesOnUndo(Item);
                }
            }
        }
        else
        {
            Point = this.Points[this.Index--];

            // Откатываем все действия в обратном порядке (относительно их выполенения)
            for (var Index = Point.Items.length - 1; Index >= 0; Index--)
            {
                var Item = Point.Items[Index];
				if (Item.Data)
				{
					Item.Data.Undo();
					arrChanges.push(Item.Data);
				}
				this.private_UpdateContentChangesOnUndo(Item);
            }
        }

        if (null != Point)
            this.Document.SetSelectionState( Point.State );

		if(!window['AscCommon'].g_specialPasteHelper.specialPasteStart)
		{
			window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide(true);
		}
		
        return arrChanges;
    },

    Redo : function()
    {
		var arrChanges = [];

        // Проверяем можно ли сделать Redo
        if (true !== this.Can_Redo())
            return null;

        this.Document.RemoveSelection(true);
        
        var Point = this.Points[++this.Index];

        // Выполняем все действия в прямом порядке
        for ( var Index = 0; Index < Point.Items.length; Index++ )
        {
            var Item = Point.Items[Index];

			if (Item.Data)
			{
				Item.Data.Redo();
				arrChanges.push(Item.Data);
			}
			this.private_UpdateContentChangesOnRedo(Item);
        }

        // Восстанавливаем состояние на следующую точку
        var State = null;
        if ( this.Index === this.Points.length - 1 )
            State = this.LastState;
        else
            State = this.Points[this.Index + 1].State;

        this.Document.SetSelectionState( State );
		
		if(!window['AscCommon'].g_specialPasteHelper.pasteStart)
		{
			window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
		}
		
        return arrChanges;
    },

	/**
	 * Создаем новую точку в истории
	 * @param {number} nDescription - идентификатор производимого действия
	 * @param {object} [oSelectionState=undefined] - сохраненное состояние редактора до начала действия (если не задано используем состояние на текущий момент)
	 * @returns {boolean}
	 */
    Create_NewPoint : function(nDescription, oSelectionState)
    {
		if ( 0 !== this.TurnOffHistory )
			return false;

		if (this.Document && this.Document.ClearListsCache)
			this.Document.ClearListsCache();

        this.CanNotAddChanges = false;
		this.CollectChanges   = false;

		if (null !== this.SavedIndex && this.Index < this.SavedIndex)
            this.Set_SavedIndex(this.Index);

        this.ClearAdditional();

        this.CheckUnionLastPoints();
        
        var State = oSelectionState ? oSelectionState : this.Document.GetSelectionState();
        var Items = [];
        var Time  = new Date().getTime();

        // Создаем новую точку
        this.Points[++this.Index] =
        {
            State      : State, // Текущее состояние документа (курсор, селект)
            Items      : Items, // Массив изменений, начиная с текущего момента
            Time       : Time,  // Текущее время
            Additional : {},    // Дополнительная информация
            Description: nDescription
        };

        // Удаляем ненужные точки
        this.Points.length = this.Index + 1;
		
		if(!window['AscCommon'].g_specialPasteHelper.pasteStart)
		{
			window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
		}

		return true;
    },

	/**
	 * Специальная функция, для создания точки, чтобы отловить все изменения, которые происходят. После использования
	 * данная точка ДОЛЖНА быть удалена через функцию Remove_LastPoint.
	 */
	CreateNewPointForCollectChanges : function()
	{
		this.Points[++this.Index] = {
			State       : null,
			Items       : [],
			Time        : null,
			Additional  : {},
			Description : -1
		};

		this.Points.length  = this.Index + 1;
		this.CollectChanges = true;
	},
    
    Remove_LastPoint : function()
	{
		if (this.Index > -1)
		{
			this.CollectChanges = false;
			this.Index--;
			this.Points.length = this.Index + 1;
		}
	},

    Is_LastPointEmpty : function()
    {
        if (!this.Points[this.Index] || this.Points[this.Index].Items.length <= 0)
            return true;

        return false;
    },

    Is_LastPointNeedRecalc : function()
    {
        if (!this.Points[this.Index])
            return false;

        var RecalcData = this.Get_RecalcData();
        if (RecalcData.Flow.length > 0 || RecalcData.HdrFtr.length > 0 || -1 !== RecalcData.Inline.Pos || true === RecalcData.Drawings.All)
            return true;

        for(var Key in RecalcData.Drawings.Map)
        {
            if(null != AscCommon.g_oTableId.Get_ById(Key))
            {
                return true;
            }
        }

        return false;
    },

    Clear_Redo : function()
    {
        // Удаляем ненужные точки
        this.Points.length = this.Index + 1;
    },

    // Регистрируем новое изменение:
    // Class - объект, в котором оно произошло
    // Data  - сами изменения
	Add : function(_Class, Data)
	{
		if (!this.CanAddChanges())
			return;

		this._CheckCanNotAddChanges();

		// Заглушка на случай, если у нас во время создания одной точки в истории, после нескольких изменений идет
		// пересчет, потом снова добавляются изменения и снова запускается пересчет и т.д.
		if (this.RecIndex >= this.Index)
			this.RecIndex = this.Index - 1;

		var Binary_Pos = this.BinaryWriter.GetCurPosition();

		var Class;
		if (_Class) {
            Class = _Class.GetClass();
            Data = _Class;

            this.BinaryWriter.WriteString2(Class.Get_Id());
            this.BinaryWriter.WriteLong(_Class.Type);
            _Class.WriteToBinary(this.BinaryWriter);
        }

        if (Class && Class.SetIsRecalculated && (!_Class || _Class.IsNeedRecalculate()))
        	Class.SetIsRecalculated(false);

		var Binary_Len = this.BinaryWriter.GetCurPosition() - Binary_Pos;
		var Item       = {
			Class  : Class,
			Data   : Data,
			Binary : {
				Pos : Binary_Pos,
				Len : Binary_Len
			},

			NeedRecalc : !this.MinorChanges && (!_Class || _Class.IsNeedRecalculate() || _Class.IsNeedRecalculateLineNumbers())
		};

		this.Points[this.Index].Items.push(Item);

		if (!this.CollaborativeEditing)
			return;

		if (_Class)
		{
			if (_Class.IsContentChange())
			{
				var bAdd  = _Class.IsAdd();
				var Count = _Class.GetItemsCount();

				var ContentChanges = new AscCommon.CContentChangesElement(bAdd == true ? AscCommon.contentchanges_Add : AscCommon.contentchanges_Remove, Data.Pos, Count, Item);
				Class.Add_ContentChanges(ContentChanges);
				this.CollaborativeEditing.Add_NewDC(Class);

				if (true === bAdd)
					this.CollaborativeEditing.Update_DocumentPositionsOnAdd(Class, Data.Pos);
				else
					this.CollaborativeEditing.Update_DocumentPositionsOnRemove(Class, Data.Pos, Count);
			}
		    if(_Class.IsPosExtChange()){
                this.CollaborativeEditing.AddPosExtChanges(Item, _Class);
            }
		}
	},

    RecalcData_Add : function(Data)
    {
        if (true !== this.RecalculateData.Update)
            return;

        if ( "undefined" === typeof(Data) || null === Data )
            return;

        switch ( Data.Type )
        {
            case AscDFH.historyitem_recalctype_Flow:
            {
                var bNew = true;
                for ( var Index = 0; Index < this.RecalculateData.Flow.length; Index++ )
                {
                    if ( this.RecalculateData.Flow[Index] === Data.Data )
                    {
                        bNew = false;
                        break;
                    }
                }

                if ( true === bNew )
                    this.RecalculateData.Flow.push( Data.Data );

                break;
            }

            case AscDFH.historyitem_recalctype_HdrFtr:
            {
                if ( null === Data.Data )
                    break;

                var bNew = true;
                for ( var Index = 0; Index < this.RecalculateData.HdrFtr.length; Index++ )
                {
                    if ( this.RecalculateData.HdrFtr[Index] === Data.Data )
                    {
                        bNew = false;
                        break;
                    }
                }

                if ( true === bNew )
                    this.RecalculateData.HdrFtr.push( Data.Data );

                break
            }

            case AscDFH.historyitem_recalctype_Inline:
            {
                if ( (Data.Data.Pos < this.RecalculateData.Inline.Pos) || (Data.Data.Pos === this.RecalculateData.Inline.Pos && Data.Data.PageNum < this.RecalculateData.Inline.PageNum) || this.RecalculateData.Inline.Pos < 0 )
                {
                    this.RecalculateData.Inline.Pos     = Data.Data.Pos;
                    this.RecalculateData.Inline.PageNum = Data.Data.PageNum;
                }

                break;
            }
            case AscDFH.historyitem_recalctype_Drawing:
            {
                if(!this.RecalculateData.Drawings.All)
                {
                    if(Data.All)
                    {
                        this.RecalculateData.Drawings.All = true;
                    }
                    else
                    {
                        if(Data.Theme)
                        {
                            this.RecalculateData.Drawings.ThemeInfo =
                            {
                                Theme: true,
                                ArrInd: Data.ArrInd
                            }
                        }
                        else if(Data.ColorScheme)
                        {
                            this.RecalculateData.Drawings.ThemeInfo =
                            {
                                ColorScheme: true,
                                ArrInd: Data.ArrInd
                            }
                        }
                        else
                        {
                            this.RecalculateData.Drawings.Map[Data.Object.Get_Id()] = Data.Object;
                        }
                    }
                }
                break;
            }

			case AscDFH.historyitem_recalctype_NotesEnd:
			{
				this.RecalculateData.NotesEnd     = true;
				this.RecalculateData.NotesEndPage = Data.PageNum;
				break;
			}
        }
    },

	AddChangedStyleToRecalculateData : function(sId, oStyle)
	{
		if (!this.RecalculateData.ChangedStyles)
			this.RecalculateData.ChangedStyles = {};

		if (this.RecalculateData.ChangedStyles[sId] === oStyle)
			return false;

		this.RecalculateData.ChangedStyles[sId] = oStyle;
		return true;
	},

	AddChangedNumberingToRecalculateData : function(NumId, Lvl, oNum)
	{
		if (!this.RecalculateData.ChangedNums)
			this.RecalculateData.ChangedNums = {};

		if (!this.RecalculateData.ChangedNums[NumId])
			this.RecalculateData.ChangedNums[NumId] = {};

		if (this.RecalculateData.ChangedNums[NumId][Lvl] === oNum)
			return false;

		this.RecalculateData.ChangedNums[NumId][Lvl] = oNum;
		return true;
	},

    Add_RecalcNumPr : function(NumPr)
    {
        if (undefined !== NumPr && null !== NumPr && undefined !== NumPr.NumId)
            this.RecalculateData.NumPr[NumPr.NumId] = true;
    },

    Add_RecalcTableGrid : function(TableId)
    {
        this.RecalculateData.Tables[TableId] = true;
    },

	AddLineNumbersToRecalculateData : function()
	{
		this.RecalculateData.LineNumbers = true;
	},

    OnEnd_GetRecalcData : function()
    {
        // Пересчитываем таблицы
        for (var TableId in this.RecalculateData.Tables)
        {
            var Table = AscCommon.g_oTableId.Get_ById(TableId);
            if (null !== Table && Table.Is_UseInDocument())
            {
                if (true === Table.Check_ChangedTableGrid())
                {
                    Table.Refresh_RecalcData2(0, 0);
                }
            }
        }

        // Делаем это, чтобы пересчитались ячейки таблиц, в которых есть заданная нумерация. Но нам не нужно менять
        // начальную точку пересчета здесь, т.к. начальная точка уже рассчитана правильно.
        this.RecalculateData.Update = false;
        for (var NumId in this.RecalculateData.NumPr)
        {
            var NumPr = new CNumPr();
            NumPr.NumId = NumId;
            for (var Lvl = 0; Lvl < 9; ++Lvl)
            {
                NumPr.Lvl = Lvl;
                var AllParagraphs = this.Document.GetAllParagraphsByNumbering(NumPr);
                var Count = AllParagraphs.length;
                for (var Index = 0; Index < Count; ++Index)
                {
                    var Para = AllParagraphs[Index];
                    Para.Refresh_RecalcData2(0);
                }
            }
        }
        this.RecalculateData.NumPr = [];
        this.RecalculateData.Update = true;
    },

	CheckUnionLastPoints : function()
    {
        // Не объединяем точки в истории, когда отключается пересчет.
        // TODO: Неправильно изменяется RecalcIndex
        if (true !== this.Document.Is_OnRecalculate())
            return false;

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
            NewDescription = AscDFH.historydescription_Document_CompositeInput;
        }
        else
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

        if (0 !== StartIndex1)
            Point1.Items.splice(0, 1);

        if (0 !== StartIndex2)
            Point2.Items.splice(0, 1);

        var NewPoint =
        {
            State      : Point1.State,
            Items      : Point1.Items.concat(Point2.Items),
            Time       : Point1.Time,
            Additional : {},
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
	},

    CanRemoveLastPoint : function()
    {
        if (this.Points.length <= 0
            || (true !== this.Is_UserSaveMode() && null !== this.SavedIndex && this.SavedIndex >= this.Points.length - 1)
            || (true === this.Is_UserSaveMode() && null !== this.UserSavedIndex && this.UserSavedIndex >= this.Points.length - 1))
            return false;

        return true;
    },

    TurnOff : function()
    {
		this.TurnOffHistory++;
    },

    TurnOn : function()
    {
		this.TurnOffHistory--;
		if(this.TurnOffHistory < 0)
			this.TurnOffHistory = 0;
    },

	/** @returns {boolean} */
    Is_On : function()
    {
		return (0 === this.TurnOffHistory);
    },

	/** @returns {boolean} */
	IsOn : function()
	{
		return (0 === this.TurnOffHistory);
	},

	Reset_SavedIndex : function(IsUserSave)
	{
		this.SavedIndex = (null === this.SavedIndex && -1 === this.Index ? null : this.Index);
		if (true === this.Is_UserSaveMode())
		{
			if (true === IsUserSave)
			{
				this.UserSavedIndex = this.Index;
				this.ForceSave      = false;
			}
		}
		else
		{
			this.ForceSave = false;
		}
	},

    Set_SavedIndex : function(Index)
    {
      this.SavedIndex = Index;
      if (true === this.Is_UserSaveMode()) {
        if (null !== this.UserSavedIndex && this.UserSavedIndex > this.SavedIndex) {
          this.UserSavedIndex = Index;
          this.ForceSave = true;
        }
      } else {
        this.ForceSave = true;
      }
    },

	Have_Changes : function(IsNotUserSave)
	{
		if (!this.Document || this.Document.IsViewModeInReview())
			return false;

		var checkIndex = (this.Is_UserSaveMode() && !IsNotUserSave) ? this.UserSavedIndex : this.SavedIndex;
		if (-1 === this.Index && null === checkIndex && false === this.ForceSave)
			return false;

		if (this.Index != checkIndex || true === this.ForceSave)
			return true;

		return false;
	},

    Get_RecalcData : function(oRecalcData, arrChanges, nChangeStart, nChangeEnd)
    {
        if (oRecalcData)
        {
            this.RecalculateData = oRecalcData;
        }
        else if (arrChanges)
		{
			this.private_ClearRecalcData();

			var nStart = undefined !== nChangeStart ? nChangeStart : 0;
			var nEnd   = undefined !== nChangeEnd ? nChangeEnd : arrChanges.length - 1;
			for (var nIndex = nStart; nIndex <= nEnd; ++nIndex)
			{
				var oChange = arrChanges[nIndex];
				oChange.RefreshRecalcData();
			}

			this.private_PostProcessingRecalcData();
		}
        else
        {
            if (this.Index >= 0)
            {
                this.private_ClearRecalcData();

                for (var Pos = this.RecIndex + 1; Pos <= this.Index; Pos++)
                {
                    // Считываем изменения, начиная с последней точки, и смотрим что надо пересчитать.
                    var Point = this.Points[Pos];

                    // Выполняем все действия в прямом порядке
                    for (var Index = 0; Index < Point.Items.length; Index++)
                    {
                        var Item = Point.Items[Index];

                        if (true === Item.NeedRecalc)
                            Item.Class.Refresh_RecalcData(Item.Data);
                    }
                }

                this.private_PostProcessingRecalcData();
            }
        }

        this.OnEnd_GetRecalcData();
        return this.RecalculateData;
    },

    Reset_RecalcIndex : function()
    {
        this.RecIndex = this.Index;
    },

    Set_Additional_ExtendDocumentToPos : function()
    {
        if ( this.Index >= 0 )
        {
            this.Points[this.Index].Additional.ExtendDocumentToPos = true;
        }
    },

    Is_ExtendDocumentToPos : function()
    {
        if ( undefined === this.Points[this.Index] || undefined === this.Points[this.Index].Additional || undefined === this.Points[this.Index].Additional.ExtendDocumentToPos )
            return false;

        return true;
    },

    Get_EditingTime : function(dTime)
    {
        var Count = this.Points.length;

        var TimeLine = [];
        for ( var Index = 0; Index < Count; Index++ )
        {
            var PointTime = this.Points[Index].Time;
            TimeLine.push( { t0 : PointTime - dTime, t1 : PointTime } );
        }

        Count = TimeLine.length;
        for ( var Index = 1; Index < Count; Index++ )
        {
            var CurrEl = TimeLine[Index];
            var PrevEl = TimeLine[Index - 1];
            if ( CurrEl.t0 <= PrevEl.t1 )
            {
                PrevEl.t1 = CurrEl.t1;
                TimeLine.splice( Index, 1 );
                Index--;
                Count--;
            }
        }

        Count = TimeLine.length;
        var OverallTime = 0;
        for ( var Index = 0; Index < Count; Index++ )
        {
            OverallTime += TimeLine[Index].t1 - TimeLine[Index].t0;
        }

        return OverallTime;
    },

    Get_DocumentPositionBinary : function()
    {
        var PosInfo = this.Document.Get_DocumentPositionInfoForCollaborative();
        return AscCommon.CollaborativeEditing.GetDocumentPositionBinary(this.BinaryWriter, PosInfo);
    },

    _CheckCanNotAddChanges : function() {
        try {
            if (this.CanNotAddChanges && this.Api && !this.CollectChanges) {
                var tmpErr = new Error();
                if (tmpErr.stack) {
                    this.Api.CoAuthoringApi.sendChangesError(tmpErr.stack);
                }
            }
        } catch (e) {
        }
    }
};
/**
 * Проверяем, можно ли добавить изменение
 * @returns {boolean}
 */
CHistory.prototype.CanAddChanges = function()
{
	return (0 === this.TurnOffHistory && this.Index >= 0);
};
CHistory.prototype.ClearAdditional = function()
{
	if (this.Index >= 0)
		this.Points[this.Index].Additional = {};

	if (this.Api && true === this.Api.isMarkerFormat)
		this.Api.sync_MarkerFormatCallback(false);

	if (this.Api && true === this.Api.isDrawTablePen)
		this.Api.sync_TableDrawModeCallback(false);

	if (this.Api && true === this.Api.isDrawTableErase)
		this.Api.sync_TableEraseModeCallback(false);
};
CHistory.prototype.private_UpdateContentChangesOnUndo = function(Item)
{
	if (this.private_IsContentChange(Item.Class, Item.Data))
	{
		Item.Class.m_oContentChanges.RemoveByHistoryItem(Item);
	}
};
CHistory.prototype.private_UpdateContentChangesOnRedo = function(Item)
{
	if (this.private_IsContentChange(Item.Class, Item.Data))
	{
		var bAdd  = this.private_IsAddContentChange(Item.Class, Item.Data);
		var Count = this.private_GetItemsCountInContentChange(Item.Class, Item.Data);

		var ContentChanges = new AscCommon.CContentChangesElement(( bAdd == true ? AscCommon.contentchanges_Add : AscCommon.contentchanges_Remove ), Item.Data.Pos, Count, Item);
		Item.Class.Add_ContentChanges(ContentChanges);
		this.CollaborativeEditing.Add_NewDC(Item.Class);
	}
};
CHistory.prototype.private_IsContentChange = function(Class, Data)
{
	var bPresentation = !(typeof CPresentation === "undefined");
	var bSlide = !(typeof Slide === "undefined");
	if ( ( Class instanceof CDocument        && ( AscDFH.historyitem_Document_AddItem        === Data.Type || AscDFH.historyitem_Document_RemoveItem        === Data.Type ) ) ||
		(((Class instanceof CDocumentContent || Class instanceof AscFormat.CDrawingDocContent)) && ( AscDFH.historyitem_DocumentContent_AddItem === Data.Type || AscDFH.historyitem_DocumentContent_RemoveItem === Data.Type ) ) ||
		( Class instanceof CTable           && ( AscDFH.historyitem_Table_AddRow            === Data.Type || AscDFH.historyitem_Table_RemoveRow            === Data.Type ) ) ||
		( Class instanceof CTableRow        && ( AscDFH.historyitem_TableRow_AddCell        === Data.Type || AscDFH.historyitem_TableRow_RemoveCell        === Data.Type ) ) ||
		( Class instanceof Paragraph        && ( AscDFH.historyitem_Paragraph_AddItem       === Data.Type || AscDFH.historyitem_Paragraph_RemoveItem       === Data.Type ) ) ||
		( Class instanceof ParaHyperlink    && ( AscDFH.historyitem_Hyperlink_AddItem       === Data.Type || AscDFH.historyitem_Hyperlink_RemoveItem       === Data.Type ) ) ||
		( Class instanceof ParaRun          && ( AscDFH.historyitem_ParaRun_AddItem         === Data.Type || AscDFH.historyitem_ParaRun_RemoveItem         === Data.Type ) ) ||
		( bPresentation && Class instanceof CPresentation && (AscDFH.historyitem_Presentation_AddSlide === Data.Type || AscDFH.historyitem_Presentation_RemoveSlide === Data.Type)) ||
		( bSlide && Class instanceof Slide && (AscDFH.historyitem_SlideAddToSpTree === Data.Type || AscDFH.historyitem_SlideRemoveFromSpTree === Data.Type))
	)
		return true;

	return false;
};
CHistory.prototype.private_IsAddContentChange = function(Class, Data)
{
	var bPresentation = !(typeof CPresentation === "undefined");
	var bSlide = !(typeof Slide === "undefined");
	return ( ( Class instanceof CDocument        && AscDFH.historyitem_Document_AddItem        === Data.Type ) ||
		( ((Class instanceof CDocumentContent || Class instanceof AscFormat.CDrawingDocContent)) && AscDFH.historyitem_DocumentContent_AddItem === Data.Type ) ||
		( Class instanceof CTable           && AscDFH.historyitem_Table_AddRow            === Data.Type ) ||
		( Class instanceof CTableRow        && AscDFH.historyitem_TableRow_AddCell        === Data.Type ) ||
		( Class instanceof Paragraph        && AscDFH.historyitem_Paragraph_AddItem       === Data.Type ) ||
		( Class instanceof ParaHyperlink    && AscDFH.historyitem_Hyperlink_AddItem       === Data.Type ) ||
		( Class instanceof ParaRun          && AscDFH.historyitem_ParaRun_AddItem         === Data.Type ) ||
		( bPresentation && Class instanceof CPresentation && (AscDFH.historyitem_Presentation_AddSlide === Data.Type )) ||
		( bSlide && Class instanceof Slide && (AscDFH.historyitem_SlideAddToSpTree === Data.Type))
	) ? true : false;
};
CHistory.prototype.private_GetItemsCountInContentChange = function(Class, Data)
{
	if ( ( Class instanceof Paragraph ) ||  ( Class instanceof ParaHyperlink) || ( Class instanceof ParaRun ) ||
		( Class instanceof CDocument        && AscDFH.historyitem_Document_RemoveItem        === Data.Type ) ||
		( ((Class instanceof CDocumentContent || Class instanceof AscFormat.CDrawingDocContent)) && AscDFH.historyitem_DocumentContent_RemoveItem === Data.Type ) )
		return Data.Items.length;

	return 1;
};
CHistory.prototype.GetAllParagraphsForRecalcData = function(Props)
{
	if (!this.RecalculateData.AllParagraphs)
	{
		if (this.Document)
			this.RecalculateData.AllParagraphs = this.Document.GetAllParagraphs({All : true});
		else
			this.RecalculateData.AllParagraphs = [];
	}

	var arrParagraphs = [];
	if (!Props || true === Props.All)
	{
		return this.RecalculateData.AllParagraphs;
	}
	else if (true === Props.Style)
	{
		var arrStylesId = Props.StylesId;
		for (var nParaIndex = 0, nParasCount = this.RecalculateData.AllParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
		{
			var oPara = this.RecalculateData.AllParagraphs[nParaIndex];
			for (var nStyleIndex = 0, nStylesCount = arrStylesId.length; nStyleIndex < nStylesCount; ++nStyleIndex)
			{
				if (oPara.Pr.PStyle === arrStylesId[nStyleIndex])
				{
					arrParagraphs.push(oPara);
					break;
				}
			}
		}
	}
	else if (true === Props.Numbering)
	{
		for (var nParaIndex = 0, nParasCount = this.RecalculateData.AllParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
		{
			var oPara = this.RecalculateData.AllParagraphs[nParaIndex];

			var NumPr  = Props.NumPr;
			var _NumPr = oPara.GetNumPr();

			if (undefined != _NumPr && _NumPr.NumId === NumPr.NumId && (_NumPr.Lvl === NumPr.Lvl || undefined === NumPr.Lvl))
				arrParagraphs.push(oPara);
		}
	}
	return arrParagraphs;
};
CHistory.prototype.GetRecalculateIndex = function()
{
	return this.RecIndex;
};
CHistory.prototype.SetRecalculateIndex = function(nIndex)
{
	this.RecIndex = Math.min(this.Index, nIndex);
};
CHistory.prototype.SaveRedoPoints = function()
{
	var arrData = [];
	this.StoredData.push(arrData);
	for (var nIndex = this.Index + 1, nCount = this.Points.length; nIndex < nCount; ++nIndex)
	{
		arrData.push(this.Points[nIndex]);
	}
};
CHistory.prototype.PopRedoPoints = function()
{
	if (this.StoredData.length <= 0)
		return;

	var arrPoints = this.StoredData[this.StoredData.length - 1];
	this.Points.length = this.Index + 1;
	for (var nIndex = 0, nCount = arrPoints.length; nIndex < nCount; ++nIndex)
	{
		this.Points[this.Index + nIndex + 1] = arrPoints[nIndex];
	}

	this.StoredData.length = this.StoredData.length - 1;
};
CHistory.prototype.RemoveLastPoint = function()
{
	this.Remove_LastPoint();
};
CHistory.prototype.private_ClearRecalcData = function()
{
	// NumPr здесь не обнуляем
	var NumPr            = this.RecalculateData.NumPr;
	this.RecalculateData = {
		Inline   : {
			Pos     : -1,
			PageNum : 0
		},
		Flow     : [],
		HdrFtr   : [],
		Drawings : {
			All       : false,
			Map       : {},
			ThemeInfo : null
		},

		Tables        : [],
		NumPr         : NumPr,
		NotesEnd      : false,
		NotesEndPage  : 0,
		Update        : true,
		ChangedStyles : {},
		ChangedNums   : {},
		LineNumbers   : false,
		AllParagraphs : null
	};
};
/**
 * Обработка изменений после Undo/Redo всех изменений
 */
CHistory.prototype.private_PostProcessingRecalcData = function()
{
	for (var sId in this.RecalculateData.ChangedStyles)
	{
		var oStyle = this.RecalculateData.ChangedStyles[sId];
		oStyle.RecalculateRelatedParagraphs();
	}
};
	/**
	 * Получаем сколько изменений в истории уже сохранено на сервере на данный момент с учетом текущей точки в истории
	 * @returns {number}
	 */
	CHistory.prototype.GetDeleteIndex = function()
	{
		if (null === this.SavedIndex)
			return null;

		var nSum = 0;
		for (var nPointIndex = 0, nLastPoint = Math.min(this.SavedIndex, this.Index); nPointIndex <= nLastPoint; ++nPointIndex)
		{
			nSum += this.Points[nPointIndex].Items.length;
		}

		return nSum;
	};
	/**
	 * Удаляем изменения из истории, которые сохранены на сервере. Это происходит при подключении второго пользователя
	 */
	CHistory.prototype.RemovePointsByDeleteIndex = function()
	{
		var nDeleteIndex = this.GetDeleteIndex();
		if (null === nDeleteIndex)
			return;

		while (nDeleteIndex > 0 && this.Points.length > 0)
		{
			nDeleteIndex -= this.Points[0].Items.length;

			this.Points.splice(0, 1);

			if (this.Index >= 0)
				this.Index--;

			if (this.RecIndex >= 0)
				this.RecIndex--;
		}

		this.SavedIndex = null;
	};
	/**
	 * Получаем массив изменений, которые еще не были пересчитаны
	 * @returns {[]}
	 */
	CHistory.prototype.GetNonRecalculatedChanges = function()
	{
		var arrChanges = [];

		if (this.Index - this.RecIndex !== 1 && this.RecIndex >= -1)
		{
			for (var nPointIndex = this.RecIndex + 1; nPointIndex <= this.Index; ++nPointIndex)
			{
				this.GetChangesFromPoint(nPointIndex, arrChanges);
			}
		}
		else if (this.Index >= 0)
		{
			this.GetChangesFromPoint(this.Index, arrChanges);
		}

		return arrChanges;
	};
	/**
	 * Получем массив изменений	из заданной точки
	 * @param nPointIndex {number}
	 * @param [arrChanges=undefined] {[CChangesBase]}
	 * @returns {[CChangesBase]}
	 */
	CHistory.prototype.GetChangesFromPoint = function(nPointIndex, arrChanges)
	{
		if (!arrChanges)
			arrChanges = [];

		var oHPoint = this.Points[nPointIndex];
		if (oHPoint)
		{
			for (var nIndex = 0, nItemsCount = oHPoint.Items.length; nIndex < nItemsCount; ++nIndex)
			{
				arrChanges.push(oHPoint.Items[nIndex].Data);
			}
		}

		return arrChanges;
	};
	/**
	 * Проверяем перед автозаменой, что действие совершается во время набора
	 * @param oLastElement - последний элемент, добавленный перед автозаменой
	 * @param nHistoryActions - количество точек предществующих автозамене
	 * @param [nMaxTimeDelay=0] - если задано, то максимальное значение времени, которое прошло с момента последнего действия
	 * @returns {boolean}
	 */
	CHistory.prototype.CheckAsYouTypeAutoCorrect = function(oLastElement, nHistoryActions, nMaxTimeDelay)
	{
		// В nHistoryActions задано количество точек, которые предществовали автозамене, т.е.
		// выполнялись действия, которые и вызывали автозамену в итоге. Нам надо проверить предыдущую точку до заданных
		// Если в там происходило добавление заданного элемента, значит у нас был набор текста

		if (this.Index < nHistoryActions)
			return false;

		var nCurIndex = this.Index - nHistoryActions;
		while (this.private_IsPointDoAutoCorrect(nCurIndex) && nCurIndex > 0)
			nCurIndex--;

		if (nCurIndex < 0)
			return false;

		var oPoint = this.Points[nCurIndex];

		if (nMaxTimeDelay && (new Date().getTime() - oPoint.Time) > nMaxTimeDelay)
			return false;

		var nItemsCount = oPoint.Items.length;
		if ((AscDFH.historydescription_Document_AddLetter === oPoint.Description
			|| AscDFH.historydescription_Document_AddLetterUnion === oPoint.Description
			|| AscDFH.historydescription_Document_SpaceButton === oPoint.Description
			|| AscDFH.historydescription_Presentation_ParagraphAdd === oPoint.Description)
			&& nItemsCount > 0)
		{
			var oChange = oPoint.Items[nItemsCount - 1].Data;
			if (!oChange || !oChange.IsContentChange())
				return false;

			var nChangeItemsCount = oChange.GetItemsCount();
			return (nChangeItemsCount > 0 && AscDFH.historyitem_ParaRun_AddItem === oChange.GetType() && oChange.GetItem(nChangeItemsCount - 1) === oLastElement);
		}

		return false;
	};
	CHistory.prototype.private_IsPointDoAutoCorrect = function(nPointIndex)
	{
		if (nPointIndex < 0 || nPointIndex >= this.Points.length)
			return false;

		var nDescription = this.Points[nPointIndex].Description;

		return (AscDFH.historydescription_Document_AutoCorrectCommon === nDescription
			|| AscDFH.historydescription_Document_AutoCorrectFirstLetterOfSentence === nDescription
			|| AscDFH.historydescription_Document_AutoCorrectHyphensWithDash === nDescription
			|| AscDFH.historydescription_Document_AutoCorrectSmartQuotes === nDescription);
	};

	//----------------------------------------------------------export--------------------------------------------------
	window['AscCommon']          = window['AscCommon'] || {};
	window['AscCommon'].CHistory = CHistory;
	window['AscCommon'].History  = new CHistory();
})(window);
