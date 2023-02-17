/*
 * (c) Copyright Ascensio System SIA 2010-2022
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

(function(window, undefined)
{
	const FLAG_MATH          = 0x0001;
	const FLAG_SHAPE         = 0x0002;
	const FLAG_TABLE         = 0x0004;
	const FLAG_NOT_PARAGRAPH = 0x0008; // Есть не параграф в массиве

	/**
	 * Класс, используемый для вставки или переноса содержимого внутри документа
	 * @constructor
	 */
	function CSelectedContent()
	{
		this.Elements       = [];
		this.Flags          = 0;

		this.DrawingObjects = [];
		this.Comments       = [];
		this.CommentsMarks  = {};
		this.Maths          = [];

		this.LogicDocument  = null;

		this.NewCommentsGuid     = false;
		this.SaveNumberingValues = false;
		this.CopyComments        = true;
		this.MoveDrawing         = false; // Только для переноса автофигур
		this.ForceInline         = false;
		this.CursorInLastRun     = false; // TODO: Данный флаг не работает для формул и неинлайновой вставки

		this.InsertOptions = {
			Table : Asc.c_oSpecialPasteProps.overwriteCells
		};

		// Опции для отслеживания переноса
		this.TrackRevisions = false;
		this.MoveTrackId    = null;
		this.MoveTrackRuns  = [];
		this.HaveMovedParts = false;

		this.LastSection = null;

		this.AnchorPos     = null;
		this.Select        = true;
		this.ParaAnchorPos = null;
		this.Run           = null;
		this.PasteHelper   = null;

		this.IsPresentationContent = false;
	}

	CSelectedContent.prototype.Reset = function()
	{
		this.Elements = [];
		this.Flags    = 0;

		this.DrawingObjects = [];
		this.Comments       = [];
		this.Maths          = [];


		this.MoveDrawing = false;
	};
	CSelectedContent.prototype.Add = function(oElement)
	{
		this.Elements.push(oElement);
	};
	CSelectedContent.prototype.EndCollect = function(oLogicDocument)
	{
		this.private_CollectObjects();
		this.private_CheckComments(oLogicDocument);
		this.private_CheckTrackMove(oLogicDocument);
	};
	CSelectedContent.prototype.SetNewCommentsGuid = function(isNew)
	{
		this.NewCommentsGuid = isNew;
	};
	CSelectedContent.prototype.SetMoveDrawing = function(isMoveDrawing)
	{
		this.MoveDrawing = isMoveDrawing;
	};
	CSelectedContent.prototype.IsMoveDrawing = function()
	{
		return this.MoveDrawing;
	};
	CSelectedContent.prototype.SetCopyComments = function(isCopy)
	{
		this.CopyComments = isCopy;
	};
	CSelectedContent.prototype.CanConvertToMath = function()
	{
		// Проверка возможности конвертации имеющегося контента в контент для вставки в формулу
		// Если формулы уже имеются, то ничего не конвертируем
		return !(this.Flags & FLAG_NOT_PARAGRAPH);
	};
	CSelectedContent.prototype.ForceInlineInsert = function(isForce)
	{
		this.ForceInline = undefined === isForce ? true : !!isForce;
	};
	CSelectedContent.prototype.HaveShape = function()
	{
		return !!(this.Flags & FLAG_SHAPE);
	};
	CSelectedContent.prototype.HaveMath = function()
	{
		return !!(this.Flags & FLAG_MATH);
	};
	CSelectedContent.prototype.HaveTable = function()
	{
		return !!(this.Flags & FLAG_TABLE);
	};
	CSelectedContent.prototype.CanInsert = function(oAnchorPos)
	{
		if (this.Elements.length <= 0)
			return false;

		let oParagraph = oAnchorPos.Paragraph;

		var oDocContent = oParagraph.GetParent();
		if (!oDocContent)
			return false;

		// Автофигуры не вставляем в другие автофигуры, сноски и концевые сноски
		// Единственное исключение, если вставка происходит картинки в картиночное поле (для замены картинки)
		let oParentShape = oDocContent.Is_DrawingShape(true);
		if (((oParentShape && !oParentShape.isForm()) || true === oDocContent.IsFootnote()) && true === this.HaveShape())
			return false;

		// В заголовки диаграмм не вставляем формулы
		if(this.HaveMath())
		{
			if(oParagraph.bFromDocument === false)
			{
				let oDrawing = oDocContent.Is_DrawingShape(true);
				if(oDrawing)
				{
					let nDrawingType = null;
					if(oDrawing.getObjectType) 
					{
						nDrawingType = oDrawing.getObjectType();
					}
					if(nDrawingType !== AscDFH.historyitem_type_Shape)
					{
						return false;
					}
				}
			}
		}
		if (oParagraph.bFromDocument === false && (this.DrawingObjects.length > 0 || this.HaveTable()))
			return false;

		let oParaAnchorPos = oParagraph.Get_ParaNearestPos(oAnchorPos);
		if (!oParaAnchorPos || oParaAnchorPos.Classes.length < 2)
			return false;

		let oRun = oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 1];
		if (!oRun || !(oRun instanceof AscCommonWord.ParaRun))
			return false;
		
		// Пока автофигуры не поддерживаются внутри формул, запрещаем их туда всталять
		if (oRun.IsMathRun() && this.IsMoveDrawing())
			return false;

		return (oRun.IsMathRun() ? this.CanConvertToMath() : true);
	};
	CSelectedContent.prototype.Insert = function(oAnchorPos, isSelect)
	{
		if (!this.CanInsert(oAnchorPos))
			return false;

		let oParagraph     = oAnchorPos.Paragraph;
		let oDocContent    = oParagraph.GetParent();
		let oLogicDocument = oParagraph.GetLogicDocument();

		this.LogicDocument = oLogicDocument; // Может быть не задан (например при вставке в формулу в таблицах)
		this.IsPresentationContent = !oParagraph.bFromDocument;

		this.PrepareObjectsForInsert();
		this.private_CheckContentBeforePaste(oAnchorPos, oDocContent);

		let oParaAnchorPos = oParagraph.Get_ParaNearestPos(oAnchorPos);

		let oRun = oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 1];

		this.ParaAnchorPos = oParaAnchorPos;
		this.Select        = isSelect;
		this.Run           = oRun;
		this.AnchorPos     = oAnchorPos;
		this.Select        = !!isSelect;
		this.PasteHelper   = oRun ? oRun.GetParagraph() : null;

		let isLocalTrack = false;
		if (oLogicDocument && oLogicDocument.IsDocumentEditor())
		{
			isLocalTrack = oLogicDocument.GetLocalTrackRevisions();
			oLogicDocument.SetLocalTrackRevisions(false);
		}

		if (oRun.IsMathRun())
		{
			this.private_InsertToMathRun();
		}
		else if (oRun.GetParentPictureContentControl())
		{
			this.private_InsertToPictureCC();
		}
		else if (oRun.GetParentForm())
		{
			this.private_InsertToForm();
		}
		else if (this.private_IsInlineInsert())
		{
			this.private_InsertInline();
		}
		else if (this.private_IsOverwriteTableCells())
		{
			this.private_OverwriteTableCells();
		}
		else
		{
			this.private_InsertCommon();
		}

		if (false !== isLocalTrack)
			oLogicDocument.SetLocalTrackRevisions(isLocalTrack);

		return true;
	};
	CSelectedContent.prototype.ReplaceContent = function(oDocContent, isSelect)
	{
		if (this.Elements.length <= 0)
			return;

		oDocContent.ClearContent(false);
		for (let nPos = 0, nCount = this.Elements.length; nPos < nCount; ++nPos)
		{
			let oElement = this.Elements[nPos].Element;
			oDocContent.AddToContent(nPos, oElement);
		}

		if (true === isSelect)
		{
			oDocContent.SelectAll();
		}
		else
		{
			oDocContent.RemoveSelection();
			oDocContent.MoveCursorToEndPos();
		}

		oDocContent.SetThisElementCurrent();
	};
	CSelectedContent.prototype.GetPasteHelperElement = function()
	{
		return this.PasteHelper;
	};
	CSelectedContent.prototype.PrepareObjectsForInsert = function()
	{
		let oLogicDocument = this.LogicDocument;

		if (oLogicDocument && oLogicDocument.IsDocumentEditor())
		{
			if (this.NewCommentsGuid)
				this.private_CreateNewCommentsGuid();

			this.private_CopyDocPartNames();

			if (this.CopyComments)
				this.private_CopyComments();
		}
	};
	CSelectedContent.prototype.SetInsertOptionForTable = function(nType)
	{
		this.InsertOptions.Table = nType;
	};
	/**
	 * Converts current content to ParaMath if it possible. Doesn't change current SelectedContent
	 * @returns {?AscCommonWord.ParaMath}
	 * */
	CSelectedContent.prototype.ConvertToMath = function()
	{
		if (!this.CanConvertToMath())
			return null;

		var oParaMath = new AscCommonWord.ParaMath();
		oParaMath.Root.Remove_FromContent(0, oParaMath.Root.GetElementsCount());

		for (let nParaIndex = 0, nParasCount = this.Elements.length; nParaIndex < nParasCount; ++nParaIndex)
		{
			let oParagraph = this.Elements[nParaIndex].Element;
			if (!oParagraph.IsParagraph())
				continue;

			for (var nInParaPos = 0; nInParaPos < oParagraph.GetElementsCount(); ++nInParaPos)
			{
				var oElement = oParagraph.Content[nInParaPos];
				let nType    = oElement.GetType();
				if (para_Run === nType)
				{
					oParaMath.Push(oElement.ToMathRun());
				}
				else if (para_Math === nType)
				{
					oParaMath.Concat(oElement);
				}
			}
		}

		oParaMath.Root.Correct_Content(true);
		return oParaMath;
	};
	/**
	 * Устанавливаем, что сейчас происходит перенос во время рецензирования
	 * @param {boolean} isTrackRevision
	 * @param {string} sMoveId
	 */
	CSelectedContent.prototype.SetMoveTrack = function(isTrackRevision, sMoveId)
	{
		this.TrackRevisions = isTrackRevision;
		this.MoveTrackId    = sMoveId;
	};
	/**
	 * Проверяем собираем ли содержимое для переноса в рецензировании
	 * @returns {boolean}
	 */
	CSelectedContent.prototype.IsMoveTrack = function()
	{
		return this.MoveTrackId !== null;
	};
	/**
	 * @returns {boolean}
	 */
	CSelectedContent.prototype.IsTrackRevisions = function()
	{
		return this.TrackRevisions;
	};
	/**
	 * Добавляем ран, который участвует в переносе
	 * @param {ParaRun} oRun
	 */
	CSelectedContent.prototype.AddRunForMoveTrack = function(oRun)
	{
		this.MoveTrackRuns.push(oRun);
	};
	/**
	 * Устанавливаем есть ли в содержимом текст перенесенный во время рецензирования
	 * @param {boolean} isHave
	 */
	CSelectedContent.prototype.SetMovedParts = function(isHave)
	{
		this.HaveMovedParts = isHave;
	};
	/**
	 * Запрашиваем, есть ли перенесенная во время рецензирования часть
	 * @returns {boolean}
	 */
	CSelectedContent.prototype.IsHaveMovedParts = function()
	{
		return this.HaveMovedParts;
	};
	/**
	 * Запоминаем секцию, на которой закончилось выделение (если оно было в основной части документа)
	 * @param {CSectionPr} oSectPr
	 */
	CSelectedContent.prototype.SetLastSection = function(oSectPr)
	{
		this.LastSection = oSectPr;
	};
	/**
	 * Получаем секцию, на которой закончилось выделение
	 * @returns {null|CSectionPr}
	 */
	CSelectedContent.prototype.GetLastSection = function()
	{
		return this.LastSection;
	};
	/**
	 * Сохранять значения нумерации
	 * @param {boolean} isSave
	 */
	CSelectedContent.prototype.SetSaveNumberingValues = function(isSave)
	{
		this.SaveNumberingValues = isSave;
	};
	/**
	 * Заппрашиваем, нужно ли сохранять расчитанные значения нумерации
	 * @returns {boolean}
	 */
	CSelectedContent.prototype.IsSaveNumberingValues = function()
	{
		return this.SaveNumberingValues;
	};
	/**
	 * По умолчанию мы выводим курсор за пределы вставленных элементов, с данным флагом мы оставляем его
	 * внутри последнего рана
	 * NB: Данный флаг работает только для инлайновой вставки, и не в формулу
	 */
	CSelectedContent.prototype.PlaceCursorInLastInsertedRun = function(isInLast)
	{
		this.CursorInLastRun = undefined === isInLast ? true : !!isInLast;
	};
	/**
	 * Конвертируем элементы в один элемент с простым текстом
	 */
	CSelectedContent.prototype.ConvertToText = function()
	{
		var oParagraph = this.private_CreateParagraph();

		var sText = "";
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oElement = this.Elements[nIndex].Element;
			if (oElement.IsParagraph())
				sText += oElement.GetText({ParaEndToSpace : false});
		}

		var oRun = new ParaRun(oParagraph, null);
		oRun.AddText(sText);
		oParagraph.AddToContent(0, oRun);

		this.Elements.length = 0;
		this.Elements.push(new CSelectedElement(oParagraph, false));
	};
	CSelectedContent.prototype.GetText = function(oPr)
	{
		var sText = "";
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oElement = this.Elements[nIndex].Element;
			if (oElement.IsParagraph())
				sText += oElement.GetText(oPr);
		}
		return sText;
	};
	CSelectedContent.prototype.ConvertToPresentation = function(Parent)
	{
		let Elements = this.Elements.slice(0);
		this.Elements.length = 0;

		for (let nIndex = 0, nCount = Elements.length; nIndex < nCount; ++nIndex)
		{
			let oSelectedElement = Elements[nIndex];
			var oElement = oSelectedElement.Element;
			if (oElement.IsParagraph())
			{
				this.Elements.push(new CSelectedElement(AscFormat.ConvertParagraphToPPTX(oElement, Parent.DrawingDocument, Parent, true, false), oSelectedElement.SelectedAll))
			}
		}
	};
	CSelectedContent.prototype.ConvertToInline = function()
	{
		var oParagraph = this.private_CreateParagraph();

		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oElement = this.Elements[nIndex].Element;
			if (oElement.IsParagraph())
				oParagraph.ConcatContent(oElement.Content);
		}

		this.Elements.length = 0;
		this.Elements.push(new CSelectedElement(oParagraph, false));
	};
	//----------------- Private Area -----------------------------------------------------------------------------------
	CSelectedContent.prototype.private_CollectObjects = function()
	{
		for (let nPos = 0, nCount = this.Elements.length; nPos < nCount; ++nPos)
		{
			let oElement = this.Elements[nPos].Element;

			oElement.Set_DocumentPrev(0 === nPos ? null : this.Elements[nPos - 1].Element);
			oElement.Set_DocumentNext(nPos === nCount - 1 ? null : this.Elements[nPos + 1].Element);
			oElement.ProcessComplexFields();

			let arrParagraphs = oElement.GetAllParagraphs();
			for (let nParaIndex = 0, nParasCount = arrParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
			{
				let oParagraph = arrParagraphs[nParaIndex];
				oParagraph.GetAllDrawingObjects(this.DrawingObjects);
				oParagraph.GetAllComments(this.Comments);
				oParagraph.GetAllMaths(this.Maths);
			}

			if (oElement.IsParagraph() && nCount > 1)
				oElement.CorrectContent();

			if (oElement.IsTable())
				this.Flags |= FLAG_TABLE;

			if (!oElement.IsParagraph())
				this.Flags |= FLAG_NOT_PARAGRAPH;

			oElement.MoveCursorToEndPos(false);
		}

		if (this.Maths.length)
			this.Flags |= FLAG_MATH;

		for (let nPos = 0, nCount = this.DrawingObjects.length; nPos < nCount; ++nPos)
		{
			let oDrawing = this.DrawingObjects[nPos];
			if (oDrawing.IsShape() || oDrawing.IsGroup())
			{
				this.Flags |= FLAG_SHAPE;
				break;
			}
		}
	};
	CSelectedContent.prototype.private_CheckComments = function(oLogicDocument)
	{
		if (!(oLogicDocument instanceof AscCommonWord.CDocument))
			return;

		var mCommentsMarks = {};
		for (var nIndex = 0, nCount = this.Comments.length; nIndex < nCount; ++nIndex)
		{
			var oMark = this.Comments[nIndex].Comment;

			var sId = oMark.GetCommentId();
			if (!mCommentsMarks[sId])
				mCommentsMarks[sId] = {};

			if (oMark.IsCommentStart())
				mCommentsMarks[sId].Start = oMark;
			else
				mCommentsMarks[sId].End   = oMark;
		}

		// Пробегаемся по найденным комментариям и удаляем те, у которых нет начала или конца
		var oCommentsManager = oLogicDocument.GetCommentsManager();
		for (var sId in mCommentsMarks)
		{
			var oEntry = mCommentsMarks[sId];

			var oParagraph = null;
			if (!oEntry.Start && oEntry.End)
				oParagraph = oEntry.End.GetParagraph();
			else if (oEntry.Start && !oEntry.End)
				oParagraph = oEntry.Start.GetParagraph();

			var oComment = oCommentsManager.GetById(sId);
			if ((!oEntry.Start && !oEntry.End) || !oComment)
				delete mCommentsMarks[sId];
			else
				oEntry.Comment = oComment;

			if (oParagraph)
			{
				var bOldValue = oParagraph.DeleteCommentOnRemove;
				oParagraph.DeleteCommentOnRemove = false;
				oParagraph.RemoveCommentMarks(sId);
				oParagraph.DeleteCommentOnRemove = bOldValue;
				delete mCommentsMarks[sId];
			}
		}

		this.CommentsMarks = mCommentsMarks;
	};
	CSelectedContent.prototype.private_CheckTrackMove = function(oLogicDocument)
	{
		if (this.Elements.length <= 0 || !oLogicDocument || !oLogicDocument.TrackMoveId)
			return;

		var isCanMove = !this.IsHaveMovedParts();
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			if (!this.Elements[nIndex].Element.IsParagraph())
			{
				isCanMove = false;
				break;
			}
		}

		if (oLogicDocument.TrackMoveRelocation)
			isCanMove = true;

		if (isCanMove)
		{
			if (oLogicDocument.TrackMoveRelocation)
			{
				var oMarks = oLogicDocument.GetTrackRevisionsManager().GetMoveMarks(oLogicDocument.TrackMoveId);
				if (oMarks)
				{
					oMarks.To.Start.RemoveThisMarkFromDocument();
					oMarks.To.End.RemoveThisMarkFromDocument();
				}
			}

			var oStartElement = this.Elements[0].Element;
			var oEndElement   = this.Elements[this.Elements.length - 1].Element;

			var oStartParagraph = oStartElement.GetFirstParagraph();
			var oEndParagraph   = oEndElement.GetLastParagraph();

			oStartParagraph.AddToContent(0, new CParaRevisionMove(true, false, oLogicDocument.TrackMoveId));

			if (oEndParagraph !== oEndElement || this.Elements[this.Elements.length - 1].SelectedAll)
			{
				var oEndRun = oEndParagraph.GetParaEndRun();
				oEndRun.AddAfterParaEnd(new AscWord.CRunRevisionMove(false, false, oLogicDocument.TrackMoveId));

				var oInfo = new CReviewInfo();
				oInfo.Update();
				oInfo.SetMove(Asc.c_oAscRevisionsMove.MoveTo);
				oEndRun.SetReviewTypeWithInfo(reviewtype_Add, oInfo, false);
			}
			else
			{
				oEndParagraph.AddToContent(oEndParagraph.GetElementsCount(), new CParaRevisionMove(false, false, oLogicDocument.TrackMoveId));
			}

			for (var nIndex = 0, nCount = this.MoveTrackRuns.length; nIndex < nCount; ++nIndex)
			{
				var oRun  = this.MoveTrackRuns[nIndex];
				var oInfo = new CReviewInfo();
				oInfo.Update();
				oInfo.SetMove(Asc.c_oAscRevisionsMove.MoveTo);
				oRun.SetReviewTypeWithInfo(reviewtype_Add, oInfo);
			}
		}
		else
		{
			oLogicDocument.TrackMoveId = null;
		}
	};
	CSelectedContent.prototype.private_CreateNewCommentsGuid = function()
	{
		let oManager = this.LogicDocument.GetCommentsManager();
		for (var Index = 0; Index < this.Comments.length; Index++)
		{
			var comment = oManager.GetById(this.Comments[Index].Comment.CommentId);
			if (comment)
			{
				comment.CreateNewCommentsGuid();
			}
		}
	};
	CSelectedContent.prototype.private_CopyDocPartNames = function()
	{
		var arrCC = [];
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			this.Elements[nIndex].Element.GetAllContentControls(arrCC);
		}

		var oGlossaryDocument = this.LogicDocument.GetGlossaryDocument();
		for (var nIndex = 0, nCount = arrCC.length; nIndex < nCount; ++nIndex)
		{
			var oCC = arrCC[nIndex];

			var sPlaceHolderName = oCC.GetPlaceholder();
			if (sPlaceHolderName)
			{
				var oDocPart = oGlossaryDocument.GetDocPartByName(sPlaceHolderName);
				if (!oDocPart || oGlossaryDocument.IsDefaultDocPart(oDocPart))
					continue;

				var sNewName = oGlossaryDocument.GetNewName();
				oGlossaryDocument.AddDocPart(oDocPart.Copy(sNewName));
				oCC.SetPlaceholder(sNewName);
			}
		}
	};
	CSelectedContent.prototype.private_CopyComments = function()
	{
		let oLogicDocument = this.LogicDocument;

		var oCommentsManager = this.LogicDocument.GetCommentsManager();
		for (let sId in this.CommentsMarks)
		{
			let oEntry = this.CommentsMarks[sId];

			var oNewComment = oEntry.Comment.Copy();
			oCommentsManager.Add(oNewComment);

			var sNewId = oNewComment.GetId();
			oLogicDocument.GetApi().sync_AddComment(sNewId, oNewComment.GetData());
			oEntry.Start.SetCommentId(sNewId);
			oEntry.End.SetCommentId(sNewId);

			oNewComment.SetRangeStart(oEntry.Start.GetId());
			oNewComment.SetRangeEnd(oEntry.End.GetId());
		}
	};
	/**
	 * Проверяем содержимое, которые мы вставляем, в зависимости от места куда оно вставляется
	 * @param oAnchorPos {NearestPos}
	 * @param oDocContent {AscCommonWord.CDocumentContent}
	 */
	CSelectedContent.prototype.private_CheckContentBeforePaste = function(oAnchorPos, oDocContent)
	{
		var oParagraph = oAnchorPos.Paragraph;

		// Если мы вставляем в специальный контент контрол, тогда производим простую вставку текста
		var oParaState = oParagraph.SaveSelectionState();
		oParagraph.RemoveSelection();
		oParagraph.Set_ParaContentPos(oAnchorPos.ContentPos, false, -1, -1, false);
		var arrContentControls = oParagraph.GetSelectedContentControls();
		oParagraph.LoadSelectionState(oParaState);

		for (var nIndex = 0, nCount = arrContentControls.length; nIndex < nCount; ++nIndex)
		{
			if (arrContentControls[nIndex].IsComboBox() || arrContentControls[nIndex].IsDropDownList())
			{
				this.ConvertToText();
				break;
			}
		}

		if (this.IsPresentationContent)
			this.ConvertToPresentation(oDocContent);

		if (this.ForceInline)
			this.ConvertToInline();
	};
	CSelectedContent.prototype.private_AdjustSizeForInlineDrawing = function()
	{
		if (this.MoveDrawing)
			return;

		if (1 === this.DrawingObjects.length && 1 === this.Elements.length)
		{
			let oParaDrawing = this.DrawingObjects[0];
			if (oParaDrawing.IsInline())
			{
				let oElement = this.Elements[0].Element;
				if (oElement.IsParagraph())
				{
					let isAdditionalContent = oElement.CheckRunContent(function(oRun)
					{
						for (let nPos = 0, nCount = oRun.GetElementsCount(); nPos < nCount; ++nPos)
						{
							let oItem = oRun.GetElement(nPos);
							if (oItem && !oItem.IsParaEnd() && !oItem.IsDrawing())
								return true;
						}
						return false;
					});

					if (!isAdditionalContent)
						oParaDrawing.CheckFitToColumn();
				}
			}
		}
	};
	CSelectedContent.prototype.private_CheckInsertSignatures = function()
	{
		var aDrawings        = this.DrawingObjects;
		var nDrawing, oDrawing, oSp;
		var sLastSignatureId = null;
		for (nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing)
		{
			oDrawing = aDrawings[nDrawing];
			oSp      = oDrawing.GraphicObj;
			if (oSp && oSp.signatureLine)
			{
				oSp.setSignature(oSp.signatureLine);
				sLastSignatureId = oSp.signatureLine.id;
			}
		}
		if (sLastSignatureId)
		{
			editor.sendEvent("asc_onAddSignature", sLastSignatureId);
		}
	};
	CSelectedContent.prototype.private_IsInlineInsert = function()
	{
		return (1 === this.Elements.length && !this.Elements[0].SelectedAll && this.Elements[0].Element.IsParagraph() && (!this.Elements[0].Element.IsEmpty() || this.ForceInline));
	};
	CSelectedContent.prototype.private_IsOverwriteTableCells = function()
	{
		let oParagraph = this.Run.GetParagraph();
		if (!oParagraph)
			return false;

		let nDstIndex   = oParagraph.GetIndex();
		let oDocContent = oParagraph.GetParent();
		if (!oDocContent || oParagraph !== oDocContent.GetElement(nDstIndex))
			return false;

		return (Asc.c_oSpecialPasteProps.overwriteCells === this.InsertOptions.Table
			&& 1 === this.Elements.length
			&& this.Elements[0].Element.IsTable()
			&& oDocContent.GetParent() instanceof AscWord.CTableCell);
	};
	CSelectedContent.prototype.private_InsertToMathRun = function()
	{
		let oParaAnchorPos = this.ParaAnchorPos;

		let oMathContent      = oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 2];
		let nInMathContentPos = oParaAnchorPos.NearPos.ContentPos.Data[oParaAnchorPos.Classes.length - 2];

		let oInsertMath = this.ConvertToMath();
		if (oInsertMath)
		{
			let oRun = oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 1];
			let oNewRun = oRun.Split(oParaAnchorPos.NearPos.ContentPos, oParaAnchorPos.Classes.length - 1);
			oMathContent.AddToContent(nInMathContentPos + 1, oNewRun);
			oMathContent.InsertMathContent(oInsertMath.Root, nInMathContentPos + 1, this.Select);
		}
	};
	CSelectedContent.prototype.private_InsertToPictureCC = function()
	{
		let oPictureCC = this.Run.GetParentPictureContentControl()

		var oSrcPicture = null;
		for (var nIndex = 0, nCount = this.DrawingObjects.length; nIndex < nCount; ++nIndex)
		{
			if (this.DrawingObjects[nIndex].IsPicture())
			{
				oSrcPicture = this.DrawingObjects[nIndex].GraphicObj.copy();
				break;
			}
		}

		var arrParaDrawings = oPictureCC.GetAllDrawingObjects();
		if (arrParaDrawings.length > 0 && oSrcPicture)
		{
			oPictureCC.SetShowingPlcHdr(false);
			oSrcPicture.setParent(arrParaDrawings[0]);
			arrParaDrawings[0].Set_GraphicObject(oSrcPicture);

			if (oPictureCC.IsPictureForm())
				oPictureCC.UpdatePictureFormLayout();

			let oLogicDocument = this.LogicDocument;
			if (oLogicDocument)
			{
				oLogicDocument.DrawingObjects.resetSelection();
				oLogicDocument.RemoveSelection();
				oPictureCC.SelectContentControl();

				if (oLogicDocument.IsDocumentEditor() && arrParaDrawings[0].IsPicture())
					oLogicDocument.OnChangeForm(oPictureCC);
			}
		}
	};
	CSelectedContent.prototype.private_InsertToForm = function()
	{
		let oParaAnchorPos = this.ParaAnchorPos;

		let oRun  = this.Run;
		let oForm = oRun.GetParentForm();

		let nInLastClassPos = oParaAnchorPos.NearPos.ContentPos.Data[oParaAnchorPos.Classes.length - 1];

		if (oForm.IsComplexForm())
		{
			this.ConvertToInline();
			return this.private_InsertInline();
		}

		if ((!oForm.IsTextForm() && !oForm.IsComboBox()))
			return;

		let sInsertedText = this.GetText({ParaEndToSpace : false});
		if (!sInsertedText || !sInsertedText.length)
			return;

		var isPlaceHolder = oRun.GetParentForm().IsPlaceHolder();
		if (isPlaceHolder && oRun.GetParent() instanceof CInlineLevelSdt)
		{
			var oInlineLeveLSdt = oRun.GetParent();
			oInlineLeveLSdt.ReplacePlaceHolderWithContent();
			oRun            = oInlineLeveLSdt.GetElement(0);
			nInLastClassPos = 0;
		}

		let nInRunStartPos = nInLastClassPos;
		oRun.State.ContentPos = nInLastClassPos;
		oRun.AddText(sInsertedText, nInLastClassPos);
		let nInRunEndPos = oRun.State.ContentPos;

		let nLastClassLen = oRun.GetElementsCount();
		nInRunStartPos    = Math.min(nLastClassLen, Math.min(nInRunStartPos, nInRunEndPos));
		nInRunEndPos      = Math.min(nLastClassLen, nInRunEndPos);

		if (this.Select)
		{
			oRun.Selection.Use      = true;
			oRun.Selection.StartPos = nInRunStartPos;
			oRun.Selection.EndPos   = nInRunEndPos;
			oRun.State.ContentPos   = nInRunEndPos;
			oRun.SelectThisElement(1, true);
		}
		else
		{
			oRun.SetThisElementCurrent();
			oRun.State.ContentPos = nInRunEndPos;
		}
	};
	CSelectedContent.prototype.private_InsertInline = function()
	{
		let oParaAnchorPos = this.ParaAnchorPos;

		let oInlineLeveLSdt = this.Run.GetParent();
		if (oInlineLeveLSdt instanceof CInlineLevelSdt
			&& (oInlineLeveLSdt.IsPlaceHolder() || oInlineLeveLSdt.IsContentControlTemporary()))
		{
			if (oInlineLeveLSdt.IsContentControlTemporary())
			{
				let oResult = oInlineLeveLSdt.RemoveContentControlWrapper();

				let oSdtParent = oResult.Parent;
				let oSdtPos    = oResult.Pos;
				let oSdtCount  = oResult.Count;

				if (!oSdtParent
					|| oParaAnchorPos.Classes.length < 3
					|| oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 2] !== oInlineLeveLSdt
					|| oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 3] !== oSdtParent)
					return;

				let oRun = new ParaRun(undefined, false);
				oRun.SetPr(oInlineLeveLSdt.GetDefaultTextPr().Copy());

				oSdtParent.RemoveFromContent(oSdtPos, oSdtCount);
				oSdtParent.AddToContent(oSdtPos, oRun);

				oParaAnchorPos.Classes.length--;
				oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 1] = oRun;
				oParaAnchorPos.NearPos.ContentPos.Update(oSdtPos, oParaAnchorPos.Classes.length - 2);
				oParaAnchorPos.NearPos.ContentPos.Update(0, oParaAnchorPos.Classes.length - 1);
			}
			else
			{
				oInlineLeveLSdt.ReplacePlaceHolderWithContent();
				oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 1] = oInlineLeveLSdt.GetElement(0);
				oParaAnchorPos.NearPos.ContentPos.Update(0, oParaAnchorPos.Classes.length - 2);
				oParaAnchorPos.NearPos.ContentPos.Update(0, oParaAnchorPos.Classes.length - 1);
			}
		}

		let oRun    = oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 1];
		let oNewRun = oRun.Split(oParaAnchorPos.NearPos.ContentPos, oParaAnchorPos.Classes.length - 1);

		let oParent      = oParaAnchorPos.Classes[oParaAnchorPos.Classes.length - 2];
		let nInParentPos = oParaAnchorPos.NearPos.ContentPos.Data[oParaAnchorPos.Classes.length - 2];

		oParent.AddToContent(nInParentPos + 1, oNewRun);

		let oParagraph     = this.Elements[0].Element;
		let nElementsCount = oParagraph.Content.length - 1; // Последний ран с para_End не добавляем

		let isSelect = this.Select && !this.MoveDrawing;
		for (let nPos = 0; nPos < nElementsCount; ++nPos)
		{
			let oItem = oParagraph.GetElement(nPos);
			oParent.AddToContent(nInParentPos + 1 + nPos, oItem);

			if (isSelect)
				oItem.SelectAll();
			else
				oItem.RemoveSelection();
		}

		if (this.MoveDrawing)
		{
		}
		else if (isSelect)
		{
			oParent.Selection.Use      = true;
			oParent.Selection.StartPos = nInParentPos + 1;
			oParent.Selection.EndPos   = nInParentPos + 1 + nElementsCount - 1;
			oParent.SelectThisElement(1, true);
		}
		else
		{
			oParent.RemoveSelection();
			oParent.SetThisElementCurrent();
			if (this.CursorInLastRun)
			{
				oParent.SetCurrentPos(nInParentPos + nElementsCount);
				oParent.GetElement(nInParentPos + nElementsCount).MoveCursorToEndPos();
			}
			else
			{
				oParent.SetCurrentPos(nInParentPos + 1 + nElementsCount);
				oParent.GetElement(nInParentPos + nElementsCount).MoveCursorToStartPos();
			}
		}

		if (oParent.CorrectContent)
			oParent.CorrectContent();

		if (this.LogicDocument && this.LogicDocument.IsDocumentEditor())
			this.private_AdjustSizeForInlineDrawing();

		this.private_CheckInsertSignatures();
	};
	CSelectedContent.prototype.private_OverwriteTableCells = function()
	{
		let oTableCell = this.Run.GetParagraph().GetParent().GetParent();
		return oTableCell.InsertTableContent(this.Elements[0].Element);
	};
	CSelectedContent.prototype.private_InsertCommon = function()
	{
		let oParagraph = this.Run.GetParagraph();
		if (!oParagraph)
			return;

		let nDstIndex   = oParagraph.GetIndex();
		let oDocContent = oParagraph.GetParent();
		if (!oDocContent || oParagraph !== oDocContent.GetElement(nDstIndex))
			return;

		oParagraph.RemoveSelection();
		oParagraph.MoveCursorToAnchorPos(this.AnchorPos);

		let oParagraphS, oParagraphE, nInsertPos;
		if (oParagraph.IsCursorAtBegin())
		{
			oParagraphS = null;
			oParagraphE = oParagraph;
			nInsertPos  = nDstIndex;
		}
		else
		{
			oParagraphS = oParagraph;
			oParagraphE = new Paragraph(oParagraph.DrawingDocument, undefined, this.IsPresentationContent);
			oParagraphS.Split(oParagraphE);
			oDocContent.AddToContent(nDstIndex + 1, oParagraphE);
			nInsertPos  = nDstIndex + 1;
		}

		let nSelectionStart = nInsertPos;
		let nStartPos = 0;
		if (oParagraphS
			&& this.Elements[0].Element.IsParagraph()
			&& -1 !== oParagraphS.GetIndex())
		{
			let nParagraphSPos   = oParagraphS.GetIndex();
			let oInsertParagraph = this.Elements[0].Element;
			oInsertParagraph.ConcatBefore(oParagraphS, this.Select ? 1 : 0);

			oDocContent.AddToContent(nParagraphSPos, oInsertParagraph);
			oDocContent.RemoveFromContent(nParagraphSPos + 1, 1);
			nSelectionStart = nParagraphSPos;
			nStartPos++;

			oParagraphS = oInsertParagraph;
		}

		let nEndPos   = this.Elements.length - 1;
		let isConcatE = false;
		if (oParagraphE
			&& this.Elements.length > 1
			&& this.Elements[nEndPos].Element.IsParagraph()
			&& !this.Elements[nEndPos].SelectedAll)
		{
			oParagraphE.ConcatBefore(this.Elements[nEndPos].Element, this.Select ? -1 : 0);
			nEndPos--;
			isConcatE = true;
		}
		else
		{
			oParagraphE.MoveCursorToStartPos();
		}

		for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
		{
			let oElement = this.Elements[nPos].Element;
			oDocContent.AddToContent(nInsertPos++, oElement);

			if (this.Select)
			{
				oElement.SelectAll(1);
			}
			else
			{
				oElement.RemoveSelection();
				oElement.MoveCursorToEndPos();
			}
		}
		let nSelectionEnd = isConcatE ? oParagraphE.GetIndex() : nInsertPos - 1;

		if (this.Select)
		{
			oDocContent.Selection.Use      = true;
			oDocContent.Selection.StartPos = nSelectionStart;
			oDocContent.Selection.EndPos   = nSelectionEnd;
			oDocContent.CurPos.ContentPos  = nSelectionEnd;
			oDocContent.SetThisElementCurrent();
		}
		else
		{
			if (oParagraphS && oParagraphS !== oParagraphE)
			{
				oParagraphS.RemoveSelection();
				oParagraphS.MoveCursorToEndPos();
			}

			oParagraphE.RemoveSelection();
			oDocContent.CurPos.ContentPos = nInsertPos;
			oDocContent.SetThisElementCurrent();
		}

		this.private_CheckInsertSignatures();

		if (isConcatE && oParagraphE)
			this.PasteHelper = oParagraphE;
		else
			this.PasteHelper = this.Elements[this.Elements.length - 1].Element;
	};
	CSelectedContent.prototype.private_GetDrawingDocument = function()
	{
		let _editor = editor;
		if (!_editor && Asc && Asc.editor)
			_editor = Asc.editor;
		
		if (!_editor)
			return null;
		
		return _editor.getDrawingDocument();
	};
	CSelectedContent.prototype.private_CreateParagraph = function()
	{
		return new AscWord.CParagraph(this.private_GetDrawingDocument(), undefined, this.IsPresentationContent);
	};

	/**
	 * @param oElement
	 * @param isSelectedAll
	 * @constructor
	 */
	function CSelectedElement(oElement, isSelectedAll)
	{
		this.Element     = oElement;
		this.SelectedAll = isSelectedAll;
	}


	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CSelectedContent = CSelectedContent;
	window['AscCommonWord'].CSelectedElement = CSelectedElement;

})(window);
