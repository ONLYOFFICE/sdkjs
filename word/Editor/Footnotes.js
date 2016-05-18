/*
 *
 * (c) Copyright Ascensio System Limited 2010-2016
 *
 * This program is freeware. You can redistribute it and/or modify it under the terms of the GNU
 * General Public License (GPL) version 3 as published by the Free Software Foundation (https://www.gnu.org/copyleft/gpl.html).
 * In accordance with Section 7(a) of the GNU GPL its Section 15 shall be amended to the effect that
 * Ascensio System SIA expressly excludes the warranty of non-infringement of any third-party rights.
 *
 * THIS PROGRAM IS DISTRIBUTED WITHOUT ANY WARRANTY; WITHOUT EVEN THE IMPLIED WARRANTY OF MERCHANTABILITY OR
 * FITNESS FOR A PARTICULAR PURPOSE. For more details, see GNU GPL at https://www.gnu.org/copyleft/gpl.html
 *
 * You can contact Ascensio System SIA by email at sales@onlyoffice.com
 *
 * The interactive user interfaces in modified source and object code versions of ONLYOFFICE must display
 * Appropriate Legal Notices, as required under Section 5 of the GNU GPL version 3.
 *
 * Pursuant to Section 7  3(b) of the GNU GPL you must retain the original ONLYOFFICE logo which contains
 * relevant author attributions when distributing the software. If the display of the logo in its graphic
 * form is not reasonably feasible for technical reasons, you must include the words "Powered by ONLYOFFICE"
 * in every copy of the program you distribute.
 * Pursuant to Section 7  3(e) we decline to grant you any rights under trademark law for use of our trademarks.
 *
 */

"use strict";
/**
 * User: Ilja.Kirillov
 * Date: 04.05.2016
 * Time: 17:00
 */

/**
 * Класс, работающий со сносками документа.
 * @param {CDocument} LogicDocument - Ссылка на главный документ.
 * @constructor
 * @extends {CDocumentControllerBase}
 */
function CFootnotesController(LogicDocument)
{
    CFootnotesController.superclass.constructor.call(this, LogicDocument);

    this.Id = LogicDocument.Get_IdCounter().Get_NewId();

    this.Footnote = {}; // Список всех сносок с ключом - Id.
    this.Pages    = [];

    // Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
    LogicDocument.Get_TableId().Add(this, this.Id);
}

AscCommon.extendClass(CFootnotesController, CDocumentControllerBase);

/**
 * Получаем Id данного класса.
 */
CFootnotesController.prototype.Get_Id = function()
{
    return this.Id;
};
/**
 * Создаем новую сноску.
 * @returns {CFootEndnote}
 */
CFootnotesController.prototype.Create_Footnote = function()
{
    var NewFootnote = new CFootEndnote(this);
    this.Footnote[NewFootnote.Get_Id()] = NewFootnote;
    return NewFootnote;
};
/**
 * Пересчитываем сноски на заданной странице.
 */
CFootnotesController.prototype.Recalculate = function(nPageIndex, X, XLimit, Y, YLimit)
{
    if (!this.Pages[nPageIndex])
        this.Pages[nPageIndex] = new CFootEndnotePage();

    // Мы пересчет начинаем с 0, потом просто делаем сдвиг, через функцию Shift.

    var CurY = Y;

    for (var nIndex = 0; nIndex < this.Pages[nPageIndex].Elements.length; ++nIndex)
    {
        var Footnote = this.Pages[nPageIndex].Elements[nIndex];
        Footnote.Reset(X, CurY, XLimit, 10000);

        var CurPage      = 0;
        var RecalcResult = recalcresult2_NextPage;
        while (recalcresult2_End != RecalcResult)
            RecalcResult = Footnote.Recalculate_Page(CurPage++, true);

        var Bounds = Footnote.Get_PageBounds(0);
        CurY += Bounds.Bottom - Bounds.Top;
    }
};
/**
 * Получаем суммарную высоту, занимаемую сносками на заданной странице.
 * @param {number} nPageIndex
 */
CFootnotesController.prototype.Get_Height = function(nPageIndex)
{
    return 0;
    var nHeight = 0;
    for (var nIndex = 0; nIndex < this.Pages[nPageIndex].Elements.length; ++nIndex)
    {
        var Footnote = this.Pages[nPageIndex].Elements[nIndex];
        var Bounds = Footnote.Get_PageBounds(0);
        nHeight += Bounds.Bottom - Bounds.Top;
    }
    return nHeight;
};
/**
 * Добавляем заданную сноску на страницу для пересчета.
 * @param {number} nPageIndex
 * @param {CFootEndnote} oFootnote
 */
CFootnotesController.prototype.Add_FootnoteOnPage = function(nPageIndex, oFootnote)
{
    if (!this.Pages[nPageIndex])
        this.Pages[nPageIndex] = new CFootEndnotePage();

    this.Pages[nPageIndex].Elements.push(oFootnote);
};
/**
 * Проверяем, используется заданная сноска в документе.
 * @param {string} sFootnoteId
 * @returns  {boolean}
 */
CFootnotesController.prototype.Is_UseInDocument = function(sFootnoteId)
{
    // TODO: Надо бы еще проверить, если ли в документе ссылка на данную сноску.
    for (var sId in this.Footnote)
    {
        if (sId === sFootnoteId)
            return true;
    }

    return false;
};

function CFootEndnotePage()
{
    this.X      = 0;
    this.Y      = 0;
    this.XLimit = 0;
    this.YLimit = 0;

    this.Elements = [];

}





