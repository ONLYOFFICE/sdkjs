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
 * Date: 29.04.2016
 * Time: 18:09
 */

/**
 * Интерфейс для классов, которые работают с колонтитулами/автофигурами/сносками
 * @param {CDocument} LogicDocument - Ссылка на главный документ.
 * @constructor
 */
function CDocumentControllerBase(LogicDocument)
{
    this.LogicDocument   = LogicDocument;
    this.DrawingDocument = LogicDocument.Get_DrawingDocument();
}

/**
 * Получаем ссылку на CDocument.
 * @returns {CDocument}
 */
CDocumentControllerBase.prototype.Get_LogicDocument = function()
{
    return this.LogicDocument;
};
/**
 * Получаем ссылку на CDrawingDocument.
 * @returns {CDrawingDocument}
 */
CDocumentControllerBase.prototype.Get_DrawingDocument = function()
{
    return this.DrawingDocument;
};
/**
 * Является ли данный класс контроллером для колонтитулов.
 * @param {boolean} bReturnHdrFtr Если true, тогда возвращаем ссылку на колонтитул.
 * @returns {boolean | CHeaderFooter}
 */
CDocumentControllerBase.prototype.Is_HdrFtr = function(bReturnHdrFtr)
{
    if (bReturnHdrFtr)
        return null;

    return false;
};
/**
 * Является ли данный класс контроллером для автофигур.
 * @param {boolean} bReturnShape Если true, тогда возвращаем ссылку на автофигуру.
 * @returns {boolean | CShape}
 */
CDocumentControllerBase.prototype.Is_HdrFtr = function(bReturnShape)
{
    if (bReturnShape)
        return null;

    return false;
};
/**
 * Получаем абосолютный номер страницы по относительному.
 * @param {number} CurPage
 * @returns {number}
 */
CDocumentControllerBase.prototype.Get_AbsolutePage = function(CurPage)
{
    return CurPage;
};
/**
 * Получаем абосолютный номер колноки по относительному номеру страницы.
 * @param {number} CurPage
 * @returns {number}
 */
CDocumentControllerBase.prototype.Get_AbsoluteColumn = function(CurPage)
{
    return 0;
};
/**
 * Проверяем, является ли данный класс верхним, по отношению к другим классам DocumentContent, Document
 * @param {boolean} bReturnTopDocument Если true, тогда возвращаем ссылку на документ.
 * @returns {СDocument | CDocumentContent | null}
 */
CDocumentControllerBase.prototype.Is_TopDocument = function(bReturnTopDocument)
{
    if (bReturnTopDocument)
        return null;

    return false;
};
/**
 * Получаем ссылку на объект, работающий с нумерацией.
 * @returns {CNumbering}
 */
CDocumentControllerBase.prototype.Get_Numbering = function()
{
    return this.LogicDocument.Get_Numbering();
};
/**
 * Получаем ссылку на объект, работающий со стилями.
 * @returns {CStyles}
 */
CDocumentControllerBase.prototype.Get_Styles = function()
{
    return this.LogicDocument.Get_Styles();
};
/**
 * Получаем ссылку на табличный стиль для параграфа.
 * @returns {null}
 */
CDocumentControllerBase.prototype.Get_TableStyleForPara = function()
{
    return null;
};
/**
 * Получаем ссылку на стиль автофигур для параграфа.
 * @returns {null}
 */
CDocumentControllerBase.prototype.Get_ShapeStyleForPara = function()
{
    return null;
};
/**
 * Запрашиваем цвет заливки.
 * @returns {null}
 */
CDocumentControllerBase.prototype.Get_TextBackGroundColor = function()
{
    return undefined;
};
/**
 * Является ли данный класс ячейкой.
 * @returns {false}
 */
CDocumentControllerBase.prototype.Is_Cell = function()
{
    return false;
};
/**
 * Получаем настройки темы для документа.
 */
CDocumentControllerBase.prototype.Get_Theme = function()
{
    return this.LogicDocument.Get_Theme();
};
/**
 * Запрашиваем информацию о конце пересчета предыдущего элемента.
 * @param CurElement
 * @returns {null}
 */
CDocumentControllerBase.prototype.Get_PrevElementEndInfo = function(CurElement)
{
    return null;
};
