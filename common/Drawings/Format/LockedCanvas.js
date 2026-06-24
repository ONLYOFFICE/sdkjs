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

(function (undefined){
    function CLockedCanvas() {
        AscFormat.CGroupShape.call(this);
    }
    CLockedCanvas.prototype = Object.create(AscFormat.CGroupShape.prototype);
    CLockedCanvas.prototype.constructor = CLockedCanvas;
    CLockedCanvas.prototype.getObjectType = function(){
        return AscDFH.historyitem_type_LockedCanvas;
    };
    CLockedCanvas.prototype.canRotate = function () {
        return false;
    };

    CLockedCanvas.prototype.recalculateBounds = function () {
        var tr = this.localTransform;
        var arr_p_x = [];
        var arr_p_y = [];
        arr_p_x.push(tr.TransformPointX(0,0));
        arr_p_y.push(tr.TransformPointY(0,0));
        arr_p_x.push(tr.TransformPointX(this.extX,0));
        arr_p_y.push(tr.TransformPointY(this.extX,0));
        arr_p_x.push(tr.TransformPointX(this.extX,this.extY));
        arr_p_y.push(tr.TransformPointY(this.extX,this.extY));
        arr_p_x.push(tr.TransformPointX(0,this.extY));
        arr_p_y.push(tr.TransformPointY(0,this.extY));
        this.bounds.x = Math.min.apply(Math, arr_p_x);
        this.bounds.y = Math.min.apply(Math, arr_p_y);
        this.bounds.l = this.bounds.x;
        this.bounds.t = this.bounds.y;
        this.bounds.r = Math.max.apply(Math, arr_p_x);
        this.bounds.b = Math.max.apply(Math, arr_p_y);
        this.bounds.w = this.bounds.r - this.bounds.l;
        this.bounds.h = this.bounds.b - this.bounds.t;
    };


    CLockedCanvas.prototype.draw = function(graphics) {
        // if(this.parent){
        //     graphics.m_oCoordTransform.sx *= (this.parent.Extent.W / this.spPr.xfrm.extX);
        //     graphics.m_oCoordTransform.sy *= (this.parent.Extent.H / this.spPr.xfrm.extY);
        // }
        AscFormat.CGroupShape.prototype.draw.call(this, graphics);
    };

    CLockedCanvas.prototype.hitInPath = function () {
        return false;
    };
    CLockedCanvas.prototype.copy = function (oPr) {
        var copy = new CLockedCanvas();
        return this.copy2(copy, oPr);
    };

    CLockedCanvas.prototype.hitInInnerArea = function (x, y) {
        var invert_transform = this.getInvertTransform();
        if(!invert_transform) {
            return false;
        }
        var x_t = invert_transform.TransformPointX(x, y);
        var y_t = invert_transform.TransformPointY(x, y);
        return x_t > 0 && x_t < this.extX && y_t > 0 && y_t < this.extY;
    };

    CLockedCanvas.prototype.hit = function (x, y) {
        return this.hitInInnerArea(x, y);
    };


    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].CLockedCanvas = CLockedCanvas;
})();
