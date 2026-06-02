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

// Import
var CShape = AscFormat.CShape;
var CImageShape = AscFormat.CImageShape;

CImageShape.prototype.Get_Theme = CShape.prototype.Get_Theme;
CImageShape.prototype.Get_ColorMap = CShape.prototype.Get_ColorMap;
CImageShape.prototype.addToRecalculate = CShape.prototype.addToRecalculate;
CImageShape.prototype.convertPixToMM = CShape.prototype.convertPixToMM;
CImageShape.prototype.getHierarchy = CShape.prototype.getHierarchy;
CImageShape.prototype.getParentObjects = CShape.prototype.getParentObjects;
CImageShape.prototype.recalculateTransform = CShape.prototype.recalculateTransform;
CImageShape.prototype.recalculateBounds = CShape.prototype.recalculateBounds;
CImageShape.prototype.getArrayWrapPolygons = CShape.prototype.getArrayWrapPolygons;
CImageShape.prototype.getRotateAngle = CShape.prototype.getRotateAngle;
CImageShape.prototype.recalculateWrapPolygon = CShape.prototype.recalculateWrapPolygon;

CImageShape.prototype.setRecalculateInfo = function()
{
    this.recalcInfo =
    {
        recalculateBrush:          true,
        recalculatePen:            true,
        recalculateTransform:      true,
        recalculateBounds:         true,
        recalculateGeometry:       true,
        recalculateFill:           true,
        recalculateLine:           true,
        recalculateTransparent:    true,
        recalculateWrapPolygon: true
    };


    this.bNeedUpdatePosition = true;

    this.posX = null;
    this.posY = null;

    this.localTransform = new AscCommon.CMatrix();
    this.localTransformText = new AscCommon.CMatrix();
};

CImageShape.prototype.recalcBrush = function()
{
    this.recalcInfo.recalculateBrush = true;
};

CImageShape.prototype.recalcPen = function()
{
    this.recalcInfo.recalculatePen = true;
};
CImageShape.prototype.recalcTransform = function()
{
    this.recalcInfo.recalculateTransform = true;
};
CImageShape.prototype.recalcBounds = function()
{
    this.recalcInfo.recalculateBounds = true;
};
CImageShape.prototype.recalcGeometry = function()
{
    this.recalcInfo.recalculateGeometry = true;
};
CImageShape.prototype.recalcStyle = function()
{
    this.recalcInfo.recalculateStyle = true;
};
CImageShape.prototype.recalcFill = function()
{
    this.recalcInfo.recalculateFill = true;
};
CImageShape.prototype.recalcLine = function()
{
    this.recalcInfo.recalculateLine = true;
};
CImageShape.prototype.recalcTransparent = function()
{
    this.recalcInfo.recalculateTransparent = true;
};
CImageShape.prototype.recalcWrapPolygon = function()
{
    this.recalcInfo.recalculateWrapPolygon = true;
};

CImageShape.prototype.handleUpdatePosition = function()
{
    this.recalcTransform();
    this.recalcBounds();
    this.addToRecalculate();
};
CImageShape.prototype.handleUpdateExtents = function()
{
    this.recalcGeometry();
    this.recalcBounds();
    this.recalcTransform();
    this.recalcWrapPolygon();
    this.addToRecalculate();
};
CImageShape.prototype.handleUpdateRot = function()
{
    this.recalcTransform();
    this.recalcBounds();
    this.recalcWrapPolygon();
    this.addToRecalculate();
};
CImageShape.prototype.handleUpdateFlip = function()
{
    this.recalcTransform();
    this.recalcWrapPolygon();
    this.addToRecalculate();
};
CImageShape.prototype.handleUpdateFill = function()
{
    this.recalcBrush();
    this.recalcFill();
    this.recalcTransparent();
    this.addToRecalculate();
};
CImageShape.prototype.handleUpdateGeometry = function()
{
    this.recalcBounds();
    this.recalcGeometry();
    this.addToRecalculate();
};

CImageShape.prototype.recalculateLocalTransform = CShape.prototype.recalculateLocalTransform;
CImageShape.prototype.convertPixToMM = CShape.prototype.convertPixToMM;
CImageShape.prototype.getCompiledStyle = CShape.prototype.getCompiledStyle;
CImageShape.prototype.getHierarchy = CShape.prototype.getHierarchy;
CImageShape.prototype.getParentObjects = CShape.prototype.getParentObjects;
CImageShape.prototype.recalculate = function ()
{
    if(this.bDeleted)
        return;
    AscFormat.ExecuteNoHistory(function() {
        var bRecalcShadow = this.recalcInfo.recalculateBrush ||
            this.recalcInfo.recalculatePen ||
            this.recalcInfo.recalculateTransform ||
            this.recalcInfo.recalculateGeometry ||
            this.recalcInfo.recalculateBounds;
        if (this.recalcInfo.recalculateBrush) {
            this.recalculateBrush();
            this.recalcInfo.recalculateBrush = false;
        }

        if (this.recalcInfo.recalculatePen) {
            this.recalculatePen();
            this.recalcInfo.recalculatePen = false;
        }
        if (this.recalcInfo.recalculateTransform) {
            this.recalculateTransform();
            this.recalcInfo.recalculateTransform = false;
        }

        if(this.recalcInfo.recalculateGeometry) {
            this.recalculateGeometry();
            this.recalcInfo.recalculateGeometry = false;
        }
        if(this.recalcInfo.recalculateBounds) {
            this.recalculateBounds();
            this.recalcInfo.recalculateBounds = false;
        }
        if(this.recalcInfo.recalculateWrapPolygon)
        {
            this.recalculateWrapPolygon();
            this.recalcInfo.recalculateWrapPolygon = false;
        }
        if(bRecalcShadow)
        {
            this.recalculateShdw();
        }

        this.clearCropObject();
        this.bNeedUpdatePosition = true;
    }, this, []);
};
CImageShape.prototype.canChangeWrapPolygon = CShape.prototype.canChangeWrapPolygon;
CImageShape.prototype.hitInInnerArea = CShape.prototype.hitInInnerArea;
CImageShape.prototype.hitInPath = CShape.prototype.hitInPath;
CImageShape.prototype.check_bounds = CShape.prototype.check_bounds;
CImageShape.prototype.updatePosition = CShape.prototype.updatePosition;
CImageShape.prototype.getDrawingDocument = CShape.prototype.getDrawingDocument;
CImageShape.prototype.getArrayWrapIntervals = CShape.prototype.getArrayWrapIntervals;
CImageShape.prototype.setStartPage = CShape.prototype.setStartPage;
CImageShape.prototype.getRecalcObject = CShape.prototype.getRecalcObject;
CImageShape.prototype.setRecalcObject = CShape.prototype.setRecalcObject;
CImageShape.prototype.checkContentDrawings = function()
{};
CImageShape.prototype.checkShapeChildTransform = function()
{
};
