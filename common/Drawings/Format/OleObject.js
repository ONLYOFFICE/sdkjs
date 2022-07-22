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

(
    /**
     * @param {Window} window
     * @param {undefined} undefined
     */
function (window, undefined) {

    var spreadsheetApplicationId = 'Excel.Sheet.12';

        function COleSize(w, h){
            this.w = w;
            this.h = h;
        }
        COleSize.prototype.Write_ToBinary = function(Writer){
            Writer.WriteLong(this.w);
            Writer.WriteLong(this.h);
        };
        COleSize.prototype.Read_FromBinary = function(Reader){
            this.w = Reader.GetLong();
            this.h = Reader.GetLong();
        };

        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetData] = AscDFH.CChangesDrawingsString;
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetApplicationId] = AscDFH.CChangesDrawingsString;
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetPixSizes] = AscDFH.CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetObjectFile] = AscDFH.CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetOleType] = AscDFH.CChangesDrawingsLong;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetMathObject] = AscDFH.CChangesDrawingsObject;


        AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ChartStyleEntryDefRPr] = AscCommonWord.CTextPr;


		function CChangesOleObjectBinary(Class, Old, New, Color){
            AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
        }

        CChangesOleObjectBinary.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
        CChangesOleObjectBinary.prototype.Type = AscDFH.historyitem_ImageShapeSetBinaryData;
        CChangesOleObjectBinary.prototype.private_SetValue = function(Value)
        {
            this.Class.m_aBinaryData = Value;
        };

        CChangesOleObjectBinary.prototype.WriteToBinary = function(Writer)
        {
            Writer.WriteBool(this.New !== null);
            if(this.New !== null)
            {
                Writer.WriteLong(this.New.length);
                Writer.WriteBuffer(this.New, 0, this.New.length);
            }
        };
        CChangesOleObjectBinary.prototype.ReadFromBinary = function(Reader)
        {
            if(Reader.GetBool())
            {
                var length = Reader.GetLong();
                this.New = Reader.GetBuffer(length);
            }
        };
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetBinaryData] = CChangesOleObjectBinary;

        AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetData] = function(oClass, value){oClass.m_sData = value;};
        AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetApplicationId] = function(oClass, value){oClass.m_sApplicationId = value;};
        AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetPixSizes] = function(oClass, value){
            if(value){
                oClass.m_nPixWidth = value.w;
                oClass.m_nPixHeight = value.h;
            }
        };
        AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ImageShapeSetPixSizes] = COleSize;
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetObjectFile] = function(oClass, value){oClass.m_sObjectFile = value;};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetOleType] = function(oClass, value){oClass.m_nOleType = value;};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetMathObject] = function(oClass, value){oClass.m_oMathObject = value;};

    function COleObject()
    {
		AscFormat.CImageShape.call(this);
        this.m_sData = null;
        this.m_sApplicationId = null;
        this.m_nPixWidth = null;
        this.m_nPixHeight = null;
        this.m_sObjectFile = null;//ole object name in OOX
        this.m_nOleType = null;
        this.m_aBinaryData = null;
        this.m_oMathObject = null;
    }

    COleObject.prototype = Object.create(AscFormat.CImageShape.prototype);
    COleObject.prototype.constructor = COleObject;

    COleObject.prototype.getObjectType = function()
    {
        return AscDFH.historyitem_type_OleObject;
    };
    COleObject.prototype.setData = function(sData)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ImageShapeSetData, this.m_sData, sData));
        this.m_sData = sData;
    };
    COleObject.prototype.setApplicationId = function(sApplicationId)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ImageShapeSetApplicationId, this.m_sApplicationId, sApplicationId));
        this.m_sApplicationId = sApplicationId;
        if (this.m_sApplicationId === spreadsheetApplicationId) {
            this.setOleType(AscCommon.c_oAscOleObjectTypes.spreadsheet);
        }
    };
    COleObject.prototype.setPixSizes = function(nPixWidth, nPixHeight)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ImageShapeSetPixSizes, new COleSize(this.m_nPixWidth, this.m_nPixHeight), new COleSize(nPixWidth, nPixHeight)));
        this.m_nPixWidth = nPixWidth;
        this.m_nPixHeight = nPixHeight;
    };
    COleObject.prototype.setObjectFile = function(sObjectFile)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ImageShapeSetObjectFile, this.m_sObjectFile, sObjectFile));
        this.m_sObjectFile = sObjectFile;
    };
    COleObject.prototype.setOleType = function(nOleType)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_ImageShapeSetOleType, this.m_nOleType, nOleType));
        this.m_nOleType = nOleType;
    };
    COleObject.prototype.setBinaryData = function(aBinaryData)
    {
        AscCommon.History.Add(new CChangesOleObjectBinary(this, this.m_aBinaryData, aBinaryData, false));
        this.m_aBinaryData = aBinaryData;
    };
    COleObject.prototype.setMathObject = function(oMath)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetMathObject, this.m_oMathObject, oMath));
        this.m_oMathObject = oMath;
    };

    COleObject.prototype.canRotate = function () {
        return false;
    };

    COleObject.prototype.copy = function()
    {
        var copy = new COleObject();
        if(this.nvPicPr)
        {
            copy.setNvPicPr(this.nvPicPr.createDuplicate());
        }
        if(this.spPr)
        {
            copy.setSpPr(this.spPr.createDuplicate());
            copy.spPr.setParent(copy);
        }
        if(this.blipFill)
        {
            copy.setBlipFill(this.blipFill.createDuplicate());
        }
        if(this.style)
        {
            copy.setStyle(this.style.createDuplicate());
        }
        copy.setBDeleted(this.bDeleted);
        copy.setData(this.m_sData);
        copy.setApplicationId(this.m_sApplicationId);
        copy.setPixSizes(this.m_nPixWidth, this.m_nPixHeight);
        copy.setObjectFile(this.m_sObjectFile);
        copy.setOleType(this.m_nOleType);
        if(this.m_aBinaryData !== null)
        {
            copy.setBinaryData(this.m_aBinaryData.slice(0, this.m_aBinaryData.length));
        }
        if(this.macro !== null) {
            copy.setMacro(this.macro);
        }
        if(this.textLink !== null) {
            copy.setTextLink(this.textLink);
        }
        if(this.m_oMathObject !== null) {
            copy.setMathObject(this.m_oMathObject.Copy());
        }
        copy.cachedImage = this.getBase64Img();
        copy.cachedPixH = this.cachedPixH;
        copy.cachedPixW = this.cachedPixW;
        return copy;
    };


    COleObject.prototype.handleUpdateExtents = function(){
        AscFormat.CImageShape.prototype.handleUpdateExtents.call(this, []);
    };
    COleObject.prototype.checkTypeCorrect = function(){
        var bCorrectData = false;
        if(this.m_sData) {
            bCorrectData = true;
        }
        else if(this.m_sObjectFile) {
            bCorrectData = true;
        }
        if(!bCorrectData){
            return false;
        }
        if(this.m_sApplicationId === null){
            return false;
        }
        return true;
    };
    COleObject.prototype.replaceToMath = function () {
        if(!this.m_oMathObject) {
            return null;
        }
        if(!this.getDrawingObjectsController) {
            return null;
        }
        var oController = this.getDrawingObjectsController();
        if(!oController) {
            if(this.worksheet) {
                if(Asc && Asc.editor && Asc.editor.wb && Asc.editor.wbModel) {
                    Asc.editor.wb.getWorksheet(Asc.editor.wbModel.getWorksheetIndexByName(this.worksheet.getName()));
                    oController = this.getDrawingObjectsController();
                    if(!oController) {
                        return null;
                    }
                }
                else {
                    return null;
                }
            }
            else {
                return null;
            }
        }
        var oShape = new AscFormat.CShape();
        oShape.setBDeleted(false);
        if(this.worksheet)
            oShape.setWorksheet(this.worksheet);
        if(this.parent) {
            oShape.setParent(this.parent);
        }
        if(this.group) {
            oShape.setGroup(this.group);
        }
        var oSpPr = new AscFormat.CSpPr();
        var oXfrm = new AscFormat.CXfrm();
        oXfrm.setOffX(0);
        oXfrm.setOffY(0);
        oXfrm.setExtX(1828800/36000);
        oXfrm.setExtY(1828800/36000);
        oSpPr.setXfrm(oXfrm);
        oXfrm.setParent(oSpPr);
        oSpPr.setFill(AscFormat.CreateNoFillUniFill());
        oSpPr.setLn(AscFormat.CreateNoFillLine());
        oSpPr.setGeometry(AscFormat.CreateGeometry("rect"));
        oShape.setSpPr(oSpPr);
        oSpPr.setParent(oShape);
        oShape.createTextBody();
        var oContent = oShape.getDocContent();
        this.m_oMathObject.Correct_AfterConvertFromEquation();
        var oParagraph = oContent.Content[0];
        oParagraph.AddToContent(1, this.m_oMathObject);
        oParagraph.Correct_Content();
        oParagraph.SetParagraphAlign(AscCommon.align_Center);
        var oBodyPr = oShape.getBodyPr().createDuplicate();
        oBodyPr.rot = 0;
        oBodyPr.spcFirstLastPara = false;
        oBodyPr.vertOverflow = AscFormat.nOTOwerflow;
        oBodyPr.horzOverflow = AscFormat.nOTOwerflow;
        oBodyPr.vert = AscFormat.nVertTThorz;
        oBodyPr.wrap = AscFormat.nTWTNone;
        oBodyPr.lIns = 2.54;
        oBodyPr.tIns = 1.27;
        oBodyPr.rIns = 2.54;
        oBodyPr.bIns = 1.27;
        oBodyPr.numCol = 1;
        oBodyPr.spcCol = 0;
        oBodyPr.rtlCol = 0;
        oBodyPr.fromWordArt = false;
        oBodyPr.anchor = 1;
        oBodyPr.anchorCtr = false;
        oBodyPr.forceAA = false;
        oBodyPr.compatLnSpc = true;
        oBodyPr.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry("textNoShape");
        oBodyPr.textFit = new AscFormat.CTextFit();
        oBodyPr.textFit.type = AscFormat.text_fit_Auto;
        oShape.txBody.setBodyPr(oBodyPr);
        var nPos;
        if(this.group) {
            nPos = this.group.getPosInSpTree(this.Id);
            if(null !== nPos && nPos > -1) {
                this.group.removeFromSpTreeByPos(nPos);
                this.group.addToSpTree(nPos, oShape);
            }
        }
        else {
            nPos = this.deleteDrawingBase();
            if(null !== nPos && nPos > -1) {
                oShape.addToDrawingObjects(nPos);
            }
        }
        oShape.checkExtentsByDocContent(true, false);
        var fXc = this.x + this.extX / 2;
        var fYc = this.y + this.extY / 2;
        oShape.spPr.xfrm.setOffX(fXc - oShape.extX / 2);
        oShape.spPr.xfrm.setOffY(fYc - oShape.extY / 2);
        oShape.spPr.xfrm.setRot(this.rot);
        oShape.checkDrawingBaseCoords();
        if(this.group) {
            this.group.updateCoordinatesAfterInternalResize();
        }
        if(this.selected) {
            var nSelectStartPage = this.selectStartPage;
            this.deselect(oController);
            oShape.select(oController, nSelectStartPage);
        }
        return oShape;
    };

    COleObject.prototype.editExternal = function(Data, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight) {
        if(typeof Data === "string" && this.m_sData !== Data) {
            this.setData(Data);
        }
        if (Data instanceof Uint8Array) {
            this.setBinaryData(Data)
        }
        if(typeof sImageUrl  === "string" &&
            (!this.blipFill || this.blipFill.RasterImageId !== sImageUrl)) {
            var _blipFill           = new AscFormat.CBlipFill();
            _blipFill.RasterImageId = sImageUrl;
            this.setBlipFill(_blipFill);
        }
        if(this.m_nPixWidth !== nPixWidth || this.m_nPixHeight !== nPixHeight) {
            this.setPixSizes(nPixWidth, nPixHeight);
        }
        var fWidth_ = fWidth;
        var fHeight_ = fHeight;
        if(!AscFormat.isRealNumber(fWidth_) || !AscFormat.isRealNumber(fHeight_)) {
            var oImagePr = new Asc.asc_CImgProperty();
            oImagePr.asc_putImageUrl(sImageUrl);
            var oApi = editor || Asc.editor;
            var oSize = oImagePr.asc_getOriginSize(oApi);
            if(oSize.IsCorrect) {
                fWidth_ = oSize.Width;
                fHeight_ = oSize.Height;
            }
        }
        if(AscFormat.isRealNumber(fWidth_) && AscFormat.isRealNumber(fHeight_)) {
            var oXfrm = this.spPr && this.spPr.xfrm;
            if(oXfrm) {
                if(!AscFormat.fApproxEqual(oXfrm.extX, fWidth_) ||
                    !AscFormat.fApproxEqual(oXfrm.extY, fHeight_)) {
                    oXfrm.setExtX(fWidth_);
                    oXfrm.setExtY(fHeight_);
                    if(!this.group) {
                        if(this.drawingBase) {
                            this.checkDrawingBaseCoords();
                        }
                        if(this.parent && this.parent.CheckWH) {
                            this.parent.CheckWH();
                        }
                    }
                }
            }
        }
    };
    COleObject.prototype.GetAllOleObjects = function(sPluginId, arrObjects) {
        if(typeof sPluginId === "string" && sPluginId.length > 0) {
            if(sPluginId === this.m_sApplicationId) {
                arrObjects.push(this);
            }
        }
        else {
            arrObjects.push(this);
        }
    };
    COleObject.prototype.getDataObject = function() {
        var dWidth = 0, dHeight = 0;
        if(this.parent && this.parent.Extent) {
            var oExtent = this.parent.Extent;
            dWidth = oExtent.W;
            dHeight = oExtent.H;
        }
        else {
            if(this.spPr && this.spPr.xfrm) {
                var oXfrm = this.spPr.xfrm;
                dWidth = oXfrm.extX;
                dHeight = oXfrm.extY;
            }
        }
        var oBlipFill = this.blipFill;
        var oParaDrawing;
        var oParaDrawingChild = this;
        if(this.group) {
            oParaDrawingChild = this.getMainGroup();
        }
        if(AscCommonWord.ParaDrawing &&
            oParaDrawingChild.parent &&
            oParaDrawingChild.parent instanceof AscCommonWord.ParaDrawing) {
            oParaDrawing = oParaDrawingChild.parent;
        }
        return {
            "Data": this.m_sData,
            "ApplicationId": this.m_sApplicationId,
            "ImageData": oBlipFill ? oBlipFill.getBase64RasterImageId(false) : "",
            "Width": dWidth,
            "Height": dHeight,
            "WidthPix": this.m_nPixWidth,
            "HeightPix": this.m_nPixHeight,
            "InternalId": this.Id,
            "ParaDrawingId": oParaDrawing ? oParaDrawing.Id : ""
        }
    };

    COleObject.prototype.canEditTableOleObject = function(bReturnOle) {
        var canEdit = this.m_aBinaryData &&
          (this.m_nOleType === AscCommon.c_oAscOleObjectTypes.spreadsheet || this.m_sApplicationId === spreadsheetApplicationId);
        if (bReturnOle) {
            return canEdit ? this : null;
        }
        return !!canEdit;
    };

    function asc_putBinaryDataToFrameFromTableOleObject(oleObject)
    {
        if (oleObject instanceof AscFormat.COleObject) {
            var dataSize = oleObject.m_aBinaryData.length;
            var data = AscCommon.Base64.encode(oleObject.m_aBinaryData);
            return {
                "binary": "XLSY;v2;" + dataSize  + ";" + data,
                "isFromSheetEditor": !!oleObject.worksheet,
            };
        }
        return {
            "binary": null
        };
    };
    window['Asc'] = window['Asc'] || {};
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].COleObject = COleObject;
    window['Asc'].asc_putBinaryDataToFrameFromTableOleObject = window['Asc']['asc_putBinaryDataToFrameFromTableOleObject'] = asc_putBinaryDataToFrameFromTableOleObject;

})(window);
