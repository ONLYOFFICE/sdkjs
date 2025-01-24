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

(
    /**
     * @param {Window} window
     * @param {undefined} undefined
     */
function (window, undefined) {

    const SPREADSHEET_APPLICATION_ID = 'Excel.Sheet.12';
    const BINARY_PART_HISTORY_LIMIT = 1048576;

    function CChangesDrawingsImageId(Class, Type, OldPr, NewPr) {
        AscDFH.CChangesDrawingsString.call(this, Class, Type, OldPr, NewPr);
        this.FromLoad = false;
    }
    CChangesDrawingsImageId.prototype = Object.create(AscDFH.CChangesDrawingsString.prototype);
    CChangesDrawingsImageId.prototype.constructor = CChangesDrawingsImageId;
    CChangesDrawingsImageId.prototype.ReadFromBinary = function (reader) {
        this.FromLoad = true;
        AscDFH.CChangesDrawingsString.prototype.ReadFromBinary.call(this, reader);
    };

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

        function CChangesStartOleObjectBinary(Class, Old, New, Color) {
            AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
        }
        CChangesStartOleObjectBinary.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
        CChangesStartOleObjectBinary.prototype.constructor = CChangesStartOleObjectBinary;

        CChangesStartOleObjectBinary.prototype.Type = AscDFH.historyitem_ImageShapeSetStartBinaryData;

        CChangesStartOleObjectBinary.prototype.Undo = function () {
            if (!this.Class.partsOfBinaryData) {
                return this.Redo();
            }

            let lenOfAllBinaryData = 0;
            for (let i = 0; i < this.Class.partsOfBinaryData.length; i += 1) {
                lenOfAllBinaryData += this.Class.partsOfBinaryData[i].length;
            }

            this.Class.m_aBinaryData = new Uint8Array(lenOfAllBinaryData);
            let indexOfInsert = 0;
            for (let i = this.Class.partsOfBinaryData.length - 1; i >= 0; i -= 1) {
                const partOfBinaryData = this.Class.partsOfBinaryData[i];
                for (let j = 0; j < partOfBinaryData.length; j += 1) {
                    this.Class.m_aBinaryData[indexOfInsert] = partOfBinaryData[j];
                    indexOfInsert += 1;
                }
            }
            delete this.Class.partsOfBinaryData;
        }

        CChangesStartOleObjectBinary.prototype.Redo = function () {
            if (this.Class.partsOfBinaryData) {
                return this.Undo()
            }
            this.Class.partsOfBinaryData = [];
        }

        function CChangesPartOleObjectBinary(Class, Old, New, Color) {
            AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
        }
        CChangesPartOleObjectBinary.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
        CChangesPartOleObjectBinary.prototype.constructor = CChangesPartOleObjectBinary;
        CChangesPartOleObjectBinary.prototype.Type = AscDFH.historyitem_ImageShapeSetPartBinaryData;





        CChangesPartOleObjectBinary.prototype.private_SetValue = function (oPr) {
            if (oPr.length) {
                this.Class.partsOfBinaryData.push(oPr);
            }
        }

        CChangesPartOleObjectBinary.prototype.WriteToBinary = function(Writer)
        {
            Writer.WriteLong(this.Old.length);
            Writer.WriteBuffer(this.Old, 0, this.Old.length);

            Writer.WriteLong(this.New.length);
            Writer.WriteBuffer(this.New, 0, this.New.length);
        };
        CChangesPartOleObjectBinary.prototype.ReadFromBinary = function(Reader)
        {
            let length = Reader.GetLong();
            this.Old = new Uint8Array(Reader.GetBuffer(length));

            length = Reader.GetLong();
            this.New = new Uint8Array(Reader.GetBuffer(length));
        };

        function CChangesEndOleObjectBinary(Class, Old, New, Color) {
            AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
        }
        CChangesEndOleObjectBinary.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
        CChangesEndOleObjectBinary.prototype.constructor = CChangesEndOleObjectBinary;
        CChangesEndOleObjectBinary.prototype.Type = AscDFH.historyitem_ImageShapeSetEndBinaryData;

        CChangesEndOleObjectBinary.prototype.Undo = function () {
            if (this.Class.partsOfBinaryData) {
                return this.Redo();
            }
            this.Class.partsOfBinaryData = [];
        }

        CChangesEndOleObjectBinary.prototype.Redo = function () {
            if (!this.Class.partsOfBinaryData) {
                return this.Undo();
            }
            let lenOfAllBinaryData = 0;
            for (let i = 0; i < this.Class.partsOfBinaryData.length; i += 1) {
                lenOfAllBinaryData += this.Class.partsOfBinaryData[i].length;
            }
            this.Class.m_aBinaryData = new Uint8Array(lenOfAllBinaryData);

            let indexOfInsert = 0;
            for (let i = 0; i < this.Class.partsOfBinaryData.length; i += 1) {
                const partOfBinaryData = this.Class.partsOfBinaryData[i];
                for (let j = 0; j < partOfBinaryData.length; j += 1) {
                    this.Class.m_aBinaryData[indexOfInsert] = partOfBinaryData[j];
                    indexOfInsert += 1;
                }
            }
            delete this.Class.partsOfBinaryData;
        }

        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetData] = AscDFH.CChangesDrawingsString;
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetApplicationId] = AscDFH.CChangesDrawingsString;
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetPixSizes] = AscDFH.CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetObjectFile] = AscDFH.CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetDataLink] = AscDFH.CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetOleType] = AscDFH.CChangesDrawingsLong;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetMathObject] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetDrawAspect] = AscDFH.CChangesDrawingsLong;
        AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ChartStyleEntryDefRPr] = AscCommonWord.CTextPr;
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetStartBinaryData] = CChangesStartOleObjectBinary;
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetPartBinaryData] = CChangesPartOleObjectBinary;
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetEndBinaryData] = CChangesEndOleObjectBinary;
        AscDFH.changesFactory[AscDFH.historyitem_ImageShapeLoadImagesfromContent] = CChangesDrawingsImageId;

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
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetDataLink] = function(oClass, value){oClass.m_sDataLink = value;};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetOleType] = function(oClass, value){oClass.m_nOleType = value;};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetMathObject] = function(oClass, value){oClass.m_oMathObject = value;};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetDrawAspect] = function(oClass, value){oClass.m_nDrawAspect = value;};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeLoadImagesfromContent] = function(oClass, sValue, bFromLoad) {
            if (bFromLoad) {
                if(AscCommon.CollaborativeEditing) {
                    if(sValue && sValue.length > 0) {
                        AscCommon.CollaborativeEditing.Add_NewImage(sValue);
                    }
                }
            }
        };


        let EOLEDrawAspect =
            {
                oledrawaspectContent: 0,
                oledrawaspectIcon: 1
            };

    function COleObject()
    {
		AscFormat.CImageShape.call(this);
        this.m_sData = null;
        this.m_sApplicationId = null;
        this.m_nPixWidth = null;
        this.m_nPixHeight = null;
        this.m_sObjectFile = null;//ole object name in OOX
        this.m_nOleType = null;
        this.m_aBinaryData = new Uint8Array(0);
        this.m_oMathObject = null;
        this.m_sDataLink = null;
        this.m_nDrawAspect = AscFormat.EOLEDrawAspect.oledrawaspectContent;
        this.m_bShowAsIcon = false;
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
    COleObject.prototype.setDrawAspect = function (oPr) {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_ImageShapeSetDrawAspect, this.m_nDrawAspect, oPr));
        this.m_nDrawAspect = oPr;
    };
    COleObject.prototype.loadImagesFromContent = function (arrImagesId) {
        for (let i = 0; i < arrImagesId.length; i += 1) {
            AscCommon.History.Add(new CChangesDrawingsImageId(this, AscDFH.historyitem_ImageShapeLoadImagesfromContent, '', arrImagesId[i]));
        }
    };
    COleObject.prototype.setApplicationId = function(sApplicationId)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ImageShapeSetApplicationId, this.m_sApplicationId, sApplicationId));
        this.m_sApplicationId = sApplicationId;
        if (this.m_sApplicationId === SPREADSHEET_APPLICATION_ID) {
            this.setOleType(AscCommon.c_oAscOleObjectTypes.spreadsheet);
        }
    };
    COleObject.prototype.setPixSizes = function(nPixWidth, nPixHeight)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ImageShapeSetPixSizes, new COleSize(this.m_nPixWidth, this.m_nPixHeight), new COleSize(nPixWidth, nPixHeight)));
        this.m_nPixWidth = nPixWidth;
        this.m_nPixHeight = nPixHeight;
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
    COleObject.prototype.setDataLink = function(sDataLink)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ImageShapeSetDataLink, this.m_sDataLink, sDataLink));
        this.m_sDataLink = sDataLink;
    };
    COleObject.prototype.setOleType = function(nOleType)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_ImageShapeSetOleType, this.m_nOleType, nOleType));
        this.m_nOleType = nOleType;
    };
    COleObject.prototype.setBinaryData = function(aBinaryData)
    {
        const maxLen = aBinaryData.length > this.m_aBinaryData.length ? aBinaryData.length : this.m_aBinaryData.length;
        const oldParts = [];
        const newParts = [];
        const amountOfParts = Math.ceil(maxLen / BINARY_PART_HISTORY_LIMIT);
        for (let i = 0; i < amountOfParts; i += 1) {
            oldParts.push(this.m_aBinaryData.slice(i * BINARY_PART_HISTORY_LIMIT, (i + 1) * BINARY_PART_HISTORY_LIMIT));
            newParts.push(aBinaryData.slice(i * BINARY_PART_HISTORY_LIMIT, (i + 1) * BINARY_PART_HISTORY_LIMIT));
        }
        AscCommon.History.Add(new CChangesStartOleObjectBinary(this, null, null, false));
        for (let i = 0; i < amountOfParts; i += 1) {
            AscCommon.History.Add(new CChangesPartOleObjectBinary(this, oldParts[i], newParts[i], false));
        }
        AscCommon.History.Add(new CChangesEndOleObjectBinary(this, null, null, false));
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

    COleObject.prototype.copy = function(oPr)
    {
        const copy = new COleObject();
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
        if(this.m_aBinaryData.length !== 0)
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
        if(!oPr || false !== oPr.cacheImage) {
            copy.cachedImage = this.getBase64Img();
            copy.cachedPixH = this.cachedPixH;
            copy.cachedPixW = this.cachedPixW;
        }
        return copy;
    };


    COleObject.prototype.handleUpdateExtents = function(){
        AscFormat.CImageShape.prototype.handleUpdateExtents.call(this, []);
    };
    COleObject.prototype.checkTypeCorrect = function(){
        let bCorrectData = false;
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
        let oController = this.getDrawingObjectsController();
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
        const oShape = new AscFormat.CShape();
        oShape.setBDeleted(false);
        if(this.worksheet)
            oShape.setWorksheet(this.worksheet);
        if(this.parent) {
            oShape.setParent(this.parent);
        }
        if(this.group) {
            oShape.setGroup(this.group);
        }
        const oSpPr = new AscFormat.CSpPr();
        const oXfrm = new AscFormat.CXfrm();
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
        const oContent = oShape.getDocContent();
        this.m_oMathObject.Correct_AfterConvertFromEquation();
        const oParagraph = oContent.Content[0];
        oParagraph.AddToContent(1, this.m_oMathObject);
        oParagraph.Correct_Content();
        oParagraph.SetParagraphAlign(AscCommon.align_Center);
        const oBodyPr = oShape.getBodyPr().createDuplicate();
        oBodyPr.rot = 0;
        oBodyPr.spcFirstLastPara = false;
        oBodyPr.vertOverflow = AscFormat.nVOTOverflow;
        oBodyPr.horzOverflow = AscFormat.nHOTOverflow;
        oBodyPr.vert = AscFormat.nVertTThorz;
        oBodyPr.wrap = AscFormat.nTWTNone;
        oBodyPr.setDefaultInsets();
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
        if(this.group) {
            const nPos = this.group.getPosInSpTree(this.Id);
            if(null !== nPos && nPos > -1) {
                this.group.removeFromSpTreeByPos(nPos);
                this.group.addToSpTree(nPos, oShape);
            }
        }
        else {
            const nPos = this.deleteDrawingBase();
            if(null !== nPos && nPos > -1) {
                oShape.addToDrawingObjects(nPos);
            }
        }
        oShape.checkExtentsByDocContent(true, false);
        const fXc = this.x + this.extX / 2;
        const fYc = this.y + this.extY / 2;
        oShape.spPr.xfrm.setOffX(fXc - oShape.extX / 2);
        oShape.spPr.xfrm.setOffY(fYc - oShape.extY / 2);
        oShape.spPr.xfrm.setRot(this.rot);
        oShape.checkDrawingBaseCoords();
        if(this.group) {
            this.group.updateCoordinatesAfterInternalResize();
        }
        if(this.selected) {
            const nSelectStartPage = this.selectStartPage;
            this.deselect(oController);
            oShape.select(oController, nSelectStartPage);
        }
        return oShape;
    };

    COleObject.prototype.editExternal = function(Data, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory) {
        if (arrImagesForAddToHistory) {
            this.loadImagesFromContent(arrImagesForAddToHistory);
        }

        if(typeof Data === "string" && this.m_sData !== Data) {
            this.setData(Data);
        }
        if (Data instanceof Uint8Array) {
            this.setBinaryData(Data)
        }
        if (this.m_nDrawAspect === AscFormat.EOLEDrawAspect.oledrawaspectContent && !this.m_bShowAsIcon) {
            if(typeof sImageUrl  === "string" &&
                (!this.blipFill || this.blipFill.RasterImageId !== sImageUrl)) {
                const _blipFill = this.blipFill ? this.blipFill.createDuplicate() : new AscFormat.CBlipFill();
                _blipFill.setRasterImageId(sImageUrl);
                this.setBlipFill(_blipFill);
            }
            if(this.m_nPixWidth !== nPixWidth || this.m_nPixHeight !== nPixHeight) {
                this.setPixSizes(nPixWidth, nPixHeight);
            }
            let fWidth_ = fWidth;
            let fHeight_ = fHeight;
            if(!AscFormat.isRealNumber(fWidth_) || !AscFormat.isRealNumber(fHeight_)) {
                const oImagePr = new Asc.asc_CImgProperty();
                oImagePr.asc_putImageUrl(sImageUrl);
                const oApi = editor || Asc.editor;
                const oSize = oImagePr.asc_getOriginSize(oApi);
                if(oSize.IsCorrect) {
                    fWidth_ = oSize.Width;
                    fHeight_ = oSize.Height;
                }
            }
            if(AscFormat.isRealNumber(fWidth_) && AscFormat.isRealNumber(fHeight_)) {
                const oXfrm = this.spPr && this.spPr.xfrm;
                if(oXfrm) {
                    if (oXfrm.isZero()) {
                        oXfrm.setOffX(this.x || 0);
                        oXfrm.setOffY(this.y || 0);
                    }
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
        let dWidth = 0, dHeight = 0;
        if(this.parent && this.parent.Extent) {
            const oExtent = this.parent.Extent;
            dWidth = oExtent.W;
            dHeight = oExtent.H;
        }
        else {
            if(this.spPr && this.spPr.xfrm) {
                const oXfrm = this.spPr.xfrm;
                dWidth = oXfrm.extX;
                dHeight = oXfrm.extY;
            }
        }
        const oBlipFill = this.blipFill;

        let oParaDrawingChild = this;
        let oParaDrawing;
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
    COleObject.prototype.convertDataObjectToPluginData = function(oDataObject) {
        var pluginData = new Asc.CPluginData();
        pluginData.setAttribute("data", oDataObject["Data"]);
        pluginData.setAttribute("guid", oDataObject["ApplicationId"]);
        pluginData.setAttribute("width", oDataObject["Width"]);
        pluginData.setAttribute("height", oDataObject["Height"]);
        pluginData.setAttribute("widthPix", oDataObject["WidthPix"]);
        pluginData.setAttribute("heightPix", oDataObject["HeightPix"]);
        pluginData.setAttribute("objectId", oDataObject["InternalId"]);
        return pluginData;
    };
    COleObject.prototype.runPlugin = function() {
        let pluginData = this.getPluginData();
        Asc.editor.asc_pluginRun(pluginData.getAttribute("guid"), 0, pluginData);
    };
    COleObject.prototype.callPluginOnResize = function() {
        let pluginData = this.getPluginData();
        let width = this.getXfrmExtX();
        let height = this.getXfrmExtX();
        if(width && height) {
            pluginData.setAttribute("width", width);
            pluginData.setAttribute("height", height);
        }
        Asc.editor.asc_pluginResize(pluginData);
    };
    COleObject.prototype.getPluginData = function() {
        let oDataObject = this.getDataObject();
        return this.convertDataObjectToPluginData(oDataObject);
    };
    COleObject.prototype.getPluginDataObject = function() {
        let oPluginData = this.getPluginData();
        return oPluginData.getDataObject();
    };

    COleObject.prototype.canEditTableOleObject = function(bReturnOle) {
        const canEdit = this.m_aBinaryData.length !== 0 &&
          (this.m_nOleType === AscCommon.c_oAscOleObjectTypes.spreadsheet || this.m_sApplicationId === SPREADSHEET_APPLICATION_ID);
        if (bReturnOle) {
            return canEdit ? this : null;
        }
        return !!canEdit;
    };

    COleObject.prototype.fillDataLink = function(sId, reader) {
        if(sId) {
            let rel = reader.rels.getRelationship(sId);
            if (rel) {
                if ("Internal" === rel.targetMode && rel.targetFullName) {
                    this.setDataLink(rel.targetFullName.slice(1));
                }
            }
        }
    };

    COleObject.prototype.getTypeName = function () {
        return AscCommon.translateManager.getValue("Object");
    };

    function asc_putBinaryDataToFrameFromTableOleObject(oOleObject)
    {
        if (oOleObject instanceof AscFormat.COleObject) {
            const oApi = Asc.editor || editor;
            if (!oApi.isOpenedChartFrame) {
                oApi.isOleEditor = true;
                oApi.asc_onOpenChartFrame();
                const oController = oApi.getGraphicController();
                if (oController) {
                    AscFormat.ExecuteNoHistory(function () {
                        oController.checkSelectedObjectsAndCallback(function () {}, [], false);
                    }, this, []);
                }
            }
            const nDataSize = oOleObject.m_aBinaryData.length;
            const sData = AscCommon.Base64.encode(oOleObject.m_aBinaryData);
            const nImageWidth = oOleObject.extX * AscCommon.g_dKoef_mm_to_pix;
            const nImageHeight = oOleObject.extY * AscCommon.g_dKoef_mm_to_pix;
            const documentImageUrls = AscCommon.g_oDocumentUrls.urls;

            return {
                "binary": "XLSY;v2;" + nDataSize  + ";" + sData,
                "isFromSheetEditor": !!oOleObject.worksheet,
                "imageWidth": nImageWidth,
                "imageHeight": nImageHeight,
                "documentImageUrls": documentImageUrls
            };
        }
        return {
            "binary": null
        };
    }
    window['Asc'] = window['Asc'] || {};
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].COleObject = COleObject;
    window['Asc'].asc_putBinaryDataToFrameFromTableOleObject = window['Asc']['asc_putBinaryDataToFrameFromTableOleObject'] = asc_putBinaryDataToFrameFromTableOleObject;
    window['AscFormat'].EOLEDrawAspect = EOLEDrawAspect;

})(window);
