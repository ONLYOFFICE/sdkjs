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

// ---------------------------------------------------------------
function CAscSlideTransition()
{
    AscFormat.CBaseNoIdObject.call(this);
    this.TransitionType     = undefined;
    this.TransitionOption   = undefined;
    this.TransitionDuration = undefined;

    this.SlideAdvanceOnMouseClick   = undefined;
    this.SlideAdvanceAfter          = undefined;
    this.SlideAdvanceDuration       = undefined;
    this.ShowLoop                   = undefined;
}
AscFormat.InitClass(CAscSlideTransition, AscFormat.CBaseNoIdObject, 0);

CAscSlideTransition.prototype.put_TransitionType = function(v) { this.TransitionType = v; };
CAscSlideTransition.prototype.get_TransitionType = function() { return this.TransitionType; };
CAscSlideTransition.prototype.put_TransitionOption = function(v) { this.TransitionOption = v; };
CAscSlideTransition.prototype.get_TransitionOption = function() { return this.TransitionOption; };
CAscSlideTransition.prototype.put_TransitionDuration = function(v) { this.TransitionDuration = v; };
CAscSlideTransition.prototype.get_TransitionDuration = function() { return this.TransitionDuration; };

CAscSlideTransition.prototype.put_SlideAdvanceOnMouseClick = function(v) { this.SlideAdvanceOnMouseClick = v; };
CAscSlideTransition.prototype.get_SlideAdvanceOnMouseClick = function() { return this.SlideAdvanceOnMouseClick; };
CAscSlideTransition.prototype.put_SlideAdvanceAfter = function(v) { this.SlideAdvanceAfter = v; };
CAscSlideTransition.prototype.get_SlideAdvanceAfter = function() { return this.SlideAdvanceAfter; };
CAscSlideTransition.prototype.put_SlideAdvanceDuration = function(v) { this.SlideAdvanceDuration = v; };
CAscSlideTransition.prototype.get_SlideAdvanceDuration = function() { return this.SlideAdvanceDuration; };
CAscSlideTransition.prototype.put_ShowLoop = function(v) { this.ShowLoop = v; };
CAscSlideTransition.prototype.get_ShowLoop = function() { return this.ShowLoop; };
CAscSlideTransition.prototype.applyProps = function(v)
{
    if (undefined !== v.TransitionType && null !== v.TransitionType)
        this.TransitionType = v.TransitionType;
    if (undefined !== v.TransitionOption && null !== v.TransitionOption)
        this.TransitionOption = v.TransitionOption;
    if (undefined !== v.TransitionDuration && null !== v.TransitionDuration)
        this.TransitionDuration = v.TransitionDuration;

    if (undefined !== v.SlideAdvanceOnMouseClick && null !== v.SlideAdvanceOnMouseClick)
        this.SlideAdvanceOnMouseClick = v.SlideAdvanceOnMouseClick;
    if (undefined !== v.SlideAdvanceAfter && null !== v.SlideAdvanceAfter)
        this.SlideAdvanceAfter = v.SlideAdvanceAfter;
    if (undefined !== v.SlideAdvanceDuration && null !== v.SlideAdvanceDuration)
        this.SlideAdvanceDuration = v.SlideAdvanceDuration;
    if (undefined !== v.ShowLoop && null !== v.ShowLoop)
        this.ShowLoop = v.ShowLoop;
};
CAscSlideTransition.prototype.createDuplicate = function(v)
{
    var _slideT = new Asc.CAscSlideTransition();

    _slideT.TransitionType     = this.TransitionType;
    _slideT.TransitionOption   = this.TransitionOption;
    _slideT.TransitionDuration = this.TransitionDuration;

    _slideT.SlideAdvanceOnMouseClick   = this.SlideAdvanceOnMouseClick;
    _slideT.SlideAdvanceAfter          = this.SlideAdvanceAfter;
    _slideT.SlideAdvanceDuration       = this.SlideAdvanceDuration;
    _slideT.ShowLoop                   = this.ShowLoop;

    return _slideT;
};
CAscSlideTransition.prototype.makeDuplicate = function(_slideT)
{
    this.fillObject(_slideT);
};
CAscSlideTransition.prototype.fillObject = function(_slideT)
{
    if (!_slideT)
        return;

    _slideT.TransitionType     = this.TransitionType;
    _slideT.TransitionOption   = this.TransitionOption;
    _slideT.TransitionDuration = this.TransitionDuration;

    _slideT.SlideAdvanceOnMouseClick   = this.SlideAdvanceOnMouseClick;
    _slideT.SlideAdvanceAfter          = this.SlideAdvanceAfter;
    _slideT.SlideAdvanceDuration       = this.SlideAdvanceDuration;
    _slideT.ShowLoop                   = this.ShowLoop;
};

CAscSlideTransition.prototype.createDuplicate = function()
{
    let oCopy = new CAscSlideTransition();
    this.fillObject(oCopy);
    return oCopy;
};
CAscSlideTransition.prototype.setUndefinedOptions = function()
{
    this.TransitionType     = undefined;
    this.TransitionOption   = undefined;
    this.TransitionDuration = undefined;

    this.SlideAdvanceOnMouseClick   = undefined;
    this.SlideAdvanceAfter          = undefined;
    this.SlideAdvanceDuration       = undefined;
    this.ShowLoop                   = undefined;
};
CAscSlideTransition.prototype.setDefaultParams = function()
{
    this.TransitionType     = c_oAscSlideTransitionTypes.None;
    this.TransitionOption   = -1;
    this.TransitionDuration = 500;

    this.SlideAdvanceOnMouseClick   = true;
    this.SlideAdvanceAfter          = false;
    this.SlideAdvanceDuration       = 10000;
    this.ShowLoop                   = true;
};
CAscSlideTransition.prototype.getTypeAndOption = function()
{
    let oResult = {Type: null, Option: null};
    if(this.TransitionType === c_oAscSlideTransitionTypes.Random)
    {
        let aTypes = [
            Asc.c_oAscSlideTransitionTypes.Fade,
            Asc.c_oAscSlideTransitionTypes.Push,
            Asc.c_oAscSlideTransitionTypes.Wipe,
            Asc.c_oAscSlideTransitionTypes.Split,
            Asc.c_oAscSlideTransitionTypes.UnCover,
            Asc.c_oAscSlideTransitionTypes.Cover,
            Asc.c_oAscSlideTransitionTypes.Clock,
            Asc.c_oAscSlideTransitionTypes.Zoom,
            Asc.c_oAscSlideTransitionTypes.Morph,
            Asc.c_oAscSlideTransitionTypes.Cut,
            Asc.c_oAscSlideTransitionTypes.Blinds,
            Asc.c_oAscSlideTransitionTypes.Checker,
            Asc.c_oAscSlideTransitionTypes.Comb,
            Asc.c_oAscSlideTransitionTypes.Circle,
            Asc.c_oAscSlideTransitionTypes.Diamond,
            Asc.c_oAscSlideTransitionTypes.Dissolve,
            Asc.c_oAscSlideTransitionTypes.Plus,
            Asc.c_oAscSlideTransitionTypes.RandomBar,
            Asc.c_oAscSlideTransitionTypes.BoxZoom,
            Asc.c_oAscSlideTransitionTypes.Vortex,
            Asc.c_oAscSlideTransitionTypes.Ripple,
            Asc.c_oAscSlideTransitionTypes.Honeycomb,
            Asc.c_oAscSlideTransitionTypes.Doors,
            Asc.c_oAscSlideTransitionTypes.Window,
            Asc.c_oAscSlideTransitionTypes.Pan,
            Asc.c_oAscSlideTransitionTypes.Glitter,
            Asc.c_oAscSlideTransitionTypes.Flash,
            Asc.c_oAscSlideTransitionTypes.Shred,
            Asc.c_oAscSlideTransitionTypes.Reveal
        ];
        let nType = AscCommon.getArrayRandomElement(aTypes);
        let nOption = null;
        switch (nType) {
            case Asc.c_oAscSlideTransitionTypes.Fade: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Fade_Smoothly,
                    c_oAscSlideTransitionParams.Fade_Through_Black
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Push:
            case Asc.c_oAscSlideTransitionTypes.Wipe:
            case Asc.c_oAscSlideTransitionTypes.UnCover:
            case Asc.c_oAscSlideTransitionTypes.Cover: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Param_Left,
                    c_oAscSlideTransitionParams.Param_Top,
                    c_oAscSlideTransitionParams.Param_Right,
                    c_oAscSlideTransitionParams.Param_Bottom,
                    c_oAscSlideTransitionParams.Param_TopLeft,
                    c_oAscSlideTransitionParams.Param_TopRight,
                    c_oAscSlideTransitionParams.Param_BottomLeft,
                    c_oAscSlideTransitionParams.Param_BottomRight
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Split: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Split_VerticalIn,
                    c_oAscSlideTransitionParams.Split_VerticalOut,
                    c_oAscSlideTransitionParams.Split_HorizontalIn,
                    c_oAscSlideTransitionParams.Split_HorizontalOut
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Clock: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Clock_Clockwise,
                    c_oAscSlideTransitionParams.Clock_Counterclockwise,
                    c_oAscSlideTransitionParams.Clock_Wedge
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Zoom: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Zoom_In,
                    c_oAscSlideTransitionParams.Zoom_Out,
                    c_oAscSlideTransitionParams.Zoom_AndRotate
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Morph: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Morph_Objects,
                    c_oAscSlideTransitionParams.Morph_Words,
                    c_oAscSlideTransitionParams.Morph_Letters
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Cut: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Cut_Default,
                    c_oAscSlideTransitionParams.Cut_ThroughBlack
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Blinds:
            case Asc.c_oAscSlideTransitionTypes.Checker:
            case Asc.c_oAscSlideTransitionTypes.Comb:
            case Asc.c_oAscSlideTransitionTypes.RandomBar: {
                const horzVertMap = {};
                horzVertMap[Asc.c_oAscSlideTransitionTypes.Blinds] = [c_oAscSlideTransitionParams.Blinds_Horizontal, c_oAscSlideTransitionParams.Blinds_Vertical];
                horzVertMap[Asc.c_oAscSlideTransitionTypes.Checker] = [c_oAscSlideTransitionParams.Checker_Horizontal, c_oAscSlideTransitionParams.Checker_Vertical];
                horzVertMap[Asc.c_oAscSlideTransitionTypes.Comb] = [c_oAscSlideTransitionParams.Comb_Horizontal, c_oAscSlideTransitionParams.Comb_Vertical];
                horzVertMap[Asc.c_oAscSlideTransitionTypes.RandomBar] = [c_oAscSlideTransitionParams.RandomBar_Horizontal, c_oAscSlideTransitionParams.RandomBar_Vertical];
                nOption = AscCommon.getArrayRandomElement(horzVertMap[nType]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Circle:    { nOption = c_oAscSlideTransitionParams.Circle_Default; break; }
            case Asc.c_oAscSlideTransitionTypes.Diamond:   { nOption = c_oAscSlideTransitionParams.Diamond_Default; break; }
            case Asc.c_oAscSlideTransitionTypes.Dissolve:  { nOption = c_oAscSlideTransitionParams.Dissolve_Default; break; }
            case Asc.c_oAscSlideTransitionTypes.Plus:      { nOption = c_oAscSlideTransitionParams.Plus_Default; break; }
            case Asc.c_oAscSlideTransitionTypes.BoxZoom: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.BoxZoom_In,
                    c_oAscSlideTransitionParams.BoxZoom_Out
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Vortex: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Vortex_Left, c_oAscSlideTransitionParams.Vortex_Right,
                    c_oAscSlideTransitionParams.Vortex_Up, c_oAscSlideTransitionParams.Vortex_Down
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Ripple: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Ripple_Center, c_oAscSlideTransitionParams.Ripple_LeftUp,
                    c_oAscSlideTransitionParams.Ripple_RightUp, c_oAscSlideTransitionParams.Ripple_LeftDown,
                    c_oAscSlideTransitionParams.Ripple_RightDown
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Honeycomb: { nOption = c_oAscSlideTransitionParams.Honeycomb_Default; break; }
            case Asc.c_oAscSlideTransitionTypes.Flash:     { nOption = c_oAscSlideTransitionParams.Flash_Default; break; }
            case Asc.c_oAscSlideTransitionTypes.Doors: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Doors_Horizontal, c_oAscSlideTransitionParams.Doors_Vertical
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Window: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Window_Horizontal, c_oAscSlideTransitionParams.Window_Vertical
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Pan: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Pan_Left, c_oAscSlideTransitionParams.Pan_Right,
                    c_oAscSlideTransitionParams.Pan_Up, c_oAscSlideTransitionParams.Pan_Down
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Glitter: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Glitter_Left_Diamond, c_oAscSlideTransitionParams.Glitter_Right_Diamond,
                    c_oAscSlideTransitionParams.Glitter_Up_Diamond, c_oAscSlideTransitionParams.Glitter_Down_Diamond,
                    c_oAscSlideTransitionParams.Glitter_Left_Hexagon, c_oAscSlideTransitionParams.Glitter_Right_Hexagon,
                    c_oAscSlideTransitionParams.Glitter_Up_Hexagon, c_oAscSlideTransitionParams.Glitter_Down_Hexagon
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Shred: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Shred_StripIn, c_oAscSlideTransitionParams.Shred_StripOut,
                    c_oAscSlideTransitionParams.Shred_RectangleIn, c_oAscSlideTransitionParams.Shred_RectangleOut
                ]);
                break;
            }
            case Asc.c_oAscSlideTransitionTypes.Reveal: {
                nOption = AscCommon.getArrayRandomElement([
                    c_oAscSlideTransitionParams.Reveal_SmoothLeft, c_oAscSlideTransitionParams.Reveal_SmoothRight,
                    c_oAscSlideTransitionParams.Reveal_BlackLeft, c_oAscSlideTransitionParams.Reveal_BlackRight
                ]);
                break;
            }
        }
        oResult.Type = nType;
        oResult.Option = nOption;
    }
    else
    {
        oResult.Type = this.TransitionType;
        oResult.Option = this.TransitionOption;
    }
    return oResult;
};

CAscSlideTransition.prototype.Write_ToBinary = function(w)
{
    w.WriteBool(AscFormat.isRealNumber(this.TransitionType));
    if(AscFormat.isRealNumber(this.TransitionType))
        w.WriteLong(this.TransitionType);

    w.WriteBool(AscFormat.isRealNumber(this.TransitionOption));
    if(AscFormat.isRealNumber(this.TransitionOption))
        w.WriteLong(this.TransitionOption);

    w.WriteBool(AscFormat.isRealNumber(this.TransitionDuration));
    if(AscFormat.isRealNumber(this.TransitionDuration))
        w.WriteLong(this.TransitionDuration);


    w.WriteBool(AscFormat.isRealBool(this.SlideAdvanceOnMouseClick));
    if(AscFormat.isRealBool(this.SlideAdvanceOnMouseClick))
        w.WriteBool(this.SlideAdvanceOnMouseClick);

    w.WriteBool(AscFormat.isRealBool(this.SlideAdvanceAfter));
    if(AscFormat.isRealBool(this.SlideAdvanceAfter))
        w.WriteBool(this.SlideAdvanceAfter);

    w.WriteBool(AscFormat.isRealNumber(this.SlideAdvanceDuration));
    if(AscFormat.isRealNumber(this.SlideAdvanceDuration))
        w.WriteLong(this.SlideAdvanceDuration);
    AscFormat.writeBool(w, this.ShowLoop);
};

CAscSlideTransition.prototype.Read_FromBinary = function(r)
{

    if(r.GetBool())
        this.TransitionType = r.GetLong();

    if(r.GetBool())
        this.TransitionOption = r.GetLong();


    if(r.GetBool())
        this.TransitionDuration = r.GetLong();


    if(r.GetBool())
        this.SlideAdvanceOnMouseClick = r.GetBool();


    if(r.GetBool())
        this.SlideAdvanceAfter = r.GetBool();

    if(r.GetBool())
        this.SlideAdvanceDuration = r.GetLong();
    this.ShowLoop = AscFormat.readBool(r);
};

CAscSlideTransition.prototype.ToArray = function()
{
    var _ret = [];
    _ret.push(this.TransitionType);
    _ret.push(this.TransitionOption);
    _ret.push(this.TransitionDuration);

    _ret.push(this.SlideAdvanceOnMouseClick);
    _ret.push(this.SlideAdvanceAfter);
    _ret.push(this.SlideAdvanceDuration);
    _ret.push(this.ShowLoop);
    return _ret;
};

CAscSlideTransition.prototype.parseXmlParameters = function (_type, _paramNames, _paramValues) {
	let typeMatched = false;
    if (_paramNames.length === _paramValues.length && typeof _type === "string" && _type.length > 0)
    {
        var _len = _paramNames.length;
        if ("p:fade" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Fade;
            this.TransitionOption = c_oAscSlideTransitionParams.Fade_Smoothly;

            if (1 === _len && _paramNames[0] === "thruBlk" && _paramValues[0] === "1")
            {
                this.TransitionOption = c_oAscSlideTransitionParams.Fade_Through_Black;
            }
        }
        else if ("p:push" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Push;
            this.TransitionOption = c_oAscSlideTransitionParams.Param_Bottom;

            if (1 === _len && _paramNames[0] === "dir")
            {
                if ("l" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Right;
                if ("r" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Left;
                if ("d" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Top;
            }
        }
        else if ("p:wipe" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Wipe;
            this.TransitionOption = c_oAscSlideTransitionParams.Param_Right;

            if (1 === _len && _paramNames[0] === "dir")
            {
                if ("u" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Bottom;
                if ("r" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Left;
                if ("d" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Top;
            }
        }
        else if ("p:strips" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Wipe;
            this.TransitionOption = c_oAscSlideTransitionParams.Param_TopRight;

            if (1 === _len && _paramNames[0] === "dir")
            {
                if ("rd" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_TopLeft;
                if ("ru" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_BottomLeft;
                if ("lu" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_BottomRight;
            }
        }
        else if ("p:cover" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Cover;
            this.TransitionOption = c_oAscSlideTransitionParams.Param_Right;

            if (1 === _len && _paramNames[0] === "dir")
            {
                if ("u" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Bottom;
                if ("r" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Left;
                if ("d" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Top;
                if ("rd" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_TopLeft;
                if ("ru" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_BottomLeft;
                if ("lu" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_BottomRight;
                if ("ld" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_TopRight;
            }
        }
        else if ("p:pull" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.UnCover;
            this.TransitionOption = c_oAscSlideTransitionParams.Param_Right;

            if (1 === _len && _paramNames[0] === "dir")
            {
                if ("u" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Bottom;
                if ("r" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Left;
                if ("d" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_Top;
                if ("rd" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_TopLeft;
                if ("ru" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_BottomLeft;
                if ("lu" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_BottomRight;
                if ("ld" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Param_TopRight;
            }
        }
        else if ("p:split" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Split;

            var _is_vert = true;
            var _is_out = true;

            for (var i = 0; i < _len; i++)
            {
                if (_paramNames[i] === "orient")
                {
                    _is_vert = (_paramValues[i] === "vert") ? true : false;
                }
                else if (_paramNames[i] === "dir")
                {
                    _is_out = (_paramValues[i] === "out") ? true : false;
                }
            }

            if (_is_vert)
            {
                if (_is_out)
                    this.TransitionOption = c_oAscSlideTransitionParams.Split_VerticalOut;
                else
                    this.TransitionOption = c_oAscSlideTransitionParams.Split_VerticalIn;
            }
            else
            {
                if (_is_out)
                    this.TransitionOption = c_oAscSlideTransitionParams.Split_HorizontalOut;
                else
                    this.TransitionOption = c_oAscSlideTransitionParams.Split_HorizontalIn;
            }
        }
        else if ("p:wheel" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Clock;
            this.TransitionOption = c_oAscSlideTransitionParams.Clock_Clockwise;
        }
        else if ("p14:wheelReverse" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Clock;
            this.TransitionOption = c_oAscSlideTransitionParams.Clock_Counterclockwise;
        }
        else if ("p:wedge" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Clock;
            this.TransitionOption = c_oAscSlideTransitionParams.Clock_Wedge;
        }
        else if ("p14:warp" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Zoom;
            this.TransitionOption = c_oAscSlideTransitionParams.Zoom_Out;

            if (1 === _len && _paramNames[0] === "dir")
            {
                if ("in" === _paramValues[0])
                    this.TransitionOption = c_oAscSlideTransitionParams.Zoom_In;
            }
        }
        else if ("p:newsflash" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Zoom;
            this.TransitionOption = c_oAscSlideTransitionParams.Zoom_AndRotate;
        }
        else if ("p159:morph" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Morph;
            this.TransitionOption = c_oAscSlideTransitionParams.Morph_Objects;
            if(_paramNames[0] === "option")
            {
                if ("byObject" === _paramValues[0])
                {
                    this.TransitionOption = c_oAscSlideTransitionParams.Morph_Objects;
                }
                else if("byWord" === _paramValues[0])
                {
                    this.TransitionOption = c_oAscSlideTransitionParams.Morph_Words;
                }
                else if("byChar" === _paramValues[0])
                {
                    this.TransitionOption = c_oAscSlideTransitionParams.Morph_Letters;
                }
            }
        }
        else if("p:random" === _type)
        {
			typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Random;
        }
        else if ("p:cut" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Cut;
            this.TransitionOption = c_oAscSlideTransitionParams.Cut_Default;
            if (1 === _len && _paramNames[0] === "thruBlk" && (_paramValues[0] === "1" || _paramValues[0] === "true"))
            {
                this.TransitionOption = c_oAscSlideTransitionParams.Cut_ThroughBlack;
            }
        }
        else if ("p:blinds" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Blinds;
            this.TransitionOption = c_oAscSlideTransitionParams.Blinds_Horizontal;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "vert" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Blinds_Vertical;
            }
        }
        else if ("p:checker" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Checker;
            this.TransitionOption = c_oAscSlideTransitionParams.Checker_Horizontal;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "vert" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Checker_Vertical;
            }
        }
        else if ("p:comb" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Comb;
            this.TransitionOption = c_oAscSlideTransitionParams.Comb_Horizontal;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "vert" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Comb_Vertical;
            }
        }
        else if ("p:circle" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Circle;
            this.TransitionOption = c_oAscSlideTransitionParams.Circle_Default;
        }
        else if ("p:diamond" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Diamond;
            this.TransitionOption = c_oAscSlideTransitionParams.Diamond_Default;
        }
        else if ("p:dissolve" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Dissolve;
            this.TransitionOption = c_oAscSlideTransitionParams.Dissolve_Default;
        }
        else if ("p:plus" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Plus;
            this.TransitionOption = c_oAscSlideTransitionParams.Plus_Default;
        }
        else if ("p:randomBar" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.RandomBar;
            this.TransitionOption = c_oAscSlideTransitionParams.RandomBar_Horizontal;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "vert" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.RandomBar_Vertical;
            }
        }
        else if ("p:zoom" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.BoxZoom;
            this.TransitionOption = c_oAscSlideTransitionParams.BoxZoom_Out;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "in" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.BoxZoom_In;
            }
        }
        else if ("p14:vortex" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Vortex;
            this.TransitionOption = c_oAscSlideTransitionParams.Vortex_Left;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir") {
                    if ("r" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Vortex_Right;
                    else if ("u" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Vortex_Up;
                    else if ("d" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Vortex_Down;
                }
            }
        }
        else if ("p14:switch" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Switch;
            this.TransitionOption = c_oAscSlideTransitionParams.Switch_Left;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "r" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Switch_Right;
            }
        }
        else if ("p14:flip" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Flip;
            this.TransitionOption = c_oAscSlideTransitionParams.Flip_Left;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "r" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Flip_Right;
            }
        }
        else if ("p14:ripple" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Ripple;
            this.TransitionOption = c_oAscSlideTransitionParams.Ripple_Center;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir") {
                    if ("lu" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Ripple_LeftUp;
                    else if ("ru" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Ripple_RightUp;
                    else if ("ld" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Ripple_LeftDown;
                    else if ("rd" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Ripple_RightDown;
                }
            }
        }
        else if ("p14:honeycomb" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Honeycomb;
            this.TransitionOption = c_oAscSlideTransitionParams.Honeycomb_Default;
        }
        else if ("p14:prism" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Prism;
            let dir = "l", isContent = false, isInverted = false;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir") dir = _paramValues[i];
                else if (_paramNames[i] === "isContent") isContent = (_paramValues[i] === "1" || _paramValues[i] === "true");
                else if (_paramNames[i] === "isInverted") isInverted = (_paramValues[i] === "1" || _paramValues[i] === "true");
            }
            const prismMap = {
                "l_0_0": c_oAscSlideTransitionParams.Prism_Left,
                "r_0_0": c_oAscSlideTransitionParams.Prism_Right,
                "u_0_0": c_oAscSlideTransitionParams.Prism_Up,
                "d_0_0": c_oAscSlideTransitionParams.Prism_Down,
                "l_0_1": c_oAscSlideTransitionParams.Prism_Left_Inverted,
                "r_0_1": c_oAscSlideTransitionParams.Prism_Right_Inverted,
                "u_0_1": c_oAscSlideTransitionParams.Prism_Up_Inverted,
                "d_0_1": c_oAscSlideTransitionParams.Prism_Down_Inverted,
                "l_1_0": c_oAscSlideTransitionParams.Prism_Left_Content,
                "r_1_0": c_oAscSlideTransitionParams.Prism_Right_Content,
                "u_1_0": c_oAscSlideTransitionParams.Prism_Up_Content,
                "d_1_0": c_oAscSlideTransitionParams.Prism_Down_Content,
                "l_1_1": c_oAscSlideTransitionParams.Prism_Left_ContentInv,
                "r_1_1": c_oAscSlideTransitionParams.Prism_Right_ContentInv,
                "u_1_1": c_oAscSlideTransitionParams.Prism_Up_ContentInv,
                "d_1_1": c_oAscSlideTransitionParams.Prism_Down_ContentInv
            };
            const key = dir + "_" + (isContent ? "1" : "0") + "_" + (isInverted ? "1" : "0");
            this.TransitionOption = prismMap[key] !== undefined ? prismMap[key] : c_oAscSlideTransitionParams.Prism_Left;
        }
        else if ("p14:doors" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Doors;
            this.TransitionOption = c_oAscSlideTransitionParams.Doors_Horizontal;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "vert" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Doors_Vertical;
            }
        }
        else if ("p14:window" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Window;
            this.TransitionOption = c_oAscSlideTransitionParams.Window_Horizontal;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "vert" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Window_Vertical;
            }
        }
        else if ("p14:ferris" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Ferris;
            this.TransitionOption = c_oAscSlideTransitionParams.Ferris_Left;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "r" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Ferris_Right;
            }
        }
        else if ("p14:gallery" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Gallery;
            this.TransitionOption = c_oAscSlideTransitionParams.Gallery_Left;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "r" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Gallery_Right;
            }
        }
        else if ("p14:conveyor" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Conveyor;
            this.TransitionOption = c_oAscSlideTransitionParams.Conveyor_Left;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir" && "r" === _paramValues[i])
                    this.TransitionOption = c_oAscSlideTransitionParams.Conveyor_Right;
            }
        }
        else if ("p14:pan" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Pan;
            this.TransitionOption = c_oAscSlideTransitionParams.Pan_Left;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir") {
                    if ("r" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Pan_Right;
                    else if ("u" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Pan_Up;
                    else if ("d" === _paramValues[i]) this.TransitionOption = c_oAscSlideTransitionParams.Pan_Down;
                }
            }
        }
        else if ("p14:glitter" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Glitter;
            let dir = "l", pattern = "diamond";
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir") dir = _paramValues[i];
                else if (_paramNames[i] === "pattern") pattern = _paramValues[i];
            }
            const glitterMap = {
                "l_diamond": c_oAscSlideTransitionParams.Glitter_Left_Diamond,
                "r_diamond": c_oAscSlideTransitionParams.Glitter_Right_Diamond,
                "u_diamond": c_oAscSlideTransitionParams.Glitter_Up_Diamond,
                "d_diamond": c_oAscSlideTransitionParams.Glitter_Down_Diamond,
                "l_hexagon": c_oAscSlideTransitionParams.Glitter_Left_Hexagon,
                "r_hexagon": c_oAscSlideTransitionParams.Glitter_Right_Hexagon,
                "u_hexagon": c_oAscSlideTransitionParams.Glitter_Up_Hexagon,
                "d_hexagon": c_oAscSlideTransitionParams.Glitter_Down_Hexagon
            };
            this.TransitionOption = glitterMap[dir + "_" + pattern] || c_oAscSlideTransitionParams.Glitter_Left_Diamond;
        }
        else if ("p14:flythrough" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Flythrough;
            let dir = "in", hasBounce = false;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "dir") dir = _paramValues[i];
                else if (_paramNames[i] === "hasBounce") hasBounce = (_paramValues[i] === "1" || _paramValues[i] === "true");
            }
            if (dir === "in" && !hasBounce) this.TransitionOption = c_oAscSlideTransitionParams.Flythrough_In;
            else if (dir === "in" && hasBounce) this.TransitionOption = c_oAscSlideTransitionParams.Flythrough_In_Bounce;
            else if (dir === "out" && !hasBounce) this.TransitionOption = c_oAscSlideTransitionParams.Flythrough_Out;
            else this.TransitionOption = c_oAscSlideTransitionParams.Flythrough_Out_Bounce;
        }
        else if ("p14:flash" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Flash;
            this.TransitionOption = c_oAscSlideTransitionParams.Flash_Default;
        }
        else if ("p14:shred" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Shred;
            let pattern = "strip", dir = "in";
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "pattern") pattern = _paramValues[i];
                else if (_paramNames[i] === "dir") dir = _paramValues[i];
            }
            const shredMap = {
                "strip_in": c_oAscSlideTransitionParams.Shred_StripIn,
                "strip_out": c_oAscSlideTransitionParams.Shred_StripOut,
                "rectangle_in": c_oAscSlideTransitionParams.Shred_RectangleIn,
                "rectangle_out": c_oAscSlideTransitionParams.Shred_RectangleOut
            };
            this.TransitionOption = shredMap[pattern + "_" + dir] || c_oAscSlideTransitionParams.Shred_StripIn;
        }
        else if ("p14:reveal" === _type)
        {
            typeMatched = true;
            this.TransitionType = c_oAscSlideTransitionTypes.Reveal;
            let thruBlk = false, dir = "l";
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "thruBlk") thruBlk = (_paramValues[i] === "1" || _paramValues[i] === "true");
                else if (_paramNames[i] === "dir") dir = _paramValues[i];
            }
            const revealMap = {
                "0_l": c_oAscSlideTransitionParams.Reveal_SmoothLeft,
                "0_r": c_oAscSlideTransitionParams.Reveal_SmoothRight,
                "1_l": c_oAscSlideTransitionParams.Reveal_BlackLeft,
                "1_r": c_oAscSlideTransitionParams.Reveal_BlackRight
            };
            this.TransitionOption = revealMap[(thruBlk ? "1" : "0") + "_" + dir] || c_oAscSlideTransitionParams.Reveal_SmoothLeft;
        }
        else if ("p15:prstTrans" === _type)
        {
            typeMatched = true;
            let prst = null, invX = false, invY = false;
            for (let i = 0; i < _len; i++) {
                if (_paramNames[i] === "prst") prst = _paramValues[i];
                else if (_paramNames[i] === "invX") invX = (_paramValues[i] === "1" || _paramValues[i] === "true");
                else if (_paramNames[i] === "invY") invY = (_paramValues[i] === "1" || _paramValues[i] === "true");
            }
            const prstTypeMap = {
                "fallOver": c_oAscSlideTransitionTypes.FallOver,
                "drape": c_oAscSlideTransitionTypes.Drape,
                "curtains": c_oAscSlideTransitionTypes.Curtains,
                "wind": c_oAscSlideTransitionTypes.Wind,
                "prestige": c_oAscSlideTransitionTypes.Prestige,
                "fracture": c_oAscSlideTransitionTypes.Fracture,
                "crush": c_oAscSlideTransitionTypes.Crush,
                "peelOff": c_oAscSlideTransitionTypes.PeelOff,
                "pageCurlDouble": c_oAscSlideTransitionTypes.PageCurlDouble,
                "pageCurlSingle": c_oAscSlideTransitionTypes.PageCurlSingle,
                "airplane": c_oAscSlideTransitionTypes.Airplane,
                "origami": c_oAscSlideTransitionTypes.Origami
            };
            const prstBaseParamMap = {
                "fallOver": c_oAscSlideTransitionParams.FallOver_Default,
                "drape": c_oAscSlideTransitionParams.Drape_Default,
                "curtains": c_oAscSlideTransitionParams.Curtains_Default,
                "wind": c_oAscSlideTransitionParams.Wind_Default,
                "prestige": c_oAscSlideTransitionParams.Prestige_Default,
                "fracture": c_oAscSlideTransitionParams.Fracture_Default,
                "crush": c_oAscSlideTransitionParams.Crush_Default,
                "peelOff": c_oAscSlideTransitionParams.PeelOff_Default,
                "pageCurlDouble": c_oAscSlideTransitionParams.PageCurlDouble_Default,
                "pageCurlSingle": c_oAscSlideTransitionParams.PageCurlSingle_Default,
                "airplane": c_oAscSlideTransitionParams.Airplane_Default,
                "origami": c_oAscSlideTransitionParams.Origami_Default
            };
            if (prstTypeMap[prst] !== undefined) {
                this.TransitionType = prstTypeMap[prst];
                const invOffset = (invX ? 1 : 0) + (invY ? 2 : 0);
                this.TransitionOption = prstBaseParamMap[prst] + invOffset;
            } else {
                typeMatched = false;
                this.TransitionType = c_oAscSlideTransitionTypes.Fade;
                this.TransitionOption = c_oAscSlideTransitionParams.Fade_Smoothly;
            }
        }
        else if ("p:none" !== _type)
        {
            this.TransitionType = c_oAscSlideTransitionTypes.Fade;
            this.TransitionOption = c_oAscSlideTransitionParams.Fade_Smoothly;
        }
		else if ("p:none" === _type) {
			typeMatched = true;
		}
    }
	return typeMatched;
};
CAscSlideTransition.prototype.fillXmlParams = function (aAttrNames, aAttrValues) {
    let sNodeName = null;
    switch (this.TransitionType)
    {
        case c_oAscSlideTransitionTypes.Fade:
        {
            sNodeName = "p:fade";
            switch (this.TransitionOption)
            {
                case c_oAscSlideTransitionParams.Fade_Smoothly:
                {
                    aAttrNames.push("thruBlk");
                    aAttrValues.push("0");
                    break;
                }
                case c_oAscSlideTransitionParams.Fade_Through_Black:
                {
                    aAttrNames.push("thruBlk");
                    aAttrValues.push("1");
                    break;
                }
                default:
                    break;
            }
            break;
        }
        case c_oAscSlideTransitionTypes.Push:
        {
            sNodeName = "p:push";
            switch (this.TransitionOption)
            {
                case c_oAscSlideTransitionParams.Param_Left:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("r");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Right:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("l");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Top:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("d");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Bottom:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("u");
                    break;
                }
                default:
                    break;
            }
            break;
        }
        case c_oAscSlideTransitionTypes.Wipe:
        {
            switch (this.TransitionOption)
            {
                case c_oAscSlideTransitionParams.Param_Left:
                {
                    sNodeName = "p:wipe";
                    aAttrNames.push("dir");
                    aAttrValues.push("r");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Right:
                {
                    sNodeName = "p:wipe";
                    aAttrNames.push("dir");
                    aAttrValues.push("l");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Top:
                {
                    sNodeName = "p:wipe";
                    aAttrNames.push("dir");
                    aAttrValues.push("d");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Bottom:
                {
                    sNodeName = "p:wipe";
                    aAttrNames.push("dir");
                    aAttrValues.push("u");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_TopLeft:
                {
                    sNodeName = "p:strips";
                    aAttrNames.push("dir");
                    aAttrValues.push("rd");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_TopRight:
                {
                    sNodeName = "p:strips";
                    aAttrNames.push("dir");
                    aAttrValues.push("ld");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_BottomLeft:
                {
                    sNodeName = "p:strips";
                    aAttrNames.push("dir");
                    aAttrValues.push("ru");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_BottomRight:
                {
                    sNodeName = "p:strips";
                    aAttrNames.push("dir");
                    aAttrValues.push("lu");
                    break;
                }
                default:
                    break;
            }
            break;
        }
        case c_oAscSlideTransitionTypes.Split:
        {
            sNodeName = "p:split";
            switch (this.TransitionOption)
            {
                case c_oAscSlideTransitionParams.Split_HorizontalIn:
                {
                    aAttrNames.push("orient");
                    aAttrNames.push("dir");
                    aAttrValues.push("horz");
                    aAttrValues.push("in");
                    break;
                }
                case c_oAscSlideTransitionParams.Split_HorizontalOut:
                {
                    aAttrNames.push("orient");
                    aAttrNames.push("dir");
                    aAttrValues.push("horz");
                    aAttrValues.push("out");
                    break;
                }
                case c_oAscSlideTransitionParams.Split_VerticalIn:
                {
                    aAttrNames.push("orient");
                    aAttrNames.push("dir");
                    aAttrValues.push("vert");
                    aAttrValues.push("in");
                    break;
                }
                case c_oAscSlideTransitionParams.Split_VerticalOut:
                {
                    aAttrNames.push("orient");
                    aAttrNames.push("dir");
                    aAttrValues.push("vert");
                    aAttrValues.push("out");
                    break;
                }
                default:
                    break;
            }
            break;
        }
        case c_oAscSlideTransitionTypes.UnCover:
        case c_oAscSlideTransitionTypes.Cover:
        {
            if (this.TransitionType === c_oAscSlideTransitionTypes.Cover)
                sNodeName = "p:cover";
            else
                sNodeName = "p:pull";

            switch (this.TransitionOption)
            {
                case c_oAscSlideTransitionParams.Param_Left:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("r");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Right:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("l");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Top:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("d");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_Bottom:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("u");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_TopLeft:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("rd");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_TopRight:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("ld");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_BottomLeft:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("ru");
                    break;
                }
                case c_oAscSlideTransitionParams.Param_BottomRight:
                {
                    aAttrNames.push("dir");
                    aAttrValues.push("lu");
                    break;
                }
                default:
                    break;
            }
            break;
        }
        case c_oAscSlideTransitionTypes.Clock:
        {
            switch (this.TransitionOption)
            {
                case c_oAscSlideTransitionParams.Clock_Clockwise:
                {
                    sNodeName = "p:wheel";
                    aAttrNames.push("spokes");
                    aAttrValues.push("1");
                    break;
                }
                case c_oAscSlideTransitionParams.Clock_Counterclockwise:
                {
                    sNodeName = "p14:wheelReverse";
                    aAttrNames.push("spokes");
                    aAttrValues.push("1");
                    break;
                }
                case c_oAscSlideTransitionParams.Clock_Wedge:
                {
                    sNodeName = "p:wedge";
                    break;
                }
                default:
                    break;
            }
            break;
        }
        case c_oAscSlideTransitionTypes.Zoom:
        {
            switch (this.TransitionOption)
            {
                case c_oAscSlideTransitionParams.Zoom_In:
                {
                    sNodeName = "p14:warp";
                    aAttrNames.push("dir");
                    aAttrValues.push("in");
                    break;
                }
                case c_oAscSlideTransitionParams.Zoom_Out:
                {
                    sNodeName = "p14:warp";
                    aAttrNames.push("dir");
                    aAttrValues.push("out");
                    break;
                }
                case c_oAscSlideTransitionParams.Zoom_AndRotate:
                {
                    sNodeName = "p:newsflash";
                    break;
                }
                default:
                    break;
            }
            break;
        }
        case c_oAscSlideTransitionTypes.Morph:
        {
            sNodeName = "p159:morph";
            aAttrNames.push("option");
            switch (this.TransitionOption)
            {
                case c_oAscSlideTransitionParams.Morph_Objects:
                {
                    aAttrValues.push("byObject");
                    break;
                }
                case c_oAscSlideTransitionParams.Morph_Words:
                {
                    aAttrValues.push("byWord");
                    break;
                }
                case c_oAscSlideTransitionParams.Morph_Letters:
                {
                    aAttrValues.push("byChar");
                    break;
                }
                default:
                {
                    aAttrValues.push("byObject");
                    break;
                }
            }
            break;
        }
        case c_oAscSlideTransitionTypes.Random:
        {
            sNodeName = "p:random";
            break;
        }
        case c_oAscSlideTransitionTypes.Cut:
        {
            sNodeName = "p:cut";
            aAttrNames.push("thruBlk");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Cut_ThroughBlack ? "1" : "0");
            break;
        }
        case c_oAscSlideTransitionTypes.Blinds:
        {
            sNodeName = "p:blinds";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Blinds_Vertical ? "vert" : "horz");
            break;
        }
        case c_oAscSlideTransitionTypes.Checker:
        {
            sNodeName = "p:checker";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Checker_Vertical ? "vert" : "horz");
            break;
        }
        case c_oAscSlideTransitionTypes.Comb:
        {
            sNodeName = "p:comb";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Comb_Vertical ? "vert" : "horz");
            break;
        }
        case c_oAscSlideTransitionTypes.Circle:
        {
            sNodeName = "p:circle";
            break;
        }
        case c_oAscSlideTransitionTypes.Diamond:
        {
            sNodeName = "p:diamond";
            break;
        }
        case c_oAscSlideTransitionTypes.Dissolve:
        {
            sNodeName = "p:dissolve";
            break;
        }
        case c_oAscSlideTransitionTypes.Plus:
        {
            sNodeName = "p:plus";
            break;
        }
        case c_oAscSlideTransitionTypes.RandomBar:
        {
            sNodeName = "p:randomBar";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.RandomBar_Vertical ? "vert" : "horz");
            break;
        }
        case c_oAscSlideTransitionTypes.BoxZoom:
        {
            sNodeName = "p:zoom";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.BoxZoom_In ? "in" : "out");
            break;
        }
        case c_oAscSlideTransitionTypes.Vortex:
        {
            sNodeName = "p14:vortex";
            const vortexDirMap = {};
            vortexDirMap[c_oAscSlideTransitionParams.Vortex_Left] = "l";
            vortexDirMap[c_oAscSlideTransitionParams.Vortex_Right] = "r";
            vortexDirMap[c_oAscSlideTransitionParams.Vortex_Up] = "u";
            vortexDirMap[c_oAscSlideTransitionParams.Vortex_Down] = "d";
            aAttrNames.push("dir");
            aAttrValues.push(vortexDirMap[this.TransitionOption] || "l");
            break;
        }
        case c_oAscSlideTransitionTypes.Switch:
        {
            sNodeName = "p14:switch";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Switch_Right ? "r" : "l");
            break;
        }
        case c_oAscSlideTransitionTypes.Flip:
        {
            sNodeName = "p14:flip";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Flip_Right ? "r" : "l");
            break;
        }
        case c_oAscSlideTransitionTypes.Ripple:
        {
            sNodeName = "p14:ripple";
            const rippleDirMap = {};
            rippleDirMap[c_oAscSlideTransitionParams.Ripple_LeftUp] = "lu";
            rippleDirMap[c_oAscSlideTransitionParams.Ripple_RightUp] = "ru";
            rippleDirMap[c_oAscSlideTransitionParams.Ripple_LeftDown] = "ld";
            rippleDirMap[c_oAscSlideTransitionParams.Ripple_RightDown] = "rd";
            rippleDirMap[c_oAscSlideTransitionParams.Ripple_Center] = "center";
            aAttrNames.push("dir");
            aAttrValues.push(rippleDirMap[this.TransitionOption] || "center");
            break;
        }
        case c_oAscSlideTransitionTypes.Honeycomb:
        {
            sNodeName = "p14:honeycomb";
            break;
        }
        case c_oAscSlideTransitionTypes.Prism:
        {
            sNodeName = "p14:prism";
            const prismAttrs = {};
            prismAttrs[c_oAscSlideTransitionParams.Prism_Left]           = ["l", "0", "0"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Right]          = ["r", "0", "0"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Up]             = ["u", "0", "0"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Down]           = ["d", "0", "0"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Left_Inverted]  = ["l", "0", "1"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Right_Inverted] = ["r", "0", "1"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Up_Inverted]    = ["u", "0", "1"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Down_Inverted]  = ["d", "0", "1"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Left_Content]   = ["l", "1", "0"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Right_Content]  = ["r", "1", "0"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Up_Content]     = ["u", "1", "0"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Down_Content]   = ["d", "1", "0"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Left_ContentInv]  = ["l", "1", "1"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Right_ContentInv] = ["r", "1", "1"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Up_ContentInv]    = ["u", "1", "1"];
            prismAttrs[c_oAscSlideTransitionParams.Prism_Down_ContentInv]  = ["d", "1", "1"];
            const pVals = prismAttrs[this.TransitionOption] || ["l", "0", "0"];
            aAttrNames.push("dir");        aAttrValues.push(pVals[0]);
            aAttrNames.push("isContent");  aAttrValues.push(pVals[1]);
            aAttrNames.push("isInverted"); aAttrValues.push(pVals[2]);
            break;
        }
        case c_oAscSlideTransitionTypes.Doors:
        {
            sNodeName = "p14:doors";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Doors_Vertical ? "vert" : "horz");
            break;
        }
        case c_oAscSlideTransitionTypes.Window:
        {
            sNodeName = "p14:window";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Window_Vertical ? "vert" : "horz");
            break;
        }
        case c_oAscSlideTransitionTypes.Ferris:
        {
            sNodeName = "p14:ferris";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Ferris_Right ? "r" : "l");
            break;
        }
        case c_oAscSlideTransitionTypes.Gallery:
        {
            sNodeName = "p14:gallery";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Gallery_Right ? "r" : "l");
            break;
        }
        case c_oAscSlideTransitionTypes.Conveyor:
        {
            sNodeName = "p14:conveyor";
            aAttrNames.push("dir");
            aAttrValues.push(this.TransitionOption === c_oAscSlideTransitionParams.Conveyor_Right ? "r" : "l");
            break;
        }
        case c_oAscSlideTransitionTypes.Pan:
        {
            sNodeName = "p14:pan";
            const panDirMap = {};
            panDirMap[c_oAscSlideTransitionParams.Pan_Left] = "l";
            panDirMap[c_oAscSlideTransitionParams.Pan_Right] = "r";
            panDirMap[c_oAscSlideTransitionParams.Pan_Up] = "u";
            panDirMap[c_oAscSlideTransitionParams.Pan_Down] = "d";
            aAttrNames.push("dir");
            aAttrValues.push(panDirMap[this.TransitionOption] || "l");
            break;
        }
        case c_oAscSlideTransitionTypes.Glitter:
        {
            sNodeName = "p14:glitter";
            const glitterAttrs = {};
            glitterAttrs[c_oAscSlideTransitionParams.Glitter_Left_Diamond]  = ["l", "diamond"];
            glitterAttrs[c_oAscSlideTransitionParams.Glitter_Right_Diamond] = ["r", "diamond"];
            glitterAttrs[c_oAscSlideTransitionParams.Glitter_Up_Diamond]    = ["u", "diamond"];
            glitterAttrs[c_oAscSlideTransitionParams.Glitter_Down_Diamond]  = ["d", "diamond"];
            glitterAttrs[c_oAscSlideTransitionParams.Glitter_Left_Hexagon]  = ["l", "hexagon"];
            glitterAttrs[c_oAscSlideTransitionParams.Glitter_Right_Hexagon] = ["r", "hexagon"];
            glitterAttrs[c_oAscSlideTransitionParams.Glitter_Up_Hexagon]    = ["u", "hexagon"];
            glitterAttrs[c_oAscSlideTransitionParams.Glitter_Down_Hexagon]  = ["d", "hexagon"];
            const gVals = glitterAttrs[this.TransitionOption] || ["l", "diamond"];
            aAttrNames.push("dir");     aAttrValues.push(gVals[0]);
            aAttrNames.push("pattern"); aAttrValues.push(gVals[1]);
            break;
        }
        case c_oAscSlideTransitionTypes.Flythrough:
        {
            sNodeName = "p14:flythrough";
            const ftAttrs = {};
            ftAttrs[c_oAscSlideTransitionParams.Flythrough_In]         = ["in", "0"];
            ftAttrs[c_oAscSlideTransitionParams.Flythrough_In_Bounce]  = ["in", "1"];
            ftAttrs[c_oAscSlideTransitionParams.Flythrough_Out]        = ["out", "0"];
            ftAttrs[c_oAscSlideTransitionParams.Flythrough_Out_Bounce] = ["out", "1"];
            const fVals = ftAttrs[this.TransitionOption] || ["in", "0"];
            aAttrNames.push("dir");       aAttrValues.push(fVals[0]);
            aAttrNames.push("hasBounce"); aAttrValues.push(fVals[1]);
            break;
        }
        case c_oAscSlideTransitionTypes.Flash:
        {
            sNodeName = "p14:flash";
            break;
        }
        case c_oAscSlideTransitionTypes.Shred:
        {
            sNodeName = "p14:shred";
            const shredAttrs = {};
            shredAttrs[c_oAscSlideTransitionParams.Shred_StripIn]      = ["strip", "in"];
            shredAttrs[c_oAscSlideTransitionParams.Shred_StripOut]     = ["strip", "out"];
            shredAttrs[c_oAscSlideTransitionParams.Shred_RectangleIn]  = ["rectangle", "in"];
            shredAttrs[c_oAscSlideTransitionParams.Shred_RectangleOut] = ["rectangle", "out"];
            const sVals = shredAttrs[this.TransitionOption] || ["strip", "in"];
            aAttrNames.push("pattern"); aAttrValues.push(sVals[0]);
            aAttrNames.push("dir");     aAttrValues.push(sVals[1]);
            break;
        }
        case c_oAscSlideTransitionTypes.Reveal:
        {
            sNodeName = "p14:reveal";
            const revealAttrs = {};
            revealAttrs[c_oAscSlideTransitionParams.Reveal_SmoothLeft]  = ["0", "l"];
            revealAttrs[c_oAscSlideTransitionParams.Reveal_SmoothRight] = ["0", "r"];
            revealAttrs[c_oAscSlideTransitionParams.Reveal_BlackLeft]   = ["1", "l"];
            revealAttrs[c_oAscSlideTransitionParams.Reveal_BlackRight]  = ["1", "r"];
            const rVals = revealAttrs[this.TransitionOption] || ["0", "l"];
            aAttrNames.push("thruBlk"); aAttrValues.push(rVals[0]);
            aAttrNames.push("dir");     aAttrValues.push(rVals[1]);
            break;
        }
        case c_oAscSlideTransitionTypes.FallOver:
        case c_oAscSlideTransitionTypes.Drape:
        case c_oAscSlideTransitionTypes.Curtains:
        case c_oAscSlideTransitionTypes.Wind:
        case c_oAscSlideTransitionTypes.Prestige:
        case c_oAscSlideTransitionTypes.Fracture:
        case c_oAscSlideTransitionTypes.Crush:
        case c_oAscSlideTransitionTypes.PeelOff:
        case c_oAscSlideTransitionTypes.PageCurlDouble:
        case c_oAscSlideTransitionTypes.PageCurlSingle:
        case c_oAscSlideTransitionTypes.Airplane:
        case c_oAscSlideTransitionTypes.Origami:
        {
            sNodeName = "p15:prstTrans";
            const typeToPreset = {};
            typeToPreset[c_oAscSlideTransitionTypes.FallOver]       = "fallOver";
            typeToPreset[c_oAscSlideTransitionTypes.Drape]          = "drape";
            typeToPreset[c_oAscSlideTransitionTypes.Curtains]       = "curtains";
            typeToPreset[c_oAscSlideTransitionTypes.Wind]           = "wind";
            typeToPreset[c_oAscSlideTransitionTypes.Prestige]       = "prestige";
            typeToPreset[c_oAscSlideTransitionTypes.Fracture]       = "fracture";
            typeToPreset[c_oAscSlideTransitionTypes.Crush]          = "crush";
            typeToPreset[c_oAscSlideTransitionTypes.PeelOff]        = "peelOff";
            typeToPreset[c_oAscSlideTransitionTypes.PageCurlDouble] = "pageCurlDouble";
            typeToPreset[c_oAscSlideTransitionTypes.PageCurlSingle] = "pageCurlSingle";
            typeToPreset[c_oAscSlideTransitionTypes.Airplane]       = "airplane";
            typeToPreset[c_oAscSlideTransitionTypes.Origami]        = "origami";
            const prstBaseMap = {};
            prstBaseMap[c_oAscSlideTransitionTypes.FallOver]       = c_oAscSlideTransitionParams.FallOver_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.Drape]          = c_oAscSlideTransitionParams.Drape_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.Curtains]       = c_oAscSlideTransitionParams.Curtains_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.Wind]           = c_oAscSlideTransitionParams.Wind_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.Prestige]       = c_oAscSlideTransitionParams.Prestige_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.Fracture]       = c_oAscSlideTransitionParams.Fracture_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.Crush]          = c_oAscSlideTransitionParams.Crush_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.PeelOff]        = c_oAscSlideTransitionParams.PeelOff_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.PageCurlDouble] = c_oAscSlideTransitionParams.PageCurlDouble_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.PageCurlSingle] = c_oAscSlideTransitionParams.PageCurlSingle_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.Airplane]       = c_oAscSlideTransitionParams.Airplane_Default;
            prstBaseMap[c_oAscSlideTransitionTypes.Origami]        = c_oAscSlideTransitionParams.Origami_Default;
            aAttrNames.push("prst");
            aAttrValues.push(typeToPreset[this.TransitionType]);
            const base = prstBaseMap[this.TransitionType];
            const offset = this.TransitionOption - base;
            if (offset & 1) { aAttrNames.push("invX"); aAttrValues.push("1"); }
            if (offset & 2) { aAttrNames.push("invY"); aAttrValues.push("1"); }
            break;
        }
        default:
            break;
    }
    return sNodeName;
};

AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideSetTransition] = CAscSlideTransition;


// информация о темах --------------------------------------------

function CAscThemeInfo(themeInfo)
{
    this.ThemeInfo = themeInfo;
    this.Index = -1000;
}
CAscThemeInfo.prototype.get_Name = function() { return this.ThemeInfo.Name; };
CAscThemeInfo.prototype.get_Url = function() { return this.ThemeInfo.Url; };
CAscThemeInfo.prototype.get_Image = function() { return this.ThemeInfo.Thumbnail; };
CAscThemeInfo.prototype.get_Index = function() { return this.Index; };

function CLayoutThumbnail()
{
    this.Index = 0;
    this.Name = "";
    this.Type = 15;
    this.Image = "";

    this.Width = 0;
    this.Height = 0;
}

CLayoutThumbnail.prototype.getIndex = function() { return this.Index; };
CLayoutThumbnail.prototype.getType = function() { return this.Type; };
CLayoutThumbnail.prototype.get_Image = function() { return this.Image; };
CLayoutThumbnail.prototype.get_Name = function() { return this.Name; };
CLayoutThumbnail.prototype.get_Width = function() { return this.Width; };
CLayoutThumbnail.prototype.get_Height = function() { return this.Height; };


function CompareTransitions(transition1, transition2){
    if(!transition1 || !transition2){
        return null;
    }
    var ret = new Asc.CAscSlideTransition();
    if(transition1.TransitionType === transition2.TransitionType){
        ret.TransitionType = transition1.TransitionType;
    }
    if(transition1.TransitionOption === transition2.TransitionOption){
        ret.TransitionOption = transition1.TransitionOption;
    }
    if(transition1.TransitionDuration === transition2.TransitionDuration){
        ret.TransitionDuration = transition1.TransitionDuration;
    }
    if(transition1.SlideAdvanceOnMouseClick === transition2.SlideAdvanceOnMouseClick){
        ret.SlideAdvanceOnMouseClick = transition1.SlideAdvanceOnMouseClick;
    }
    if(transition1.SlideAdvanceAfter === transition2.SlideAdvanceAfter){
        ret.SlideAdvanceAfter = transition1.SlideAdvanceAfter;
    }
    if(transition1.SlideAdvanceDuration === transition2.SlideAdvanceDuration){
        ret.SlideAdvanceDuration = transition1.SlideAdvanceDuration;
    }
    if(transition1.ShowLoop === transition2.ShowLoop){
        ret.ShowLoop = transition1.ShowLoop;
    }
    return ret;
}

function CAscDateTime() {
    this.DateTime = null;
    this.CustomDateTime = null;
    this.Lang = null;
}

CAscDateTime.prototype['get_DateTime'] = CAscDateTime.prototype.get_DateTime = function(){return this.DateTime;};
CAscDateTime.prototype['put_DateTime'] = CAscDateTime.prototype.put_DateTime = function(v){this.DateTime = v;};
CAscDateTime.prototype['get_CustomDateTime']  = CAscDateTime.prototype.get_CustomDateTime = function(){return this.CustomDateTime;};
CAscDateTime.prototype['put_CustomDateTime']  = CAscDateTime.prototype.put_CustomDateTime = function(v){this.CustomDateTime = v;};
CAscDateTime.prototype['get_Lang'] = CAscDateTime.prototype.get_Lang = function(){return this.Lang;};
CAscDateTime.prototype['put_Lang'] = CAscDateTime.prototype.put_Lang = function(v){this.Lang = v;};
CAscDateTime.prototype['get_DateTimeExamples'] = CAscDateTime.prototype.get_DateTimeExamples = function(){
    var oMap = {
        "datetime1": null,
        "datetime2": null,
        "datetime3": null,
        "datetime4": null,
        "datetime5": null,
        "datetime6": null,
        "datetime7": null,
        "datetime8": null,
        "datetime9": null,
        "datetime10": null,
        "datetime11": null,
        "datetime12": null,
        "datetime13": null
    };
    AscFormat.ExecuteNoHistory(function () {
        var oParaField = new AscCommonWord.CPresentationField();
        oParaField.RecalcInfo.TextPr = false;
        oParaField.CompiledPr = new CTextPr();
        oParaField.CompiledPr.InitDefault();
        oParaField.CompiledPr.Lang.Val = this.Lang;
        for(var key in oMap) {
            if(oMap.hasOwnProperty(key)) {
                oParaField.FieldType = key;
                 let sVal = oParaField.private_GetString();
                 if(sVal) {
                     oMap[key] = sVal;
                 }
            }
        }
    }, this, []);
    return oMap;

};

function CAscHFProps() {
    this.Footer = null;
    this.Header = null;
    this.DateTime = null;

    this.ShowDateTime = null;
    this.ShowSlideNum = null;
    this.ShowFooter = null;
    this.ShowHeader = null;

    this.ShowOnTitleSlide = null;


    this.api = null;
    this.DivId = null;
    this.slide = null;
    this.notes = null;
}

CAscHFProps.prototype['get_Footer'] = CAscHFProps.prototype.get_Footer = function(){return this.Footer;};
CAscHFProps.prototype['get_Header'] = CAscHFProps.prototype.get_Header = function(){return this.Header;};
CAscHFProps.prototype['get_DateTime'] = CAscHFProps.prototype.get_DateTime = function(){return this.DateTime;};
CAscHFProps.prototype['get_ShowSlideNum'] = CAscHFProps.prototype.get_ShowSlideNum = function(){return this.ShowSlideNum;};
CAscHFProps.prototype['get_ShowOnTitleSlide'] = CAscHFProps.prototype.get_ShowOnTitleSlide = function(){return this.ShowOnTitleSlide;};
CAscHFProps.prototype['get_ShowFooter'] = CAscHFProps.prototype.get_ShowFooter = function(){return this.ShowFooter;};
CAscHFProps.prototype['get_ShowHeader'] = CAscHFProps.prototype.get_ShowHeader = function(){return this.ShowHeader;};
CAscHFProps.prototype['get_ShowDateTime'] = CAscHFProps.prototype.get_ShowDateTime = function(){return this.ShowDateTime;};

CAscHFProps.prototype['put_ShowOnTitleSlide'] = CAscHFProps.prototype.put_ShowOnTitleSlide = function(v){this.ShowOnTitleSlide = v;};
CAscHFProps.prototype['put_Footer'] = CAscHFProps.prototype.put_Footer = function(v){this.Footer = v;};
CAscHFProps.prototype['put_Header'] = CAscHFProps.prototype.put_Header = function(v){this.Header = v;};
CAscHFProps.prototype['put_DateTime'] = CAscHFProps.prototype.put_DateTime = function(v){this.DateTime = v;};
CAscHFProps.prototype['put_ShowSlideNum'] = CAscHFProps.prototype.put_ShowSlideNum = function(v){this.ShowSlideNum = v;};
CAscHFProps.prototype['put_ShowFooter'] = CAscHFProps.prototype.put_ShowFooter = function(v){this.ShowFooter = v;};
CAscHFProps.prototype['put_ShowHeader'] = CAscHFProps.prototype.put_ShowHeader = function(v){this.ShowHeader = v;};
CAscHFProps.prototype['put_ShowDateTime'] = CAscHFProps.prototype.put_ShowDateTime = function(v){this.ShowDateTime = v;};

CAscHFProps.prototype['put_DivId'] = CAscHFProps.prototype.put_DivId = function(v){this.DivId = v;};
CAscHFProps.prototype['updateView'] = CAscHFProps.prototype.updateView = function(){
    if(!this.api) {
        return;
    }
    var oCanvas = AscCommon.checkCanvasInDiv(this.DivId);
    if(!oCanvas) {
        return;
    }
    const oPresentation = this.api.private_GetLogicDocument();
    var oContext = oCanvas.getContext('2d');
    oContext.clearRect(0, 0, oCanvas.width, oCanvas.height);
    var oSp, nPhType, aSpTree, oSlideObject = null, l, t, r, b;
    var i;
    let dWidth, dHeight;
    if(this.slide) {
        oSlideObject = this.slide.Layout;
        dWidth = oPresentation.GetWidthMM();
        dHeight = oPresentation.GetHeightMM();
    }
    else if(this.notes) {
        oSlideObject = this.notes.Master;
        dWidth = oPresentation.GetNotesWidthMM();
        dHeight = oPresentation.GetNotesHeightMM();
    }
    if(oSlideObject) {
        aSpTree = oSlideObject.cSld.spTree;

        oContext.fillStyle = "#FFFFFF";
        oContext.fillRect(0, 0, oCanvas.width, oCanvas.height);
        const rPR = AscCommon.AscBrowser.retinaPixelRatio;
        const nLineWidth = Math.round(rPR);
        oContext.lineWidth = nLineWidth;
        oContext.fillStyle = "#000000";
        if(Array.isArray(aSpTree)) {
            for(i = 0; i < aSpTree.length; ++i) {
                oSp = aSpTree[i];
                if(oSp.isPlaceholder()) {
                    oSp.recalculate();
                    l = ((oSp.x / dWidth * oCanvas.width) >> 0) + nLineWidth;
                    t = ((oSp.y / dHeight * oCanvas.height) >> 0) + nLineWidth;
                    r = (((oSp.x + oSp.extX)/ dWidth * oCanvas.width) >> 0);
                    b = (((oSp.y + oSp.extY)/ dHeight * oCanvas.height) >> 0);
                    if(r <= oCanvas.width && r + nLineWidth >= oCanvas.width) {
                        r = oCanvas.width - nLineWidth - 1;
                    }
                    if(b <= oCanvas.height && b + nLineWidth >= oCanvas.height) {
                        b = oCanvas.height - nLineWidth - 1;
                    }
                    nPhType = oSp.getPlaceholderType();
                    oContext.beginPath();
                    if(nPhType === AscFormat.phType_dt ||
                    nPhType === AscFormat.phType_ftr ||
                    nPhType === AscFormat.phType_hdr ||
                    nPhType === AscFormat.phType_sldNum) {
                        editor.WordControl.m_oDrawingDocument.AutoShapesTrack.AddRect(oContext, l, t, r, b, true);
                        oContext.closePath();
                        oContext.stroke();
                        if(nPhType === AscFormat.phType_dt && this.ShowDateTime
                            || nPhType === AscFormat.phType_ftr && this.ShowFooter
                            || nPhType === AscFormat.phType_hdr && this.ShowHeader
                            || nPhType === AscFormat.phType_sldNum && this.ShowSlideNum) {
                            oContext.fill();
                        }
                    }
                    else {
                        editor.WordControl.m_oDrawingDocument.AutoShapesTrack.AddRectDashClever(oContext, l, t, r, b, 3, 3, true);
                        oContext.closePath();
                    }
                }
            }
        }
    }
    //return oCanvas.toDataURL("image/png");
};
CAscHFProps.prototype['put_Api'] = CAscHFProps.prototype.put_Api = function(v){this.api = v;};


function CAscHF() {
    this.Slide = null;
    this.Notes = null;
}

CAscHF.prototype['put_Slide'] = CAscHF.prototype.put_Slide = function(v){this.Slide = v;};
CAscHF.prototype['get_Slide'] = CAscHF.prototype.get_Slide = function(){return this.Slide;};
CAscHF.prototype['put_Notes'] = CAscHF.prototype.put_Notes = function(v){this.Notes = v;};
CAscHF.prototype['get_Notes'] = CAscHF.prototype.get_Notes = function(){return this.Notes;};

//------------------------------------------------------------export----------------------------------------------------
window['Asc'] = window['Asc'] || {};
window['AscCommonSlide'] = window['AscCommonSlide'] || {};



window['AscCommonSlide']['CAscDateTime'] = window['AscCommonSlide'].CAscDateTime = CAscDateTime;
window['AscCommonSlide']['CAscHFProps'] = window['AscCommonSlide'].CAscHFProps = CAscHFProps;
window['AscCommonSlide']['CAscHF'] = window['AscCommonSlide'].CAscHF = CAscHF;

window['Asc']['CAscSlideTransition'] = CAscSlideTransition;
window['AscCommonSlide'].CompareTransitions = CompareTransitions;
CAscSlideTransition.prototype['put_TransitionType'] = CAscSlideTransition.prototype.put_TransitionType;
CAscSlideTransition.prototype['get_TransitionType'] = CAscSlideTransition.prototype.get_TransitionType;
CAscSlideTransition.prototype['put_TransitionOption'] = CAscSlideTransition.prototype.put_TransitionOption;
CAscSlideTransition.prototype['get_TransitionOption'] = CAscSlideTransition.prototype.get_TransitionOption;
CAscSlideTransition.prototype['put_TransitionDuration'] = CAscSlideTransition.prototype.put_TransitionDuration;
CAscSlideTransition.prototype['get_TransitionDuration'] = CAscSlideTransition.prototype.get_TransitionDuration;
CAscSlideTransition.prototype['put_SlideAdvanceOnMouseClick'] = CAscSlideTransition.prototype.put_SlideAdvanceOnMouseClick;
CAscSlideTransition.prototype['get_SlideAdvanceOnMouseClick'] = CAscSlideTransition.prototype.get_SlideAdvanceOnMouseClick;
CAscSlideTransition.prototype['put_SlideAdvanceAfter'] = CAscSlideTransition.prototype.put_SlideAdvanceAfter;
CAscSlideTransition.prototype['get_SlideAdvanceAfter'] = CAscSlideTransition.prototype.get_SlideAdvanceAfter;
CAscSlideTransition.prototype['put_SlideAdvanceDuration'] = CAscSlideTransition.prototype.put_SlideAdvanceDuration;
CAscSlideTransition.prototype['get_SlideAdvanceDuration'] = CAscSlideTransition.prototype.get_SlideAdvanceDuration;
CAscSlideTransition.prototype['put_ShowLoop'] = CAscSlideTransition.prototype.put_ShowLoop;
CAscSlideTransition.prototype['get_ShowLoop'] = CAscSlideTransition.prototype.get_ShowLoop;
CAscSlideTransition.prototype['applyProps'] = CAscSlideTransition.prototype.applyProps;
CAscSlideTransition.prototype['createDuplicate'] = CAscSlideTransition.prototype.createDuplicate;
CAscSlideTransition.prototype['makeDuplicate'] = CAscSlideTransition.prototype.makeDuplicate;
CAscSlideTransition.prototype['setUndefinedOptions'] = CAscSlideTransition.prototype.setUndefinedOptions;
CAscSlideTransition.prototype['setDefaultParams'] = CAscSlideTransition.prototype.setDefaultParams;
CAscSlideTransition.prototype['Write_ToBinary'] = CAscSlideTransition.prototype.Write_ToBinary;
CAscSlideTransition.prototype['Read_FromBinary'] = CAscSlideTransition.prototype.Read_FromBinary;

window['AscCommonSlide'].CAscThemeInfo = CAscThemeInfo;
CAscThemeInfo.prototype['get_Name'] = CAscThemeInfo.prototype.get_Name;
CAscThemeInfo.prototype['get_Url'] = CAscThemeInfo.prototype.get_Url;
CAscThemeInfo.prototype['get_Image'] = CAscThemeInfo.prototype.get_Image;
CAscThemeInfo.prototype['get_Index'] = CAscThemeInfo.prototype.get_Index;

CLayoutThumbnail.prototype['getIndex'] = CLayoutThumbnail.prototype.getIndex;
CLayoutThumbnail.prototype['getType'] = CLayoutThumbnail.prototype.getType;
CLayoutThumbnail.prototype['get_Image'] = CLayoutThumbnail.prototype.get_Image;
CLayoutThumbnail.prototype['get_Name'] = CLayoutThumbnail.prototype.get_Name;
CLayoutThumbnail.prototype['get_Width'] = CLayoutThumbnail.prototype.get_Width;
CLayoutThumbnail.prototype['get_Height'] = CLayoutThumbnail.prototype.get_Height;
