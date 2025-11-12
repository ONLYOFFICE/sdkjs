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

(function(window, undefined){
    // Import
    var HitInBezier4 = AscFormat.HitInBezier4;
    
    // arcTo new version
    function Arc3(ctx, fX, fY, fWidth, fHeight, fStartAngle, fSweepAngle, ellipseRotation)
    {
        // Params transform
        var sin1 = Math.sin(fStartAngle);
        var cos1 = Math.cos(fStartAngle);

        var __x = cos1 / fWidth;
        var __y = sin1 / fHeight;
        var l = 1 / Math.sqrt(__x * __x + __y * __y);

        var cx = fX - l * cos1;
        var cy = fY - l * sin1;

        // ellipseRotation can be undefined in old version calls and it is passed down anyway
        Arc2(ctx, cx - fWidth, cy - fHeight, 2 * fWidth, 2 * fHeight, fStartAngle, fSweepAngle, ellipseRotation);
    }

    function Arc2(ctx, fX, fY, fWidth, fHeight, fStartAngle, fSweepAngle, ellipseRotation)
    {
        // Calls Ellipse or EllipseArc
        if (0 >= fWidth || 0 >= fHeight)
            return;

        fStartAngle = -fStartAngle;
        fSweepAngle = -fSweepAngle;

        if (false /*is path closed*/ )
        {
            var fStartX = fX + fWidth / 2.0 + fWidth / 2 * Math.cos( AngToEllPrm( fStartAngle, fWidth / 2, fHeight / 2 ) );
            var fStartY = fY + fHeight / 2.0 - fHeight / 2 * Math.sin( AngToEllPrm ( fStartAngle, fWidth / 2, fHeight / 2 ) );

            if ( fSweepAngle < (2 * Math.PI) )
            {
                ctx._m(fStartX, fStartY);
            }
        }

        var bClockDirection = false;
        var fEndAngle = (2 * Math.PI) -(fSweepAngle + fStartAngle);
        var fSrtAngle = (2 * Math.PI) - fStartAngle;
        // by this moment we have endAngle( = fEndAngle) and startAngle( = fSrtAngle) clockwise

        if( fSweepAngle > 0 )
            bClockDirection = true;

        if(Math.abs(fSweepAngle) >= (2 * Math.PI))
        {
            Ellipse(ctx, fX + fWidth / 2, fY + fHeight / 2, fWidth / 2, fHeight / 2);
        }
        else
        {
            EllipseArc(ctx, fX + fWidth / 2, fY + fHeight / 2, fWidth / 2, fHeight / 2, fSrtAngle, fEndAngle, bClockDirection, ellipseRotation);
        }
    }

    function AngToEllPrm(fAngle, fXRad, fYRad)
    {
        // to ellipse parametric equation angle parameter - t
        // https://www.mathopenref.com/coordparamellipse.html
        return Math.atan2( Math.sin( fAngle ) / fYRad,  Math.cos( fAngle ) / fXRad );
    }

    function Ellipse(ctx, fX, fY, fXRad, fYRad)
    {
        ctx._m(fX - fXRad, fY);

        var c_fKappa = 0.552;
        ctx._c(fX - fXRad, fY + fYRad * c_fKappa, fX - fXRad * c_fKappa, fY + fYRad, fX, fY + fYRad);
        ctx._c(fX + fXRad * c_fKappa, fY + fYRad, fX + fXRad, fY + fYRad * c_fKappa, fX + fXRad, fY);
        ctx._c(fX + fXRad, fY - fYRad * c_fKappa, fX + fXRad * c_fKappa, fY - fYRad, fX, fY - fYRad);
        ctx._c(fX - fXRad * c_fKappa, fY - fYRad, fX - fXRad, fY - fYRad * c_fKappa, fX - fXRad, fY);
    }

    function EllipseArc(ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, bClockDirection, ellipseRotation)
    {
        // If arc comes through 0 angle - splits one to two EllipseArc2 calls
        while ( fAngle1 < 0 )
            fAngle1 += (2 * Math.PI);

        while ( fAngle1 > (2 * Math.PI) )
            fAngle1 -= (2 * Math.PI);

        while ( fAngle2 < 0 )
            fAngle2 += (2 * Math.PI);

        while ( fAngle2 >= (2 * Math.PI) )
            fAngle2 -= (2 * Math.PI);

        let rotatePoint;
        // fX = cx, fY = cy, fAngle1 = startAngle, fAngle2 = endAngle
        // rotatePoint added later for ellipse rotation. it is arc start point.
        // it is not used if ellipseRotation not used
        if (ellipseRotation !== undefined) {
            rotatePoint = {x : fX + fXRad * Math.cos( AngToEllPrm( fAngle1, fXRad, fYRad ) ),
                y : fY + fYRad * Math.sin( AngToEllPrm( fAngle1, fXRad, fYRad ) )};
        }

        // let rotatePoint = {x : fX - fXRad, y : fY};
        if ( !bClockDirection )
        {
            if ( fAngle1 <= fAngle2 )
                EllipseArc2(ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, false, ellipseRotation, rotatePoint);
            else
            {
                // if arc comes through 0 angle - split it
                EllipseArc2(ctx, fX, fY, fXRad, fYRad, fAngle1, 2 * Math.PI, false, ellipseRotation, rotatePoint);
                EllipseArc2(ctx, fX, fY, fXRad, fYRad, 0, fAngle2, false, ellipseRotation, rotatePoint);
            }
        }
        else
        {
            if ( fAngle1 >= fAngle2 )
                EllipseArc2(ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, true, ellipseRotation, rotatePoint);
            else
            {
                // if arc comes through 0 angle - split it
                EllipseArc2(ctx, fX, fY, fXRad, fYRad, fAngle1, 0, true, ellipseRotation, rotatePoint);
                EllipseArc2(ctx, fX, fY, fXRad, fYRad, 2 * Math.PI, fAngle2, true, ellipseRotation, rotatePoint);
            }
        }
    }

    function EllipseArc2(ctx, fX, fY, fXRad, fYRad, dAngle1, dAngle2, bClockDirection, ellipseRotation, rotatePoint)
    {
        // Splits calls to EllipseArc3 on each quadrant
        var nFirstPointQuard  = ((2 * dAngle1 / Math.PI) >> 0) + 1;
        var nSecondPointQuard = ((2 * dAngle2 / Math.PI) >> 0) + 1;
        nSecondPointQuard = Math.min( 4, Math.max( 1, nSecondPointQuard ) );
        nFirstPointQuard  = Math.min( 4, Math.max( 1, nFirstPointQuard ) );

        // fX = cx, fY = cy, fAngle1 = startAngle, fAngle2 = endAngle
        var fStartX = fX + fXRad * Math.cos( AngToEllPrm( dAngle1, fXRad, fYRad ) );
        var fStartY = fY + fYRad * Math.sin( AngToEllPrm( dAngle1, fXRad, fYRad ) );

        var EndPoint = {X: 0, Y: 0};
        //ctx._l(fStartX, fStartY);

        var fCurX = fStartX, fCurY = fStartY;
        var dStartAngle = dAngle1;
        var dEndAngle = 0;

        if ( !bClockDirection )
        {
            for( var nIndex = nFirstPointQuard; nIndex <= nSecondPointQuard; nIndex++ )
            {
                if ( nIndex == nSecondPointQuard )
                    dEndAngle = dAngle2;
                else
                    dEndAngle = nIndex * Math.PI / 2;
                if ( !( nIndex == nFirstPointQuard ) )
                    dStartAngle = (nIndex - 1 ) * Math.PI / 2;

                EndPoint = EllipseArc3(ctx, fX, fY, fXRad, fYRad, AngToEllPrm( dStartAngle, fXRad, fYRad ), AngToEllPrm( dEndAngle, fXRad, fYRad ), false, ellipseRotation, rotatePoint);
            }
        }
        else
        {
            for( var nIndex = nFirstPointQuard; nIndex >= nSecondPointQuard; nIndex-- )
            {
                if ( nIndex == nFirstPointQuard )
                    dStartAngle = dAngle1;
                else
                    dStartAngle = nIndex * Math.PI / 2;
                if ( !( nIndex == nSecondPointQuard ) )
                    dEndAngle = (nIndex - 1 ) * Math.PI / 2;
                else
                    dEndAngle = dAngle2;

                EndPoint = EllipseArc3(ctx, fX, fY, fXRad, fYRad, AngToEllPrm( dStartAngle, fXRad, fYRad ), AngToEllPrm( dEndAngle, fXRad, fYRad ), false, ellipseRotation, rotatePoint);
            }
        }
    }

    function EllipseArc3(ctx, fX, fY, fXRad, fYRad, dAngle1, dAngle2, bClockDirection, ellipseRotation, rotatePoint)
    {
        // fX = cx, fY = cy, fAngle1 = startAngle after AngToEllPrm, fAngle2 = endAngle after AngToEllPrm

        var fAlpha = Math.sin( dAngle2 - dAngle1 ) * ( Math.sqrt( 4.0 + 3.0 * Math.tan( (dAngle2 - dAngle1) / 2.0 ) * Math.tan( (dAngle2 - dAngle1) / 2.0 ) ) - 1.0 ) / 3.0;

        var sin1 = Math.sin(dAngle1);
        var cos1 = Math.cos(dAngle1);
        var sin2 = Math.sin(dAngle2);
        var cos2 = Math.cos(dAngle2);

        // calculate start point
        var fX1 = fX + fXRad * cos1;
        var fY1 = fY + fYRad * sin1;

        // calculate end point
        var fX2 = fX + fXRad * cos2;
        var fY2 = fY + fYRad * sin2;

        // bezier control point 1
        var fCX1 = fX1 - fAlpha * fXRad * sin1;
        var fCY1 = fY1 + fAlpha * fYRad * cos1;

        // bezier control point 2
        var fCX2 = fX2 + fAlpha * fXRad * sin2;
        var fCY2 = fY2 - fAlpha * fYRad * cos2;

        if (ellipseRotation !== undefined) {
            // https://www.figma.com/file/hs43oiAUyuoqFULVoJ5lyZ/EllipticArcConvert?type=design&node-id=265-2&mode=design&t=pla9h2OkZAde8Xim-0
            // ellipseRotation = 0;
            // TODO import
            /**
             * afin rotate clockwise
             * @param {number} x
             * @param {number} y
             * @param {number} radiansRotateAngle radians Rotate AntiClockWise Angle. E.g. 30 degrees rotates does DOWN.
             * @returns {{x: number, y: number}} point
             */
            function rotatePointAroundCordsStartClockWise(x, y, radiansRotateAngle) {
                let newX = x * Math.cos(radiansRotateAngle) + y * Math.sin(radiansRotateAngle);
                let newY = x * (-1) * Math.sin(radiansRotateAngle) + y * Math.cos(radiansRotateAngle);
                return {x : newX, y: newY};
            }

            let startPointRelative = {x : fX1 - rotatePoint.x, y : rotatePoint.y - fY1};
            let endPointRelative = {x : fX2 - rotatePoint.x, y : rotatePoint.y - fY2};
            let controlPoint1Relative = {x : fCX1 - rotatePoint.x, y : rotatePoint.y - fCY1};
            let controlPoint2Relative = {x : fCX2 - rotatePoint.x, y : rotatePoint.y - fCY2};

            let startPointRelativeRotated = rotatePointAroundCordsStartClockWise(startPointRelative.x, startPointRelative.y, ellipseRotation);
            let endPointRelativeRotated = rotatePointAroundCordsStartClockWise(endPointRelative.x, endPointRelative.y, ellipseRotation);
            let controlPoint1RelativeRotated = rotatePointAroundCordsStartClockWise(controlPoint1Relative.x, controlPoint1Relative.y, ellipseRotation);
            let controlPoint2RelativeRotated = rotatePointAroundCordsStartClockWise(controlPoint2Relative.x, controlPoint2Relative.y, ellipseRotation);

            // go to initial cord system
            let startPointRotated = {x : startPointRelativeRotated.x + rotatePoint.x, y : rotatePoint.y - startPointRelativeRotated.y};
            let endPointRotated = {x : endPointRelativeRotated.x + rotatePoint.x, y : rotatePoint.y - endPointRelativeRotated.y};
            let controlPoint1Rotated = {x : controlPoint1RelativeRotated.x + rotatePoint.x, y : rotatePoint.y - controlPoint1RelativeRotated.y};
            let controlPoint2Rotated = {x : controlPoint2RelativeRotated.x + rotatePoint.x, y : rotatePoint.y - controlPoint2RelativeRotated.y};

            fX1 = startPointRotated.x;
            fY1 = startPointRotated.y;
            fX2 = endPointRotated.x;
            fY2 = endPointRotated.y;
            fCX1 = controlPoint1Rotated.x;
            fCY1 = controlPoint1Rotated.y;
            fCX2 = controlPoint2Rotated.x;
            fCY2 = controlPoint2Rotated.y;
        }

        if ( !bClockDirection )
        {
            ctx._c(fCX1, fCY1, fCX2, fCY2, fX2, fY2);
            return {X: fX2, Y: fY2};
        }
        else
        {
            ctx._c(fCX2, fCY2, fCX1, fCY1, fX1, fY1);
            return {X: fX1, Y: fY1};
        }
    }

    // ----------------------------------------------------------------------- //

    function _ArcToOnCanvas(context, start_x, start_y, width_r, height_r, start_ang, sweep_ang)
    {
        var _sin = Math.sin(start_ang);
        var _cos = Math.cos(start_ang);

        var _x = _cos / width_r;
        var _y = _sin / height_r;
        var _l = 1 / Math.sqrt(_x * _x + _y * _y);

        var _cx = start_x - _l * _cos;
        var _cy = start_y - _l * _sin;


        ArcTo2OnCanvas(context, _cx - width_r, _cy - height_r, 2 * width_r, 2 * height_r, start_ang, sweep_ang);
    }

    function ArcTo2OnCanvas(context, _l_c_x, _l_c_y, width,  height, start_ang, sweep_ang)
    {
        if (0 >= width || 0 >= height)
            return;

        start_ang = -start_ang;
        sweep_ang = -sweep_ang;

        var bClockDirection = false;
        var fEndAngle = (2 * Math.PI) - (sweep_ang + start_ang);
        var fSrtAngle = (2 * Math.PI) - start_ang;
        if (sweep_ang > 0)
        {
            bClockDirection = true;
        }

        if (Math.abs(sweep_ang) >= (2 * Math.PI))
        {
            EllipseOnCanvas(context, _l_c_x + width / 2, _l_c_y + height / 2, width / 2, height / 2);
        }
        else
        {
            EllipseArcOnCanvas(context, _l_c_x + width / 2, _l_c_y + height / 2, width / 2, height / 2, fSrtAngle, fEndAngle, bClockDirection);
        }
    }

    function EllipseOnCanvas(ctx, fX, fY, fXRad, fYRad)
    {
        ctx.moveTo(fX - fXRad, fY);
        var c_fKappa = 0.552;
        ctx.bezierCurveTo(fX - fXRad, fY + fYRad * c_fKappa, fX - fXRad * c_fKappa, fY + fYRad, fX, fY + fYRad);
        ctx.bezierCurveTo(fX + fXRad * c_fKappa, fY + fYRad, fX + fXRad, fY + fYRad * c_fKappa, fX + fXRad, fY);
        ctx.bezierCurveTo(fX + fXRad, fY - fYRad * c_fKappa, fX + fXRad * c_fKappa, fY - fYRad, fX, fY - fYRad);
        ctx.bezierCurveTo(fX - fXRad * c_fKappa, fY - fYRad, fX - fXRad, fY - fYRad * c_fKappa, fX - fXRad, fY);
    }


    function EllipseArcOnCanvas(ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, bClockDirection)
    {
        while ( fAngle1 < 0 )
            fAngle1 += (2 * Math.PI);

        while ( fAngle1 > (2 * Math.PI) )
            fAngle1 -= (2 * Math.PI);

        while ( fAngle2 < 0 )
            fAngle2 += (2 * Math.PI);

        while ( fAngle2 >= (2 * Math.PI) )
            fAngle2 -= (2 * Math.PI);

        if ( !bClockDirection )
        {
            if ( fAngle1 <= fAngle2 )
                EllipseArc2OnCanvas(ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, false);
            else
            {
                EllipseArc2OnCanvas(ctx, fX, fY, fXRad, fYRad, fAngle1, 2 * Math.PI, false);
                EllipseArc2OnCanvas(ctx, fX, fY, fXRad, fYRad, 0, fAngle2, false);
            }
        }
        else
        {
            if ( fAngle1 >= fAngle2 )
                EllipseArc2OnCanvas(ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, true);
            else
            {
                EllipseArc2OnCanvas(ctx, fX, fY, fXRad, fYRad, fAngle1, 0, true);
                EllipseArc2OnCanvas(ctx, fX, fY, fXRad, fYRad, 2 * Math.PI, fAngle2, true);
            }
        }
    }


    function EllipseArc2OnCanvas(ctx, fX, fY, fXRad, fYRad, dAngle1, dAngle2, bClockDirection)
    {
        var nFirstPointQuard  = ((2 * dAngle1 / Math.PI) >> 0) + 1;
        var nSecondPointQuard = ((2 * dAngle2 / Math.PI) >> 0) + 1;
        nSecondPointQuard = Math.min( 4, Math.max( 1, nSecondPointQuard ) );
        nFirstPointQuard  = Math.min( 4, Math.max( 1, nFirstPointQuard ) );

        var fStartX = fX + fXRad * Math.cos( AngToEllPrm( dAngle1, fXRad, fYRad ) );
        var fStartY = fY + fYRad * Math.sin( AngToEllPrm( dAngle1, fXRad, fYRad ) );

        var EndPoint = {X: 0, Y: 0};
        ctx.lineTo(fStartX, fStartY);

        var fCurX = fStartX, fCurY = fStartY;
        var dStartAngle = dAngle1;
        var dEndAngle = 0;

        if ( !bClockDirection )
        {
            for( var nIndex = nFirstPointQuard; nIndex <= nSecondPointQuard; nIndex++ )
            {
                if ( nIndex == nSecondPointQuard )
                    dEndAngle = dAngle2;
                else
                    dEndAngle = nIndex * Math.PI / 2;
                if ( !( nIndex == nFirstPointQuard ) )
                    dStartAngle = (nIndex - 1 ) * Math.PI / 2;

                EndPoint = EllipseArc3OnCanvas(ctx, fX, fY, fXRad, fYRad, AngToEllPrm( dStartAngle, fXRad, fYRad ), AngToEllPrm( dEndAngle, fXRad, fYRad ), false);
            }
        }
        else
        {
            for( var nIndex = nFirstPointQuard; nIndex >= nSecondPointQuard; nIndex-- )
            {
                if ( nIndex == nFirstPointQuard )
                    dStartAngle = dAngle1;
                else
                    dStartAngle = nIndex * Math.PI / 2;
                if ( !( nIndex == nSecondPointQuard ) )
                    dEndAngle = (nIndex - 1 ) * Math.PI / 2;
                else
                    dEndAngle = dAngle2;

                EndPoint = EllipseArc3OnCanvas(ctx, fX, fY, fXRad, fYRad, AngToEllPrm( dStartAngle, fXRad, fYRad ), AngToEllPrm( dEndAngle, fXRad, fYRad ), false);
            }
        }
    }

    function EllipseArc3OnCanvas(ctx, fX, fY, fXRad, fYRad, dAngle1, dAngle2, bClockDirection)
    {
        var fAlpha = Math.sin( dAngle2 - dAngle1 ) * ( Math.sqrt( 4.0 + 3.0 * Math.tan( (dAngle2 - dAngle1) / 2.0 ) * Math.tan( (dAngle2 - dAngle1) / 2.0 ) ) - 1.0 ) / 3.0;

        var sin1 = Math.sin(dAngle1);
        var cos1 = Math.cos(dAngle1);
        var sin2 = Math.sin(dAngle2);
        var cos2 = Math.cos(dAngle2);

        var fX1 = fX + fXRad * cos1;
        var fY1 = fY + fYRad * sin1;

        var fX2 = fX + fXRad * cos2;
        var fY2 = fY + fYRad * sin2;

        var fCX1 = fX1 - fAlpha * fXRad * sin1;
        var fCY1 = fY1 + fAlpha * fYRad * cos1;

        var fCX2 = fX2 + fAlpha * fXRad * sin2;
        var fCY2 = fY2 - fAlpha * fYRad * cos2;

        if ( !bClockDirection )
        {
            ctx.bezierCurveTo(fCX1, fCY1, fCX2, fCY2, fX2, fY2);
            return {X: fX2, Y: fY2};
        }
        else
        {
            ctx.bezierCurveTo(fCX2, fCY2, fCX1, fCY1, fX1, fY1);
            return {X: fX1, Y: fY1};
        }
    }

    function _HitToArc(context, px, py,    start_x, start_y, width_r, height_r, start_ang, sweep_ang)
    {
        var _sin = Math.sin(start_ang);
        var _cos = Math.cos(start_ang);

        var _x = _cos / width_r;
        var _y = _sin / height_r;
        var _l = 1 / Math.sqrt(_x * _x + _y * _y);

        var _cx = start_x - _l * _cos;
        var _cy = start_y - _l * _sin;


        return HitToArc2(px, py, context, _cx - width_r, _cy - height_r, 2 * width_r, 2 * height_r, start_ang, sweep_ang);
    }

    function HitToArc2(px, py, context, _l_c_x, _l_c_y, width,  height, start_ang, sweep_ang)
    {
        if (0 >= width || 0 >= height)
            return;

        start_ang = -start_ang;
        sweep_ang = -sweep_ang;

        var bClockDirection = false;
        var fEndAngle = (2 * Math.PI) - (sweep_ang + start_ang);
        var fSrtAngle = (2 * Math.PI) - start_ang;
        if (sweep_ang > 0)
        {
            bClockDirection = true;
        }

        if (Math.abs(sweep_ang) >= (2 * Math.PI))
        {
            return HitToEllipseOnCanvas(px, py, context, _l_c_x + width / 2, _l_c_y + height / 2, width / 2, height / 2);
        }
        else
        {
            return HitToEllipseArcOnCanvas(px, py, context, _l_c_x + width / 2, _l_c_y + height / 2, width / 2, height / 2, fSrtAngle, fEndAngle, bClockDirection);
        }
    }

    function HitToEllipseOnCanvas(px, py, ctx, fX, fY, fXRad, fYRad)
    {
        var c_fKappa = 0.552;
        return HitInBezier4(ctx, px, py, fX - fXRad, fY, fX - fXRad, fY + fYRad * c_fKappa, fX - fXRad * c_fKappa, fY + fYRad, fX, fY + fYRad) ||
            HitInBezier4(ctx, px, py, fX, fY + fYRad, fX + fXRad * c_fKappa, fY + fYRad, fX + fXRad, fY + fYRad * c_fKappa, fX + fXRad, fY)||
            HitInBezier4(ctx, px, py, fX + fXRad, fY, fX + fXRad, fY - fYRad * c_fKappa, fX + fXRad * c_fKappa, fY - fYRad, fX, fY - fYRad)||
            HitInBezier4(ctx, px, py, fX, fY - fYRad, fX - fXRad * c_fKappa, fY - fYRad, fX - fXRad, fY - fYRad * c_fKappa, fX - fXRad, fY);
    }


    function HitToEllipseArcOnCanvas(px, py, ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, bClockDirection)
    {
        while ( fAngle1 < 0 )
            fAngle1 += (2 * Math.PI);

        while ( fAngle1 > (2 * Math.PI) )
            fAngle1 -= (2 * Math.PI);

        while ( fAngle2 < 0 )
            fAngle2 += (2 * Math.PI);

        while ( fAngle2 >= (2 * Math.PI) )
            fAngle2 -= (2 * Math.PI);

        if ( !bClockDirection )
        {
            if ( fAngle1 <= fAngle2 )
                return HitToEllipseArc2OnCanvas(px, py,ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, false);
            else
            {
                return HitToEllipseArc2OnCanvas(px, py,ctx, fX, fY, fXRad, fYRad, fAngle1, 2 * Math.PI, false)|| HitToEllipseArc2OnCanvas(px, py,ctx, fX, fY, fXRad, fYRad, 0, fAngle2, false);
            }
        }
        else
        {
            if ( fAngle1 >= fAngle2 )
                return HitToEllipseArc2OnCanvas(px, py,ctx, fX, fY, fXRad, fYRad, fAngle1, fAngle2, true);
            else
            {
                return HitToEllipseArc2OnCanvas(px, py, ctx, fX, fY, fXRad, fYRad, fAngle1, 0, true) || HitToEllipseArc2OnCanvas(px, py,ctx, fX, fY, fXRad, fYRad, 2 * Math.PI, fAngle2, true);
            }
        }
    }


    function HitToEllipseArc2OnCanvas(px, py, ctx, fX, fY, fXRad, fYRad, dAngle1, dAngle2, bClockDirection)
    {
        var nFirstPointQuard  = ((2 * dAngle1 / Math.PI) >> 0) + 1;
        var nSecondPointQuard = ((2 * dAngle2 / Math.PI) >> 0) + 1;
        nSecondPointQuard = Math.min( 4, Math.max( 1, nSecondPointQuard ) );
        nFirstPointQuard  = Math.min( 4, Math.max( 1, nFirstPointQuard ) );

        var fStartX = fX + fXRad * Math.cos( AngToEllPrm( dAngle1, fXRad, fYRad ) );
        var fStartY = fY + fYRad * Math.sin( AngToEllPrm( dAngle1, fXRad, fYRad ) );

        var EndPoint = {X: fStartX, Y: fStartY, hit : false};
        var dStartAngle = dAngle1;
        var dEndAngle = 0;

        if ( !bClockDirection )
        {
            for( var nIndex = nFirstPointQuard; nIndex <= nSecondPointQuard; nIndex++ )
            {
                if ( nIndex == nSecondPointQuard )
                    dEndAngle = dAngle2;
                else
                    dEndAngle = nIndex * Math.PI / 2;
                if ( !( nIndex == nFirstPointQuard ) )
                    dStartAngle = (nIndex - 1 ) * Math.PI / 2;


                EndPoint = HitToEllipseArc3OnCanvas(px, py, EndPoint, ctx, fX, fY, fXRad, fYRad, AngToEllPrm( dStartAngle, fXRad, fYRad ), AngToEllPrm( dEndAngle, fXRad, fYRad ), false);
                if(EndPoint.hit)
                {
                    return true;
                }
            }
        }
        else
        {
            for( var nIndex = nFirstPointQuard; nIndex >= nSecondPointQuard; nIndex-- )
            {
                if ( nIndex == nFirstPointQuard )
                    dStartAngle = dAngle1;
                else
                    dStartAngle = nIndex * Math.PI / 2;
                if ( !( nIndex == nSecondPointQuard ) )
                    dEndAngle = (nIndex - 1 ) * Math.PI / 2;
                else
                    dEndAngle = dAngle2;

                EndPoint = HitToEllipseArc3OnCanvas(px, py, EndPoint, ctx, fX, fY, fXRad, fYRad, AngToEllPrm( dStartAngle, fXRad, fYRad ), AngToEllPrm( dEndAngle, fXRad, fYRad ), false);
                if(EndPoint.hit)
                {
                    return true;
                }
            }
        }
        return false;
    }

    function HitToEllipseArc3OnCanvas(px, py, EndPoint, ctx, fX, fY, fXRad, fYRad, dAngle1, dAngle2, bClockDirection)
    {
        var fAlpha = Math.sin( dAngle2 - dAngle1 ) * ( Math.sqrt( 4.0 + 3.0 * Math.tan( (dAngle2 - dAngle1) / 2.0 ) * Math.tan( (dAngle2 - dAngle1) / 2.0 ) ) - 1.0 ) / 3.0;

        var sin1 = Math.sin(dAngle1);
        var cos1 = Math.cos(dAngle1);
        var sin2 = Math.sin(dAngle2);
        var cos2 = Math.cos(dAngle2);

        var fX1 = fX + fXRad * cos1;
        var fY1 = fY + fYRad * sin1;

        var fX2 = fX + fXRad * cos2;
        var fY2 = fY + fYRad * sin2;

        var fCX1 = fX1 - fAlpha * fXRad * sin1;
        var fCY1 = fY1 + fAlpha * fYRad * cos1;

        var fCX2 = fX2 + fAlpha * fXRad * sin2;
        var fCY2 = fY2 - fAlpha * fYRad * cos2;

        if ( !bClockDirection )
        {
            return {X: fX2, Y: fY2, hit :  HitInBezier4(ctx, px, py,EndPoint.X, EndPoint.Y, fCX1, fCY1, fCX2, fCY2, fX2, fY2)};
        }
        else
        {
            return {X: fX1, Y: fY1, hit :  HitInBezier4(ctx, px, py,EndPoint.X, EndPoint.Y, fCX2, fCY2, fCX1, fCY1, fX1, fY1)};
        }
    }

    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].ArcToCurvers = Arc3;
    window['AscFormat'].HitToArc = _HitToArc;
    window['AscFormat'].ArcToOnCanvas = _ArcToOnCanvas;
    window['AscFormat'].EllipseOnCanvas = EllipseOnCanvas;

})(window);
