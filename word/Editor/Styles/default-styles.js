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

(function()
{
	const DEFAULT_STYLES = {
		PARAGRAPH : {
			"Normal" : {
				QFormat : true
			},
			
			"Heading 1" : {
				BasedOn : "Normal",
				Next    : "Normal",
				
				ParaPr : {
					KeepNext : true,
					KeepLines : true,
					Spacing : {
						Before : 360 * g_dKoef_twips_to_mm,
						After  : 80 * g_dKoef_twips_to_mm
					},
					OutlineLvl : 0
				},
				
				TextPr : {
				
				}
			},
			
			"Heading 2" : {
			
			},
			
			"Heading 3" : {
			
			},
			
			"Heading 4" : {
			
			},
			
			"Heading 5" : {
			
			},
			
			"Heading 6" : {
			
			},
			
			"Heading 7" : {
			
			},
			"Heading 8" : {
			
			},
			"Heading 9" : {
			
			},
		},
		
		
		
	};
	
	
	function generateDefaultStyles(styleManager)
	{
		let styleMap = {};
		
		for (let i = 0; i < DEFAULT_STYLES.length; ++i)
		{
			let style = AscWord.CStyle.fromObject(DEFAULT_STYLES[i]);
			if (!style)
				continue;
			
			let styleId = styleManager.Add(style);
			styleMap[style.GetName()] = {
				style   : style,
				styleId : styleId,
				next    : undefined !== DEFAULT_STYLES[i].Next ? DEFAULT_STYLES[i].Next : undefined,
				basedOn : undefined !== DEFAULT_STYLES[i].BasedOn ? DEFAULT_STYLES[i].BasedOn : undefined,
				link    : undefined !== DEFAULT_STYLES[i].Link ? DEFAULT_STYLES[i].Link : undefined;
			};
		}
		
		
		for (let styleName in styleMap)
		{
			let entry = styleMap[styleName];
			let style = entry.style;
			
			if (entry.next && styleMap[entry.next])
				style.SetNext(styleMap[entry.next]);
			
			if (entry.basedOn && styleMap[entry.basedOn])
				style.SetBasedOn(styleMap[entry.basedOn]);
			
			if (entry.link)
			{
				if (styleMap[entry.link])
					style.SetLink(styleMap[entry.link]);
				else
				{
					// TODO: Generate linked pair
				}
			}
		}
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DEFAULT_STYLES = DEFAULT_STYLES;
})(window);
