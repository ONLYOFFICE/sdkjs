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
	
	/**
	 * @constructor
	 * @extends {CParagraphContentBase}
	 */
	function ParagraphPermBase()
	{
		CParagraphContentBase.call(this);
		this.Id = AscCommon.g_oIdCounter.Get_NewId();
		AscCommon.g_oTableId.Add(this, this.Id);
	}
	
	ParagraphPermBase.prototype = Object.create(CParagraphContentBase.prototype);
	ParagraphPermBase.prototype.constructor = ParagraphPermBase;
	Object.assign(ParagraphPermBase.prototype, AscWord.AnnotationMarkBase.prototype);
	
	ParagraphPermBase.prototype.Get_Id = function()
	{
		return this.Id;
	};
	ParagraphPermBase.prototype.GetId = function()
	{
		return this.Id;
	};
	ParagraphPermBase.prototype.Copy = function()
	{
		return new this.constructor();
	};
	ParagraphPermBase.prototype.PreDelete = function()
	{
		let logicDocument = this.GetLogicDocument();
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return;
		
		logicDocument.OnDeleteAnnotationMark(this);
	};
	ParagraphPermBase.prototype.SetParagraph = function(p)
	{
		CParagraphContentBase.prototype.SetParagraph.call(this, p);
		AscWord.registerPermRangeMark(this);
	};
	ParagraphPermBase.prototype.GetAllPermRangeMarks = function(marks)
	{
		marks.push(this);
	};
	/**
	 * Очень важно, что в режимах комментирования и просмотра, мы проход через данный элемент считаем как перемещение курсора,
	 * т.к. на этом завязано выделение текущего слова и применение настроек к текущему слову (оно должно применяться
	 * только к той части, что внутри разрешенного диапазона)
	 * @returns {boolean}
	 */
	ParagraphPermBase.prototype.IsStopCursorOnEntryExit = function()
	{
		let logicDocument = this.GetLogicDocument();
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return false;
		
		let api = logicDocument.GetApi();
		return api.isRestrictionComments() || api.isRestrictionView();
	};
	ParagraphPermBase.prototype.isPermMark = function()
	{
		return true;
	};
	ParagraphPermBase.prototype.getRangeId = function()
	{
		return this.rangeId;
	};
	ParagraphPermBase.prototype.getColFirst = function()
	{
		return this.colFirst;
	};
	ParagraphPermBase.prototype.getColLast = function()
	{
		return this.colLast;
	};
	ParagraphPermBase.prototype.getDisplacedByCustomXml = function()
	{
		return this.displacedByCustomXml;
	};
	ParagraphPermBase.prototype.getEd = function()
	{
		return this.ed;
	};
	ParagraphPermBase.prototype.getEdGrp = function()
	{
		return this.edGrp;
	};
	//----------------------------------------------------------------------------------------------------------------------
	// Collaboration
	//----------------------------------------------------------------------------------------------------------------------
	ParagraphPermBase.prototype.Refresh_RecalcData = function()
	{
	};
	
	/**
	 * @param rangeId
	 * @param colFirst
	 * @param colLast
	 * @param displacedByCustomXml
	 * @param ed
	 * @param edGrp
	 * @constructor
	 * @extends {ParagraphPermBase}
	 */
	function ParagraphPermStart(rangeId, colFirst, colLast, displacedByCustomXml, ed, edGrp)
	{
		this.rangeId              = rangeId;
		this.colFirst             = undefined !== colFirst && null !== colFirst ? colFirst : undefined;
		this.colLast              = undefined !== colLast && null !== colLast ? colLast : undefined;
		this.displacedByCustomXml = undefined !== displacedByCustomXml && null !== displacedByCustomXml ? displacedByCustomXml : undefined;
		this.ed                   = undefined !== ed && null !== ed ? ed : undefined;
		this.edGrp                = undefined !== edGrp && null !== edGrp ? edGrp : undefined;
		
		ParagraphPermBase.call(this);
		
		this.Type = para_PermStart;
	}
	ParagraphPermStart.prototype = Object.create(ParagraphPermBase.prototype);
	ParagraphPermStart.prototype.constructor = ParagraphPermStart;
	
	ParagraphPermStart.fromObject = function(obj)
	{
		if (!obj)
			obj = {};
		
		return new ParagraphPermStart(obj.id, obj.colFirst, obj.colLast, obj.displacedByCustomXml, obj.ed, obj.edGrp);
	};
	ParagraphPermStart.prototype.isStart = function()
	{
		return true;
	};
	ParagraphPermStart.prototype.Draw_HighLights = function(PDSH)
	{
		PDSH.addPermRange(this.rangeId);
	};
	ParagraphPermStart.prototype.Recalculate_PageEndInfo = function(PRSI, curLine, curRange)
	{
		PRSI.addPermRange(this.rangeId);
	};
	ParagraphPermStart.prototype.RecalculateEndInfo = function(PRSI)
	{
		PRSI.addPermRange(this.rangeId);
	};
	ParagraphPermStart.prototype.GetCurrentPermRanges = function(permRanges, isCurrent)
	{
		let pos = permRanges.indexOf(this.rangeId);
		// TODO: Такого не должно быть
		if (-1 !== pos)
			return;
		
		permRanges.push(this.rangeId);
	};
	ParagraphPermStart.prototype.Copy = function()
	{
		return new ParagraphPermStart(this.rangeId, this.colFirst, this.colLast, this.displacedByCustomXml, this.ed, this.edGrp);
	};
	ParagraphPermStart.prototype.Write_ToBinary2 = function(writer)
	{
		writer.WriteLong(AscDFH.historyitem_type_ParagraphPermStart);
		
		writer.WriteString2("" + this.Id);
		writer.WriteString2("" + this.rangeId);
		
		let startPos = writer.GetCurPosition();
		writer.Skip(4);
		let flags = 0;
		
		if (undefined !== this.colFirst)
		{
			writer.WriteLong(this.colFirst);
			flags |= 1;
		}
		
		if (undefined !== this.colLast)
		{
			writer.WriteLong(this.colLast);
			flags |= 2;
		}
		
		if (undefined !== this.displacedByCustomXml)
		{
			writer.WriteByte(this.displacedByCustomXml);
			flags |= 4;
		}
		
		if (undefined !== this.ed)
		{
			writer.WriteString2(this.ed);
			flags |= 8;
		}
		
		if (undefined !== this.edGrp)
		{
			writer.WriteByte(this.edGrp);
			flags |= 16;
		}
		
		let endPos = writer.GetCurPosition();
		writer.Seek(startPos);
		writer.WriteLong(flags);
		writer.Seek(endPos);
	};
	ParagraphPermStart.prototype.Read_FromBinary2 = function(reader)
	{
		this.Id      = reader.GetString2();
		this.rangeId = reader.GetString2();
		
		let flags = reader.GetLong();
		
		if (flags & 1)
			this.colFirst = reader.GetLong();
		
		if (flags & 2)
			this.colLast = reader.GetLong();
		
		if (flags & 4)
			this.displacedByCustomXml = reader.GetByte();
		
		if (flags & 8)
			this.ed = reader.GetString2();
		
		if (flags & 16)
			this.edGrp = reader.GetByte();
	};
	
	/**
	 * @param rangeId
	 * @constructor
	 * @extends {ParagraphPermBase}
	 */
	function ParagraphPermEnd(rangeId)
	{
		this.rangeId = rangeId;
		ParagraphPermBase.call(this);
		
		this.Type = para_PermEnd;
	}
	ParagraphPermEnd.prototype = Object.create(ParagraphPermBase.prototype);
	ParagraphPermEnd.prototype.constructor = ParagraphPermEnd;
	
	ParagraphPermEnd.fromObject = function(obj)
	{
		if (!obj)
			return null;
		
		return new ParagraphPermEnd(obj.id);
	};
	ParagraphPermEnd.prototype.isEnd = function()
	{
		return true;
	};
	ParagraphPermEnd.prototype.Draw_HighLights = function(PDSH)
	{
		PDSH.removePermRange(this.rangeId);
	};
	ParagraphPermEnd.prototype.Recalculate_PageEndInfo = function(PRSI, curLine, curRange)
	{
		PRSI.removePermRange(this.rangeId);
	};
	ParagraphPermEnd.prototype.RecalculateEndInfo = function(PRSI)
	{
		PRSI.removePermRange(this.rangeId);
	};
	ParagraphPermEnd.prototype.GetCurrentPermRanges = function(permRanges, isCurrent)
	{
		let pos = permRanges.indexOf(this.rangeId);
		
		// TODO: Такого не должно быть
		if (-1 === pos)
			return;
		
		if (pos === permRanges.length - 1)
			--permRanges.length;
		else
			permRanges.splice(pos, 1);
	};
	ParagraphPermEnd.prototype.Copy = function()
	{
		return new ParagraphPermEnd(this.rangeId);
	};
	ParagraphPermEnd.prototype.Write_ToBinary2 = function(writer)
	{
		writer.WriteLong(AscDFH.historyitem_type_ParagraphPermEnd);
		
		writer.WriteString2("" + this.Id);
		writer.WriteString2("" + this.rangeId);
	};
	ParagraphPermEnd.prototype.Read_FromBinary2 = function(reader)
	{
		this.Id      = reader.GetString2();
		this.rangeId = reader.GetString2();
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphPermStart = ParagraphPermStart;
	AscWord.ParagraphPermEnd   = ParagraphPermEnd;
})();

