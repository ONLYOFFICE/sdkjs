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

(function()
{
	//------------------------------------------------------------------------------------------------------------------
	// CustomXmlManager
	//------------------------------------------------------------------------------------------------------------------
	AscDFH.historyitem_CustomXmlManager_Add    = AscDFH.historyitem_type_CustomXmlManager | 1;
	AscDFH.historyitem_CustomXmlManager_Remove = AscDFH.historyitem_type_CustomXmlManager | 2;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBase}
	 */
	function CChangesCustomXmlManagerAdd(Class, Id, xml)
	{
		AscDFH.CChangesBase.call(this, Class);
		
		this.Id		= Id;
		this.xml	= xml;
	}
	AscDFH.InheritBaseChange(
		CChangesCustomXmlManagerAdd,
		AscDFH.CChangesBase,
		AscDFH.historyitem_CustomXmlManager_Add
	);
	AscDFH.CChangesCustomXmlManagerAdd = CChangesCustomXmlManagerAdd;
	CChangesCustomXmlManagerAdd.prototype.Undo = function()
	{
		let oXml = this.Class.m_arrXmlById[this.Id];
		if (oXml)
		{
			delete this.Class.m_arrXmlById[this.Id];
			for (let i = 0; i < this.Class.xml.length; i++)
			{
				if (this.Class.xml[i] === oXml)
				{
					this.Class.xml.splice(i, 1);
				}
			}
		}
	};
	CChangesCustomXmlManagerAdd.prototype.Redo = function()
	{
		this.Class.m_arrXmlById[this.Id] = this.xml;
		this.Class.xml.push(this.xml);
	};
	CChangesCustomXmlManagerAdd.prototype.WriteToBinary = function(Writer)
	{
		// String : Id customXML
		Writer.WriteString2(this.Id);
	};
	CChangesCustomXmlManagerAdd.prototype.ReadFromBinary = function(Reader)
	{
		// String : Id customXML
		this.Id      = Reader.GetString2();
		this.xml = AscCommon.g_oTableId.Get_ById(this.Id);
	};
	CChangesCustomXmlManagerAdd.prototype.CreateReverseChange = function()
	{
		return new CChangesCustomXmlManagerRemove(this.Class, this.Id, this.xml);
	};
	CChangesCustomXmlManagerAdd.prototype.Merge = function(oChange)
	{
		if (this.Class !== oChange.Class)
			return true;
		
		if ((AscDFH.historyitem_CustomXmlManager_Add === oChange.Type || AscDFH.historyitem_CustomXmlManager_Remove === oChange.Type) && this.Id === oChange.Id)
			return false;
		
		return true;
	};
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBase}
	 */
	function CChangesCustomXmlManagerRemove(Class, Id, xml)
	{
		AscDFH.CChangesBase.call(this, Class);
		
		this.Id		= Id;
		this.xml	= xml;
	}
	AscDFH.InheritBaseChange(
		CChangesCustomXmlManagerRemove,
		AscDFH.CChangesBase,
		AscDFH.historyitem_CustomXmlManager_Remove
	);
	AscDFH.CChangesCustomXmlManagerRemove = CChangesCustomXmlManagerRemove;
	CChangesCustomXmlManagerRemove.prototype.Undo = function()
	{
		this.Class.m_arrXmlById[this.Id] = this.xml;
		this.Class.xml.push(this.xml);
	};
	CChangesCustomXmlManagerRemove.prototype.Redo = function()
	{
		let xml = this.Class.m_arrXmlById[this.Id];
		if (xml)
		{
			delete this.Class.m_arrXmlById[this.Id];
			for (let i = 0; i < this.Class.xml.length; i++)
			{
				if (this.Class.xml[i] === xml)
				{
					this.Class.xml.splice(i, 1);
				}
			}
		}
	};
	CChangesCustomXmlManagerRemove.prototype.WriteToBinary = function(Writer)
	{
		// String : Id customXML
		Writer.WriteString2(this.Id);
	};
	CChangesCustomXmlManagerRemove.prototype.ReadFromBinary = function(Reader)
	{
		// String : Id customXML
		this.Id		= Reader.GetString2();
		this.xml	= AscCommon.g_oTableId.Get_ById(this.Id);
	};
	CChangesCustomXmlManagerRemove.prototype.CreateReverseChange = function()
	{
		return new CChangesCustomXmlManagerAdd(this.Class, this.Id, this.xml);
	};
	CChangesCustomXmlManagerRemove.prototype.Merge = function(oChange)
	{
		if (this.Class !== oChange.Class)
			return true;
		
		if ((AscDFH.historyitem_CustomXmlManager_Add === oChange.Type || AscDFH.historyitem_CustomXmlManager_Remove === oChange.Type) && this.Id === oChange.Id)
			return false;
		
		return true;
	};
	
	
	
	//------------------------------------------------------------------------------------------------------------------
	// CustomXml
	//------------------------------------------------------------------------------------------------------------------
	AscDFH.historyitem_CustomXmlManager_Content_Start = AscDFH.historyitem_type_CustomXml | 1;
	AscDFH.historyitem_CustomXmlManager_Content_Part  = AscDFH.historyitem_type_CustomXml | 2;
	AscDFH.historyitem_CustomXmlManager_Content_End   = AscDFH.historyitem_type_CustomXml | 3;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesCustomXmlContentStart(Class, Old, New, Color)
	{
		AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
	}
	AscDFH.InheritBaseChange(
		CChangesCustomXmlContentStart,
		AscDFH.CChangesBaseProperty,
		AscDFH.historyitem_CustomXmlManager_Content_Start
	);
	AscDFH.CChangesCustomXmlContentStart = CChangesCustomXmlContentStart;
	CChangesCustomXmlContentStart.prototype.Undo = function()
	{
		if (!this.Class.partsOfCustomXml)
		{
			return this.Redo();
		}
		
		this.Class.m_aCustomXmlData = this.Class.partsOfCustomXml.reverse().join("");
		this.Class.addContentByXMLString(this.Class.m_aCustomXmlData);
		delete this.Class.partsOfCustomXml;
	};
	CChangesCustomXmlContentStart.prototype.Redo = function()
	{
		if (this.Class.partsOfCustomXml)
		{
			return this.Undo();
		}
		this.Class.partsOfCustomXml = [];
	};
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesCustomXmlContentPart(Class, Old, New, Color)
	{
		AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
	}
	
	AscDFH.InheritBaseChange(
		CChangesCustomXmlContentPart,
		AscDFH.CChangesBaseProperty,
		AscDFH.historyitem_CustomXmlManager_Content_Part
	);
	AscDFH.CChangesCustomXmlContentPart = CChangesCustomXmlContentPart;
	CChangesCustomXmlContentPart.prototype.private_SetValue = function(oPr)
	{
		if (oPr.length)
		{
			this.Class.partsOfCustomXml.push(oPr);
		}
	};
	CChangesCustomXmlContentPart.prototype.WriteToBinary    = function(Writer)
	{
		Writer.WriteString2(this.Old);
		Writer.WriteString2(this.New);
	};
	CChangesCustomXmlContentPart.prototype.ReadFromBinary   = function(Reader)
	{
		this.Old = Reader.GetString2();
		this.New = Reader.GetString2();
	};
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesCustomXmlContentEnd(Class, Old, New, Color)
	{
		AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
	}
	
	AscDFH.InheritBaseChange(
		CChangesCustomXmlContentEnd,
		AscDFH.CChangesBaseProperty,
		AscDFH.historyitem_CustomXmlManager_Content_End
	);
	AscDFH.CChangesCustomXmlContentEnd = CChangesCustomXmlContentEnd;
	CChangesCustomXmlContentEnd.prototype.Undo = function()
	{
		if (this.Class.partsOfCustomXml)
		{
			return this.Redo();
		}
		this.Class.partsOfCustomXml = [];
	};
	CChangesCustomXmlContentEnd.prototype.Redo = function()
	{
		if (!this.Class.partsOfCustomXml)
		{
			return this.Undo();
		}
		
		this.Class.m_aCustomXmlData = this.Class.partsOfCustomXml.join("");
		this.Class.addContentByXMLString(this.Class.m_aCustomXmlData);
		delete this.Class.partsOfCustomXml;
	};
})();


