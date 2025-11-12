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

$( function () {
	var xml = '<w:node1 attr0="val0" attr1="val1">';
			xml += '<node2 a:attr2="val&amp;">';
				xml += '<node3/>';
				xml += '<node4 attr3="val&apos;"/>';
				xml += '<node5>text&quot;</node5>';
				xml += '<node6>text6</node6>';
			xml += '</node2>';
			xml += '<node2>';
				xml += '<node3/>';
				xml += '<node4 attr3="val&gt;"/>';
				xml += '<node5>text&lt;</node5>';
				xml += '<node6>text6</node6>';
			xml += '</node2>';
		xml += '</w:node1>';

	module( "stax xml reader" );
	test( "Read", function () {
		var res = ''
		var reader = new StaxParser(xml);
		while (reader.Read()) {
			switch(reader.GetEventType()) {
				case EasySAXEvent.START_ELEMENT:
					res += '<';
					res += reader.GetName();
					while (reader.MoveToNextAttribute()) {
						res += ' ';
						res += reader.GetName();
						res += '=\"';
						res += reader.GetValue();
						res += '\"';
					}
					res += '>';
					break;
				case EasySAXEvent.CHARACTERS:
					res += reader.GetValue();
					break;
				case EasySAXEvent.END_ELEMENT:
					res += '</';
					res += 'end';
					//res += reader.GetName();
					res += '>';
					break;
			}
		}
		strictEqual( res, '<w:node1 attr0="val0" attr1="val1"><node2 a:attr2="val&amp;"><node3></end><node4 attr3="val&apos;"></end><node5>text&quot;</end><node6>text6</end></end><node2><node3></end><node4 attr3="val&gt;"></end><node5>text&lt;</end><node6>text6</end></end></end>' , 'Read');
	} );
	test( "Read2", function () {
		var res = ''
		var reader = new StaxParser(xml);
		while (reader.Read()) {
			switch(reader.GetEventType()) {
				case EasySAXEvent.START_ELEMENT:
					res += '<';
					res += reader.GetName();
					reader.SkipAttributes();
					res += '>';
					break;
				case EasySAXEvent.CHARACTERS:
					res += reader.GetValue();
					break;
			}
		}
		strictEqual( res, '<w:node1><node2><node3><node4><node5>text&quot;<node6>text6<node2><node3><node4><node5>text&lt;<node6>text6' , 'Read2');
	} );
	test( "ReadNextNode", function () {
		var res = ''
		var reader = new StaxParser(xml);
		while (reader.ReadNextNode()) {
			res += '<';
			res += reader.GetName();
			res += '/>';
		}
		strictEqual( res, '<w:node1/><node2/><node3/><node4/><node5/><node6/><node2/><node3/><node4/><node5/><node6/>' , 'ReadNextNode');
	} );
	test( "ReadNextSiblingNode", function () {
		var res = ''
		var reader = new StaxParser(xml);
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '/>';
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			res += '<';
			res += reader.GetName();
			while (reader.MoveToNextAttribute()) {
				res += ' ';
				res += reader.GetName();
				res += '=\"';
				res += reader.GetValue();
				res += '\"';
			}
			res += '/>';
		}
		strictEqual( res, '<w:node1/><node2 a:attr2="val&amp;"/><node2/>' , 'ReadNextSiblingNode');
	} );
	test( "ReadNextSiblingNode2", function () {
		var res = ''
		var reader = new StaxParser(xml);
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '/>';
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '/>';
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			res += '<';
			res += reader.GetName();
			res += '/>';
		}
		strictEqual( res, '<w:node1/><node2/><node3/><node4/><node5/><node6/>' , 'ReadNextSiblingNode2');
	} );
	test( "ReadTillEnd", function () {
		var res = ''
		var reader = new StaxParser(xml);
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '/>';
		while (reader.ReadNextNode()) {
			res += '<';
			res += reader.GetName();
			while (reader.MoveToNextAttribute()) {
				res += ' ';
				res += reader.GetName();
				res += '=\"';
				res += reader.GetValue();
				res += '\"';
			}
			res += '/>';
			reader.ReadTillEnd();
		}
		strictEqual( res, '<w:node1/><node2 a:attr2="val&amp;"/><node2/>' , 'ReadTillEnd');
	} );
	test( "ReadTillEnd2", function () {
		var res = ''
		var reader = new StaxParser(xml);
		reader.ReadNextNode();
		reader.ReadNextNode();
		reader.ReadNextNode();
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		while (reader.MoveToNextAttribute()) {
			res += ' ';
			res += reader.GetName();
			res += '=\"';
			res += reader.GetValue();
			res += '\"';
		}
		res += '/>';
		reader.ReadTillEnd();
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '/>';
		reader.ReadNextNode();
		reader.ReadNextNode();
		reader.ReadNextNode();
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '/>';
		reader.ReadTillEnd();
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '/>';
		strictEqual( res, '<node4 attr3="val&apos;"/><node5/><node4/><node5/>' , 'ReadTillEnd');
	} );
	test( "MoveToNextAttribute", function () {
		var res = ''
		var reader = new StaxParser(xml);
		while (reader.ReadNextNode()) {
			res += '<';
			res += reader.GetName();
			while(reader.MoveToNextAttribute()) {
				res += ' ';
				res += reader.GetName();
				res += '=\"';
				res += reader.GetValue();
				res += '\"';
			}
			res += '/>';
		}
		strictEqual( res, '<w:node1 attr0="val0" attr1="val1"/><node2 a:attr2="val&amp;"/><node3/><node4 attr3="val&apos;"/><node5/><node6/><node2/><node3/><node4 attr3="val&gt;"/><node5/><node6/>' , 'MoveToNextAttribute');

		var xml2 = '<w:node1  attr0  ="val0" attr1=  "val1" attr3  =  "val3" attr4=\'val4\' attr5 = \'val5\' />';
		var res = ''
		var reader = new StaxParser(xml2);
		while (reader.ReadNextNode()) {
			res += '<';
			res += reader.GetName();
			while(reader.MoveToNextAttribute()) {
				res += ' ';
				res += reader.GetName();
				res += '=\"';
				res += reader.GetValue();
				res += '\"';
			}
			res += '/>';
		}
		strictEqual( res, '<w:node1 attr0="val0" attr1="val1" attr3="val3" attr4="val4" attr5="val5"/>' , 'MoveToNextAttribute');
	} );
	test( "GetDepth", function () {
		var res = ''
		var reader = new StaxParser(xml);
		while (reader.ReadNextNode()) {
			res += '|';
			res += reader.GetDepth();
		}
		strictEqual( res, '|1|2|3|3|3|3|2|3|3|3|3' , 'GetDepth');
	} );
	test( "GetNameNoNS", function () {
		var res = ''
		var reader = new StaxParser(xml);
		while (reader.Read()) {
			switch(reader.GetEventType()) {
				case EasySAXEvent.START_ELEMENT:
					res += '<';
					res += reader.GetNameNoNS();
					while (reader.MoveToNextAttribute()) {
						res += ' ';
						res += reader.GetNameNoNS();
						res += '=\"';
						res += reader.GetValue();
						res += '\"';
					}
					res += '>';
					break;
			}
		}
		strictEqual( res, '<node1 attr0="val0" attr1="val1"><node2 attr2="val&amp;"><node3><node4 attr3="val&apos;"><node5><node6><node2><node3><node4 attr3="val&gt;"><node5><node6>' , 'Read');
	} );
	test( "GetValueDecodeXml", function () {
		var res = ''
		var reader = new StaxParser(xml);
		while (reader.Read()) {
			switch(reader.GetEventType()) {
				case EasySAXEvent.START_ELEMENT:
					res += '<';
					res += reader.GetName();
					while (reader.MoveToNextAttribute()) {
						res += ' ';
						res += reader.GetName();
						res += '=\"';
						res += reader.GetValueDecodeXml();
						res += '\"';
					}
					res += '>';
					break;
				case EasySAXEvent.CHARACTERS:
					res += reader.GetValueDecodeXml();
					break;
			}
		}
		strictEqual( res, '<w:node1 attr0="val0" attr1="val1"><node2 a:attr2="val&"><node3><node4 attr3="val\'"><node5>text"<node6>text6<node2><node3><node4 attr3="val>"><node5>text<<node6>text6' , 'GetValueDecodeXml');
	} );
	test( "GetText", function () {
		var res = ''
		var reader = new StaxParser(xml);
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '>';
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '>';
		res += reader.GetText();
		reader.ReadNextNode();
		res += '<';
		res += reader.GetName();
		res += '>';
		res += reader.GetText();
		strictEqual( res, '<w:node1><node2>text&quot;text6<node2>text&lt;text6' , 'GetText');
	} );
} );
