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

$(function () {
	
	
	QUnit.testStart( function() { Reset() } );
	
	QUnit.test("Rich text content control load from/save CustomXML", async function (assert)
	{
		CreateCustomXMLForDocument(getCheck(oCustomXMLData.linearXML));
		
		let c1	= 	CreateFilledContentControl(0);
		CreateDataBindingForCC(c1);
		
		let p	= c1.GetFirstParagraph();
		assert.strictEqual(
			AscTest.GetParagraphText(p),
			'12345+2',
			"Content in Paragraph after load CustomXML"
		);
		
		let p1 = MoveCursorToCC(c1);
		AscTest.AddTextToParagraph(p1, "990")
		let str = oXMLManager.getCustomXMLString(oXMLManager.getCustomXml(0));
		
		assert.strictEqual(
			str,
			`<?xml version="1.0" encoding="UTF-8"?>\n<documentData xmlns="http://example.com/picture"><simpleText>&lt;?xml version="1.0" standalone="yes"?&gt;&lt;?mso-application progid="Word.Document"?&gt;&lt;pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"&gt;&lt;pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"&gt;&lt;pkg:xmlData&gt;&lt;Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"&gt;&lt;Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/&gt;&lt;/Relationships&gt;&lt;/pkg:xmlData&gt;&lt;/pkg:part&gt;&lt;pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"&gt;&lt;pkg:xmlData&gt;&lt;w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"&gt;&lt;w:body&gt;&lt;w:p w:rsidR="00000000" w:rsidRDefault="001D1115"&gt;&lt;w:r&gt;&lt;w:rPr&gt;&lt;w:lang w:val="en-US"/&gt;&lt;/w:rPr&gt;&lt;w:t&gt;12345+2&lt;/w:t&gt;&lt;/w:r&gt;&lt;/w:p&gt;&lt;w:sectPr w:rsidR="00000000"&gt;&lt;w:pgSz w:w="12240" w:h="15840"/&gt;&lt;w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/&gt;&lt;w:cols w:space="720"/&gt;&lt;/w:sectPr&gt;&lt;/w:body&gt;&lt;/w:document&gt;&lt;/pkg:xmlData&gt;&lt;/pkg:part&gt;&lt;pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="256"&gt;&lt;pkg:xmlData&gt;&lt;Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"&gt;&lt;Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/&gt;&lt;/Relationships&gt;&lt;/pkg:xmlData&gt;&lt;/pkg:part&gt;&lt;pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"&gt;&lt;pkg:xmlData&gt;&lt;w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh"&gt;&lt;w:docDefaults&gt;&lt;w:rPrDefault&gt;&lt;w:rPr&gt;&lt;w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/&gt;&lt;w:sz w:val="22"/&gt;&lt;w:szCs w:val="22"/&gt;&lt;w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/&gt;&lt;/w:rPr&gt;&lt;/w:rPrDefault&gt;&lt;w:pPrDefault&gt;&lt;w:pPr&gt;&lt;w:spacing w:after="160" w:line="259" w:lineRule="auto"/&gt;&lt;/w:pPr&gt;&lt;/w:pPrDefault&gt;&lt;/w:docDefaults&gt;&lt;w:style w:type="paragraph" w:default="1" w:styleId="a"&gt;&lt;w:name w:val="Normal"/&gt;&lt;w:qFormat/&gt;&lt;w:pPr&gt;&lt;w:spacing w:after="200" w:line="276" w:lineRule="auto"/&gt;&lt;/w:pPr&gt;&lt;/w:style&gt;&lt;w:style w:type="character" w:default="1" w:styleId="a0"&gt;&lt;w:name w:val="Default Paragraph Font"/&gt;&lt;w:uiPriority w:val="1"/&gt;&lt;w:semiHidden/&gt;&lt;w:unhideWhenUsed/&gt;&lt;/w:style&gt;&lt;w:style w:type="table" w:default="1" w:styleId="a1"&gt;&lt;w:name w:val="Normal Table"/&gt;&lt;w:uiPriority w:val="99"/&gt;&lt;w:semiHidden/&gt;&lt;w:unhideWhenUsed/&gt;&lt;w:tblPr&gt;&lt;w:tblInd w:w="0" w:type="dxa"/&gt;&lt;w:tblCellMar&gt;&lt;w:top w:w="0" w:type="dxa"/&gt;&lt;w:left w:w="108" w:type="dxa"/&gt;&lt;w:bottom w:w="0" w:type="dxa"/&gt;&lt;w:right w:w="108" w:type="dxa"/&gt;&lt;/w:tblCellMar&gt;&lt;/w:tblPr&gt;&lt;/w:style&gt;&lt;w:style w:type="numbering" w:default="1" w:styleId="a2"&gt;&lt;w:name w:val="No List"/&gt;&lt;w:uiPriority w:val="99"/&gt;&lt;w:semiHidden/&gt;&lt;w:unhideWhenUsed/&gt;&lt;/w:style&gt;&lt;/w:styles&gt;&lt;/pkg:xmlData&gt;&lt;/pkg:part&gt;&lt;/pkg:package&gt;</simpleText></documentData>`,
			"Check load CustomXML"
		);
		
		SetDataToContentControl(c1, "qwe");
		assert.strictEqual(
			oXMLManager.getCustomXMLString(oXMLManager.getCustomXml(0)),
			'<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>qwe</simpleText></documentData>',
			"Check saved CustomXML"
		);
	});
});
