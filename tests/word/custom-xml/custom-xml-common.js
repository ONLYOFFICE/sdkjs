/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

AscWord.CustomXmlManager.prototype.isSupported = function()
{
	return true;
};

function MockZLibEngine ()
{
	this.files = [];
	
	this.open = function (buf) {
		//console.log("ZLibFake: Open archive...");
		
		if (Array.isArray(buf)) {
			this.files = [];
			this.data = {};
			buf.forEach(item => {
				if (item.path && item.data) {
					this.files.push(item.path);
					this.data[item.path] = item.data;
				}
			});
			return this.files;
		}
	}
	this.create = function () {
		//console.log("ZLibFake: Creating new archive...");
		
		this.files = [];
		this.data = {};
		return true;
	}
	this.save = function () {
		//console.log("ZLibFake: Saving archive...");
		
		return this.files.map(path => ({
			path: path,
			data: this.data[path]
		}));
	}
	this.getFile = function (path) {
		//console.log(`ZLibFake: Getting file: ${path}`);
		
		const file = this.data[path] || null;
		return file ? file : null;
	}
	this.addFile = function (path, data) {
		//console.log(`ZLibFake: Add file: ${path}`);
		
		if (this.files.indexOf(path) === -1) {
			this.files.push(path);
		}
		this.data[path] = data;
		return true;
	}
	this.removeFile = function (path) {
		//console.log(`ZLibFake: Removing file: ${path}`);
		
		const index = this.files.indexOf(path);
		if (index > -1) {
			this.files.splice(index, 1);
			delete this.data[path];
			return true;
		}
		return false;
	}
	this.close = function() {
		//console.log("ZLibFake: Closing archive...");
		
		this.files = [];
		this.data = {};
	}
	this.getImageBlob = function (path) {
		// console.log(`ZLibFake: Getting image blob: ${path}`);
		
		// return new Blob();
	}
	this.getPaths = function () {
		console.log("ZLibFake: Getting all file paths...");
		
		return this.files;
	}
}

function MockZLib () {
	this.engine = new MockZLibEngine();
	this.files = [];
	
	this.open = function (buf) {
		this.files = this.engine.open(buf);
		return this.files.length > 0;
	}
	this.create = function () {
		return this.engine.create();
	}
	this.save = function () {
		return this.engine.save();
	}
	this.getFile = function (path) {
		return this.engine.getFile(path);
	}
	this.addFile = function (path, data) {
		this.files.push(path);
		return this.engine.addFile(path, data);
	}
	this.removeFile = function (path) {
		const index = this.files.indexOf(path);
		if (index > -1) {
			this.files.splice(index, 1); // Удаляем имя файла
		}
		
		return this.engine.removeFile(path);
	}
	this.close = function () {
		return this.engine.close();
	}
	this.getImageBlob = function (path) {
		return this.engine.getImageBlob(path);
	}
	this.getPaths = function () {
		return this.engine.getPaths();
	}
}

AscCommon.ZLib			= MockZLib;
let logicDocument		= AscTest.CreateLogicDocument();
let oXMLManager			= logicDocument.getCustomXmlManager();
const oCustomXMLData	= {
	'date': "2000-01-01",
	'dateFormatted': "01-01-2000",
	'checkboxTrue': true,
	'checkboxFalse': false,
	'checkbox0': 0,
	'checkbox1': 1,
	'checkboxMess': "hello",
	'checkboxMess2': "hello123",
	'linearXML': '&lt;?xml version="1.0" standalone="yes"?&gt;&lt;?mso-application progid="Word.Document"?&gt;&lt;pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"&gt;&lt;pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"&gt;&lt;pkg:xmlData&gt;&lt;Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"&gt;&lt;Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/&gt;&lt;/Relationships&gt;&lt;/pkg:xmlData&gt;&lt;/pkg:part&gt;&lt;pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"&gt;&lt;pkg:xmlData&gt;&lt;w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"&gt;&lt;w:body&gt;&lt;w:p w:rsidR="00000000" w:rsidRDefault="001D1115"&gt;&lt;w:r&gt;&lt;w:rPr&gt;&lt;w:lang w:val="en-US"/&gt;&lt;/w:rPr&gt;&lt;w:t&gt;12345+2&lt;/w:t&gt;&lt;/w:r&gt;&lt;/w:p&gt;&lt;w:sectPr w:rsidR="00000000"&gt;&lt;w:pgSz w:w="12240" w:h="15840"/&gt;&lt;w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/&gt;&lt;w:cols w:space="720"/&gt;&lt;/w:sectPr&gt;&lt;/w:body&gt;&lt;/w:document&gt;&lt;/pkg:xmlData&gt;&lt;/pkg:part&gt;&lt;pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="256"&gt;&lt;pkg:xmlData&gt;&lt;Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"&gt;&lt;Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/&gt;&lt;/Relationships&gt;&lt;/pkg:xmlData&gt;&lt;/pkg:part&gt;&lt;pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"&gt;&lt;pkg:xmlData&gt;&lt;w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh"&gt;&lt;w:docDefaults&gt;&lt;w:rPrDefault&gt;&lt;w:rPr&gt;&lt;w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/&gt;&lt;w:sz w:val="22"/&gt;&lt;w:szCs w:val="22"/&gt;&lt;w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/&gt;&lt;/w:rPr&gt;&lt;/w:rPrDefault&gt;&lt;w:pPrDefault&gt;&lt;w:pPr&gt;&lt;w:spacing w:after="160" w:line="259" w:lineRule="auto"/&gt;&lt;/w:pPr&gt;&lt;/w:pPrDefault&gt;&lt;/w:docDefaults&gt;&lt;w:style w:type="paragraph" w:default="1" w:styleId="a"&gt;&lt;w:name w:val="Normal"/&gt;&lt;w:qFormat/&gt;&lt;w:pPr&gt;&lt;w:spacing w:after="200" w:line="276" w:lineRule="auto"/&gt;&lt;/w:pPr&gt;&lt;/w:style&gt;&lt;w:style w:type="character" w:default="1" w:styleId="a0"&gt;&lt;w:name w:val="Default Paragraph Font"/&gt;&lt;w:uiPriority w:val="1"/&gt;&lt;w:semiHidden/&gt;&lt;w:unhideWhenUsed/&gt;&lt;/w:style&gt;&lt;w:style w:type="table" w:default="1" w:styleId="a1"&gt;&lt;w:name w:val="Normal Table"/&gt;&lt;w:uiPriority w:val="99"/&gt;&lt;w:semiHidden/&gt;&lt;w:unhideWhenUsed/&gt;&lt;w:tblPr&gt;&lt;w:tblInd w:w="0" w:type="dxa"/&gt;&lt;w:tblCellMar&gt;&lt;w:top w:w="0" w:type="dxa"/&gt;&lt;w:left w:w="108" w:type="dxa"/&gt;&lt;w:bottom w:w="0" w:type="dxa"/&gt;&lt;w:right w:w="108" w:type="dxa"/&gt;&lt;/w:tblCellMar&gt;&lt;/w:tblPr&gt;&lt;/w:style&gt;&lt;w:style w:type="numbering" w:default="1" w:styleId="a2"&gt;&lt;w:name w:val="No List"/&gt;&lt;w:uiPriority w:val="99"/&gt;&lt;w:semiHidden/&gt;&lt;w:unhideWhenUsed/&gt;&lt;/w:style&gt;&lt;/w:styles&gt;&lt;/pkg:xmlData&gt;&lt;/pkg:part&gt;&lt;/pkg:package&gt;',
	'onePicture' : '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAIBAQIBAQICAgICAgICAwUDAwMDAwYEBAMFBwYHBwcGBwcICQsJCAgKCAcHCg0KCgsMDAwMBwkODw0MDgsMDAz/2wBDAQICAgMDAwYDAwYMCAcIDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAz/wAARCAEsASwDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9/KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo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',
	'twoPicture' : '/9j/4AAQSkZJRgABAQEAYABgAAD/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAEAAAAAAAD/2wBDAAIBAQIBAQICAgICAgICAwUDAwMDAwYEBAMFBwYHBwcGBwcICQsJCAgKCAcHCg0KCgsMDAwMBwkODw0MDgsMDAz/2wBDAQICAgMDAwYDAwYMCAcIDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAz/wAARCAEsASwDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9/KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo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y3t/W79T6m/4Ix/8Em/FP/BWn9q+z8J2K3mmeBNAaK+8Za/EoA0uyLHEUbMCv2mba6RKQeQzkFY3r+xb4IfBPwr+zf8ACTw/4F8E6LZ+HvCnhezSw0zT7VcR28S+55ZiSWZ2JZmZmYkkk/mX/wAE2v8Agpz/AME6f+CY37LGifDHwN8ctBkW1H2rWNWk0LU1utev2UCW6mxbHlsAKmSERUQEhcn3z/iJJ/Yj/wCi9aF/4J9U/wDkau6vVioqjS+Fat9339Fry36Xdk20cNGnKUvbVN3suy009Xa7+7W139xUV8O/8RJP7Ef/AEXrQv8AwT6p/wDI1H/EST+xH/0XrQv/AAT6p/8AI1ch1H3FRXw7/wARJP7Ef/RetC/8E+qf/I1H/EST+xH/ANF60L/wT6p/8jUAfcVFfDv/ABEk/sR/9F60L/wT6p/8jUf8RJP7Ef8A0XrQv/BPqn/yNQB9xUV8O/8AEST+xH/0XrQv/BPqn/yNX2J8LfidoXxq+GugeMPC+oJq3hvxRp8Gq6XepG8a3drPGskUgVwGAZGU4YA88gVXK7c1tP6/yJ5lfl6m9RRRUlBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFfx8/8ABzn/AMpvPjZ/120v/wBNVnX9g1fHv7Tv/BA/9k39sn43638R/iR8KP8AhJPGfiIxNqOo/wDCT6zZ/aDFEkKfure7jiXEcaD5UGcZOSSa550ZOvGotkmvvt/kdFOso0pwe8rfgz+MWiv6/f8AiFx/YT/6Ib/5efiD/wCTqP8AiFx/YT/6Ib/5efiD/wCTq6DnP5AqK/r9/wCIXH9hP/ohv/l5+IP/AJOo/wCIXH9hP/ohv/l5+IP/AJOoA/kCor+v3/iFx/YT/wCiG/8Al5+IP/k6j/iFx/YT/wCiG/8Al5+IP/k6gD+QKiv6/f8AiFx/YT/6Ib/5efiD/wCTqP8AiFx/YT/6Ib/5efiD/wCTqAP5AqK/r9/4hcf2E/8Aohv/AJefiD/5Oo/4hcf2E/8Aohv/AJefiD/5OoA/kY8AeC774k+O9F8O6XD9o1PX7+DTbSLOPMmmkWNFz7swFf3kfCzwLbfC/wCGPhzwzZqqWfh3S7bTIFAwFjhiWNR+Sivjv4Y/8G3P7Fvwc+JHh/xd4c+C66f4g8L6jb6tply3izXLhbe5gkWWJzHJetG+11U7XVlOMEEZFfcVdPto/V1SW/M2/uSj915ff93P7Juv7V7JWXzd3+UbfMKKKK5joCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA/9k=',
}
const oCustomXMLs		= {
	"withoutContent" : "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText></simpleText></documentData>\"",
	"date" : "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>" + oCustomXMLData.date + "</simpleText></documentData>\"",
	'checkboxTrue': "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>" + oCustomXMLData.checkboxTrue + "</simpleText></documentData>\"",
	'checkboxFalse': "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>" + oCustomXMLData.checkboxFalse + "</simpleText></documentData>\"",
	'checkbox0': "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>" + oCustomXMLData.checkbox0 + "</simpleText></documentData>\"",
	'checkbox1': "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>" + oCustomXMLData.checkbox1 + "</simpleText></documentData>\"",
	'checkboxMess': "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>" + oCustomXMLData.checkboxMess + "</simpleText></documentData>\"",
	'checkboxMess2': "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>" + oCustomXMLData.checkboxMess2 + "</simpleText></documentData>\"",
	'picture': "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<documentData xmlns=\"http://example.com/picture\"><simpleText>" + oCustomXMLData.onePicture + "</simpleText></documentData>\"",
	'notValidData': oCustomXMLData.onePicture,
	'checkboxTrueAnotherXML': "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<weather>" + oCustomXMLData.checkboxTrue + "</weather>",
}

function CreateContentControl(isInline, nPos)
{
	let para = AscTest.CreateParagraph();
	
	logicDocument.AddToContent(nPos, para);
	para.SetThisElementCurrent();
	return logicDocument.AddContentControl(isInline ? c_oAscSdtLevelType.Inline : c_oAscSdtLevelType.Block);
}
function CreateDataBindingForCC(contentControl, prefix, itemId, xpath, checkSum)
{
	let oDataBinding = new AscWord.DataBinding(
		prefix === undefined ? "xmlns:ns0='http://example.com/picture' " : prefix,
		itemId === undefined ? "{694325A8-B1C9-407B-A2C2-E2DD1740AA5E}" : itemId,
		xpath === undefined ? "/ns0:documentData[1]/ns0:simpleText[1]" : xpath,
		checkSum === undefined ? "Gt6wYg==" : checkSum,
	);
	
	contentControl.Pr.DataBinding = oDataBinding;
	contentControl.checkDataBinding();
}
function CreateCustomXMLForDocument(strContent, ItemId, Uri)
{
	let oXML					= new AscWord.CustomXml();
	
	oXML.addContentByXMLString(
		strContent !== undefined
			? strContent
			: oCustomXMLs.withoutContent
	);
	
	oXML.itemId		= ItemId === undefined
		? "{694325A8-B1C9-407B-A2C2-E2DD1740AA5E}"
		: ItemId;
	
	oXML.uri		= Uri === undefined
		? ['http://example.com/picture']
		: Uri;
	
	oXMLManager.add(oXML);
}
function SetDataToContentControl (oCC, strData)
{
	oXMLManager.setContentByDataBinding(oCC.Pr.DataBinding, strData);
	oCC.checkDataBinding();
}
function MoveCursorToCC(oCC, isToStart)
{
	let paragraph = oCC.GetLastParagraph();
	if (!paragraph || !(paragraph instanceof AscWord.Paragraph))
		return;
	
	paragraph.SetThisElementCurrent();
	
	if (false === isToStart)
		paragraph.MoveCursorToEndPos();
	else
		paragraph.MoveCursorToStartPos();
	
	return paragraph;
}

function CreateDateCC(nPos, isInline)
{
	let cc = CreateFilledContentControl(nPos, isInline);
	let dateTimePr = new AscWord.CSdtDatePickerPr();
	dateTimePr.SetDateFormat('DD-MM-YYYY');
	cc.ApplyDatePickerPr(dateTimePr, true);
	return cc;
}
function CreateCheckBoxCC(nPos, isInline)
{
	let cc = CreateFilledContentControl(nPos, isInline)
	let checkboxPr = new AscWord.CSdtCheckBoxPr();
	cc.ApplyCheckBoxPr(checkboxPr);
	return cc;
}
// поле со списком
function CreateComboBox(nPos, isInline)
{
	let cc = CreateFilledContentControl(nPos, isInline);
	let comboBoxPr = new AscWord.CSdtComboBoxPr();
	cc.ApplyComboBoxPr(comboBoxPr, true);
	return cc;
}
// раскрывающийся список
function CreateDropDown(nPos, isInline)
{
	let cc = CreateFilledContentControl(nPos, isInline);
	let comboBoxPr = new AscWord.CSdtComboBoxPr();
	cc.ApplyDropDownListPr(comboBoxPr, true);
	return cc;
}
function CreatePicture(nPos, isInline)
{
	let cc = CreateFilledContentControl(nPos, isInline);
	let picturePr = new AscWord.CSdtPictureFormPr();
	cc.ApplyPicturePr(picturePr);
	return cc;
}
function CreateText(nPos, isInline)
{
	let cc = CreateFilledContentControl(nPos, isInline);
	cc.SetContentControlText(true)
	return cc;
}

function CreateFilledContentControl(nPos, isInline)
{
	return CreateContentControl(isInline, nPos)
}
function CheckContentParagraph(assert, oContentArr, arrSample)
{
	for (let i = 0; i < arrSample.length; i++)
	{
		let oCurStr = arrSample[i];
		let oCurContent = oContentArr[i].GetText();
		assert.strictEqual(oCurStr, oCurContent, oCurContent);
	}
}

function getCheck(strContent)
{
	return '<?xml version="1.0" encoding="UTF-8"?>\n' +
		'<documentData xmlns="http://example.com/picture">' +
		'<simpleText>'+ strContent +'</simpleText>' +
		'</documentData>';
}
function Reset()
{
	AscTest.ClearDocument();
	oXMLManager.xml = []
}
