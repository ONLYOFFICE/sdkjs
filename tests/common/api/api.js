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

	QUnit.module("Test api for the all editors");


	QUnit.test("Test asc_getUrlType", function (assert)
	{
		const editor = new AscCommon.baseEditorsApi({});

		let test = [
			["http://foo.com/blah_blah", AscCommon.c_oAscUrlType.Http],
			["http://foo.com/blah_blah_(wikipedia)_(again)", AscCommon.c_oAscUrlType.Http],
			["https://www.example.com/foo/?bar=baz&inga=42&quux", AscCommon.c_oAscUrlType.Http],
			["http://userid:password@example.com:8080", AscCommon.c_oAscUrlType.Http],
			["http://userid@example.com:8080/", AscCommon.c_oAscUrlType.Http],
			["http://142.42.1.1", AscCommon.c_oAscUrlType.Http],
			["http://142.42.1.1:8080/", AscCommon.c_oAscUrlType.Http],
			["http://foo.com/blah_(wikipedia)_blah#cite-1", AscCommon.c_oAscUrlType.Http],
			["http://foo.bar/?q=Test%20URL-encoded%20stuff", AscCommon.c_oAscUrlType.Http],
			["http://a.b-c.de", AscCommon.c_oAscUrlType.Http],
			["ftp://public.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["ftp://user001:secretpassword@private.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["ftps://user001:secretpassword@private.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["http://a.b-c.de", AscCommon.c_oAscUrlType.Http],
			["http://مثال.إختبار", AscCommon.c_oAscUrlType.Unsafe],//todo Http
			["http://фывап.ролдж", AscCommon.c_oAscUrlType.Http],
			["http://", AscCommon.c_oAscUrlType.Unsafe],//todo Invalid ?
			["http:///a", AscCommon.c_oAscUrlType.Unsafe],
			["http://.www.foo.bar/", AscCommon.c_oAscUrlType.Unsafe],

			["mysite@ourearth.com", AscCommon.c_oAscUrlType.Http],//todo Email
			["my.ownsite@ourearth.org", AscCommon.c_oAscUrlType.Email],
			["mysite@you.me.net", AscCommon.c_oAscUrlType.Http],//todo Email
			["mysite@.com.my", AscCommon.c_oAscUrlType.Email],//todo Invalid
			["@you.me.net", AscCommon.c_oAscUrlType.Http],//todo Invalid
			[".mysite@mysite.org", AscCommon.c_oAscUrlType.Email],//todo Invalid
			["mysite()*@gmail.com", AscCommon.c_oAscUrlType.Invalid],

			["smb://192.168.56.1/e/Testfolder/TestFile.docx", AscCommon.c_oAscUrlType.Unsafe],

			["tessa://tessaclient.EPD/?Action=OpenCard&ID=c40076f5-daa9-4929-8f66-d3fd6ae2dcb1", AscCommon.c_oAscUrlType.Unsafe],
			["joplin://x-callback-url/openFolder?id=1234", AscCommon.c_oAscUrlType.Unsafe],

			["file://localhost/etc/fstab", AscCommon.c_oAscUrlType.Unsafe],
			["file:///etc/fstab", AscCommon.c_oAscUrlType.Unsafe],
			["file://localhost/c:/WINDOWS/clock.avi", AscCommon.c_oAscUrlType.Unsafe],
			["file:///c:/WINDOWS/clock.avi", AscCommon.c_oAscUrlType.Unsafe],
			["file://\"C:\\Users\\User\\Documents\\About.pdf\"", AscCommon.c_oAscUrlType.Invalid],
			["file://'C:\\Users\\User\\Documents\\About.pdf'", AscCommon.c_oAscUrlType.Invalid],

			["/home/user/123.txt", AscCommon.c_oAscUrlType.Invalid],
			["123.txt", AscCommon.c_oAscUrlType.Http],
			["../../123.txt", AscCommon.c_oAscUrlType.Invalid],
		];
		for(let i = 0; i < test.length; ++i) {
			assert.strictEqual(editor.asc_getUrlType(test[i][0]), test[i][1], "Check " + test[i][0]);
		}
	});

	QUnit.test("Test asc_getUrlType desktop", function (assert)
	{
		let oldAscDesktopEditor = window["AscDesktopEditor"];
		window["AscDesktopEditor"] = {"IsLocalFile": function(){return true;}};

		const editor = new AscCommon.baseEditorsApi({});

		let test = [
			["http://foo.com/blah_blah", AscCommon.c_oAscUrlType.Http],
			["http://foo.com/blah_blah_(wikipedia)_(again)", AscCommon.c_oAscUrlType.Http],
			["https://www.example.com/foo/?bar=baz&inga=42&quux", AscCommon.c_oAscUrlType.Http],
			["http://userid:password@example.com:8080", AscCommon.c_oAscUrlType.Http],
			["http://userid@example.com:8080/", AscCommon.c_oAscUrlType.Http],
			["http://142.42.1.1", AscCommon.c_oAscUrlType.Http],
			["http://142.42.1.1:8080/", AscCommon.c_oAscUrlType.Http],
			["http://foo.com/blah_(wikipedia)_blah#cite-1", AscCommon.c_oAscUrlType.Http],
			["http://foo.bar/?q=Test%20URL-encoded%20stuff", AscCommon.c_oAscUrlType.Http],
			["http://a.b-c.de", AscCommon.c_oAscUrlType.Http],
			["ftp://public.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["ftp://user001:secretpassword@private.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["ftps://user001:secretpassword@private.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["http://a.b-c.de", AscCommon.c_oAscUrlType.Http],
			["http://مثال.إختبار", AscCommon.c_oAscUrlType.Unsafe],//todo Http
			["http://фывап.ролдж", AscCommon.c_oAscUrlType.Http],
			["http://", AscCommon.c_oAscUrlType.Unsafe],//todo Invalid ?
			["http:///a", AscCommon.c_oAscUrlType.Unsafe],
			["http://.www.foo.bar/", AscCommon.c_oAscUrlType.Unsafe],

			["mysite@ourearth.com", AscCommon.c_oAscUrlType.Http],//todo Email
			["my.ownsite@ourearth.org", AscCommon.c_oAscUrlType.Email],
			["mysite@you.me.net", AscCommon.c_oAscUrlType.Http],//todo Email
			["mysite@.com.my", AscCommon.c_oAscUrlType.Email],//todo Invalid
			["@you.me.net", AscCommon.c_oAscUrlType.Http],//todo Invalid
			[".mysite@mysite.org", AscCommon.c_oAscUrlType.Email],//todo Invalid
			["mysite()*@gmail.com", AscCommon.c_oAscUrlType.Invalid],

			["smb://192.168.56.1/e/Testfolder/TestFile.docx", AscCommon.c_oAscUrlType.Unsafe],

			["tessa://tessaclient.EPD/?Action=OpenCard&ID=c40076f5-daa9-4929-8f66-d3fd6ae2dcb1", AscCommon.c_oAscUrlType.Unsafe],

			["file://localhost/etc/fstab", AscCommon.c_oAscUrlType.Unsafe],
			["file:///etc/fstab", AscCommon.c_oAscUrlType.Unsafe],
			["file://localhost/c:/WINDOWS/clock.avi", AscCommon.c_oAscUrlType.Unsafe],
			["file:///c:/WINDOWS/clock.avi", AscCommon.c_oAscUrlType.Unsafe],
			["file://\"C:\\Users\\User\\Documents\\About.pdf\"", AscCommon.c_oAscUrlType.Invalid],
			["file://'C:\\Users\\User\\Documents\\About.pdf'", AscCommon.c_oAscUrlType.Invalid],

			["/home/user/123.txt", AscCommon.c_oAscUrlType.Unsafe],
			["123.txt", AscCommon.c_oAscUrlType.Unsafe],
			["../../123.txt", AscCommon.c_oAscUrlType.Unsafe],
		];
		for(let i = 0; i < test.length; ++i) {
			assert.strictEqual(editor.asc_getUrlType(test[i][0]), test[i][1], "Check " + test[i][0]);
		}
		window["AscDesktopEditor"] = oldAscDesktopEditor;
	});
});
