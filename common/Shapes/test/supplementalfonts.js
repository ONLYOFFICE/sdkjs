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

/*
 * Regression tests for the <a:font script="..."/> entries of a theme font collection
 * (record 3 of a FontCollection) through the PPTY writer and reader.
 *
 * From this directory:  node supplementalfonts.js   (exit code 1 if anything fails)
 * Or open supplementalfonts.html and read the browser console.
 *
 * The record layout under test is the one the C++ side writes, in
 * core/OOXML/PPTXFormat/Logic/{FontCollection,SupplementalFont}.cpp: record 3 holds a
 * count followed by that many subrecords of type 0, each an attribute block with
 * 0 = script and 1 = typeface.
 */
(function () {
	"use strict";

	/* Decided up front: the node branch below defines window, so this cannot be
	   re-tested later on. */
	var isNode = (typeof window === "undefined");
	var root;
	if (!isNode) {
		/* The html page has already set up the namespaces and loaded the sources. */
		root = window;
	} else {
		/* Under node, give the sources the globals they write into, then load them. */
		root = global;
		root.window = root;
		root.AscCommon = root.AscCommon || {};
		root.Asc = root.Asc || {};
		root.AscFormat = root.AscFormat || {};
		root.AscDFH = root.AscDFH || {};
		root.AscCommon.c_dScalePPTXSizes = 36000;
		root.Asc.c_oAscShdClear = 0;
		root.Asc.c_oAscColor = {};
		root.Asc.c_oAscFill = {};
		require("../../SerializeCommonWordExcel.js");
		require("../SerializeWriter.js");
		require("../Serialize.js");
	}

	var failures = 0, total = 0;

	function log(s) {
		if (typeof console !== "undefined" && console.log) {
			console.log(s);
		}
	}

	function check(name, actual, expected) {
		total++;
		var ok = (actual === expected);
		if (!ok) {
			failures++;
		}
		log((ok ? "PASS  " : "FAIL  ") + name + "   got " + actual + ", expected " + expected);
	}

	/* AscFormat.FontCollection does not load outside the editor, so the tests use a
	   stand-in carrying only what the reader and the writer touch. The code under test
	   is the real writer and the real reader. */
	function TestCollection() {
		this.latin = null;
		this.ea = null;
		this.cs = null;
		this.supplementalFont = [];
	}
	TestCollection.prototype.setLatin = function (v) { this.latin = v; };
	TestCollection.prototype.setEA = function (v) { this.ea = v; };
	TestCollection.prototype.setCS = function (v) { this.cs = v; };
	TestCollection.prototype.addSupplementalFont = function (script, typeface) {
		this.supplementalFont.push({ script: script, typeface: typeface });
	};

	function makeCollection(fonts) {
		var coll = new TestCollection();
		coll.latin = "Cambria";
		coll.ea = "";
		coll.cs = "";
		for (var i = 0; i < fonts.length; ++i) {
			coll.addSupplementalFont(fonts[i][0], fonts[i][1]);
		}
		return coll;
	}

	function describe(coll) {
		var parts = [];
		for (var i = 0; i < coll.supplementalFont.length; ++i) {
			parts.push(coll.supplementalFont[i].script + "=" + coll.supplementalFont[i].typeface);
		}
		return parts.join(",");
	}

	/* Writes the collection as record 0, plus an optional marker record after it, and
	   returns the bytes. */
	function write(coll, markerType) {
		var writer = new root.AscCommon.CBinaryFileWriter();
		writer.WriteRecord1(0, coll, writer.WriteFontCollection);
		if (undefined !== markerType) {
			writer.StartRecord(markerType);
			writer.WriteULong(0);
			writer.EndRecord();
		}
		return writer.GetData();
	}

	/* The same record, but with an attribute this build does not know placed before the
	   two it does. Hand-built, because the writer only ever emits known attributes. */
	function writeWithUnknownAttribute(markerType) {
		var writer = new root.AscCommon.CBinaryFileWriter();
		writer.StartRecord(0);
		writer.WriteRecord1(0, { Name: "Cambria", Index: -1 }, writer.WriteTextFontTypeface);
		writer.StartRecord(3);
		writer.WriteULong(1);
		writer.StartRecord(0);
		writer.WriteUChar(root.AscCommon.g_nodeAttributeStart);
		writer._WriteString1(9, "unknown-attribute-value");
		writer._WriteString1(0, "Jpan");
		writer._WriteString1(1, "MS PGothic");
		writer.WriteUChar(root.AscCommon.g_nodeAttributeEnd);
		writer.EndRecord();
		writer.EndRecord();
		writer.EndRecord();
		writer.StartRecord(markerType);
		writer.WriteULong(0);
		writer.EndRecord();
		return writer.GetData();
	}

	/* The same again, but with the attribute block left unterminated. Past the end of the
	   data FileStream.GetUChar keeps returning 0, so a parser that reads on until it sees
	   an end marker never finishes on input like this. */
	function writeWithoutEndMarker(markerType) {
		var writer = new root.AscCommon.CBinaryFileWriter();
		writer.StartRecord(0);
		writer.WriteRecord1(0, { Name: "Cambria", Index: -1 }, writer.WriteTextFontTypeface);
		writer.StartRecord(3);
		writer.WriteULong(1);
		writer.StartRecord(0);
		writer.WriteUChar(root.AscCommon.g_nodeAttributeStart);
		writer._WriteString1(9, "aaaaaaaaaaaaaaaa");
		writer.EndRecord();
		writer.EndRecord();
		writer.EndRecord();
		writer.StartRecord(markerType);
		writer.WriteULong(0);
		writer.EndRecord();
		return writer.GetData();
	}

	/* Reads the collection back. Returns the reader so the caller can look at what
	   follows the record. The byte reads are capped, so a parse that does not terminate
	   fails a test rather than hanging the run. */
	function read(bytes, coll) {
		var loader = new root.AscCommon.BinaryPPTYLoader();
		var stream = new root.AscCommon.FileStream(bytes, bytes.length);
		loader.stream = stream;
		stream.GetUChar(); /* the record type written above */

		var reads = 0;
		var readByte = stream.GetUChar;
		stream.GetUChar = function () {
			if (++reads > 100000) {
				throw new Error("the parse did not terminate");
			}
			return readByte.apply(stream, arguments);
		};
		try {
			loader.ReadFontCollection(coll);
		} finally {
			stream.GetUChar = readByte;
		}
		return loader;
	}

	function run() {
		var fonts = [["Jpan", "ＭＳ Ｐゴシック"],
			["Hang", "맑은 고딕"],
			["Hans", "宋体"],
			["Hant", "新細明體"],
			["Arab", "Times New Roman"]];

		/* The case this was written for: the per-script fonts of a theme have to come
		   back out of the round trip. Without them an application has nothing but the
		   latin font to fall back on for East Asian text. */
		var src = makeCollection(fonts);
		var dst = new TestCollection();
		read(write(src), dst);
		check("script fonts survive the round trip", describe(dst), describe(src));
		check("script font count", dst.supplementalFont.length, fonts.length);

		/* The three named fonts must be unaffected. */
		check("latin survives", dst.latin, "Cambria");
		check("ea survives", dst.ea, "");
		check("cs survives", dst.cs, "");

		/* A collection with no script fonts stays empty, and must not upset the reader. */
		var empty = makeCollection([]);
		var emptyBack = new TestCollection();
		read(write(empty), emptyBack);
		check("no script fonts stays empty", emptyBack.supplementalFont.length, 0);
		check("latin survives without script fonts", emptyBack.latin, "Cambria");

		/* The reader has to leave the stream at the end of record 3, or everything after
		   the font collection is read from the wrong offset. */
		var marked = read(write(src, 7), new TestCollection());
		check("stream is positioned after the record", marked.stream.GetUChar(), 7);

		/* An attribute this build does not know leaves the rest of that entry misread -
		   the two typefaces after it may or may not survive, so nothing is asserted about
		   them - but the damage is contained: the entry is still counted, the named fonts
		   are untouched, and the record that follows the font collection is still found.
		   Reaching the end of the entry does depend on the parse landing on the end
		   marker, which is how every attribute block in these files is read. */
		var oddColl = new TestCollection();
		var oddLoader = read(writeWithUnknownAttribute(7), oddColl);
		check("unknown attribute: entry still counted", oddColl.supplementalFont.length, 1);
		check("unknown attribute: latin untouched", oddColl.latin, "Cambria");
		check("unknown attribute: next record still found", oddLoader.stream.GetUChar(), 7);

		/* And the reader has to give up on an attribute block that never ends, rather than
		   look for an end marker that is not there. */
		var truncColl = new TestCollection();
		var truncLoader = null;
		try {
			truncLoader = read(writeWithoutEndMarker(7), truncColl);
		} catch (e) {
			log("  (" + e.message + ")");
		}
		check("unterminated block: the parse stops", null !== truncLoader, true);
		check("unterminated block: next record still found",
			truncLoader ? truncLoader.stream.GetUChar() : -1, 7);

		/* The record layout itself, as the C++ writer produces it: type 3, then the
		   record length, then the number of entries. */
		var bytes = write(src);
		var at = -1;
		for (var i = 0; i + 8 < bytes.length; ++i) {
			if (3 === bytes[i] && fonts.length === bytes[i + 5] &&
				0 === bytes[i + 6] && 0 === bytes[i + 7] && 0 === bytes[i + 8]) {
				at = i;
				break;
			}
		}
		check("record 3 carries the entry count", at >= 0, true);

		log("");
		log(failures ? (failures + " of " + total + " failed") : ("all " + total + " passed"));
		return failures;
	}

	if (isNode) {
		process.exitCode = run() ? 1 : 0;
	} else {
		root.runSupplementalFontsTests = run;
	}
})();
