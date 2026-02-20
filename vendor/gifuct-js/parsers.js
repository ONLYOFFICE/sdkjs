/*
 * gifuct-js - https://github.com/matt-way/gifuct-js
 * Copyright (c) 2015 Matt Way
 * Licensed under the MIT License
 *
 * Modified for ONLYOFFICE: wrapped in IIFE, added GIFUCT namespace
 */

'use strict';

(function (window, undefined) {

	window["GIFUCT"] = window.GIFUCT = window.GIFUCT || {};



	// a set of common parsers used with DataParser

	var Parsers = {
		// read a byte
		readByte: function() {
			return function(stream) {
				return stream.readByte();
			};
		},
		// read an array of bytes
		readBytes: function(length) {
			return function(stream) {
				return stream.readBytes(length);
			};
		},
		// read a string from bytes
		readString: function(length) {
			return function(stream) {
				return stream.readString(length);
			};
		},
		// read an unsigned int (with endian)
		readUnsigned: function(littleEndian) {
			return function(stream) {
				return stream.readUnsigned(littleEndian);
			};
		},
		// read an array of byte sets
		readArray: function(size, countFunc) {
			return function(stream, obj, parent) {
				var count = countFunc(stream, obj, parent);
				var arr = new Array(count);
				for (var i = 0; i < count; i++) {
					arr[i] = stream.readBytes(size);
				}
				return arr;
			};
		}
	};

	window.GIFUCT.Parsers = Parsers;	
	window.GIFUCT.GetParsers = function() {
		return Parsers;
	};
})(window);
