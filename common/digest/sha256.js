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

(function () {

	'use strict';

	window["AscCommon"] = window.AscCommon = (window["AscCommon"] || {});

	var EXTRA = [-2147483648, 8388608, 32768, 128];
	var SHIFT = [24, 16, 8, 0];
	var K = [
		0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
		0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
		0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
		0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
		0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
		0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
		0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
		0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
	];

	function Sha256() {
		this.blocks = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

		this.h0 = 0x6a09e667;
		this.h1 = 0xbb67ae85;
		this.h2 = 0x3c6ef372;
		this.h3 = 0xa54ff53a;
		this.h4 = 0x510e527f;
		this.h5 = 0x9b05688c;
		this.h6 = 0x1f83d9ab;
		this.h7 = 0x5be0cd19;

		this.block = this.start = this.bytes = this.hBytes = 0;
		this.finalized = this.hashed = false;
		this.first = true;
	}

	Sha256.prototype.update = function (message) {
		if (this.finalized) {
			return;
		}

		let code, index = 0, i, length = message.length, blocks = this.blocks;
		let isString = (typeof message === "string");

		while (index < length) {
			if (this.hashed) {
				this.hashed = false;
				blocks[0] = this.block;
				blocks[16] = blocks[1] = blocks[2] = blocks[3] =
					blocks[4] = blocks[5] = blocks[6] = blocks[7] =
						blocks[8] = blocks[9] = blocks[10] = blocks[11] =
							blocks[12] = blocks[13] = blocks[14] = blocks[15] = 0;
			}

			if (!isString) {
				for (i = this.start; index < length && i < 64; ++index)
				{
					blocks[i >> 2] |= message[index] << SHIFT[i++ & 3];
				}
			} else {
				for (i = this.start; index < length && i < 64; ++index) {
					code = message.charCodeAt(index);
					if (code < 0x80) {
						blocks[i >> 2] |= code << SHIFT[i++ & 3];
					} else if (code < 0x800) {
						blocks[i >> 2] |= (0xc0 | (code >> 6)) << SHIFT[i++ & 3];
						blocks[i >> 2] |= (0x80 | (code & 0x3f)) << SHIFT[i++ & 3];
					} else if (code < 0xd800 || code >= 0xe000) {
						blocks[i >> 2] |= (0xe0 | (code >> 12)) << SHIFT[i++ & 3];
						blocks[i >> 2] |= (0x80 | ((code >> 6) & 0x3f)) << SHIFT[i++ & 3];
						blocks[i >> 2] |= (0x80 | (code & 0x3f)) << SHIFT[i++ & 3];
					} else {
						code = 0x10000 + (((code & 0x3ff) << 10) | (message.charCodeAt(++index) & 0x3ff));
						blocks[i >> 2] |= (0xf0 | (code >> 18)) << SHIFT[i++ & 3];
						blocks[i >> 2] |= (0x80 | ((code >> 12) & 0x3f)) << SHIFT[i++ & 3];
						blocks[i >> 2] |= (0x80 | ((code >> 6) & 0x3f)) << SHIFT[i++ & 3];
						blocks[i >> 2] |= (0x80 | (code & 0x3f)) << SHIFT[i++ & 3];
					}
				}
			}

			this.lastByteIndex = i;
			this.bytes += i - this.start;
			if (i >= 64) {
				this.block = blocks[16];
				this.start = i - 64;
				this.hash();
				this.hashed = true;
			} else {
				this.start = i;
			}
		}
		if (this.bytes > 4294967295) {
			this.hBytes += this.bytes / 4294967296 << 0;
			this.bytes = this.bytes % 4294967296;
		}
		return this;
	};

	Sha256.prototype.finalize = function () {
		if (this.finalized) {
			return;
		}
		this.finalized = true;
		var blocks = this.blocks, i = this.lastByteIndex;
		blocks[16] = this.block;
		blocks[i >> 2] |= EXTRA[i & 3];
		this.block = blocks[16];
		if (i >= 56) {
			if (!this.hashed) {
				this.hash();
			}
			blocks[0] = this.block;
			blocks[16] = blocks[1] = blocks[2] = blocks[3] =
				blocks[4] = blocks[5] = blocks[6] = blocks[7] =
					blocks[8] = blocks[9] = blocks[10] = blocks[11] =
						blocks[12] = blocks[13] = blocks[14] = blocks[15] = 0;
		}
		blocks[14] = this.hBytes << 3 | this.bytes >>> 29;
		blocks[15] = this.bytes << 3;
		this.hash();
	};

	Sha256.prototype.hash = function () {
		var a = this.h0, b = this.h1, c = this.h2, d = this.h3, e = this.h4, f = this.h5, g = this.h6,
			h = this.h7, blocks = this.blocks, j, s0, s1, maj, t1, t2, ch, ab, da, cd, bc;

		for (j = 16; j < 64; ++j) {
			// rightrotate
			t1 = blocks[j - 15];
			s0 = ((t1 >>> 7) | (t1 << 25)) ^ ((t1 >>> 18) | (t1 << 14)) ^ (t1 >>> 3);
			t1 = blocks[j - 2];
			s1 = ((t1 >>> 17) | (t1 << 15)) ^ ((t1 >>> 19) | (t1 << 13)) ^ (t1 >>> 10);
			blocks[j] = blocks[j - 16] + s0 + blocks[j - 7] + s1 << 0;
		}

		bc = b & c;
		for (j = 0; j < 64; j += 4) {
			if (this.first) {
				if (this.is224) {
					ab = 300032;
					t1 = blocks[0] - 1413257819;
					h = t1 - 150054599 << 0;
					d = t1 + 24177077 << 0;
				} else {
					ab = 704751109;
					t1 = blocks[0] - 210244248;
					h = t1 - 1521486534 << 0;
					d = t1 + 143694565 << 0;
				}
				this.first = false;
			} else {
				s0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
				s1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
				ab = a & b;
				maj = ab ^ (a & c) ^ bc;
				ch = (e & f) ^ (~e & g);
				t1 = h + s1 + ch + K[j] + blocks[j];
				t2 = s0 + maj;
				h = d + t1 << 0;
				d = t1 + t2 << 0;
			}
			s0 = ((d >>> 2) | (d << 30)) ^ ((d >>> 13) | (d << 19)) ^ ((d >>> 22) | (d << 10));
			s1 = ((h >>> 6) | (h << 26)) ^ ((h >>> 11) | (h << 21)) ^ ((h >>> 25) | (h << 7));
			da = d & a;
			maj = da ^ (d & b) ^ ab;
			ch = (h & e) ^ (~h & f);
			t1 = g + s1 + ch + K[j + 1] + blocks[j + 1];
			t2 = s0 + maj;
			g = c + t1 << 0;
			c = t1 + t2 << 0;
			s0 = ((c >>> 2) | (c << 30)) ^ ((c >>> 13) | (c << 19)) ^ ((c >>> 22) | (c << 10));
			s1 = ((g >>> 6) | (g << 26)) ^ ((g >>> 11) | (g << 21)) ^ ((g >>> 25) | (g << 7));
			cd = c & d;
			maj = cd ^ (c & a) ^ da;
			ch = (g & h) ^ (~g & e);
			t1 = f + s1 + ch + K[j + 2] + blocks[j + 2];
			t2 = s0 + maj;
			f = b + t1 << 0;
			b = t1 + t2 << 0;
			s0 = ((b >>> 2) | (b << 30)) ^ ((b >>> 13) | (b << 19)) ^ ((b >>> 22) | (b << 10));
			s1 = ((f >>> 6) | (f << 26)) ^ ((f >>> 11) | (f << 21)) ^ ((f >>> 25) | (f << 7));
			bc = b & c;
			maj = bc ^ (b & d) ^ cd;
			ch = (f & g) ^ (~f & h);
			t1 = e + s1 + ch + K[j + 3] + blocks[j + 3];
			t2 = s0 + maj;
			e = a + t1 << 0;
			a = t1 + t2 << 0;
		}

		this.h0 = this.h0 + a << 0;
		this.h1 = this.h1 + b << 0;
		this.h2 = this.h2 + c << 0;
		this.h3 = this.h3 + d << 0;
		this.h4 = this.h4 + e << 0;
		this.h5 = this.h5 + f << 0;
		this.h6 = this.h6 + g << 0;
		this.h7 = this.h7 + h << 0;
	};

	Sha256.prototype.getResult = function () {
		this.finalize();

		let result = new Uint8Array(32);
		if (typeof DataView !== 'undefined')
		{
			var dataView = new DataView(result.buffer);
			dataView.setUint32(0, this.h0);
			dataView.setUint32(4, this.h1);
			dataView.setUint32(8, this.h2);
			dataView.setUint32(12, this.h3);
			dataView.setUint32(16, this.h4);
			dataView.setUint32(20, this.h5);
			dataView.setUint32(24, this.h6);
			dataView.setUint32(28, this.h7);
			return result;
		}

		let h0 = this.h0, h1 = this.h1, h2 = this.h2, h3 = this.h3, h4 = this.h4, h5 = this.h5,
			h6 = this.h6, h7 = this.h7;

		result[0]  = ((h0 >> 28) & 0x0F) + ((h0 >> 24) & 0x0F);
		result[1]  = ((h0 >> 20) & 0x0F) + ((h0 >> 24) & 0x0F);
		result[2]  = ((h0 >> 12) & 0x0F) + ((h0 >> 24) & 0x0F);
		result[3]  = ((h0 >> 4) & 0x0F)  + ((h0 >> 0) & 0x0F);
		result[4]  = ((h1 >> 28) & 0x0F) + ((h1 >> 24) & 0x0F);
		result[5]  = ((h1 >> 20) & 0x0F) + ((h1 >> 24) & 0x0F);
		result[6]  = ((h1 >> 12) & 0x0F) + ((h1 >> 24) & 0x0F);
		result[7]  = ((h1 >> 4) & 0x0F)  + ((h1 >> 0) & 0x0F);
		result[8]  = ((h2 >> 28) & 0x0F) + ((h2 >> 24) & 0x0F);
		result[9]  = ((h2 >> 20) & 0x0F) + ((h2 >> 24) & 0x0F);
		result[10] = ((h2 >> 12) & 0x0F) + ((h2 >> 24) & 0x0F);
		result[11] = ((h2 >> 4) & 0x0F)  + ((h2 >> 0) & 0x0F);
		result[12] = ((h3 >> 28) & 0x0F) + ((h3 >> 24) & 0x0F);
		result[13] = ((h3 >> 20) & 0x0F) + ((h3 >> 24) & 0x0F);
		result[14] = ((h3 >> 12) & 0x0F) + ((h3 >> 24) & 0x0F);
		result[15] = ((h3 >> 4) & 0x0F)  + ((h3 >> 0) & 0x0F);
		result[16] = ((h4 >> 28) & 0x0F) + ((h4 >> 24) & 0x0F);
		result[17] = ((h4 >> 20) & 0x0F) + ((h4 >> 24) & 0x0F);
		result[18] = ((h4 >> 12) & 0x0F) + ((h4 >> 24) & 0x0F);
		result[19] = ((h4 >> 4) & 0x0F)  + ((h4 >> 0) & 0x0F);
		result[20] = ((h5 >> 28) & 0x0F) + ((h5 >> 24) & 0x0F);
		result[21] = ((h5 >> 20) & 0x0F) + ((h5 >> 24) & 0x0F);
		result[22] = ((h5 >> 12) & 0x0F) + ((h5 >> 24) & 0x0F);
		result[23] = ((h5 >> 4) & 0x0F)  + ((h5 >> 0) & 0x0F);
		result[24] = ((h6 >> 28) & 0x0F) + ((h6 >> 24) & 0x0F);
		result[25] = ((h6 >> 20) & 0x0F) + ((h6 >> 24) & 0x0F);
		result[26] = ((h6 >> 12) & 0x0F) + ((h6 >> 24) & 0x0F);
		result[27] = ((h6 >> 4) & 0x0F)  + ((h6 >> 0) & 0x0F);
		result[28] = ((h7 >> 28) & 0x0F) + ((h7 >> 24) & 0x0F);
		result[29] = ((h7 >> 20) & 0x0F) + ((h7 >> 24) & 0x0F);
		result[30] = ((h7 >> 12) & 0x0F) + ((h7 >> 24) & 0x0F);
		result[31] = ((h7 >> 4) & 0x0F)  + ((h7 >> 0) & 0x0F);

		return result;
	};

	window.AscCommon.Digest = window.AscCommon.Digest || {};

	window.AscCommon.Digest.sha256 = function(data, offset, length) {
		let engine = new Sha256();
		if (undefined !== offset || undefined !== length)
			engine.update(new Uint8Array(data.buffer, offset || 0, length || data.length));
		else
			engine.update(data);
		return engine.getResult();
	};

})();
