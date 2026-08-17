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

(function (window, undefined) {

	window["AscCommonSlide"] = window['AscCommonSlide'] || {};

	const MAX_DELAY_CS = 65535;
	const MAX_COLORS = 256;
	const HISTOGRAM_SIZE = 32768;
	const DISPOSAL_LEAVE = 1;
	const LZW_MAX_CODES = 4096;
	const HASH_SIZE = 5003;

	function CGIFStream() {
		this.data = new Uint8Array(65536);
		this.pos = 0;
	}
	CGIFStream.prototype.reserve = function (count) {
		if (this.pos + count <= this.data.length) {
			return;
		}
		let length = this.data.length;
		while (length < this.pos + count) {
			length *= 2;
		}
		const data = new Uint8Array(length);
		data.set(this.data.subarray(0, this.pos));
		this.data = data;
	};
	CGIFStream.prototype.writeByte = function (value) {
		this.reserve(1);
		this.data[this.pos++] = value & 0xFF;
	};
	CGIFStream.prototype.writeShort = function (value) {
		this.reserve(2);
		this.data[this.pos++] = value & 0xFF;
		this.data[this.pos++] = (value >> 8) & 0xFF;
	};
	CGIFStream.prototype.writeBytes = function (bytes, count) {
		const length = undefined === count ? bytes.length : count;
		this.reserve(length);
		for (let index = 0; index < length; ++index) {
			this.data[this.pos++] = bytes[index] & 0xFF;
		}
	};
	CGIFStream.prototype.writeString = function (text) {
		this.reserve(text.length);
		for (let index = 0; index < text.length; ++index) {
			this.data[this.pos++] = text.charCodeAt(index) & 0xFF;
		}
	};
	CGIFStream.prototype.readShort = function (position) {
		return this.data[position] | (this.data[position + 1] << 8);
	};
	CGIFStream.prototype.patchShort = function (position, value) {
		this.data[position] = value & 0xFF;
		this.data[position + 1] = (value >> 8) & 0xFF;
	};
	CGIFStream.prototype.toUint8Array = function () {
		return this.data.subarray(0, this.pos);
	};

	function CGIFPalette() {
		this.colors = new Uint8Array(MAX_COLORS * 3);
		this.count = 0;
		this.lut = new Int16Array(HISTOGRAM_SIZE);
	}
	CGIFPalette.prototype.getBits = function () {
		let bits = 1;
		while ((1 << bits) < this.count) {
			++bits;
		}
		return bits;
	};
	CGIFPalette.prototype.getTableSize = function () {
		return 1 << this.getBits();
	};

	function CGIFQuantizer() {
		this.counts = new Uint32Array(HISTOGRAM_SIZE);
		this.sumR = new Uint32Array(HISTOGRAM_SIZE);
		this.sumG = new Uint32Array(HISTOGRAM_SIZE);
		this.sumB = new Uint32Array(HISTOGRAM_SIZE);
		this.usedBins = new Int32Array(HISTOGRAM_SIZE);
		this.usedCount = 0;
	}
	CGIFQuantizer.prototype.reset = function () {
		for (let index = 0; index < this.usedCount; ++index) {
			const bin = this.usedBins[index];
			this.counts[bin] = 0;
			this.sumR[bin] = 0;
			this.sumG[bin] = 0;
			this.sumB[bin] = 0;
		}
		this.usedCount = 0;
	};
	CGIFQuantizer.prototype.addPixel = function (r, g, b) {
		const bin = ((r >> 3) << 10) | ((g >> 3) << 5) | (b >> 3);
		if (0 === this.counts[bin]) {
			this.usedBins[this.usedCount++] = bin;
		}
		++this.counts[bin];
		this.sumR[bin] += r;
		this.sumG[bin] += g;
		this.sumB[bin] += b;
	};
	CGIFQuantizer.prototype.getBinColor = function (bin, color) {
		const count = this.counts[bin];
		color[0] = (this.sumR[bin] / count) | 0;
		color[1] = (this.sumG[bin] / count) | 0;
		color[2] = (this.sumB[bin] / count) | 0;
		return color;
	};
	CGIFQuantizer.prototype.buildPalette = function (palette, maxColors) {
		const color = [0, 0, 0];
		palette.count = 0;
		if (this.usedCount <= maxColors) {
			for (let index = 0; index < this.usedCount; ++index) {
				const bin = this.usedBins[index];
				this.getBinColor(bin, color);
				palette.colors[palette.count * 3] = color[0];
				palette.colors[palette.count * 3 + 1] = color[1];
				palette.colors[palette.count * 3 + 2] = color[2];
				palette.lut[bin] = palette.count;
				++palette.count;
			}
			return;
		}
		this.buildPaletteMedianCut(palette, maxColors);
	};
	CGIFQuantizer.prototype.buildPaletteMedianCut = function (palette, maxColors) {
		const bins = this.usedBins.slice(0, this.usedCount);
		const boxes = [{start: 0, end: this.usedCount, weight: this.getBoxWeight(bins, 0, this.usedCount)}];
		while (boxes.length < maxColors) {
			let boxIndex = -1;
			let bestWeight = -1;
			for (let index = 0; index < boxes.length; ++index) {
				const box = boxes[index];
				if (box.end - box.start < 2) {
					continue;
				}
				if (box.weight > bestWeight) {
					bestWeight = box.weight;
					boxIndex = index;
				}
			}
			if (-1 === boxIndex) {
				break;
			}
			const box = boxes[boxIndex];
			const axis = this.getBoxLongestAxis(bins, box);
			const shift = 2 === axis ? 0 : (1 === axis ? 5 : 10);
			bins.subarray(box.start, box.end).sort(function (first, second) {
				return ((first >> shift) & 0x1F) - ((second >> shift) & 0x1F);
			});
			const middle = this.getBoxMedian(bins, box);
			boxes.splice(boxIndex + 1, 0, {
				start: middle,
				end: box.end,
				weight: this.getBoxWeight(bins, middle, box.end)
			});
			box.end = middle;
			box.weight = this.getBoxWeight(bins, box.start, middle);
		}
		const color = [0, 0, 0];
		palette.count = boxes.length;
		for (let index = 0; index < boxes.length; ++index) {
			const box = boxes[index];
			let weight = 0;
			let totalR = 0;
			let totalG = 0;
			let totalB = 0;
			for (let position = box.start; position < box.end; ++position) {
				const bin = bins[position];
				const count = this.counts[bin];
				this.getBinColor(bin, color);
				totalR += color[0] * count;
				totalG += color[1] * count;
				totalB += color[2] * count;
				weight += count;
				palette.lut[bin] = index;
			}
			if (0 === weight) {
				weight = 1;
			}
			palette.colors[index * 3] = (totalR / weight) | 0;
			palette.colors[index * 3 + 1] = (totalG / weight) | 0;
			palette.colors[index * 3 + 2] = (totalB / weight) | 0;
		}
	};
	CGIFQuantizer.prototype.getBoxWeight = function (bins, start, end) {
		let weight = 0;
		for (let position = start; position < end; ++position) {
			weight += this.counts[bins[position]];
		}
		return weight;
	};
	CGIFQuantizer.prototype.getBoxLongestAxis = function (bins, box) {
		let minR = 31, maxR = 0, minG = 31, maxG = 0, minB = 31, maxB = 0;
		for (let position = box.start; position < box.end; ++position) {
			const bin = bins[position];
			const r = (bin >> 10) & 0x1F;
			const g = (bin >> 5) & 0x1F;
			const b = bin & 0x1F;
			if (r < minR) { minR = r; }
			if (r > maxR) { maxR = r; }
			if (g < minG) { minG = g; }
			if (g > maxG) { maxG = g; }
			if (b < minB) { minB = b; }
			if (b > maxB) { maxB = b; }
		}
		const rangeR = maxR - minR;
		const rangeG = maxG - minG;
		const rangeB = maxB - minB;
		if (rangeG >= rangeR && rangeG >= rangeB) {
			return 1;
		}
		return rangeR >= rangeB ? 0 : 2;
	};
	CGIFQuantizer.prototype.getBoxMedian = function (bins, box) {
		let total = 0;
		for (let position = box.start; position < box.end; ++position) {
			total += this.counts[bins[position]];
		}
		let half = 0;
		for (let position = box.start; position < box.end - 1; ++position) {
			half += this.counts[bins[position]];
			if (half * 2 >= total) {
				return position + 1;
			}
		}
		return box.end - 1;
	};

	function CGIFLzw(stream) {
		this.stream = stream;
		this.hashTable = new Int32Array(HASH_SIZE);
		this.codeTable = new Int32Array(HASH_SIZE);
		this.block = new Uint8Array(256);
		this.blockSize = 0;
		this.accumulator = 0;
		this.accumulatorBits = 0;
		this.codeBits = 0;
		this.initialBits = 0;
		this.maxCode = 0;
		this.freeEntry = 0;
		this.clearCode = 0;
		this.endCode = 0;
	}
	CGIFLzw.prototype.encode = function (indices, count, minCodeSize) {
		this.initialBits = minCodeSize + 1;
		this.codeBits = this.initialBits;
		this.clearCode = 1 << minCodeSize;
		this.endCode = this.clearCode + 1;
		this.freeEntry = this.clearCode + 2;
		this.maxCode = (1 << this.codeBits) - 1;
		this.blockSize = 0;
		this.accumulator = 0;
		this.accumulatorBits = 0;

		let hashShift = 0;
		let value = HASH_SIZE;
		while (value < 65536) {
			++hashShift;
			value *= 2;
		}
		hashShift = 8 - hashShift;

		this.clearHash();
		this.writeCode(this.clearCode);

		const hashTable = this.hashTable;
		const codeTable = this.codeTable;
		let prefix = indices[0];
		for (let position = 1; position < count; ++position) {
			const suffix = indices[position];
			const key = ((suffix << 12) | prefix) + 1;
			let slot = (suffix << hashShift) ^ prefix;
			if (slot >= HASH_SIZE) {
				slot -= HASH_SIZE;
			}
			let found = false;
			if (0 !== hashTable[slot]) {
				if (hashTable[slot] === key) {
					prefix = codeTable[slot];
					found = true;
				}
				else {
					let step = 0 === slot ? 1 : HASH_SIZE - slot;
					do {
						slot -= step;
						if (slot < 0) {
							slot += HASH_SIZE;
						}
						if (hashTable[slot] === key) {
							prefix = codeTable[slot];
							found = true;
							break;
						}
					} while (0 !== hashTable[slot]);
				}
			}
			if (found) {
				continue;
			}
			this.writeCode(prefix);
			prefix = suffix;
			if (this.freeEntry < LZW_MAX_CODES) {
				hashTable[slot] = key;
				codeTable[slot] = this.freeEntry;
				++this.freeEntry;
				if (this.freeEntry > this.maxCode + 1 && this.codeBits < 12) {
					++this.codeBits;
					this.maxCode = (1 << this.codeBits) - 1;
				}
			}
			else {
				this.writeCode(this.clearCode);
				this.clearHash();
				this.codeBits = this.initialBits;
				this.maxCode = (1 << this.codeBits) - 1;
				this.freeEntry = this.clearCode + 2;
			}
		}
		this.writeCode(prefix);
		this.writeCode(this.endCode);
		this.flushBits();
		this.flushBlock();
		this.stream.writeByte(0);
	};
	CGIFLzw.prototype.clearHash = function () {
		for (let index = 0; index < HASH_SIZE; ++index) {
			this.hashTable[index] = 0;
		}
	};
	CGIFLzw.prototype.writeCode = function (code) {
		this.accumulator |= code << this.accumulatorBits;
		this.accumulatorBits += this.codeBits;
		while (this.accumulatorBits >= 8) {
			this.writeBlockByte(this.accumulator & 0xFF);
			this.accumulator >>>= 8;
			this.accumulatorBits -= 8;
		}
	};
	CGIFLzw.prototype.flushBits = function () {
		if (this.accumulatorBits > 0) {
			this.writeBlockByte(this.accumulator & 0xFF);
			this.accumulator = 0;
			this.accumulatorBits = 0;
		}
	};
	CGIFLzw.prototype.writeBlockByte = function (value) {
		this.block[this.blockSize++] = value;
		if (255 === this.blockSize) {
			this.flushBlock();
		}
	};
	CGIFLzw.prototype.flushBlock = function () {
		if (0 === this.blockSize) {
			return;
		}
		this.stream.writeByte(this.blockSize);
		this.stream.writeBytes(this.block, this.blockSize);
		this.blockSize = 0;
	};

	function CGIFEncoder(width, height, loopCount) {
		this.width = width;
		this.height = height;
		this.loopCount = undefined === loopCount ? 0 : loopCount;

		this.stream = new CGIFStream();
		this.quantizer = new CGIFQuantizer();
		this.palette = new CGIFPalette();
		this.lzw = new CGIFLzw(this.stream);

		this.previous = null;
		this.indices = new Uint8Array(width * height);
		this.frameCount = 0;
		this.lastDelayPos = -1;
		this.finished = false;

		this.writeHeader();
	}
	CGIFEncoder.prototype.writeHeader = function () {
		this.stream.writeString("GIF89a");
		this.stream.writeShort(this.width);
		this.stream.writeShort(this.height);
		this.stream.writeByte(0x70);
		this.stream.writeByte(0);
		this.stream.writeByte(0);

		this.stream.writeByte(0x21);
		this.stream.writeByte(0xFF);
		this.stream.writeByte(0x0B);
		this.stream.writeString("NETSCAPE2.0");
		this.stream.writeByte(0x03);
		this.stream.writeByte(0x01);
		this.stream.writeShort(this.loopCount);
		this.stream.writeByte(0);
	};
	CGIFEncoder.prototype.getFrameCount = function () {
		return this.frameCount;
	};
	CGIFEncoder.prototype.addDelay = function (delayCs) {
		if (this.lastDelayPos < 0) {
			return false;
		}
		const delay = Math.min(MAX_DELAY_CS, this.stream.readShort(this.lastDelayPos) + delayCs);
		this.stream.patchShort(this.lastDelayPos, delay);
		return true;
	};
	CGIFEncoder.prototype.getChangedRect = function (rgba) {
		const previous = this.previous;
		let left = this.width;
		let top = this.height;
		let right = -1;
		let bottom = -1;
		for (let y = 0; y < this.height; ++y) {
			const rowStart = y * this.width;
			for (let x = 0; x < this.width; ++x) {
				const offset = (rowStart + x) * 4;
				if (rgba[offset] !== previous[offset] ||
					rgba[offset + 1] !== previous[offset + 1] ||
					rgba[offset + 2] !== previous[offset + 2]) {
					if (x < left) { left = x; }
					if (x > right) { right = x; }
					if (y < top) { top = y; }
					bottom = y;
				}
			}
		}
		if (-1 === right) {
			return null;
		}
		return {x: left, y: top, w: right - left + 1, h: bottom - top + 1};
	};
	CGIFEncoder.prototype.writeFrame = function (rgba, delayCs) {
		if (this.finished) {
			return false;
		}
		let rect = {x: 0, y: 0, w: this.width, h: this.height};
		let isDelta = false;
		if (this.previous) {
			rect = this.getChangedRect(rgba);
			if (!rect) {
				this.addDelay(delayCs);
				return false;
			}
			isDelta = true;
		}

		const transparentIndex = this.buildFrameIndices(rgba, rect, isDelta);
		this.writeFrameData(rect, delayCs, transparentIndex);

		if (!this.previous) {
			this.previous = new Uint8Array(rgba.length);
		}
		this.previous.set(rgba);
		++this.frameCount;
		return true;
	};
	CGIFEncoder.prototype.buildFrameIndices = function (rgba, rect, isDelta) {
		const quantizer = this.quantizer;
		const palette = this.palette;
		const previous = this.previous;
		quantizer.reset();

		for (let y = 0; y < rect.h; ++y) {
			const rowStart = (rect.y + y) * this.width + rect.x;
			for (let x = 0; x < rect.w; ++x) {
				const offset = (rowStart + x) * 4;
				if (isDelta &&
					rgba[offset] === previous[offset] &&
					rgba[offset + 1] === previous[offset + 1] &&
					rgba[offset + 2] === previous[offset + 2]) {
					continue;
				}
				quantizer.addPixel(rgba[offset], rgba[offset + 1], rgba[offset + 2]);
			}
		}

		const maxColors = isDelta ? MAX_COLORS - 1 : MAX_COLORS;
		quantizer.buildPalette(palette, maxColors);
		const transparentIndex = isDelta ? palette.count : -1;
		if (isDelta) {
			palette.colors[palette.count * 3] = 0;
			palette.colors[palette.count * 3 + 1] = 0;
			palette.colors[palette.count * 3 + 2] = 0;
			++palette.count;
		}

		const indices = this.indices;
		let position = 0;
		for (let y = 0; y < rect.h; ++y) {
			const rowStart = (rect.y + y) * this.width + rect.x;
			for (let x = 0; x < rect.w; ++x) {
				const offset = (rowStart + x) * 4;
				if (isDelta &&
					rgba[offset] === previous[offset] &&
					rgba[offset + 1] === previous[offset + 1] &&
					rgba[offset + 2] === previous[offset + 2]) {
					indices[position++] = transparentIndex;
					continue;
				}
				const bin = ((rgba[offset] >> 3) << 10) | ((rgba[offset + 1] >> 3) << 5) | (rgba[offset + 2] >> 3);
				indices[position++] = palette.lut[bin];
			}
		}
		return transparentIndex;
	};
	CGIFEncoder.prototype.writeFrameData = function (rect, delayCs, transparentIndex) {
		const stream = this.stream;
		const palette = this.palette;

		stream.writeByte(0x21);
		stream.writeByte(0xF9);
		stream.writeByte(0x04);
		stream.writeByte((DISPOSAL_LEAVE << 2) | (transparentIndex >= 0 ? 1 : 0));
		this.lastDelayPos = stream.pos;
		stream.writeShort(Math.min(MAX_DELAY_CS, delayCs));
		stream.writeByte(transparentIndex >= 0 ? transparentIndex : 0);
		stream.writeByte(0);

		const bits = palette.getBits();
		stream.writeByte(0x2C);
		stream.writeShort(rect.x);
		stream.writeShort(rect.y);
		stream.writeShort(rect.w);
		stream.writeShort(rect.h);
		stream.writeByte(0x80 | (bits - 1));

		const tableSize = palette.getTableSize();
		stream.writeBytes(palette.colors, palette.count * 3);
		for (let index = palette.count; index < tableSize; ++index) {
			stream.writeByte(0);
			stream.writeByte(0);
			stream.writeByte(0);
		}

		const minCodeSize = Math.max(2, bits);
		stream.writeByte(minCodeSize);
		this.lzw.encode(this.indices, rect.w * rect.h, minCodeSize);
	};
	CGIFEncoder.prototype.finish = function () {
		if (!this.finished) {
			this.stream.writeByte(0x3B);
			this.finished = true;
		}
		return this.stream.toUint8Array();
	};

	window['AscCommonSlide'].CGIFEncoder = CGIFEncoder;

})(window);
