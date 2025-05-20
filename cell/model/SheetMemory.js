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

(function(window, undefined) {
	/**
	 * Class for managing sheet memory with efficient storage
	 * @constructor
	 * @param {number} structSize - Size of the structure in bytes (must be a multiple of 8)
	 * @param {number} maxIndex - Maximum index allowed
	 */
	function SheetMemory(structSize, maxIndex) {
		//todo separate structure for data and style
		this.dataBuffer = null; // ArrayBuffer to store the data
		this.dataUint8 = null; // Uint8Array view on the buffer
		this.dataFloat = null; // Float64Array view on the buffer
		this.dataInt32 = null; // Int32Array view on the buffer
		this.indexA = -1;
		this.indexB = -1;
		
		// Ensure structSize is a multiple of 8 for proper Float64Array alignment
		if (structSize % 8 !== 0) {
			this.structSize = ((structSize + 7) >> 3) << 3; // Round up to the nearest multiple of 8
			console.warn("SheetMemory: structSize must be a multiple of 8. Adjusted from " + structSize + " to " + this.structSize);
		} else {
			this.structSize = structSize;
		}
		
		this.maxIndex = maxIndex;
	}

	/**
	 * Ensures the given size is properly aligned for typed array views
	 * @private
	 * @param {number} size - Size to align
	 * @returns {number} Aligned size
	 */
	SheetMemory.prototype._alignSize = function(size) {
		// Align to 8 bytes (Float64Array requirement)
		return (size + 7) & ~7; // Equivalent to Math.ceil(size / 8) * 8
	};

	/**
	 * Creates typed array views for the given buffer
	 * @private
	 * @param {ArrayBuffer} buffer - The ArrayBuffer to create views for
	 */
	SheetMemory.prototype._createViews = function(buffer) {
		this.dataUint8 = new Uint8Array(buffer);
		this.dataFloat = new Float64Array(buffer);
		this.dataInt32 = new Int32Array(buffer);
	};

	/**
	 * Checks and ensures that the specified index is valid and the data array is properly sized
	 * @param {number} index - Index to check
	 */
	SheetMemory.prototype.checkIndex = function(index) {
		if (index > this.maxIndex) {
			index = this.maxIndex;
		}
		if (this.dataBuffer) {
			let allocatedCount = this.getAllocatedCount();
			if (index > this.indexB) {
				if (this.indexA + allocatedCount - 1 < index) {
					// Calculate needed array size, with growth factor for better amortized performance
					let newAllocatedCount = Math.min(Math.max((1.5 * allocatedCount) | 0, index - this.indexA + 1), (this.maxIndex - this.indexA + 1));
					if (newAllocatedCount > allocatedCount) {
						let oldData = this.dataUint8;
						// Ensure buffer size is aligned for typed arrays
						let bufferSize = this._alignSize(newAllocatedCount * this.structSize);
						this.dataBuffer = new ArrayBuffer(bufferSize);
						this._createViews(this.dataBuffer);
						// Fast copy of old data
						this.dataUint8.set(oldData);
					}
				}
				this.indexB = index;
			} else if (index < this.indexA) {
				let oldData = this.dataUint8;
				let oldIndexA = this.indexA;
				this.indexA = Math.max(0, index);
				let diff = oldIndexA - this.indexA;
				// Ensure buffer size is aligned for typed arrays
				let bufferSize = this._alignSize((allocatedCount + diff) * this.structSize);
				this.dataBuffer = new ArrayBuffer(bufferSize);
				this._createViews(this.dataBuffer);
				// When copying with offset, must use this pattern
				this.dataUint8.set(oldData, diff * this.structSize);
			}
		} else {
			this.indexA = this.indexB = index;
			// Initial allocation - optimize for common case scenarios
			let newAllocatedCount = Math.min(32, (this.maxIndex - this.indexA + 1));
			// Ensure buffer size is aligned for typed arrays
			let bufferSize = this._alignSize(newAllocatedCount * this.structSize);
			this.dataBuffer = new ArrayBuffer(bufferSize);
			this._createViews(this.dataBuffer);
		}
	};

	/**
	 * Checks if the specified index is within the allocated range
	 * @param {number} index - Index to check
	 * @returns {boolean} True if index is within the allocated range
	 */
	SheetMemory.prototype.hasIndex = function(index) {
		return this.indexA <= index && index <= this.indexB;
	};

	/**
	 * Gets the minimum index in the memory
	 * @returns {number} Minimum index
	 */
	SheetMemory.prototype.getMinIndex = function() {
		return this.indexA;
	};

	/**
	 * Gets the maximum index in the memory
	 * @returns {number} Maximum index
	 */
	SheetMemory.prototype.getMaxIndex = function() {
		return this.indexB;
	};

	/**
	 * Gets the allocated count
	 * @returns {number} Allocated count
	 */
	SheetMemory.prototype.getAllocatedCount = function() {
		if (!this.dataBuffer) {
			return 0;
		}
		// Calculate the actual element count from the buffer size, considering alignment
		return Math.floor(this.dataBuffer.byteLength / this.structSize);
	};

	/**
	 * Creates a clone of this SheetMemory
	 * @returns {SheetMemory} Cloned SheetMemory object
	 */
	SheetMemory.prototype.clone = function() {
		const sheetMemory = new SheetMemory(this.structSize, this.maxIndex);
		if (this.dataBuffer) {
			// Use the exact same buffer size for efficiency
			sheetMemory.dataBuffer = new ArrayBuffer(this.dataBuffer.byteLength);
			// Use helper method to create views consistently
			sheetMemory._createViews(sheetMemory.dataBuffer);
			// Copy data once
			sheetMemory.dataUint8.set(this.dataUint8);
		}
		sheetMemory.indexA = this.indexA;
		sheetMemory.indexB = this.indexB;
		return sheetMemory;
	};

	/**
	 * Deletes a range of indices
	 * @param {number} start - Start index
	 * @param {number} deleteCount - Number of indices to delete
	 */
	SheetMemory.prototype.deleteRange = function(start, deleteCount) {
		let delA = start;
		let delB = start + deleteCount - 1;
		if (delA > this.indexB) {
			return;
		}
		if (delA <= this.indexA) {
			if (delB < this.indexA) {
				this.indexA -= deleteCount;
				this.indexB -= deleteCount;
			} else if (delB >= this.indexB) {
				// Reset buffer to initial state instead of nullifying dataUint8
				this.dataBuffer = null;
				this.dataUint8 = null;
				this.dataFloat = null;
				this.dataInt32 = null;
				this.indexA = this.indexB = -1;
			} else {
				let endOffset = (delB + 1 - this.indexA) * this.structSize;
				this.dataUint8.set(this.dataUint8.subarray(endOffset), 0);
				this.dataUint8.fill(0, (this.indexB - delB) * this.structSize);
				this.indexA = delA;
				this.indexB -= deleteCount;
			}
		} else {
			if (delB >= this.indexB) {
				this.dataUint8.fill(0, (delA - this.indexA) * this.structSize);
				this.indexB = delA - 1;
			} else {
				let startOffset = (delA - this.indexA) * this.structSize;
				let endOffset = (delB + 1 - this.indexA) * this.structSize;
				this.dataUint8.set(this.dataUint8.subarray(endOffset), startOffset);
				this.dataUint8.fill(0, (this.indexB - this.indexA + 1 - deleteCount) * this.structSize);
				this.indexB -= deleteCount;
			}
		}
	};

	/**
	 * Inserts a range of indices
	 * @param {number} start - Start index
	 * @param {number} insertCount - Number of indices to insert
	 */
	SheetMemory.prototype.insertRange = function(start, insertCount) {
		let insA = start;
		let insB = start + insertCount;
		if (insA > this.indexB) {
			return;
		}
		if (insA <= this.indexA) {
			this.indexA += insertCount;
			this.indexB += insertCount;
		} else {
			this.checkIndex(this.indexB + insertCount);
			const newCount = (this.indexB + 1 - this.indexA);
			const startOffset = (insA - this.indexA) * this.structSize;
			const endOffset = (insB - this.indexA) * this.structSize;
			const endData = (newCount - insertCount) * this.structSize;
			this.dataUint8.set(this.dataUint8.subarray(startOffset, endData), endOffset);
			this.dataUint8.fill(0, startOffset, endOffset);
		}
	};

	/**
	 * Copies a range from another SheetMemory
	 * @param {SheetMemory} sheetMemory - Source SheetMemory
	 * @param {number} startFrom - Start index in source
	 * @param {number} startTo - Start index in destination
	 * @param {number} count - Number of indices to copy
	 */
	SheetMemory.prototype.copyRange = function(sheetMemory, startFrom, startTo, count) {
		let dataCopy, startToSrc = startTo, countSrc = count;
		if (startFrom <= sheetMemory.indexB && startFrom + count - 1 >= sheetMemory.indexA) {
			if (startFrom < sheetMemory.indexA) {
				let diff = sheetMemory.indexA - startFrom;
				startTo += diff;
				count -= diff;
				startFrom = sheetMemory.indexA;
			}
			if (startFrom + count - 1 > sheetMemory.indexB) {
				count -= startFrom + count - 1 - sheetMemory.indexB;
			}
			if (count > 0) {
				let startOffsetFrom = (startFrom - sheetMemory.indexA) * this.structSize;
				let endOffsetFrom = (startFrom - sheetMemory.indexA + count) * this.structSize;
				dataCopy = sheetMemory.dataUint8.slice(startOffsetFrom, endOffsetFrom);
			}
		}
		this.clear(startToSrc, startToSrc + countSrc);
		if(dataCopy) {
			this.checkIndex(startTo);
			this.checkIndex(startTo + count - 1);
			let startOffsetTo = (startTo - this.indexA) * this.structSize;
			this.dataUint8.set(dataCopy, startOffsetTo);
		}
	};

	/**
	 * Sets an area by row
	 * @param {number} from - Source row index
	 * @param {number} to - Destination start row index
	 * @param {number} toCount - Number of destination rows
	 */
	SheetMemory.prototype.setAreaByRow = function(from, to, toCount) {
		let fromCount = 1;
		if (from <= this.indexB && from + fromCount - 1 >= this.indexA) {
			//todo from < this.indexA
			let fromStartOffset = Math.max(0, (from - this.indexA)) * this.structSize;
			let fromEndOffset = (Math.min((from + fromCount), this.indexB + 1) - this.indexA) * this.structSize;
			let fromSubArray = this.dataUint8.subarray(fromStartOffset, fromEndOffset);
			this.checkIndex(to + toCount - 1);
			for (let i = to; i < to + toCount && i <= this.indexB; i += fromCount) {
				this.dataUint8.set(fromSubArray, (i - this.indexA) * this.structSize);
			}
		}
	};

	/**
	 * Sets an area by row
	 * @param {number} from - Source row index
	 * @param {number} fromCount - Number of source rows
	 * @param {number} to - Destination start row index
	 * @param {number} toCount - Number of destination rows
	 */
	SheetMemory.prototype.copyRangeByChunk = function(from, fromCount, to, toCount) {
		this.setAreaByRow(from, to, toCount);
	}

	/**
	 * Clears a range of indices
	 * @param {number} start - Start index
	 * @param {number} end - End index (exclusive)
	 */
	SheetMemory.prototype.clear = function(start, end) {
		start = Math.max(start, this.indexA);
		end = Math.min(end, this.indexB + 1);
		if (start < end) {
			this.dataUint8.fill(0, (start - this.indexA) * this.structSize, (end - this.indexA) * this.structSize);
		}
	};

	/**
	 * Gets an unsigned 8-bit integer
	 * @param {number} index - Index
	 * @param {number} offset - Offset within the structure
	 * @returns {number} Unsigned 8-bit integer
	 */
	SheetMemory.prototype.getUint8 = function(index, offset) {
		offset += (index - this.indexA) * this.structSize;
		return this.dataUint8[offset];
	};

	/**
	 * Sets an unsigned 8-bit integer
	 * @param {number} index - Index
	 * @param {number} offset - Offset within the structure
	 * @param {number} val - Value to set
	 */
	SheetMemory.prototype.setUint8 = function(index, offset, val) {
		offset += (index - this.indexA) * this.structSize;
		this.dataUint8[offset] = val;
	};

	/**
	 * Gets a signed 32-bit integer
	 * @param {number} index - Index
	 * @param {number} offset - Offset within the structure
	 * @returns {number} Signed 32-bit integer
	 */
	SheetMemory.prototype.getInt32 = function(index, offset) {
		// Calculate byte offset and convert to Int32 index
		// We need to ensure the offset is aligned to 4-byte boundary
		const byteOffset = (index - this.indexA) * this.structSize + offset;
		return this.dataInt32[byteOffset >> 2]; // Division by 4 to get Int32 index
	};
	

	/**
	 * Sets a signed 32-bit integer
	 * @param {number} index - Index
	 * @param {number} offset - Offset within the structure
	 * @param {number} val - Value to set
	 */
	SheetMemory.prototype.setInt32 = function(index, offset, val) {
		// Calculate byte offset and convert to Int32 index
		// We need to ensure the offset is aligned to 4-byte boundary
		const byteOffset = (index - this.indexA) * this.structSize + offset;
		this.dataInt32[byteOffset >> 2] = val; // Division by 4 to get Int32 index
	};

	/**
	 * Gets a 64-bit floating point number
	 * @param {number} index - Index
	 * @param {number} offset - Offset within the structure
	 * @returns {number} 64-bit floating point number
	 */
	SheetMemory.prototype.getFloat64 = function(index, offset) {
		// Calculate byte offset and convert to Float64 index
		// We need to ensure the offset is aligned to 8-byte boundary
		const byteOffset = (index - this.indexA) * this.structSize + offset;
		return this.dataFloat[byteOffset >> 3]; // Division by 8 to get Float64 index
	};

	/**
	 * Sets a 64-bit floating point number
	 * @param {number} index - Index
	 * @param {number} offset - Offset within the structure
	 * @param {number} val - Value to set
	 */
	SheetMemory.prototype.setFloat64 = function(index, offset, val) {
		// Calculate byte offset and convert to Float64 index
		// We need to ensure the offset is aligned to 8-byte boundary
		const byteOffset = (index - this.indexA) * this.structSize + offset;
		this.dataFloat[byteOffset >> 3] = val; // Division by 8 to get Float64 index
	};

	// Export
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].SheetMemory = SheetMemory;
})(window);
