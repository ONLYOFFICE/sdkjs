(function(){
	if (undefined === String.fromCodePoint)
	{
		String.fromCodePoint = function()
		{
			let codesUtf16 = [];
			var length = arguments.length;
			if (!length)
				return "";

			for (let index = 0; index < length; index++)
			{
				let code = arguments[index];
				if (code < 0x10000)
				{
					codesUtf16.push(code);
				}
				else
				{
					code -= 0x10000;
					codesUtf16.push(0xD800 | (code >> 10));
					codesUtf16.push(0xDC00 | (code & 0x3FF));
				}
			}

			return String.fromCharCode.apply(null, codesUtf16);
		};
	}

	if (undefined === String.prototype.codePointAt)
	{
		String.prototype.codePointAt = function(position)
		{
			if (position < 0 || position >= this.length)
				return undefined;

			let first = this.charCodeAt(position);

			if ((first >= 0xD800) && (first <= 0xDBFF) && (this.length > (position + 1)))
			{
				let second = this.charCodeAt(position + 1);
				if (second >= 0xDC00 && second <= 0xDFFF)
				{
					return ((first & 0x3FF) << 10) | (second & 0x3FF) + 0x10000;
				}
			}

			return first;
		};
	}

	if (undefined === String.prototype.replaceAll)
	{
		String.prototype.replaceAll = function(str, newStr)
		{
			if (Object.prototype.toString.call(str).toLowerCase() === '[object regexp]')
				return this.replace(str, newStr);
			return this.split(str).join(newStr);
		};
	}

	String.prototype.sentenceCase = function ()
	{
		var trimStr = this.trim();
		if (trimStr.length === 0)
		{
			return "";
		} else if (trimStr.length === 1)
		{
			return trimStr[0].toUpperCase();
		}
		return trimStr[0].toUpperCase() + trimStr.slice(1);
	}
	String.prototype['sentenceCase'] = String.prototype.sentenceCase;

	if (undefined === String.prototype.startsWith)
	{
		String.prototype.startsWith = function (str)
		{
			return this.indexOf(str) === 0;
		};
		String.prototype['startsWith'] = String.prototype.startsWith;
	}

	if (undefined ===  String.prototype.endsWith)
	{
		String.prototype.endsWith = function (suffix)
		{
			return this.indexOf(suffix, this.length - suffix.length) !== -1;
		};
		String.prototype['endsWith'] = String.prototype.endsWith;
	}

	if (undefined === String.prototype.trim)
	{
		String.prototype.trim = function()
		{
			var reg = new RegExp('^[\\s\uFEFF\xA0]+|[\\s\uFEFF\xA0]+$', 'g');
			return this.replace(reg, '');
		};
	}

	if (typeof String.prototype.repeat !== 'function')
	{
		String.prototype.repeat = function (count)
		{
			'use strict';
			if (this == null)
			{
				throw new TypeError('can\'t convert ' + this + ' to object');
			}
			var str = '' + this;
			count = +count;
			if (count != count)
			{
				count = 0;
			}
			if (count < 0)
			{
				throw new RangeError('repeat count must be non-negative');
			}
			if (count == Infinity)
			{
				throw new RangeError('repeat count must be less than infinity');
			}
			count = Math.floor(count);
			if (str.length == 0 || count == 0)
			{
				return '';
			}
			// Обеспечение того, что count является 31-битным целым числом, позволяет нам значительно
			// соптимизировать главную часть функции. Впрочем, большинство современных (на август
			// 2014 года) браузеров не обрабатывают строки, длиннее 1 << 28 символов, так что:
			if (str.length * count >= 1 << 28)
			{
				throw new RangeError('repeat count must not overflow maximum string size');
			}
			var rpt = '';
			for (; ;)
			{
				if ((count & 1) == 1)
				{
					rpt += str;
				}
				count >>>= 1;
				if (count == 0)
				{
					break;
				}
				str += str;
			}
			return rpt;
		};
		String.prototype['repeat'] = String.prototype.repeat;
	}

	if (typeof String.prototype.padStart !== 'function')
	{
		String.prototype.padStart = function padStart(targetLength,padString)
		{
			targetLength = targetLength>>0; //floor if number or convert non-number to 0;
			padString = String(padString || ' ');
			if (this.length > targetLength)
			{
				return String(this);
			}
			else
			{
				targetLength = targetLength-this.length;
				if (targetLength > padString.length)
				{
					padString += padString.repeat(targetLength/padString.length); //append to original to ensure we are longer than needed
				}
				return padString.slice(0,targetLength) + String(this);
			}
		};
		String.prototype['padStart'] = String.prototype.padStart;
	}

	String.prototype.strongMatch = function (regExp)
	{
		if (regExp && regExp instanceof RegExp)
		{
			var arr = this.toString().match(regExp);
			return !!(arr && arr.length > 0 && arr[0].length == this.length);
		}

		return false;
	};

	if (undefined === String.prototype.includes)
	{
		String.prototype.includes = function(search, start)
		{
			if (typeof start !== 'number')
				start = 0;

			if (start + search.length > this.length)
				return false;
			return this.indexOf(search, start) !== -1;
		};
	}

	function CUnicodeIterator(str)
	{
		this._position = 0;
		this._index = 0;
		this._str = str;
	}

	CUnicodeIterator.prototype.isOutside = function()
	{
		return (this._index >= this._str.length);
	};
	CUnicodeIterator.prototype.isInside = function()
	{
		return (this._index < this._str.length);
	};
	CUnicodeIterator.prototype.value = function()
	{
		if (this._index >= this._str.length)
			return 0;

		var nCharCode = this._str.charCodeAt(this._index);
		if (!AscCommon.isLeadingSurrogateChar(nCharCode))
			return nCharCode;

		if ((this._str.length - 1) === this._index)
			return nCharCode; // error

		var nTrailingChar = this._str.charCodeAt(this._index + 1);
		return AscCommon.decodeSurrogateChar(nCharCode, nTrailingChar);
	};
	CUnicodeIterator.prototype.next = function()
	{
		if (this._index >= this._str.length)
			return;

		this._position++;
		if (!AscCommon.isLeadingSurrogateChar(this._str.charCodeAt(this._index)))
		{
			++this._index;
			return;
		}

		if (this._index === (this._str.length - 1))
		{
			++this._index;
			return;
		}

		this._index += 2;
	};
	CUnicodeIterator.prototype.position = function()
	{
		return this._position;
	};
	CUnicodeIterator.prototype.check = CUnicodeIterator.prototype.isInside;

	/**
	 * @returns {CUnicodeIterator}
	 */
	String.prototype.getUnicodeIterator = function()
	{
		return new CUnicodeIterator(this);
	};
	/**
	 * @returns {number[]}
	 */
	String.prototype.codePointsArray = function(codePoints)
	{
		let _codePoints = codePoints ? codePoints : [];

		for (let iter = this.getUnicodeIterator(); iter.check(); iter.next())
			_codePoints.push(iter.value());

		return _codePoints;
	};

})();
