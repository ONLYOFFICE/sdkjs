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
})();
