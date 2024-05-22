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

(function(window, undefined) {

	function getMemoryPathIE(name)
	{
		if (self["AscViewer"] && self["AscViewer"]["baseUrl"])
			return self["AscViewer"]["baseUrl"] + name;
		return name;
	}

	var baseFontsPath = "../../../../fonts/";

	var FS = undefined;

	// correct fetch for desktop application

var printErr = undefined;
var print    = undefined;

var fetch = ("undefined" !== typeof window) ? window.fetch : (("undefined" !== typeof self) ? self.fetch : null);
var getBinaryPromise = null;

function internal_isLocal()
{
	if (window.navigator && window.navigator.userAgent.toLowerCase().indexOf("ascdesktopeditor") < 0)
		return false;
	if (window.location && window.location.protocol == "file:")
		return true;
	if (window.document && window.document.currentScript && 0 == window.document.currentScript.src.indexOf("file:///"))
		return true;
	return false;
}

if (internal_isLocal())
{
	fetch = undefined; // fetch not support file:/// scheme
	getBinaryPromise = function()
	{
		var wasmPath = "ascdesktop://fonts/" + wasmBinaryFile.substr(8);
		return new Promise(function (resolve, reject)
		{
			var xhr = new XMLHttpRequest();
			xhr.open('GET', wasmPath, true);
			xhr.responseType = 'arraybuffer';

			if (xhr.overrideMimeType)
				xhr.overrideMimeType('text/plain; charset=x-user-defined');
			else
				xhr.setRequestHeader('Accept-Charset', 'x-user-defined');

			xhr.onload = function ()
			{
				if (this.status == 200)
					resolve(new Uint8Array(this.response));
			};
			xhr.send(null);
		});
	}
}
else
{
	getBinaryPromise = function() { return getBinaryPromise2.apply(undefined, arguments); }
}


	//polyfill

	(function(){

	if (undefined !== String.prototype.fromUtf8 &&
		undefined !== String.prototype.toUtf8)
		return;

	var STRING_UTF8_BUFFER_LENGTH = 1024;
	var STRING_UTF8_BUFFER = new ArrayBuffer(STRING_UTF8_BUFFER_LENGTH);

	/**
	 * Read string from utf8
	 * @param {Uint8Array} buffer
	 * @param {number} [start=0]
	 * @param {number} [len]
	 * @returns {string}
	 */
	String.prototype.fromUtf8 = function(buffer, start, len) {
		if (undefined === start)
			start = 0;
		if (undefined === len)
			len = buffer.length - start;

		var result = "";
		var index  = start;
		var end = start + len;
		while (index < end)
		{
			var u0 = buffer[index++];
			if (!(u0 & 128))
			{
				result += String.fromCharCode(u0);
				continue;
			}
			var u1 = buffer[index++] & 63;
			if ((u0 & 224) == 192)
			{
				result += String.fromCharCode((u0 & 31) << 6 | u1);
				continue;
			}
			var u2 = buffer[index++] & 63;
			if ((u0 & 240) == 224)
				u0 = (u0 & 15) << 12 | u1 << 6 | u2;
			else
				u0 = (u0 & 7) << 18 | u1 << 12 | u2 << 6 | buffer[index++] & 63;
			if (u0 < 65536)
				result += String.fromCharCode(u0);
			else
			{
				var ch = u0 - 65536;
				result += String.fromCharCode(55296 | ch >> 10, 56320 | ch & 1023);
			}
		}
		return result;
	};

	/**
	 * Convert string to utf8 array
	 * @returns {Uint8Array}
	 */
	String.prototype.toUtf8 = function(isNoEndNull, isUseBuffer) {
		var inputLen = this.length;
		var testLen  = 6 * inputLen + 1;
		var tmpStrings = (isUseBuffer && testLen < STRING_UTF8_BUFFER_LENGTH) ? STRING_UTF8_BUFFER : new ArrayBuffer(testLen);

		var code  = 0;
		var index = 0;

		var outputIndex = 0;
		var outputDataTmp = new Uint8Array(tmpStrings);
		var outputData = outputDataTmp;

		while (index < inputLen)
		{
			code = this.charCodeAt(index++);
			if (code >= 0xD800 && code <= 0xDFFF && index < inputLen)
				code = 0x10000 + (((code & 0x3FF) << 10) | (0x03FF & this.charCodeAt(index++)));

			if (code < 0x80)
				outputData[outputIndex++] = code;
			else if (code < 0x0800)
			{
				outputData[outputIndex++] = 0xC0 | (code >> 6);
				outputData[outputIndex++] = 0x80 | (code & 0x3F);
			}
			else if (code < 0x10000)
			{
				outputData[outputIndex++] = 0xE0 | (code >> 12);
				outputData[outputIndex++] = 0x80 | ((code >> 6) & 0x3F);
				outputData[outputIndex++] = 0x80 | (code & 0x3F);
			}
			else if (code < 0x1FFFFF)
			{
				outputData[outputIndex++] = 0xF0 | (code >> 18);
				outputData[outputIndex++] = 0x80 | ((code >> 12) & 0x3F);
				outputData[outputIndex++] = 0x80 | ((code >> 6) & 0x3F);
				outputData[outputIndex++] = 0x80 | (code & 0x3F);
			}
			else if (code < 0x3FFFFFF)
			{
				outputData[outputIndex++] = 0xF8 | (code >> 24);
				outputData[outputIndex++] = 0x80 | ((code >> 18) & 0x3F);
				outputData[outputIndex++] = 0x80 | ((code >> 12) & 0x3F);
				outputData[outputIndex++] = 0x80 | ((code >> 6) & 0x3F);
				outputData[outputIndex++] = 0x80 | (code & 0x3F);
			}
			else if (code < 0x7FFFFFFF)
			{
				outputData[outputIndex++] = 0xFC | (code >> 30);
				outputData[outputIndex++] = 0x80 | ((code >> 24) & 0x3F);
				outputData[outputIndex++] = 0x80 | ((code >> 18) & 0x3F);
				outputData[outputIndex++] = 0x80 | ((code >> 12) & 0x3F);
				outputData[outputIndex++] = 0x80 | ((code >> 6) & 0x3F);
				outputData[outputIndex++] = 0x80 | (code & 0x3F);
			}
		}

		if (isNoEndNull !== true)
			outputData[outputIndex++] = 0;

		return new Uint8Array(tmpStrings, 0, outputIndex);
	};

	function StringPointer(pointer, len)
	{
		this.ptr = pointer;
		this.length = len;
	}
	StringPointer.prototype.free = function()
	{
		if (0 !== this.ptr)
			Module["_free"](this.ptr);
	};

	String.prototype.toUtf8Pointer = function(isNoEndNull) {
		var tmp = this.toUtf8(isNoEndNull, true);
		var pointer = Module["_malloc"](tmp.length);
		if (0 == pointer)
			return null;

		Module["HEAP8"].set(tmp, pointer);
		return new StringPointer(pointer, tmp.length);		
	};

})();


	var Module=typeof Module!="undefined"?Module:{};var moduleOverrides=Object.assign({},Module);var arguments_=[];var thisProgram="./this.program";var quit_=(status,toThrow)=>{throw toThrow};var ENVIRONMENT_IS_WEB=true;var ENVIRONMENT_IS_WORKER=false;var scriptDirectory="";function locateFile(path){if(Module["locateFile"]){return Module["locateFile"](path,scriptDirectory)}return scriptDirectory+path}var read_,readAsync,readBinary;if(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER){if(ENVIRONMENT_IS_WORKER){scriptDirectory=self.location.href}else if(typeof document!="undefined"&&document.currentScript){scriptDirectory=document.currentScript.src}if(scriptDirectory.indexOf("blob:")!==0){scriptDirectory=scriptDirectory.substr(0,scriptDirectory.replace(/[?#].*/,"").lastIndexOf("/")+1)}else{scriptDirectory=""}{read_=url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.send(null);return xhr.responseText};if(ENVIRONMENT_IS_WORKER){readBinary=url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.responseType="arraybuffer";xhr.send(null);return new Uint8Array(xhr.response)}}readAsync=(url,onload,onerror)=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,true);xhr.responseType="arraybuffer";xhr.onload=()=>{if(xhr.status==200||xhr.status==0&&xhr.response){onload(xhr.response);return}onerror()};xhr.onerror=onerror;xhr.send(null)}}}else{}var out=Module["print"]||console.log.bind(console);var err=Module["printErr"]||console.error.bind(console);Object.assign(Module,moduleOverrides);moduleOverrides=null;if(Module["arguments"])arguments_=Module["arguments"];if(Module["thisProgram"])thisProgram=Module["thisProgram"];if(Module["quit"])quit_=Module["quit"];var wasmBinary;if(Module["wasmBinary"])wasmBinary=Module["wasmBinary"];if(typeof WebAssembly!="object"){abort("no native wasm support detected")}var wasmMemory;var ABORT=false;var EXITSTATUS;var HEAP8,HEAPU8,HEAP16,HEAPU16,HEAP32,HEAPU32,HEAPF32,HEAPF64;function updateMemoryViews(){var b=wasmMemory.buffer;Module["HEAP8"]=HEAP8=new Int8Array(b);Module["HEAP16"]=HEAP16=new Int16Array(b);Module["HEAPU8"]=HEAPU8=new Uint8Array(b);Module["HEAPU16"]=HEAPU16=new Uint16Array(b);Module["HEAP32"]=HEAP32=new Int32Array(b);Module["HEAPU32"]=HEAPU32=new Uint32Array(b);Module["HEAPF32"]=HEAPF32=new Float32Array(b);Module["HEAPF64"]=HEAPF64=new Float64Array(b)}var __ATPRERUN__=[];var __ATINIT__=[];var __ATPOSTRUN__=[function(){window["AscViewer"] && window["AscViewer"]["onLoadModule"] && window["AscViewer"]["onLoadModule"]();}];var runtimeInitialized=false;function preRun(){if(Module["preRun"]){if(typeof Module["preRun"]=="function")Module["preRun"]=[Module["preRun"]];while(Module["preRun"].length){addOnPreRun(Module["preRun"].shift())}}callRuntimeCallbacks(__ATPRERUN__)}function initRuntime(){runtimeInitialized=true;callRuntimeCallbacks(__ATINIT__)}function postRun(){if(Module["postRun"]){if(typeof Module["postRun"]=="function")Module["postRun"]=[Module["postRun"]];while(Module["postRun"].length){addOnPostRun(Module["postRun"].shift())}}callRuntimeCallbacks(__ATPOSTRUN__)}function addOnPreRun(cb){__ATPRERUN__.unshift(cb)}function addOnInit(cb){__ATINIT__.unshift(cb)}function addOnPostRun(cb){__ATPOSTRUN__.unshift(cb)}var runDependencies=0;var runDependencyWatcher=null;var dependenciesFulfilled=null;function addRunDependency(id){runDependencies++;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}}function removeRunDependency(id){runDependencies--;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}if(runDependencies==0){if(runDependencyWatcher!==null){clearInterval(runDependencyWatcher);runDependencyWatcher=null}if(dependenciesFulfilled){var callback=dependenciesFulfilled;dependenciesFulfilled=null;callback()}}}function abort(what){if(Module["onAbort"]){Module["onAbort"](what)}what="Aborted("+what+")";err(what);ABORT=true;EXITSTATUS=1;what+=". Build with -sASSERTIONS for more info.";var e=new WebAssembly.RuntimeError(what);throw e}var dataURIPrefix="data:application/octet-stream;base64,";var isDataURI=filename=>filename.startsWith(dataURIPrefix);var wasmBinaryFile;wasmBinaryFile="drawingfile.wasm";if(!isDataURI(wasmBinaryFile)){wasmBinaryFile=locateFile(wasmBinaryFile)}function getBinarySync(file){if(file==wasmBinaryFile&&wasmBinary){return new Uint8Array(wasmBinary)}if(readBinary){return readBinary(file)}throw"both async and sync fetching of the wasm failed"}function getBinaryPromise2(binaryFile){if(!wasmBinary&&(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER)){if(typeof fetch=="function"){return fetch(binaryFile,{credentials:"same-origin"}).then(response=>{if(!response["ok"]){throw"failed to load wasm binary file at '"+binaryFile+"'"}return response["arrayBuffer"]()}).catch(()=>getBinarySync(binaryFile))}}return Promise.resolve().then(()=>getBinarySync(binaryFile))}function instantiateArrayBuffer(binaryFile,imports,receiver){return getBinaryPromise(binaryFile).then(binary=>WebAssembly.instantiate(binary,imports)).then(instance=>instance).then(receiver,reason=>{err(`failed to asynchronously prepare wasm: ${reason}`);abort(reason)})}function instantiateAsync(binary,binaryFile,imports,callback){if(!binary&&typeof WebAssembly.instantiateStreaming=="function"&&!isDataURI(binaryFile)&&typeof fetch=="function"){return fetch(binaryFile,{credentials:"same-origin"}).then(response=>{var result=WebAssembly.instantiateStreaming(response,imports);return result.then(callback,function(reason){err(`wasm streaming compile failed: ${reason}`);err("falling back to ArrayBuffer instantiation");return instantiateArrayBuffer(binaryFile,imports,callback)})})}return instantiateArrayBuffer(binaryFile,imports,callback)}function createWasm(){var info={"a":wasmImports};function receiveInstance(instance,module){wasmExports=instance.exports;wasmMemory=wasmExports["eb"];updateMemoryViews();wasmTable=wasmExports["gb"];addOnInit(wasmExports["fb"]);removeRunDependency("wasm-instantiate");return wasmExports}addRunDependency("wasm-instantiate");function receiveInstantiationResult(result){receiveInstance(result["instance"])}if(Module["instantiateWasm"]){try{return Module["instantiateWasm"](info,receiveInstance)}catch(e){err(`Module.instantiateWasm callback failed with error: ${e}`);return false}}instantiateAsync(wasmBinary,wasmBinaryFile,info,receiveInstantiationResult);return{}}var tempDouble;function js_get_stream_id(data,status){return self.AscViewer.CheckStreamId(data,status)}function js_free_id(data){self.AscViewer.Free(data);return 1}function ExitStatus(status){this.name="ExitStatus";this.message=`Program terminated with exit(${status})`;this.status=status}var callRuntimeCallbacks=callbacks=>{while(callbacks.length>0){callbacks.shift()(Module)}};var noExitRuntime=Module["noExitRuntime"]||true;var UTF8Decoder=typeof TextDecoder!="undefined"?new TextDecoder("utf8"):undefined;var UTF8ArrayToString=(heapOrArray,idx,maxBytesToRead)=>{var endIdx=idx+maxBytesToRead;var endPtr=idx;while(heapOrArray[endPtr]&&!(endPtr>=endIdx))++endPtr;if(endPtr-idx>16&&heapOrArray.buffer&&UTF8Decoder){return UTF8Decoder.decode(heapOrArray.subarray(idx,endPtr))}var str="";while(idx<endPtr){var u0=heapOrArray[idx++];if(!(u0&128)){str+=String.fromCharCode(u0);continue}var u1=heapOrArray[idx++]&63;if((u0&224)==192){str+=String.fromCharCode((u0&31)<<6|u1);continue}var u2=heapOrArray[idx++]&63;if((u0&240)==224){u0=(u0&15)<<12|u1<<6|u2}else{u0=(u0&7)<<18|u1<<12|u2<<6|heapOrArray[idx++]&63}if(u0<65536){str+=String.fromCharCode(u0)}else{var ch=u0-65536;str+=String.fromCharCode(55296|ch>>10,56320|ch&1023)}}return str};var UTF8ToString=(ptr,maxBytesToRead)=>ptr?UTF8ArrayToString(HEAPU8,ptr,maxBytesToRead):"";var ___assert_fail=(condition,filename,line,func)=>{abort(`Assertion failed: ${UTF8ToString(condition)}, at: `+[filename?UTF8ToString(filename):"unknown filename",line,func?UTF8ToString(func):"unknown function"])};var exceptionCaught=[];var uncaughtExceptionCount=0;var ___cxa_begin_catch=ptr=>{var info=new ExceptionInfo(ptr);if(!info.get_caught()){info.set_caught(true);uncaughtExceptionCount--}info.set_rethrown(false);exceptionCaught.push(info);___cxa_increment_exception_refcount(info.excPtr);return info.get_exception_ptr()};var exceptionLast=0;var ___cxa_end_catch=()=>{_setThrew(0,0);var info=exceptionCaught.pop();___cxa_decrement_exception_refcount(info.excPtr);exceptionLast=0};function ExceptionInfo(excPtr){this.excPtr=excPtr;this.ptr=excPtr-24;this.set_type=function(type){HEAPU32[this.ptr+4>>2]=type};this.get_type=function(){return HEAPU32[this.ptr+4>>2]};this.set_destructor=function(destructor){HEAPU32[this.ptr+8>>2]=destructor};this.get_destructor=function(){return HEAPU32[this.ptr+8>>2]};this.set_caught=function(caught){caught=caught?1:0;HEAP8[this.ptr+12>>0]=caught};this.get_caught=function(){return HEAP8[this.ptr+12>>0]!=0};this.set_rethrown=function(rethrown){rethrown=rethrown?1:0;HEAP8[this.ptr+13>>0]=rethrown};this.get_rethrown=function(){return HEAP8[this.ptr+13>>0]!=0};this.init=function(type,destructor){this.set_adjusted_ptr(0);this.set_type(type);this.set_destructor(destructor)};this.set_adjusted_ptr=function(adjustedPtr){HEAPU32[this.ptr+16>>2]=adjustedPtr};this.get_adjusted_ptr=function(){return HEAPU32[this.ptr+16>>2]};this.get_exception_ptr=function(){var isPointer=___cxa_is_pointer_type(this.get_type());if(isPointer){return HEAPU32[this.excPtr>>2]}var adjusted=this.get_adjusted_ptr();if(adjusted!==0)return adjusted;return this.excPtr}}var ___resumeException=ptr=>{if(!exceptionLast){exceptionLast=ptr}throw exceptionLast};var findMatchingCatch=args=>{var thrown=exceptionLast;if(!thrown){setTempRet0(0);return 0}var info=new ExceptionInfo(thrown);info.set_adjusted_ptr(thrown);var thrownType=info.get_type();if(!thrownType){setTempRet0(0);return thrown}for(var arg in args){var caughtType=args[arg];if(caughtType===0||caughtType===thrownType){break}var adjusted_ptr_addr=info.ptr+16;if(___cxa_can_catch(caughtType,thrownType,adjusted_ptr_addr)){setTempRet0(caughtType);return thrown}}setTempRet0(thrownType);return thrown};var ___cxa_find_matching_catch_2=()=>findMatchingCatch([]);var ___cxa_find_matching_catch_3=arg0=>findMatchingCatch([arg0]);var ___cxa_rethrow=()=>{var info=exceptionCaught.pop();if(!info){abort("no exception to throw")}var ptr=info.excPtr;if(!info.get_rethrown()){exceptionCaught.push(info);info.set_rethrown(true);info.set_caught(false);uncaughtExceptionCount++}exceptionLast=ptr;throw exceptionLast};var ___cxa_throw=(ptr,type,destructor)=>{var info=new ExceptionInfo(ptr);info.init(type,destructor);exceptionLast=ptr;uncaughtExceptionCount++;throw exceptionLast};var ___cxa_uncaught_exceptions=()=>uncaughtExceptionCount;var SYSCALLS={varargs:undefined,get(){var ret=HEAP32[+SYSCALLS.varargs>>2];SYSCALLS.varargs+=4;return ret},getp(){return SYSCALLS.get()},getStr(ptr){var ret=UTF8ToString(ptr);return ret}};function ___syscall_fcntl64(fd,cmd,varargs){SYSCALLS.varargs=varargs;return 0}var ___syscall_fstat64=(fd,buf)=>{};var lengthBytesUTF8=str=>{var len=0;for(var i=0;i<str.length;++i){var c=str.charCodeAt(i);if(c<=127){len++}else if(c<=2047){len+=2}else if(c>=55296&&c<=57343){len+=4;++i}else{len+=3}}return len};var stringToUTF8Array=(str,heap,outIdx,maxBytesToWrite)=>{if(!(maxBytesToWrite>0))return 0;var startIdx=outIdx;var endIdx=outIdx+maxBytesToWrite-1;for(var i=0;i<str.length;++i){var u=str.charCodeAt(i);if(u>=55296&&u<=57343){var u1=str.charCodeAt(++i);u=65536+((u&1023)<<10)|u1&1023}if(u<=127){if(outIdx>=endIdx)break;heap[outIdx++]=u}else if(u<=2047){if(outIdx+1>=endIdx)break;heap[outIdx++]=192|u>>6;heap[outIdx++]=128|u&63}else if(u<=65535){if(outIdx+2>=endIdx)break;heap[outIdx++]=224|u>>12;heap[outIdx++]=128|u>>6&63;heap[outIdx++]=128|u&63}else{if(outIdx+3>=endIdx)break;heap[outIdx++]=240|u>>18;heap[outIdx++]=128|u>>12&63;heap[outIdx++]=128|u>>6&63;heap[outIdx++]=128|u&63}}heap[outIdx]=0;return outIdx-startIdx};var stringToUTF8=(str,outPtr,maxBytesToWrite)=>stringToUTF8Array(str,HEAPU8,outPtr,maxBytesToWrite);var ___syscall_getcwd=(buf,size)=>{};var ___syscall_getdents64=(fd,dirp,count)=>{};function ___syscall_ioctl(fd,op,varargs){SYSCALLS.varargs=varargs;return 0}var ___syscall_lstat64=(path,buf)=>{};var ___syscall_mkdirat=(dirfd,path,mode)=>{};var ___syscall_newfstatat=(dirfd,path,buf,flags)=>{};function ___syscall_openat(dirfd,path,flags,varargs){SYSCALLS.varargs=varargs}var ___syscall_readlinkat=(dirfd,path,buf,bufsize)=>{};var ___syscall_rmdir=path=>{};var ___syscall_stat64=(path,buf)=>{};var ___syscall_unlinkat=(dirfd,path,flags)=>{};var ___syscall_utimensat=(dirfd,path,times,flags)=>{};var nowIsMonotonic=true;var __emscripten_get_now_is_monotonic=()=>nowIsMonotonic;var __emscripten_throw_longjmp=()=>{throw Infinity};var isLeapYear=year=>year%4===0&&(year%100!==0||year%400===0);var MONTH_DAYS_LEAP_CUMULATIVE=[0,31,60,91,121,152,182,213,244,274,305,335];var MONTH_DAYS_REGULAR_CUMULATIVE=[0,31,59,90,120,151,181,212,243,273,304,334];var ydayFromDate=date=>{var leap=isLeapYear(date.getFullYear());var monthDaysCumulative=leap?MONTH_DAYS_LEAP_CUMULATIVE:MONTH_DAYS_REGULAR_CUMULATIVE;var yday=monthDaysCumulative[date.getMonth()]+date.getDate()-1;return yday};var convertI32PairToI53Checked=(lo,hi)=>hi+2097152>>>0<4194305-!!lo?(lo>>>0)+hi*4294967296:NaN;var __mktime_js=function(tmPtr){var ret=(()=>{var date=new Date(HEAP32[tmPtr+20>>2]+1900,HEAP32[tmPtr+16>>2],HEAP32[tmPtr+12>>2],HEAP32[tmPtr+8>>2],HEAP32[tmPtr+4>>2],HEAP32[tmPtr>>2],0);var dst=HEAP32[tmPtr+32>>2];var guessedOffset=date.getTimezoneOffset();var start=new Date(date.getFullYear(),0,1);var summerOffset=new Date(date.getFullYear(),6,1).getTimezoneOffset();var winterOffset=start.getTimezoneOffset();var dstOffset=Math.min(winterOffset,summerOffset);if(dst<0){HEAP32[tmPtr+32>>2]=Number(summerOffset!=winterOffset&&dstOffset==guessedOffset)}else if(dst>0!=(dstOffset==guessedOffset)){var nonDstOffset=Math.max(winterOffset,summerOffset);var trueOffset=dst>0?dstOffset:nonDstOffset;date.setTime(date.getTime()+(trueOffset-guessedOffset)*6e4)}HEAP32[tmPtr+24>>2]=date.getDay();var yday=ydayFromDate(date)|0;HEAP32[tmPtr+28>>2]=yday;HEAP32[tmPtr>>2]=date.getSeconds();HEAP32[tmPtr+4>>2]=date.getMinutes();HEAP32[tmPtr+8>>2]=date.getHours();HEAP32[tmPtr+12>>2]=date.getDate();HEAP32[tmPtr+16>>2]=date.getMonth();HEAP32[tmPtr+20>>2]=date.getYear();return date.getTime()/1e3})();return setTempRet0((tempDouble=ret,+Math.abs(tempDouble)>=1?tempDouble>0?+Math.floor(tempDouble/4294967296)>>>0:~~+Math.ceil((tempDouble-+(~~tempDouble>>>0))/4294967296)>>>0:0)),ret>>>0};function __mmap_js(len,prot,flags,fd,offset_low,offset_high,allocated,addr){var offset=convertI32PairToI53Checked(offset_low,offset_high);return-52}function __munmap_js(addr,len,prot,flags,fd,offset_low,offset_high){var offset=convertI32PairToI53Checked(offset_low,offset_high)}var stringToNewUTF8=str=>{var size=lengthBytesUTF8(str)+1;var ret=_malloc(size);if(ret)stringToUTF8(str,ret,size);return ret};var __tzset_js=(timezone,daylight,tzname)=>{var currentYear=(new Date).getFullYear();var winter=new Date(currentYear,0,1);var summer=new Date(currentYear,6,1);var winterOffset=winter.getTimezoneOffset();var summerOffset=summer.getTimezoneOffset();var stdTimezoneOffset=Math.max(winterOffset,summerOffset);HEAPU32[timezone>>2]=stdTimezoneOffset*60;HEAP32[daylight>>2]=Number(winterOffset!=summerOffset);function extractZone(date){var match=date.toTimeString().match(/\(([A-Za-z ]+)\)$/);return match?match[1]:"GMT"}var winterName=extractZone(winter);var summerName=extractZone(summer);var winterNamePtr=stringToNewUTF8(winterName);var summerNamePtr=stringToNewUTF8(summerName);if(summerOffset<winterOffset){HEAPU32[tzname>>2]=winterNamePtr;HEAPU32[tzname+4>>2]=summerNamePtr}else{HEAPU32[tzname>>2]=summerNamePtr;HEAPU32[tzname+4>>2]=winterNamePtr}};var _abort=()=>{abort("")};var _emscripten_date_now=()=>Date.now();var _emscripten_get_now;_emscripten_get_now=()=>performance.now();var _emscripten_memcpy_js=(dest,src,num)=>HEAPU8.copyWithin(dest,src,src+num);var getHeapMax=()=>2147483648;var growMemory=size=>{var b=wasmMemory.buffer;var pages=(size-b.byteLength+65535)/65536;try{wasmMemory.grow(pages);updateMemoryViews();return 1}catch(e){}};var _emscripten_resize_heap=requestedSize=>{var oldSize=HEAPU8.length;requestedSize>>>=0;var maxHeapSize=getHeapMax();if(requestedSize>maxHeapSize){return false}var alignUp=(x,multiple)=>x+(multiple-x%multiple)%multiple;for(var cutDown=1;cutDown<=4;cutDown*=2){var overGrownHeapSize=oldSize*(1+.2/cutDown);overGrownHeapSize=Math.min(overGrownHeapSize,requestedSize+100663296);var newSize=Math.min(maxHeapSize,alignUp(Math.max(requestedSize,overGrownHeapSize),65536));var replacement=growMemory(newSize);if(replacement){return true}}return false};var ENV={};var getExecutableName=()=>thisProgram||"./this.program";var getEnvStrings=()=>{if(!getEnvStrings.strings){var lang=(typeof navigator=="object"&&navigator.languages&&navigator.languages[0]||"C").replace("-","_")+".UTF-8";var env={"USER":"web_user","LOGNAME":"web_user","PATH":"/","PWD":"/","HOME":"/home/web_user","LANG":lang,"_":getExecutableName()};for(var x in ENV){if(ENV[x]===undefined)delete env[x];else env[x]=ENV[x]}var strings=[];for(var x in env){strings.push(`${x}=${env[x]}`)}getEnvStrings.strings=strings}return getEnvStrings.strings};var stringToAscii=(str,buffer)=>{for(var i=0;i<str.length;++i){HEAP8[buffer++>>0]=str.charCodeAt(i)}HEAP8[buffer>>0]=0};var _environ_get=(__environ,environ_buf)=>{var bufSize=0;getEnvStrings().forEach((string,i)=>{var ptr=environ_buf+bufSize;HEAPU32[__environ+i*4>>2]=ptr;stringToAscii(string,ptr);bufSize+=string.length+1});return 0};var _environ_sizes_get=(penviron_count,penviron_buf_size)=>{var strings=getEnvStrings();HEAPU32[penviron_count>>2]=strings.length;var bufSize=0;strings.forEach(string=>bufSize+=string.length+1);HEAPU32[penviron_buf_size>>2]=bufSize;return 0};var runtimeKeepaliveCounter=0;var keepRuntimeAlive=()=>noExitRuntime||runtimeKeepaliveCounter>0;var _proc_exit=code=>{EXITSTATUS=code;if(!keepRuntimeAlive()){if(Module["onExit"])Module["onExit"](code);ABORT=true}quit_(code,new ExitStatus(code))};var exitJS=(status,implicit)=>{EXITSTATUS=status;_proc_exit(status)};var _exit=exitJS;var _fd_close=fd=>52;var _fd_read=(fd,iov,iovcnt,pnum)=>52;function _fd_seek(fd,offset_low,offset_high,whence,newOffset){var offset=convertI32PairToI53Checked(offset_low,offset_high);return 70}var printCharBuffers=[null,[],[]];var printChar=(stream,curr)=>{var buffer=printCharBuffers[stream];if(curr===0||curr===10){(stream===1?out:err)(UTF8ArrayToString(buffer,0));buffer.length=0}else{buffer.push(curr)}};var _fd_write=(fd,iov,iovcnt,pnum)=>{var num=0;for(var i=0;i<iovcnt;i++){var ptr=HEAPU32[iov>>2];var len=HEAPU32[iov+4>>2];iov+=8;for(var j=0;j<len;j++){printChar(fd,HEAPU8[ptr+j])}num+=len}HEAPU32[pnum>>2]=num;return 0};var _llvm_eh_typeid_for=type=>type;var arraySum=(array,index)=>{var sum=0;for(var i=0;i<=index;sum+=array[i++]){}return sum};var MONTH_DAYS_LEAP=[31,29,31,30,31,30,31,31,30,31,30,31];var MONTH_DAYS_REGULAR=[31,28,31,30,31,30,31,31,30,31,30,31];var addDays=(date,days)=>{var newDate=new Date(date.getTime());while(days>0){var leap=isLeapYear(newDate.getFullYear());var currentMonth=newDate.getMonth();var daysInCurrentMonth=(leap?MONTH_DAYS_LEAP:MONTH_DAYS_REGULAR)[currentMonth];if(days>daysInCurrentMonth-newDate.getDate()){days-=daysInCurrentMonth-newDate.getDate()+1;newDate.setDate(1);if(currentMonth<11){newDate.setMonth(currentMonth+1)}else{newDate.setMonth(0);newDate.setFullYear(newDate.getFullYear()+1)}}else{newDate.setDate(newDate.getDate()+days);return newDate}}return newDate};function intArrayFromString(stringy,dontAddNull,length){var len=length>0?length:lengthBytesUTF8(stringy)+1;var u8array=new Array(len);var numBytesWritten=stringToUTF8Array(stringy,u8array,0,u8array.length);if(dontAddNull)u8array.length=numBytesWritten;return u8array}var writeArrayToMemory=(array,buffer)=>{HEAP8.set(array,buffer)};var _strftime=(s,maxsize,format,tm)=>{var tm_zone=HEAPU32[tm+40>>2];var date={tm_sec:HEAP32[tm>>2],tm_min:HEAP32[tm+4>>2],tm_hour:HEAP32[tm+8>>2],tm_mday:HEAP32[tm+12>>2],tm_mon:HEAP32[tm+16>>2],tm_year:HEAP32[tm+20>>2],tm_wday:HEAP32[tm+24>>2],tm_yday:HEAP32[tm+28>>2],tm_isdst:HEAP32[tm+32>>2],tm_gmtoff:HEAP32[tm+36>>2],tm_zone:tm_zone?UTF8ToString(tm_zone):""};var pattern=UTF8ToString(format);var EXPANSION_RULES_1={"%c":"%a %b %d %H:%M:%S %Y","%D":"%m/%d/%y","%F":"%Y-%m-%d","%h":"%b","%r":"%I:%M:%S %p","%R":"%H:%M","%T":"%H:%M:%S","%x":"%m/%d/%y","%X":"%H:%M:%S","%Ec":"%c","%EC":"%C","%Ex":"%m/%d/%y","%EX":"%H:%M:%S","%Ey":"%y","%EY":"%Y","%Od":"%d","%Oe":"%e","%OH":"%H","%OI":"%I","%Om":"%m","%OM":"%M","%OS":"%S","%Ou":"%u","%OU":"%U","%OV":"%V","%Ow":"%w","%OW":"%W","%Oy":"%y"};for(var rule in EXPANSION_RULES_1){pattern=pattern.replace(new RegExp(rule,"g"),EXPANSION_RULES_1[rule])}var WEEKDAYS=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];var MONTHS=["January","February","March","April","May","June","July","August","September","October","November","December"];function leadingSomething(value,digits,character){var str=typeof value=="number"?value.toString():value||"";while(str.length<digits){str=character[0]+str}return str}function leadingNulls(value,digits){return leadingSomething(value,digits,"0")}function compareByDay(date1,date2){function sgn(value){return value<0?-1:value>0?1:0}var compare;if((compare=sgn(date1.getFullYear()-date2.getFullYear()))===0){if((compare=sgn(date1.getMonth()-date2.getMonth()))===0){compare=sgn(date1.getDate()-date2.getDate())}}return compare}function getFirstWeekStartDate(janFourth){switch(janFourth.getDay()){case 0:return new Date(janFourth.getFullYear()-1,11,29);case 1:return janFourth;case 2:return new Date(janFourth.getFullYear(),0,3);case 3:return new Date(janFourth.getFullYear(),0,2);case 4:return new Date(janFourth.getFullYear(),0,1);case 5:return new Date(janFourth.getFullYear()-1,11,31);case 6:return new Date(janFourth.getFullYear()-1,11,30)}}function getWeekBasedYear(date){var thisDate=addDays(new Date(date.tm_year+1900,0,1),date.tm_yday);var janFourthThisYear=new Date(thisDate.getFullYear(),0,4);var janFourthNextYear=new Date(thisDate.getFullYear()+1,0,4);var firstWeekStartThisYear=getFirstWeekStartDate(janFourthThisYear);var firstWeekStartNextYear=getFirstWeekStartDate(janFourthNextYear);if(compareByDay(firstWeekStartThisYear,thisDate)<=0){if(compareByDay(firstWeekStartNextYear,thisDate)<=0){return thisDate.getFullYear()+1}return thisDate.getFullYear()}return thisDate.getFullYear()-1}var EXPANSION_RULES_2={"%a":date=>WEEKDAYS[date.tm_wday].substring(0,3),"%A":date=>WEEKDAYS[date.tm_wday],"%b":date=>MONTHS[date.tm_mon].substring(0,3),"%B":date=>MONTHS[date.tm_mon],"%C":date=>{var year=date.tm_year+1900;return leadingNulls(year/100|0,2)},"%d":date=>leadingNulls(date.tm_mday,2),"%e":date=>leadingSomething(date.tm_mday,2," "),"%g":date=>getWeekBasedYear(date).toString().substring(2),"%G":date=>getWeekBasedYear(date),"%H":date=>leadingNulls(date.tm_hour,2),"%I":date=>{var twelveHour=date.tm_hour;if(twelveHour==0)twelveHour=12;else if(twelveHour>12)twelveHour-=12;return leadingNulls(twelveHour,2)},"%j":date=>leadingNulls(date.tm_mday+arraySum(isLeapYear(date.tm_year+1900)?MONTH_DAYS_LEAP:MONTH_DAYS_REGULAR,date.tm_mon-1),3),"%m":date=>leadingNulls(date.tm_mon+1,2),"%M":date=>leadingNulls(date.tm_min,2),"%n":()=>"\n","%p":date=>{if(date.tm_hour>=0&&date.tm_hour<12){return"AM"}return"PM"},"%S":date=>leadingNulls(date.tm_sec,2),"%t":()=>"\t","%u":date=>date.tm_wday||7,"%U":date=>{var days=date.tm_yday+7-date.tm_wday;return leadingNulls(Math.floor(days/7),2)},"%V":date=>{var val=Math.floor((date.tm_yday+7-(date.tm_wday+6)%7)/7);if((date.tm_wday+371-date.tm_yday-2)%7<=2){val++}if(!val){val=52;var dec31=(date.tm_wday+7-date.tm_yday-1)%7;if(dec31==4||dec31==5&&isLeapYear(date.tm_year%400-1)){val++}}else if(val==53){var jan1=(date.tm_wday+371-date.tm_yday)%7;if(jan1!=4&&(jan1!=3||!isLeapYear(date.tm_year)))val=1}return leadingNulls(val,2)},"%w":date=>date.tm_wday,"%W":date=>{var days=date.tm_yday+7-(date.tm_wday+6)%7;return leadingNulls(Math.floor(days/7),2)},"%y":date=>(date.tm_year+1900).toString().substring(2),"%Y":date=>date.tm_year+1900,"%z":date=>{var off=date.tm_gmtoff;var ahead=off>=0;off=Math.abs(off)/60;off=off/60*100+off%60;return(ahead?"+":"-")+String("0000"+off).slice(-4)},"%Z":date=>date.tm_zone,"%%":()=>"%"};pattern=pattern.replace(/%%/g,"\0\0");for(var rule in EXPANSION_RULES_2){if(pattern.includes(rule)){pattern=pattern.replace(new RegExp(rule,"g"),EXPANSION_RULES_2[rule](date))}}pattern=pattern.replace(/\0\0/g,"%");var bytes=intArrayFromString(pattern,false);if(bytes.length>maxsize){return 0}writeArrayToMemory(bytes,s);return bytes.length-1};var _strftime_l=(s,maxsize,format,tm,loc)=>_strftime(s,maxsize,format,tm);var wasmTableMirror=[];var wasmTable;var getWasmTableEntry=funcPtr=>{var func=wasmTableMirror[funcPtr];if(!func){if(funcPtr>=wasmTableMirror.length)wasmTableMirror.length=funcPtr+1;wasmTableMirror[funcPtr]=func=wasmTable.get(funcPtr)}return func};var wasmImports={i:___assert_fail,q:___cxa_begin_catch,w:___cxa_end_catch,a:___cxa_find_matching_catch_2,h:___cxa_find_matching_catch_3,L:___cxa_rethrow,A:___cxa_throw,cb:___cxa_uncaught_exceptions,e:___resumeException,ba:___syscall_fcntl64,xa:___syscall_fstat64,ta:___syscall_getcwd,pa:___syscall_getdents64,Aa:___syscall_ioctl,ua:___syscall_lstat64,qa:___syscall_mkdirat,va:___syscall_newfstatat,Z:___syscall_openat,Y:___syscall_readlinkat,W:___syscall_rmdir,wa:___syscall_stat64,X:___syscall_unlinkat,na:___syscall_utimensat,ya:__emscripten_get_now_is_monotonic,db:__emscripten_throw_longjmp,_a:__mktime_js,Ya:__mmap_js,Za:__munmap_js,oa:__tzset_js,u:_abort,$:_emscripten_date_now,_:_emscripten_get_now,za:_emscripten_memcpy_js,ma:_emscripten_resize_heap,ra:_environ_get,sa:_environ_sizes_get,B:_exit,J:_fd_close,aa:_fd_read,$a:_fd_seek,Q:_fd_write,t:invoke_di,M:invoke_dii,D:invoke_diii,Da:invoke_fif,la:invoke_fiii,s:invoke_i,c:invoke_ii,O:invoke_iidd,Ta:invoke_iidddddd,La:invoke_iiddiii,d:invoke_iii,ga:invoke_iiiddddd,Na:invoke_iiiddiii,N:invoke_iiiff,Oa:invoke_iiiffff,j:invoke_iiii,k:invoke_iiiii,Fa:invoke_iiiiid,Va:invoke_iiiiifi,p:invoke_iiiiii,S:invoke_iiiiiiddiiiii,m:invoke_iiiiiii,y:invoke_iiiiiiii,C:invoke_iiiiiiiii,G:invoke_iiiiiiiiii,T:invoke_iiiiiiiiiii,E:invoke_iiiiiiiiiiii,ia:invoke_iiiiiiiiiiiiiiiiiiiiiiiiiii,ab:invoke_ji,Xa:invoke_jiiii,n:invoke_v,Ga:invoke_vdii,b:invoke_vi,I:invoke_vid,U:invoke_vidd,ja:invoke_viddd,Sa:invoke_vidddddddd,Ua:invoke_viddi,ka:invoke_vidi,Pa:invoke_viffffi,g:invoke_vii,z:invoke_viid,Ka:invoke_viidddd,Ja:invoke_viiddddddi,Ba:invoke_viiddiiiii,Ea:invoke_viif,f:invoke_viii,Ha:invoke_viiid,da:invoke_viiiddiiiii,fa:invoke_viiiddiiiiii,Ia:invoke_viiidi,Ma:invoke_viiidiiiddddd,l:invoke_viiii,K:invoke_viiiid,ca:invoke_viiiidii,o:invoke_viiiii,ea:invoke_viiiiid,ha:invoke_viiiiiff,r:invoke_viiiiii,x:invoke_viiiiiii,F:invoke_viiiiiiii,V:invoke_viiiiiiiii,H:invoke_viiiiiiiiii,R:invoke_viiiiiiiiiii,Ca:invoke_viiiiiiiiiiii,Wa:invoke_viiiiiiiiiiiiii,P:invoke_viiiiiiiiiiiiiii,Qa:js_free_id,Ra:js_get_stream_id,v:_llvm_eh_typeid_for,bb:_strftime_l};var wasmExports=createWasm();var ___wasm_call_ctors=()=>(___wasm_call_ctors=wasmExports["fb"])();var ___cxa_free_exception=a0=>(___cxa_free_exception=wasmExports["__cxa_free_exception"])(a0);var _malloc=Module["_malloc"]=a0=>(_malloc=Module["_malloc"]=wasmExports["hb"])(a0);var _free=Module["_free"]=a0=>(_free=Module["_free"]=wasmExports["ib"])(a0);var setTempRet0=a0=>(setTempRet0=wasmExports["jb"])(a0);var ___errno_location=()=>(___errno_location=wasmExports["kb"])();var _InitializeFontsBin=Module["_InitializeFontsBin"]=(a0,a1)=>(_InitializeFontsBin=Module["_InitializeFontsBin"]=wasmExports["lb"])(a0,a1);var _InitializeFontsBase64=Module["_InitializeFontsBase64"]=(a0,a1)=>(_InitializeFontsBase64=Module["_InitializeFontsBase64"]=wasmExports["mb"])(a0,a1);var _InitializeFontsRanges=Module["_InitializeFontsRanges"]=a0=>(_InitializeFontsRanges=Module["_InitializeFontsRanges"]=wasmExports["nb"])(a0);var _SetFontBinary=Module["_SetFontBinary"]=(a0,a1,a2)=>(_SetFontBinary=Module["_SetFontBinary"]=wasmExports["ob"])(a0,a1,a2);var _IsFontBinaryExist=Module["_IsFontBinaryExist"]=a0=>(_IsFontBinaryExist=Module["_IsFontBinaryExist"]=wasmExports["pb"])(a0);var _GetType=Module["_GetType"]=(a0,a1)=>(_GetType=Module["_GetType"]=wasmExports["qb"])(a0,a1);var _Open=Module["_Open"]=(a0,a1,a2)=>(_Open=Module["_Open"]=wasmExports["rb"])(a0,a1,a2);var _GetErrorCode=Module["_GetErrorCode"]=a0=>(_GetErrorCode=Module["_GetErrorCode"]=wasmExports["sb"])(a0);var _Close=Module["_Close"]=a0=>(_Close=Module["_Close"]=wasmExports["tb"])(a0);var _GetInfo=Module["_GetInfo"]=a0=>(_GetInfo=Module["_GetInfo"]=wasmExports["ub"])(a0);var _GetPixmap=Module["_GetPixmap"]=(a0,a1,a2,a3,a4)=>(_GetPixmap=Module["_GetPixmap"]=wasmExports["vb"])(a0,a1,a2,a3,a4);var _GetGlyphs=Module["_GetGlyphs"]=(a0,a1)=>(_GetGlyphs=Module["_GetGlyphs"]=wasmExports["wb"])(a0,a1);var _GetLinks=Module["_GetLinks"]=(a0,a1)=>(_GetLinks=Module["_GetLinks"]=wasmExports["xb"])(a0,a1);var _GetStructure=Module["_GetStructure"]=a0=>(_GetStructure=Module["_GetStructure"]=wasmExports["yb"])(a0);var _GetInteractiveFormsInfo=Module["_GetInteractiveFormsInfo"]=a0=>(_GetInteractiveFormsInfo=Module["_GetInteractiveFormsInfo"]=wasmExports["zb"])(a0);var _GetInteractiveFormsFonts=Module["_GetInteractiveFormsFonts"]=(a0,a1)=>(_GetInteractiveFormsFonts=Module["_GetInteractiveFormsFonts"]=wasmExports["Ab"])(a0,a1);var _GetInteractiveFormsAP=Module["_GetInteractiveFormsAP"]=(a0,a1,a2,a3,a4,a5,a6,a7)=>(_GetInteractiveFormsAP=Module["_GetInteractiveFormsAP"]=wasmExports["Bb"])(a0,a1,a2,a3,a4,a5,a6,a7);var _GetButtonIcons=Module["_GetButtonIcons"]=(a0,a1,a2,a3,a4,a5)=>(_GetButtonIcons=Module["_GetButtonIcons"]=wasmExports["Cb"])(a0,a1,a2,a3,a4,a5);var _GetAnnotationsInfo=Module["_GetAnnotationsInfo"]=(a0,a1)=>(_GetAnnotationsInfo=Module["_GetAnnotationsInfo"]=wasmExports["Db"])(a0,a1);var _GetAnnotationsAP=Module["_GetAnnotationsAP"]=(a0,a1,a2,a3,a4,a5,a6)=>(_GetAnnotationsAP=Module["_GetAnnotationsAP"]=wasmExports["Eb"])(a0,a1,a2,a3,a4,a5,a6);var _GetFontBinary=Module["_GetFontBinary"]=(a0,a1)=>(_GetFontBinary=Module["_GetFontBinary"]=wasmExports["Fb"])(a0,a1);var _DestroyTextInfo=Module["_DestroyTextInfo"]=a0=>(_DestroyTextInfo=Module["_DestroyTextInfo"]=wasmExports["Gb"])(a0);var _IsNeedCMap=Module["_IsNeedCMap"]=a0=>(_IsNeedCMap=Module["_IsNeedCMap"]=wasmExports["Hb"])(a0);var _SetCMapData=Module["_SetCMapData"]=(a0,a1,a2)=>(_SetCMapData=Module["_SetCMapData"]=wasmExports["Ib"])(a0,a1,a2);var _ScanPage=Module["_ScanPage"]=(a0,a1,a2)=>(_ScanPage=Module["_ScanPage"]=wasmExports["Jb"])(a0,a1,a2);var _GetImageBase64=Module["_GetImageBase64"]=(a0,a1)=>(_GetImageBase64=Module["_GetImageBase64"]=wasmExports["Kb"])(a0,a1);var _GetImageBase64Len=Module["_GetImageBase64Len"]=a0=>(_GetImageBase64Len=Module["_GetImageBase64Len"]=wasmExports["Lb"])(a0);var _GetImageBase64Ptr=Module["_GetImageBase64Ptr"]=a0=>(_GetImageBase64Ptr=Module["_GetImageBase64Ptr"]=wasmExports["Mb"])(a0);var _GetImageBase64Free=Module["_GetImageBase64Free"]=a0=>(_GetImageBase64Free=Module["_GetImageBase64Free"]=wasmExports["Nb"])(a0);var _setThrew=(a0,a1)=>(_setThrew=wasmExports["Ob"])(a0,a1);var stackSave=()=>(stackSave=wasmExports["Pb"])();var stackRestore=a0=>(stackRestore=wasmExports["Qb"])(a0);var ___cxa_decrement_exception_refcount=a0=>(___cxa_decrement_exception_refcount=wasmExports["Rb"])(a0);var ___cxa_increment_exception_refcount=a0=>(___cxa_increment_exception_refcount=wasmExports["Sb"])(a0);var ___cxa_can_catch=(a0,a1,a2)=>(___cxa_can_catch=wasmExports["Tb"])(a0,a1,a2);var ___cxa_is_pointer_type=a0=>(___cxa_is_pointer_type=wasmExports["Ub"])(a0);var dynCall_ji=Module["dynCall_ji"]=(a0,a1)=>(dynCall_ji=Module["dynCall_ji"]=wasmExports["Vb"])(a0,a1);var dynCall_jiiii=Module["dynCall_jiiii"]=(a0,a1,a2,a3,a4)=>(dynCall_jiiii=Module["dynCall_jiiii"]=wasmExports["Wb"])(a0,a1,a2,a3,a4);var ___start_em_js=Module["___start_em_js"]=2729080;var ___stop_em_js=Module["___stop_em_js"]=2729249;function invoke_iii(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiii(index,a1,a2,a3,a4,a5){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_ii(index,a1){var sp=stackSave();try{return getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vii(index,a1,a2){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiii(index,a1,a2,a3,a4){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viii(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vi(index,a1){var sp=stackSave();try{getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiii(index,a1,a2,a3,a4,a5){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiii(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_v(index){var sp=stackSave();try{getWasmTableEntry(index)()}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_i(index){var sp=stackSave();try{return getWasmTableEntry(index)()}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiifi(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viddi(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vidi(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iidddddd(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vid(index,a1,a2){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vidd(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iidd(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vidddddddd(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viddd(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiiiiiiiiiiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiddiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viffffi(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiffff(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiff(index,a1,a2,a3,a4){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiff(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viid(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiddiii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiidiiiddddd(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_dii(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiddiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiddddd(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiddiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_di(index,a1){var sp=stackSave();try{return getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viidddd(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiddddddi(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiid(index,a1,a2,a3,a4,a5){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiidi(index,a1,a2,a3,a4,a5){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiid(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vdii(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiid(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiid(index,a1,a2,a3,a4,a5){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_diii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viif(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_fif(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiddiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiidii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiddiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_fiii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_ji(index,a1){var sp=stackSave();try{return dynCall_ji(index,a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_jiiii(index,a1,a2,a3,a4){var sp=stackSave();try{return dynCall_jiiii(index,a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}var calledRun;dependenciesFulfilled=function runCaller(){if(!calledRun)run();if(!calledRun)dependenciesFulfilled=runCaller};function run(){if(runDependencies>0){return}preRun();if(runDependencies>0){return}function doRun(){if(calledRun)return;calledRun=true;Module["calledRun"]=true;if(ABORT)return;initRuntime();if(Module["onRuntimeInitialized"])Module["onRuntimeInitialized"]();postRun()}if(Module["setStatus"]){Module["setStatus"]("Running...");setTimeout(function(){setTimeout(function(){Module["setStatus"]("")},1);doRun()},1)}else{doRun()}}if(Module["preInit"]){if(typeof Module["preInit"]=="function")Module["preInit"]=[Module["preInit"]];while(Module["preInit"].length>0){Module["preInit"].pop()()}}run();


	self.drawingFile = null;

	function CBinaryReader(data, start, size)
	{
		this.data = data;
		this.pos = start;
		this.limit = start + size;
	}
	CBinaryReader.prototype.readByte = function()
	{
		let val = this.data[this.pos];
		this.pos += 1;
		return val;
	};
	CBinaryReader.prototype.readInt = function()
	{
		let val = this.data[this.pos] | this.data[this.pos + 1] << 8 | this.data[this.pos + 2] << 16 | this.data[this.pos + 3] << 24;
		this.pos += 4;
		return val;
	};
	CBinaryReader.prototype.readDouble = function()
	{
		return this.readInt() / 100;
	};
	CBinaryReader.prototype.readDouble2 = function()
	{
		return this.readInt() / 10000;
	};
	CBinaryReader.prototype.readString = function()
	{
		let len = this.readInt();
		let val = String.prototype.fromUtf8(this.data, this.pos, len);
		this.pos += len;
		return val;
	};
	CBinaryReader.prototype.readData = function()
	{
		let len = this.readInt();
		let val = this.data.slice(this.pos, this.pos + len);
		this.pos += len;
		return val;
	};
	CBinaryReader.prototype.isValid = function()
	{
		return (this.pos < this.limit) ? true : false;
	};
	CBinaryReader.prototype.Skip = function(nPos)
	{
		this.pos += nPos;
	};

	function CBinaryWriter()
	{
		this.size = 100000;
		this.dataSize = 0;
		this.buffer = new Uint8Array(this.size);
	}
	CBinaryWriter.prototype.checkAlloc = function(addition)
	{
		if ((this.dataSize + addition) <= this.size)
			return;

		let newSize = Math.max(this.size * 2, this.size + addition);
		let newBuffer = new Uint8Array(newSize);
		newBuffer.set(this.buffer, 0);

		this.size = newSize;
		this.buffer = newBuffer;
	};
	CBinaryWriter.prototype.writeUint = function(value)
	{
		this.checkAlloc(4);
		let val = (value>2147483647)?value-4294967296:value;
		this.buffer[this.dataSize++] = (val) & 0xFF;
		this.buffer[this.dataSize++] = (val >>> 8) & 0xFF;
		this.buffer[this.dataSize++] = (val >>> 16) & 0xFF;
		this.buffer[this.dataSize++] = (val >>> 24) & 0xFF;
	};
	CBinaryWriter.prototype.writeString = function(value)
	{
		let valueUtf8 = value.toUtf8();
		this.checkAlloc(valueUtf8.length);
		this.buffer.set(valueUtf8, this.dataSize);
		this.dataSize += valueUtf8.length;
	};

	var UpdateFontsSource = {
		Undefined   : 0,
		Page        : 1,
		Annotation  : 2,
		Forms       : 4
	};

	function CFile()
	{
		this.nativeFile = 0;
		this.stream = -1;
		this.stream_size = 0;
		this.type = -1;
		this.pages = [];
		this.info = null;
		this._isNeedPassword = false;

		// for async fonts loader
		this.fontPageIndex = -1;
		this.fontPageUpdateType = UpdateFontsSource.Undefined;
		this.fontStreams = {};

		this.scannedImages = {};
	}

	CFile.prototype["loadFromData"] = function(arrayBuffer)
	{
		let data = new Uint8Array(arrayBuffer);
		let _stream = Module["_malloc"](data.length);
		Module["HEAP8"].set(data, _stream);
		this.nativeFile = Module["_Open"](_stream, data.length, 0);
		let error = Module["_GetErrorCode"](this.nativeFile);
		this.stream = _stream;
		this.stream_size = data.length;
		this.type = Module["_GetType"](_stream, data.length);
		self.drawingFile = this;
		if (!error)
			this.getInfo();
		this._isNeedPassword = (4 === error) ? true : false;

		// 0 - ok
		// 4 - password
		// else - error
		return error;
	};
	CFile.prototype["loadFromDataWithPassword"] = function(password)
	{
		if (0 != this.nativeFile)
			Module["_Close"](this.nativeFile);

		let passBuffer = password.toUtf8();
		let passPointer = Module["_malloc"](passBuffer.length);
		Module["HEAP8"].set(passBuffer, passPointer);
		this.nativeFile = Module["_Open"](this.stream, this.stream_size, passPointer);
		Module["_free"](passPointer);
		let error = Module["_GetErrorCode"](this.nativeFile);
		this.type = Module["_GetType"](this.stream, this.stream_size);
		self.drawingFile = this;
		if (!error)
			this.getInfo();
		this._isNeedPassword = (4 === error) ? true : false;

		// 0 - ok
		// 4 - password
		// else - error
		return error;
	};
	CFile.prototype["getFileAsBase64"] = function()
	{
		if (0 >= this.stream)
			return "";

		return new Uint8Array(Module["HEAP8"].buffer, this.stream, this.stream_size);
	};
	CFile.prototype["isNeedPassword"] = function()
	{
		return this._isNeedPassword;
	};
	CFile.prototype["isNeedCMap"] = function()
	{
		if (!this.nativeFile)
			return false;

		let isNeed = Module["_IsNeedCMap"](this.nativeFile);
		return (isNeed === 1) ? true : false;
	};
	CFile.prototype["setCMap"] = function(memoryBuffer)
	{
		if (!this.nativeFile)
			return;

		let pointer = Module["_malloc"](memoryBuffer.length);
		Module.HEAP8.set(memoryBuffer, pointer);
		Module["_SetCMapData"](this.nativeFile, pointer, memoryBuffer.length);
	};
	CFile.prototype["getInfo"] = function()
	{
		if (!this.nativeFile)
			return false;

		let _info = Module["_GetInfo"](this.nativeFile);
		if (_info == 0)
			return false;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, _info, 4);
		if (lenArray == null)
			return false;

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
			return false;

		let buffer = new Uint8Array(Module["HEAP8"].buffer, _info + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);

		this.StartID = reader.readInt();

		let _pages = reader.readInt();
		for (let i = 0; i < _pages; i++)
		{
			let rec = {};
			rec["W"] = reader.readInt();
			rec["H"] = reader.readInt();
			rec["Dpi"] = reader.readInt();
			rec["Rotate"] = reader.readInt();
			rec["originIndex"] = i;
			rec.fonts = [];
			rec.fontsUpdateType = UpdateFontsSource.Undefined;
			rec.text = null;
			this.pages.push(rec);
		}
		let json_info = reader.readString();
		try
		{
			this.info = JSON.parse(json_info);
		} catch(err) {}

		Module["_free"](_info);
		return this.pages.length > 0;
	};
	CFile.prototype["close"] = function()
	{
		Module["_Close"](this.nativeFile);
		this.nativeFile = 0;
		this.pages = [];
		this.info = null;
		this.StartID = null;
		if (this.stream > 0)
			Module["_free"](this.stream);
		this.stream = -1;
		self.drawingFile = null;
	};

	CFile.prototype["getPages"] = function()
	{
		return this.pages;
	};

	CFile.prototype["openForms"] = function()
	{

	};

	CFile.prototype["getDocumentInfo"] = function()
	{
		return this.info;
	};
	
	CFile.prototype["getStartID"] = function()
	{
		return this.StartID;
	};

	function getOriginPage(pages, originIndex)
	{
		for (let i = 0; i < pages.length; ++i)
		{
			if (pages[i]["originIndex"] == originIndex)
				return pages[i];
		}
		return null;
	}

	CFile.prototype["getPagePixmap"] = function(pageIndex, width, height, backgroundColor)
	{
		let page = getOriginPage(this.pages, pageIndex);
		if (!page || page.fonts.length > 0)
		{
			// waiting fonts
			return null;
		}

		this.lockPageNumForFontsLoader(pageIndex, UpdateFontsSource.Page);
		let retValue = Module["_GetPixmap"](this.nativeFile, pageIndex, width, height, backgroundColor === undefined ? 0xFFFFFF : backgroundColor);
		this.unlockPageNumForFontsLoader();

		if (page.fonts.length > 0)
		{
			// waiting fonts
			Module["_free"](retValue);
			retValue = null;
		}
		return retValue;
	};
	CFile.prototype["getGlyphs"] = function(pageIndex)
	{
		let page = getOriginPage(this.pages, pageIndex);
		if (!page || page.fonts.length > 0)
		{
			// waiting fonts
			return null;
		}

		this.lockPageNumForFontsLoader(pageIndex, UpdateFontsSource.Page);
		let retValue = Module["_GetGlyphs"](this.nativeFile, pageIndex);
		// there is no need to delete the result; this buffer is used as a text buffer 
		// for text commands on other pages. After receiving ALL text pages, 
		// you need to call destroyTextInfo()
		this.unlockPageNumForFontsLoader();

		if (page.fonts.length > 0)
		{
			// waiting fonts
			retValue = null;
		}

		if (null == retValue)
			return null;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, retValue, 5);
		let len = lenArray[0];
		len -= 20;

		if (this.onUpdateStatistics)
			this.onUpdateStatistics(lenArray[1], lenArray[2], lenArray[3], lenArray[4]);

		if (len <= 0)
		{
			return [];
		}

		let textCommandsSrc = new Uint8Array(Module["HEAP8"].buffer, retValue + 20, len);
		let textCommands = new Uint8Array(len);
		textCommands.set(textCommandsSrc);

		textCommandsSrc = null;
		return textCommands;
	};
	CFile.prototype["destroyTextInfo"] = function()
	{
		Module["_DestroyTextInfo"]();
	};
	CFile.prototype["getLinks"] = function(pageIndex)
	{
		let res = [];
		let ext = Module["_GetLinks"](this.nativeFile, pageIndex);
		if (ext == 0)
			return res;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, ext, 4);
		if (lenArray == null)
			return res;

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
			return res;

		let buffer = new Uint8Array(Module["HEAP8"].buffer, ext + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);

		while (reader.isValid())
		{
			let rec = {};
			rec["link"] = reader.readString();
			rec["dest"] = reader.readDouble();
			rec["x"] = reader.readDouble();
			rec["y"] = reader.readDouble();
			rec["w"] = reader.readDouble();
			rec["h"] = reader.readDouble();
			res.push(rec);
		}

		Module["_free"](ext);
		return res;
	};

	function readAction(reader, rec)
	{
		let SType = reader.readByte();
		// 0 - Unknown, 1 - GoTo, 2 - GoToR, 3 - GoToE, 4 - Launch
		// 5 - Thread, 6 - URI, 7 - Sound, 8 - Movie, 9 - Hide
		// 10 - Named, 11 - SubmitForm, 12 - ResetForm, 13 - ImportData
		// 14 - JavaScript, 15 - SetOCGState, 16 - Rendition
		// 17 - Trans, 18 - GoTo3DView
		rec["S"] = SType;
		if (SType == 14)
		{
			rec["JS"] = reader.readString();
		}
		else if (SType == 1)
		{
			rec["page"] = reader.readInt();
			rec["kind"] = reader.readByte();
			// 0 - XYZ
			// 1 - Fit
			// 2 - FitH
			// 3 - FitV
			// 4 - FitR
			// 5 - FitB
			// 6 - FitBH
			// 7 - FitBV
			switch (rec["kind"])
			{
				case 0:
				case 2:
				case 3:
				case 6:
				case 7:
				{
					let nFlag = reader.readByte();
					if (nFlag & (1 << 0))
						rec["left"] = reader.readDouble();
					if (nFlag & (1 << 1))
						rec["top"]  = reader.readDouble();
					if (nFlag & (1 << 2))
						rec["zoom"] = reader.readDouble();
					break;
				}
				case 4:
				{
					rec["left"]   = reader.readDouble();
					rec["bottom"] = reader.readDouble();
					rec["right"]  = reader.readDouble();
					rec["top"]    = reader.readDouble();
					break;
				}
				case 1:
				case 5:
				default:
					break;
			}
		}
		else if (SType == 10)
		{
			rec["N"] = reader.readString();
		}
		else if (SType == 6)
		{
			rec["URI"] = reader.readString();
		}
		else if (SType == 9)
		{
			rec["H"] = reader.readByte();
			let m = reader.readInt();
		    rec["T"] = [];
			// array of annotation names - rec["name"]
		    for (let j = 0; j < m; ++j)
				rec["T"].push(reader.readString());
		}
		else if (SType == 12)
		{
			rec["Flags"] = reader.readInt();
			let m = reader.readInt();
		    rec["Fields"] = [];
			// array of annotation names - rec["name"]
		    for (let j = 0; j < m; ++j)
				rec["Fields"].push(reader.readString());
		}
		let NextAction = reader.readByte();
		if (NextAction)
		{
			rec["Next"] = {};
			readAction(reader, rec["Next"]);
		}
	}
	function readAnnot(reader, rec)
	{
		rec["AP"] = {};
		// Annot
		// number for relations with AP
		rec["AP"]["i"] = reader.readInt();
		rec["annotflag"] = reader.readInt();
		// 12.5.3
		let bHidden = (rec["annotflag"] >> 1) & 1; // Hidden
		let bPrint = (rec["annotflag"] >> 2) & 1; // Print
		rec["noZoom"] = (rec["annotflag"] >> 3) & 1; // NoZoom
		rec["noRotate"] = (rec["annotflag"] >> 4) & 1; // NoRotate
		let bNoView = (rec["annotflag"] >> 5) & 1; // NoView
		rec["locked"] = (rec["annotflag"] >> 7) & 1; // Locked
		rec["ToggleNoView"] = (rec["annotflag"] >> 8) & 1; // ToggleNoView
		rec["lockedC"] = (rec["annotflag"] >> 9) & 1; // LockedContents
		// 0 - visible, 1 - hidden, 2 - noPrint, 3 - noView
		rec["display"] = 0;
		if (bHidden)
			rec["display"] = 1;
		else
		{
			if (bPrint)
			{
				if (bNoView)
					rec["display"] = 3;
				else
					rec["display"] = 0;
			}
			else
			{
				if (bNoView)
					rec["display"] = 0; // ??? no hidden, but noView and no print
				else
					rec["display"] = 2;
			}
		}
		rec["page"] = reader.readInt();
		// offsets like getStructure and viewer.navigate
		rec["rect"] = {};
		rec["rect"]["x1"] = reader.readDouble2();
		rec["rect"]["y1"] = reader.readDouble2();
		rec["rect"]["x2"] = reader.readDouble2();
		rec["rect"]["y2"] = reader.readDouble2();
		let flags = reader.readInt();
		// Unique name - NM
		if (flags & (1 << 0))
			rec["UniqueName"] = reader.readString();
		// Alternate annotation text - Contents
		if (flags & (1 << 1))
			rec["Contents"] = reader.readString();
		// Border effect - BE
		if (flags & (1 << 2))
		{
			rec["BE"] = {};
			rec["BE"]["S"] = reader.readByte();
			rec["BE"]["I"] = reader.readDouble();
		}
		// Special annotation color - С
		if (flags & (1 << 3))
		{
			let n = reader.readInt();
			rec["C"] = [];
			for (let i = 0; i < n; ++i)
				rec["C"].push(reader.readDouble2());
		}
		// Border/BS
		if (flags & (1 << 4))
		{
			// 0 - solid, 1 - beveled, 2 - dashed, 3 - inset, 4 - underline
			rec["border"] = reader.readByte();
			rec["borderWidth"] = reader.readDouble();
			// Border Dash Pattern
			if (rec["border"] == 2)
			{
				let n = reader.readInt();
				rec["dashed"] = [];
				for (let i = 0; i < n; ++i)
					rec["dashed"].push(reader.readDouble());
			}
		}
		// Date of last change - M
		if (flags & (1 << 5))
			rec["LastModified"] = reader.readString();
		rec["AP"]["have"] = (flags >> 6) & 1;
	}
	function readAnnotAP(reader, AP)
	{
		// number for relations with AP
		AP["i"] = reader.readInt();
		AP["x"] = reader.readInt();
		AP["y"] = reader.readInt();
		AP["w"] = reader.readInt();
		AP["h"] = reader.readInt();
		let n = reader.readInt();
		for (let i = 0; i < n; ++i)
		{
			let APType = reader.readString();
			if (!AP[APType])
				AP[APType] = {};
			let APi = AP[APType];
			let ASType = reader.readString();
			if (ASType)
			{
				AP[APType][ASType] = {};
				APi = AP[APType][ASType];
			}
			let np1 = reader.readInt();
			let np2 = reader.readInt();
			// this memory needs to be deleted
			APi["retValue"] = np2 << 32 | np1;
			// 0 - Normal, 1 - Multiply, 2 - Screen, 3 - Overlay, 4 - Darken, 5 - Lighten, 6 - ColorDodge, 7 - ColorBurn, 8 - HardLight,
			// 9 - SoftLight, 10 - Difference, 11 - Exclusion, 12 - Hue, 13 - Saturation, 14 - Color, 15 - Luminosity
			APi["BlendMode"] = reader.readByte();
		}
	}
	function getWidgetFonts(nativeFile, type)
	{
		let res = [];
		let ext = Module["_GetInteractiveFormsFonts"](nativeFile, type);
		if (ext == 0)
			return res;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, ext, 4);
		if (lenArray == null)
		{
			Module["_free"](ext);
			return res;
		}

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
		{
			Module["_free"](ext);
			return res;
		}

		let buffer = new Uint8Array(Module["HEAP8"].buffer, ext + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);
		
		while (reader.isValid())
		{
			let n = reader.readInt();
			for (let i = 0; i < n; ++i)
				res.push(reader.readString());
		}
		
		Module["_free"](ext);
		return res;
	}

	CFile.prototype["getInteractiveFormsInfo"] = function()
	{
		let res = {};
		let ext = Module["_GetInteractiveFormsInfo"](this.nativeFile);
		if (ext == 0)
			return res;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, ext, 4);
		if (lenArray == null)
		{
			Module["_free"](ext);
			return res;
		}

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
		{
			Module["_free"](ext);
			return res;
		}

		let buffer = new Uint8Array(Module["HEAP8"].buffer, ext + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);

		if (!reader.isValid())
		{
			Module["_free"](ext);
			return res;
		}

		let k = reader.readInt();
		if (k > 0)
			res["CO"] = [];
		for (let i = 0; i < k; ++i)
			res["CO"].push(reader.readInt());
		
		k = reader.readInt();
		if (k > 0)
			res["Parents"] = [];
		for (let i = 0; i < k; ++i)
		{
			let rec = {};
			rec["i"] = reader.readInt();
			let flags = reader.readInt();
			if (flags & (1 << 0))
				rec["name"] = reader.readString();
			if (flags & (1 << 1))
				rec["value"] = reader.readString();
			if (flags & (1 << 2))
				rec["defaultValue"] = reader.readString();
			if (flags & (1 << 3))
			{
				let n = reader.readInt();
				rec["curIdxs"] = [];
				for (let i = 0; i < n; ++i)
					rec["curIdxs"].push(reader.readInt());
			}
			if (flags & (1 << 4))
				rec["Parent"] = reader.readInt();
			if (flags & (1 << 5))
			{
				let n = reader.readInt();
				rec["value"] = [];
				for (let i = 0; i < n; ++i)
					rec["value"].push(reader.readString());
			}
			if (flags & (1 << 6))
			{
				let n = reader.readInt();
				rec["Opt"] = [];
				for (let i = 0; i < n; ++i)
					rec["Opt"].push(reader.readString());
			}
			res["Parents"].push(rec);
		}

		res["Fields"] = [];
		k = reader.readInt();
		for (let q = 0; reader.isValid() && q < k; ++q)
		{
			let rec = {};
			// Widget type - FT
			// 26 - Unknown, 27 - button, 28 - radiobutton, 29 - checkbox, 30 - text, 31 - combobox, 32 - listbox, 33 - signature
			rec["type"] = reader.readByte();
			// Annot
			readAnnot(reader, rec);
			// Widget
			rec["font"] = {};
			rec["font"]["name"] = reader.readString();
			rec["font"]["size"] = reader.readDouble();
			rec["font"]["style"] = reader.readInt();
			let tc = reader.readInt();
			if (tc)
			{
				rec["font"]["color"] = [];
				for (let i = 0; i < tc; ++i)
					rec["font"]["color"].push(reader.readDouble2());
			}
			// 0 - left-justified, 1 - centered, 2 - right-justified
			rec["alignment"] = reader.readByte();
			rec["flag"] = reader.readInt();
			// 12.7.3.1
			rec["readOnly"] = (rec["flag"] >> 0) & 1; // ReadOnly
			rec["required"] = (rec["flag"] >> 1) & 1; // Required
			rec["noexport"] = (rec["flag"] >> 2) & 1; // NoExport
			let flags = reader.readInt();
			// Alternative field name, used in tooltip and error messages - TU
			if (flags & (1 << 0))
				rec["userName"] = reader.readString();
			// Default style string (CSS2 format) - DS
			if (flags & (1 << 1))
				rec["defaultStyle"] = reader.readString();
			// Actual font
			if (flags & (1 << 2))
				rec["font"]["actual"] = reader.readString();
			// Selection mode - H
			// 0 - none, 1 - invert, 2 - push, 3 - outline
			if (flags & (1 << 3))
				rec["highlight"] = reader.readByte();
			// Font key
			if (flags & (1 << 4))
				rec["font"]["key"] = reader.readString();
			// Border color - BC. Even if the border is not specified by BS/Border, 
			// then if BC is present, a default border is provided (solid, thickness 1). 
			// If the text annotation has MaxLen, borders appear for each character
			if (flags & (1 << 5))
			{
				let n = reader.readInt();
				rec["BC"] = [];
				for (let i = 0; i < n; ++i)
					rec["BC"].push(reader.readDouble2());
			}
			// Rotate an annotation relative to the page - R
			if (flags & (1 << 6))
				rec["rotate"] = reader.readInt();
			// Annotation background color - BG
			if (flags & (1 << 7))
			{
				let n = reader.readInt();
				rec["BG"] = [];
				for (let i = 0; i < n; ++i)
					rec["BG"].push(reader.readDouble2());
			}
			// Default value - DV
			if (flags & (1 << 8))
				rec["defaultValue"] = reader.readString();
			if (flags & (1 << 17))
				rec["Parent"] = reader.readInt();
			if (flags & (1 << 18))
				rec["name"] = reader.readString();
			if (flags & (1 << 19))
				rec["font"]["AP"] = reader.readString();
			// Action
			let nAction = reader.readInt();
			if (nAction > 0)
				rec["AA"] = {};
			for (let i = 0; i < nAction; ++i)
			{
				let AAType = reader.readString();
				rec["AA"][AAType] = {};
				readAction(reader, rec["AA"][AAType]);
			}
			// Widget types
			if (rec["type"] == 27)
			{
				if (flags & (1 << 9))
					rec["value"] = reader.readString();
				let IFflags = reader.readInt();
				// Header - СA
				if (flags & (1 << 10))
					rec["caption"] = reader.readString();
				// Rollover header - RC
				if (flags & (1 << 11))
					rec["rolloverCaption"] = reader.readString();
				// Alternate header - AC
				if (flags & (1 << 12))
					rec["alternateCaption"] = reader.readString();
				// Header position - TP
				if (flags & (1 << 13))
					// 0 - textOnly, 1 - iconOnly, 2 - iconTextV, 3 - textIconV, 4 - iconTextH, 5 - textIconH, 6 - overlay
					rec["position"] = reader.readByte();
				// Icons - IF
				if (IFflags & (1 << 0))
				{
					rec["IF"] = {};
					// Scaling IF.SW
					// 0 - Always, 1 - Never, 2 - too big, 3 - too small
					if (IFflags & (1 << 1))
						rec["IF"]["SW"] = reader.readByte();
					// Scaling type - IF.S
					// 0 - Proportional, 1 - Anamorphic
					if (IFflags & (1 << 2))
						rec["IF"]["S"] = reader.readByte();
					if (IFflags & (1 << 3))
					{
						rec["IF"]["A"] = [];
						rec["IF"]["A"].push(reader.readDouble());
						rec["IF"]["A"].push(reader.readDouble());
					}
					rec["IF"]["FB"] = (IFflags >> 4) & 1;
				}
			}
			else if (rec["type"] == 29 || rec["type"] == 28)
			{
				if (flags & (1 << 9))
					rec["value"] = reader.readString();
				// 0 - check, 1 - cross, 2 - diamond, 3 - circle, 4 - star, 5 - square
				rec["style"] = reader.readByte();
			    if (flags & (1 << 14))
					rec["ExportValue"] = reader.readString();
				// 12.7.4.2.1
				rec["NoToggleToOff"]  = (rec["flag"] >> 14) & 1; // NoToggleToOff
				rec["radiosInUnison"] = (rec["flag"] >> 25) & 1; // RadiosInUnison
			}
			else if (rec["type"] == 30)
			{
				if (flags & (1 << 9))
					rec["value"] = reader.readString();
				if (flags & (1 << 10))
					rec["maxLen"] = reader.readInt();
				if (rec["flag"] & (1 << 25))
					rec["richValue"] = reader.readString();
				// 12.7.4.3
				rec["multiline"]       = (rec["flag"] >> 12) & 1; // Multiline
				rec["password"]        = (rec["flag"] >> 13) & 1; // Password
				rec["fileSelect"]      = (rec["flag"] >> 20) & 1; // FileSelect
				rec["doNotSpellCheck"] = (rec["flag"] >> 22) & 1; // DoNotSpellCheck
				rec["doNotScroll"]     = (rec["flag"] >> 23) & 1; // DoNotScroll
				rec["comb"]            = (rec["flag"] >> 24) & 1; // Comb
				rec["richText"]        = (rec["flag"] >> 25) & 1; // RichText
			}
			else if (rec["type"] == 31 || rec["type"] == 32)
			{
				if (flags & (1 << 9))
					rec["value"] = reader.readString();
				if (flags & (1 << 10))
				{
					let n = reader.readInt();
					rec["opt"] = [];
					for (let i = 0; i < n; ++i)
					{
						let opt1 = reader.readString();
						let opt2 = reader.readString();
						if (opt1 == "")
							rec["opt"].push(opt2);
						else
							rec["opt"].push([opt2, opt1]);
					}
				}
				if (flags & (1 << 11))
					rec["TI"] = reader.readInt();
				if (flags & (1 << 12))
				{
					let n = reader.readInt();
					rec["curIdxs"] = [];
					for (let i = 0; i < n; ++i)
						rec["curIdxs"].push(reader.readInt());
				}
				if (flags & (1 << 13))
				{
					let n = reader.readInt();
					rec["value"] = [];
					for (let i = 0; i < n; ++i)
						rec["value"].push(reader.readString());
				}
				// 12.7.4.4
				rec["editable"]          = (rec["flag"] >> 18) & 1; // Edit
				rec["multipleSelection"] = (rec["flag"] >> 21) & 1; // MultiSelect
				rec["doNotSpellCheck"]   = (rec["flag"] >> 22) & 1; // DoNotSpellCheck
				rec["commitOnSelChange"] = (rec["flag"] >> 26) & 1; // CommitOnSelChange
			}
			else if (rec["type"] == 33)
			{
				rec["Sig"] = (flags >> 9) & 1;
			}
			
			res["Fields"].push(rec);
		}

		Module["_free"](ext);
		return res;
	};
	CFile.prototype["getInteractiveFormsEmbeddedFonts"] = function()
	{
		return getWidgetFonts(this.nativeFile, 1);
	};
	CFile.prototype["getInteractiveFormsStandardFonts"] = function()
	{
		return getWidgetFonts(this.nativeFile, 2);
	};
	// optional nWidget     - rec["AP"]["i"]
	// optional sView       - N/D/R
	// optional sButtonView - state pushbutton-annotation - Off/Yes(or rec["ExportValue"])
	CFile.prototype["getInteractiveFormsAP"] = function(pageIndex, width, height, backgroundColor, nWidget, sView, sButtonView)
	{
		let nView = -1;
		if (sView)
		{
			if (sView == "N")
				nView = 0;
			else if (sView == "D")
				nView = 1;
			else if (sView == "R")
				nView = 2;
		}
		let nButtonView = -1;
		if (sButtonView)
			nButtonView = (sButtonView == "Off" ? 0 : 1);

		let res = [];
		this.lockPageNumForFontsLoader(pageIndex, UpdateFontsSource.Forms);
		let ext = Module["_GetInteractiveFormsAP"](this.nativeFile, width, height, backgroundColor === undefined ? 0xFFFFFF : backgroundColor, pageIndex, nWidget === undefined ? -1 : nWidget, nView, nButtonView);
		this.unlockPageNumForFontsLoader();
		if (ext == 0)
			return res;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, ext, 4);
		if (lenArray == null)
			return res;

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
			return res;

		let buffer = new Uint8Array(Module["HEAP8"].buffer, ext + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);

		while (reader.isValid())
		{
			// Annotation view
			let AP = {};
			readAnnotAP(reader, AP);
			res.push(AP);
		}

		Module["_free"](ext);
		return res;
	};
	// optional bBase64   - true/false base64 result
	// optional nWidget ...
	// optional sIconView - icon - I/RI/IX
	CFile.prototype["getButtonIcons"] = function(pageIndex, width, height, backgroundColor, bBase64, nWidget, sIconView)
	{
		let nView = -1;
		if (sIconView)
		{
			if (sIconView == "I")
				nView = 0;
			else if (sIconView == "RI")
				nView = 1;
			else if (sIconView == "IX")
				nView = 2;
		}

		let res = {};
		let ext = Module["_GetButtonIcons"](this.nativeFile, backgroundColor === undefined ? 0xFFFFFF : backgroundColor, pageIndex, bBase64 ? 1 : 0, nWidget === undefined ? -1 : nWidget, nView);
		if (ext == 0)
			return res;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, ext, 4);
		if (lenArray == null)
			return res;

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
			return res;

		let buffer = new Uint8Array(Module["HEAP8"].buffer, ext + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);
		
		res["MK"] = [];
		res["View"] = [];

		while (reader.isValid())
		{
			// View pushbutton annotation
			let MK = {};
			// Relation with AP
			MK["i"] = reader.readInt();
			let n = reader.readInt();
			for (let i = 0; i < n; ++i)
			{
				let MKType = reader.readString();
				MK[MKType] = reader.readInt();
				let unique = reader.readByte();
				if (unique)
				{
					let ViewMK = {};
					ViewMK["j"] = MK[MKType];
					ViewMK["w"] = reader.readInt();
					ViewMK["h"] = reader.readInt();
					if (bBase64)
					{
						// base64 string with image
						ViewMK["retValue"] = reader.readString();
					}
					else
					{
						let np1 = reader.readInt();
						let np2 = reader.readInt();
						// this memory needs to be deleted
						ViewMK["retValue"] = np2 << 32 | np1;
					}
					res["View"].push(ViewMK);
				}
			}
			res["MK"].push(MK);
		}

		Module["_free"](ext);
		return res;
	};
	// optional pageIndex - get annotations from specific page
	CFile.prototype["getAnnotationsInfo"] = function(pageIndex)
	{
		let res = [];
		let ext = Module["_GetAnnotationsInfo"](this.nativeFile, pageIndex === undefined ? -1 : pageIndex);
		if (ext == 0)
			return res;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, ext, 4);
		if (lenArray == null)
		{
			Module["_free"](ext);
			return res;
		}

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
		{
			Module["_free"](ext);
			return res;
		}

		let buffer = new Uint8Array(Module["HEAP8"].buffer, ext + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);

		if (!reader.isValid())
		{
			Module["_free"](ext);
			return res;
		}
		
		while (reader.isValid())
		{
			let rec = {};
			// Annotation type
			// 0 - Text, 1 - Link, 2 - FreeText, 3 - Line, 4 - Square, 5 - Circle,
			// 6 - Polygon, 7 - PolyLine, 8 - Highlight, 9 - Underline, 10 - Squiggly, 
			// 11 - Strikeout, 12 - Stamp, 13 - Caret, 14 - Ink, 15 - Popup, 16 - FileAttachment, 
			// 17 - Sound, 18 - Movie, 19 - Widget, 20 - Screen, 21 - PrinterMark,
			// 22 - TrapNet, 23 - Watermark, 24 - 3D, 25 - Redact
			rec["Type"] = reader.readByte();
			// Annot
			readAnnot(reader, rec);
			// Markup
			let flags = 0;
			if ((rec["Type"] < 18 && rec["Type"] != 1 && rec["Type"] != 15) || rec["Type"] == 25)
			{
				flags = reader.readInt();
				if (flags & (1 << 0))
					rec["Popup"] = reader.readInt();
				// T
				if (flags & (1 << 1))
					rec["User"] = reader.readString();
				// CA
				if (flags & (1 << 2))
					rec["CA"] = reader.readDouble();
				// RC
				if (flags & (1 << 3))
				{
					let n = reader.readInt();
					rec["RC"] = [];
					for (let i = 0; i < n; ++i)
					{
						let oFont = {};
						// 0 - left, 1 - centered, 2 - right, 3 - justify
						oFont["alignment"] = reader.readByte();
						let nFontFlag = reader.readInt();
						oFont["bold"] = (nFontFlag >> 0) & 1;
						oFont["italic"] = (nFontFlag >> 1) & 1;
						oFont["strikethrough"] = (nFontFlag >> 3) & 1;
						oFont["underlined"] = (nFontFlag >> 4) & 1;
						if (nFontFlag & (1 << 5))
							oFont["vertical"] = reader.readDouble();
						if (nFontFlag & (1 << 6))
							oFont["actual"] = reader.readString();
						oFont["size"] = reader.readDouble();
						oFont["color"] = [];
						oFont["color"].push(reader.readDouble2());
						oFont["color"].push(reader.readDouble2());
						oFont["color"].push(reader.readDouble2());
						oFont["name"] = reader.readString();
						oFont["text"] = reader.readString();
						rec["RC"].push(oFont);
					}
				}
				// CreationDate
				if (flags & (1 << 4))
					rec["CreationDate"] = reader.readString();
				// IRT
				if (flags & (1 << 5))
					rec["RefTo"] = reader.readInt();
				// RT
				// 0 - R, 1 - Group
				if (flags & (1 << 6))
					rec["RefToReason"] = reader.readByte();
				// Subj
				if (flags & (1 << 7))
					rec["Subj"] = reader.readString();
			}
			// Text
			if (rec["Type"] == 0)
			{
				// Bachground color - C->IC
				if (rec["C"])
				{
					rec["IC"] = rec["C"];
					delete rec["C"];
				}
				rec["Open"] = (flags >> 15) & 1;
				// icon - Name
				// 0 - Check, 1 - Checkmark, 2 - Circle, 3 - Comment, 4 - Cross, 5 - CrossHairs, 6 - Help, 7 - Insert, 8 - Key, 9 - NewParagraph, 10 - Note, 11 - Paragraph, 12 - RightArrow, 13 - RightPointer, 14 - Star, 15 - UpArrow, 16 - UpLeftArrow
				if (flags & (1 << 16))
					rec["Icon"] = reader.readByte();
				// StateModel
				// 0 - Marked, 1 - Review
				if (flags & (1 << 17))
					rec["StateModel"] = reader.readByte();
				// State
				// 0 - Marked, 1 - Unmarked, 2 - Accepted, 3 - Rejected, 4 - Cancelled, 5 - Completed, 6 - None
				if (flags & (1 << 18))
					rec["State"] = reader.readByte();
				
			}
			// Line
			else if (rec["Type"] == 3)
			{
				// L
				rec["L"] = [];
				for (let i = 0; i < 4; ++i)
					rec["L"].push(reader.readDouble());
				// LE
				// 0 - Square, 1 - Circle, 2 - Diamond, 3 - OpenArrow, 4 - ClosedArrow, 5 - None, 6 - Butt, 7 - ROpenArrow, 8 - RClosedArrow, 9 - Slash
				if (flags & (1 << 15))
				{
					rec["LE"] = [];
					rec["LE"].push(reader.readByte());
					rec["LE"].push(reader.readByte());
				}
				// IC
				if (flags & (1 << 16))
				{
					let n = reader.readInt();
					rec["IC"] = [];
					for (let i = 0; i < n; ++i)
						rec["IC"].push(reader.readDouble2());
				}
				// LL
				if (flags & (1 << 17))
					rec["LL"] = reader.readDouble();
				// LLE
				if (flags & (1 << 18))
					rec["LLE"] = reader.readDouble();
				// Cap
				rec["Cap"] = (flags >> 19) & 1;
				// IT
				// 0 - LineDimension, 1 - LineArrow
				if (flags & (1 << 20))
					rec["IT"] = reader.readByte();
				// LLO
				if (flags & (1 << 21))
					rec["LLO"] = reader.readDouble();
				// CP
				// 0 - Inline, 1 - Top
				if (flags & (1 << 22))
					rec["CP"] = reader.readByte();
				// CO
				if (flags & (1 << 23))
				{
					rec["CO"] = [];
					rec["CO"].push(reader.readDouble());
					rec["CO"].push(reader.readDouble());
				}
			}
			// Ink
			else if (rec["Type"] == 14)
			{
				// offsets like getStructure and viewer.navigate
				let n = reader.readInt();
				rec["InkList"] = [];
				for (let i = 0; i < n; ++i)
				{
					rec["InkList"][i] = [];
					let m = reader.readInt();
					for (let j = 0; j < m; ++j)
						rec["InkList"][i].push(reader.readDouble());
				}
			}
			// Highlight, Underline, Squiggly, Strikeout
			else if (rec["Type"] > 7 && rec["Type"] < 12)
			{
				// QuadPoints
				let n = reader.readInt();
				rec["QuadPoints"] = [];
				for (let i = 0; i < n; ++i)
					rec["QuadPoints"].push(reader.readDouble());
			}
			// Square, Circle
			else if (rec["Type"] == 4 || rec["Type"] == 5)
			{
				// Rect and RD differences
				if (flags & (1 << 15))
				{
					rec["RD"] = [];
					for (let i = 0; i < 4; ++i)
						rec["RD"].push(reader.readDouble());
				}
				// IC
				if (flags & (1 << 16))
				{
					let n = reader.readInt();
					rec["IC"] = [];
					for (let i = 0; i < n; ++i)
						rec["IC"].push(reader.readDouble2());
				}
			}
			// Polygon, PolyLine
			else if (rec["Type"] == 6 || rec["Type"] == 7)
			{
				let nVertices = reader.readInt();
				rec["Vertices"] = [];
				for (let i = 0; i < nVertices; ++i)
					rec["Vertices"].push(reader.readDouble());
				// LE
				// 0 - Square, 1 - Circle, 2 - Diamond, 3 - OpenArrow, 4 - ClosedArrow, 5 - None, 6 - Butt, 7 - ROpenArrow, 8 - RClosedArrow, 9 - Slash
				if (flags & (1 << 15))
				{
					rec["LE"] = [];
					rec["LE"].push(reader.readByte());
					rec["LE"].push(reader.readByte());
				}
				// IC
				if (flags & (1 << 16))
				{
					let n = reader.readInt();
					rec["IC"] = [];
					for (let i = 0; i < n; ++i)
						rec["IC"].push(reader.readDouble2());
				}
				// IT
				// 0 - PolygonCloud, 1 - PolyLineDimension, 2 - PolygonDimension
				if (flags & (1 << 20))
					rec["IT"] = reader.readByte();
			}
			// Popup
			/*
			else if (rec["Type"] == 15)
			{
				flags = reader.readInt();
				rec["Open"] = (flags >> 0) & 1;
				// Link to parent-annotation
				if (flags & (1 << 1))
					rec["PopupParent"] = reader.readInt();
			}
			*/
			// FreeText
			else if (rec["Type"] == 2)
			{
				// Bachground color - C->IC
				if (rec["C"])
				{
					rec["IC"] = rec["C"];
					delete rec["C"];
				}
				// 0 - left-justified, 1 - centered, 2 - right-justified
				rec["alignment"] = reader.readByte();
				rec["Rotate"] = reader.readInt();
				// Rect and RD differences
				if (flags & (1 << 15))
				{
					rec["RD"] = [];
					for (let i = 0; i < 4; ++i)
						rec["RD"].push(reader.readDouble());
				}
				// CL
				if (flags & (1 << 16))
				{
					let n = reader.readInt();
					rec["CL"] = [];
					for (let i = 0; i < n; ++i)
						rec["CL"].push(reader.readDouble());
				}
				// Default style (CSS2 format) - DS
				if (flags & (1 << 17))
					rec["defaultStyle"] = reader.readString();
				// LE
				// 0 - Square, 1 - Circle, 2 - Diamond, 3 - OpenArrow, 4 - ClosedArrow, 5 - None, 6 - Butt, 7 - ROpenArrow, 8 - RClosedArrow, 9 - Slash
				if (flags & (1 << 18))
					rec["LE"] = reader.readByte();
				// IT
				// 0 - FreeText, 1 - FreeTextCallout, 2 - FreeTextTypeWriter
				if (flags & (1 << 20))
					rec["IT"] = reader.readByte();
				// Border color - from DA (write to C)
				if (flags & (1 << 21))
				{
					let n = reader.readInt();
					rec["C"] = [];
					for (let i = 0; i < n; ++i)
						rec["C"].push(reader.readDouble2());
				}
			}
			// Caret
			else if (rec["Type"] == 13)
			{
				// Rect and RD differenses
				if (flags & (1 << 15))
				{
					rec["RD"] = [];
					for (let i = 0; i < 4; ++i)
						rec["RD"].push(reader.readDouble());
				}
				// Sy
				// 0 - None, 1 - P, 2 - S
				if (flags & (1 << 16))
					rec["Sy"] = reader.readByte();
			}
			// FileAttachment
			else if (rec["Type"] == 16)
			{
				if (flags & (1 << 15))
					rec["Icon"] = reader.readString();
				if (flags & (1 << 16))
					rec["FS"] = reader.readString();
				if (flags & (1 << 17))
				{
					rec["F"] = {};
					rec["F"]["FileName"] = reader.readString();
				}
				if (flags & (1 << 18))
				{
					rec["UF"] = {};
					rec["UF"]["FileName"] = reader.readString();
				}
				if (flags & (1 << 19))
				{
					rec["DOS"] = {};
					rec["DOS"]["FileName"] = reader.readString();
				}
				if (flags & (1 << 20))
				{
					rec["Mac"] = {};
					rec["Mac"]["FileName"] = reader.readString();
				}
				if (flags & (1 << 21))
				{
					rec["Unix"] = {};
					rec["Unix"]["FileName"] = reader.readString();
				}
				if (flags & (1 << 22))
				{
					rec["ID"] = [];
					rec["ID"].push(reader.readString());
					rec["ID"].push(reader.readString());
				}
				rec["V"] = flags & (1 << 23);
				if (flags & (1 << 24))
				{
					let flag = reader.readInt();
					if (flag & (1 << 0))
					{
						let n = reader.readInt();
						let np1 = reader.readInt();
						let np2 = reader.readInt();
						let pPoint = np2 << 32 | np1;
						rec["F"]["File"] = new Uint8Array(Module["HEAP8"].buffer, pPoint, n);
						Module["_free"](pPoint);
					}
					if (flag & (1 << 1))
					{
						let n = reader.readInt();
						let np1 = reader.readInt();
						let np2 = reader.readInt();
						let pPoint = np2 << 32 | np1;
						rec["UF"]["File"] = new Uint8Array(Module["HEAP8"].buffer, pPoint, n);
						Module["_free"](pPoint);
					}
					if (flag & (1 << 2))
					{
						let n = reader.readInt();
						let np1 = reader.readInt();
						let np2 = reader.readInt();
						let pPoint = np2 << 32 | np1;
						rec["DOS"]["File"] = new Uint8Array(Module["HEAP8"].buffer, pPoint, n);
						Module["_free"](pPoint);
					}
					if (flag & (1 << 3))
					{
						let n = reader.readInt();
						let np1 = reader.readInt();
						let np2 = reader.readInt();
						let pPoint = np2 << 32 | np1;
						rec["Mac"]["File"] = new Uint8Array(Module["HEAP8"].buffer, pPoint, n);
						Module["_free"](pPoint);
					}
					if (flag & (1 << 4))
					{
						let n = reader.readInt();
						let np1 = reader.readInt();
						let np2 = reader.readInt();
						let pPoint = np2 << 32 | np1;
						rec["Unix"]["File"] = new Uint8Array(Module["HEAP8"].buffer, pPoint, n);
						Module["_free"](pPoint);
					}
				}
				if (flags & (1 << 26))
					rec["Desc"] = reader.readString();
			}
			res.push(rec);
		}
		
		Module["_free"](ext);
		return res;
	};
	// optional nAnnot ...
	// optional sView ...
	CFile.prototype["getAnnotationsAP"] = function(pageIndex, width, height, backgroundColor, nAnnot, sView)
	{
		let nView = -1;
		if (sView)
		{
			if (sView == "N")
				nView = 0;
			else if (sView == "D")
				nView = 1;
			else if (sView == "R")
				nView = 2;
		}

		let res = [];
		this.lockPageNumForFontsLoader(pageIndex, UpdateFontsSource.Annotation);
		let ext = Module["_GetAnnotationsAP"](this.nativeFile, width, height, backgroundColor === undefined ? 0xFFFFFF : backgroundColor, pageIndex, nAnnot === undefined ? -1 : nAnnot, nView);
		this.unlockPageNumForFontsLoader();
		if (ext == 0)
			return res;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, ext, 4);
		if (lenArray == null)
			return res;

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
			return res;

		let buffer = new Uint8Array(Module["HEAP8"].buffer, ext + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);
		
		while (reader.isValid())
		{
			// Annotation view
			let AP = {};
			readAnnotAP(reader, AP);
			res.push(AP);
		}

		Module["_free"](ext);
		return res;
	};
	CFile.prototype["getStructure"] = function()
	{
		let res = [];
		let str = Module["_GetStructure"](this.nativeFile);
		if (str == 0)
			return res;
		let lenArray = new Int32Array(Module["HEAP8"].buffer, str, 4);
		if (lenArray == null)
			return res;
		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
			return res;

		let buffer = new Uint8Array(Module["HEAP8"].buffer, str + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);

		while (reader.isValid())
		{
			let rec = {};
			rec["page"]  = reader.readInt();
			rec["level"] = reader.readInt();
			rec["y"]  = reader.readDouble();
			rec["description"] = reader.readString();
			res.push(rec);
		}

		Module["_free"](str);
		return res;
	};
	CFile.prototype["getFontByID"] = function(ID)
	{
		let res = null;
		if (ID === undefined)
			return res;
		
		let idBuffer = ID.toUtf8();
		let idPointer = Module["_malloc"](idBuffer.length);
		Module["HEAP8"].set(idBuffer, idPointer);
		let ext = Module["_GetFontBinary"](this.nativeFile, idPointer);
		Module["_free"](idPointer);
		if (ext == 0)
			return res;
		
		let lenArray = new Int32Array(Module["HEAP8"].buffer, ext, 4);
		if (lenArray == null)
		{
			Module["_free"](ext);
			return res;
		}

		let len = lenArray[0];
		len -= 4;
		if (len <= 0)
		{
			Module["_free"](ext);
			return res;
		}

		let buffer = new Uint8Array(Module["HEAP8"].buffer, ext + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);
		
		while (reader.isValid())
		{
			let nFontLength = reader.readInt();
			let np1 = reader.readInt();
			let np2 = reader.readInt();
			let pFontPoint = np2 << 32 | np1;
			
			res = new Uint8Array(Module["HEAP8"].buffer, pFontPoint, nFontLength);
		}

		Module["_free"](ext);
		
		return res;
	};

	CFile.prototype["scanPage"] = function(page, mode)
	{
		let data = Module["_ScanPage"](this.nativeFile, page, (mode === undefined) ? 0 : mode);
		if (data == 0)
			return [];

		let lenArray = new Int32Array(Module["HEAP8"].buffer, data, 4);
		if (lenArray == null)
			return [];

		let len = lenArray[0];
		if (0 == len)
			return [];

		let buffer = new Uint8Array(Module["HEAP8"].buffer, data + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);

		let shapesCount = reader.readInt();
		let shapes = new Array(shapesCount);

		for (let i = 0; i < shapesCount; i++)
			shapes[i] = reader.readString();

		Module["_free"](data);
		return shapes;
	};

	CFile.prototype["getImageBase64"] = function(rId)
	{
		let strId = "" + rId;
		if (this.scannedImages[strId])
			return this.scannedImages[strId];

		let strPtr = Module["_GetImageBase64"](this.nativeFile, rId);
		if (0 == strPtr)
		{
			this.scannedImages[strId] = "error";
			return this.scannedImages[strId];
		}

		let len = Module["_GetImageBase64Len"](strPtr);
		let ptr = Module["_GetImageBase64Ptr"](strPtr);

		var buffer = new Uint8Array(Module["HEAP8"].buffer, ptr, len);
		this.scannedImages[strId] = String.prototype.fromUtf8(buffer, 0, len);
		Module["_GetImageBase64Free"](strPtr);
		return this.scannedImages[strId];
	};

	CFile.prototype["changeImageUrl"] = function(baseUrl, resultUrl)
	{
		for (let i in this.scannedImages)
		{
			if (this.scannedImages[i] == baseUrl)
				this.scannedImages[i] = resultUrl;
		}
	};

	CFile.prototype.memory = function()
	{
		return Module["HEAP8"];
	};
	CFile.prototype.free = function(pointer)
	{
		Module["_free"](pointer);
	};

	CFile.prototype.lockPageNumForFontsLoader = function(pageIndex, type)
	{
		this.fontPageIndex = pageIndex;
		this.fontPageUpdateType = type;		
	};
	CFile.prototype.unlockPageNumForFontsLoader = function()
	{
		this.fontPageIndex = -1;
		drawingFile.fontPageUpdateType = UpdateFontsSource.Undefined;
	};
	
	self["AscViewer"]["CDrawingFile"] = CFile;
	self["AscViewer"]["InitializeFonts"] = function(basePath) {
		if (undefined !== basePath && "" !== basePath)
			baseFontsPath = basePath;
		if (!window["g_fonts_selection_bin"])
			return;
		let memoryBuffer = window["g_fonts_selection_bin"].toUtf8();
		let pointer = Module["_malloc"](memoryBuffer.length);
		Module.HEAP8.set(memoryBuffer, pointer);
		Module["_InitializeFontsBase64"](pointer, memoryBuffer.length);
		Module["_free"](pointer);
		delete window["g_fonts_selection_bin"];

		// ranges
		let rangesBuffer = new CBinaryWriter();
		let ranges = AscFonts.getSymbolRanges();

		let rangesCount = ranges.length;
		rangesBuffer.writeUint(rangesCount);
		for (let i = 0; i < rangesCount; i++)
		{
			rangesBuffer.writeString(ranges[i].getName());
			rangesBuffer.writeUint(ranges[i].getStart());
			rangesBuffer.writeUint(ranges[i].getEnd());
		}

		let rangesFinalLen = rangesBuffer.dataSize;
		let rangesFinal = new Uint8Array(rangesBuffer.buffer.buffer, 0, rangesFinalLen);
		pointer = Module["_malloc"](rangesFinalLen);
		Module.HEAP8.set(rangesFinal, pointer);
		Module["_InitializeFontsRanges"](pointer, rangesFinalLen);
		Module["_free"](pointer);
	};
	self["AscViewer"]["Free"] = function(pointer) {
		Module["_free"](pointer);
	};
	
	function addToArrayAsDictionary(arr, value)
	{
		let isFound = false;
		for (let i = 0, len = arr.length; i < len; i++)
		{
			if (arr[i] == value)
			{
				isFound = true;
				break;
			}
		}
		if (!isFound)
			arr.push(value);
		return isFound;
	}

	self["AscViewer"]["CheckStreamId"] = function(data, status) {
		let drawingFile = self.drawingFile;
		if (!drawingFile)
			return;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, data, 4);
		let len = lenArray[0];
		len -= 4;

		let buffer = new Uint8Array(Module["HEAP8"].buffer, data + 4, len);
		let reader = new CBinaryReader(buffer, 0, len);

		let name = reader.readString();
		let style = 0;
		if (reader.readInt() != 0)
			style |= 1;//AscFonts.FontStyle.FontStyleBold;
		if (reader.readInt() != 0)
			style |= 2;//AscFonts.FontStyle.FontStyleItalic;

		let file = AscFonts.pickFont(name, style);
		let fileId = file.GetID();
		let fileStatus = file.GetStatus();

		if (fileStatus === 0)
		{
			// font was loaded
			fontToMemory(file, true);
		}
		else
		{
			drawingFile.fontStreams[fileId] = drawingFile.fontStreams[fileId] || {};
			drawingFile.fontStreams[fileId].pages = drawingFile.fontStreams[fileId].pages || [];
			addToArrayAsDictionary(drawingFile.fontStreams[fileId].pages, drawingFile.fontPageIndex);
			addToArrayAsDictionary(drawingFile.pages[drawingFile.fontPageIndex].fonts, fileId);

			drawingFile.pages[drawingFile.fontPageIndex].fontsUpdateType |= drawingFile.fontPageUpdateType;

			// font can be loading in editor
			if (undefined === file.externalCallback)
			{
				let _t = file;
				file.externalCallback = function() {
					fontToMemory(_t, true);

					let pages = drawingFile.fontStreams[fileId].pages;
					delete drawingFile.fontStreams[fileId];

					let pagesRepaint_Page        = [];
					let pagesRepaint_Annotation  = [];
					let pagesRepaint_Forms       = [];

					for (let i = 0, len = pages.length; i < len; i++)
					{
						let pageNum = pages[i];
						let pageObj = drawingFile.pages[pageNum];
						let fonts = pageObj.fonts;
						
						for (let j = 0, len_fonts = fonts.length; j < len_fonts; j++)
						{
							if (fonts[j] == fileId)
							{
								fonts.splice(j, 1);
								break;
							}
						}

						if (0 == fonts.length)
						{
							if (pageObj.fontsUpdateType & UpdateFontsSource.Page)
								pagesRepaint_Page.push(pageNum);

							if (pageObj.fontsUpdateType & UpdateFontsSource.Annotation)
								pagesRepaint_Annotation.push(pageNum);

							if (pageObj.fontsUpdateType & UpdateFontsSource.Forms)
								pagesRepaint_Forms.push(pageNum);

							pageObj.fontsUpdateType = UpdateFontsSource.Undefined;
						}
					}

					if (pagesRepaint_Page.length > 0 && drawingFile.onRepaintPages)
						drawingFile.onRepaintPages(pagesRepaint_Page);

					if (pagesRepaint_Annotation.length > 0 && drawingFile.onRepaintAnnotations)
						drawingFile.onRepaintAnnotations(pagesRepaint_Annotation);

					if (pagesRepaint_Forms.length > 0 && drawingFile.onRepaintForms)
						drawingFile.onRepaintForms(pagesRepaint_Forms);

					delete _t.externalCallback;
				};

				if (2 !== file.LoadFontAsync)
					file.LoadFontAsync(baseFontsPath, null);
			}
		}

		let memoryBuffer = fileId.toUtf8();
		let pointer = Module["_malloc"](memoryBuffer.length);
		Module.HEAP8.set(memoryBuffer, pointer);
		Module["HEAP8"][status] = (fileStatus == 0) ? 1 : 0;
		return pointer;
	};

	function fontToMemory(file, isCheck)
	{
		let idBuffer = file.GetID().toUtf8();
		let idPointer = Module["_malloc"](idBuffer.length);
		Module["HEAP8"].set(idBuffer, idPointer);

		if (isCheck)
		{
			let nExist = Module["_IsFontBinaryExist"](idPointer);
			if (nExist != 0)
			{
				Module["_free"](idPointer);
				return;
			}
		}

		let stream_index = file.GetStreamIndex();
		
		let stream = AscFonts.getFontStream(stream_index);
		let streamPointer = Module["_malloc"](stream.size);
		Module["HEAP8"].set(stream.data, streamPointer);

		Module["_SetFontBinary"](idPointer, streamPointer, stream.size);

		Module["_free"](streamPointer);
		Module["_free"](idPointer);
	}
})(window, undefined);
