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

var AscFonts = window['AscFonts'];

if (window["NATIVE_EDITOR_ENJINE"])
	window.setImmediate = function(fn) { fn(); };

var setImmediate = window.setImmediate;

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
	getBinaryPromise = function() { return getBinaryPromise2(); }
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


var Module=typeof Module!="undefined"?Module:{};var moduleOverrides=Object.assign({},Module);var arguments_=[];var thisProgram="./this.program";var quit_=(status,toThrow)=>{throw toThrow};var ENVIRONMENT_IS_WEB=true;var ENVIRONMENT_IS_WORKER=false;var scriptDirectory="";function locateFile(path){if(Module["locateFile"]){return Module["locateFile"](path,scriptDirectory)}return scriptDirectory+path}var read_,readAsync,readBinary,setWindowTitle;if(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER){if(ENVIRONMENT_IS_WORKER){scriptDirectory=self.location.href}else if(typeof document!="undefined"&&document.currentScript){scriptDirectory=document.currentScript.src}if(scriptDirectory.indexOf("blob:")!==0){scriptDirectory=scriptDirectory.substr(0,scriptDirectory.replace(/[?#].*/,"").lastIndexOf("/")+1)}else{scriptDirectory=""}{read_=(url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.send(null);return xhr.responseText});if(ENVIRONMENT_IS_WORKER){readBinary=(url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.responseType="arraybuffer";xhr.send(null);return new Uint8Array(xhr.response)})}readAsync=((url,onload,onerror)=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,true);xhr.responseType="arraybuffer";xhr.onload=(()=>{if(xhr.status==200||xhr.status==0&&xhr.response){onload(xhr.response);return}onerror()});xhr.onerror=onerror;xhr.send(null)})}setWindowTitle=(title=>document.title=title)}else{}var out=Module["print"]||console.log.bind(console);var err=Module["printErr"]||console.warn.bind(console);Object.assign(Module,moduleOverrides);moduleOverrides=null;if(Module["arguments"])arguments_=Module["arguments"];if(Module["thisProgram"])thisProgram=Module["thisProgram"];if(Module["quit"])quit_=Module["quit"];var tempRet0=0;var setTempRet0=value=>{tempRet0=value};var getTempRet0=()=>tempRet0;var wasmBinary;if(Module["wasmBinary"])wasmBinary=Module["wasmBinary"];var noExitRuntime=Module["noExitRuntime"]||true;if(typeof WebAssembly!="object"){abort("no native wasm support detected")}var wasmMemory;var ABORT=false;var EXITSTATUS;var UTF8Decoder=typeof TextDecoder!="undefined"?new TextDecoder("utf8"):undefined;function UTF8ArrayToString(heapOrArray,idx,maxBytesToRead){var endIdx=idx+maxBytesToRead;var endPtr=idx;while(heapOrArray[endPtr]&&!(endPtr>=endIdx))++endPtr;if(endPtr-idx>16&&heapOrArray.buffer&&UTF8Decoder){return UTF8Decoder.decode(heapOrArray.subarray(idx,endPtr))}else{var str="";while(idx<endPtr){var u0=heapOrArray[idx++];if(!(u0&128)){str+=String.fromCharCode(u0);continue}var u1=heapOrArray[idx++]&63;if((u0&224)==192){str+=String.fromCharCode((u0&31)<<6|u1);continue}var u2=heapOrArray[idx++]&63;if((u0&240)==224){u0=(u0&15)<<12|u1<<6|u2}else{u0=(u0&7)<<18|u1<<12|u2<<6|heapOrArray[idx++]&63}if(u0<65536){str+=String.fromCharCode(u0)}else{var ch=u0-65536;str+=String.fromCharCode(55296|ch>>10,56320|ch&1023)}}}return str}function UTF8ToString(ptr,maxBytesToRead){return ptr?UTF8ArrayToString(HEAPU8,ptr,maxBytesToRead):""}function stringToUTF8Array(str,heap,outIdx,maxBytesToWrite){if(!(maxBytesToWrite>0))return 0;var startIdx=outIdx;var endIdx=outIdx+maxBytesToWrite-1;for(var i=0;i<str.length;++i){var u=str.charCodeAt(i);if(u>=55296&&u<=57343){var u1=str.charCodeAt(++i);u=65536+((u&1023)<<10)|u1&1023}if(u<=127){if(outIdx>=endIdx)break;heap[outIdx++]=u}else if(u<=2047){if(outIdx+1>=endIdx)break;heap[outIdx++]=192|u>>6;heap[outIdx++]=128|u&63}else if(u<=65535){if(outIdx+2>=endIdx)break;heap[outIdx++]=224|u>>12;heap[outIdx++]=128|u>>6&63;heap[outIdx++]=128|u&63}else{if(outIdx+3>=endIdx)break;heap[outIdx++]=240|u>>18;heap[outIdx++]=128|u>>12&63;heap[outIdx++]=128|u>>6&63;heap[outIdx++]=128|u&63}}heap[outIdx]=0;return outIdx-startIdx}function lengthBytesUTF8(str){var len=0;for(var i=0;i<str.length;++i){var u=str.charCodeAt(i);if(u>=55296&&u<=57343)u=65536+((u&1023)<<10)|str.charCodeAt(++i)&1023;if(u<=127)++len;else if(u<=2047)len+=2;else if(u<=65535)len+=3;else len+=4}return len}function writeArrayToMemory(array,buffer){HEAP8.set(array,buffer)}function writeAsciiToMemory(str,buffer,dontAddNull){for(var i=0;i<str.length;++i){HEAP8[buffer++>>0]=str.charCodeAt(i)}if(!dontAddNull)HEAP8[buffer>>0]=0}var buffer,HEAP8,HEAPU8,HEAP16,HEAPU16,HEAP32,HEAPU32,HEAPF32,HEAPF64;function updateGlobalBufferAndViews(buf){buffer=buf;Module["HEAP8"]=HEAP8=new Int8Array(buf);Module["HEAP16"]=HEAP16=new Int16Array(buf);Module["HEAP32"]=HEAP32=new Int32Array(buf);Module["HEAPU8"]=HEAPU8=new Uint8Array(buf);Module["HEAPU16"]=HEAPU16=new Uint16Array(buf);Module["HEAPU32"]=HEAPU32=new Uint32Array(buf);Module["HEAPF32"]=HEAPF32=new Float32Array(buf);Module["HEAPF64"]=HEAPF64=new Float64Array(buf)}var INITIAL_MEMORY=Module["INITIAL_MEMORY"]||16777216;var wasmTable;var __ATPRERUN__=[];var __ATINIT__=[];var __ATPOSTRUN__=[function(){window["AscFonts"].onLoadModule();}];var runtimeInitialized=false;function keepRuntimeAlive(){return noExitRuntime}function preRun(){if(Module["preRun"]){if(typeof Module["preRun"]=="function")Module["preRun"]=[Module["preRun"]];while(Module["preRun"].length){addOnPreRun(Module["preRun"].shift())}}callRuntimeCallbacks(__ATPRERUN__)}function initRuntime(){runtimeInitialized=true;callRuntimeCallbacks(__ATINIT__)}function postRun(){if(Module["postRun"]){if(typeof Module["postRun"]=="function")Module["postRun"]=[Module["postRun"]];while(Module["postRun"].length){addOnPostRun(Module["postRun"].shift())}}callRuntimeCallbacks(__ATPOSTRUN__)}function addOnPreRun(cb){__ATPRERUN__.unshift(cb)}function addOnInit(cb){__ATINIT__.unshift(cb)}function addOnPostRun(cb){__ATPOSTRUN__.unshift(cb)}var runDependencies=0;var runDependencyWatcher=null;var dependenciesFulfilled=null;function addRunDependency(id){runDependencies++;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}}function removeRunDependency(id){runDependencies--;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}if(runDependencies==0){if(runDependencyWatcher!==null){clearInterval(runDependencyWatcher);runDependencyWatcher=null}if(dependenciesFulfilled){var callback=dependenciesFulfilled;dependenciesFulfilled=null;callback()}}}Module["preloadedImages"]={};Module["preloadedAudios"]={};function abort(what){{if(Module["onAbort"]){Module["onAbort"](what)}}what="Aborted("+what+")";err(what);ABORT=true;EXITSTATUS=1;what+=". Build with -s ASSERTIONS=1 for more info.";var e=new WebAssembly.RuntimeError(what);throw e}var dataURIPrefix="data:application/octet-stream;base64,";function isDataURI(filename){return filename.startsWith(dataURIPrefix)}var wasmBinaryFile;wasmBinaryFile="fonts.wasm";if(!isDataURI(wasmBinaryFile)){wasmBinaryFile=locateFile(wasmBinaryFile)}function getBinary(file){try{if(file==wasmBinaryFile&&wasmBinary){return new Uint8Array(wasmBinary)}if(readBinary){return readBinary(file)}else{throw"both async and sync fetching of the wasm failed"}}catch(err){abort(err)}}function getBinaryPromise2(){if(!wasmBinary&&(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER)){if(typeof fetch=="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){if(!response["ok"]){throw"failed to load wasm binary file at '"+wasmBinaryFile+"'"}return response["arrayBuffer"]()}).catch(function(){return getBinary(wasmBinaryFile)})}}return Promise.resolve().then(function(){return getBinary(wasmBinaryFile)})}function createWasm(){var info={"a":asmLibraryArg};function receiveInstance(instance,module){var exports=instance.exports;Module["asm"]=exports;wasmMemory=Module["asm"]["za"];updateGlobalBufferAndViews(wasmMemory.buffer);wasmTable=Module["asm"]["Ba"];addOnInit(Module["asm"]["Aa"]);removeRunDependency("wasm-instantiate")}addRunDependency("wasm-instantiate");function receiveInstantiationResult(result){receiveInstance(result["instance"])}function instantiateArrayBuffer(receiver){return getBinaryPromise().then(function(binary){return WebAssembly.instantiate(binary,info)}).then(function(instance){return instance}).then(receiver,function(reason){err("failed to asynchronously prepare wasm: "+reason);abort(reason)})}function instantiateAsync(){if(!wasmBinary&&typeof WebAssembly.instantiateStreaming=="function"&&!isDataURI(wasmBinaryFile)&&typeof fetch=="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){var result=WebAssembly.instantiateStreaming(response,info);return result.then(receiveInstantiationResult,function(reason){err("wasm streaming compile failed: "+reason);err("falling back to ArrayBuffer instantiation");return instantiateArrayBuffer(receiveInstantiationResult)})})}else{return instantiateArrayBuffer(receiveInstantiationResult)}}if(Module["instantiateWasm"]){try{var exports=Module["instantiateWasm"](info,receiveInstance);return exports}catch(e){err("Module.instantiateWasm callback failed with error: "+e);return false}}instantiateAsync();return{}}function callRuntimeCallbacks(callbacks){while(callbacks.length>0){var callback=callbacks.shift();if(typeof callback=="function"){callback(Module);continue}var func=callback.func;if(typeof func=="number"){if(callback.arg===undefined){getWasmTableEntry(func)()}else{getWasmTableEntry(func)(callback.arg)}}else{func(callback.arg===undefined?null:callback.arg)}}}var wasmTableMirror=[];function getWasmTableEntry(funcPtr){var func=wasmTableMirror[funcPtr];if(!func){if(funcPtr>=wasmTableMirror.length)wasmTableMirror.length=funcPtr+1;wasmTableMirror[funcPtr]=func=wasmTable.get(funcPtr)}return func}function ___cxa_allocate_exception(size){return _malloc(size+16)+16}function ExceptionInfo(excPtr){this.excPtr=excPtr;this.ptr=excPtr-16;this.set_type=function(type){HEAP32[this.ptr+4>>2]=type};this.get_type=function(){return HEAP32[this.ptr+4>>2]};this.set_destructor=function(destructor){HEAP32[this.ptr+8>>2]=destructor};this.get_destructor=function(){return HEAP32[this.ptr+8>>2]};this.set_refcount=function(refcount){HEAP32[this.ptr>>2]=refcount};this.set_caught=function(caught){caught=caught?1:0;HEAP8[this.ptr+12>>0]=caught};this.get_caught=function(){return HEAP8[this.ptr+12>>0]!=0};this.set_rethrown=function(rethrown){rethrown=rethrown?1:0;HEAP8[this.ptr+13>>0]=rethrown};this.get_rethrown=function(){return HEAP8[this.ptr+13>>0]!=0};this.init=function(type,destructor){this.set_type(type);this.set_destructor(destructor);this.set_refcount(0);this.set_caught(false);this.set_rethrown(false)};this.add_ref=function(){var value=HEAP32[this.ptr>>2];HEAP32[this.ptr>>2]=value+1};this.release_ref=function(){var prev=HEAP32[this.ptr>>2];HEAP32[this.ptr>>2]=prev-1;return prev===1}}function CatchInfo(ptr){this.free=function(){_free(this.ptr);this.ptr=0};this.set_base_ptr=function(basePtr){HEAP32[this.ptr>>2]=basePtr};this.get_base_ptr=function(){return HEAP32[this.ptr>>2]};this.set_adjusted_ptr=function(adjustedPtr){HEAP32[this.ptr+4>>2]=adjustedPtr};this.get_adjusted_ptr_addr=function(){return this.ptr+4};this.get_adjusted_ptr=function(){return HEAP32[this.ptr+4>>2]};this.get_exception_ptr=function(){var isPointer=___cxa_is_pointer_type(this.get_exception_info().get_type());if(isPointer){return HEAP32[this.get_base_ptr()>>2]}var adjusted=this.get_adjusted_ptr();if(adjusted!==0)return adjusted;return this.get_base_ptr()};this.get_exception_info=function(){return new ExceptionInfo(this.get_base_ptr())};if(ptr===undefined){this.ptr=_malloc(8);this.set_adjusted_ptr(0)}else{this.ptr=ptr}}var exceptionCaught=[];function exception_addRef(info){info.add_ref()}var uncaughtExceptionCount=0;function ___cxa_begin_catch(ptr){var catchInfo=new CatchInfo(ptr);var info=catchInfo.get_exception_info();if(!info.get_caught()){info.set_caught(true);uncaughtExceptionCount--}info.set_rethrown(false);exceptionCaught.push(catchInfo);exception_addRef(info);return catchInfo.get_exception_ptr()}var exceptionLast=0;function ___cxa_free_exception(ptr){return _free(new ExceptionInfo(ptr).ptr)}function exception_decRef(info){if(info.release_ref()&&!info.get_rethrown()){var destructor=info.get_destructor();if(destructor){getWasmTableEntry(destructor)(info.excPtr)}___cxa_free_exception(info.excPtr)}}function ___cxa_end_catch(){_setThrew(0);var catchInfo=exceptionCaught.pop();exception_decRef(catchInfo.get_exception_info());catchInfo.free();exceptionLast=0}function ___resumeException(catchInfoPtr){var catchInfo=new CatchInfo(catchInfoPtr);var ptr=catchInfo.get_base_ptr();if(!exceptionLast){exceptionLast=ptr}catchInfo.free();throw ptr}function ___cxa_find_matching_catch_2(){var thrown=exceptionLast;if(!thrown){setTempRet0(0);return 0|0}var info=new ExceptionInfo(thrown);var thrownType=info.get_type();var catchInfo=new CatchInfo;catchInfo.set_base_ptr(thrown);catchInfo.set_adjusted_ptr(thrown);if(!thrownType){setTempRet0(0);return catchInfo.ptr|0}var typeArray=Array.prototype.slice.call(arguments);for(var i=0;i<typeArray.length;i++){var caughtType=typeArray[i];if(caughtType===0||caughtType===thrownType){break}if(___cxa_can_catch(caughtType,thrownType,catchInfo.get_adjusted_ptr_addr())){setTempRet0(caughtType);return catchInfo.ptr|0}}setTempRet0(thrownType);return catchInfo.ptr|0}function ___cxa_find_matching_catch_3(){var thrown=exceptionLast;if(!thrown){setTempRet0(0);return 0|0}var info=new ExceptionInfo(thrown);var thrownType=info.get_type();var catchInfo=new CatchInfo;catchInfo.set_base_ptr(thrown);catchInfo.set_adjusted_ptr(thrown);if(!thrownType){setTempRet0(0);return catchInfo.ptr|0}var typeArray=Array.prototype.slice.call(arguments);for(var i=0;i<typeArray.length;i++){var caughtType=typeArray[i];if(caughtType===0||caughtType===thrownType){break}if(___cxa_can_catch(caughtType,thrownType,catchInfo.get_adjusted_ptr_addr())){setTempRet0(caughtType);return catchInfo.ptr|0}}setTempRet0(thrownType);return catchInfo.ptr|0}function ___cxa_rethrow(){var catchInfo=exceptionCaught.pop();if(!catchInfo){abort("no exception to throw")}var info=catchInfo.get_exception_info();var ptr=catchInfo.get_base_ptr();if(!info.get_rethrown()){exceptionCaught.push(catchInfo);info.set_rethrown(true);info.set_caught(false);uncaughtExceptionCount++}else{catchInfo.free()}exceptionLast=ptr;throw ptr}function ___cxa_throw(ptr,type,destructor){var info=new ExceptionInfo(ptr);info.init(type,destructor);exceptionLast=ptr;uncaughtExceptionCount++;throw ptr}function ___cxa_uncaught_exceptions(){return uncaughtExceptionCount}var SYSCALLS={buffers:[null,[],[]],printChar:function(stream,curr){var buffer=SYSCALLS.buffers[stream];if(curr===0||curr===10){(stream===1?out:err)(UTF8ArrayToString(buffer,0));buffer.length=0}else{buffer.push(curr)}},varargs:undefined,get:function(){SYSCALLS.varargs+=4;var ret=HEAP32[SYSCALLS.varargs-4>>2];return ret},getStr:function(ptr){var ret=UTF8ToString(ptr);return ret},get64:function(low,high){return low}};function ___syscall_fcntl64(fd,cmd,varargs){SYSCALLS.varargs=varargs;return 0}function ___syscall_ioctl(fd,op,varargs){SYSCALLS.varargs=varargs;return 0}function ___syscall_openat(dirfd,path,flags,varargs){SYSCALLS.varargs=varargs}function ___syscall_rmdir(path){}function ___syscall_stat64(path,buf){}function ___syscall_unlinkat(dirfd,path,flags){}function __emscripten_date_now(){return Date.now()}var nowIsMonotonic=true;function __emscripten_get_now_is_monotonic(){return nowIsMonotonic}function __emscripten_throw_longjmp(){throw Infinity}function _abort(){abort("")}var _emscripten_get_now;_emscripten_get_now=(()=>performance.now());function _emscripten_memcpy_big(dest,src,num){HEAPU8.copyWithin(dest,src,src+num)}function _emscripten_get_heap_max(){return 2147483648}function emscripten_realloc_buffer(size){try{wasmMemory.grow(size-buffer.byteLength+65535>>>16);updateGlobalBufferAndViews(wasmMemory.buffer);return 1}catch(e){}}function _emscripten_resize_heap(requestedSize){var oldSize=HEAPU8.length;requestedSize=requestedSize>>>0;var maxHeapSize=_emscripten_get_heap_max();if(requestedSize>maxHeapSize){return false}let alignUp=(x,multiple)=>x+(multiple-x%multiple)%multiple;for(var cutDown=1;cutDown<=4;cutDown*=2){var overGrownHeapSize=oldSize*(1+.2/cutDown);overGrownHeapSize=Math.min(overGrownHeapSize,requestedSize+100663296);var newSize=Math.min(maxHeapSize,alignUp(Math.max(requestedSize,overGrownHeapSize),65536));var replacement=emscripten_realloc_buffer(newSize);if(replacement){return true}}return false}var ENV={};function getExecutableName(){return thisProgram||"./this.program"}function getEnvStrings(){if(!getEnvStrings.strings){var lang=(typeof navigator=="object"&&navigator.languages&&navigator.languages[0]||"C").replace("-","_")+".UTF-8";var env={"USER":"web_user","LOGNAME":"web_user","PATH":"/","PWD":"/","HOME":"/home/web_user","LANG":lang,"_":getExecutableName()};for(var x in ENV){if(ENV[x]===undefined)delete env[x];else env[x]=ENV[x]}var strings=[];for(var x in env){strings.push(x+"="+env[x])}getEnvStrings.strings=strings}return getEnvStrings.strings}function _environ_get(__environ,environ_buf){var bufSize=0;getEnvStrings().forEach(function(string,i){var ptr=environ_buf+bufSize;HEAP32[__environ+i*4>>2]=ptr;writeAsciiToMemory(string,ptr);bufSize+=string.length+1});return 0}function _environ_sizes_get(penviron_count,penviron_buf_size){var strings=getEnvStrings();HEAP32[penviron_count>>2]=strings.length;var bufSize=0;strings.forEach(function(string){bufSize+=string.length+1});HEAP32[penviron_buf_size>>2]=bufSize;return 0}function _exit(status){exit(status)}function _fd_close(fd){return 0}function _fd_read(fd,iov,iovcnt,pnum){var stream=SYSCALLS.getStreamFromFD(fd);var num=SYSCALLS.doReadv(stream,iov,iovcnt);HEAP32[pnum>>2]=num;return 0}function _fd_seek(fd,offset_low,offset_high,whence,newOffset){}function _fd_write(fd,iov,iovcnt,pnum){var num=0;for(var i=0;i<iovcnt;i++){var ptr=HEAP32[iov>>2];var len=HEAP32[iov+4>>2];iov+=8;for(var j=0;j<len;j++){SYSCALLS.printChar(fd,HEAPU8[ptr+j])}num+=len}HEAP32[pnum>>2]=num;return 0}function _getTempRet0(){return getTempRet0()}function _llvm_eh_typeid_for(type){return type}function _setTempRet0(val){setTempRet0(val)}function __isLeapYear(year){return year%4===0&&(year%100!==0||year%400===0)}function __arraySum(array,index){var sum=0;for(var i=0;i<=index;sum+=array[i++]){}return sum}var __MONTH_DAYS_LEAP=[31,29,31,30,31,30,31,31,30,31,30,31];var __MONTH_DAYS_REGULAR=[31,28,31,30,31,30,31,31,30,31,30,31];function __addDays(date,days){var newDate=new Date(date.getTime());while(days>0){var leap=__isLeapYear(newDate.getFullYear());var currentMonth=newDate.getMonth();var daysInCurrentMonth=(leap?__MONTH_DAYS_LEAP:__MONTH_DAYS_REGULAR)[currentMonth];if(days>daysInCurrentMonth-newDate.getDate()){days-=daysInCurrentMonth-newDate.getDate()+1;newDate.setDate(1);if(currentMonth<11){newDate.setMonth(currentMonth+1)}else{newDate.setMonth(0);newDate.setFullYear(newDate.getFullYear()+1)}}else{newDate.setDate(newDate.getDate()+days);return newDate}}return newDate}function _strftime(s,maxsize,format,tm){var tm_zone=HEAP32[tm+40>>2];var date={tm_sec:HEAP32[tm>>2],tm_min:HEAP32[tm+4>>2],tm_hour:HEAP32[tm+8>>2],tm_mday:HEAP32[tm+12>>2],tm_mon:HEAP32[tm+16>>2],tm_year:HEAP32[tm+20>>2],tm_wday:HEAP32[tm+24>>2],tm_yday:HEAP32[tm+28>>2],tm_isdst:HEAP32[tm+32>>2],tm_gmtoff:HEAP32[tm+36>>2],tm_zone:tm_zone?UTF8ToString(tm_zone):""};var pattern=UTF8ToString(format);var EXPANSION_RULES_1={"%c":"%a %b %d %H:%M:%S %Y","%D":"%m/%d/%y","%F":"%Y-%m-%d","%h":"%b","%r":"%I:%M:%S %p","%R":"%H:%M","%T":"%H:%M:%S","%x":"%m/%d/%y","%X":"%H:%M:%S","%Ec":"%c","%EC":"%C","%Ex":"%m/%d/%y","%EX":"%H:%M:%S","%Ey":"%y","%EY":"%Y","%Od":"%d","%Oe":"%e","%OH":"%H","%OI":"%I","%Om":"%m","%OM":"%M","%OS":"%S","%Ou":"%u","%OU":"%U","%OV":"%V","%Ow":"%w","%OW":"%W","%Oy":"%y"};for(var rule in EXPANSION_RULES_1){pattern=pattern.replace(new RegExp(rule,"g"),EXPANSION_RULES_1[rule])}var WEEKDAYS=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];var MONTHS=["January","February","March","April","May","June","July","August","September","October","November","December"];function leadingSomething(value,digits,character){var str=typeof value=="number"?value.toString():value||"";while(str.length<digits){str=character[0]+str}return str}function leadingNulls(value,digits){return leadingSomething(value,digits,"0")}function compareByDay(date1,date2){function sgn(value){return value<0?-1:value>0?1:0}var compare;if((compare=sgn(date1.getFullYear()-date2.getFullYear()))===0){if((compare=sgn(date1.getMonth()-date2.getMonth()))===0){compare=sgn(date1.getDate()-date2.getDate())}}return compare}function getFirstWeekStartDate(janFourth){switch(janFourth.getDay()){case 0:return new Date(janFourth.getFullYear()-1,11,29);case 1:return janFourth;case 2:return new Date(janFourth.getFullYear(),0,3);case 3:return new Date(janFourth.getFullYear(),0,2);case 4:return new Date(janFourth.getFullYear(),0,1);case 5:return new Date(janFourth.getFullYear()-1,11,31);case 6:return new Date(janFourth.getFullYear()-1,11,30)}}function getWeekBasedYear(date){var thisDate=__addDays(new Date(date.tm_year+1900,0,1),date.tm_yday);var janFourthThisYear=new Date(thisDate.getFullYear(),0,4);var janFourthNextYear=new Date(thisDate.getFullYear()+1,0,4);var firstWeekStartThisYear=getFirstWeekStartDate(janFourthThisYear);var firstWeekStartNextYear=getFirstWeekStartDate(janFourthNextYear);if(compareByDay(firstWeekStartThisYear,thisDate)<=0){if(compareByDay(firstWeekStartNextYear,thisDate)<=0){return thisDate.getFullYear()+1}else{return thisDate.getFullYear()}}else{return thisDate.getFullYear()-1}}var EXPANSION_RULES_2={"%a":function(date){return WEEKDAYS[date.tm_wday].substring(0,3)},"%A":function(date){return WEEKDAYS[date.tm_wday]},"%b":function(date){return MONTHS[date.tm_mon].substring(0,3)},"%B":function(date){return MONTHS[date.tm_mon]},"%C":function(date){var year=date.tm_year+1900;return leadingNulls(year/100|0,2)},"%d":function(date){return leadingNulls(date.tm_mday,2)},"%e":function(date){return leadingSomething(date.tm_mday,2," ")},"%g":function(date){return getWeekBasedYear(date).toString().substring(2)},"%G":function(date){return getWeekBasedYear(date)},"%H":function(date){return leadingNulls(date.tm_hour,2)},"%I":function(date){var twelveHour=date.tm_hour;if(twelveHour==0)twelveHour=12;else if(twelveHour>12)twelveHour-=12;return leadingNulls(twelveHour,2)},"%j":function(date){return leadingNulls(date.tm_mday+__arraySum(__isLeapYear(date.tm_year+1900)?__MONTH_DAYS_LEAP:__MONTH_DAYS_REGULAR,date.tm_mon-1),3)},"%m":function(date){return leadingNulls(date.tm_mon+1,2)},"%M":function(date){return leadingNulls(date.tm_min,2)},"%n":function(){return"\n"},"%p":function(date){if(date.tm_hour>=0&&date.tm_hour<12){return"AM"}else{return"PM"}},"%S":function(date){return leadingNulls(date.tm_sec,2)},"%t":function(){return"\t"},"%u":function(date){return date.tm_wday||7},"%U":function(date){var days=date.tm_yday+7-date.tm_wday;return leadingNulls(Math.floor(days/7),2)},"%V":function(date){var val=Math.floor((date.tm_yday+7-(date.tm_wday+6)%7)/7);if((date.tm_wday+371-date.tm_yday-2)%7<=2){val++}if(!val){val=52;var dec31=(date.tm_wday+7-date.tm_yday-1)%7;if(dec31==4||dec31==5&&__isLeapYear(date.tm_year%400-1)){val++}}else if(val==53){var jan1=(date.tm_wday+371-date.tm_yday)%7;if(jan1!=4&&(jan1!=3||!__isLeapYear(date.tm_year)))val=1}return leadingNulls(val,2)},"%w":function(date){return date.tm_wday},"%W":function(date){var days=date.tm_yday+7-(date.tm_wday+6)%7;return leadingNulls(Math.floor(days/7),2)},"%y":function(date){return(date.tm_year+1900).toString().substring(2)},"%Y":function(date){return date.tm_year+1900},"%z":function(date){var off=date.tm_gmtoff;var ahead=off>=0;off=Math.abs(off)/60;off=off/60*100+off%60;return(ahead?"+":"-")+String("0000"+off).slice(-4)},"%Z":function(date){return date.tm_zone},"%%":function(){return"%"}};pattern=pattern.replace(/%%/g,"\0\0");for(var rule in EXPANSION_RULES_2){if(pattern.includes(rule)){pattern=pattern.replace(new RegExp(rule,"g"),EXPANSION_RULES_2[rule](date))}}pattern=pattern.replace(/\0\0/g,"%");var bytes=intArrayFromString(pattern,false);if(bytes.length>maxsize){return 0}writeArrayToMemory(bytes,s);return bytes.length-1}function _strftime_l(s,maxsize,format,tm){return _strftime(s,maxsize,format,tm)}function intArrayFromString(stringy,dontAddNull,length){var len=length>0?length:lengthBytesUTF8(stringy)+1;var u8array=new Array(len);var numBytesWritten=stringToUTF8Array(stringy,u8array,0,u8array.length);if(dontAddNull)u8array.length=numBytesWritten;return u8array}var asmLibraryArg={"q":___cxa_allocate_exception,"u":___cxa_begin_catch,"t":___cxa_end_catch,"b":___cxa_find_matching_catch_2,"j":___cxa_find_matching_catch_3,"M":___cxa_free_exception,"Z":___cxa_rethrow,"E":___cxa_throw,"ia":___cxa_uncaught_exceptions,"g":___resumeException,"X":___syscall_fcntl64,"ra":___syscall_ioctl,"Y":___syscall_openat,"ma":___syscall_rmdir,"la":___syscall_stat64,"na":___syscall_unlinkat,"ta":__emscripten_date_now,"sa":__emscripten_get_now_is_monotonic,"ja":__emscripten_throw_longjmp,"L":_abort,"ua":_emscripten_memcpy_big,"ka":_emscripten_resize_heap,"oa":_environ_get,"pa":_environ_sizes_get,"D":_exit,"Q":_fd_close,"qa":_fd_read,"da":_fd_seek,"W":_fd_write,"a":_getTempRet0,"B":invoke_di,"U":invoke_diii,"V":invoke_fiii,"x":invoke_i,"c":invoke_ii,"T":invoke_iid,"z":invoke_iidd,"K":invoke_iidddddd,"f":invoke_iii,"va":invoke_iiidd,"wa":invoke_iiidddd,"e":invoke_iiii,"xa":invoke_iiiidddd,"n":invoke_iiiii,"ha":invoke_iiiiid,"r":invoke_iiiiii,"ba":invoke_iiiiiidd,"o":invoke_iiiiiii,"C":invoke_iiiiiiii,"P":invoke_iiiiiiiii,"J":invoke_iiiiiiiiiiii,"ca":invoke_jiiii,"p":invoke_v,"k":invoke_vi,"F":invoke_vid,"v":invoke_vidd,"aa":invoke_viddddiii,"h":invoke_vidi,"d":invoke_vii,"fa":invoke_viid,"s":invoke_viidd,"R":invoke_viidi,"$":invoke_viidiii,"i":invoke_viii,"S":invoke_viiiddiidd,"O":invoke_viiiffi,"m":invoke_viiii,"y":invoke_viiiii,"_":invoke_viiiiidd,"H":invoke_viiiiii,"w":invoke_viiiiiii,"N":invoke_viiiiiiii,"ya":invoke_viiiiiiiii,"G":invoke_viiiiiiiiii,"I":invoke_viiiiiiiiiiiiiii,"ea":invoke_viijii,"A":_llvm_eh_typeid_for,"l":_setTempRet0,"ga":_strftime_l};var asm=createWasm();var ___wasm_call_ctors=Module["___wasm_call_ctors"]=function(){return(___wasm_call_ctors=Module["___wasm_call_ctors"]=Module["asm"]["Aa"]).apply(null,arguments)};var _malloc=Module["_malloc"]=function(){return(_malloc=Module["_malloc"]=Module["asm"]["Ca"]).apply(null,arguments)};var _free=Module["_free"]=function(){return(_free=Module["_free"]=Module["asm"]["Da"]).apply(null,arguments)};var _ASC_FT_Malloc=Module["_ASC_FT_Malloc"]=function(){return(_ASC_FT_Malloc=Module["_ASC_FT_Malloc"]=Module["asm"]["Ea"]).apply(null,arguments)};var _ASC_FT_Free=Module["_ASC_FT_Free"]=function(){return(_ASC_FT_Free=Module["_ASC_FT_Free"]=Module["asm"]["Fa"]).apply(null,arguments)};var _ASC_FT_Init=Module["_ASC_FT_Init"]=function(){return(_ASC_FT_Init=Module["_ASC_FT_Init"]=Module["asm"]["Ga"]).apply(null,arguments)};var _ASC_FT_Done_FreeType=Module["_ASC_FT_Done_FreeType"]=function(){return(_ASC_FT_Done_FreeType=Module["_ASC_FT_Done_FreeType"]=Module["asm"]["Ha"]).apply(null,arguments)};var _ASC_FT_Set_TrueType_HintProp=Module["_ASC_FT_Set_TrueType_HintProp"]=function(){return(_ASC_FT_Set_TrueType_HintProp=Module["_ASC_FT_Set_TrueType_HintProp"]=Module["asm"]["Ia"]).apply(null,arguments)};var _ASC_FT_Open_Face=Module["_ASC_FT_Open_Face"]=function(){return(_ASC_FT_Open_Face=Module["_ASC_FT_Open_Face"]=Module["asm"]["Ja"]).apply(null,arguments)};var _ASC_FT_Done_Face=Module["_ASC_FT_Done_Face"]=function(){return(_ASC_FT_Done_Face=Module["_ASC_FT_Done_Face"]=Module["asm"]["Ka"]).apply(null,arguments)};var _ASC_FT_SetCMapForCharCode=Module["_ASC_FT_SetCMapForCharCode"]=function(){return(_ASC_FT_SetCMapForCharCode=Module["_ASC_FT_SetCMapForCharCode"]=Module["asm"]["La"]).apply(null,arguments)};var _ASC_FT_GetFaceInfo=Module["_ASC_FT_GetFaceInfo"]=function(){return(_ASC_FT_GetFaceInfo=Module["_ASC_FT_GetFaceInfo"]=Module["asm"]["Ma"]).apply(null,arguments)};var _ASC_FT_GetFaceMaxAdvanceX=Module["_ASC_FT_GetFaceMaxAdvanceX"]=function(){return(_ASC_FT_GetFaceMaxAdvanceX=Module["_ASC_FT_GetFaceMaxAdvanceX"]=Module["asm"]["Na"]).apply(null,arguments)};var _ASC_FT_GetKerningX=Module["_ASC_FT_GetKerningX"]=function(){return(_ASC_FT_GetKerningX=Module["_ASC_FT_GetKerningX"]=Module["asm"]["Oa"]).apply(null,arguments)};var _ASC_FT_Set_Transform=Module["_ASC_FT_Set_Transform"]=function(){return(_ASC_FT_Set_Transform=Module["_ASC_FT_Set_Transform"]=Module["asm"]["Pa"]).apply(null,arguments)};var _ASC_FT_Set_Char_Size=Module["_ASC_FT_Set_Char_Size"]=function(){return(_ASC_FT_Set_Char_Size=Module["_ASC_FT_Set_Char_Size"]=Module["asm"]["Qa"]).apply(null,arguments)};var _ASC_FT_Load_Glyph=Module["_ASC_FT_Load_Glyph"]=function(){return(_ASC_FT_Load_Glyph=Module["_ASC_FT_Load_Glyph"]=Module["asm"]["Ra"]).apply(null,arguments)};var _ASC_FT_Glyph_Get_CBox=Module["_ASC_FT_Glyph_Get_CBox"]=function(){return(_ASC_FT_Glyph_Get_CBox=Module["_ASC_FT_Glyph_Get_CBox"]=Module["asm"]["Sa"]).apply(null,arguments)};var _ASC_FT_Get_Glyph_Measure_Params=Module["_ASC_FT_Get_Glyph_Measure_Params"]=function(){return(_ASC_FT_Get_Glyph_Measure_Params=Module["_ASC_FT_Get_Glyph_Measure_Params"]=Module["asm"]["Ta"]).apply(null,arguments)};var _ASC_FT_Get_Glyph_Render_Params=Module["_ASC_FT_Get_Glyph_Render_Params"]=function(){return(_ASC_FT_Get_Glyph_Render_Params=Module["_ASC_FT_Get_Glyph_Render_Params"]=Module["asm"]["Ua"]).apply(null,arguments)};var _ASC_FT_Get_Glyph_Render_Buffer=Module["_ASC_FT_Get_Glyph_Render_Buffer"]=function(){return(_ASC_FT_Get_Glyph_Render_Buffer=Module["_ASC_FT_Get_Glyph_Render_Buffer"]=Module["asm"]["Va"]).apply(null,arguments)};var _ASC_HB_LanguageFromString=Module["_ASC_HB_LanguageFromString"]=function(){return(_ASC_HB_LanguageFromString=Module["_ASC_HB_LanguageFromString"]=Module["asm"]["Wa"]).apply(null,arguments)};var _ASC_HB_ShapeText=Module["_ASC_HB_ShapeText"]=function(){return(_ASC_HB_ShapeText=Module["_ASC_HB_ShapeText"]=Module["asm"]["Xa"]).apply(null,arguments)};var _ASC_HB_FontFree=Module["_ASC_HB_FontFree"]=function(){return(_ASC_HB_FontFree=Module["_ASC_HB_FontFree"]=Module["asm"]["Ya"]).apply(null,arguments)};var _Zlib_Malloc=Module["_Zlib_Malloc"]=function(){return(_Zlib_Malloc=Module["_Zlib_Malloc"]=Module["asm"]["Za"]).apply(null,arguments)};var _Zlib_Free=Module["_Zlib_Free"]=function(){return(_Zlib_Free=Module["_Zlib_Free"]=Module["asm"]["_a"]).apply(null,arguments)};var _Zlib_Create=Module["_Zlib_Create"]=function(){return(_Zlib_Create=Module["_Zlib_Create"]=Module["asm"]["$a"]).apply(null,arguments)};var _Zlib_Open=Module["_Zlib_Open"]=function(){return(_Zlib_Open=Module["_Zlib_Open"]=Module["asm"]["ab"]).apply(null,arguments)};var _Zlib_Close=Module["_Zlib_Close"]=function(){return(_Zlib_Close=Module["_Zlib_Close"]=Module["asm"]["bb"]).apply(null,arguments)};var _Zlib_AddFile=Module["_Zlib_AddFile"]=function(){return(_Zlib_AddFile=Module["_Zlib_AddFile"]=Module["asm"]["cb"]).apply(null,arguments)};var _Zlib_RemoveFile=Module["_Zlib_RemoveFile"]=function(){return(_Zlib_RemoveFile=Module["_Zlib_RemoveFile"]=Module["asm"]["db"]).apply(null,arguments)};var _Zlib_GetPaths=Module["_Zlib_GetPaths"]=function(){return(_Zlib_GetPaths=Module["_Zlib_GetPaths"]=Module["asm"]["eb"]).apply(null,arguments)};var _Zlib_GetFile=Module["_Zlib_GetFile"]=function(){return(_Zlib_GetFile=Module["_Zlib_GetFile"]=Module["asm"]["fb"]).apply(null,arguments)};var _Zlib_Save=Module["_Zlib_Save"]=function(){return(_Zlib_Save=Module["_Zlib_Save"]=Module["asm"]["gb"]).apply(null,arguments)};var _Raster_DecodeFile=Module["_Raster_DecodeFile"]=function(){return(_Raster_DecodeFile=Module["_Raster_DecodeFile"]=Module["asm"]["hb"]).apply(null,arguments)};var _Raster_GetDecodedBuffer=Module["_Raster_GetDecodedBuffer"]=function(){return(_Raster_GetDecodedBuffer=Module["_Raster_GetDecodedBuffer"]=Module["asm"]["ib"]).apply(null,arguments)};var _Raster_GetWidth=Module["_Raster_GetWidth"]=function(){return(_Raster_GetWidth=Module["_Raster_GetWidth"]=Module["asm"]["jb"]).apply(null,arguments)};var _Raster_GetHeight=Module["_Raster_GetHeight"]=function(){return(_Raster_GetHeight=Module["_Raster_GetHeight"]=Module["asm"]["kb"]).apply(null,arguments)};var _Raster_GetStride=Module["_Raster_GetStride"]=function(){return(_Raster_GetStride=Module["_Raster_GetStride"]=Module["asm"]["lb"]).apply(null,arguments)};var _Raster_Destroy=Module["_Raster_Destroy"]=function(){return(_Raster_Destroy=Module["_Raster_Destroy"]=Module["asm"]["mb"]).apply(null,arguments)};var _Raster_EncodeImageData=Module["_Raster_EncodeImageData"]=function(){return(_Raster_EncodeImageData=Module["_Raster_EncodeImageData"]=Module["asm"]["nb"]).apply(null,arguments)};var _Raster_Encode=Module["_Raster_Encode"]=function(){return(_Raster_Encode=Module["_Raster_Encode"]=Module["asm"]["ob"]).apply(null,arguments)};var _Raster_GetEncodedSize=Module["_Raster_GetEncodedSize"]=function(){return(_Raster_GetEncodedSize=Module["_Raster_GetEncodedSize"]=Module["asm"]["pb"]).apply(null,arguments)};var _Raster_GetEncodedBuffer=Module["_Raster_GetEncodedBuffer"]=function(){return(_Raster_GetEncodedBuffer=Module["_Raster_GetEncodedBuffer"]=Module["asm"]["qb"]).apply(null,arguments)};var _Raster_DestroyEncodedData=Module["_Raster_DestroyEncodedData"]=function(){return(_Raster_DestroyEncodedData=Module["_Raster_DestroyEncodedData"]=Module["asm"]["rb"]).apply(null,arguments)};var _Image_GetFormat=Module["_Image_GetFormat"]=function(){return(_Image_GetFormat=Module["_Image_GetFormat"]=Module["asm"]["sb"]).apply(null,arguments)};var _hyphenCreateApplication=Module["_hyphenCreateApplication"]=function(){return(_hyphenCreateApplication=Module["_hyphenCreateApplication"]=Module["asm"]["tb"]).apply(null,arguments)};var _hyphenDestroyApplication=Module["_hyphenDestroyApplication"]=function(){return(_hyphenDestroyApplication=Module["_hyphenDestroyApplication"]=Module["asm"]["ub"]).apply(null,arguments)};var _hyphenLoadDictionary=Module["_hyphenLoadDictionary"]=function(){return(_hyphenLoadDictionary=Module["_hyphenLoadDictionary"]=Module["asm"]["vb"]).apply(null,arguments)};var _hyphenCheckDictionary=Module["_hyphenCheckDictionary"]=function(){return(_hyphenCheckDictionary=Module["_hyphenCheckDictionary"]=Module["asm"]["wb"]).apply(null,arguments)};var _hyphenWord=Module["_hyphenWord"]=function(){return(_hyphenWord=Module["_hyphenWord"]=Module["asm"]["xb"]).apply(null,arguments)};var _setThrew=Module["_setThrew"]=function(){return(_setThrew=Module["_setThrew"]=Module["asm"]["yb"]).apply(null,arguments)};var stackSave=Module["stackSave"]=function(){return(stackSave=Module["stackSave"]=Module["asm"]["zb"]).apply(null,arguments)};var stackRestore=Module["stackRestore"]=function(){return(stackRestore=Module["stackRestore"]=Module["asm"]["Ab"]).apply(null,arguments)};var ___cxa_can_catch=Module["___cxa_can_catch"]=function(){return(___cxa_can_catch=Module["___cxa_can_catch"]=Module["asm"]["Bb"]).apply(null,arguments)};var ___cxa_is_pointer_type=Module["___cxa_is_pointer_type"]=function(){return(___cxa_is_pointer_type=Module["___cxa_is_pointer_type"]=Module["asm"]["Cb"]).apply(null,arguments)};var dynCall_viijii=Module["dynCall_viijii"]=function(){return(dynCall_viijii=Module["dynCall_viijii"]=Module["asm"]["Db"]).apply(null,arguments)};var dynCall_jiiii=Module["dynCall_jiiii"]=function(){return(dynCall_jiiii=Module["dynCall_jiiii"]=Module["asm"]["Eb"]).apply(null,arguments)};function invoke_v(index){var sp=stackSave();try{getWasmTableEntry(index)()}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vii(index,a1,a2){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iii(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vi(index,a1){var sp=stackSave();try{getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiii(index,a1,a2,a3,a4){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viii(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_ii(index,a1){var sp=stackSave();try{return getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiii(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiii(index,a1,a2,a3,a4,a5){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiffi(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiii(index,a1,a2,a3,a4,a5){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_i(index){var sp=stackSave();try{return getWasmTableEntry(index)()}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiidddd(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiidddd(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_di(index,a1){var sp=stackSave();try{return getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iid(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiidd(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vidi(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vid(index,a1,a2){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viidd(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iidd(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiidd(index,a1,a2,a3,a4){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiddiidd(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iidddddd(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viddddiii(index,a1,a2,a3,a4,a5,a6,a7,a8){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vidd(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viidi(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viidiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiidd(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiid(index,a1,a2,a3,a4,a5){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_fiii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_diii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viid(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viijii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{dynCall_viijii(index,a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_jiiii(index,a1,a2,a3,a4){var sp=stackSave();try{return dynCall_jiiii(index,a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}var calledRun;function ExitStatus(status){this.name="ExitStatus";this.message="Program terminated with exit("+status+")";this.status=status}dependenciesFulfilled=function runCaller(){if(!calledRun)run();if(!calledRun)dependenciesFulfilled=runCaller};function run(args){args=args||arguments_;if(runDependencies>0){return}preRun();if(runDependencies>0){return}function doRun(){if(calledRun)return;calledRun=true;Module["calledRun"]=true;if(ABORT)return;initRuntime();if(Module["onRuntimeInitialized"])Module["onRuntimeInitialized"]();postRun()}if(Module["setStatus"]){Module["setStatus"]("Running...");setTimeout(function(){setTimeout(function(){Module["setStatus"]("")},1);doRun()},1)}else{doRun()}}Module["run"]=run;function exit(status,implicit){EXITSTATUS=status;procExit(status)}function procExit(code){EXITSTATUS=code;if(!keepRuntimeAlive()){if(Module["onExit"])Module["onExit"](code);ABORT=true}quit_(code,new ExitStatus(code))}if(Module["preInit"]){if(typeof Module["preInit"]=="function")Module["preInit"]=[Module["preInit"]];while(Module["preInit"].length>0){Module["preInit"].pop()()}}run();


function CReturnObject()
{
	this.error = 0;
	this.freeObj = 0;
}
CReturnObject.prototype.free = function()
{
	Module["_ASC_FT_Free"](this.freeObj);
};

let g_return_obj = new CReturnObject();
let g_return_obj_count = new CReturnObject();
g_return_obj_count.count = 0;

AscFonts.CopyStreamToMemory = function(data, size)
{
	var fontStreamPointer = Module["_ASC_FT_Malloc"](size);
	Module["HEAP8"].set(data, fontStreamPointer);
	return fontStreamPointer;
};

AscFonts.GetUint8ArrayFromPointer = function(pointer, size)
{
	return new Uint8Array(Module["HEAP8"].buffer, pointer, size);
};

function CShapeString(size)
{
	this.size = size;
	this.pointer = Module["_malloc"](size);
}
CShapeString.prototype.getBuffer = function()
{
	return new Uint8Array(Module["HEAPU8"].buffer, this.pointer, this.size);
};
CShapeString.prototype.free = function()
{
	Module["_free"](this.pointer);
};
CShapeString.prototype.set = function(index, value)
{
	Module["HEAPU8"][this.pointer + index] = value;
};

AscFonts.AllocString = function(size)
{
	return new CShapeString(size);
};

AscFonts.FT_CreateLibrary = Module["_ASC_FT_Init"];
AscFonts.FT_Done_Library = Module["_ASC_FT_Done_FreeType"];
AscFonts.FT_Set_TrueType_HintProp = Module["_ASC_FT_Set_TrueType_HintProp"];

AscFonts.FT_Open_Face = Module["_ASC_FT_Open_Face"];
AscFonts.FT_Done_Face = Module["_ASC_FT_Done_Face"];
AscFonts.FT_SetCMapForCharCode = Module["_ASC_FT_SetCMapForCharCode"];
AscFonts.FT_GetKerningX = Module["_ASC_FT_GetKerningX"];
AscFonts.FT_GetFaceMaxAdvanceX = Module["_ASC_FT_GetFaceMaxAdvanceX"];
AscFonts.FT_Set_Transform = Module["_ASC_FT_Set_Transform"];
AscFonts.FT_Set_Char_Size = Module["_ASC_FT_Set_Char_Size"];
AscFonts.FT_GetFaceInfo = function(face, reader)
{
	let pointer = Module["_ASC_FT_GetFaceInfo"](face);
	if (!pointer)
	{
		g_return_obj.error = 1;
		return g_return_obj;
	}

	var len_buffer = Math.min((Module["HEAP8"].length - pointer), 1000); //max 230 symbols on name & style
	reader.init(new Uint8Array(Module["HEAP8"].buffer, pointer, len_buffer));

	g_return_obj.freeObj = pointer;
	g_return_obj.error = 0;
	return g_return_obj;
};

AscFonts.FT_Load_Glyph = Module["_ASC_FT_Load_Glyph"];
AscFonts.FT_SetCMapForCharCode = Module["_ASC_FT_SetCMapForCharCode"];
AscFonts.FT_Get_Glyph_Measure_Params = function(face, vector_worker, reader)
{
	let pointer = Module["_ASC_FT_Get_Glyph_Measure_Params"](face, vector_worker ? 1 : 0);
	if (!pointer)
	{
		g_return_obj_count.error = 1;
		return g_return_obj_count;
	}

	let len = !vector_worker ? 15 : Module["HEAP32"][pointer >> 2];
	if (vector_worker)
		len = Module["HEAP32"][pointer >> 2];

	reader.init(new Uint8Array(Module["HEAP8"].buffer, pointer + 4, 4 * (len - 1)));
	g_return_obj_count.freeObj = pointer;
	g_return_obj_count.count = len;
	g_return_obj_count.error = 0;
	return g_return_obj_count;
};
AscFonts.FT_Get_Glyph_Render_Params = function(face, render_mode, reader)
{
	let pointer = Module["_ASC_FT_Get_Glyph_Render_Params"](face, render_mode);
	if (!pointer)
	{
		g_return_obj_count.error = 1;
		return g_return_obj_count;
	}

	reader.init(new Uint8Array(Module["HEAP8"].buffer, pointer, 4 * 6));

	g_return_obj.freeObj = pointer;
	g_return_obj.error = 0;
	return g_return_obj;
};
AscFonts.FT_Get_Glyph_Render_Buffer = function(face, size)
{
	var pointer = Module["_ASC_FT_Get_Glyph_Render_Buffer"](face);
	return new Uint8Array(Module["HEAP8"].buffer, pointer, size);
};

let hb_cache_languages = {};
AscFonts.HB_FontFree = Module["ASC_HB_FontFree"];
AscFonts.HB_ShapeText = function(fontFile, text, features, script, direction, language, reader)
{
	if (!hb_cache_languages[language])
	{
		let langBuffer = language.toUtf8();
		var langPointer = Module["_malloc"](langBuffer.length);
		Module["HEAP8"].set(langBuffer, langBuffer);
		hb_cache_languages[language] = Module["_ASC_HB_LanguageFromString"](langPointer);
		Module["_free"](langPointer);
	}

	let pointer = Module["_ASC_HB_ShapeText"](fontFile["GetFace"](), fontFile["GetHBFont"](), text.pointer, features, script, direction, hb_cache_languages[language]);
	if (!pointer)
	{
		g_return_obj_count.error = 1;
		return g_return_obj_count;
	}

	let buffer = Module["HEAP8"];
	let len = (buffer[pointer + 3] & 0xFF) << 24 | (buffer[pointer + 2] & 0xFF) << 16 | (buffer[pointer + 1] & 0xFF) << 8 | (buffer[pointer] & 0xFF);

	reader.init(buffer, pointer + 4, len - 4);
	fontFile["SetHBFont"](reader.readPointer64());

	g_return_obj_count.freeObj = pointer;
	g_return_obj_count.count = (len - 12) / 26;
	g_return_obj_count.error = 0;
	return g_return_obj_count;
};

/**
 * Class representing a zip archive creator/reader.
 * @constructor
 */
function ZLib()
{
	this.engine = 0; // указатель на нативный класс Zlib
	this.files = {};
}

/**
 * Check loaded wasm/asmjs module
 */
ZLib.prototype.isModuleInit = false;

/**
 * Open archive from bytes
 * @param {Uint8Array | ArrayBuffer} buf
 * @returns {boolean} success or not
 */
ZLib.prototype.open = function(buf)
{
	if (!this.isModuleInit)
		return false;

	if (this.engine)
		this.close();

	if (!buf)
		return false;

	var arrayBuffer = (undefined !== buf.byteLength) ? new Uint8Array(buf) : buf;

	// TODO: открыли архив, и заполнили this.files
	// объектами { path : null }

	// копируем память в память webasm
	var FileRawDataSize = arrayBuffer.length;
	var FileRawData = Module["_Zlib_Malloc"](FileRawDataSize);
	if (0 == FileRawData)
		return false;
	Module["HEAP8"].set(arrayBuffer, FileRawData);

	// грузим данные
	this.engine = Module["_Zlib_Open"](FileRawData, FileRawDataSize);
	if (0 == this.engine)
	{
		Module["_Zlib_Free"](FileRawData);
		return false;
	}

	// получаем пути в архиве
	var pointer = Module["_Zlib_GetPaths"](this.engine);
	if (0 == pointer)
	{
		Module["_Zlib_Close"](this.engine);
		Module["_Zlib_Free"](FileRawData);
		return false;
	}
	var lenArray = new Int32Array(Module["HEAP8"].buffer, pointer, 4);
	var len = lenArray[0];
	len -= 4;

	var buffer = new Uint8Array(Module["HEAP8"].buffer, pointer + 4, len);
	var index = 0;
	while (index < len)
	{
		var lenRec = buffer[index] | buffer[index + 1] << 8 | buffer[index + 2] << 16 | buffer[index + 3] << 24;
		index += 4;
		var _path = "".fromUtf8(buffer, index, lenRec);
		this.files[_path] = null;
		index += lenRec;
	}
	Module["_Zlib_Free"](FileRawData);
	Module["_Zlib_Free"](pointer);
	return true;
};

/**
 * Create new archive
 * @returns {boolean} success or not
 */
ZLib.prototype.create = function()
{
	if (!this.isModuleInit)
		return false;

	if (this.engine)
		this.close();

	this.engine = Module["_Zlib_Create"]();
	return !!this.engine;
};

/**
 * Save archive from current files
 * @returns {Uint8Array | null} zip-archive bytes, or null if error
 */
ZLib.prototype.save = function()
{
	if (!this.isModuleInit || !this.engine)
		return null;

	var pointerZip = Module["_Zlib_Save"](this.engine);
	if (0 == pointerZip)
		return null;

	var _lenFile = new Int32Array(Module["HEAP8"].buffer, pointerZip, 4);
	var len = _lenFile[0];
	var zip = new Uint8Array(Module["HEAP8"].buffer, pointerZip + 4, len);
	return zip;
};

/**
 * Get all file paths in archive
 * @returns {Array}
 */
ZLib.prototype.getPaths = function()
{
	var retFiles = [];
	if (!this.files)
		return retFiles;

	for (var path in this.files) 
	{
		if (this.files.hasOwnProperty(path))
			retFiles.push(path);
	}
	return retFiles;
};

/**
 * Get uncomressed file from archive
 * @param {string} path
 * @returns {Uint8Array | null} bytes of uncompressed data, or null if error
 */
ZLib.prototype.getFile = function(path)
{
	if (!this.isModuleInit || !this.engine)
		return null;

	// проверяем - есть ли файл вообще?
	if (undefined === this.files[path])
		return null;

	// проверяем - может мы уже его разжимали?
	if (null !== this.files[path])
	{
		if (this.files[path].l > 0)
		{
			return new Uint8Array(Module["HEAP8"].buffer, this.files[path].p, this.files[path].l);
		}
		else
		{
			var _lenFile = new Int32Array(Module["HEAP8"].buffer, this.files[path].p, 4);
			var len = _lenFile[0];
			return new Uint8Array(Module["HEAP8"].buffer, this.files[path].p + 4, len);
		}
	}

	var tmp = path.toUtf8();
	var pointer = Module["_Zlib_Malloc"](tmp.length);
	if (0 == pointer)
		return null;
	Module["HEAP8"].set(tmp, pointer);

	var pointerFile = Module["_Zlib_GetFile"](this.engine, pointer);
	if (0 == pointerFile) 
	{
		Module["_Zlib_Free"](pointer);
		return null;
	}

	var _lenFile = new Int32Array(Module["HEAP8"].buffer, pointerFile, 4);
	var len = _lenFile[0];

	Module["_Zlib_Free"](pointer);
	this.files[path] = { p : pointerFile, l : 0};
	return new Uint8Array(Module["HEAP8"].buffer, pointerFile + 4, len);
};

/**
 * Add uncomressed file to archive
 * @param {string} path
 * @param {Uint8Array} new file in archive
 * @returns {boolean} success or not
 */
ZLib.prototype.addFile = function(path, data)
{
	if (!this.isModuleInit || !this.engine)
		return false;

	if (!data)
		return false;

	// проверяем - может такой файл уже есть? тогда его надо сначала удалить?
	if (undefined !== this.files[path])
		this.removeFile(path);

	var tmp = path.toUtf8();
	var pointer = Module["_Zlib_Malloc"](tmp.length);
	if (0 == pointer)
		return false;
	Module["HEAP8"].set(tmp, pointer);

	var arrayBuffer = (undefined !== data.byteLength) ? new Uint8Array(data) : data;

	var FileRawDataSize = arrayBuffer.length;
	var FileRawData = Module["_Zlib_Malloc"](FileRawDataSize);
	if (0 == FileRawData)
	{
		Module["_Zlib_Free"](pointer);
		return false;
	}
	Module["HEAP8"].set(arrayBuffer, FileRawData);
	
	Module["_Zlib_AddFile"](this.engine, pointer, FileRawData, FileRawDataSize);

	this.files[path] = { p : FileRawData, l : FileRawDataSize};
	Module["_Zlib_Free"](pointer);
	return true;
};

/**
 * Remove file from archive
 * @param {string} path
 * @returns {boolean} success or not
 */
ZLib.prototype.removeFile = function(path)
{
	if (!this.isModuleInit || !this.engine)
		return false;

	// проверяем - может такого файла и нет?
	if (undefined === this.files[path])
		return false;
		
	var tmp = path.toUtf8();
	var pointer = Module["_Zlib_Malloc"](tmp.length);
	if (0 == pointer)
		return false;
	Module["HEAP8"].set(tmp, pointer);
	
	Module["_Zlib_RemoveFile"](this.engine, pointer);

	if (this.files[path] && this.files[path].p)
	{
		Module["_Zlib_Free"](this.files[path].p);
		delete this.files[path];
	}
	Module["_Zlib_Free"](pointer);
	return true;
};

/**
 * Close & remove all used memory in archive
 * @returns {undefined}
 */
ZLib.prototype.close = function()
{
	if (!this.isModuleInit || !this.engine)
		return;

	for (var i in this.files)
	{
		if (this.files[i] && this.files[i].p)
			Module["_Zlib_Free"](this.files[i].p);
	}

	this.files = {};
	if (this.engine)
		Module["_Zlib_Free"](this.engine);
	this.engine = 0;
};

/**
 * Get image type
 * @returns {Number}
 */
ZLib.prototype.getImageType = function(path)
{
	let fileData = this.getFile(path);
	return Module["_Image_GetFormat"](this.files[path].p + 4, fileData.length);
};

/**
 * Get image in needed format
 * @returns {Uint8Array}
 */
ZLib.prototype.getImageAsFormat = function(path, format)
{
	let fileData = this.getFile(path);
	let encodedData = Module["_Raster_Encode"](this.files[path].p + 4, fileData.length, format);
	let encodedSize = Module["_Raster_GetEncodedSize"](encodedData);
	let encodedBuffer = Module["_Raster_GetEncodedBuffer"](encodedData);

	let copyData = new Uint8Array(encodedSize);
	copyData.set(new Uint8Array(Module["HEAP8"].buffer, encodedBuffer, encodedSize));

	Module["_Raster_DestroyEncodedData"](encodedData);

	return copyData;
};
/**
 * Get image as svg (for simple test)
 * @returns {string}
 */
ZLib.prototype.getImageAsSvg = function(path)
{
	let fileData = this.getFile(path);
	let encodedData = Module["_Raster_Encode"](this.files[path].p + 4, fileData.length, 24);
	let encodedSize = Module["_Raster_GetEncodedSize"](encodedData);
	let encodedBuffer = Module["_Raster_GetEncodedBuffer"](encodedData);

	let string = String.prototype.fromUtf8(new Uint8Array(Module["HEAP8"].buffer, encodedBuffer, encodedSize));

	Module["_Raster_DestroyEncodedData"](encodedData);

	return string;
};
/**
 * Get image blob for browser
 * @returns {Blob}
 */
ZLib.prototype.getImageBlob = function(path)
{
	let imageType = this.getImageType(path);
	if (imageType != 10 && imageType != 21)
	{
		return new Blob([this.getFile(path)], {type:AscCommon.openXml.GetMimeType(AscCommon.GetFileExtension(path))});
	}

	let fileData = this.getFile(path);
	let encodedData = Module["_Raster_Encode"](this.files[path].p + 4, fileData.length, 24);
	let encodedSize = Module["_Raster_GetEncodedSize"](encodedData);
	let encodedBuffer = Module["_Raster_GetEncodedBuffer"](encodedData);

	let blob = new Blob([new Uint8Array(Module["HEAP8"].buffer, encodedBuffer, encodedSize)], {type : AscCommon.openXml.GetMimeType("svg")});

	Module["_Raster_DestroyEncodedData"](encodedData);

	return blob;
};

window.AscCommon = window.AscCommon || {};
window.AscCommon.CZLibEngineJS = ZLib;

var hyphenApplication = 0;
AscFonts.Hyphen_Init = function()
{
	hyphenApplication = Module["_hyphenCreateApplication"]();
};
AscFonts.Hyphen_Destroy = function()
{
	Module["_hyphenDestroyApplication"](hyphenApplication);
};
AscFonts.Hyphen_CheckDictionary = function(lang)
{
	return false;
};
AscFonts.Hyphen_LoadDictionary = function(lang, data)
{
	let dictSize = data.byteLength;
	let dictPointer = Module["_malloc"](dictSize);
	Module["HEAP8"].set(new Uint8ClampedArray(data), dictPointer);

	let result = Module["_hyphenLoadDictionary"](hyphenApplication, lang, dictPointer, dictSize);

	Module["_free"](dictPointer);
		
	return (result === 0) ? true : false;
};

function GetUtf8SymbolLen(c)
{
	if (0x00 == (c & 0x80))
		return 1;
	else if (0x00 == (c & 0x20))
		return 2;
	else if (0x00 == (c & 0x10))
		return 3;
	else if (0x00 == (c & 0x0F))
		return 4;
	else if (0x00 == (c & 0x08))
		return 4;
	else if (0x00 == (c & 0x04))
		return 5;
	return 6;
}

AscFonts.Hyphen_Word = function(lang, word)
{
	let wordPointer = word.toUtf8Pointer();
	let wordLen = wordPointer.length;
	let hyphens = [];

	if (wordPointer) 
	{
		const ptr = Module._hyphenWord(hyphenApplication, lang, wordPointer.ptr, wordLen);
		
		let curUnicode = new Uint8ClampedArray(Module["HEAP8"].buffer, wordPointer.ptr, wordLen);
		let posUnicode = 0;
		let posUtf8 = 0;

		let vector = new Uint8ClampedArray(Module["HEAP8"].buffer, ptr, wordLen + 5);

		let pos = 0;
		while (vector[pos] != 0)
		{
			if (1 === (vector[pos] & 1))
			{
				while (posUtf8 < (pos + 1))
				{
					++posUnicode;
					posUtf8 += GetUtf8SymbolLen(curUnicode[posUtf8]);
				}				
				hyphens.push(posUnicode);
			}

			pos++;
		}

		wordPointer.free();
	}
	return hyphens;
};

AscFonts.onLoadModule();

})(window, undefined);
