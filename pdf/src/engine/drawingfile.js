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


	var Module=typeof Module!="undefined"?Module:{};var moduleOverrides=Object.assign({},Module);var arguments_=[];var thisProgram="./this.program";var quit_=(status,toThrow)=>{throw toThrow};var ENVIRONMENT_IS_WEB=true;var ENVIRONMENT_IS_WORKER=false;var scriptDirectory="";function locateFile(path){if(Module["locateFile"]){return Module["locateFile"](path,scriptDirectory)}return scriptDirectory+path}var read_,readAsync,readBinary,setWindowTitle;if(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER){if(ENVIRONMENT_IS_WORKER){scriptDirectory=self.location.href}else if(typeof document!="undefined"&&document.currentScript){scriptDirectory=document.currentScript.src}if(scriptDirectory.indexOf("blob:")!==0){scriptDirectory=scriptDirectory.substr(0,scriptDirectory.replace(/[?#].*/,"").lastIndexOf("/")+1)}else{scriptDirectory=""}{read_=(url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.send(null);return xhr.responseText});if(ENVIRONMENT_IS_WORKER){readBinary=(url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.responseType="arraybuffer";xhr.send(null);return new Uint8Array(xhr.response)})}readAsync=((url,onload,onerror)=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,true);xhr.responseType="arraybuffer";xhr.onload=(()=>{if(xhr.status==200||xhr.status==0&&xhr.response){onload(xhr.response);return}onerror()});xhr.onerror=onerror;xhr.send(null)})}setWindowTitle=(title=>document.title=title)}else{}var out=Module["print"]||console.log.bind(console);var err=Module["printErr"]||console.warn.bind(console);Object.assign(Module,moduleOverrides);moduleOverrides=null;if(Module["arguments"])arguments_=Module["arguments"];if(Module["thisProgram"])thisProgram=Module["thisProgram"];if(Module["quit"])quit_=Module["quit"];var tempRet0=0;var setTempRet0=value=>{tempRet0=value};var getTempRet0=()=>tempRet0;var wasmBinary;if(Module["wasmBinary"])wasmBinary=Module["wasmBinary"];var noExitRuntime=Module["noExitRuntime"]||true;if(typeof WebAssembly!="object"){abort("no native wasm support detected")}var wasmMemory;var ABORT=false;var EXITSTATUS;var UTF8Decoder=typeof TextDecoder!="undefined"?new TextDecoder("utf8"):undefined;function UTF8ArrayToString(heapOrArray,idx,maxBytesToRead){var endIdx=idx+maxBytesToRead;var endPtr=idx;while(heapOrArray[endPtr]&&!(endPtr>=endIdx))++endPtr;if(endPtr-idx>16&&heapOrArray.buffer&&UTF8Decoder){return UTF8Decoder.decode(heapOrArray.subarray(idx,endPtr))}else{var str="";while(idx<endPtr){var u0=heapOrArray[idx++];if(!(u0&128)){str+=String.fromCharCode(u0);continue}var u1=heapOrArray[idx++]&63;if((u0&224)==192){str+=String.fromCharCode((u0&31)<<6|u1);continue}var u2=heapOrArray[idx++]&63;if((u0&240)==224){u0=(u0&15)<<12|u1<<6|u2}else{u0=(u0&7)<<18|u1<<12|u2<<6|heapOrArray[idx++]&63}if(u0<65536){str+=String.fromCharCode(u0)}else{var ch=u0-65536;str+=String.fromCharCode(55296|ch>>10,56320|ch&1023)}}}return str}function UTF8ToString(ptr,maxBytesToRead){return ptr?UTF8ArrayToString(HEAPU8,ptr,maxBytesToRead):""}function stringToUTF8Array(str,heap,outIdx,maxBytesToWrite){if(!(maxBytesToWrite>0))return 0;var startIdx=outIdx;var endIdx=outIdx+maxBytesToWrite-1;for(var i=0;i<str.length;++i){var u=str.charCodeAt(i);if(u>=55296&&u<=57343){var u1=str.charCodeAt(++i);u=65536+((u&1023)<<10)|u1&1023}if(u<=127){if(outIdx>=endIdx)break;heap[outIdx++]=u}else if(u<=2047){if(outIdx+1>=endIdx)break;heap[outIdx++]=192|u>>6;heap[outIdx++]=128|u&63}else if(u<=65535){if(outIdx+2>=endIdx)break;heap[outIdx++]=224|u>>12;heap[outIdx++]=128|u>>6&63;heap[outIdx++]=128|u&63}else{if(outIdx+3>=endIdx)break;heap[outIdx++]=240|u>>18;heap[outIdx++]=128|u>>12&63;heap[outIdx++]=128|u>>6&63;heap[outIdx++]=128|u&63}}heap[outIdx]=0;return outIdx-startIdx}function lengthBytesUTF8(str){var len=0;for(var i=0;i<str.length;++i){var u=str.charCodeAt(i);if(u>=55296&&u<=57343)u=65536+((u&1023)<<10)|str.charCodeAt(++i)&1023;if(u<=127)++len;else if(u<=2047)len+=2;else if(u<=65535)len+=3;else len+=4}return len}function allocateUTF8(str){var size=lengthBytesUTF8(str)+1;var ret=_malloc(size);if(ret)stringToUTF8Array(str,HEAP8,ret,size);return ret}function writeArrayToMemory(array,buffer){HEAP8.set(array,buffer)}function writeAsciiToMemory(str,buffer,dontAddNull){for(var i=0;i<str.length;++i){HEAP8[buffer++>>0]=str.charCodeAt(i)}if(!dontAddNull)HEAP8[buffer>>0]=0}var buffer,HEAP8,HEAPU8,HEAP16,HEAPU16,HEAP32,HEAPU32,HEAPF32,HEAPF64;function updateGlobalBufferAndViews(buf){buffer=buf;Module["HEAP8"]=HEAP8=new Int8Array(buf);Module["HEAP16"]=HEAP16=new Int16Array(buf);Module["HEAP32"]=HEAP32=new Int32Array(buf);Module["HEAPU8"]=HEAPU8=new Uint8Array(buf);Module["HEAPU16"]=HEAPU16=new Uint16Array(buf);Module["HEAPU32"]=HEAPU32=new Uint32Array(buf);Module["HEAPF32"]=HEAPF32=new Float32Array(buf);Module["HEAPF64"]=HEAPF64=new Float64Array(buf)}var INITIAL_MEMORY=Module["INITIAL_MEMORY"]||16777216;var wasmTable;var __ATPRERUN__=[];var __ATINIT__=[];var __ATPOSTRUN__=[function(){window["AscViewer"] && window["AscViewer"]["onLoadModule"] && window["AscViewer"]["onLoadModule"]();}];var runtimeInitialized=false;function keepRuntimeAlive(){return noExitRuntime}function preRun(){if(Module["preRun"]){if(typeof Module["preRun"]=="function")Module["preRun"]=[Module["preRun"]];while(Module["preRun"].length){addOnPreRun(Module["preRun"].shift())}}callRuntimeCallbacks(__ATPRERUN__)}function initRuntime(){runtimeInitialized=true;callRuntimeCallbacks(__ATINIT__)}function postRun(){if(Module["postRun"]){if(typeof Module["postRun"]=="function")Module["postRun"]=[Module["postRun"]];while(Module["postRun"].length){addOnPostRun(Module["postRun"].shift())}}callRuntimeCallbacks(__ATPOSTRUN__)}function addOnPreRun(cb){__ATPRERUN__.unshift(cb)}function addOnInit(cb){__ATINIT__.unshift(cb)}function addOnPostRun(cb){__ATPOSTRUN__.unshift(cb)}var runDependencies=0;var runDependencyWatcher=null;var dependenciesFulfilled=null;function addRunDependency(id){runDependencies++;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}}function removeRunDependency(id){runDependencies--;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}if(runDependencies==0){if(runDependencyWatcher!==null){clearInterval(runDependencyWatcher);runDependencyWatcher=null}if(dependenciesFulfilled){var callback=dependenciesFulfilled;dependenciesFulfilled=null;callback()}}}Module["preloadedImages"]={};Module["preloadedAudios"]={};function abort(what){{if(Module["onAbort"]){Module["onAbort"](what)}}what="Aborted("+what+")";err(what);ABORT=true;EXITSTATUS=1;what+=". Build with -s ASSERTIONS=1 for more info.";var e=new WebAssembly.RuntimeError(what);throw e}var dataURIPrefix="data:application/octet-stream;base64,";function isDataURI(filename){return filename.startsWith(dataURIPrefix)}var wasmBinaryFile;wasmBinaryFile="drawingfile.wasm";if(!isDataURI(wasmBinaryFile)){wasmBinaryFile=locateFile(wasmBinaryFile)}function getBinary(file){try{if(file==wasmBinaryFile&&wasmBinary){return new Uint8Array(wasmBinary)}if(readBinary){return readBinary(file)}else{throw"both async and sync fetching of the wasm failed"}}catch(err){abort(err)}}function getBinaryPromise2(){if(!wasmBinary&&(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER)){if(typeof fetch=="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){if(!response["ok"]){throw"failed to load wasm binary file at '"+wasmBinaryFile+"'"}return response["arrayBuffer"]()}).catch(function(){return getBinary(wasmBinaryFile)})}}return Promise.resolve().then(function(){return getBinary(wasmBinaryFile)})}function createWasm(){var info={"a":asmLibraryArg};function receiveInstance(instance,module){var exports=instance.exports;Module["asm"]=exports;wasmMemory=Module["asm"]["cb"];updateGlobalBufferAndViews(wasmMemory.buffer);wasmTable=Module["asm"]["eb"];addOnInit(Module["asm"]["db"]);removeRunDependency("wasm-instantiate")}addRunDependency("wasm-instantiate");function receiveInstantiationResult(result){receiveInstance(result["instance"])}function instantiateArrayBuffer(receiver){return getBinaryPromise().then(function(binary){return WebAssembly.instantiate(binary,info)}).then(function(instance){return instance}).then(receiver,function(reason){err("failed to asynchronously prepare wasm: "+reason);abort(reason)})}function instantiateAsync(){if(!wasmBinary&&typeof WebAssembly.instantiateStreaming=="function"&&!isDataURI(wasmBinaryFile)&&typeof fetch=="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){var result=WebAssembly.instantiateStreaming(response,info);return result.then(receiveInstantiationResult,function(reason){err("wasm streaming compile failed: "+reason);err("falling back to ArrayBuffer instantiation");return instantiateArrayBuffer(receiveInstantiationResult)})})}else{return instantiateArrayBuffer(receiveInstantiationResult)}}if(Module["instantiateWasm"]){try{var exports=Module["instantiateWasm"](info,receiveInstance);return exports}catch(e){err("Module.instantiateWasm callback failed with error: "+e);return false}}instantiateAsync();return{}}function js_free_id(data){self.AscViewer.Free(data);return 1}function js_get_stream_id(data,status){return self.AscViewer.CheckStreamId(data,status)}function callRuntimeCallbacks(callbacks){while(callbacks.length>0){var callback=callbacks.shift();if(typeof callback=="function"){callback(Module);continue}var func=callback.func;if(typeof func=="number"){if(callback.arg===undefined){getWasmTableEntry(func)()}else{getWasmTableEntry(func)(callback.arg)}}else{func(callback.arg===undefined?null:callback.arg)}}}var wasmTableMirror=[];function getWasmTableEntry(funcPtr){var func=wasmTableMirror[funcPtr];if(!func){if(funcPtr>=wasmTableMirror.length)wasmTableMirror.length=funcPtr+1;wasmTableMirror[funcPtr]=func=wasmTable.get(funcPtr)}return func}function ___assert_fail(condition,filename,line,func){abort("Assertion failed: "+UTF8ToString(condition)+", at: "+[filename?UTF8ToString(filename):"unknown filename",line,func?UTF8ToString(func):"unknown function"])}function ___cxa_allocate_exception(size){return _malloc(size+16)+16}function ExceptionInfo(excPtr){this.excPtr=excPtr;this.ptr=excPtr-16;this.set_type=function(type){HEAP32[this.ptr+4>>2]=type};this.get_type=function(){return HEAP32[this.ptr+4>>2]};this.set_destructor=function(destructor){HEAP32[this.ptr+8>>2]=destructor};this.get_destructor=function(){return HEAP32[this.ptr+8>>2]};this.set_refcount=function(refcount){HEAP32[this.ptr>>2]=refcount};this.set_caught=function(caught){caught=caught?1:0;HEAP8[this.ptr+12>>0]=caught};this.get_caught=function(){return HEAP8[this.ptr+12>>0]!=0};this.set_rethrown=function(rethrown){rethrown=rethrown?1:0;HEAP8[this.ptr+13>>0]=rethrown};this.get_rethrown=function(){return HEAP8[this.ptr+13>>0]!=0};this.init=function(type,destructor){this.set_type(type);this.set_destructor(destructor);this.set_refcount(0);this.set_caught(false);this.set_rethrown(false)};this.add_ref=function(){var value=HEAP32[this.ptr>>2];HEAP32[this.ptr>>2]=value+1};this.release_ref=function(){var prev=HEAP32[this.ptr>>2];HEAP32[this.ptr>>2]=prev-1;return prev===1}}function CatchInfo(ptr){this.free=function(){_free(this.ptr);this.ptr=0};this.set_base_ptr=function(basePtr){HEAP32[this.ptr>>2]=basePtr};this.get_base_ptr=function(){return HEAP32[this.ptr>>2]};this.set_adjusted_ptr=function(adjustedPtr){HEAP32[this.ptr+4>>2]=adjustedPtr};this.get_adjusted_ptr_addr=function(){return this.ptr+4};this.get_adjusted_ptr=function(){return HEAP32[this.ptr+4>>2]};this.get_exception_ptr=function(){var isPointer=___cxa_is_pointer_type(this.get_exception_info().get_type());if(isPointer){return HEAP32[this.get_base_ptr()>>2]}var adjusted=this.get_adjusted_ptr();if(adjusted!==0)return adjusted;return this.get_base_ptr()};this.get_exception_info=function(){return new ExceptionInfo(this.get_base_ptr())};if(ptr===undefined){this.ptr=_malloc(8);this.set_adjusted_ptr(0)}else{this.ptr=ptr}}var exceptionCaught=[];function exception_addRef(info){info.add_ref()}var uncaughtExceptionCount=0;function ___cxa_begin_catch(ptr){var catchInfo=new CatchInfo(ptr);var info=catchInfo.get_exception_info();if(!info.get_caught()){info.set_caught(true);uncaughtExceptionCount--}info.set_rethrown(false);exceptionCaught.push(catchInfo);exception_addRef(info);return catchInfo.get_exception_ptr()}function ___cxa_call_unexpected(exception){err("Unexpected exception thrown, this is not properly supported - aborting");ABORT=true;throw exception}var exceptionLast=0;function ___cxa_free_exception(ptr){return _free(new ExceptionInfo(ptr).ptr)}function exception_decRef(info){if(info.release_ref()&&!info.get_rethrown()){var destructor=info.get_destructor();if(destructor){getWasmTableEntry(destructor)(info.excPtr)}___cxa_free_exception(info.excPtr)}}function ___cxa_end_catch(){_setThrew(0);var catchInfo=exceptionCaught.pop();exception_decRef(catchInfo.get_exception_info());catchInfo.free();exceptionLast=0}function ___resumeException(catchInfoPtr){var catchInfo=new CatchInfo(catchInfoPtr);var ptr=catchInfo.get_base_ptr();if(!exceptionLast){exceptionLast=ptr}catchInfo.free();throw ptr}function ___cxa_find_matching_catch_2(){var thrown=exceptionLast;if(!thrown){setTempRet0(0);return 0|0}var info=new ExceptionInfo(thrown);var thrownType=info.get_type();var catchInfo=new CatchInfo;catchInfo.set_base_ptr(thrown);catchInfo.set_adjusted_ptr(thrown);if(!thrownType){setTempRet0(0);return catchInfo.ptr|0}var typeArray=Array.prototype.slice.call(arguments);for(var i=0;i<typeArray.length;i++){var caughtType=typeArray[i];if(caughtType===0||caughtType===thrownType){break}if(___cxa_can_catch(caughtType,thrownType,catchInfo.get_adjusted_ptr_addr())){setTempRet0(caughtType);return catchInfo.ptr|0}}setTempRet0(thrownType);return catchInfo.ptr|0}function ___cxa_find_matching_catch_3(){var thrown=exceptionLast;if(!thrown){setTempRet0(0);return 0|0}var info=new ExceptionInfo(thrown);var thrownType=info.get_type();var catchInfo=new CatchInfo;catchInfo.set_base_ptr(thrown);catchInfo.set_adjusted_ptr(thrown);if(!thrownType){setTempRet0(0);return catchInfo.ptr|0}var typeArray=Array.prototype.slice.call(arguments);for(var i=0;i<typeArray.length;i++){var caughtType=typeArray[i];if(caughtType===0||caughtType===thrownType){break}if(___cxa_can_catch(caughtType,thrownType,catchInfo.get_adjusted_ptr_addr())){setTempRet0(caughtType);return catchInfo.ptr|0}}setTempRet0(thrownType);return catchInfo.ptr|0}function ___cxa_rethrow(){var catchInfo=exceptionCaught.pop();if(!catchInfo){abort("no exception to throw")}var info=catchInfo.get_exception_info();var ptr=catchInfo.get_base_ptr();if(!info.get_rethrown()){exceptionCaught.push(catchInfo);info.set_rethrown(true);info.set_caught(false);uncaughtExceptionCount++}else{catchInfo.free()}exceptionLast=ptr;throw ptr}function ___cxa_throw(ptr,type,destructor){var info=new ExceptionInfo(ptr);info.init(type,destructor);exceptionLast=ptr;uncaughtExceptionCount++;throw ptr}function ___cxa_uncaught_exceptions(){return uncaughtExceptionCount}var SYSCALLS={buffers:[null,[],[]],printChar:function(stream,curr){var buffer=SYSCALLS.buffers[stream];if(curr===0||curr===10){(stream===1?out:err)(UTF8ArrayToString(buffer,0));buffer.length=0}else{buffer.push(curr)}},varargs:undefined,get:function(){SYSCALLS.varargs+=4;var ret=HEAP32[SYSCALLS.varargs-4>>2];return ret},getStr:function(ptr){var ret=UTF8ToString(ptr);return ret},get64:function(low,high){return low}};function ___syscall_fcntl64(fd,cmd,varargs){SYSCALLS.varargs=varargs;return 0}function ___syscall_getcwd(buf,size){}function ___syscall_getdents64(fd,dirp,count){}function ___syscall_ioctl(fd,op,varargs){SYSCALLS.varargs=varargs;return 0}function ___syscall_lstat64(path,buf){}function ___syscall_mkdir(path,mode){path=SYSCALLS.getStr(path);return SYSCALLS.doMkdir(path,mode)}function ___syscall_openat(dirfd,path,flags,varargs){SYSCALLS.varargs=varargs}function ___syscall_readlinkat(dirfd,path,buf,bufsize){path=SYSCALLS.getStr(path);path=SYSCALLS.calculateAt(dirfd,path);return SYSCALLS.doReadlink(path,buf,bufsize)}function ___syscall_rmdir(path){}function ___syscall_stat64(path,buf){}function ___syscall_unlinkat(dirfd,path,flags){}function __emscripten_date_now(){return Date.now()}var nowIsMonotonic=true;function __emscripten_get_now_is_monotonic(){return nowIsMonotonic}function __emscripten_throw_longjmp(){throw Infinity}function __gmtime_js(time,tmPtr){var date=new Date(HEAP32[time>>2]*1e3);HEAP32[tmPtr>>2]=date.getUTCSeconds();HEAP32[tmPtr+4>>2]=date.getUTCMinutes();HEAP32[tmPtr+8>>2]=date.getUTCHours();HEAP32[tmPtr+12>>2]=date.getUTCDate();HEAP32[tmPtr+16>>2]=date.getUTCMonth();HEAP32[tmPtr+20>>2]=date.getUTCFullYear()-1900;HEAP32[tmPtr+24>>2]=date.getUTCDay();var start=Date.UTC(date.getUTCFullYear(),0,1,0,0,0,0);var yday=(date.getTime()-start)/(1e3*60*60*24)|0;HEAP32[tmPtr+28>>2]=yday}function __mktime_js(tmPtr){var date=new Date(HEAP32[tmPtr+20>>2]+1900,HEAP32[tmPtr+16>>2],HEAP32[tmPtr+12>>2],HEAP32[tmPtr+8>>2],HEAP32[tmPtr+4>>2],HEAP32[tmPtr>>2],0);var dst=HEAP32[tmPtr+32>>2];var guessedOffset=date.getTimezoneOffset();var start=new Date(date.getFullYear(),0,1);var summerOffset=new Date(date.getFullYear(),6,1).getTimezoneOffset();var winterOffset=start.getTimezoneOffset();var dstOffset=Math.min(winterOffset,summerOffset);if(dst<0){HEAP32[tmPtr+32>>2]=Number(summerOffset!=winterOffset&&dstOffset==guessedOffset)}else if(dst>0!=(dstOffset==guessedOffset)){var nonDstOffset=Math.max(winterOffset,summerOffset);var trueOffset=dst>0?dstOffset:nonDstOffset;date.setTime(date.getTime()+(trueOffset-guessedOffset)*6e4)}HEAP32[tmPtr+24>>2]=date.getDay();var yday=(date.getTime()-start.getTime())/(1e3*60*60*24)|0;HEAP32[tmPtr+28>>2]=yday;HEAP32[tmPtr>>2]=date.getSeconds();HEAP32[tmPtr+4>>2]=date.getMinutes();HEAP32[tmPtr+8>>2]=date.getHours();HEAP32[tmPtr+12>>2]=date.getDate();HEAP32[tmPtr+16>>2]=date.getMonth();return date.getTime()/1e3|0}function __mmap_js(addr,len,prot,flags,fd,off,allocated,builtin){return-52}function __munmap_js(addr,len,prot,flags,fd,offset){}function _tzset_impl(timezone,daylight,tzname){var currentYear=(new Date).getFullYear();var winter=new Date(currentYear,0,1);var summer=new Date(currentYear,6,1);var winterOffset=winter.getTimezoneOffset();var summerOffset=summer.getTimezoneOffset();var stdTimezoneOffset=Math.max(winterOffset,summerOffset);HEAP32[timezone>>2]=stdTimezoneOffset*60;HEAP32[daylight>>2]=Number(winterOffset!=summerOffset);function extractZone(date){var match=date.toTimeString().match(/\(([A-Za-z ]+)\)$/);return match?match[1]:"GMT"}var winterName=extractZone(winter);var summerName=extractZone(summer);var winterNamePtr=allocateUTF8(winterName);var summerNamePtr=allocateUTF8(summerName);if(summerOffset<winterOffset){HEAP32[tzname>>2]=winterNamePtr;HEAP32[tzname+4>>2]=summerNamePtr}else{HEAP32[tzname>>2]=summerNamePtr;HEAP32[tzname+4>>2]=winterNamePtr}}function __tzset_js(timezone,daylight,tzname){if(__tzset_js.called)return;__tzset_js.called=true;_tzset_impl(timezone,daylight,tzname)}function _abort(){abort("")}var _emscripten_get_now;_emscripten_get_now=(()=>performance.now());function _emscripten_memcpy_big(dest,src,num){HEAPU8.copyWithin(dest,src,src+num)}function _emscripten_get_heap_max(){return 2147483648}function emscripten_realloc_buffer(size){try{wasmMemory.grow(size-buffer.byteLength+65535>>>16);updateGlobalBufferAndViews(wasmMemory.buffer);return 1}catch(e){}}function _emscripten_resize_heap(requestedSize){var oldSize=HEAPU8.length;requestedSize=requestedSize>>>0;var maxHeapSize=_emscripten_get_heap_max();if(requestedSize>maxHeapSize){return false}let alignUp=(x,multiple)=>x+(multiple-x%multiple)%multiple;for(var cutDown=1;cutDown<=4;cutDown*=2){var overGrownHeapSize=oldSize*(1+.2/cutDown);overGrownHeapSize=Math.min(overGrownHeapSize,requestedSize+100663296);var newSize=Math.min(maxHeapSize,alignUp(Math.max(requestedSize,overGrownHeapSize),65536));var replacement=emscripten_realloc_buffer(newSize);if(replacement){return true}}return false}var ENV={};function getExecutableName(){return thisProgram||"./this.program"}function getEnvStrings(){if(!getEnvStrings.strings){var lang=(typeof navigator=="object"&&navigator.languages&&navigator.languages[0]||"C").replace("-","_")+".UTF-8";var env={"USER":"web_user","LOGNAME":"web_user","PATH":"/","PWD":"/","HOME":"/home/web_user","LANG":lang,"_":getExecutableName()};for(var x in ENV){if(ENV[x]===undefined)delete env[x];else env[x]=ENV[x]}var strings=[];for(var x in env){strings.push(x+"="+env[x])}getEnvStrings.strings=strings}return getEnvStrings.strings}function _environ_get(__environ,environ_buf){var bufSize=0;getEnvStrings().forEach(function(string,i){var ptr=environ_buf+bufSize;HEAP32[__environ+i*4>>2]=ptr;writeAsciiToMemory(string,ptr);bufSize+=string.length+1});return 0}function _environ_sizes_get(penviron_count,penviron_buf_size){var strings=getEnvStrings();HEAP32[penviron_count>>2]=strings.length;var bufSize=0;strings.forEach(function(string){bufSize+=string.length+1});HEAP32[penviron_buf_size>>2]=bufSize;return 0}function _exit(status){exit(status)}function _fd_close(fd){return 0}function _fd_read(fd,iov,iovcnt,pnum){var stream=SYSCALLS.getStreamFromFD(fd);var num=SYSCALLS.doReadv(stream,iov,iovcnt);HEAP32[pnum>>2]=num;return 0}function _fd_seek(fd,offset_low,offset_high,whence,newOffset){}function _fd_write(fd,iov,iovcnt,pnum){var num=0;for(var i=0;i<iovcnt;i++){var ptr=HEAP32[iov>>2];var len=HEAP32[iov+4>>2];iov+=8;for(var j=0;j<len;j++){SYSCALLS.printChar(fd,HEAPU8[ptr+j])}num+=len}HEAP32[pnum>>2]=num;return 0}function _getTempRet0(){return getTempRet0()}function _llvm_eh_typeid_for(type){return type}function _setTempRet0(val){setTempRet0(val)}function __isLeapYear(year){return year%4===0&&(year%100!==0||year%400===0)}function __arraySum(array,index){var sum=0;for(var i=0;i<=index;sum+=array[i++]){}return sum}var __MONTH_DAYS_LEAP=[31,29,31,30,31,30,31,31,30,31,30,31];var __MONTH_DAYS_REGULAR=[31,28,31,30,31,30,31,31,30,31,30,31];function __addDays(date,days){var newDate=new Date(date.getTime());while(days>0){var leap=__isLeapYear(newDate.getFullYear());var currentMonth=newDate.getMonth();var daysInCurrentMonth=(leap?__MONTH_DAYS_LEAP:__MONTH_DAYS_REGULAR)[currentMonth];if(days>daysInCurrentMonth-newDate.getDate()){days-=daysInCurrentMonth-newDate.getDate()+1;newDate.setDate(1);if(currentMonth<11){newDate.setMonth(currentMonth+1)}else{newDate.setMonth(0);newDate.setFullYear(newDate.getFullYear()+1)}}else{newDate.setDate(newDate.getDate()+days);return newDate}}return newDate}function _strftime(s,maxsize,format,tm){var tm_zone=HEAP32[tm+40>>2];var date={tm_sec:HEAP32[tm>>2],tm_min:HEAP32[tm+4>>2],tm_hour:HEAP32[tm+8>>2],tm_mday:HEAP32[tm+12>>2],tm_mon:HEAP32[tm+16>>2],tm_year:HEAP32[tm+20>>2],tm_wday:HEAP32[tm+24>>2],tm_yday:HEAP32[tm+28>>2],tm_isdst:HEAP32[tm+32>>2],tm_gmtoff:HEAP32[tm+36>>2],tm_zone:tm_zone?UTF8ToString(tm_zone):""};var pattern=UTF8ToString(format);var EXPANSION_RULES_1={"%c":"%a %b %d %H:%M:%S %Y","%D":"%m/%d/%y","%F":"%Y-%m-%d","%h":"%b","%r":"%I:%M:%S %p","%R":"%H:%M","%T":"%H:%M:%S","%x":"%m/%d/%y","%X":"%H:%M:%S","%Ec":"%c","%EC":"%C","%Ex":"%m/%d/%y","%EX":"%H:%M:%S","%Ey":"%y","%EY":"%Y","%Od":"%d","%Oe":"%e","%OH":"%H","%OI":"%I","%Om":"%m","%OM":"%M","%OS":"%S","%Ou":"%u","%OU":"%U","%OV":"%V","%Ow":"%w","%OW":"%W","%Oy":"%y"};for(var rule in EXPANSION_RULES_1){pattern=pattern.replace(new RegExp(rule,"g"),EXPANSION_RULES_1[rule])}var WEEKDAYS=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];var MONTHS=["January","February","March","April","May","June","July","August","September","October","November","December"];function leadingSomething(value,digits,character){var str=typeof value=="number"?value.toString():value||"";while(str.length<digits){str=character[0]+str}return str}function leadingNulls(value,digits){return leadingSomething(value,digits,"0")}function compareByDay(date1,date2){function sgn(value){return value<0?-1:value>0?1:0}var compare;if((compare=sgn(date1.getFullYear()-date2.getFullYear()))===0){if((compare=sgn(date1.getMonth()-date2.getMonth()))===0){compare=sgn(date1.getDate()-date2.getDate())}}return compare}function getFirstWeekStartDate(janFourth){switch(janFourth.getDay()){case 0:return new Date(janFourth.getFullYear()-1,11,29);case 1:return janFourth;case 2:return new Date(janFourth.getFullYear(),0,3);case 3:return new Date(janFourth.getFullYear(),0,2);case 4:return new Date(janFourth.getFullYear(),0,1);case 5:return new Date(janFourth.getFullYear()-1,11,31);case 6:return new Date(janFourth.getFullYear()-1,11,30)}}function getWeekBasedYear(date){var thisDate=__addDays(new Date(date.tm_year+1900,0,1),date.tm_yday);var janFourthThisYear=new Date(thisDate.getFullYear(),0,4);var janFourthNextYear=new Date(thisDate.getFullYear()+1,0,4);var firstWeekStartThisYear=getFirstWeekStartDate(janFourthThisYear);var firstWeekStartNextYear=getFirstWeekStartDate(janFourthNextYear);if(compareByDay(firstWeekStartThisYear,thisDate)<=0){if(compareByDay(firstWeekStartNextYear,thisDate)<=0){return thisDate.getFullYear()+1}else{return thisDate.getFullYear()}}else{return thisDate.getFullYear()-1}}var EXPANSION_RULES_2={"%a":function(date){return WEEKDAYS[date.tm_wday].substring(0,3)},"%A":function(date){return WEEKDAYS[date.tm_wday]},"%b":function(date){return MONTHS[date.tm_mon].substring(0,3)},"%B":function(date){return MONTHS[date.tm_mon]},"%C":function(date){var year=date.tm_year+1900;return leadingNulls(year/100|0,2)},"%d":function(date){return leadingNulls(date.tm_mday,2)},"%e":function(date){return leadingSomething(date.tm_mday,2," ")},"%g":function(date){return getWeekBasedYear(date).toString().substring(2)},"%G":function(date){return getWeekBasedYear(date)},"%H":function(date){return leadingNulls(date.tm_hour,2)},"%I":function(date){var twelveHour=date.tm_hour;if(twelveHour==0)twelveHour=12;else if(twelveHour>12)twelveHour-=12;return leadingNulls(twelveHour,2)},"%j":function(date){return leadingNulls(date.tm_mday+__arraySum(__isLeapYear(date.tm_year+1900)?__MONTH_DAYS_LEAP:__MONTH_DAYS_REGULAR,date.tm_mon-1),3)},"%m":function(date){return leadingNulls(date.tm_mon+1,2)},"%M":function(date){return leadingNulls(date.tm_min,2)},"%n":function(){return"\n"},"%p":function(date){if(date.tm_hour>=0&&date.tm_hour<12){return"AM"}else{return"PM"}},"%S":function(date){return leadingNulls(date.tm_sec,2)},"%t":function(){return"\t"},"%u":function(date){return date.tm_wday||7},"%U":function(date){var days=date.tm_yday+7-date.tm_wday;return leadingNulls(Math.floor(days/7),2)},"%V":function(date){var val=Math.floor((date.tm_yday+7-(date.tm_wday+6)%7)/7);if((date.tm_wday+371-date.tm_yday-2)%7<=2){val++}if(!val){val=52;var dec31=(date.tm_wday+7-date.tm_yday-1)%7;if(dec31==4||dec31==5&&__isLeapYear(date.tm_year%400-1)){val++}}else if(val==53){var jan1=(date.tm_wday+371-date.tm_yday)%7;if(jan1!=4&&(jan1!=3||!__isLeapYear(date.tm_year)))val=1}return leadingNulls(val,2)},"%w":function(date){return date.tm_wday},"%W":function(date){var days=date.tm_yday+7-(date.tm_wday+6)%7;return leadingNulls(Math.floor(days/7),2)},"%y":function(date){return(date.tm_year+1900).toString().substring(2)},"%Y":function(date){return date.tm_year+1900},"%z":function(date){var off=date.tm_gmtoff;var ahead=off>=0;off=Math.abs(off)/60;off=off/60*100+off%60;return(ahead?"+":"-")+String("0000"+off).slice(-4)},"%Z":function(date){return date.tm_zone},"%%":function(){return"%"}};pattern=pattern.replace(/%%/g,"\0\0");for(var rule in EXPANSION_RULES_2){if(pattern.includes(rule)){pattern=pattern.replace(new RegExp(rule,"g"),EXPANSION_RULES_2[rule](date))}}pattern=pattern.replace(/\0\0/g,"%");var bytes=intArrayFromString(pattern,false);if(bytes.length>maxsize){return 0}writeArrayToMemory(bytes,s);return bytes.length-1}function _strftime_l(s,maxsize,format,tm){return _strftime(s,maxsize,format,tm)}function intArrayFromString(stringy,dontAddNull,length){var len=length>0?length:lengthBytesUTF8(stringy)+1;var u8array=new Array(len);var numBytesWritten=stringToUTF8Array(stringy,u8array,0,u8array.length);if(dontAddNull)u8array.length=numBytesWritten;return u8array}var asmLibraryArg={"i":___assert_fail,"F":___cxa_allocate_exception,"s":___cxa_begin_catch,"la":___cxa_call_unexpected,"y":___cxa_end_catch,"b":___cxa_find_matching_catch_2,"j":___cxa_find_matching_catch_3,"I":___cxa_free_exception,"Q":___cxa_rethrow,"E":___cxa_throw,"$a":___cxa_uncaught_exceptions,"f":___resumeException,"aa":___syscall_fcntl64,"ya":___syscall_getcwd,"sa":___syscall_getdents64,"Ga":___syscall_ioctl,"za":___syscall_lstat64,"va":___syscall_mkdir,"U":___syscall_openat,"ra":___syscall_readlinkat,"qa":___syscall_rmdir,"Aa":___syscall_stat64,"T":___syscall_unlinkat,"_":__emscripten_date_now,"Ba":__emscripten_get_now_is_monotonic,"ab":__emscripten_throw_longjmp,"Ca":__gmtime_js,"Da":__mktime_js,"ta":__mmap_js,"ua":__munmap_js,"Ea":__tzset_js,"w":_abort,"Z":_emscripten_get_now,"Fa":_emscripten_memcpy_big,"bb":_emscripten_resize_heap,"wa":_environ_get,"xa":_environ_sizes_get,"D":_exit,"M":_fd_close,"$":_fd_read,"Za":_fd_seek,"V":_fd_write,"a":_getTempRet0,"v":invoke_di,"ca":invoke_dii,"N":invoke_diii,"Ha":invoke_fif,"pa":invoke_fiii,"u":invoke_i,"e":invoke_ii,"G":invoke_iidd,"Ua":invoke_iidddddd,"ha":invoke_iiddiii,"c":invoke_iii,"fa":invoke_iiiddddd,"ja":invoke_iiiddiii,"ka":invoke_iiiff,"Pa":invoke_iiiffff,"k":invoke_iiii,"l":invoke_iiiii,"Ja":invoke_iiiiid,"ga":invoke_iiiiiddiii,"Wa":invoke_iiiiifi,"o":invoke_iiiiii,"W":invoke_iiiiiiddiiiii,"p":invoke_iiiiiii,"z":invoke_iiiiiiii,"B":invoke_iiiiiiiii,"H":invoke_iiiiiiiiii,"X":invoke_iiiiiiiiiii,"J":invoke_iiiiiiiiiiii,"ma":invoke_iiiiiiiiiiiiiiiiiiiiiiiiiii,"Ya":invoke_jiiii,"q":invoke_v,"Ka":invoke_vdii,"d":invoke_vi,"O":invoke_vid,"R":invoke_vidd,"na":invoke_viddd,"Ta":invoke_vidddddddd,"Va":invoke_viddi,"oa":invoke_vidi,"Qa":invoke_viffffi,"h":invoke_vii,"C":invoke_viid,"Na":invoke_viidddd,"Ma":invoke_viiddddddi,"Ia":invoke_viif,"g":invoke_viii,"ia":invoke_viiid,"ea":invoke_viiiddiiiiii,"La":invoke_viiidi,"Oa":invoke_viiidiiiddddd,"n":invoke_viiii,"P":invoke_viiiid,"t":invoke_viiiii,"ba":invoke_viiiiid,"r":invoke_viiiiii,"A":invoke_viiiiiii,"K":invoke_viiiiiiii,"Y":invoke_viiiiiiiii,"L":invoke_viiiiiiiiii,"da":invoke_viiiiiiiiiiii,"Xa":invoke_viiiiiiiiiiiiii,"S":invoke_viiiiiiiiiiiiiii,"Ra":js_free_id,"Sa":js_get_stream_id,"x":_llvm_eh_typeid_for,"m":_setTempRet0,"_a":_strftime_l};var asm=createWasm();var ___wasm_call_ctors=Module["___wasm_call_ctors"]=function(){return(___wasm_call_ctors=Module["___wasm_call_ctors"]=Module["asm"]["db"]).apply(null,arguments)};var _free=Module["_free"]=function(){return(_free=Module["_free"]=Module["asm"]["fb"]).apply(null,arguments)};var _malloc=Module["_malloc"]=function(){return(_malloc=Module["_malloc"]=Module["asm"]["gb"]).apply(null,arguments)};var _InitializeFontsBin=Module["_InitializeFontsBin"]=function(){return(_InitializeFontsBin=Module["_InitializeFontsBin"]=Module["asm"]["hb"]).apply(null,arguments)};var _InitializeFontsBase64=Module["_InitializeFontsBase64"]=function(){return(_InitializeFontsBase64=Module["_InitializeFontsBase64"]=Module["asm"]["ib"]).apply(null,arguments)};var _InitializeFontsRanges=Module["_InitializeFontsRanges"]=function(){return(_InitializeFontsRanges=Module["_InitializeFontsRanges"]=Module["asm"]["jb"]).apply(null,arguments)};var _SetFontBinary=Module["_SetFontBinary"]=function(){return(_SetFontBinary=Module["_SetFontBinary"]=Module["asm"]["kb"]).apply(null,arguments)};var _IsFontBinaryExist=Module["_IsFontBinaryExist"]=function(){return(_IsFontBinaryExist=Module["_IsFontBinaryExist"]=Module["asm"]["lb"]).apply(null,arguments)};var _GetType=Module["_GetType"]=function(){return(_GetType=Module["_GetType"]=Module["asm"]["mb"]).apply(null,arguments)};var _Open=Module["_Open"]=function(){return(_Open=Module["_Open"]=Module["asm"]["nb"]).apply(null,arguments)};var _GetErrorCode=Module["_GetErrorCode"]=function(){return(_GetErrorCode=Module["_GetErrorCode"]=Module["asm"]["ob"]).apply(null,arguments)};var _Close=Module["_Close"]=function(){return(_Close=Module["_Close"]=Module["asm"]["pb"]).apply(null,arguments)};var _GetInfo=Module["_GetInfo"]=function(){return(_GetInfo=Module["_GetInfo"]=Module["asm"]["qb"]).apply(null,arguments)};var _GetPixmap=Module["_GetPixmap"]=function(){return(_GetPixmap=Module["_GetPixmap"]=Module["asm"]["rb"]).apply(null,arguments)};var _GetGlyphs=Module["_GetGlyphs"]=function(){return(_GetGlyphs=Module["_GetGlyphs"]=Module["asm"]["sb"]).apply(null,arguments)};var _GetLinks=Module["_GetLinks"]=function(){return(_GetLinks=Module["_GetLinks"]=Module["asm"]["tb"]).apply(null,arguments)};var _GetStructure=Module["_GetStructure"]=function(){return(_GetStructure=Module["_GetStructure"]=Module["asm"]["ub"]).apply(null,arguments)};var _GetInteractiveFormsInfo=Module["_GetInteractiveFormsInfo"]=function(){return(_GetInteractiveFormsInfo=Module["_GetInteractiveFormsInfo"]=Module["asm"]["vb"]).apply(null,arguments)};var _GetInteractiveFormsAP=Module["_GetInteractiveFormsAP"]=function(){return(_GetInteractiveFormsAP=Module["_GetInteractiveFormsAP"]=Module["asm"]["wb"]).apply(null,arguments)};var _GetButtonIcons=Module["_GetButtonIcons"]=function(){return(_GetButtonIcons=Module["_GetButtonIcons"]=Module["asm"]["xb"]).apply(null,arguments)};var _GetAnnotationsInfo=Module["_GetAnnotationsInfo"]=function(){return(_GetAnnotationsInfo=Module["_GetAnnotationsInfo"]=Module["asm"]["yb"]).apply(null,arguments)};var _GetAnnotationsAP=Module["_GetAnnotationsAP"]=function(){return(_GetAnnotationsAP=Module["_GetAnnotationsAP"]=Module["asm"]["zb"]).apply(null,arguments)};var _DestroyTextInfo=Module["_DestroyTextInfo"]=function(){return(_DestroyTextInfo=Module["_DestroyTextInfo"]=Module["asm"]["Ab"]).apply(null,arguments)};var _IsNeedCMap=Module["_IsNeedCMap"]=function(){return(_IsNeedCMap=Module["_IsNeedCMap"]=Module["asm"]["Bb"]).apply(null,arguments)};var _SetCMapData=Module["_SetCMapData"]=function(){return(_SetCMapData=Module["_SetCMapData"]=Module["asm"]["Cb"]).apply(null,arguments)};var _setThrew=Module["_setThrew"]=function(){return(_setThrew=Module["_setThrew"]=Module["asm"]["Db"]).apply(null,arguments)};var stackSave=Module["stackSave"]=function(){return(stackSave=Module["stackSave"]=Module["asm"]["Eb"]).apply(null,arguments)};var stackRestore=Module["stackRestore"]=function(){return(stackRestore=Module["stackRestore"]=Module["asm"]["Fb"]).apply(null,arguments)};var ___cxa_can_catch=Module["___cxa_can_catch"]=function(){return(___cxa_can_catch=Module["___cxa_can_catch"]=Module["asm"]["Gb"]).apply(null,arguments)};var ___cxa_is_pointer_type=Module["___cxa_is_pointer_type"]=function(){return(___cxa_is_pointer_type=Module["___cxa_is_pointer_type"]=Module["asm"]["Hb"]).apply(null,arguments)};var dynCall_jiiii=Module["dynCall_jiiii"]=function(){return(dynCall_jiiii=Module["dynCall_jiiii"]=Module["asm"]["Ib"]).apply(null,arguments)};function invoke_iii(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiii(index,a1,a2,a3,a4,a5){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_ii(index,a1){var sp=stackSave();try{return getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vii(index,a1,a2){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiii(index,a1,a2,a3,a4){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viii(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vi(index,a1){var sp=stackSave();try{getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiii(index,a1,a2,a3,a4,a5){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiii(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_v(index){var sp=stackSave();try{getWasmTableEntry(index)()}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_i(index){var sp=stackSave();try{return getWasmTableEntry(index)()}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiifi(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viddi(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vidi(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iidddddd(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vid(index,a1,a2){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vidd(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iidd(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vidddddddd(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viddd(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiiiiiiiiiiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viffffi(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiffff(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiff(index,a1,a2,a3,a4){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viid(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiddiii(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiid(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiidiiiddddd(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiddiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiddiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiddiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiddddd(index,a1,a2,a3,a4,a5,a6,a7){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiddiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_di(index,a1){var sp=stackSave();try{return getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_dii(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viidddd(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiddddddi(index,a1,a2,a3,a4,a5,a6,a7,a8,a9){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiid(index,a1,a2,a3,a4,a5){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiidi(index,a1,a2,a3,a4,a5){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vdii(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiid(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiid(index,a1,a2,a3,a4,a5){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_diii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viif(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_fif(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_fiii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiiiiiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_jiiii(index,a1,a2,a3,a4){var sp=stackSave();try{return dynCall_jiiii(index,a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}var calledRun;function ExitStatus(status){this.name="ExitStatus";this.message="Program terminated with exit("+status+")";this.status=status}dependenciesFulfilled=function runCaller(){if(!calledRun)run();if(!calledRun)dependenciesFulfilled=runCaller};function run(args){args=args||arguments_;if(runDependencies>0){return}preRun();if(runDependencies>0){return}function doRun(){if(calledRun)return;calledRun=true;Module["calledRun"]=true;if(ABORT)return;initRuntime();if(Module["onRuntimeInitialized"])Module["onRuntimeInitialized"]();postRun()}if(Module["setStatus"]){Module["setStatus"]("Running...");setTimeout(function(){setTimeout(function(){Module["setStatus"]("")},1);doRun()},1)}else{doRun()}}Module["run"]=run;function exit(status,implicit){EXITSTATUS=status;procExit(status)}function procExit(code){EXITSTATUS=code;if(!keepRuntimeAlive()){if(Module["onExit"])Module["onExit"](code);ABORT=true}quit_(code,new ExitStatus(code))}if(Module["preInit"]){if(typeof Module["preInit"]=="function")Module["preInit"]=[Module["preInit"]];while(Module["preInit"].length>0){Module["preInit"].pop()()}}run();


	self.drawingFileCurrentPageIndex = -1;
	self.fontStreams = {};
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

	function CFile()
	{
		this.nativeFile = 0;
		this.stream = -1;
		this.stream_size = 0;
		this.type = -1;
		this.pages = [];
		this.info = null;
		this._isNeedPassword = false;
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

		let _pages = reader.readInt();
		for (let i = 0; i < _pages; i++)
		{
			let rec = {};
			rec["W"] = reader.readInt();
			rec["H"] = reader.readInt();
			rec["Dpi"] = reader.readInt();
			rec["Rotate"] = reader.readInt();
			rec.fonts = [];
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

	CFile.prototype["getPagePixmap"] = function(pageIndex, width, height, backgroundColor)
	{
		if (this.pages[pageIndex].fonts.length > 0)
		{
			// ждем загрузки шрифтов для этой страницы
			return null;
		}

		self.drawingFileCurrentPageIndex = pageIndex;
		let retValue = Module["_GetPixmap"](this.nativeFile, pageIndex, width, height, backgroundColor === undefined ? 0xFFFFFF : backgroundColor);
		self.drawingFileCurrentPageIndex = -1;

		if (this.pages[pageIndex].fonts.length > 0)
		{
			// ждем загрузки шрифтов для этой страницы
			Module["_free"](retValue);
			retValue = null;
		}
		return retValue;
	};
	CFile.prototype["getGlyphs"] = function(pageIndex)
	{
		if (this.pages[pageIndex].fonts.length > 0)
		{
			// ждем загрузки шрифтов для этой страницы
			return null;
		}

		self.drawingFileCurrentPageIndex = pageIndex;
		let retValue = Module["_GetGlyphs"](this.nativeFile, pageIndex);
		// удалять результат не надо, этот буфер используется в качестве текстового буфера 
		// для текстовых команд других страниц. После получения ВСЕХ текстовых страниц - 
		// нужно вызвать destroyTextInfo()
		self.drawingFileCurrentPageIndex = -1;

		if (this.pages[pageIndex].fonts.length > 0)
		{
			// ждем загрузки шрифтов для этой страницы
			retValue = null;
		}

		if (null == retValue)
			return null;

		let lenArray = new Int32Array(Module["HEAP8"].buffer, retValue, 5);
		let len = lenArray[0];
		len -= 20;

		if (self.drawingFile.onUpdateStatistics)
			self.drawingFile.onUpdateStatistics(lenArray[1], lenArray[2], lenArray[3], lenArray[4]);

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
			rec["H"] = reader.readInt();
			let m = reader.readInt();
		    rec["T"] = [];
			// В массиве используются имена аннотаций - rec["name"]
		    for (let j = 0; j < m; ++j)
				rec["T"].push(reader.readString());
		}
		else if (SType == 12)
		{
			rec["Flags"] = reader.readInt();
			let m = reader.readInt();
		    rec["Fields"] = [];
			// В массиве используются имена аннотаций - rec["name"]
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
		// Номер для сопоставление с AP
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
		// Необходимо смещение полученных координат как у getStructure и viewer.navigate
		rec["rect"] = {};
		rec["rect"]["x1"] = reader.readDouble2();
		rec["rect"]["y1"] = reader.readDouble2();
		rec["rect"]["x2"] = reader.readDouble2();
		rec["rect"]["y2"] = reader.readDouble2();
		let flags = reader.readInt();
		// Уникальное имя - NM
		if (flags & (1 << 0))
			rec["UniqueName"] = reader.readString();
		// Альтернативный текст аннотации - Contents
		if (flags & (1 << 1))
			rec["Contents"] = reader.readString();
		// Эффекты границы - BE
		if (flags & (1 << 2))
		{
			rec["BE"] = {};
			rec["BE"]["S"] = reader.readByte();
			rec["BE"]["I"] = reader.readDouble();
		}
		// Специальный цвет аннотации - С
		if (flags & (1 << 3))
		{
			let n = reader.readInt();
			rec["C"] = [];
			for (let i = 0; i < n; ++i)
				rec["C"].push(reader.readDouble());
		}
		// Границы - Border/BS
		if (flags & (1 << 4))
		{
			// 0 - solid, 1 - beveled, 2 - dashed, 3 - inset, 4 - underline
			rec["border"] = reader.readByte();
			rec["borderWidth"] = reader.readDouble();
			// Dash Pattern границы
			if (rec["border"] == 2)
			{
				rec["dashed"] = [];
				rec["dashed"].push(reader.readDouble());
				rec["dashed"].push(reader.readDouble());
			}
		}
		// Дата последнего изменения - M
		if (flags & (1 << 5))
			rec["LastModified"] = reader.readString();
	}
	function readAnnotAP(reader, AP)
	{
		// Номер для сопоставление с AP
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
			// Указатель на память, аналогичный возвращаемому getPagePixmap. Память необходимо освободить
			APi["retValue"] = np2 << 32 | np1;
			let k = reader.readInt();
			if (k != 0)
				APi["fontInfo"] = [];
			for (let j = 0; j < k; ++j)
			{
				let fontInfo = {};
				fontInfo["text"] = reader.readString();
				fontInfo["fontName"] = reader.readString();
				fontInfo["fontSize"] = reader.readDouble();
				APi["fontInfo"].push(fontInfo);
			}
		}
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
			// В массиве используются имена аннотаций - rec["name"]
			res["CO"].push(reader.readString());
		
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
				rec["Parent"] = reader.readInt();
			res["Parents"].push(rec);
		}

		res["Fields"] = [];
		k = reader.readInt();
		for (let q = 0; reader.isValid() && q < k; ++q)
		{
			let rec = {};
			// Annot
			readAnnot(reader, rec);
			// Widget
			let tc = reader.readInt();
			if (tc)
			{
				rec["textColor"] = [];
				for (let i = 0; i < tc; ++i)
					rec["textColor"].push(reader.readDouble());
			}
			// 0 - left-justified, 1 - centered, 2 - right-justified
			rec["alignment"] = reader.readByte();
			// Тип аннотации виджета - FT
			// 0 - Unknown, 1 - button, 2 - radiobutton, 3 - checkbox
			// 4 - text, 5 - combobox, 6 - listbox, 7 - signature
			rec["type"] = reader.readByte();
			rec["flag"] = reader.readInt();
			let flags = reader.readInt();
			// Альтернативное имя поля, используется во всплывающей подсказке и сообщениях об ошибке - TU
			if (flags & (1 << 0))
				rec["userName"] = reader.readString();
			// Строка стиля по умолчанию (в формате CSS2) - DS
			if (flags & (1 << 1))
				rec["defaultStyle"] = reader.readString();
			// Режим выделения - H
			// 0 - none, 1 - invert, 2 - push, 3 - outline
			if (flags & (1 << 3))
				rec["highlight"] = reader.readByte();
			// Цвет границ - BC. Даже если граница не задана BS/Border, то при наличии BC предоставляется граница по-умолчанию (сплошная, толщиной 1)
			// При наличии MaxLen у text-аннотации границы появляются у каждого символа
			if (flags & (1 << 5))
			{
				let n = reader.readInt();
				rec["BC"] = [];
				for (let i = 0; i < n; ++i)
					rec["BC"].push(reader.readDouble());
			}
			// Поворот аннотации относительно страницы - R
			if (flags & (1 << 6))
				rec["rotate"] = reader.readInt();
			// Цвет фона аннотации - BG
			if (flags & (1 << 7))
			{
				let n = reader.readInt();
				rec["BG"] = [];
				for (let i = 0; i < n; ++i)
					rec["BG"].push(reader.readDouble());
			}
			// Значение по-умолчанию - DV
			if (flags & (1 << 8))
				rec["defaultValue"] = reader.readString();
			if (flags & (1 << 17))
				rec["Parent"] = reader.readInt();
			if (flags & (1 << 18))
				rec["name"] = reader.readString();
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
			if (rec["type"] == 3 || rec["type"] == 2 || rec["type"] == 1)
			{
				rec["value"] = (flags & (1 << 9)) ? "Yes" : "Off";
				let IFflags = reader.readInt();
				// Характеристики внешнего вида - MK
				if (rec["type"] == 1)
				{
					// Заголовок - СА
					if (flags & (1 << 10))
						rec["caption"] = reader.readString();
					// Заголовок прокрутки - RC
					if (flags & (1 << 11))
						rec["rolloverCaption"] = reader.readString();
					// Альтернативный заголовок - AC
					if (flags & (1 << 12))
						rec["alternateCaption"] = reader.readString();
				}
				else
					// 0 - check, 1 - cross, 2 - diamond, 3 - circle, 4 - star, 5 - square
					rec["style"] = reader.readByte();
				// Расположение заголовка - TP
				if (flags & (1 << 13))
					// 0 - textOnly, 1 - iconOnly, 2 - iconTextV, 3 - textIconV, 4 - iconTextH, 5 - textIconH, 6 - overlay
					rec["position"] = reader.readByte();
				// Справочный словарь иконок - IF
				if (IFflags & (1 << 0))
				{
					rec["IF"] = {};
					// Обстоятельства масштабирования IF.SW
					// 0 - Always, 1 - Never, 2 - too big, 3 - too small
					if (IFflags & (1 << 1))
						rec["IF"]["SW"] = reader.readByte();
					// Тип масштабирования - IF.S
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
			    if (flags & (1 << 14))
				{
					rec["NameOfYes"] = reader.readString();
					if (flags & (1 << 9))
						rec["value"] = rec["NameOfYes"];
				}
				// 12.7.4.2.1
				rec["NoToggleToOff"]  = (rec["flag"] >> 14) & 1; // NoToggleToOff
				rec["radiosInUnison"] = (rec["flag"] >> 25) & 1; // RadiosInUnison
			}
			else if (rec["type"] == 4)
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
			else if (rec["type"] == 5 || rec["type"] == 6)
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
				// 12.7.4.4
				rec["editable"]          = (rec["flag"] >> 18) & 1; // Edit
				rec["multipleSelection"] = (rec["flag"] >> 21) & 1; // MultiSelect
				rec["doNotSpellCheck"]   = (rec["flag"] >> 22) & 1; // DoNotSpellCheck
				rec["commitOnSelChange"] = (rec["flag"] >> 26) & 1; // CommitOnSelChange
			}
			// 12.7.3.1
			rec["readOnly"] = (rec["flag"] >> 0) & 1; // ReadOnly
			rec["required"] = (rec["flag"] >> 1) & 1; // Required
			rec["noexport"] = (rec["flag"] >> 2) & 1; // NoExport

			res["Fields"].push(rec);
		}

		Module["_free"](ext);
		return res;
	};
	// необязательный nWidget - номер сопоставления аннотации - rec["AP"]["i"]
	// необязательный sView   - определенный внешний вид аннотации - N/D/R
	// необязательный sButtonView - состояние pushbutton-аннотации - Off/Yes(или rec["NameOfYes"])
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
		let ext = Module["_GetInteractiveFormsAP"](this.nativeFile, width, height, backgroundColor === undefined ? 0xFFFFFF : backgroundColor, pageIndex, nWidget === undefined ? -1 : nWidget, nView, nButtonView);
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
			// Внешний вид аннотации
			let AP = {};
			readAnnotAP(reader, AP);
			res.push(AP);
		}

		Module["_free"](ext);
		return res;
	};
	// необязательный nWidget ...
	// необязательный sIconView - определенная иконка - I/RI/IX
	CFile.prototype["getButtonIcons"] = function(pageIndex, width, height, backgroundColor, nWidget, sIconView)
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
		let ext = Module["_GetButtonIcons"](this.nativeFile, width, height, backgroundColor === undefined ? 0xFFFFFF : backgroundColor, pageIndex, nWidget === undefined ? -1 : nWidget, nView);
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
			// Внешний вид pushbutton аннотации
			let MK = {};
			// Номер для сопоставление с AP
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
					let np1 = reader.readInt();
					let np2 = reader.readInt();
					// Указатель на память, аналогичный возвращаемому getPagePixmap. Память необходимо освободить
					ViewMK["retValue"] = np2 << 32 | np1;
					res["View"].push(ViewMK);
				}
			}
			res["MK"].push(MK);
		}

		Module["_free"](ext);
		return res;
	};
	// необязательный pageIndex - получить данные для аннотаций на конкретной странице
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
			// Тип аннотации
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
				// Номер AP popup аннотации для сопоставления
				if (flags & (1 << 0))
					rec["Popup"] = reader.readInt();
				// Текстовая метка пользователя - T
				if (flags & (1 << 1))
					rec["User"] = reader.readString();
				// Значение непрозрачности - CA
				if (flags & (1 << 2))
					rec["CA"] = reader.readDouble();
				// Форматированный текст - RC
				if (flags & (1 << 3))
					rec["RC"] = reader.readString();
				// Дата создания - CreationDate
				if (flags & (1 << 4))
					rec["CreationDate"] = reader.readString();
				// Ссылка на аннотацию ответ - IRT
				if (flags & (1 << 5))
					rec["RefTo"] = reader.readInt();
				// Тип аннотации ответа - RT
				// 0 - R, 1 - Group
				if (flags & (1 << 6))
					rec["RefToReason"] = reader.readByte();
				// Краткое описание - Subj
				if (flags & (1 << 7))
					rec["Subj"] = reader.readString();
			}
			// Text
			if (rec["Type"] == 0)
			{
				rec["Open"] = (flags >> 15) & 1;
				// иконка - Name
				// 0 - Comment, 1 - Key, 2 - Note, 3 - Help, 4 - NewParagraph, 5 - Paragraph, 6 - Insert
				if (flags & (1 << 16))
					rec["Icon"] = reader.readByte();
				// Модель состояния - StateModel
				// 0 - Marked, 1 - Review
				if (flags & (1 << 17))
					rec["StateModel"] = reader.readByte();
				// Состояние - State
				// 0 - Marked, 1 - Unmarked, 2 - Accepted, 3 - Rejected, 4 - Cancelled, 5 - Completed, 6 - None
				if (flags & (1 << 18))
					rec["State"] = reader.readByte();
				
			}
			// Line
			else if (rec["Type"] == 3)
			{
				// Координаты линии - L
				rec["L"] = [];
				for (let i = 0; i < 4; ++i)
					rec["L"].push(reader.readDouble());
				// Стили окончания линии - LE
				// 0 - Square, 1 - Circle, 2 - Diamond, 3 - OpenArrow, 4 - ClosedArrow, 5 - None, 6 - Butt, 7 - ROpenArrow, 8 - RClosedArrow, 9 - Slash
				if (flags & (1 << 15))
				{
					rec["LE"] = [];
					rec["LE"].push(reader.readByte());
					rec["LE"].push(reader.readByte());
				}
				// Цвет окончаний линии - IC
				if (flags & (1 << 16))
				{
					let n = reader.readInt();
					rec["IC"] = [];
					for (let i = 0; i < n; ++i)
						rec["IC"].push(reader.readDouble());
				}
				// Длина линий выноски - LL
				if (flags & (1 << 17))
					rec["LL"] = reader.readDouble();
				// Продолжение линий выноски - LLE
				if (flags & (1 << 18))
					rec["LLE"] = reader.readDouble();
				// Местоположение заголовка - Cap
				rec["Cap"] = (flags >> 19) & 1;
				// Назначение аннотации - IT
				// 0 - LineDimension, 1 - LineArrow
				if (flags & (1 << 20))
					rec["IT"] = reader.readByte();
				// Длина смещения выноски - LLO
				if (flags & (1 << 21))
					rec["LLO"] = reader.readDouble();
				// Расположение заголовка аннотации - CP
				// 0 - Inline, 1 - Top
				if (flags & (1 << 22))
					rec["CP"] = reader.readByte();
				// Смещение текста подписи - CO
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
				// Необходимо смещение полученных координат как у getStructure и viewer.navigate
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
				// Координаты - QuadPoints
				let n = reader.readInt();
				rec["QuadPoints"] = [];
				for (let i = 0; i < n; ++i)
					rec["QuadPoints"].push(reader.readDouble());
			}
			// Square, Circle
			else if (rec["Type"] == 4 || rec["Type"] == 5)
			{
				// Различия Rect и фактического размера - RD
				if (flags & (1 << 15))
				{
					rec["RD"] = [];
					for (let i = 0; i < 4; ++i)
						rec["RD"].push(reader.readDouble());
				}
				// Цвет заполнения - IC
				if (flags & (1 << 16))
				{
					let n = reader.readInt();
					rec["IC"] = [];
					for (let i = 0; i < n; ++i)
						rec["IC"].push(reader.readDouble());
				}
			}
			// Polygon, PolyLine
			else if (rec["Type"] == 6 || rec["Type"] == 7)
			{
				let nVertices = reader.readInt();
				rec["Vertices"] = [];
				for (let i = 0; i < nVertices; ++i)
					rec["Vertices"].push(reader.readDouble());
				// Стили окончания линии - LE
				// 0 - Square, 1 - Circle, 2 - Diamond, 3 - OpenArrow, 4 - ClosedArrow, 5 - None, 6 - Butt, 7 - ROpenArrow, 8 - RClosedArrow, 9 - Slash
				if (flags & (1 << 15))
				{
					rec["LE"] = [];
					rec["LE"].push(reader.readByte());
					rec["LE"].push(reader.readByte());
				}
				// Цвет окончаний линии - IC
				if (flags & (1 << 16))
				{
					let n = reader.readInt();
					rec["IC"] = [];
					for (let i = 0; i < n; ++i)
						rec["IC"].push(reader.readDouble());
				}
				// Назначение аннотации - IT
				// 0 - PolygonCloud, 1 - PolyLineDimension, 2 - PolygonDimension
				if (flags & (1 << 20))
					rec["IT"] = reader.readByte();
			}
			// Popup
			else if (rec["Type"] == 15)
			{
				flags = reader.readInt();
				rec["Open"] = (flags >> 0) & 1;
				// Ссылка на аннотацию-родителя
				if (flags & (1 << 1))
					rec["PopupParent"] = reader.readInt();
			}
			// FreeText
			else if (rec["Type"] == 2)
			{
				// 0 - left-justified, 1 - centered, 2 - right-justified
				rec["alignment"] = reader.readByte();
				// Различия Rect и фактического размера - RD
				if (flags & (1 << 15))
				{
					rec["RD"] = [];
					for (let i = 0; i < 4; ++i)
						rec["RD"].push(reader.readDouble());
				}
				// Координаты выноски - CL
				if (flags & (1 << 16))
				{
					let n = reader.readInt();
					rec["CL"] = [];
					for (let i = 0; i < n; ++i)
						rec["CL"].push(reader.readDouble());
				}
				// Строка стиля по умолчанию (в формате CSS2) - DS
				if (flags & (1 << 17))
					rec["defaultStyle"] = reader.readString();
				// Стиль окончания линии - LE
				// 0 - Square, 1 - Circle, 2 - Diamond, 3 - OpenArrow, 4 - ClosedArrow, 5 - None, 6 - Butt, 7 - ROpenArrow, 8 - RClosedArrow, 9 - Slash
				if (flags & (1 << 18))
					rec["LE"] = reader.readByte();
				// Назначение аннотации - IT
				// 0 - FreeText, 1 - FreeTextCallout, 2 - FreeTextTypeWriter
				if (flags & (1 << 20))
					rec["IT"] = reader.readByte();
			}
			// Caret
			else if (rec["Type"] == 13)
			{
				// Различия Rect и фактического размера - RD
				if (flags & (1 << 15))
				{
					rec["RD"] = [];
					for (let i = 0; i < 4; ++i)
						rec["RD"].push(reader.readDouble());
				}
				// Связанный символ - Sy
				// 0 - P, 1 - None
				if (flags & (1 << 16))
					rec["Sy"] = reader.readByte();
			}
			res.push(rec);
		}
		
		Module["_free"](ext);
		return res;
	};
	// необязательный nAnnot ...
	// необязательный sView ...
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
		let ext = Module["_GetAnnotationsAP"](this.nativeFile, width, height, backgroundColor === undefined ? 0xFFFFFF : backgroundColor, pageIndex, nAnnot === undefined ? -1 : nAnnot, nView);
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
			// Внешний вид аннотации
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

	CFile.prototype.memory = function()
	{
		return Module["HEAP8"];
	};
	CFile.prototype.free = function(pointer)
	{
		Module["_free"](pointer);
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
			// шрифт загружен.
			fontToMemory(file, true);
		}
		else
		{
			self.fontStreams[fileId] = self.fontStreams[fileId] || {};
			self.fontStreams[fileId].pages = self.fontStreams[fileId].pages || [];
			addToArrayAsDictionary(self.fontStreams[fileId].pages, self.drawingFileCurrentPageIndex);

			if (self.drawingFile)
			{
				addToArrayAsDictionary(self.drawingFile.pages[self.drawingFileCurrentPageIndex].fonts, fileId);
			}

			// шрифт может грузиться в редакторе
			if (undefined === file.externalCallback)
			{
				let _t = file;
				file.externalCallback = function() {
					fontToMemory(_t, true);

					let pages = self.fontStreams[fileId].pages;
					delete self.fontStreams[fileId];
					let pagesRepaint = [];
					for (let i = 0, len = pages.length; i < len; i++)
					{
						let pageObj = self.drawingFile.pages[pages[i]];
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
							pagesRepaint.push(pages[i]);
					}

					if (pagesRepaint.length > 0)
					{
						if (self.drawingFile.onRepaintPages)
							self.drawingFile.onRepaintPages(pagesRepaint);
					}

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

		// не скидываем стрим, чтобы можно было использовать его а fonts.js
		//let streams = AscFonts.getFontStreams();
		//streams[stream_index] = null;
		//streams[stream_index] = AscFonts.updateFontStreamNative(streamPointer, stream.size);

		Module["_SetFontBinary"](idPointer, streamPointer, stream.size);

		Module["_free"](streamPointer);
		Module["_free"](idPointer);
	}
})(window, undefined);
