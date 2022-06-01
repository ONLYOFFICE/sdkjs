/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

	window['AscFonts'] = window['AscFonts'] || {};
	var AscFonts = window['AscFonts'];

	AscFonts.TT_INTERPRETER_VERSION_35 = 35;
	AscFonts.TT_INTERPRETER_VERSION_38 = 38;
	AscFonts.TT_INTERPRETER_VERSION_40 = 40;

	

// correct fetch for desktop application

var printErr = undefined;
var print    = undefined;

var fetch = self.fetch;
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
			len = buffer.length;

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
	String.prototype.toUtf8 = function(isNoEndNull) {
		var inputLen = this.length;
		var testLen  = 6 * inputLen + 1;
		var tmpStrings = new ArrayBuffer(testLen);

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

})();


var Module=typeof Module!="undefined"?Module:{};var moduleOverrides=Object.assign({},Module);var arguments_=[];var thisProgram="./this.program";var quit_=(status,toThrow)=>{throw toThrow};var ENVIRONMENT_IS_WEB=true;var ENVIRONMENT_IS_WORKER=false;var scriptDirectory="";function locateFile(path){if(Module["locateFile"]){return Module["locateFile"](path,scriptDirectory)}return scriptDirectory+path}var read_,readAsync,readBinary,setWindowTitle;if(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER){if(ENVIRONMENT_IS_WORKER){scriptDirectory=self.location.href}else if(typeof document!="undefined"&&document.currentScript){scriptDirectory=document.currentScript.src}if(scriptDirectory.indexOf("blob:")!==0){scriptDirectory=scriptDirectory.substr(0,scriptDirectory.replace(/[?#].*/,"").lastIndexOf("/")+1)}else{scriptDirectory=""}{read_=(url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.send(null);return xhr.responseText});if(ENVIRONMENT_IS_WORKER){readBinary=(url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.responseType="arraybuffer";xhr.send(null);return new Uint8Array(xhr.response)})}readAsync=((url,onload,onerror)=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,true);xhr.responseType="arraybuffer";xhr.onload=(()=>{if(xhr.status==200||xhr.status==0&&xhr.response){onload(xhr.response);return}onerror()});xhr.onerror=onerror;xhr.send(null)})}setWindowTitle=(title=>document.title=title)}else{}var out=Module["print"]||console.log.bind(console);var err=Module["printErr"]||console.warn.bind(console);Object.assign(Module,moduleOverrides);moduleOverrides=null;if(Module["arguments"])arguments_=Module["arguments"];if(Module["thisProgram"])thisProgram=Module["thisProgram"];if(Module["quit"])quit_=Module["quit"];var tempRet0=0;var setTempRet0=value=>{tempRet0=value};var getTempRet0=()=>tempRet0;var wasmBinary;if(Module["wasmBinary"])wasmBinary=Module["wasmBinary"];var noExitRuntime=Module["noExitRuntime"]||true;if(typeof WebAssembly!="object"){abort("no native wasm support detected")}var wasmMemory;var ABORT=false;var EXITSTATUS;var UTF8Decoder=typeof TextDecoder!="undefined"?new TextDecoder("utf8"):undefined;function UTF8ArrayToString(heapOrArray,idx,maxBytesToRead){var endIdx=idx+maxBytesToRead;var endPtr=idx;while(heapOrArray[endPtr]&&!(endPtr>=endIdx))++endPtr;if(endPtr-idx>16&&heapOrArray.buffer&&UTF8Decoder){return UTF8Decoder.decode(heapOrArray.subarray(idx,endPtr))}else{var str="";while(idx<endPtr){var u0=heapOrArray[idx++];if(!(u0&128)){str+=String.fromCharCode(u0);continue}var u1=heapOrArray[idx++]&63;if((u0&224)==192){str+=String.fromCharCode((u0&31)<<6|u1);continue}var u2=heapOrArray[idx++]&63;if((u0&240)==224){u0=(u0&15)<<12|u1<<6|u2}else{u0=(u0&7)<<18|u1<<12|u2<<6|heapOrArray[idx++]&63}if(u0<65536){str+=String.fromCharCode(u0)}else{var ch=u0-65536;str+=String.fromCharCode(55296|ch>>10,56320|ch&1023)}}}return str}function UTF8ToString(ptr,maxBytesToRead){return ptr?UTF8ArrayToString(HEAPU8,ptr,maxBytesToRead):""}function writeAsciiToMemory(str,buffer,dontAddNull){for(var i=0;i<str.length;++i){HEAP8[buffer++>>0]=str.charCodeAt(i)}if(!dontAddNull)HEAP8[buffer>>0]=0}var buffer,HEAP8,HEAPU8,HEAP16,HEAPU16,HEAP32,HEAPU32,HEAPF32,HEAPF64;function updateGlobalBufferAndViews(buf){buffer=buf;Module["HEAP8"]=HEAP8=new Int8Array(buf);Module["HEAP16"]=HEAP16=new Int16Array(buf);Module["HEAP32"]=HEAP32=new Int32Array(buf);Module["HEAPU8"]=HEAPU8=new Uint8Array(buf);Module["HEAPU16"]=HEAPU16=new Uint16Array(buf);Module["HEAPU32"]=HEAPU32=new Uint32Array(buf);Module["HEAPF32"]=HEAPF32=new Float32Array(buf);Module["HEAPF64"]=HEAPF64=new Float64Array(buf)}var INITIAL_MEMORY=Module["INITIAL_MEMORY"]||16777216;var wasmTable;var __ATPRERUN__=[];var __ATINIT__=[];var __ATPOSTRUN__=[function(){window["AscFonts"].onLoadModule();}];var runtimeInitialized=false;function preRun(){if(Module["preRun"]){if(typeof Module["preRun"]=="function")Module["preRun"]=[Module["preRun"]];while(Module["preRun"].length){addOnPreRun(Module["preRun"].shift())}}callRuntimeCallbacks(__ATPRERUN__)}function initRuntime(){runtimeInitialized=true;callRuntimeCallbacks(__ATINIT__)}function postRun(){if(Module["postRun"]){if(typeof Module["postRun"]=="function")Module["postRun"]=[Module["postRun"]];while(Module["postRun"].length){addOnPostRun(Module["postRun"].shift())}}callRuntimeCallbacks(__ATPOSTRUN__)}function addOnPreRun(cb){__ATPRERUN__.unshift(cb)}function addOnInit(cb){__ATINIT__.unshift(cb)}function addOnPostRun(cb){__ATPOSTRUN__.unshift(cb)}var runDependencies=0;var runDependencyWatcher=null;var dependenciesFulfilled=null;function addRunDependency(id){runDependencies++;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}}function removeRunDependency(id){runDependencies--;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}if(runDependencies==0){if(runDependencyWatcher!==null){clearInterval(runDependencyWatcher);runDependencyWatcher=null}if(dependenciesFulfilled){var callback=dependenciesFulfilled;dependenciesFulfilled=null;callback()}}}Module["preloadedImages"]={};Module["preloadedAudios"]={};function abort(what){{if(Module["onAbort"]){Module["onAbort"](what)}}what="Aborted("+what+")";err(what);ABORT=true;EXITSTATUS=1;what+=". Build with -s ASSERTIONS=1 for more info.";var e=new WebAssembly.RuntimeError(what);throw e}var dataURIPrefix="data:application/octet-stream;base64,";function isDataURI(filename){return filename.startsWith(dataURIPrefix)}var wasmBinaryFile;wasmBinaryFile="fonts.wasm";if(!isDataURI(wasmBinaryFile)){wasmBinaryFile=locateFile(wasmBinaryFile)}function getBinary(file){try{if(file==wasmBinaryFile&&wasmBinary){return new Uint8Array(wasmBinary)}if(readBinary){return readBinary(file)}else{throw"both async and sync fetching of the wasm failed"}}catch(err){abort(err)}}function getBinaryPromise2(){if(!wasmBinary&&(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER)){if(typeof fetch=="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){if(!response["ok"]){throw"failed to load wasm binary file at '"+wasmBinaryFile+"'"}return response["arrayBuffer"]()}).catch(function(){return getBinary(wasmBinaryFile)})}}return Promise.resolve().then(function(){return getBinary(wasmBinaryFile)})}function createWasm(){var info={"a":asmLibraryArg};function receiveInstance(instance,module){var exports=instance.exports;Module["asm"]=exports;wasmMemory=Module["asm"]["K"];updateGlobalBufferAndViews(wasmMemory.buffer);wasmTable=Module["asm"]["N"];addOnInit(Module["asm"]["L"]);removeRunDependency("wasm-instantiate")}addRunDependency("wasm-instantiate");function receiveInstantiationResult(result){receiveInstance(result["instance"])}function instantiateArrayBuffer(receiver){return getBinaryPromise().then(function(binary){return WebAssembly.instantiate(binary,info)}).then(function(instance){return instance}).then(receiver,function(reason){err("failed to asynchronously prepare wasm: "+reason);abort(reason)})}function instantiateAsync(){if(!wasmBinary&&typeof WebAssembly.instantiateStreaming=="function"&&!isDataURI(wasmBinaryFile)&&typeof fetch=="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){var result=WebAssembly.instantiateStreaming(response,info);return result.then(receiveInstantiationResult,function(reason){err("wasm streaming compile failed: "+reason);err("falling back to ArrayBuffer instantiation");return instantiateArrayBuffer(receiveInstantiationResult)})})}else{return instantiateArrayBuffer(receiveInstantiationResult)}}if(Module["instantiateWasm"]){try{var exports=Module["instantiateWasm"](info,receiveInstance);return exports}catch(e){err("Module.instantiateWasm callback failed with error: "+e);return false}}instantiateAsync();return{}}function callRuntimeCallbacks(callbacks){while(callbacks.length>0){var callback=callbacks.shift();if(typeof callback=="function"){callback(Module);continue}var func=callback.func;if(typeof func=="number"){if(callback.arg===undefined){getWasmTableEntry(func)()}else{getWasmTableEntry(func)(callback.arg)}}else{func(callback.arg===undefined?null:callback.arg)}}}var wasmTableMirror=[];function getWasmTableEntry(funcPtr){var func=wasmTableMirror[funcPtr];if(!func){if(funcPtr>=wasmTableMirror.length)wasmTableMirror.length=funcPtr+1;wasmTableMirror[funcPtr]=func=wasmTable.get(funcPtr)}return func}function ___cxa_allocate_exception(size){return _malloc(size+16)+16}function ExceptionInfo(excPtr){this.excPtr=excPtr;this.ptr=excPtr-16;this.set_type=function(type){HEAP32[this.ptr+4>>2]=type};this.get_type=function(){return HEAP32[this.ptr+4>>2]};this.set_destructor=function(destructor){HEAP32[this.ptr+8>>2]=destructor};this.get_destructor=function(){return HEAP32[this.ptr+8>>2]};this.set_refcount=function(refcount){HEAP32[this.ptr>>2]=refcount};this.set_caught=function(caught){caught=caught?1:0;HEAP8[this.ptr+12>>0]=caught};this.get_caught=function(){return HEAP8[this.ptr+12>>0]!=0};this.set_rethrown=function(rethrown){rethrown=rethrown?1:0;HEAP8[this.ptr+13>>0]=rethrown};this.get_rethrown=function(){return HEAP8[this.ptr+13>>0]!=0};this.init=function(type,destructor){this.set_type(type);this.set_destructor(destructor);this.set_refcount(0);this.set_caught(false);this.set_rethrown(false)};this.add_ref=function(){var value=HEAP32[this.ptr>>2];HEAP32[this.ptr>>2]=value+1};this.release_ref=function(){var prev=HEAP32[this.ptr>>2];HEAP32[this.ptr>>2]=prev-1;return prev===1}}function CatchInfo(ptr){this.free=function(){_free(this.ptr);this.ptr=0};this.set_base_ptr=function(basePtr){HEAP32[this.ptr>>2]=basePtr};this.get_base_ptr=function(){return HEAP32[this.ptr>>2]};this.set_adjusted_ptr=function(adjustedPtr){HEAP32[this.ptr+4>>2]=adjustedPtr};this.get_adjusted_ptr_addr=function(){return this.ptr+4};this.get_adjusted_ptr=function(){return HEAP32[this.ptr+4>>2]};this.get_exception_ptr=function(){var isPointer=___cxa_is_pointer_type(this.get_exception_info().get_type());if(isPointer){return HEAP32[this.get_base_ptr()>>2]}var adjusted=this.get_adjusted_ptr();if(adjusted!==0)return adjusted;return this.get_base_ptr()};this.get_exception_info=function(){return new ExceptionInfo(this.get_base_ptr())};if(ptr===undefined){this.ptr=_malloc(8);this.set_adjusted_ptr(0)}else{this.ptr=ptr}}var exceptionCaught=[];function exception_addRef(info){info.add_ref()}var uncaughtExceptionCount=0;function ___cxa_begin_catch(ptr){var catchInfo=new CatchInfo(ptr);var info=catchInfo.get_exception_info();if(!info.get_caught()){info.set_caught(true);uncaughtExceptionCount--}info.set_rethrown(false);exceptionCaught.push(catchInfo);exception_addRef(info);return catchInfo.get_exception_ptr()}var exceptionLast=0;function ___cxa_free_exception(ptr){return _free(new ExceptionInfo(ptr).ptr)}function exception_decRef(info){if(info.release_ref()&&!info.get_rethrown()){var destructor=info.get_destructor();if(destructor){getWasmTableEntry(destructor)(info.excPtr)}___cxa_free_exception(info.excPtr)}}function ___cxa_end_catch(){_setThrew(0);var catchInfo=exceptionCaught.pop();exception_decRef(catchInfo.get_exception_info());catchInfo.free();exceptionLast=0}function ___resumeException(catchInfoPtr){var catchInfo=new CatchInfo(catchInfoPtr);var ptr=catchInfo.get_base_ptr();if(!exceptionLast){exceptionLast=ptr}catchInfo.free();throw ptr}function ___cxa_find_matching_catch_2(){var thrown=exceptionLast;if(!thrown){setTempRet0(0);return 0|0}var info=new ExceptionInfo(thrown);var thrownType=info.get_type();var catchInfo=new CatchInfo;catchInfo.set_base_ptr(thrown);catchInfo.set_adjusted_ptr(thrown);if(!thrownType){setTempRet0(0);return catchInfo.ptr|0}var typeArray=Array.prototype.slice.call(arguments);for(var i=0;i<typeArray.length;i++){var caughtType=typeArray[i];if(caughtType===0||caughtType===thrownType){break}if(___cxa_can_catch(caughtType,thrownType,catchInfo.get_adjusted_ptr_addr())){setTempRet0(caughtType);return catchInfo.ptr|0}}setTempRet0(thrownType);return catchInfo.ptr|0}function ___cxa_find_matching_catch_3(){var thrown=exceptionLast;if(!thrown){setTempRet0(0);return 0|0}var info=new ExceptionInfo(thrown);var thrownType=info.get_type();var catchInfo=new CatchInfo;catchInfo.set_base_ptr(thrown);catchInfo.set_adjusted_ptr(thrown);if(!thrownType){setTempRet0(0);return catchInfo.ptr|0}var typeArray=Array.prototype.slice.call(arguments);for(var i=0;i<typeArray.length;i++){var caughtType=typeArray[i];if(caughtType===0||caughtType===thrownType){break}if(___cxa_can_catch(caughtType,thrownType,catchInfo.get_adjusted_ptr_addr())){setTempRet0(caughtType);return catchInfo.ptr|0}}setTempRet0(thrownType);return catchInfo.ptr|0}function ___cxa_throw(ptr,type,destructor){var info=new ExceptionInfo(ptr);info.init(type,destructor);exceptionLast=ptr;uncaughtExceptionCount++;throw ptr}var SYSCALLS={buffers:[null,[],[]],printChar:function(stream,curr){var buffer=SYSCALLS.buffers[stream];if(curr===0||curr===10){(stream===1?out:err)(UTF8ArrayToString(buffer,0));buffer.length=0}else{buffer.push(curr)}},varargs:undefined,get:function(){SYSCALLS.varargs+=4;var ret=HEAP32[SYSCALLS.varargs-4>>2];return ret},getStr:function(ptr){var ret=UTF8ToString(ptr);return ret},get64:function(low,high){return low}};function ___syscall_fcntl64(fd,cmd,varargs){SYSCALLS.varargs=varargs;return 0}function ___syscall_ioctl(fd,op,varargs){SYSCALLS.varargs=varargs;return 0}function ___syscall_openat(dirfd,path,flags,varargs){SYSCALLS.varargs=varargs}function _abort(){abort("")}function _emscripten_memcpy_big(dest,src,num){HEAPU8.copyWithin(dest,src,src+num)}function _emscripten_get_heap_max(){return 2147483648}function emscripten_realloc_buffer(size){try{wasmMemory.grow(size-buffer.byteLength+65535>>>16);updateGlobalBufferAndViews(wasmMemory.buffer);return 1}catch(e){}}function _emscripten_resize_heap(requestedSize){var oldSize=HEAPU8.length;requestedSize=requestedSize>>>0;var maxHeapSize=_emscripten_get_heap_max();if(requestedSize>maxHeapSize){return false}let alignUp=(x,multiple)=>x+(multiple-x%multiple)%multiple;for(var cutDown=1;cutDown<=4;cutDown*=2){var overGrownHeapSize=oldSize*(1+.2/cutDown);overGrownHeapSize=Math.min(overGrownHeapSize,requestedSize+100663296);var newSize=Math.min(maxHeapSize,alignUp(Math.max(requestedSize,overGrownHeapSize),65536));var replacement=emscripten_realloc_buffer(newSize);if(replacement){return true}}return false}var ENV={};function getExecutableName(){return thisProgram||"./this.program"}function getEnvStrings(){if(!getEnvStrings.strings){var lang=(typeof navigator=="object"&&navigator.languages&&navigator.languages[0]||"C").replace("-","_")+".UTF-8";var env={"USER":"web_user","LOGNAME":"web_user","PATH":"/","PWD":"/","HOME":"/home/web_user","LANG":lang,"_":getExecutableName()};for(var x in ENV){if(ENV[x]===undefined)delete env[x];else env[x]=ENV[x]}var strings=[];for(var x in env){strings.push(x+"="+env[x])}getEnvStrings.strings=strings}return getEnvStrings.strings}function _environ_get(__environ,environ_buf){var bufSize=0;getEnvStrings().forEach(function(string,i){var ptr=environ_buf+bufSize;HEAP32[__environ+i*4>>2]=ptr;writeAsciiToMemory(string,ptr);bufSize+=string.length+1});return 0}function _environ_sizes_get(penviron_count,penviron_buf_size){var strings=getEnvStrings();HEAP32[penviron_count>>2]=strings.length;var bufSize=0;strings.forEach(function(string){bufSize+=string.length+1});HEAP32[penviron_buf_size>>2]=bufSize;return 0}function _fd_close(fd){return 0}function _fd_read(fd,iov,iovcnt,pnum){var stream=SYSCALLS.getStreamFromFD(fd);var num=SYSCALLS.doReadv(stream,iov,iovcnt);HEAP32[pnum>>2]=num;return 0}function _fd_seek(fd,offset_low,offset_high,whence,newOffset){}function _fd_write(fd,iov,iovcnt,pnum){var num=0;for(var i=0;i<iovcnt;i++){var ptr=HEAP32[iov>>2];var len=HEAP32[iov+4>>2];iov+=8;for(var j=0;j<len;j++){SYSCALLS.printChar(fd,HEAPU8[ptr+j])}num+=len}HEAP32[pnum>>2]=num;return 0}function _getTempRet0(){return getTempRet0()}function _llvm_eh_typeid_for(type){return type}var asmLibraryArg={"r":___cxa_allocate_exception,"o":___cxa_begin_catch,"s":___cxa_end_catch,"b":___cxa_find_matching_catch_2,"g":___cxa_find_matching_catch_3,"z":___cxa_free_exception,"q":___cxa_throw,"d":___resumeException,"w":___syscall_fcntl64,"E":___syscall_ioctl,"F":___syscall_openat,"p":_abort,"G":_emscripten_memcpy_big,"A":_emscripten_resize_heap,"B":_environ_get,"C":_environ_sizes_get,"u":_fd_close,"D":_fd_read,"y":_fd_seek,"v":_fd_write,"a":_getTempRet0,"j":invoke_ii,"e":invoke_iii,"l":invoke_iiii,"h":invoke_iiiii,"I":invoke_iiiiii,"m":invoke_iiiiiii,"x":invoke_iiiiiiiii,"J":invoke_v,"c":invoke_vi,"i":invoke_vii,"k":invoke_viii,"n":invoke_viiiffi,"f":invoke_viiii,"H":invoke_viiiiii,"t":_llvm_eh_typeid_for};var asm=createWasm();var ___wasm_call_ctors=Module["___wasm_call_ctors"]=function(){return(___wasm_call_ctors=Module["___wasm_call_ctors"]=Module["asm"]["L"]).apply(null,arguments)};var _FT_Load_Glyph=Module["_FT_Load_Glyph"]=function(){return(_FT_Load_Glyph=Module["_FT_Load_Glyph"]=Module["asm"]["M"]).apply(null,arguments)};var _FT_Set_Transform=Module["_FT_Set_Transform"]=function(){return(_FT_Set_Transform=Module["_FT_Set_Transform"]=Module["asm"]["O"]).apply(null,arguments)};var _FT_Done_Face=Module["_FT_Done_Face"]=function(){return(_FT_Done_Face=Module["_FT_Done_Face"]=Module["asm"]["P"]).apply(null,arguments)};var _FT_Set_Char_Size=Module["_FT_Set_Char_Size"]=function(){return(_FT_Set_Char_Size=Module["_FT_Set_Char_Size"]=Module["asm"]["Q"]).apply(null,arguments)};var _FT_Get_Glyph=Module["_FT_Get_Glyph"]=function(){return(_FT_Get_Glyph=Module["_FT_Get_Glyph"]=Module["asm"]["R"]).apply(null,arguments)};var _FT_Done_FreeType=Module["_FT_Done_FreeType"]=function(){return(_FT_Done_FreeType=Module["_FT_Done_FreeType"]=Module["asm"]["S"]).apply(null,arguments)};var _malloc=Module["_malloc"]=function(){return(_malloc=Module["_malloc"]=Module["asm"]["T"]).apply(null,arguments)};var _free=Module["_free"]=function(){return(_free=Module["_free"]=Module["asm"]["U"]).apply(null,arguments)};var _ASC_FT_Malloc=Module["_ASC_FT_Malloc"]=function(){return(_ASC_FT_Malloc=Module["_ASC_FT_Malloc"]=Module["asm"]["V"]).apply(null,arguments)};var _ASC_FT_Free=Module["_ASC_FT_Free"]=function(){return(_ASC_FT_Free=Module["_ASC_FT_Free"]=Module["asm"]["W"]).apply(null,arguments)};var _ASC_FT_Init=Module["_ASC_FT_Init"]=function(){return(_ASC_FT_Init=Module["_ASC_FT_Init"]=Module["asm"]["X"]).apply(null,arguments)};var _ASC_FT_Open_Face=Module["_ASC_FT_Open_Face"]=function(){return(_ASC_FT_Open_Face=Module["_ASC_FT_Open_Face"]=Module["asm"]["Y"]).apply(null,arguments)};var _ASC_FT_SetCMapForCharCode=Module["_ASC_FT_SetCMapForCharCode"]=function(){return(_ASC_FT_SetCMapForCharCode=Module["_ASC_FT_SetCMapForCharCode"]=Module["asm"]["Z"]).apply(null,arguments)};var _ASC_FT_GetFaceInfo=Module["_ASC_FT_GetFaceInfo"]=function(){return(_ASC_FT_GetFaceInfo=Module["_ASC_FT_GetFaceInfo"]=Module["asm"]["_"]).apply(null,arguments)};var _ASC_FT_GetFaceMaxAdvanceX=Module["_ASC_FT_GetFaceMaxAdvanceX"]=function(){return(_ASC_FT_GetFaceMaxAdvanceX=Module["_ASC_FT_GetFaceMaxAdvanceX"]=Module["asm"]["$"]).apply(null,arguments)};var _ASC_FT_GetKerningX=Module["_ASC_FT_GetKerningX"]=function(){return(_ASC_FT_GetKerningX=Module["_ASC_FT_GetKerningX"]=Module["asm"]["aa"]).apply(null,arguments)};var _ASC_FT_Glyph_Get_CBox=Module["_ASC_FT_Glyph_Get_CBox"]=function(){return(_ASC_FT_Glyph_Get_CBox=Module["_ASC_FT_Glyph_Get_CBox"]=Module["asm"]["ba"]).apply(null,arguments)};var _ASC_FT_Get_Glyph_Measure_Params=Module["_ASC_FT_Get_Glyph_Measure_Params"]=function(){return(_ASC_FT_Get_Glyph_Measure_Params=Module["_ASC_FT_Get_Glyph_Measure_Params"]=Module["asm"]["ca"]).apply(null,arguments)};var _ASC_FT_Get_Glyph_Render_Params=Module["_ASC_FT_Get_Glyph_Render_Params"]=function(){return(_ASC_FT_Get_Glyph_Render_Params=Module["_ASC_FT_Get_Glyph_Render_Params"]=Module["asm"]["da"]).apply(null,arguments)};var _ASC_FT_Get_Glyph_Render_Buffer=Module["_ASC_FT_Get_Glyph_Render_Buffer"]=function(){return(_ASC_FT_Get_Glyph_Render_Buffer=Module["_ASC_FT_Get_Glyph_Render_Buffer"]=Module["asm"]["ea"]).apply(null,arguments)};var _ASC_FT_Set_Transform=Module["_ASC_FT_Set_Transform"]=function(){return(_ASC_FT_Set_Transform=Module["_ASC_FT_Set_Transform"]=Module["asm"]["fa"]).apply(null,arguments)};var _ASC_FT_Set_TrueType_HintProp=Module["_ASC_FT_Set_TrueType_HintProp"]=function(){return(_ASC_FT_Set_TrueType_HintProp=Module["_ASC_FT_Set_TrueType_HintProp"]=Module["asm"]["ga"]).apply(null,arguments)};var _ASC_HB_LanguageFromString=Module["_ASC_HB_LanguageFromString"]=function(){return(_ASC_HB_LanguageFromString=Module["_ASC_HB_LanguageFromString"]=Module["asm"]["ha"]).apply(null,arguments)};var _ASC_HP_ShapeText=Module["_ASC_HP_ShapeText"]=function(){return(_ASC_HP_ShapeText=Module["_ASC_HP_ShapeText"]=Module["asm"]["ia"]).apply(null,arguments)};var _ASC_HP_FontFree=Module["_ASC_HP_FontFree"]=function(){return(_ASC_HP_FontFree=Module["_ASC_HP_FontFree"]=Module["asm"]["ja"]).apply(null,arguments)};var _setThrew=Module["_setThrew"]=function(){return(_setThrew=Module["_setThrew"]=Module["asm"]["ka"]).apply(null,arguments)};var stackSave=Module["stackSave"]=function(){return(stackSave=Module["stackSave"]=Module["asm"]["la"]).apply(null,arguments)};var stackRestore=Module["stackRestore"]=function(){return(stackRestore=Module["stackRestore"]=Module["asm"]["ma"]).apply(null,arguments)};var ___cxa_can_catch=Module["___cxa_can_catch"]=function(){return(___cxa_can_catch=Module["___cxa_can_catch"]=Module["asm"]["na"]).apply(null,arguments)};var ___cxa_is_pointer_type=Module["___cxa_is_pointer_type"]=function(){return(___cxa_is_pointer_type=Module["___cxa_is_pointer_type"]=Module["asm"]["oa"]).apply(null,arguments)};function invoke_v(index){var sp=stackSave();try{getWasmTableEntry(index)()}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiii(index,a1,a2,a3){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vii(index,a1,a2){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iii(index,a1,a2){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_vi(index,a1){var sp=stackSave();try{getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiii(index,a1,a2,a3,a4){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viii(index,a1,a2,a3){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_ii(index,a1){var sp=stackSave();try{return getWasmTableEntry(index)(a1)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiiiii(index,a1,a2,a3,a4,a5,a6,a7,a8){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6,a7,a8)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiii(index,a1,a2,a3,a4){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiii(index,a1,a2,a3,a4,a5){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiffi(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_iiiiiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{return getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}function invoke_viiiiii(index,a1,a2,a3,a4,a5,a6){var sp=stackSave();try{getWasmTableEntry(index)(a1,a2,a3,a4,a5,a6)}catch(e){stackRestore(sp);if(e!==e+0)throw e;_setThrew(1,0)}}var calledRun;dependenciesFulfilled=function runCaller(){if(!calledRun)run();if(!calledRun)dependenciesFulfilled=runCaller};function run(args){args=args||arguments_;if(runDependencies>0){return}preRun();if(runDependencies>0){return}function doRun(){if(calledRun)return;calledRun=true;Module["calledRun"]=true;if(ABORT)return;initRuntime();if(Module["onRuntimeInitialized"])Module["onRuntimeInitialized"]();postRun()}if(Module["setStatus"]){Module["setStatus"]("Running...");setTimeout(function(){setTimeout(function(){Module["setStatus"]("")},1);doRun()},1)}else{doRun()}}Module["run"]=run;if(Module["preInit"]){if(typeof Module["preInit"]=="function")Module["preInit"]=[Module["preInit"]];while(Module["preInit"].length>0){Module["preInit"].pop()()}}run();


Module.CreateLibrary = Module["_ASC_FT_Init"];
Module.FT_Set_TrueType_HintProp = Module["_ASC_FT_Set_TrueType_HintProp"];

Module.FT_Load_Glyph = Module["_FT_Load_Glyph"];
Module.FT_Set_Transform = Module["_ASC_FT_Set_Transform"];
Module.FT_Set_Char_Size = Module["_FT_Set_Char_Size"];

Module.FT_SetCMapForCharCode = Module["_ASC_FT_SetCMapForCharCode"];
Module.FT_GetKerningX = Module["_ASC_FT_GetKerningX"];
Module.FT_GetFaceMaxAdvanceX = Module["_ASC_FT_GetFaceMaxAdvanceX"];

Module.FT_Done_Face = Module["_FT_Done_Face"];
Module.HP_FontFree = Module["_ASC_HP_FontFree"];

Module.CreateNativeStream = function(typedArray)
{
	var fontStreamPointer = Module["_ASC_FT_Malloc"](typedArray.size);
	Module["HEAP8"].set(typedArray.data, fontStreamPointer);
	return { asc_marker: true, data: fontStreamPointer, len: typedArray.size};
};

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

Module.GetFaceInfo = function(face, reader)
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

Module.FT_Open_Face = function(library, memory, len, face_index)
{
	return Module["_ASC_FT_Open_Face"](library, memory, len, face_index);
};

Module.FT_Get_Glyph_Measure_Params = function(face, vector_worker, reader)
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

Module.FT_Get_Glyph_Render_Params = function(face, render_mode, reader)
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

AscFonts.FT_Get_Glyph_Render_Buffer = function(face, rasterInfo)
{
	var pointer = Module["_ASC_FT_Get_Glyph_Render_Buffer"](face);
	return new Uint8Array(Module["HEAP8"].buffer, pointer, rasterInfo.pitch * rasterInfo.rows);
};

Module.HP_ShapeText = function(fontFile, text, features, script, direction, language, reader)
{
	if (!Module.hb_cache_languages[language])
	{
		let langBuffer = language.toUtf8();
		var langPointer = Module["_malloc"](langBuffer.length);
		Module["HEAP8"].set(langBuffer, langBuffer);

		Module.hb_cache_languages[language] = langPointer;
	}

	let pointer = Module["_ASC_HP_ShapeText"](fontFile.m_pFace, fontFile.m_pHBFont, text, features, script, direction, AscFonts.hb_cache_languages[language]);
	if (!pointer)
	{
		g_return_obj_count.error = 1;
		return g_return_obj_count;
	}

	let buffer = Module["HEAP8"];
	let len = (buffer[pointer + 3] & 0xFF) << 24 | (buffer[pointer + 2] & 0xFF) << 16 | (buffer[pointer + 1] & 0xFF) << 8 | (buffer[pointer] & 0xFF);

	reader.init(buffer, pointer + 4, len - 4);
	fontFile.m_pHBFont = reader.readPointer64();

	g_return_obj_count.freeObj = pointer;
	g_return_obj_count.count = (len - 12) / 26;
	g_return_obj_count.error = 0;
	return g_return_obj_count;
};

// this memory is not freed
Module.hb_cache_languages = {};


	AscFonts.CreateLibrary = Module.CreateLibrary;
	AscFonts.FT_Set_TrueType_HintProp = Module.FT_Set_TrueType_HintProp;

	// Create stream from typed array
	// Module.CreateNativeStream(typed_array);
	AscFonts.CreateNativeStreamByIndex = function(stream_index)
	{
		let stream = AscFonts.g_fonts_streams[stream_index];
		if (stream && true !== stream.asc_marker)
		{
			AscFonts.g_fonts_streams[stream_index] = Module.CreateNativeStream(stream);
		}
	};

	function CBinaryReader(data, start, size)
	{
		this.data = data;
		this.pos = start;
		this.limit = start + size;
	}
	CBinaryReader.prototype.init = function(data, start, size)
	{
		this.data = data;
		this.pos = start;
		this.limit = start + size;
	}
	CBinaryReader.prototype.readInt = function()
	{
		var val = (this.data[this.pos] & 0xFF) | (this.data[this.pos + 1] & 0xFF) << 8 | (this.data[this.pos + 2] & 0xFF) << 16 | (this.data[this.pos + 3] & 0xFF) << 24;
		this.pos += 4;
		return val;
	};
	CBinaryReader.prototype.readUInt = function()
	{
		var val = (this.data[this.pos] & 0xFF) | (this.data[this.pos + 1] & 0xFF) << 8 | (this.data[this.pos + 2] & 0xFF) << 16 | (this.data[this.pos + 3] & 0xFF) << 24;
		this.pos += 4;
		return (val < 0) ? val + 4294967296 : val;
	};
	CBinaryReader.prototype.readByte = function()
	{
		return (this.data[this.pos++] & 0xFF);
	};
	CBinaryReader.prototype.readPointer64 = function()
	{
		let i1 = this.readUInt();
		let i2 = this.readUInt();
		if (i2 === 0)
			return i1;
		return i2 * 4294967296 + i1;
	};
	CBinaryReader.prototype.isValid = function()
	{
		return (this.pos < this.limit) ? true : false;
	};
	const READER = new CBinaryReader(null, 0, 0);

	function CFaceInfo()
	{
		this.units_per_EM = 0;
		this.ascender = 0;
		this.descender = 0;
		this.height = 0;
		this.face_flags = 0;
		this.num_faces = 0;
		this.num_glyphs = 0;
		this.num_charmaps = 0;
		this.style_flags = 0;
		this.face_index = 0;

		this.family_name = "";

		this.style_name = "";

		this.os2_version = 0;
		this.os2_usWeightClass = 0;
		this.os2_fsSelection = 0;
		this.os2_usWinAscent = 0;
		this.os2_usWinDescent = 0;
		this.os2_usDefaultChar = 0;
		this.os2_sTypoAscender = 0;
		this.os2_sTypoDescender = 0;
		this.os2_sTypoLineGap = 0;

		this.os2_ulUnicodeRange1 = 0;
		this.os2_ulUnicodeRange2 = 0;
		this.os2_ulUnicodeRange3 = 0;
		this.os2_ulUnicodeRange4 = 0;
		this.os2_ulCodePageRange1 = 0;
		this.os2_ulCodePageRange2 = 0;

		this.os2_nSymbolic = 0;

		this.header_yMin = 0;
		this.header_yMax = 0;

		this.monochromeSizes = [];
	}

	CFaceInfo.prototype.load = function(face)
	{
		let errorObj = Module.GetFaceInfo(face, READER);
		if (errorObj.error)
			return;

		this.units_per_EM 	= READER.readUInt();
		this.ascender 		= READER.readInt();
		this.descender 		= READER.readInt();
		this.height 		= READER.readInt();
		this.face_flags 	= READER.readInt();
		this.num_faces 		= READER.readInt();
		this.num_glyphs 	= READER.readInt();
		this.num_charmaps 	= READER.readInt();
		this.style_flags 	= READER.readInt();
		this.face_index 	= READER.readInt();

		var c = READER.readInt();
		while (c)
		{
			this.family_name += String.fromCharCode(c);
			c = READER.readInt();
		}

		c = READER.readInt();
		while (c)
		{
			this.style_name += String.fromCharCode(c);
			c = READER.readInt();
		}

		this.os2_version 		= READER.readInt();
		this.os2_usWeightClass 	= READER.readInt();
		this.os2_fsSelection 	= READER.readInt();
		this.os2_usWinAscent 	= READER.readInt();
		this.os2_usWinDescent 	= READER.readInt();
		this.os2_usDefaultChar 	= READER.readInt();
		this.os2_sTypoAscender 	= READER.readInt();
		this.os2_sTypoDescender = READER.readInt();
		this.os2_sTypoLineGap 	= READER.readInt();

		this.os2_ulUnicodeRange1 	= READER.readUInt();
		this.os2_ulUnicodeRange2 	= READER.readUInt();
		this.os2_ulUnicodeRange3 	= READER.readUInt();
		this.os2_ulUnicodeRange4 	= READER.readUInt();
		this.os2_ulCodePageRange1 	= READER.readUInt();
		this.os2_ulCodePageRange2 	= READER.readUInt();

		this.os2_nSymbolic 			= READER.readInt();
		this.header_yMin 			= READER.readInt();
		this.header_yMax 			= READER.readInt();

		var fixedSizesCount = READER.readInt();
		for (var i = 0; i < fixedSizesCount; i++)
			this.monochromeSizes.push(READER.readInt());

		errorObj.free();
	};

	function CGlyphMetrics()
	{
		this.bbox_xMin = 0;
		this.bbox_yMin = 0;
		this.bbox_xMax = 0;
		this.bbox_yMax = 0;

		this.width          = 0;
		this.height         = 0;

		this.horiAdvance    = 0;
		this.horiBearingX   = 0;
		this.horiBearingY   = 0;

		this.vertAdvance    = 0;
		this.vertBearingX   = 0;
		this.vertBearingY   = 0;

		this.linearHoriAdvance = 0;
		this.linearVertAdvance = 0;
	}

	function CGlyphBitmapImage()
	{
		this.left   = 0;
		this.top    = 0;
		this.width  = 0;
		this.rows   = 0;
		this.pitch  = 0;
		this.mode   = 0;
	}

	AscFonts.CFaceInfo = CFaceInfo;
	AscFonts.CGlyphMetrics = CGlyphMetrics;
	AscFonts.CGlyphBitmapImage = CGlyphBitmapImage;

	AscFonts.FT_Open_Face = function(library, stream, face_index)
	{
		return Module.FT_Open_Face(library, stream.data, stream.len, face_index);
	};

	AscFonts.FT_Glyph_Get_Measure = function(face, vector_worker, painter)
	{
		let errorObj = Module.FT_Get_Glyph_Measure_Params(face, vector_worker ? 1 : 0, READER);
		if (errorObj.error)
			return null;

		let len = errorObj.count;

		var info = new CGlyphMetrics();
		info.bbox_xMin     = READER.readInt();
		info.bbox_yMin     = READER.readInt();
		info.bbox_xMax     = READER.readInt();
		info.bbox_yMax     = READER.readInt();

		info.width         = READER.readInt();
		info.height        = READER.readInt();

		info.horiAdvance   = READER.readInt();
		info.horiBearingX  = READER.readInt();
		info.horiBearingY  = READER.readInt();

		info.vertAdvance   = READER.readInt();
		info.vertBearingX  = READER.readInt();
		info.vertBearingY  = READER.readInt();

		info.linearHoriAdvance     = READER.readInt();
		info.linearVertAdvance     = READER.readInt();

		if (vector_worker)
		{
			painter.start(vector_worker);

			let pos = 15;
			while (pos < len)
			{
				switch (READER.readInt())
				{
					case 0:
					{
						painter._move_to(READER.readInt(), READER.readInt(), vector_worker);
						break;
					}
					case 1:
					{
						painter._line_to(READER.readInt(), READER.readInt(), vector_worker);
						break;
					}
					case 2:
					{
						painter._conic_to(READER.readInt(), READER.readInt(), READER.readInt(), READER.readInt(), vector_worker);
						break;
					}
					case 3:
					{
						painter._cubic_to(READER.readInt(), READER.readInt(), READER.readInt(), READER.readInt(), READER.readInt(), READER.readInt(), vector_worker);
						break;
					}
					default:
						break;
				}
			}

			painter.end(vector_worker);
		}

		errorObj.free();
		return info;
	};

	AscFonts.FT_Glyph_Get_Raster = function(face, render_mode)
	{
		let errorObj = Module.FT_Get_Glyph_Render_Params(face, render_mode, READER);
		if (errorObj.error)
			return null;

		var info = new CGlyphBitmapImage();
		info.left    = READER.readInt();
		info.top     = READER.readInt();
		info.width   = READER.readInt();
		info.rows    = READER.readInt();
		info.pitch   = READER.readInt();
		info.mode    = READER.readInt();

		errorObj.free();
		return info;
	};

	AscFonts.FT_Load_Glyph = Module.FT_Load_Glyph;
	AscFonts.FT_Set_Transform = Module.FT_Set_Transform;
	AscFonts.FT_Set_Char_Size = Module.FT_Set_Char_Size;

	AscFonts.FT_SetCMapForCharCode = Module.FT_SetCMapForCharCode;
	AscFonts.FT_GetKerningX = Module.FT_GetKerningX;
	AscFonts.FT_GetFaceMaxAdvanceX = Module.FT_GetFaceMaxAdvanceX;

	AscFonts.FT_Done_Face = Module.FT_Done_Face;
	AscFonts.HP_FontFree = Module.HP_FontFree;

	AscFonts.FT_Get_Glyph_Render_Buffer = function(face, rasterInfo, isCopyToRasterMemory)
	{
		let buffer = Module.FT_Get_Glyph_Render_Buffer(face, rasterInfo);

		if (!isCopyToRasterMemory)
			return buffer;

		AscFonts.raster_memory.CheckSize(rasterInfo.width, rasterInfo.rows);

		let offsetSrc = 0;
		let offsetDst = 3;
		let dstData = AscFonts.raster_memory.m_oBuffer.data;

		if (rasterInfo.pitch >= rasterInfo.width)
		{
			for (let j = 0; j < rasterInfo.rows; ++j, offsetSrc += rasterInfo.pitch)
			{
				offsetDst = 3 + j * AscFonts.raster_memory.pitch;
				for (let i = 0; i < rasterInfo.width; i++, offsetDst += 4)
				{
					dstData[offsetDst] = buffer[offsetSrc + i];
				}
			}
		}
		else
		{
			var bitNumber = 0;
			var byteNumber = 0;
			for (let j = 0; j < rasterInfo.rows; ++j, offsetSrc += rasterInfo.pitch)
			{
				offsetDst = 3 + j * AscFonts.raster_memory.pitch;
				bitNumber = 0;
				byteNumber = 0;
				for (let i = 0; i < rasterInfo.width; i++, offsetDst += 4, bitNumber++)
				{
					if (8 === bitNumber)
					{
						bitNumber = 0;
						byteNumber++;
					}
					dstData[offsetDst] = (buffer[offsetSrc + byteNumber] & (1 << (7 - bitNumber))) ? 255 : 0;
				}
			}
		}
	};

	const STRING_MAX_LEN = AscFonts.GRAPHEME_STRING_MAX_LEN;
	const COEF           = AscFonts.GRAPHEME_COEF;
	let   STRING_POINTER = null;
	let   STRING_LEN     = 0;
	const CLUSTER        = new Uint8Array(STRING_MAX_LEN);
	let   CLUSTER_LEN    = 0;
	let   CLUSTER_MAX    = 0;
	const LIGATURE       = 2;

	function CClusterUtf8Calculator()
	{
		this.currentCluster = 0;
		this.currentCodePoint = 0;
	}
	CClusterUtf8Calculator.prototype.start = function()
	{
		this.currentCluster = 0;
		this.currentCodePoint = 0;
	}
	CClusterUtf8Calculator.prototype.getCodePointsCount = function(cluster)
	{
		let nCodePointsCount = 0;

		if (cluster > this.currentCluster)
		{
			// TODO: RTL
		}
		else
		{
			while (this.currentCluster < cluster)
			{
				this.currentCluster += CLUSTER[this.currentCodePoint + nCodePointsCount];
				nCodePointsCount++;
			}
		}

		this.currentCodePoint += nCodePointsCount;
		return nCodePointsCount;
	}
	const CODEPOINTS_CALCULATOR = new CClusterUtf8Calculator();

	AscFonts.HB_StartString = function()
	{
		if (!STRING_POINTER)
			STRING_POINTER = Module.AllocString(6 * STRING_MAX_LEN + 1);

		STRING_LEN  = 0;
		CLUSTER_LEN = 0;
		CLUSTER_MAX = 0;
	};
	AscFonts.HB_AppendToString = function(code)
	{
		let arrBuffer   = STRING_POINTER.buffer;
		let nClusterLen = -1;

		if (code < 0x80)
		{
			arrBuffer[STRING_LEN++] = code;
			nClusterLen = 1;
		}
		else if (code < 0x0800)
		{
			arrBuffer[STRING_LEN++] = (0xC0 | (code >> 6));
			arrBuffer[STRING_LEN++] = (0x80 | (code & 0x3F));
			nClusterLen = 2;
		}
		else if (code < 0x10000)
		{
			arrBuffer[STRING_LEN++] = (0xE0 | (code >> 12));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 6) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | (code & 0x3F));
			nClusterLen = 3;
		}
		else if (code < 0x1FFFFF)
		{
			arrBuffer[STRING_LEN++] = (0xF0 | (code >> 18));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 12) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 6) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | (code & 0x3F));
			nClusterLen = 4;
		}
		else if (code < 0x3FFFFFF)
		{
			arrBuffer[STRING_LEN++] = (0xF8 | (code >> 24));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 18) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 12) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 6) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | (code & 0x3F));
			nClusterLen = 5;
		}
		else if (code < 0x7FFFFFFF)
		{
			arrBuffer[STRING_LEN++] = (0xFC | (code >> 30));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 24) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 18) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 12) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | ((code >> 6) & 0x3F));
			arrBuffer[STRING_LEN++] = (0x80 | (code & 0x3F));
			nClusterLen = 6;
		}

		if (-1 !== nClusterLen)
		{
			CLUSTER[CLUSTER_LEN++] = nClusterLen;
			CLUSTER_MAX += nClusterLen;
		}
	};
	AscFonts.HB_EndString = function()
	{
		STRING_POINTER.buffer[STRING_LEN++] = 0;
	};
	AscFonts.HB_GetStringLength = function()
	{
		return STRING_LEN;
	};
	AscFonts.HB_ShapeString = function(textShaper, fontId, fontStyle, fontFile, features, script, direction, language)
	{
		if (!STRING_POINTER)
			return;

		let retObj = Module.HP_ShapeText(fontFile, STRING_POINTER, features, script, direction, language, READER);
		if (retObj.error)
			return;

		CODEPOINTS_CALCULATOR.start();
		let prevCluster = -1, type, flags, gid, cluster, x_advance, y_advance, x_offset, y_offset;
		let isLigature = false;
		let nWidth     = 0;
		let reader = READER;
		let glyphsCount = retObj.count;
		for (let i = 0; i < glyphsCount; i++)
		{
			type      = reader.readByte();
			flags     = reader.readByte();
			gid       = reader.readInt();
			cluster   = reader.readInt();
			x_advance = reader.readInt();
			y_advance = reader.readInt();
			x_offset  = reader.readInt();
			y_offset  = reader.readInt();

			if (cluster !== prevCluster && -1 !== prevCluster)
			{
				textShaper.FlushGrapheme(AscFonts.GetGrapheme(), nWidth, CODEPOINTS_CALCULATOR.getCodePointsCount(cluster), isLigature);
				nWidth = 0;
			}

			if (cluster !== prevCluster)
			{
				prevCluster = cluster;
				isLigature  = LIGATURE === type;
				AscFonts.InitGrapheme(fontId, fontStyle);
			}

			AscFonts.AddGlyphToGrapheme(gid, x_advance, y_advance, x_offset, y_offset);
			nWidth += x_advance * COEF;
		}
		textShaper.FlushGrapheme(AscFonts.GetGrapheme(), nWidth, CODEPOINTS_CALCULATOR.getCodePointsCount(CLUSTER_MAX), isLigature);

		retObj.free();
	};

	AscFonts.HB_Shape = function(fontFile, text, features, script, direction, language)
	{
		if (text.length === 0)
			return null;

		AscFonts.HB_StartString();
		for (let iter = text.getUnicodeIterator(); iter.check(); iter.next())
		{
			AscFonts.HB_AppendToString(iter.value());
		}
		AscFonts.HB_EndString();

		let retObj = Module.HP_ShapeText(fontFile, STRING_POINTER, features, script, direction, language, READER);
		if (retObj.error)
			return;

		let glyphs = [];
		let glyphsCount = retObj.count;
		for (let i = 0; i < glyphsCount; i++)
		{
			let glyph = {};

			glyph.type = reader.readByte();
			glyph.flags = reader.readByte();

			glyph.gid = reader.readInt();
			glyph.cluster = reader.readInt();

			glyph.x_advance = reader.readInt();
			glyph.y_advance = reader.readInt();
			glyph.x_offset = reader.readInt();
			glyph.y_offset = reader.readInt();

			glyphs.push(glyph);
		}

		let utf8_start = 0;
		let utf8_end = 0;

		if (direction === AscFonts.HB_DIRECTION.HB_DIRECTION_RTL)
		{
			for (let i = glyphsCount - 1; i >= 0; i--)
			{
				if (i === 0)
					utf8_end = textBuffer.length - 1;
				else
					utf8_end = glyphs[i - 1].cluster;

				glyphs[i].text = String.prototype.fromUtf8(STRING_POINTER.buffer, utf8_start, utf8_end - utf8_start);
				utf8_start = utf8_end;
			}
		}
		else
		{
			for (let i = 0; i < glyphsCount; i++)
			{
				if (i === (glyphsCount - 1))
					utf8_end = textBuffer.length - 1;
				else
					utf8_end = glyphs[i + 1].cluster;

				glyphs[i].text = String.prototype.fromUtf8(STRING_POINTER.buffer, utf8_start, utf8_end - utf8_start);
				utf8_start = utf8_end;
			}
		}

		retObj.free();
		return glyphs;
	};

	AscFonts.onLoadModule();

})(window, undefined);
