/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/app.js");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/app.js":
/*!********************!*\
  !*** ./src/app.js ***!
  \********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
eval("/* global util:true, fabric:true, Office:true, OfficeExtension:true, Word:true */\n\n\n\n// load appUtilities module using commonJS syntax\n\nvar util = __webpack_require__(/*! ./appUtilities.js */ \"./src/appUtilities.js\");\n\n(function () {\n\tvar messageBanner;\n\t// var allRangeLength = 0;\n\n\tOffice.initialize = function () {\n\t\t$(document).ready(function () {\n\t\t\t// initialize FabricUI notification mechanism and hide it\n\t\t\tvar element = document.querySelector('.ms-MessageBanner');\n\t\t\tmessageBanner = new fabric.MessageBanner(element);\n\t\t\tmessageBanner.hideBanner();\n\n\t\t\t// check Office\n\t\t\tif (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {\n\t\t\t\tconsole.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');\n\t\t\t}\n\n\t\t\tvar docx = Office.context.document;\n\n\t\t\t// pull into 'live settings' the data (if any) that is stored in the file\n\t\t\tdocx.settings.refreshAsync(function () {\n\t\t\t\t// get userTerms from live settings and show them in ui\n\t\t\t\t['add', 'minus'].forEach(function (cmd) {\n\t\t\t\t\taddToShownUserTerms(cmd, docx.settings.get('userTerms-' + cmd) || []);\n\t\t\t\t});\n\t\t\t});\n\n\t\t\t$('#user-term-add').on('keydown', function (e) {\n\t\t\t\tif (e.keyCode === 13) {\n\t\t\t\t\tkeydownHandler('add', $(this));\n\t\t\t\t}\n\t\t\t});\n\t\t\t$('#user-term-minus').on('keydown', function (e) {\n\t\t\t\tif (e.keyCode === 13) {\n\t\t\t\t\tkeydownHandler('minus', $(this));\n\t\t\t\t}\n\t\t\t});\n\n\t\t\t$('#user-terms-add-container').on('click', '.user-term', function () {\n\t\t\t\tremoveClickHandler('add', $(this));\n\t\t\t});\n\t\t\t$('#user-terms-minus-container').on('click', '.user-term', function () {\n\t\t\t\tremoveClickHandler('minus', $(this));\n\t\t\t});\n\n\t\t\t$('#select-btn').on('click', selectDefParas);\n\t\t\t$('#select-btn-text').text('Select Definition Paragraphs');\n\n\t\t\t$('#parse-btn').on('click', parseVocabTerms);\n\t\t\t$('#parse-btn-text').text('Parse Selected');\n\t\t});\n\t};\n\n\t/* UI Functions */\n\tfunction keydownHandler(cmd, elm) {\n\t\tvar inpVal = elm.val().trim();\n\n\t\tif (!inpVal) {\n\t\t\treturn; //bail\n\t\t}\n\n\t\t// add to shown user terms if not a dupe\n\t\tif (getShownUserTerms(cmd).indexOf(inpVal) === -1) {\n\t\t\taddToShownUserTerms(cmd, [inpVal]);\n\t\t\telm.val(''); //clear input\n\t\t}\n\n\t\t// sync to settings if not a dupe\n\t\tvar docx = Office.context.document;\n\t\tvar userTerms = docx.settings.get('userTerms-' + cmd) || [];\n\t\tif (userTerms.indexOf(inpVal) === -1) {\n\t\t\tuserTerms.push(inpVal);\n\t\t\tuserTerms.sort(util.sortByAlphabet);\n\t\t\tdocx.settings.set('userTerms-' + cmd, userTerms);\n\t\t\tdocx.settings.saveAsync();\n\t\t}\n\t}\n\n\tfunction removeClickHandler(cmd, elm) {\n\t\tvar val = elm.text();\n\t\telm.remove();\n\n\t\t// sync to settings\n\t\tvar docx = Office.context.document;\n\t\tvar userTerms = docx.settings.get('userTerms-' + cmd);\n\t\tif (userTerms) {\n\t\t\tuserTerms.splice(userTerms.indexOf(val), 1);\n\t\t\tdocx.settings.set('userTerms-' + cmd, userTerms);\n\t\t\tdocx.settings.saveAsync();\n\t\t}\n\t}\n\n\tfunction getShownUserTerms(cmd) {\n\t\tvar userTerms = [];\n\n\t\t$('#user-terms-' + cmd + '-container .user-term').each(function () {\n\t\t\tuserTerms.push($(this).text());\n\t\t});\n\n\t\treturn userTerms;\n\t}\n\n\tfunction addToShownUserTerms(cmd, arrayOfTerms) {\n\t\tvar container = $('#user-terms-' + cmd + '-container');\n\t\tvar frag = document.createDocumentFragment();\n\n\t\tarrayOfTerms.forEach(function (term) {\n\t\t\tvar div = document.createElement('div');\n\t\t\tdiv.classList.add('user-term');\n\t\t\tdiv.textContent = term;\n\t\t\tfrag.appendChild(div);\n\t\t});\n\t\tcontainer.prepend(frag);\n\n\t\treturn container;\n\t}\n\n\tfunction showNotification(header, content) {\n\t\t$(\"#notification-header\").text(header);\n\t\t$(\"#notification-body\").text(content);\n\t\tmessageBanner.showBanner();\n\t\tmessageBanner.toggleExpansion();\n\t}\n\n\tfunction errHandler(error) {\n\t\tconsole.log(\"Error: \" + error);\n\n\t\tif (error instanceof OfficeExtension.Error) {\n\t\t\tconsole.log(\"Debug info: \" + JSON.stringify(error.debugInfo));\n\t\t} else if (/TypeError: Unable to get property 'getRange'/.test(error)) {\n\t\t\tvar header = 'Error:';\n\t\t\tvar content = 'There are no definition paragraphs to select';\n\t\t\tshowNotification(header, content);\n\t\t}\n\t}\n\n\t/* Operative Functions */\n\t/* function selectAll() {\r\n \tWord.run(function (context) {\r\n \t\t// queue command to select whole doc\r\n \t\tcontext.document.body.select();\r\n \n \t\t// queue command to load/return all the paragraphs as a range\r\n \t\tvar allRange = context.document.body.paragraphs;\r\n \t\tcontext.load(allRange, 'text');\r\n \n \t\treturn context.sync().then(function () {\r\n \t\t\t// if successful, store allRange.items.length in global var\r\n \t\t\tallRangeLength = allRange.items.length;\r\n \t\t\tconsole.log('allRangeLength', allRangeLength);\r\n \t\t});\r\n \t})\r\n \t.catch(errHandler);\r\n } */\n\n\tfunction bifurcateParas(paras) {\n\t\tvar rexqtBeginning = /(^|(\\(\\w{1,3}\\)\\s+?))“[^”]+”/;\n\n\t\tvar startIndex = paras.findIndex(function (p) {\n\t\t\treturn rexqtBeginning.test(p);\n\t\t});\n\n\t\tvar revStartIndex = paras.slice(0).reverse().findIndex(function (p) {\n\t\t\treturn rexqtBeginning.test(p);\n\t\t});\n\t\tvar endIndex = paras.length - (revStartIndex + 1);\n\n\t\t/* let defParas = paras\r\n  \t.filter(function (p, i) {\r\n  \t\treturn i >= startIndex && i <= endIndex;\r\n  \t});\r\n  \t\tlet plainParas = paras\r\n  \t.filter(function (p, i) {\r\n  \t\treturn i < startIndex || i > endIndex;\r\n  \t});\r\n  \t\treturn [defParas, plainParas]; */\n\n\t\treturn [startIndex, endIndex];\n\t}\n\n\tfunction selectDefParas() {\n\t\tWord.run(function (context) {\n\t\t\t// queue command to load/return all the paragraphs as a range\n\t\t\tvar allRange = context.document.body.paragraphs;\n\t\t\tcontext.load(allRange, 'text');\n\n\t\t\treturn context.sync().then(function () {\n\t\t\t\tvar allParas = allRange.items.map(function (p) {\n\t\t\t\t\treturn p.text.trim();\n\t\t\t\t});\n\n\t\t\t\tvar indices = bifurcateParas(allParas);\n\t\t\t\tvar startIndex = indices[0];\n\t\t\t\tvar endIndex = indices[1];\n\t\t\t\tvar startRange = allRange.items[startIndex].getRange();\n\t\t\t\tvar endRange = allRange.items[endIndex].getRange();\n\n\t\t\t\tvar expandedRange = endRange.expandTo(startRange);\n\t\t\t\texpandedRange.select();\n\n\t\t\t\treturn context.sync();\n\t\t\t});\n\t\t}).catch(errHandler);\n\t}\n\n\t/* function getCrossRefDefs(paras) {\r\n \tconst rexFirstSentence = /^.+?\\.(?:\\s|$)/;\r\n \treturn paras\r\n \t\t.map(function (p) {\r\n \t\t\treturn p.match(rexFirstSentence);\r\n \t\t})\r\n \t\t.reduce(function (accumArray, matchArray) {\r\n \t\t\treturn accumArray.concat(matchArray); //flatten into a single array of strings\r\n \t\t}, [])\r\n \t\t.filter(function (sentence) {\r\n \t\t\treturn /\\b(meaning|defined|definition)s*?\\b/.test(sentence);\r\n \t\t})\r\n \t\t.filter(function (sentence) {\r\n \t\t\treturn /^“/.test(sentence);\r\n \t\t})\r\n \t\t.filter(function (sentence) {\r\n \t\t\treturn sentence[0].split(' ').length < 30;\r\n \t\t});\r\n } */\n\n\tfunction parseVocabTerms() {\n\t\tWord.run(function (context) {\n\t\t\t// queue command to load/return all the paragraphs as a range\n\t\t\tvar allRange = context.document.body.paragraphs;\n\t\t\tcontext.load(allRange, 'text');\n\n\t\t\treturn context.sync().then(function () {\n\t\t\t\tvar paras = allRange.items.map(function (p) {\n\t\t\t\t\treturn p.text.trim();\n\t\t\t\t}).filter(function (p) {\n\t\t\t\t\treturn p; //filter out empty items in array\n\t\t\t\t});\n\t\t\t\tconsole.log('paras', paras);\n\n\t\t\t\t// check agst global var to confirm that whole doc is still selected\n\t\t\t\t/* if (paras.length === allRangeLength) {\r\n    \t// if so, trim paragraph collection (in place) from the end\r\n    \tlet revLastIndex = paras.slice(0).reverse()\r\n    \t\t.findIndex(function (item) {\r\n    \t\t\treturn /^“[^”]+”/.test(item);\r\n    \t\t});\r\n    \tparas.splice((revLastIndex * -1));\r\n    \tconsole.log('SPLICED PARAS', paras);\r\n    \t\t} else {\r\n    \t// otherwise, reset global var and don't trim paragraph collection\r\n    \tallRangeLength = 0;\r\n    } */\n\n\t\t\t\t/* START HERE */\n\t\t\t\tvar pojo = Object.create(null);\n\t\t\t\tvar lastTerm;\n\n\t\t\t\tparas.forEach(function (p) {\n\t\t\t\t\tif (!/^\\*/.test(p)) {\n\t\t\t\t\t\tvar arr = p.split('\\t');\n\t\t\t\t\t\tconsole.log('arr', arr);\n\n\t\t\t\t\t\t//set 'term' for this para and subsequent SYNONYM/ANTONYM paras\n\t\t\t\t\t\tvar term = lastTerm = arr[0];\n\n\t\t\t\t\t\t//create term object within pojo\n\t\t\t\t\t\tpojo[term] = Object.create(null);\n\n\t\t\t\t\t\t//add definition thereto\n\t\t\t\t\t\tpojo[term].def = arr[1];\n\t\t\t\t\t} else if (/SYNONYMS/.test(p)) {\n\t\t\t\t\t\tvar synos = p.replace('*SYNONYMS:*', '').trim();\n\t\t\t\t\t\tconsole.log('synos', synos);\n\n\t\t\t\t\t\tpojo[lastTerm].synos = synos;\n\t\t\t\t\t} else if (/ANTONYMS/.test(p)) {\n\t\t\t\t\t\tvar antos = p.replace('*ANTONYMS:*', '').trim();\n\t\t\t\t\t\tconsole.log('antos', antos);\n\n\t\t\t\t\t\tpojo[lastTerm].antos = antos;\n\t\t\t\t\t} else {\n\t\t\t\t\t\tconsole.log('error parsing empty para');\n\t\t\t\t\t}\n\t\t\t\t});\n\t\t\t\tlastTerm = '';\n\n\t\t\t\tvar sortedPojo = util.sortObject(pojo, util.sortByAlphabet);\n\t\t\t\tconsole.log('debug sortedPojo', sortedPojo);\n\t\t\t\t/* END HERE */\n\n\t\t\t\t// Throw error if pojo is empty\n\t\t\t\tif (!Object.keys(sortedPojo).length) {\n\t\t\t\t\tvar header = 'Error:';\n\t\t\t\t\tvar content = 'No definition paragraphs have been selected';\n\t\t\t\t\tshowNotification(header, content);\n\n\t\t\t\t\treturn context.sync(); //bail\n\t\t\t\t}\n\n\t\t\t\t// Create master array of individual term tables\n\t\t\t\tvar masterTableArray = [];\n\n\t\t\t\tObject.keys(sortedPojo).forEach(function (term) {\n\t\t\t\t\tvar termTableArray = [];\n\t\t\t\t\tvar termObj = sortedPojo[term];\n\n\t\t\t\t\t//populate termTableArray\n\t\t\t\t\ttermTableArray.push([term]);\n\t\t\t\t\ttermTableArray.push([termObj.def]);\n\n\t\t\t\t\tif (termObj.synos) {\n\t\t\t\t\t\ttermTableArray.push(['SYNONYMS', termObj.synos]);\n\t\t\t\t\t}\n\n\t\t\t\t\tif (termObj.antos) {\n\t\t\t\t\t\ttermTableArray.push(['ANTONYMS', termObj.antos]);\n\t\t\t\t\t}\n\n\t\t\t\t\t//push termTableArray into masterTableArray\n\t\t\t\t\tmasterTableArray.push(termTableArray);\n\t\t\t\t});\n\n\t\t\t\tvar newDoc = context.application.createDocument();\n\t\t\t\tcontext.load(newDoc);\n\n\t\t\t\treturn context.sync().then(function () {\n\t\t\t\t\t// console.log('newDoc', newDoc);\n\n\t\t\t\t\tmasterTableArray.forEach(function (termTableArray) {\n\t\t\t\t\t\tvar table = util.insertTable(newDoc.body, termTableArray);\n\t\t\t\t\t\ttable.headerRowCount = 1;\n\t\t\t\t\t\ttable.style = 'List Table 4 - Accent 1';\n\t\t\t\t\t\ttable.styleFirstColumn = false;\n\t\t\t\t\t});\n\n\t\t\t\t\treturn context.sync().then(function () {\n\t\t\t\t\t\tnewDoc.open();\n\n\t\t\t\t\t\treturn context.sync();\n\t\t\t\t\t});\n\t\t\t\t});\n\t\t\t});\n\t\t}).catch(errHandler);\n\t}\n})();\n\n//# sourceURL=webpack:///./src/app.js?");

/***/ }),

/***/ "./src/appUtilities.js":
/*!*****************************!*\
  !*** ./src/appUtilities.js ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
eval("/* global Word:true */\n\n\n\n// Uses revealing module pattern to return an object consisting of exposed methods\n\nvar _typeof = typeof Symbol === \"function\" && typeof Symbol.iterator === \"symbol\" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === \"function\" && obj.constructor === Symbol && obj !== Symbol.prototype ? \"symbol\" : typeof obj; };\n\nvar util = function () {\n\tfunction createRexFromString(string, flags) {\n\t\tvar escapedString = string.replace(/[|\\\\{}()[\\]^$+*?.]/g, '\\\\$&');\n\t\treturn new RegExp(escapedString, flags);\n\t}\n\n\tfunction sortByAlphabet(A, B) {\n\t\tvar a = A.toLowerCase();\n\t\tvar b = B.toLowerCase();\n\n\t\tif (a < b) {\n\t\t\treturn -1;\n\t\t}\n\t\tif (a > b) {\n\t\t\treturn 1;\n\t\t}\n\t\treturn 0; //default return value (no sorting)\n\t}\n\n\tfunction sortByLongerLength(A, B) {\n\t\tvar a = A.length;\n\t\tvar b = B.length;\n\n\t\tif (a > b) {\n\t\t\treturn -1;\n\t\t}\n\t\tif (a < b) {\n\t\t\treturn 1;\n\t\t}\n\t\treturn 0; //default return value (no sorting)\n\t}\n\n\tfunction sortObject(src, comparator) {\n\t\tvar out = Object.create(null);\n\n\t\tObject.keys(src).sort(comparator).forEach(function (key) {\n\t\t\tif (_typeof(src[key]) == 'object' && !Array.isArray(src[key]) && !(src[key] instanceof RegExp)) {\n\t\t\t\tout[key] = sortObject(src[key], comparator); //run function again\n\t\t\t\treturn;\n\t\t\t} else {\n\t\t\t\tout[key] = src[key];\n\t\t\t}\n\t\t});\n\n\t\treturn out;\n\t}\n\n\tfunction mergeObjects(target, src) {\n\t\tvar a = target || Object.create(null);\n\t\tvar b = src || Object.create(null);\n\n\t\t// merge b into a\n\t\tObject.keys(b).forEach(function (key) {\n\t\t\ta[key] = (a[key] || 0) + (b[key] || 0);\n\t\t});\n\t}\n\n\tfunction mergeWithinObject(termObj, wordPair) {\n\t\tvar retainWord = wordPair[0];\n\t\tvar loseWord = wordPair[1];\n\n\t\tObject.keys(termObj).forEach(function (mainKey) {\n\t\t\tif (mainKey !== 'defined') {\n\t\t\t\tvar subObject = termObj[mainKey]; //can be either 'incorps' or 'usedBy' object\n\n\t\t\t\tObject.keys(subObject).forEach(function (word) {\n\t\t\t\t\tif (word === loseWord) {\n\t\t\t\t\t\tsubObject[retainWord] = (subObject[retainWord] || 0) + subObject[word];\n\t\t\t\t\t\tdelete subObject[word];\n\t\t\t\t\t}\n\t\t\t\t});\n\t\t\t}\n\t\t});\n\t}\n\n\tfunction addBullet(strOrObj) {\n\t\tvar string = (typeof strOrObj === 'undefined' ? 'undefined' : _typeof(strOrObj)) === 'object' ? strOrObj[0] : strOrObj;\n\t\treturn string.replace(/^/, '• ');\n\t}\n\n\tfunction createFirstTable(pojo) {\n\t\tvar tableArray = [['May be Circular', 'Used But Not Defined in Selection'] //header row\n\t\t];\n\t\tvar circularTerms = pojo.circular.length ? pojo.circular.map(function (pathArray) {\n\t\t\treturn pathArray.join(' ->\\r\\n').replace(/^/, '• ');\n\t\t}).join('\\r\\n') : '';\n\t\tvar notDefinedTerms = pojo.notDefined ? pojo.notDefined.map(addBullet).join('\\r\\n') : '';\n\t\tvar rowArray = [];\n\t\trowArray.push(circularTerms);\n\t\trowArray.push(notDefinedTerms);\n\t\ttableArray.push(rowArray);\n\n\t\treturn tableArray;\n\t}\n\n\tfunction createSecondTable(pojo) {\n\t\tvar tableArray = [['Cross-Reference Definitions'] //header row\n\t\t];\n\t\tvar crossRefs = pojo.crossRefs.length ? pojo.crossRefs.map(addBullet).join('\\r\\n') : '';\n\t\tvar rowArray = [];\n\t\trowArray.push(crossRefs);\n\t\ttableArray.push(rowArray);\n\n\t\treturn tableArray;\n\t}\n\n\tfunction createMainTable(pojo) {\n\t\tvar tableArray = [['Term', 'Incorporates', 'Used By', 'Defined in Selection'] //header row\n\t\t];\n\n\t\tObject.keys(pojo).forEach(function (dt) {\n\t\t\tvar incorpsObj = pojo[dt].incorps;\n\t\t\tvar incorpsTerms = incorpsObj ? Object.keys(incorpsObj).map(addBullet).join('\\r\\n') : '';\n\t\t\tvar usedByObj = pojo[dt].usedBy;\n\t\t\tvar usedByTerms = usedByObj ? Object.keys(usedByObj).map(addBullet).join('\\r\\n') : '';\n\n\t\t\tvar definedVal = pojo[dt].defined ? pojo[dt].defined : 0;\n\t\t\tvar definedTerm = definedVal === 1 ? 'yes' : definedVal === 2 ? 'yes per user' : '';\n\n\t\t\tvar rowArray = [];\n\t\t\trowArray.push(dt);\n\t\t\trowArray.push(incorpsTerms);\n\t\t\trowArray.push(usedByTerms);\n\t\t\trowArray.push(definedTerm);\n\t\t\ttableArray.push(rowArray);\n\t\t});\n\n\t\treturn tableArray;\n\t}\n\n\tfunction insertTable(docBody, tableArray) {\n\t\treturn docBody.insertTable(tableArray.length, //rowLength\n\t\ttableArray[0].length, //columnLength\n\t\tWord.InsertLocation.end, //insertPosition\n\t\ttableArray);\n\t}\n\n\treturn {\n\t\tcreateRexFromString: createRexFromString,\n\t\tsortByAlphabet: sortByAlphabet,\n\t\tsortByLongerLength: sortByLongerLength,\n\t\tsortObject: sortObject,\n\t\tmergeObjects: mergeObjects,\n\t\tmergeWithinObject: mergeWithinObject,\n\t\taddBullet: addBullet,\n\t\tcreateFirstTable: createFirstTable,\n\t\tcreateSecondTable: createSecondTable,\n\t\tcreateMainTable: createMainTable,\n\t\tinsertTable: insertTable\n\t};\n}();\n\nmodule.exports = util;\n\n//# sourceURL=webpack:///./src/appUtilities.js?");

/***/ })

/******/ });