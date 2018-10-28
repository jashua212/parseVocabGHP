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
eval("/* global util:true, fabric:true, Office:true, OfficeExtension:true, Word:true */\n\n\n\n// load appUtilities module using commonJS syntax\n\nvar util = __webpack_require__(/*! ./appUtilities.js */ \"./src/appUtilities.js\");\n\n(function () {\n\tvar messageBanner;\n\n\tOffice.initialize = function () {\n\t\t$(document).ready(function () {\n\t\t\t// initialize FabricUI notification mechanism and hide it\n\t\t\tvar element = document.querySelector('.ms-MessageBanner');\n\t\t\tmessageBanner = new fabric.MessageBanner(element);\n\t\t\tmessageBanner.hideBanner();\n\n\t\t\t// check Office\n\t\t\tif (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {\n\t\t\t\tconsole.log('Sorry. This add-in uses Word.js APIs that are not available in your version of Office.');\n\t\t\t}\n\n\t\t\t$('#parse-btn').on('click', parseVocabTerms);\n\t\t\t$('#parse-btn-text').text('Parse Selected');\n\t\t});\n\t};\n\n\t/* UI Functions */\n\tfunction showNotification(header, content) {\n\t\t$(\"#notification-header\").text(header);\n\t\t$(\"#notification-body\").text(content);\n\t\tmessageBanner.showBanner();\n\t\tmessageBanner.toggleExpansion();\n\t}\n\n\tfunction errHandler(error) {\n\t\tconsole.log(\"Error: \" + error);\n\n\t\tif (error instanceof OfficeExtension.Error) {\n\t\t\tconsole.log(\"Debug info: \" + JSON.stringify(error.debugInfo));\n\t\t} else if (/TypeError: Unable to get property 'getRange'/.test(error)) {\n\t\t\tvar header = 'Error:';\n\t\t\tvar content = 'There are no definition paragraphs to select';\n\t\t\tshowNotification(header, content);\n\t\t}\n\t}\n\n\t/* Operative Functions */\n\tfunction addParaBreaksAndDashes(string) {\n\t\treturn (string || '').trim().replace(/;\\s+\\(/g, '\\n(') //add hard return\n\t\t.replace(/; (\\w)/g, ' — ' + '$1'); //add 'em' dash to separate alternative meanings\n\t}\n\n\tfunction addDashesAndTabs(string) {\n\t\treturn (string || '').trim().replace(/; (\\w)/g, ' — ' + '$1') //add 'em' dash to separate alternative meanings\n\t\t.replace(/\\((n|v|adj|adv)\\.\\) /g, '\\t' + '$&' + '\\t'); //add bookend tabs\n\t}\n\n\tfunction parseVocabTerms() {\n\t\tWord.run(function (context) {\n\t\t\t// queue command to load/return all the paragraphs as a range\n\t\t\tvar allRange = context.document.body.paragraphs;\n\t\t\tcontext.load(allRange, 'text');\n\n\t\t\treturn context.sync().then(function () {\n\t\t\t\tvar paras = allRange.items.map(function (p) {\n\t\t\t\t\treturn p.text.trim();\n\t\t\t\t}).filter(function (p) {\n\t\t\t\t\treturn p; //filter out empty items in array\n\t\t\t\t});\n\t\t\t\tconsole.log('paras', paras);\n\n\t\t\t\t/* START HERE */\n\t\t\t\tvar pojo = {};\n\n\t\t\t\tparas.forEach(function (p) {\n\t\t\t\t\tif (!/^\\*/.test(p)) {\n\t\t\t\t\t\tvar arr = p.split('\\t');\n\t\t\t\t\t\tconsole.log('arr', arr);\n\n\t\t\t\t\t\t//set 'term' for this para and subsequent SYNONYM/ANTONYM paras\n\t\t\t\t\t\tvar term = arr[0].trim();\n\n\t\t\t\t\t\t//create term object within pojo\n\t\t\t\t\t\tpojo[term] = Object.create(null);\n\n\t\t\t\t\t\t//add definition thereto\n\t\t\t\t\t\tpojo[term].defs = arr[1].trim().match(/\\((n|v|adj|adv)\\.\\)[^;(]+/g);\n\t\t\t\t\t} else {\n\t\t\t\t\t\tvar lastValue = util.getValueOfLastKey(pojo); //should be getting an object\n\n\t\t\t\t\t\tif (/SYNONYMS/.test(p)) {\n\t\t\t\t\t\t\tvar synos = p.replace('*SYNONYMS:*', '');\n\n\t\t\t\t\t\t\tlastValue.synos = addParaBreaksAndDashes(synos);\n\t\t\t\t\t\t} else if (/ANTONYMS/.test(p)) {\n\t\t\t\t\t\t\tvar antos = p.replace('*ANTONYMS:*', '');\n\n\t\t\t\t\t\t\tlastValue.antos = addParaBreaksAndDashes(antos);\n\t\t\t\t\t\t}\n\t\t\t\t\t}\n\t\t\t\t});\n\n\t\t\t\tvar sortedPojo = util.sortObject(pojo, util.sortByAlphabet);\n\t\t\t\tconsole.log('debug sortedPojo', sortedPojo);\n\t\t\t\t/* END HERE */\n\n\t\t\t\t// Throw error if pojo is empty\n\t\t\t\tif (!Object.keys(sortedPojo).length) {\n\t\t\t\t\tvar header = 'Error:';\n\t\t\t\t\tvar content = 'No definition paragraphs have been selected';\n\t\t\t\t\tshowNotification(header, content);\n\n\t\t\t\t\treturn context.sync(); //bail\n\t\t\t\t}\n\n\t\t\t\t// Create master array of individual term tables\n\t\t\t\tvar masterTableArray = [];\n\n\t\t\t\tObject.keys(sortedPojo).forEach(function (term) {\n\t\t\t\t\tvar termTableArray = [];\n\t\t\t\t\tvar termObj = sortedPojo[term];\n\n\t\t\t\t\t//populate termTableArray\n\t\t\t\t\ttermTableArray.push([term]);\n\n\t\t\t\t\ttermObj.defs.forEach(function (dd) {\n\t\t\t\t\t\ttermTableArray.push([addDashesAndTabs(dd)]);\n\t\t\t\t\t});\n\n\t\t\t\t\tif (termObj.synos) {\n\t\t\t\t\t\ttermTableArray.push(['synonyms:', termObj.synos]);\n\t\t\t\t\t}\n\n\t\t\t\t\tif (termObj.antos) {\n\t\t\t\t\t\ttermTableArray.push(['antonyms:', termObj.antos]);\n\t\t\t\t\t}\n\n\t\t\t\t\t//push termTableArray into masterTableArray\n\t\t\t\t\tmasterTableArray.push(termTableArray);\n\t\t\t\t});\n\n\t\t\t\t// Create separate parts of speech table\n\t\t\t\tvar partsOfSpeechTable = [['adjective', 'noun', 'adverb', 'verb']];\n\n\t\t\t\tfor (var i = 0; i < 20; i++) {\n\t\t\t\t\tpartsOfSpeechTable.push(['', '', '', '']);\n\t\t\t\t}\n\n\t\t\t\t// Create separate table array consisting solely of terms\n\t\t\t\t// should be 20 terms, divided into 4 columns and 5 rows\n\t\t\t\tvar termsOnlyTableArray = [[], [], [], [], []];\n\n\t\t\t\tObject.keys(sortedPojo).forEach(function (term, i) {\n\t\t\t\t\tvar moduloRemainder = i % 5;\n\n\t\t\t\t\ttermsOnlyTableArray[moduloRemainder].push(term);\n\t\t\t\t});\n\n\t\t\t\tvar newDoc = context.application.createDocument();\n\t\t\t\tcontext.load(newDoc);\n\n\t\t\t\treturn context.sync().then(function () {\n\t\t\t\t\t// console.log('newDoc', newDoc);\n\t\t\t\t\tconsole.log('masterTableArray', masterTableArray);\n\t\t\t\t\tvar newDocBody = newDoc.body;\n\n\t\t\t\t\tnewDocBody.font.name = 'Arial';\n\t\t\t\t\tnewDocBody.font.size = 11;\n\n\t\t\t\t\t// insert and style each individual term table\n\t\t\t\t\tmasterTableArray.forEach(function (termTableArray) {\n\t\t\t\t\t\tvar table = util.insertTable(newDocBody, termTableArray);\n\t\t\t\t\t\ttable.headerRowCount = 0;\n\t\t\t\t\t\ttable.style = 'Grid Table 1 Light - Accent 1';\n\t\t\t\t\t\ttable.styleFirstColumn = false;\n\t\t\t\t\t});\n\n\t\t\t\t\t// insert and style the partsOfSpeechTable\n\t\t\t\t\tvar partsOfSpeechTable = util.insertTable(newDocBody, partsOfSpeechTable);\n\t\t\t\t\tpartsOfSpeechTable.style = 'Table Grid Light';\n\n\t\t\t\t\t// insert and style the termsOnlyTableArray\n\t\t\t\t\tvar allTermsTable = util.insertTable(newDocBody, termsOnlyTableArray);\n\t\t\t\t\tallTermsTable.style = 'Table Grid Light';\n\n\t\t\t\t\treturn context.sync().then(function () {\n\t\t\t\t\t\tnewDoc.open();\n\n\t\t\t\t\t\treturn context.sync();\n\t\t\t\t\t});\n\t\t\t\t});\n\t\t\t});\n\t\t}).catch(errHandler);\n\t}\n})();\n\n//# sourceURL=webpack:///./src/app.js?");

/***/ }),

/***/ "./src/appUtilities.js":
/*!*****************************!*\
  !*** ./src/appUtilities.js ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
eval("/* global Word:true */\n\n\n\n// Uses revealing module pattern to return an object consisting of exposed methods\n\nvar _typeof = typeof Symbol === \"function\" && typeof Symbol.iterator === \"symbol\" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === \"function\" && obj.constructor === Symbol && obj !== Symbol.prototype ? \"symbol\" : typeof obj; };\n\nvar util = function () {\n\tfunction createRexFromString(string, flags) {\n\t\tvar escapedString = string.replace(/[|\\\\{}()[\\]^$+*?.]/g, '\\\\$&');\n\t\treturn new RegExp(escapedString, flags);\n\t}\n\n\tfunction sortByAlphabet(A, B) {\n\t\tvar a = A.toLowerCase();\n\t\tvar b = B.toLowerCase();\n\n\t\tif (a < b) {\n\t\t\treturn -1;\n\t\t}\n\t\tif (a > b) {\n\t\t\treturn 1;\n\t\t}\n\t\treturn 0; //default return value (no sorting)\n\t}\n\n\tfunction sortByLongerLength(A, B) {\n\t\tvar a = A.length;\n\t\tvar b = B.length;\n\n\t\tif (a > b) {\n\t\t\treturn -1;\n\t\t}\n\t\tif (a < b) {\n\t\t\treturn 1;\n\t\t}\n\t\treturn 0; //default return value (no sorting)\n\t}\n\n\tfunction sortObject(src, comparator) {\n\t\tvar out = Object.create(null);\n\n\t\tObject.keys(src).sort(comparator).forEach(function (key) {\n\t\t\tif (_typeof(src[key]) == 'object' && !Array.isArray(src[key]) && !(src[key] instanceof RegExp)) {\n\t\t\t\tout[key] = sortObject(src[key], comparator); //run function again\n\t\t\t\treturn;\n\t\t\t} else {\n\t\t\t\tout[key] = src[key];\n\t\t\t}\n\t\t});\n\n\t\treturn out;\n\t}\n\n\tfunction mergeObjects(target, src) {\n\t\tvar a = target || Object.create(null);\n\t\tvar b = src || Object.create(null);\n\n\t\t// merge b into a\n\t\tObject.keys(b).forEach(function (key) {\n\t\t\ta[key] = (a[key] || 0) + (b[key] || 0);\n\t\t});\n\t}\n\n\tfunction mergeWithinObject(termObj, wordPair) {\n\t\tvar retainWord = wordPair[0];\n\t\tvar loseWord = wordPair[1];\n\n\t\tObject.keys(termObj).forEach(function (mainKey) {\n\t\t\tif (mainKey !== 'defined') {\n\t\t\t\tvar subObject = termObj[mainKey]; //can be either 'incorps' or 'usedBy' object\n\n\t\t\t\tObject.keys(subObject).forEach(function (word) {\n\t\t\t\t\tif (word === loseWord) {\n\t\t\t\t\t\tsubObject[retainWord] = (subObject[retainWord] || 0) + subObject[word];\n\t\t\t\t\t\tdelete subObject[word];\n\t\t\t\t\t}\n\t\t\t\t});\n\t\t\t}\n\t\t});\n\t}\n\n\tfunction getValueOfLastKey(obj) {\n\t\t//this function won't work in IE\n\t\tvar objKeysArray = Object.keys(obj);\n\t\tvar length = objKeysArray.length;\n\n\t\treturn obj[objKeysArray[length - 1]]; //value for last key of obj\n\t}\n\n\tfunction addBullet(strOrObj) {\n\t\tvar string = (typeof strOrObj === 'undefined' ? 'undefined' : _typeof(strOrObj)) === 'object' ? strOrObj[0] : strOrObj;\n\t\treturn string.replace(/^/, '• ');\n\t}\n\n\tfunction createFirstTable(pojo) {\n\t\tvar tableArray = [['May be Circular', 'Used But Not Defined in Selection'] //header row\n\t\t];\n\t\tvar circularTerms = pojo.circular.length ? pojo.circular.map(function (pathArray) {\n\t\t\treturn pathArray.join(' ->\\r\\n').replace(/^/, '• ');\n\t\t}).join('\\r\\n') : '';\n\t\tvar notDefinedTerms = pojo.notDefined ? pojo.notDefined.map(addBullet).join('\\r\\n') : '';\n\t\tvar rowArray = [];\n\t\trowArray.push(circularTerms);\n\t\trowArray.push(notDefinedTerms);\n\t\ttableArray.push(rowArray);\n\n\t\treturn tableArray;\n\t}\n\n\tfunction createSecondTable(pojo) {\n\t\tvar tableArray = [['Cross-Reference Definitions'] //header row\n\t\t];\n\t\tvar crossRefs = pojo.crossRefs.length ? pojo.crossRefs.map(addBullet).join('\\r\\n') : '';\n\t\tvar rowArray = [];\n\t\trowArray.push(crossRefs);\n\t\ttableArray.push(rowArray);\n\n\t\treturn tableArray;\n\t}\n\n\tfunction createMainTable(pojo) {\n\t\tvar tableArray = [['Term', 'Incorporates', 'Used By', 'Defined in Selection'] //header row\n\t\t];\n\n\t\tObject.keys(pojo).forEach(function (dt) {\n\t\t\tvar incorpsObj = pojo[dt].incorps;\n\t\t\tvar incorpsTerms = incorpsObj ? Object.keys(incorpsObj).map(addBullet).join('\\r\\n') : '';\n\t\t\tvar usedByObj = pojo[dt].usedBy;\n\t\t\tvar usedByTerms = usedByObj ? Object.keys(usedByObj).map(addBullet).join('\\r\\n') : '';\n\n\t\t\tvar definedVal = pojo[dt].defined ? pojo[dt].defined : 0;\n\t\t\tvar definedTerm = definedVal === 1 ? 'yes' : definedVal === 2 ? 'yes per user' : '';\n\n\t\t\tvar rowArray = [];\n\t\t\trowArray.push(dt);\n\t\t\trowArray.push(incorpsTerms);\n\t\t\trowArray.push(usedByTerms);\n\t\t\trowArray.push(definedTerm);\n\t\t\ttableArray.push(rowArray);\n\t\t});\n\n\t\treturn tableArray;\n\t}\n\n\tfunction insertTable(docBody, tableArray) {\n\t\treturn docBody.insertTable(tableArray.length, //rowLength\n\t\tMath.max(2, tableArray[0].length), //columnLength\n\t\tWord.InsertLocation.end, //insertPosition\n\t\ttableArray);\n\t}\n\n\treturn {\n\t\tcreateRexFromString: createRexFromString,\n\t\tsortByAlphabet: sortByAlphabet,\n\t\tsortByLongerLength: sortByLongerLength,\n\t\tsortObject: sortObject,\n\t\tmergeObjects: mergeObjects,\n\t\tmergeWithinObject: mergeWithinObject,\n\t\tgetValueOfLastKey: getValueOfLastKey,\n\t\taddBullet: addBullet,\n\t\tcreateFirstTable: createFirstTable,\n\t\tcreateSecondTable: createSecondTable,\n\t\tcreateMainTable: createMainTable,\n\t\tinsertTable: insertTable\n\t};\n}();\n\nmodule.exports = util;\n\n//# sourceURL=webpack:///./src/appUtilities.js?");

/***/ })

/******/ });