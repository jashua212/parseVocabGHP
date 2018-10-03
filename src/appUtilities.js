/* global Word:true */

'use strict';

// Uses revealing module pattern to return an object consisting of exposed methods
var util = (function () {
	function createRexFromString(string, flags) {
		var escapedString = string.replace(/[|\\{}()[\]^$+*?.]/g, '\\$&');
		return new RegExp(escapedString, flags);
	}

	function sortByAlphabet(A, B) {
		var a = A.toLowerCase();
		var b = B.toLowerCase();

		if (a < b) {
			return -1;
		}
		if (a > b) {
			return 1;
		}
		return 0; //default return value (no sorting)
	}

	function sortByLongerLength(A, B) {
		var a = A.length;
		var b = B.length;

		if (a > b) {
			return -1;
		}
		if (a < b) {
			return 1;
		}
		return 0; //default return value (no sorting)
	}

	function sortObject(src, comparator) {
		var out = Object.create(null);

		Object.keys(src).sort(comparator).forEach(function (key) {
			if (typeof src[key] == 'object' &&
				!Array.isArray(src[key]) &&
				!(src[key] instanceof RegExp)
			) {
				out[key] = sortObject(src[key], comparator); //run function again
				return;
			} else {
				out[key] = src[key];
			}
		});

		return out;
	}

	function mergeObjects(target, src) {
		var a = target || Object.create(null);
		var b = src || Object.create(null);

		// merge b into a
		Object.keys(b).forEach(function (key) {
			a[key] = (a[key] || 0) + (b[key] || 0);
		});
	}

	function mergeWithinObject(termObj, wordPair) {
		var retainWord = wordPair[0];
		var loseWord = wordPair[1];

		Object.keys(termObj).forEach(function (mainKey) {
			if (mainKey !== 'defined') {
				var subObject = termObj[mainKey]; //can be either 'incorps' or 'usedBy' object

				Object.keys(subObject).forEach(function (word) {
					if (word === loseWord) {
						subObject[retainWord] = (subObject[retainWord] || 0) + subObject[word];
						delete subObject[word];
					}
				});
			}
		});
	}

	function addBullet(strOrObj) {
		var string = typeof strOrObj === 'object' ? strOrObj[0] : strOrObj;
		return string.replace(/^/, '• ');
	}

	function createFirstTable(pojo) {
		var tableArray = [
			['May be Circular', 'Used But Not Defined in Selection'] //header row
		];
		var circularTerms = pojo.circular.length ? pojo.circular.map(function (pathArray) {
			return pathArray.join(' ->\r\n').replace(/^/, '• ');
		}).join('\r\n') : '';
		var notDefinedTerms = pojo.notDefined ? pojo.notDefined.map(addBullet).join('\r\n') : '';
		var rowArray = [];
		rowArray.push(circularTerms);
		rowArray.push(notDefinedTerms);
		tableArray.push(rowArray);

		return tableArray;
	}

	function createSecondTable(pojo) {
		var tableArray = [
			['Cross-Reference Definitions'] //header row
		];
		var crossRefs = pojo.crossRefs.length ? pojo.crossRefs.map(addBullet).join('\r\n') : '';
		var rowArray = [];
		rowArray.push(crossRefs);
		tableArray.push(rowArray);

		return tableArray;
	}

	function createMainTable(pojo) {
		var tableArray = [
			['Term', 'Incorporates', 'Used By', 'Defined in Selection'] //header row
		];

		Object.keys(pojo).forEach(function (dt) {
			var incorpsObj = pojo[dt].incorps;
			var incorpsTerms = incorpsObj ? Object.keys(incorpsObj).map(addBullet).join('\r\n') : '';
			var usedByObj = pojo[dt].usedBy;
			var usedByTerms = usedByObj ? Object.keys(usedByObj).map(addBullet).join('\r\n') : '';

			var definedVal = pojo[dt].defined ? pojo[dt].defined : 0;
			var definedTerm = definedVal === 1 ? 'yes' : (definedVal === 2 ? 'yes per user' : '');

			var rowArray = [];
			rowArray.push(dt);
			rowArray.push(incorpsTerms);
			rowArray.push(usedByTerms);
			rowArray.push(definedTerm);
			tableArray.push(rowArray);
		});

		return tableArray;
	}

	function insertTable(docBody, tableArray) {
		return docBody.insertTable(
			tableArray.length, //rowLength
			tableArray[0].length, //columnLength
			Word.InsertLocation.end, //insertPosition
			tableArray
		);
	}

	return {
		createRexFromString: createRexFromString,
		sortByAlphabet: sortByAlphabet,
		sortByLongerLength: sortByLongerLength,
		sortObject: sortObject,
		mergeObjects: mergeObjects,
		mergeWithinObject: mergeWithinObject,
		addBullet: addBullet,
		createFirstTable: createFirstTable,
		createSecondTable: createSecondTable,
		createMainTable: createMainTable,
		insertTable: insertTable
	};
})();

module.exports = util;
