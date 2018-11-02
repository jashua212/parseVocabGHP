/* global util:true, fabric:true, Office:true, OfficeExtension:true, Word:true */

'use strict';

// load appUtilities module using commonJS syntax
const util = require('./appUtilities.js');

(function () {
	var messageBanner;

	Office.initialize = function () {
		$(document).ready(function () {
			// initialize FabricUI notification mechanism and hide it
			var element = document.querySelector('.ms-MessageBanner');
			messageBanner = new fabric.MessageBanner(element);
			messageBanner.hideBanner();

			// check Office
			if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
				console.log('Sorry. This add-in uses Word.js APIs that are not available in your version of Office.');
			}

			$('#vocab-parse-btn').on('click', parseVocabTerms);
			$('#vocab-parse-btn-text').text('Parse Vocabulary');

			$('#annot-parse-btn').on('click', getBodyHtml);
			$('#annot-parse-btn-text').text('Parse Annotations');
		});
	};

	/* UI Functions */
	function showNotification(header, content) {
		$("#notification-header").text(header);
		$("#notification-body").text(content);
		messageBanner.showBanner();
		messageBanner.toggleExpansion();
	}

	function errHandler(error) {
		console.log("Error: " + error);

		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));

		} else if (/TypeError: Unable to get property 'getRange'/.test(error)) {
			var header = 'Error:';
			var content = 'There are no definition paragraphs to select';
			showNotification(header, content);
		}
	}

	/* Operative Functions */
	function addParaBreaksAndDashes(string) {
		return (string || '')
			.trim()
			.replace(/;\s+\(/g, '\n(') //add hard return
			.replace(/; (\w)/g, ' — ' + '$1'); //add 'em' dash to separate alternative meanings
	}

	function addDashesAndTabs(string) {
		return (string || '')
			.trim()
			.replace(/; (\w)/g, ' — ' + '$1') //add 'em' dash to separate alternative meanings
			.replace(/\((n|v|adj|adv)\.\) /g, '\t' + '$&' + '\t'); //add bookend tabs
	}

	function parseVocabTerms() {
		Word.run(function (context) {
			// queue command to load/return all the paragraphs as a range
			var allRange = context.document.body.paragraphs;
			context.load(allRange, 'text');

			return context.sync().then(function () {
				var paras = allRange.items
					.map(function (p) {
						return p.text.trim();
					})
					.filter(function (p) {
						return p; //filter out empty items in array
					});
					console.log('paras', paras);

				/* START HERE */
				var pojo = {};

				paras.forEach(function (p) {
					if (!/(syn|ant)onym/i.test(p)) {
						let arr = p.split('\t');
						console.log('arr', arr);

						//set 'term' for this para and subsequent SYNONYM/ANTONYM paras
						var term = arr[0].trim();

						//create term object within pojo
						pojo[term] = Object.create(null);

						//add definitions to term object
						var defs = arr[1]
							.trim()
							.match(/\((n|v|adj|adv)\.\)[^(]+/g);

						pojo[term].defs = defs;

						//add parts of speech row to term object
						var posRow = ['', '', '', ''];

						defs.forEach(function (dd) {
							var match = dd.match(/\((\w+?)\.\)/);

							if (match) {
								if (match[1] == 'adj') {
									posRow[0] = term;
								} else if (match[1] == 'n') {
									posRow[1] = term;
								} else if (match[1] == 'adv') {
									posRow[2] = term;
								} else if (match[1] == 'v') {
									posRow[3] = term;
								}
							}
						});

						pojo[term].posRow = posRow;

					} else {
						var lastValue = util.getValueOfLastKey(pojo); //should be getting an object

						if (/SYNONYMS/.test(p)) {
							var synos = p.replace('*SYNONYMS:*', '');

							lastValue.synos = addParaBreaksAndDashes(synos);

						} else if (/ANTONYMS/.test(p)) {
							var antos = p.replace('*ANTONYMS:*', '');

							lastValue.antos = addParaBreaksAndDashes(antos);
						}
					}
				});

				var sortedPojo = util.sortObject(pojo, util.sortByAlphabet);
				console.log('debug sortedPojo', sortedPojo);
				/* END HERE */

				// Throw error if pojo is empty
				if (!Object.keys(sortedPojo).length) {
					var header = 'Error:';
					var content = 'No definition paragraphs have been selected';
					showNotification(header, content);

					return context.sync(); //bail
				}

				// Create master array of individual term tables
				var masterTableArray = [];

				Object.keys(sortedPojo).forEach(function (term) {
					var termTableArray = [];
					var termObj = sortedPojo[term];

					//populate termTableArray
					termTableArray.push([term]);

					termObj.defs.forEach(function (dd) {
						termTableArray.push([addDashesAndTabs(dd)]);
					});

					if (termObj.synos) {
						termTableArray.push(['synonyms:', termObj.synos]);
					}

					if (termObj.antos) {
						termTableArray.push(['antonyms:', termObj.antos]);
					}

					//push termTableArray into masterTableArray
					masterTableArray.push(termTableArray);
				});

				// Create separate parts of speech table
				var partsOfSpeechTableArray = [
					['adjective', 'noun', 'adverb', 'verb'],
					['', '', '', '']
				];

				Object.keys(sortedPojo).forEach(function (term) {
					partsOfSpeechTableArray.push(sortedPojo[term].posRow);
				});

				// Create separate table array consisting solely of terms
				// should be 20 terms, divided into 4 columns and 5 rows
				var termsOnlyTableArray = [
					[], [], [], [], []
				];

				Object.keys(sortedPojo).forEach(function (term, i) {
					var moduloRemainder = i % 5;

					termsOnlyTableArray[moduloRemainder].push(term);
				});


				var newDoc = context.application.createDocument();
				context.load(newDoc);

				return context.sync().then(function () {
					// console.log('newDoc', newDoc);
					console.log('masterTableArray', masterTableArray);
					var newDocBody = newDoc.body;

					newDocBody.font.name = 'Arial';
					newDocBody.font.size = 11;

					// insert and style each individual term table
					masterTableArray.forEach(function (termTableArray) {
						var table = util.insertTable(newDocBody, termTableArray);
						table.headerRowCount = 0;
						table.style = 'Grid Table 1 Light - Accent 1';
						table.styleFirstColumn = false;
					});

					// insert and style the partsOfSpeechTableArray
					var partsOfSpeechTable = util.insertTable(newDocBody, partsOfSpeechTableArray);
					partsOfSpeechTable.headerRowCount = 1;
					partsOfSpeechTable.style = 'Table Grid Light';

					// insert and style the termsOnlyTableArray
					var allTermsTable = util.insertTable(newDocBody, termsOnlyTableArray);
					allTermsTable.style = 'Table Grid Light';

					return context.sync().then(function () {
						newDoc.open();

						return context.sync();
					});
				});
			});
		})
		.catch(errHandler);
	}

	// doesn't work b/c don't seem to have access to hyperlinks
	/* function searchForText() {
		Word.run(function (context) {
			// queue command to load/return all the paragraphs as a range
			var searchResults = context.document.body.search(
				'zotero://open-pdf/library/items/[A-Z0-9]{1,}',
				{matchWildcards: true}
			);
			context.load(searchResults, 'text');

			return context.sync().then(function () {
				console.log(searchResults.items);
			});
		})
		.catch(errHandler);
	} */

	function getBodyHtml() {
		Word.run(function (context) {
			// queue command to load/return all the paragraphs as a range
			var body = context.document.body;
			var bodyHtml = body.getHtml();
			context.load(bodyHtml);

			return context.sync().then(function () {
				console.log("Body HTML contents: " + bodyHTML.value);
			});
		})
		.catch(errHandler);
	}

	function parseAnnotations() {
		Word.run(function (context) {
			// queue command to load/return all the paragraphs as a range
			var allRange = context.document.body.paragraphs;
			context.load(allRange, 'text');

			return context.sync().then(function () {
				var paras = allRange.items
					.map(function (p) {
						return p.text.trim();
					})
					.filter(function (p) {
						return p; //filter out empty items in array
					});
					console.log('paras', paras);

				/* START HERE */
				// going straight to table array (instead of an intermediate object)
				var tableArray = [];

				paras.forEach(function (p) {
					// if not a note -- notes have "{note on ...) " at their ends
					if (!/\(note.+?\)\s*$/i.test(p)) {
						// remove beginning and end quotation marks
						var text = p
							.replace(/^"/, '')
							.replace(/(")(\s*\(.+\))/, '$2');

						tableArray.push([text]);

					} else {
						tableArray[tableArray.length - 1].push(p);
					}
				});

				/* END HERE */
				var newDoc = context.application.createDocument();
				context.load(newDoc);

				return context.sync().then(function () {
					var newDocBody = newDoc.body;

					newDocBody.font.name = 'TimesNewRoman';
					newDocBody.font.size = 12;

					// insert and style the tableArray
					var annotTable = util.insertTable(newDocBody, tableArray, 3);
					annotTable.style = 'Table Grid Light';

					return context.sync().then(function () {
						newDoc.open();

						return context.sync();
					});
				});
			});
		})
		.catch(errHandler);
	}
})();
