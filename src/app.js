/* global util:true, fabric:true, Office:true, OfficeExtension:true, Word:true */

'use strict';

// load appUtilities module using commonJS syntax
const util = require('./appUtilities.js');

(function () {
	var messageBanner;
	// var allRangeLength = 0;

	Office.initialize = function () {
		$(document).ready(function () {
			// initialize FabricUI notification mechanism and hide it
			var element = document.querySelector('.ms-MessageBanner');
			messageBanner = new fabric.MessageBanner(element);
			messageBanner.hideBanner();

			// check Office
			if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
				console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
			}

			var docx = Office.context.document;

			// pull into 'live settings' the data (if any) that is stored in the file
			docx.settings.refreshAsync(function () {
				// get userTerms from live settings and show them in ui
				['add', 'minus'].forEach(function (cmd) {
					addToShownUserTerms(cmd, docx.settings.get('userTerms-' + cmd) || []);
				});
			});

			$('#user-term-add').on('keydown', function (e) {
				if (e.keyCode === 13) {
					keydownHandler('add', $(this));
				}
			});
			$('#user-term-minus').on('keydown', function (e) {
				if (e.keyCode === 13) {
					keydownHandler('minus', $(this));
				}
			});

			$('#user-terms-add-container').on('click', '.user-term', function () {
				removeClickHandler('add', $(this));
			});
			$('#user-terms-minus-container').on('click', '.user-term', function () {
				removeClickHandler('minus', $(this));
			});

			$('#select-btn').on('click', selectDefParas);
			$('#select-btn-text').text('Select Definition Paragraphs');

			$('#parse-btn').on('click', parseVocabTerms);
			$('#parse-btn-text').text('Parse Selected');
		});
	};

	/* UI Functions */
	function keydownHandler(cmd, elm) {
		var inpVal = elm.val().trim();

		if (!inpVal) {
			return; //bail
		}

		// add to shown user terms if not a dupe
		if (getShownUserTerms(cmd).indexOf(inpVal) === -1) {
			addToShownUserTerms(cmd, [inpVal]);
			elm.val(''); //clear input
		}

		// sync to settings if not a dupe
		var docx = Office.context.document;
		var userTerms = docx.settings.get('userTerms-' + cmd) || [];
		if (userTerms.indexOf(inpVal) === -1) {
			userTerms.push(inpVal);
			userTerms.sort(util.sortByAlphabet);
			docx.settings.set('userTerms-' + cmd, userTerms);
			docx.settings.saveAsync();
		}
	}

	function removeClickHandler(cmd, elm) {
		var val = elm.text();
		elm.remove();

		// sync to settings
		var docx = Office.context.document;
		var userTerms = docx.settings.get('userTerms-' + cmd);
		if (userTerms) {
			userTerms.splice(userTerms.indexOf(val), 1);
			docx.settings.set('userTerms-' + cmd, userTerms);
			docx.settings.saveAsync();
		}
	}

	function getShownUserTerms(cmd) {
		var userTerms = [];

		$('#user-terms-' + cmd + '-container .user-term').each(function () {
			userTerms.push($(this).text());
		});

		return userTerms;
	}

	function addToShownUserTerms(cmd, arrayOfTerms) {
		var container = $('#user-terms-' + cmd + '-container');
		var frag = document.createDocumentFragment();

		arrayOfTerms.forEach(function (term) {
			var div = document.createElement('div');
			div.classList.add('user-term');
			div.textContent = term;
			frag.appendChild(div);
		});
		container.prepend(frag);

		return container;
	}

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
	/* function selectAll() {
		Word.run(function (context) {
			// queue command to select whole doc
			context.document.body.select();

			// queue command to load/return all the paragraphs as a range
			var allRange = context.document.body.paragraphs;
			context.load(allRange, 'text');

			return context.sync().then(function () {
				// if successful, store allRange.items.length in global var
				allRangeLength = allRange.items.length;
				console.log('allRangeLength', allRangeLength);
			});
		})
		.catch(errHandler);
	} */

	function bifurcateParas(paras) {
		const rexqtBeginning = /(^|(\(\w{1,3}\)\s+?))“[^”]+”/;

		let startIndex = paras
			.findIndex(function (p) {
				return rexqtBeginning.test(p);
			});

		let revStartIndex = paras.slice(0).reverse()
			.findIndex(function (p) {
				return rexqtBeginning.test(p);
			});
		let endIndex = paras.length - (revStartIndex + 1);

		/* let defParas = paras
			.filter(function (p, i) {
				return i >= startIndex && i <= endIndex;
			});

		let plainParas = paras
			.filter(function (p, i) {
				return i < startIndex || i > endIndex;
			});

		return [defParas, plainParas]; */

		return [startIndex, endIndex];
	}

	function selectDefParas() {
		Word.run(function (context) {
			// queue command to load/return all the paragraphs as a range
			var allRange = context.document.body.paragraphs;
			context.load(allRange, 'text');

			return context.sync().then(function () {
				var allParas = allRange.items.map(function (p) {
					return p.text.trim();
				});

				var indices = bifurcateParas(allParas);
				var startIndex = indices[0];
				var endIndex = indices[1];
				var startRange = allRange.items[startIndex].getRange();
				var endRange = allRange.items[endIndex].getRange();

				var expandedRange = endRange.expandTo(startRange);
				expandedRange.select();

				return context.sync();
			});
		})
		.catch(errHandler);
	}

	/* function getCrossRefDefs(paras) {
		const rexFirstSentence = /^.+?\.(?:\s|$)/;
		return paras
			.map(function (p) {
				return p.match(rexFirstSentence);
			})
			.reduce(function (accumArray, matchArray) {
				return accumArray.concat(matchArray); //flatten into a single array of strings
			}, [])
			.filter(function (sentence) {
				return /\b(meaning|defined|definition)s*?\b/.test(sentence);
			})
			.filter(function (sentence) {
				return /^“/.test(sentence);
			})
			.filter(function (sentence) {
				return sentence[0].split(' ').length < 30;
			});
	} */

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

				// check agst global var to confirm that whole doc is still selected
				/* if (paras.length === allRangeLength) {
					// if so, trim paragraph collection (in place) from the end
					let revLastIndex = paras.slice(0).reverse()
						.findIndex(function (item) {
							return /^“[^”]+”/.test(item);
						});
					paras.splice((revLastIndex * -1));
					console.log('SPLICED PARAS', paras);

				} else {
					// otherwise, reset global var and don't trim paragraph collection
					allRangeLength = 0;
				} */

				/* START HERE */
				var pojo = Object.create(null);
				var lastTerm;

				paras.forEach(function (p) {
					if (!/^\*/.test(p)) {
						let arr = p.split('\t');
						console.log('arr', arr);

						//set 'term' for this para and subsequent SYNONYM/ANTONYM paras
						var term = lastTerm = arr[0];

						//create term object within pojo
						pojo[term] = Object.create(null);

						//add definition thereto
						pojo[term].def = arr[1];

					} else if (/SYNONYMS/.test(p)) {
						let synos = p.replace('*SYNONYMS:*', '').trim();
						console.log('synos', synos);

						pojo[lastTerm].synos = synos;

					} else if (/ANTONYMS/.test(p)) {
						let antos = p.replace('*ANTONYMS:*', '').trim();
						console.log('antos', antos);

						pojo[lastTerm].antos = antos;

					} else {
						console.log('error parsing empty para');
					}
				});
				lastTerm = '';

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
					termTableArray.push([term, ' ']);
					termTableArray.push([termObj.def, ' ']);

					if (termObj.synos) {
						termTableArray.push(['SYNONYMS', termObj.synos]);
					}

					if (termObj.antos) {
						termTableArray.push(['ANTONYMS', termObj.antos]);
					}

					//push termTableArray into masterTableArray
					masterTableArray.push(termTableArray);
				});

				var newDoc = context.application.createDocument();
				context.load(newDoc);

				return context.sync().then(function () {
					// console.log('newDoc', newDoc);
					console.log('masterTableArray', masterTableArray);

					masterTableArray.forEach(function (termTableArray) {
						var table = util.insertTable(newDoc.body, termTableArray);
						// table.headerRowCount = 0;
						/* table.style = 'List Table 4 - Accent 1';
						table.styleFirstColumn = false; */
					});

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
