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

			$('#parse-btn').on('click', parseParas);
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

	function getCrossRefDefs(paras) {
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
	}

	function parseParas() {
		keydownHandler('add', $('#user-term-add'));
		keydownHandler('minus', $('#user-term-minus'));

		Word.run(function (context) {
			// queue command to load/return all the paragraphs in the current selection as a range
			var selRange = context.document.getSelection().paragraphs;
			context.load(selRange, 'text');

			return context.sync().then(function () {
				var paras = selRange.items.map(function (p) {
					return p.text.trim();
				});
				// console.log('paras.length', paras.length);

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
				var rexPojo = Object.create(null);
				var rexqtPhrase = /(^|(\(\w{1,3}\)\s+?))“[^”]+”([^”]{1,7}“[^”]+”)*/;
				var rexqts = /“[^”]+”/g;

				/* 'REXPOJO' PASS */
				// populate rexPojo with every quoted term appearing at the beginning of each para
				paras.forEach(function (p) {
					var qtPhrase = p.match(rexqtPhrase);
					if (qtPhrase) {
						(qtPhrase[0].match(rexqts) || [])
							.map(function (qt) {
								return qt.replace(/[“”\,]/g, '');
							})
							.forEach(function (dt) {
								rexPojo[dt] = util.createRexFromString(dt, 'g'); //put in rexPojo
							});
					}
				});
				// console.log('rexPojo before adding userTerms', rexPojo);

				// add user specified terms (held in live settings) to rexPojo
				// also, store them in a variable for adjustments below
				var userTermsAdded = (Office.context.document.settings.get('userTerms-add') || []);
				userTermsAdded.forEach(function (uta) {
					rexPojo[uta] = util.createRexFromString(uta, 'g'); //put in rexPojo
				});

				// sort rexPojo by length (so longer ones get removed from para first per below, and
				// avoid creating fragments of defined terms that would be caught later by init caps)
				var sortedRexPojo = util.sortObject(rexPojo, util.sortByLongerLength); /*key*/
				// console.log('sortedRexPojo', sortedRexPojo);

				/* 'INCORPS' PASS */
				// populate 'incorps'
				var sortedRexPojoLowerCaseKeys = Object.keys(sortedRexPojo).map(function (key) {
						return key.toLowerCase();
					});
				var pojo = Object.create(null);
				var last_dts;
				var rexInitCaps = /((([A-Z][\w\-]+|\d{4})\s?(of|and|to)?\s?)(\d{4}(\-\d{1,2})?\s?)?)+/g;
				var rexLeadArticles = /^(A|An|If|The|This|That|Each|Such|Every|Following)\s/;
				var badLoneWords = ['by', 'name', 'title', 'date', 'for', 'with', 'each', 'if', 'the', 'this', 'none', 'such', 'every', 'in', 'on'];

				paras.forEach(function (p) {
					var dts;
					var qtPhrase = p.match(rexqtPhrase);
					if (qtPhrase) {
						last_dts = dts = (qtPhrase[0].match(rexqts) || [])
						.map(function (qt) {
							return qt.replace(/[“”\,]/g, '');
						});
						// the above replicates the rexPojo Pass, except that, here, we track last_dts
						// to link dts to paras that don't have quoted defined terms at their beginnings
					} else {
						dts = last_dts; //use last_dts (since this para doesn't have its own dts)
					}

					(dts || []).forEach(function (t) {
						if (!pojo[t]) {
							pojo[t] = Object.create(null); //add defined term to pojo
						}
						pojo[t].defined = 1; //track if t is a "defined term"

						// apply sortedRexPojo to find incorporated terms in para
						Object.keys(sortedRexPojo).forEach(function (key) {
							var rex = sortedRexPojo[key];
							(p.match(rex) || [])
								.filter(function (n) {
									return dts.indexOf(n) === -1; //exclude any defined terms (i.e., itself)
								})
								.forEach(function (n) {
									if (!pojo[t].incorps) {
										pojo[t].incorps = Object.create(null);
									}
									pojo[t].incorps[n] = (pojo[t].incorps[n] + 1) || 1;
								});

							// lowercase defined term in para to avoid catching fragments later /*key*/
							p = p.replace(rex, function (match) {
								return match.toLowerCase();
							});
						});

						// console.log('p', p);

						// apply init caps, after applying sortedRexPojo
						(p.match(/“[^”]+”/g) || []) //get all quoted terms contained in the p
							.map(function (qt) {
								return qt.replace(/[“”\,]/g, ''); //remove their quotation marks
							})
							.filter(function (dt) {
								return sortedRexPojoLowerCaseKeys.indexOf(dt) === -1; //exclude those caught by rexPojo pass
							})
							.filter(function (dt) {
								return /^[a-z]/.test(dt); //keep those whose first letter is lower case
							})
							.concat(p.match(rexInitCaps) || []) //CONCAT with new array of init caps

							.map(function (n) {
								return n.trim() //trim leading and trailing spaces
									.replace(rexLeadArticles, '') //trim leading articles
									.replace(/\s(of|and|to)$/, ''); //trim trailing of|and|to;
							})
							.filter(function (n) {
								return n.length && dts.indexOf(n) === -1; //exclude any defined terms
							})
							.filter(function (n) {
								return badLoneWords.indexOf(n.toLowerCase()) === -1; //exclude badLoneWords
							})
							.filter(function (n) {
								return !/^\d+$/.test(n); //exclude number-only strings
							})
							.forEach(function (n) {
								if (!pojo[t].incorps) {
									pojo[t].incorps = Object.create(null);
								}
								pojo[t].incorps[n] = (pojo[t].incorps[n] + 1) || 1;
							});
					});
				});

				/* REMOVE PASS */
				(Office.context.document.settings.get('userTerms-minus') || [])
				.forEach(function (utm) {
					Object.keys(pojo).forEach(function (key) {
						if (key === utm) {
							delete pojo[key];

						} else {
							var incorpsObj = pojo[key].incorps;

							if (incorpsObj) {
								Object.keys(incorpsObj).forEach(function (term) {
									if (term === utm ||
										// utm is plural
										term + 's' === utm ||
										term + 'es' === utm ||
										term.substring(0, term.length - 1) + 'ies' === utm ||
										// utm is singular
										utm + 's' === term ||
										utm + 'es' === term ||
										utm.substring(0, utm.length - 1) + 'ies' === term
									) {
										delete pojo[key].incorps[term];
									}
								});
							}
						}
					});
				});

				/* 'USEDBY' PASS */
				// use incorps data to populate 'usedBy'
				Object.keys(pojo).forEach(function (t) {
					// console.log(pojo[t].incorps);
					if (pojo[t].incorps) {
						Object.keys(pojo[t].incorps).forEach(function (n) {
							// console.log(n);
							if (!pojo[n]) {
								pojo[n] = Object.create(null);
							}
							if (!pojo[n].usedBy) {
								pojo[n].usedBy = Object.create(null);
							}
							var val = pojo[t].incorps[n];
							pojo[n].usedBy[t] = (pojo[n].usedBy[t] + val) || val;
						});
					}
				});

				var sortedPojo = util.sortObject(pojo, util.sortByAlphabet);
				// console.log('debug sortedPojo', sortedPojo);

				/* PLURAL PASS */
				var arrayOfWordPairs = []; //an array of arrays

				Object.keys(sortedPojo).forEach(function (second, i, self) {
					if (i > 0) {
						var first = self[i - 1]; //previous key
						var trimOneFromEnd = second.length - 1;
						// console.log(second.substring(0, trimOneFromEnd) + 'ies' === first);

						if ((second === first + 's') ||
							(second === first + 'es') ||
							(second.substring(0, trimOneFromEnd) + 'ies' === first)
						) {
							// console.log('first', first, sortedPojo[first].incorps || {});
							// console.log('second', second, sortedPojo[second].incorps) || {};

							if (sortedPojo[second].defined && !sortedPojo[first].defined) {
								// retain second form (as target)
								arrayOfWordPairs.push([second, first]);
								util.mergeObjects(
									sortedPojo[second].incorps,
									sortedPojo[first].incorps
								);
								util.mergeObjects(
									sortedPojo[second].usedBy,
									sortedPojo[first].usedBy
								);
								delete sortedPojo[first];

							} else if (!sortedPojo[second].defined) {
								// retain first form (as target)
								arrayOfWordPairs.push([first, second]);
								util.mergeObjects(
									sortedPojo[first].incorps,
									sortedPojo[second].incorps
								);
								util.mergeObjects(
									sortedPojo[first].usedBy,
									sortedPojo[second].usedBy
								);
								delete sortedPojo[second];
							}
						}
					}
				});

				// merge plural/singular terms contained within each object in sortedPojo
				arrayOfWordPairs.forEach(function (wordPair) {
					console.log(wordPair);
					Object.keys(sortedPojo).forEach(function (term) {
						util.mergeWithinObject(sortedPojo[term], wordPair);
					});
				});

				console.log('arrayOfWordPairs', arrayOfWordPairs);

				/* ANALYSIS PASS */
				var analysisPojo = Object.create(null);
				var sortedPojoKeys = Object.keys(sortedPojo);

				/* Pick out terms that are not defined in selection */
				sortedPojoKeys.forEach(function (term) {
					if (!sortedPojo[term].defined) {
						if (userTermsAdded.indexOf(term) !== -1) {
							// unless it is one of the userTermsAdded
							sortedPojo[term].defined = 2; //use 2 instead of 1 to distinguish

						} else {
							if (!analysisPojo.notDefined) {
								// use array (instead of another object) as value
								analysisPojo.notDefined = [];
							}
							analysisPojo.notDefined.push(term);
						}
					}
				});

				/* Find circular terms */
				var circularPaths = [];

				function walker(caller, target, path, depth) {
					if (sortedPojo[caller].incorps) {
						Object.keys(sortedPojo[caller].incorps).forEach(function (n) {
							// using a deep clone of 'path' -- must do so when
							// recursively invoking walker function below
							let clone = path.slice(0);

							if (n === target) {
								// clone.push(n); //can't push n b/c that screws up removal of dupes
								circularPaths.push(clone);

							} else if (sortedPojo[n] && sortedPojo[n].incorps) {
								if (clone.length < depth && clone.indexOf(n) === -1) {
									clone.push(n);
									walker(n, target, clone, depth); //recursively invoke walker
								}
							}
						});
					}
				}

				sortedPojoKeys.forEach(function (term) {
					walker(term, term, [term], 6);
				});

				analysisPojo.circular = circularPaths
					// remove dupe paths
					.filter(function (path, i, self) {
						return i === self.findIndex(function (item) {
							return item.slice(0).sort().join('') === path.slice(0).sort().join('');
						});
					})
					// add back in last path item
					.map(function (path) {
						path.push(path[0]);
						return path;
					});

				/* Pick cross-referenced definitions */
				analysisPojo.crossRefs = getCrossRefDefs(paras);
				// console.log(JSON.stringify(analysisPojo, null, 5));
				/* END HERE */

				if (!Object.keys(sortedPojo).length) {
					var header = 'Error:';
					var content = 'No definition paragraphs have been selected';
					showNotification(header, content);

					return context.sync(); //bail
				}

				var firstTableArray = util.createFirstTable(analysisPojo);
				var secondTableArray = util.createSecondTable(analysisPojo);
				var mainTableArray = util.createMainTable(sortedPojo);
				var newDoc = context.application.createDocument();
				context.load(newDoc);

				return context.sync().then(function () {
					// console.log('newDoc', newDoc);
					var firstTable = util.insertTable(newDoc.body, firstTableArray);
					firstTable.headerRowCount = 1;
					firstTable.style = 'List Table 4 - Accent 1';
					firstTable.styleFirstColumn = false;

					var secondTable = util.insertTable(newDoc.body, secondTableArray);
					secondTable.headerRowCount = 1;
					secondTable.style = 'List Table 4 - Accent 1';
					secondTable.styleFirstColumn = false;

					var mainTable = util.insertTable(newDoc.body, mainTableArray);
					mainTable.headerRowCount = 1;
					mainTable.style = 'List Table 4 - Accent 1';

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
