/*
 * (c) Copyright Ascensio System SIA 2010-2025
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
	const maxTokens = 16000;
	let apiKey = '';
	let interval = null;
	let tokenTimeot = null;
	let errTimeout = null;
	let modalTimeout = null;
	let loader = null;
	let bCreateLoader = true;
	let themeType = 'light';
	let regenerationMessageIndex = null;	//Index of the message for which a new reply is being created
	let searchModeEnabled = false;	// Regulatory search mode toggle
	let currentSearchAbortController = null;	// For cancelling search requests

	// Regulatory Search Configuration
	// Regulatory Search API - unified endpoint via agents-backend
	// All AI requests go through agents-backend for consistent architecture
	const AGENTS_BACKEND_URL = (function() {
		try {
			// Use agents-backend running on port 8000
			var hostname = window.parent.location.hostname || window.location.hostname;
			return 'http://' + hostname + ':8000';
		} catch (e) {
			return 'http://localhost:8001';
		}
	})();
	const REGULATORY_SEARCH_API = AGENTS_BACKEND_URL + '/api/regulatory-search';

	const ErrorCodes = {
		UNKNOWN: 1
	};
	const errorsMap = {
		[ErrorCodes.UNKNOWN]: {
			title: 'Something went wrong',
			description: 'Please try reloading the conversation'
		}
	};

	let scrollbarList; 

	let messagesList = {
		_list: [],
		
		_renderItemToList: function(item, index) {
			$('#chat_wrapper').removeClass('empty');

			let $chat = $('#chat');
			item.$el = $('<div class="message" style="order: ' + index + ';"></div>');
			$chat.prepend(item.$el);
			this._renderItem(item);
			$chat.scrollTop($chat[0].scrollHeight);
		},
		_renderItem: function(item) {
			item.$el.empty();
			item.$el.toggleClass('user_message', item.role == 'user');

			const me = this;
			let $messageContent = $('<div class="form-control message_content"></div>');
			let $spanMessage = $('<span class="span_message"></span>');
			let $attachedWrapper;
			let activeContent = item.getActiveContent();
			$messageContent.append($spanMessage);
			
			let c = window.markdownit();
			let htmlContent = c.render(activeContent);
			$spanMessage.css('display', 'block');
			$spanMessage.html(htmlContent);

			let plainText = htmlContent.replace(/<\/?[^>]+(>|$)/g, "").replace(/\n{3,}/g, "\n\n");
	
			if(item.role == 'user') {
				// TODO: For a future release.
				if(false && item.attachedText) {
					$attachedWrapper = $(
						'<div class="message_content_attached_wrapper collapsed">' + 
							'<div class="message_content_attached">' + 
								item.attachedText +
							'</div>' + 
							'<div class="message_content_collapse_btn noselect">' +
								'<img class="icon" draggable="false" src="' + getFormattedPathForIcon('resources/icons/light/chevron-down.png') + '"/>' +
							'</div>' +
						'</div>'
					);
					$attachedWrapper.find('.message_content_collapse_btn').on('click', function() {
						$attachedWrapper.toggleClass('collapsed');
						toggleAttachedCollapseButton($attachedWrapper);
					});
					$messageContent.prepend($attachedWrapper);
				}
			} else {
				if(item.error) {
					const errorObj = errorsMap[item.error];
					const $error = $(
						'<div class="message_content_error_title">' +
							'<img class="icon" draggable="false" src="' + getFormattedPathForIcon('resources/icons/light/error.png') + '" />' +
							'<div>' + errorObj.title + '</div>' + 
						'</div>' +
						'<div class="message_content_error_desc">' + errorObj.description + '</div>'
					);
					$messageContent.append($error);
				} else {
					let $actionButtons = $('<div class="action_buttons_list"></div>');
					actionButtons.forEach(function(button, index) {
						let buttonEl = $('<button class="action_button btn-text-default"></button>');
						buttonEl.append('<img class="icon" draggable="false" src="' + getFormattedPathForIcon(button.icon) + '"/>');
						buttonEl.on('click', function() {
							button.handler(item, activeContent, htmlContent, plainText);
						});
		
						if(button.tipOptions) {
							if(item.btnTips[index]) {
								item.btnTips[index]._deleteTooltipElement();
							}
							item.btnTips[index] = new Tooltip(buttonEl[0], button.tipOptions);
						}
						
						$actionButtons.append(buttonEl);
					});
					item.$el.append($actionButtons);
	
					if(item.content.length > 1) {
						const $repliesSwitch = $(
							'<div class="message_content_replies_switch">' + 
								'<div>' + (item.activeContentIndex + 1) + ' / ' + (item.content.length) + '</div>' +
							'</div>'
						);
	
						const $decrementBtn = $('<button><img class="decrement icon" src="' + getFormattedPathForIcon('resources/icons/light/chevron-down.png') + '"/></button>');
						item.activeContentIndex == 0 ? $decrementBtn.attr('disabled', 'disabled') : $decrementBtn.removeAttr('disabled');
						$repliesSwitch.prepend($decrementBtn);
						$decrementBtn.on('click', function() {
							item.activeContentIndex -= 1;
							me._renderItem(item);
	
						});
						
						const $incrementBtn = $('<button><img class="increment icon" src="' + getFormattedPathForIcon('resources/icons/light/chevron-down.png') + '"/></button>');
						item.activeContentIndex == item.content.length - 1 ? $incrementBtn.attr('disabled', 'disabled') : $incrementBtn.removeAttr('disabled');
						$repliesSwitch.append($incrementBtn);
						$incrementBtn.on('click', function() {
							item.activeContentIndex += 1;
							me._renderItem(item);
						});
	
						$messageContent.append($repliesSwitch);
					}
				}
			}
			item.$el.prepend($messageContent);
			if($attachedWrapper) {
				setTimeout(function() {
					toggleAttachedCollapseButton($attachedWrapper);
				}, 10);
			}
			scrollbarList.update();
		},

		set: function(array) {
			let me = this;

			array.forEach(function(item) {
				me.add(item);
			});
		},
		add: function(item) {
			const message = Object.assign({}, item);

			message.getActiveContent = function() {
				return (message.role == 'user' ? message.content : message.content[message.activeContentIndex]);
			};
			if(message.role == 'assistant') {
				message.activeContentIndex = 0;
			}
			message.btnTips = [];
			this._list.push(message)
			this._renderItemToList(message, this._list.length - 1);
		},
		pushContentForAssistant: function(messageIndex, content) {
			if(!this._list[messageIndex] || this._list[messageIndex].role != 'assistant') return;
			const message = this._list[messageIndex];
			message.content.push(content);
			message.activeContentIndex = message.content.length - 1;
			this._renderItem(message);
		},
		get: function() {
			return this._list;
		}
	};

	let attachedText = {
		set: function(text) {
			$('#attached_text_wrapper').removeClass('hidden');
			$('#attached_text').text(text);
		},
		get: function() {
			return $('#attached_text').text().trim();
		},
		clear: function() {
			$('#attached_text_wrapper').addClass('hidden', true);
			$('#attached_text').text('');
		},
		hasShow: function() {
			return !$('#attached_text_wrapper').hasClass('hidden');
		}
	};

	let actionButtons = [
		{ 
			icon: 'resources/icons/light/btn-update.png', 
			tipOptions: {
				text: 'Generate new',
				align: 'left'
			},
			handler: function(message) { 
				const messageIndex = messagesList.get().findIndex(function(item) { return item == message});
				if(messageIndex > 0) {
					regenerationMessageIndex = messageIndex;
					sendMessage(messagesList.get()[messageIndex - 1].content);
				}
			}
		},
		{ 
			icon: 'resources/icons/light/btn-copy.png', 
			tipOptions: {
				text: 'Copy',
				align: 'left'
			},
			handler: function(message, content) { 
				var prevTextareaVal = $('#input_message').val();
				$('#input_message').val(content);
				$('#input_message').select();
				document.execCommand('copy');
				$('#input_message').val(prevTextareaVal);
			}
		},
		{ 
			icon: 'resources/icons/light/btn-replace.png', 
			tipOptions: {
				text: 'Replace original text',
				align: 'left'
			},
			handler: function(message, content, htmlContent, plainText) { insertEngine('replace', plainText); }
		},
		{ 
			icon: 'resources/icons/light/btn-select-tool.png', 
			tipOptions: {
				text: 'Insert result',
				align: 'left'
			},
			handler: function(message, content, htmlContent)  { insertEngine('insert', htmlContent); }
		},
		{ 
			icon: 'resources/icons/light/btn-menu-comments.png', 
			tipOptions: {
				text: 'In comment',
				align: 'left'
			},
			handler: function(message, content, htmlContent, plainText) { insertEngine('comment', plainText); }
		},
		{ 
			icon: 'resources/icons/light/btn-ic-review.png', 
			tipOptions: {
				text: 'As review',
				align: 'left'
			},
			handler: function(message, content, htmlContent) { insertEngine('review', htmlContent);}
		}
	];
	let welcomeButtons = [
		{ text: 'Blog post', prompt: 'Blog post about' },
		{ text: 'Press release', prompt: 'Press release about' },
		{ text: 'An essay', prompt: 'An essay about' },
		{ text: 'Social media post', prompt: 'Social media post about' },
		{ text: 'Brainstorm', prompt: 'Brainstorm ideas for' },
		{ text: 'Project proposal', prompt: 'Project proposal about' },
		{ text: 'Creative story', prompt: 'Creative story about' },
		{ text: 'Make a plan', prompt: 'Make a plan about' },
		{ text: 'Get advice', prompt: 'Get advice about' }
	];


	function insertEngine(type, text) {
		window.Asc.plugin.sendToPlugin("onChatReplace", {
			type : type,
			data : text
		});
	}
	let localStorageKey = "onlyoffice_ai_chat_state";

	window.Asc.plugin.init = function() {
		scrollbarList = new PerfectScrollbar("#chat", {});
		restoreState();
		bCreateLoader = false;
		destroyLoader();

		updateTextareaSize();

		window.Asc.plugin.sendToPlugin("onWindowReady", {});

		document.getElementById('input_message_submit').addEventListener('click', function() {
			onSubmit();
		});
		document.getElementById('input_message').onkeydown = function(e) {
			if ( (e.ctrlKey || e.metaKey) && e.key === 'Enter') {
				e.target.value += '\n';
				updateTextareaSize();
			} else if (e.key === 'Enter') {
				e.preventDefault();
				e.stopPropagation();
				onSubmit();
			}
		};

		document.getElementById('input_message').addEventListener('focus', function(event) {
			$('#input_message_wrapper').addClass('focused');
		});
		document.getElementById('input_message').addEventListener('blur', function(event) {
			$('#input_message_wrapper').removeClass('focused');
		});
		document.getElementById('input_message').focus();

		document.getElementById('input_message').addEventListener('input', function(event) {
			//autosize
			updateTextareaSize();

			if (tokenTimeot) {
				clearTimeout(tokenTimeot);
				tokenTimeot = null;
			}
			tokenTimeot = setTimeout(function() {
				let text = event.target.value.trim();
				let tokens = window.Asc.OpenAIEncode(text);
				if (tokens.length > maxTokens) {
					event.target.classList.add('error_border');
				} else {
					event.target.classList.remove('error_border');
				}
				document.getElementById('cur_tokens').innerText = tokens.length;
			}, 250);
		});

		document.getElementById('attached_text_close').addEventListener('click', function() {
			attachedText.clear();
		});
		
		// TODO:
		if (true)
		{
			document.getElementById('tokens_info').style.display = "none";
		}

		document.getElementById('tokens_info').addEventListener('mouseenter', function (event) {
			event.target.children[0].classList.remove('hidden');
			if (modalTimeout) {
				clearTimeout(modalTimeout);
				modalTimeout = null;
			}
		});

		document.getElementById('tokens_info').addEventListener('mouseleave', function (event) {
			modalTimeout = setTimeout(function() {
				event.target.children[0].classList.add('hidden');
			}, 100)
		});

		document.getElementById('clear_history').onclick = function() {
			document.getElementById('chat').innerHTML = '';
			messagesList.set([]);
			document.getElementById('total_tokens').classList.remove('err-message');
			document.getElementById('total_tokens').innerText = 0;
		};

		document.getElementById("chat_wrapper").addEventListener("click", function(e) {
			if (e.target.tagName === "A") {
				e.preventDefault();
				window.open(e.target.href, "_blank");
			}
		});

		// Search toggle listener
		var searchToggle = document.getElementById('search_toggle');
		if (searchToggle) {
			searchToggle.addEventListener('change', function(e) {
				searchModeEnabled = e.target.checked;
				updateSearchModeUI();
			});
		}
	};

	function updateSearchModeUI() {
		var placeholder = searchModeEnabled 
			? 'Search regulatory documents...' 
			: window.Asc.plugin.tr('Ask AI anything');
		document.getElementById('input_message').setAttribute('placeholder', placeholder);
	}


	function onSubmit() {
		let textarea = document.getElementById('input_message');
		if (textarea.classList.contains('error_border')){
			setError('Too many tokens in your request.');
			return;
		}
		let value = textarea.value.trim();
		if (value.length) {
			if (searchModeEnabled) {
				// Use regulatory search mode
				performRegulatorySearch(value);
			} else {
				// Use normal AI chat mode
				sendMessage(value);
			}
			textarea.value = '';
			updateTextareaSize();
			document.getElementById('cur_tokens').innerText = 0;
		}
	};

	// ============================================
	// REGULATORY SEARCH FUNCTIONS
	// ============================================

	function performRegulatorySearch(query) {
		// Add user message to chat
		messagesList.add({ role: 'user', content: query });

		// Create search results container
		var searchResultsEl = createSearchResultsUI();
		var generatedAnswer = '';
		var sources = [];
		
		console.log('Regulatory Search API URL:', REGULATORY_SEARCH_API);
		console.log('Sending query:', query);
		
		// Use XMLHttpRequest for better compatibility with iframe context
		var xhr = new XMLHttpRequest();
		xhr.open('POST', REGULATORY_SEARCH_API, true);
		xhr.setRequestHeader('Content-Type', 'application/json');
		
		var buffer = '';
		var lastProcessedIndex = 0;
		
		xhr.onprogress = function() {
			// Process new data as it arrives
			var newData = xhr.responseText.substring(lastProcessedIndex);
			lastProcessedIndex = xhr.responseText.length;
			
			buffer += newData;
			var lines = buffer.split('\n\n');
			buffer = lines.pop() || '';
			
			for (var i = 0; i < lines.length; i++) {
				var line = lines[i].trim();
				if (line.startsWith('data: ')) {
					var jsonStr = line.slice(6).trim();
					if (!jsonStr) continue;
					try {
						var data = JSON.parse(jsonStr);
						console.log('SSE event:', data.type);
						
						if (data.type === 'answerChunk') {
							generatedAnswer += data.text;
							updateSearchAnswer(searchResultsEl, generatedAnswer, true);
						} else if (data.type === 'sources') {
							sources = data.sources || [];
							renderSearchSources($(searchResultsEl), sources);
						} else if (data.type === 'error') {
							setError(data.message);
						} else {
							handleSearchEvent(data, searchResultsEl, sources, function(text) {
								generatedAnswer += text;
								updateSearchAnswer(searchResultsEl, generatedAnswer, true);
							});
						}
					} catch (e) {
						console.error('Error parsing SSE data:', e, 'Line:', line);
					}
				}
			}
		};
		
		xhr.onload = function() {
			console.log('XHR complete. Status:', xhr.status);
			console.log('Total answer length:', generatedAnswer.length);
			console.log('Response text length:', xhr.responseText.length);
			
			// Check for HTTP errors
			if (xhr.status !== 200) {
				console.error('HTTP error:', xhr.status, xhr.statusText);
				setError('Search failed: ' + xhr.status + ' ' + xhr.statusText);
				return;
			}
			
			// Process any remaining buffer
			if (buffer.trim()) {
				var remainingLines = buffer.split('\n\n');
				for (var i = 0; i < remainingLines.length; i++) {
					var line = remainingLines[i].trim();
					if (line.startsWith('data: ')) {
						try {
							var data = JSON.parse(line.slice(6).trim());
							console.log('Processing remaining event:', data.type);
							if (data.type === 'answerChunk') {
								generatedAnswer += data.text;
							} else if (data.type === 'sources' && sources.length === 0) {
								sources = data.sources || [];
								renderSearchSources($(searchResultsEl), sources);
							}
						} catch (e) {
							console.error('Error parsing remaining data:', e);
						}
					}
				}
			}
			
			// Final update - remove cursor
			if (generatedAnswer && searchResultsEl) {
				console.log('Updating final answer, length:', generatedAnswer.length);
				updateSearchAnswer(searchResultsEl, generatedAnswer, false);
				
				// Add the complete response as an assistant message for history
				var formattedAnswer = formatSearchAnswerForHistory(generatedAnswer, sources);
				messagesList.add({ role: 'assistant', content: [formattedAnswer] });
			} else {
				console.warn('No answer generated or container missing. Answer:', generatedAnswer.length, 'Container:', !!searchResultsEl);
				if (!generatedAnswer) {
					setError('No answer was generated. Please try again.');
				}
			}
			
			// Mark all steps as complete
			$(searchResultsEl).find('.search_step').each(function() {
				$(this).removeClass('pending active').addClass('complete');
				$(this).find('.search_step_icon').removeClass('pending active').addClass('complete');
			});
			
			console.log('Search completed successfully');
		};
		
		xhr.onerror = function() {
			console.error('XHR error:', xhr.status, xhr.statusText);
			console.error('Response:', xhr.responseText);
			setError('Search failed. Please check your connection and try again.');
		};
		
		xhr.send(JSON.stringify({ query: query }));
	}

	function createSearchResultsUI() {
		$('#chat_wrapper').removeClass('empty');
		
		var $chat = $('#chat');
		var index = messagesList.get().length;
		
		var $container = $('<div class="message search_results_container" style="order: ' + index + ';">' +
			'<div class="search_steps">' +
				'<div class="search_step" data-step="analyze">' +
					'<span class="search_step_icon pending"></span>' +
					'<span>Analyzing</span>' +
				'</div>' +
				'<div class="search_step" data-step="decompose">' +
					'<span class="search_step_icon pending"></span>' +
					'<span>Breaking down</span>' +
				'</div>' +
				'<div class="search_step" data-step="search">' +
					'<span class="search_step_icon pending"></span>' +
					'<span>Searching</span>' +
				'</div>' +
				'<div class="search_step" data-step="synthesize">' +
					'<span class="search_step_icon pending"></span>' +
					'<span>Writing</span>' +
				'</div>' +
			'</div>' +
			'<div class="search_subqueries"></div>' +
			'<div class="search_sources"></div>' +
			'<div class="search_answer" style="display: none;">' +
				'<div class="search_answer_title">Answer</div>' +
				'<div class="search_answer_content"></div>' +
			'</div>' +
		'</div>');
		
		$chat.prepend($container);
		$chat.scrollTop($chat[0].scrollHeight);
		
		return $container[0];
	}

	function getSearchContainer(container) {
		// Safely get a jQuery container object
		if (container && container.jquery) {
			return container;
		}
		if (container && container.nodeType) {
			return $(container);
		}
		// Fallback to last search results container
		return $('.search_results_container').last();
	}

	function handleSearchEvent(data, container, sources, onAnswerChunk) {
		var $container = getSearchContainer(container);
		
		if (!$container || !$container.length) {
			console.error('handleSearchEvent: No valid container found');
			return;
		}
		
		switch (data.type) {
			case 'step':
				updateSearchStep($container, data.step, data.status);
				break;
				
			case 'subQueries':
				renderSubQueries($container, data.subQueries);
				break;
				
			case 'subQueryStatus':
				updateSubQueryStatus($container, data.id, data.status, data.resultCount);
				break;
				
			case 'sources':
				sources.length = 0;
				for (var i = 0; i < data.sources.length; i++) {
					sources.push(data.sources[i]);
				}
				renderSearchSources($container, sources);
				break;
				
			case 'answerChunk':
				onAnswerChunk(data.text);
				break;
				
			case 'done':
				$container.find('.search_step').each(function() {
					$(this).removeClass('active').addClass('complete');
					$(this).find('.search_step_icon').removeClass('pending active').addClass('complete');
				});
				break;
				
			case 'error':
				setError(data.message);
				break;
		}
		
		$('#chat').scrollTop($('#chat')[0].scrollHeight);
	}

	function updateSearchStep($container, stepId, status) {
		var $step = $container.find('[data-step="' + stepId + '"]');
		var $icon = $step.find('.search_step_icon');
		
		$step.removeClass('pending active complete').addClass(status);
		$icon.removeClass('pending active complete').addClass(status);
	}

	function renderSubQueries($container, subQueries) {
		var $subQueriesEl = $container.find('.search_subqueries');
		$subQueriesEl.empty();
		
		for (var i = 0; i < subQueries.length; i++) {
			var sq = subQueries[i];
			$subQueriesEl.append('<div class="search_subquery" data-id="' + sq.id + '">' +
				'<span class="subquery_status">○</span> ' +
				sq.query +
				(sq.intent ? ' <span class="subquery_intent">(' + sq.intent + ')</span>' : '') +
			'</div>');
		}
	}

	function updateSubQueryStatus($container, id, status, resultCount) {
		var $subQuery = $container.find('.search_subquery[data-id="' + id + '"]');
		$subQuery.removeClass('pending searching complete').addClass(status);
		
		var statusIcon = status === 'complete' ? '✓' : (status === 'searching' ? '●' : '○');
		$subQuery.find('.subquery_status').text(statusIcon);
		
		if (resultCount !== undefined) {
			$subQuery.append(' <span class="subquery_count">(' + resultCount + ' results)</span>');
		}
	}

	// Store sources globally for citation preview access
	var globalSources = [];
	
	function renderSearchSources($container, sources) {
		var $sourcesEl = $container.find('.search_sources');
		$sourcesEl.empty();
		
		// Store sources globally for citation previews
		globalSources = sources;
		
		if (sources.length === 0) return;
		
		var maxPills = 6; // Show max 6 pills, rest in expanded view
		var maxCards = Math.min(sources.length, 20);
		
		// Add header with toggle
		var $header = $('<div class="search_sources_header">' +
			'<span>Sources</span>' +
			'<span class="search_sources_toggle" data-expanded="false">View all ' + sources.length + '</span>' +
		'</div>');
		$sourcesEl.append($header);
		
		// Create horizontal scrollable pills container
		var $pillsContainer = $('<div class="search_sources_pills"></div>');
		
		for (var i = 0; i < Math.min(maxPills, sources.length); i++) {
			var source = sources[i];
			var citationNum = i + 1;
			var sourceId = source.id || 'source-' + i;
			var sourceType = (source.sourceType || source.source_type || 'DOC').toUpperCase();
			var sourceTitle = source.title || source.code || 'Source ' + citationNum;
			var sourceUrl = source.sourceUrl || source.url || '';
			
			// Get favicon or type icon
			var faviconHtml = getSourceTypeIcon(sourceType);
			
			// Get type badge class
			var typeBadgeClass = getTypeBadgeClass(sourceType);
			
			var $pill = $('<div class="search_source_pill" data-id="' + sourceId + '" data-citation-num="' + citationNum + '">' +
				'<span class="search_source_pill_num">' + citationNum + '</span>' +
				'<span class="search_source_pill_favicon">' + faviconHtml + '</span>' +
				'<div class="search_source_pill_content">' +
					'<div class="search_source_pill_title">' + escapeHtml(sourceTitle) + '</div>' +
					'<div class="search_source_pill_domain">' + escapeHtml(extractDomain(sourceUrl) || sourceType) + '</div>' +
				'</div>' +
				'<span class="search_source_type_badge ' + typeBadgeClass + '">' + sourceType + '</span>' +
			'</div>');
			
			// Add click handler to expand/show details
			(function(idx, src) {
				$pill.on('click', function(e) {
					e.stopPropagation();
					showSourceModal(src, idx + 1);
				});
			})(i, source);
			
			$pillsContainer.append($pill);
		}
		
		// Add "more" indicator if there are more sources
		if (sources.length > maxPills) {
			var $morePill = $('<div class="search_source_pill" style="background: #f3f4f6; border-style: dashed;">' +
				'<span class="search_source_pill_num" style="background: #9ca3af;">+' + (sources.length - maxPills) + '</span>' +
				'<div class="search_source_pill_content">' +
					'<div class="search_source_pill_title">More sources</div>' +
					'<div class="search_source_pill_domain">Click to view all</div>' +
				'</div>' +
			'</div>');
			$morePill.on('click', function() {
				$sourcesEl.find('.search_sources_grid').addClass('visible');
				$sourcesEl.find('.search_sources_toggle').text('Hide sources').data('expanded', true);
			});
			$pillsContainer.append($morePill);
		}
		
		$sourcesEl.append($pillsContainer);
		
		// Create expandable grid of source cards
		var $gridContainer = $('<div class="search_sources_grid"></div>');
		
		for (var i = 0; i < maxCards; i++) {
			var source = sources[i];
			var citationNum = i + 1;
			var sourceId = source.id || 'source-' + i;
			var sourceType = (source.sourceType || source.source_type || 'DOC').toUpperCase();
			var sourceTitle = source.title || source.code || 'Source ' + citationNum;
			var sourceUrl = source.sourceUrl || source.url || '';
			var snippet = source.snippet || (source.fullText || source.full_text ? (source.fullText || source.full_text).substring(0, 200) : '');
			var fullText = source.fullText || source.full_text || snippet;
			
			var typeBadgeClass = getTypeBadgeClass(sourceType);
			
			var $card = $('<div class="search_source_card" data-id="' + sourceId + '" data-citation-num="' + citationNum + '">' +
				'<div class="search_source_header">' +
					'<span class="search_source_badge">' + citationNum + '</span>' +
					'<div class="search_source_info">' +
						'<div class="search_source_title">' + escapeHtml(sourceTitle) + '</div>' +
						'<div class="search_source_meta">' +
							'<span class="search_source_type ' + typeBadgeClass + '">' + sourceType + '</span>' +
							(sourceUrl ? '<span>' + escapeHtml(extractDomain(sourceUrl)) + '</span>' : '') +
						'</div>' +
					'</div>' +
				'</div>' +
				'<div class="search_source_snippet">' + escapeHtml(snippet) + '</div>' +
				'<div class="search_source_expanded">' +
					'<div class="search_source_fulltext">' + escapeHtml(fullText) + '</div>' +
					(sourceUrl ? '<a href="' + escapeHtml(sourceUrl) + '" target="_blank" rel="noopener noreferrer" class="search_source_link">Open source ↗</a>' : '') +
				'</div>' +
			'</div>');
			
			$card.on('click', function() {
				$(this).toggleClass('expanded');
			});
			
			$gridContainer.append($card);
		}
		
		$sourcesEl.append($gridContainer);
		
		// Toggle handler for showing/hiding grid
		$header.find('.search_sources_toggle').on('click', function() {
			var isExpanded = $(this).data('expanded');
			if (isExpanded) {
				$gridContainer.removeClass('visible');
				$(this).text('View all ' + sources.length).data('expanded', false);
			} else {
				$gridContainer.addClass('visible');
				$(this).text('Hide sources').data('expanded', true);
			}
		});
	}
	
	function getSourceTypeIcon(sourceType) {
		var icons = {
			'ICH': '🏛️',
			'FDA': '🇺🇸',
			'EMA': '🇪🇺',
			'PSG': '📋',
			'GUIDANCE': '📜',
			'REGULATION': '⚖️',
			'DOC': '📄'
		};
		return icons[sourceType] || icons['DOC'];
	}
	
	function getTypeBadgeClass(sourceType) {
		var classes = {
			'ICH': 'ich',
			'FDA': 'fda',
			'EMA': 'ema',
			'PSG': 'psg'
		};
		return classes[sourceType] || 'doc';
	}
	
	function extractDomain(url) {
		if (!url) return '';
		try {
			var domain = url.replace(/^https?:\/\//, '').split('/')[0];
			return domain.length > 30 ? domain.substring(0, 27) + '...' : domain;
		} catch (e) {
			return '';
		}
	}
	
	function showSourceModal(source, citationNum) {
		var sourceUrl = source.sourceUrl || source.url || '';
		var sourceTitle = source.title || source.code || 'Source';
		var fullText = source.fullText || source.full_text || source.snippet || '';
		
		// If URL exists, open in new tab
		if (sourceUrl) {
			window.open(sourceUrl, '_blank');
		} else {
			// Otherwise, scroll to the card and expand it
			scrollToSourceCard(citationNum);
		}
	}

	function updateSearchAnswer(container, answer, showCursor) {
		var $container = getSearchContainer(container);
		
		if (!$container || !$container.length) {
			console.error('updateSearchAnswer: No valid container found');
			return;
		}
		
		var $answerEl = $container.find('.search_answer');
		var $contentEl = $answerEl.find('.search_answer_content');
		
		console.log('updateSearchAnswer called, answer length:', answer.length, 'found answer el:', $answerEl.length);
		
		if ($answerEl.length === 0) {
			console.error('Could not find .search_answer element');
			return;
		}
		
		// Ensure the answer section is visible
		$answerEl.css('display', 'block');
		$answerEl.show();
		
		var formattedAnswer = formatAnswerWithCitations(answer);
		if (showCursor) {
			formattedAnswer += '<span class="search_cursor"></span>';
		}
		
		$contentEl.html(formattedAnswer);
		
		// Also update the synthesize step to show it's active
		var $synthesizeStep = $container.find('[data-step="synthesize"]');
		if ($synthesizeStep.length) {
			$synthesizeStep.removeClass('pending').addClass('active');
			$synthesizeStep.find('.search_step_icon').removeClass('pending').addClass('active');
		}
		
		// Scroll to show answer
		var $chat = $('#chat');
		$chat.scrollTop($chat[0].scrollHeight);
		
		// Also update scrollbar
		if (scrollbarList) {
			scrollbarList.update();
		}
	}

	// Strip internal processing tags from backend response
	function stripInternalTags(text) {
		if (!text) return '';
		
		// Remove search quality reflection blocks and their content
		text = text.replace(/<search_quality_reflection>[\s\S]*?<\/search_quality_reflection>/gi, '');
		
		// Remove search quality score blocks and their content
		text = text.replace(/<search_quality_score>[\s\S]*?<\/search_quality_score>/gi, '');
		
		// Remove standalone result tags
		text = text.replace(/<\/?result>/gi, '');
		
		// Remove thinking tags if present
		text = text.replace(/<thinking>[\s\S]*?<\/thinking>/gi, '');
		
		// Remove query tags
		text = text.replace(/<\/?query>/gi, '');
		
		// Remove context tags
		text = text.replace(/<context>[\s\S]*?<\/context>/gi, '');
		
		// Remove any other XML-like internal tags with underscores
		text = text.replace(/<\/?[a-z_]+_[a-z_]+>/gi, '');
		
		// Clean up multiple newlines left behind
		text = text.replace(/\n{3,}/g, '\n\n');
		
		// Trim leading/trailing whitespace
		text = text.trim();
		
		return text;
	}

	function formatAnswerWithCitations(text) {
		if (!text) return '';
		
		// First, strip any internal processing tags from the backend
		text = stripInternalTags(text);
		
		// Then escape HTML to prevent XSS
		var formatted = escapeHtml(text);
		
		// Handle markdown links: [text](url) - must be done BEFORE citations to avoid conflicts
		// This regex matches [text](url) but not [1] citations
		formatted = formatted.replace(/\[([^\]]+)\]\(([^)]+)\)/g, function(match, linkText, url) {
			// Check if it's a citation number (just digits)
			if (/^\d+$/.test(linkText.trim())) {
				// It's a citation, handle separately
				return match;
			}
			// It's a real link
			return '<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener noreferrer" class="answer-link">' + linkText + '</a>';
		});
		
		// Handle citations: [1], [2], etc. - clickable badges with hover preview
		formatted = formatted.replace(/\[(\d+)\]/g, function(match, num) {
			return '<span class="citation" data-citation-num="' + num + '" onmouseenter="showCitationPreview(event, ' + num + ')" onmouseleave="hideCitationPreview()" onclick="scrollToSourceCard(' + num + ')">' + num + '</span>';
		});
		
		// Handle bold text: **text**
		formatted = formatted.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
		
		// Handle line breaks
		formatted = formatted.replace(/\n\n/g, '</p><p>');
		formatted = formatted.replace(/\n/g, '<br>');
		
		// Wrap in paragraph if not already wrapped
		if (!formatted.startsWith('<p>') && !formatted.startsWith('<a') && !formatted.startsWith('<strong')) {
			formatted = '<p>' + formatted + '</p>';
		}
		
		return formatted;
	}
	
	// Citation preview tooltip element
	var $citationPreview = null;
	var citationPreviewTimeout = null;
	
	function createCitationPreviewElement() {
		if ($citationPreview) return;
		
		$citationPreview = $('<div class="citation-preview">' +
			'<div class="citation-preview-header">' +
				'<span class="citation-preview-num"></span>' +
				'<div class="citation-preview-info">' +
					'<div class="citation-preview-title"></div>' +
					'<div class="citation-preview-meta"></div>' +
				'</div>' +
			'</div>' +
			'<div class="citation-preview-snippet"></div>' +
			'<a class="citation-preview-link" target="_blank" rel="noopener noreferrer">Open source ↗</a>' +
		'</div>');
		
		$('body').append($citationPreview);
	}
	
	function showCitationPreview(event, citationNum) {
		if (citationPreviewTimeout) {
			clearTimeout(citationPreviewTimeout);
		}
		
		var sourceIndex = citationNum - 1;
		if (sourceIndex < 0 || sourceIndex >= globalSources.length) return;
		
		var source = globalSources[sourceIndex];
		if (!source) return;
		
		createCitationPreviewElement();
		
		var sourceTitle = source.title || source.code || 'Source ' + citationNum;
		var sourceType = (source.sourceType || source.source_type || 'DOC').toUpperCase();
		var sourceUrl = source.sourceUrl || source.url || '';
		var snippet = source.snippet || (source.fullText || source.full_text ? (source.fullText || source.full_text).substring(0, 250) + '...' : '');
		
		$citationPreview.find('.citation-preview-num').text(citationNum);
		$citationPreview.find('.citation-preview-title').text(sourceTitle);
		$citationPreview.find('.citation-preview-meta').html(
			'<span class="search_source_type_badge ' + getTypeBadgeClass(sourceType) + '">' + sourceType + '</span>' +
			(sourceUrl ? ' · ' + extractDomain(sourceUrl) : '')
		);
		$citationPreview.find('.citation-preview-snippet').text(snippet);
		
		if (sourceUrl) {
			$citationPreview.find('.citation-preview-link').attr('href', sourceUrl).show();
		} else {
			$citationPreview.find('.citation-preview-link').hide();
		}
		
		// Position the preview near the citation
		var $target = $(event.target);
		var targetRect = event.target.getBoundingClientRect();
		var previewWidth = 320;
		var previewHeight = $citationPreview.outerHeight() || 180;
		
		var left = targetRect.left + targetRect.width / 2 - previewWidth / 2;
		var top = targetRect.bottom + 8;
		
		// Adjust if it goes off screen
		if (left < 10) left = 10;
		if (left + previewWidth > window.innerWidth - 10) {
			left = window.innerWidth - previewWidth - 10;
		}
		if (top + previewHeight > window.innerHeight - 10) {
			top = targetRect.top - previewHeight - 8;
		}
		
		$citationPreview.css({
			left: left + 'px',
			top: top + 'px'
		});
		
		// Show with animation
		setTimeout(function() {
			$citationPreview.addClass('visible');
		}, 50);
	}
	
	function hideCitationPreview() {
		citationPreviewTimeout = setTimeout(function() {
			if ($citationPreview) {
				$citationPreview.removeClass('visible');
			}
		}, 150);
	}
	
	// Make citation functions available globally
	window.showCitationPreview = showCitationPreview;
	window.hideCitationPreview = hideCitationPreview;
	
	// Function to scroll to a specific source card by citation number
	function scrollToSourceCard(citationNum) {
		var $container = $('.search_results_container').last();
		if (!$container.length) return;
		
		// First, make sure the grid is visible
		var $grid = $container.find('.search_sources_grid');
		if (!$grid.hasClass('visible')) {
			$grid.addClass('visible');
			$container.find('.search_sources_toggle').text('Hide sources').data('expanded', true);
		}
		
		// Try to find by data-citation-num attribute first (more reliable)
		var $targetCard = $container.find('.search_source_card[data-citation-num="' + citationNum + '"]');
		var $targetPill = $container.find('.search_source_pill[data-citation-num="' + citationNum + '"]');
		
		// Fallback to index-based if attribute not found
		if (!$targetCard.length) {
			var $sourceCards = $container.find('.search_source_card');
			var targetIndex = citationNum - 1; // Convert to 0-based index
			if (targetIndex >= 0 && targetIndex < $sourceCards.length) {
				$targetCard = $sourceCards.eq(targetIndex);
			}
		}
		
		if ($targetCard.length) {
			// Expand the card if collapsed
			if (!$targetCard.hasClass('expanded')) {
				$targetCard.addClass('expanded');
			}
			
			// Scroll to the card smoothly
			var scrollContainer = $('#chat');
			if (scrollContainer.length) {
				var cardTop = $targetCard.position().top + scrollContainer.scrollTop() - 20;
				scrollContainer.animate({
					scrollTop: cardTop
				}, 400);
			} else {
				// Fallback to window scroll
				$('html, body').animate({
					scrollTop: $targetCard.offset().top - 100
				}, 400);
			}
			
			// Highlight briefly
			$targetCard.addClass('citation-highlight');
			if ($targetPill.length) {
				$targetPill.addClass('citation-highlight');
			}
			setTimeout(function() {
				$targetCard.removeClass('citation-highlight');
				if ($targetPill.length) {
					$targetPill.removeClass('citation-highlight');
				}
			}, 2500);
		}
	}
	
	// Make scrollToSourceCard available globally
	window.scrollToSourceCard = scrollToSourceCard;

	function formatSearchAnswerForHistory(answer, sources) {
		var result = answer + '\n\n---\n**Sources:**\n';
		var maxSources = Math.min(sources.length, 8);
		for (var i = 0; i < maxSources; i++) {
			var source = sources[i];
			result += '[' + (i + 1) + '] ' + (source.title || source.code || 'Source') + '\n';
		}
		return result;
	}

	function escapeHtml(text) {
		if (!text) return '';
		var div = document.createElement('div');
		div.textContent = text;
		return div.innerHTML;
	}

	function updateStartPanel() {
		updateWelcomeText();
		renderWelcomeButtons();
	};

	function updateWelcomeText() {
		let welcomeText = window.Asc.plugin.tr('Welcome');
		if(window.Asc.plugin.info.userName) {
			welcomeText += ', ' + window.Asc.plugin.info.userName;
		}
		welcomeText += '!';
		$("#welcome_text").prepend('<span>' + welcomeText + '</span>');
	};

	function renderWelcomeButtons() {
		welcomeButtons.forEach(function(button) {
			let addedEl = $('<button class="welcome_button btn-text-default noselect">' + button.text + '</button>');
			addedEl.on('click', function() {
				$('#input_message').val(button.prompt + ' ').focus();
			});
			$('#welcome_buttons_list').append(addedEl);
		});
	};

	function updateTextareaSize() {
		let textarea = $('#input_message')[0];
		if(textarea) {
			textarea.style.height = "auto";
			textarea.style.height = Math.min(textarea.scrollHeight, 98) +2 + "px";
		}
	};

	function setState(state) {
		window.localStorage.setItem(localStorageKey, JSON.stringify(state));
	};

	function getState() {
		let state = window.localStorage.getItem(localStorageKey);
		return state ? JSON.parse(state) : null;
	};

	function restoreState() {
		let state = getState();
		if(!state) return;

		if(state.messages) {
			messagesList.set(state.messages);
		}
		if(state.inputValue) {
			document.getElementById('input_message').value = state.inputValue;
		}
		if(state.attachedText) {
			attachedText.set(state.attachedText);
		}
	};

	function sendMessage(text) {
		const isRegenerating = regenerationMessageIndex !== null;
		const message = { role: 'user', content: text };

		if (attachedText.hasShow()) {
			message.attachedText = attachedText.get();
			attachedText.clear();
		}
		if (!isRegenerating) {
			messagesList.add(message);
			createTyping();
		}

		let list = isRegenerating 
			? messagesList.get().slice(0, regenerationMessageIndex) 
			: messagesList.get();
		
		//Remove the errors and user messages that caused the error
		list = list.filter(function(item, index) {
			const nextItem = list[index + 1]
			return !item.error && !(nextItem && nextItem.error);
		});	
		list = list.map(function(item) {
			return { role: item.role, content: item.getActiveContent() }
		});

		window.Asc.plugin.sendToPlugin("onChatMessage", list);	
	};

	function createTyping() {
		let chatEl = $('#chat');
		let messageEl = $('<div id="loading" class="message" style="order: ' +  messagesList.get().length + ';"></div>');
		let spanMessageEl = $('<div class="span_message"></div>');
		spanMessageEl.text(window.Asc.plugin.tr('Thinking'));
		messageEl.append(spanMessageEl);
		chatEl.prepend(messageEl);
		chatEl.scrollTop(chatEl[0].scrollHeight);
		interval = setInterval(function() {
			let countDots = (spanMessageEl.text().match(/\./g) || []).length;
			countDots = countDots < 3 ? countDots + 1 : 0;
			spanMessageEl.text(window.Asc.plugin.tr('Thinking') + Array(countDots + 1).join("."));
		}, 500);
	};

	function removeTyping() {
		clearInterval(interval);
		interval = null;
		let element = document.getElementById('loading');
		element && element.remove();
		return;
	};

	function createLoader() {
		$('#loader-container').removeClass( "hidden" );
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = showLoader($('#loader-container')[0], window.Asc.plugin.tr('Loading...'));
	};

	function destroyLoader() {
		document.getElementById('chat_window').classList.remove('hidden');
		$('#loader-container').addClass( "hidden" )
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = undefined;
	};

	function setError(error) {
		document.getElementById('lb_err').innerHTML = window.Asc.plugin.tr(error);
		document.getElementById('div_err').classList.remove('hidden');
		if (errTimeout) {
			clearTimeout(errTimeout);
			errTimeout = null;
		}
		errTimeout = setTimeout(clearError, 5000);
	};

	function clearError() {
		document.getElementById('div_err').classList.add('hidden');
		document.getElementById('lb_err').innerHTML = '';
	};

	function getFormattedPathForIcon(path) {
		path = path.replace(/\/(light|dark)\//, '/' + themeType + '/');
		path = path.replace(/(\.\w+)$/, getZoomSuffixForImage() + '$1');
		return path;
	}

	//Toggle the hide of the button to collapse attached text 
	function toggleAttachedCollapseButton($wrapper) {
		const $content = $wrapper.find('.message_content_attached');
		const $btn = $wrapper.find('.message_content_collapse_btn');
		const needCollapse = $content.height() < $content[0].scrollHeight;
		const isCollapsed = $wrapper.hasClass('collapsed');
		$btn.toggleClass('hidden', !needCollapse && isCollapsed);
	}

	function onResize () {
		updateTextareaSize();

		scrollbarList && scrollbarList.update();

		$('.message_content_attached_wrapper').each(function(index, el) {
			toggleAttachedCollapseButton($(el));
		});

		$('img').each(function() {
			var el = $(this);
			var src = $(el).attr('src');
			if(!src.includes('resources/icons/')) return;
	
			var srcParts = src.split('/');
			var fileNameWithRatio = srcParts.pop();
			var clearFileName = fileNameWithRatio.replace(/@\d+(\.\d+)?x/, '');
			var newFileName = clearFileName;
			newFileName = clearFileName.replace(/(\.[^/.]+)$/, getZoomSuffixForImage() + '$1');
			srcParts.push(newFileName);
			el.attr('src', srcParts.join('/'));
		});
	}

	function getZoomSuffixForImage() {
		var ratio = Math.round(window.devicePixelRatio / 0.25) * 0.25;
		ratio = Math.max(ratio, 1);
		ratio = Math.min(ratio, 2);
		if(ratio == 1) return ''
		else {
			return '@' + ratio + 'x';
		}
	}

	function onThemeChanged(theme) {
		bCreateLoader = false;
		window.Asc.plugin.onThemeChangedBase(theme);

		themeType = theme.type || 'light';
		updateBodyThemeClasses(theme.type, theme.name);
		updateThemeVariables(theme);

		$('img.icon').each(function() {
			var src = $(this).attr('src');
			var newSrc = src.replace(/(icons\/)([^\/]+)(\/)/, '$1' + themeType + '$3');
			$(this).attr('src', newSrc);
		});
	}

	window.addEventListener("resize", onResize);
	onResize();

	window.Asc.plugin.onTranslate = function() {
		if (bCreateLoader)
			createLoader();
		let elements = document.querySelectorAll('.i18n');

		elements.forEach(function(element) {
			element.innerText = window.Asc.plugin.tr(element.innerText);
		});

		// Textarea
		document.getElementById('input_message').setAttribute('placeholder', window.Asc.plugin.tr('Ask AI anything'));

		//Action buttons
		// In this method info object must be exist
		if (this.info && this.info.editorType !== "word") {
			for (let i = actionButtons.length - 1; i >= 0; --i) {
				if (actionButtons[i].tipOptions.text == "As review") {
					actionButtons.splice(i, 1);
					break;
				}
			}
		}

		actionButtons.forEach(function(button) {
			button.tipOptions.text = window.Asc.plugin.tr(button.tipOptions.text);
		});

		//Welcome buttons
		welcomeButtons.forEach(function(button) {
			button.text = window.Asc.plugin.tr(button.text);
			button.prompt = window.Asc.plugin.tr(button.prompt);
		});

		//Reply errors
		for (var key in errorsMap) {
			if (errorsMap.hasOwnProperty(key)) {
				const errorItem = errorsMap[key];
				errorItem.text = window.Asc.plugin.tr(errorItem.title);
				errorItem.description = window.Asc.plugin.tr(errorItem.description);
			}
		}

		updateStartPanel();
	};

	window.Asc.plugin.onThemeChanged = onThemeChanged;

	window.Asc.plugin.attachEvent("onChatReply", function(reply) {
		let errorCode = null;
		if(!reply.trim()) {
			errorCode = ErrorCodes.UNKNOWN;
		}

		if(regenerationMessageIndex) {
			if(!errorCode) {
				messagesList.pushContentForAssistant(regenerationMessageIndex, reply);
			}
		} else {
			messagesList.add({ role: 'assistant', content: [reply], error: errorCode });
		}
		regenerationMessageIndex = null;
		
		removeTyping();
		document.getElementById('input_message').focus();
	});

	window.Asc.plugin.attachEvent("onAttachedText", function(text) {
		// For a future release.
		// attachedText.set(text);
	});

	window.Asc.plugin.attachEvent("onThemeChanged", onThemeChanged);

	window.Asc.plugin.attachEvent("onUpdateState", function() {
		setState({
			messages: messagesList.get(),
			inputValue: document.getElementById('input_message').value,
			attachedText: attachedText.hasShow() ? attachedText.get() : ''
		});
		window.Asc.plugin.sendToPlugin("onUpdateState");
	});

})(window, undefined);
