/*
 * AI Copilot - Main Script
 * Cursor-like interface with chat history, mode switching, and real backend integration
 */

(function(window, undefined) {
	'use strict';

	// ============================================
	// CONFIGURATION
	// ============================================
	var Config = {
		// Backend URL - change this to your deployed backend
		BACKEND_URL: 'http://localhost:8000',
		// Set to true to use dummy responses instead of real backend
		USE_DUMMY: false
	};

	// ============================================
	// STATE
	// ============================================
	var state = {
		currentChatId: null,
		chats: [],
		mode: 'ask', // 'ask' or 'agent'
		isGenerating: false,
		sidebarCollapsed: false,
		abortController: null // For cancelling fetch requests
	};

	// ============================================
	// DOM ELEMENTS
	// ============================================
	var elements = {};

	// ============================================
	// STORAGE
	// ============================================
	var Storage = {
		KEY: 'copilot_chats',
		
		load: function() {
			try {
				var data = localStorage.getItem(this.KEY);
				return data ? JSON.parse(data) : [];
			} catch (e) {
				return [];
			}
		},
		
		save: function(chats) {
			try {
				localStorage.setItem(this.KEY, JSON.stringify(chats));
			} catch (e) {
				console.warn('Could not save chats');
			}
		}
	};

	// ============================================
	// TOOL EXECUTOR (Frontend Tools)
	// ============================================
	var ToolExecutor = {
		tools: {
			'get_selected_text': async function() {
				// Try to get selected text from OnlyOffice
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					return new Promise(function(resolve) {
						window.Asc.plugin.executeMethod('GetSelectedText', [], function(result) {
							resolve(result || '');
						});
					});
				}
				// Fallback for testing outside OnlyOffice
				return window.getSelection ? window.getSelection().toString() : '';
			},
			
			'get_document_text': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var text = '';
							var count = doc.GetElementsCount();
							for (var i = 0; i < count; i++) {
								var elem = doc.GetElement(i);
								if (elem.GetClassType && elem.GetClassType() === 'paragraph') {
									text += elem.GetText() + '\n';
								}
							}
							return text.substring(0, params.max_length || 50000);
						}, false, false, resolve);
					});
				}
				return 'Document text not available (not in OnlyOffice)';
			},
			
			'get_document_outline': async function() {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var outline = [];
							var count = doc.GetElementsCount();
							for (var i = 0; i < count; i++) {
								var elem = doc.GetElement(i);
								if (elem.GetClassType && elem.GetClassType() === 'paragraph') {
									var style = elem.GetStyle();
									if (style) {
										var styleName = style.GetName();
										if (styleName && styleName.indexOf('Heading') === 0) {
											var level = parseInt(styleName.replace('Heading', '').trim()) || 1;
											outline.push({ text: elem.GetText(), level: level });
										}
									}
								}
							}
							return outline;
						}, false, false, resolve);
					});
				}
				return [];
			},
			
			'search_document': async function(params) {
				// Simplified search - in real implementation use OnlyOffice search API
				return [{ text: 'Search not implemented', context: '', position: 0 }];
			},
			
			'insert_text': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					var text = params.text;
					var format = params.format || 'markdown';
					
					if (format === 'html' || format === 'markdown') {
						// Convert markdown to HTML if needed
						var html = format === 'markdown' ? markdownToHtml(text) : text;
						return new Promise(function(resolve) {
							window.Asc.plugin.executeMethod('PasteHtml', [html], function() {
								resolve({ success: true, message: 'Text inserted' });
							});
						});
					} else {
						return new Promise(function(resolve) {
							window.Asc.plugin.executeMethod('PasteText', [text], function() {
								resolve({ success: true, message: 'Text inserted' });
							});
						});
					}
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			'replace_selection': async function(params) {
				// In OnlyOffice, replace selection is done by pasting over selection
				return ToolExecutor.tools['insert_text'](params);
			},
			
			'get_content_controls': async function() {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var controls = doc.GetAllContentControls();
							var result = [];
							for (var i = 0; i < controls.length; i++) {
								var cc = controls[i];
								result.push({
									tag: cc.GetTag() || '',
									title: cc.GetLabel() || '',
									value: cc.GetText() || ''
								});
							}
							return result;
						}, false, false, resolve);
					});
				}
				return [];
			},
			
			'fill_content_control': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var controls = doc.GetAllContentControls();
							for (var i = 0; i < controls.length; i++) {
								var cc = controls[i];
								if (cc.GetTag() === params.tag) {
									// Select the content control and set text
									cc.SetText(params.value);
									return { success: true, message: 'Content control filled' };
								}
							}
							return { success: false, error: 'Content control not found' };
						}, false, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			'add_comment': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					return new Promise(function(resolve) {
						window.Asc.plugin.executeMethod('AddComment', [{ Text: params.text }], function(result) {
							resolve({ success: true, comment_id: result || 'comment_added' });
						});
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			}
		},
		
		execute: async function(toolName, params) {
			if (!this.tools[toolName]) {
				return { success: false, error: 'Unknown tool: ' + toolName };
			}
			try {
				var result = await this.tools[toolName](params || {});
				return { success: true, result: result };
			} catch (e) {
				console.error('Tool execution error:', e);
				return { success: false, error: e.message };
			}
		}
	};

	// Simple markdown to HTML converter for insert_text
	function markdownToHtml(md) {
		if (!md) return '';
		var html = md;
		// Bold
		html = html.replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
		// Italic
		html = html.replace(/\*(.+?)\*/g, '<i>$1</i>');
		// Headers
		html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
		html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
		html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');
		// Lists
		html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
		html = html.replace(/(<li>.*<\/li>)+/g, '<ul>$&</ul>');
		// Line breaks
		html = html.replace(/\n/g, '<br>');
		return html;
	}

	// ============================================
	// DUMMY RESPONSES (for testing without backend)
	// ============================================
	var DummyResponses = {
		askMode: [
			{
				type: 'text',
				content: "I can help you understand your document. This appears to be a regulatory document. Would you like me to:\n\n1. **Summarize** the key points\n2. **Explain** specific sections\n3. **Find** particular information\n\nJust let me know what you need!"
			},
			{
				type: 'text',
				content: "Based on the document structure, I can see several important sections. The document follows standard regulatory formatting with clear headings and content controls for template fields.\n\nIs there anything specific you'd like me to explain?"
			},
			{
				type: 'text',
				content: "I've analyzed the selected text. Here's my understanding:\n\n- The section discusses compliance requirements\n- Key terms are defined in the glossary\n- References point to external regulatory guidelines\n\nWould you like more details on any of these points?"
			}
		],
		agentMode: [
			// Response with thinking and tool calls
			[
				{ type: 'thinking', content: "The user wants me to help with the document. Let me first check what's currently selected in the editor to understand the context better." },
				{ type: 'tool_call', name: 'get_selected_text', params: {}, status: 'running' },
				{ type: 'tool_result', name: 'get_selected_text', result: '"The study protocol adheres to ICH-GCP guidelines and local regulatory requirements."', status: 'success' },
				{ type: 'thinking', content: "I have the selected text. Now I'll search for related sections in the document to provide context." },
				{ type: 'tool_call', name: 'search_document', params: { query: 'ICH-GCP guidelines' }, status: 'running' },
				{ type: 'tool_result', name: 'search_document', result: 'Found 3 mentions in sections: Introduction, Methods, Compliance', status: 'success' },
				{ type: 'text', content: "I found the selected text refers to **ICH-GCP guidelines**. This is mentioned in 3 sections of your document.\n\nWould you like me to:\n- Add more detail about specific guidelines\n- Insert a reference to the official ICH-GCP document\n- Create a summary of all compliance mentions" },
				{ type: 'checkpoint' }
			],
			// Response with document modification
			[
				{ type: 'thinking', content: "I'll insert a new section about study objectives as requested. First, let me find the right location." },
				{ type: 'tool_call', name: 'get_document_outline', params: {}, status: 'running' },
				{ type: 'tool_result', name: 'get_document_outline', result: '["1. Introduction", "2. Background", "3. Methods", "4. Results", "5. Discussion"]', status: 'success' },
				{ type: 'thinking', content: "I'll insert the new section after Introduction. Let me prepare the content." },
				{ type: 'tool_call', name: 'insert_text', params: { position: 'after:Introduction', text: '## 1.1 Study Objectives\n\nThe primary objectives of this study are:\n\n1. To evaluate the efficacy of the intervention\n2. To assess safety and tolerability\n3. To characterize pharmacokinetic properties' }, status: 'running' },
				{ type: 'tool_result', name: 'insert_text', result: 'Text inserted successfully at position after "Introduction"', status: 'success' },
				{ type: 'text', content: "I've added a new **Study Objectives** section after the Introduction. The section includes:\n\n- Primary efficacy objective\n- Safety and tolerability assessment\n- Pharmacokinetic characterization\n\nYou can review and modify the content as needed." },
				{ type: 'checkpoint' }
			],
			// Response with template filling
			[
				{ type: 'thinking', content: "The user wants to fill template fields. Let me identify all content controls in the document." },
				{ type: 'tool_call', name: 'get_content_controls', params: {}, status: 'running' },
				{ type: 'tool_result', name: 'get_content_controls', result: '["SPONSOR_NAME", "PROTOCOL_NUMBER", "STUDY_TITLE", "INVESTIGATOR_NAME", "SITE_NUMBER"]', status: 'success' },
				{ type: 'thinking', content: "Found 5 template fields. I'll fill them with appropriate placeholder values." },
				{ type: 'tool_call', name: 'fill_content_control', params: { tag: 'SPONSOR_NAME', value: 'Acme Pharmaceuticals Inc.' }, status: 'running' },
				{ type: 'tool_result', name: 'fill_content_control', result: 'Content control "SPONSOR_NAME" filled successfully', status: 'success' },
				{ type: 'tool_call', name: 'fill_content_control', params: { tag: 'PROTOCOL_NUMBER', value: 'ACM-2024-001' }, status: 'running' },
				{ type: 'tool_result', name: 'fill_content_control', result: 'Content control "PROTOCOL_NUMBER" filled successfully', status: 'success' },
				{ type: 'text', content: "I've started filling the template fields. So far I've filled:\n\n- ✅ **SPONSOR_NAME**: Acme Pharmaceuticals Inc.\n- ✅ **PROTOCOL_NUMBER**: ACM-2024-001\n\nWould you like me to continue with the remaining fields (STUDY_TITLE, INVESTIGATOR_NAME, SITE_NUMBER)?" },
				{ type: 'checkpoint' }
			]
		],
		currentIndex: { ask: 0, agent: 0 }
	};

	// ============================================
	// UTILITY FUNCTIONS
	// ============================================
	function generateId() {
		return 'chat_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
	}

	function getToolIcon(toolName) {
		var icons = {
			'get_selected_text': '📋',
			'get_document_text': '📄',
			'get_document_outline': '📑',
			'search_document': '🔍',
			'get_content_controls': '📝',
			'insert_text': '✏️',
			'replace_selection': '🔄',
			'fill_content_control': '📝',
			'add_comment': '💬'
		};
		return icons[toolName] || '🔧';
	}

	function escapeHtml(text) {
		var div = document.createElement('div');
		div.textContent = text;
		return div.innerHTML;
	}

	function formatJson(obj) {
		try {
			return JSON.stringify(obj, null, 2);
		} catch (e) {
			return String(obj);
		}
	}

	// Simple markdown parser
	function parseMarkdown(text) {
		if (!text) return '';
		
		// Escape HTML first
		var html = escapeHtml(text);
		
		// Bold
		html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
		
		// Italic
		html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');
		
		// Code inline
		html = html.replace(/`(.+?)`/g, '<code>$1</code>');
		
		// Headers
		html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
		html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
		html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');
		
		// Lists
		html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
		html = html.replace(/^(\d+)\. (.+)$/gm, '<li>$2</li>');
		
		// Wrap consecutive <li> in <ul>
		html = html.replace(/(<li>.*<\/li>\n?)+/g, function(match) {
			return '<ul>' + match + '</ul>';
		});
		
		// Line breaks
		html = html.replace(/\n\n/g, '</p><p>');
		html = html.replace(/\n/g, '<br>');
		
		// Wrap in paragraph if not already wrapped
		if (!html.startsWith('<')) {
			html = '<p>' + html + '</p>';
		}
		
		return html;
	}

	// ============================================
	// CHAT MANAGEMENT
	// ============================================
	function createNewChat() {
		var chat = {
			id: generateId(),
			title: 'New Chat',
			messages: [],
			createdAt: Date.now()
		};
		state.chats.unshift(chat);
		state.currentChatId = chat.id;
		Storage.save(state.chats);
		renderChatList();
		renderMessages();
		showWelcomeScreen(true);
		updateChatTitle('New Chat');
	}

	function loadChat(chatId) {
		state.currentChatId = chatId;
		var chat = getCurrentChat();
		if (chat) {
			renderMessages();
			showWelcomeScreen(chat.messages.length === 0);
			updateChatTitle(chat.title);
			renderChatList();
		}
	}

	function deleteChat(chatId) {
		state.chats = state.chats.filter(function(c) { return c.id !== chatId; });
		Storage.save(state.chats);
		
		if (state.currentChatId === chatId) {
			if (state.chats.length > 0) {
				loadChat(state.chats[0].id);
			} else {
				createNewChat();
			}
		} else {
			renderChatList();
		}
	}

	function clearCurrentChat() {
		var chat = getCurrentChat();
		if (chat) {
			chat.messages = [];
			chat.title = 'New Chat';
			Storage.save(state.chats);
			renderMessages();
			showWelcomeScreen(true);
			updateChatTitle('New Chat');
		}
	}

	function getCurrentChat() {
		return state.chats.find(function(c) { return c.id === state.currentChatId; });
	}

	function addMessage(role, content, metadata) {
		var chat = getCurrentChat();
		if (chat) {
			var message = {
				id: generateId(),
				role: role,
				content: content,
				timestamp: Date.now(),
				metadata: metadata || {}
			};
			chat.messages.push(message);
			
			// Update chat title from first user message
			if (role === 'user' && chat.title === 'New Chat') {
				chat.title = content.substring(0, 40) + (content.length > 40 ? '...' : '');
				updateChatTitle(chat.title);
				renderChatList();
			}
			
			Storage.save(state.chats);
			return message;
		}
		return null;
	}

	// Get conversation history for backend
	function getConversationHistory() {
		var chat = getCurrentChat();
		if (!chat || !chat.messages) return [];
		
		return chat.messages
			.filter(function(m) { return m.role === 'user' || m.role === 'assistant'; })
			.slice(-20) // Last 20 messages
			.map(function(m) {
				return { role: m.role, content: m.content };
			});
	}

	// Gather document context for backend
	async function gatherDocumentContext() {
		var context = {
			selected_text: null,
			document_outline: null,
			cursor_position: 'current'
		};
		
		try {
			// Try to get selected text
			var selectedResult = await ToolExecutor.execute('get_selected_text');
			if (selectedResult.success && selectedResult.result) {
				context.selected_text = selectedResult.result;
			}
		} catch (e) {
			console.warn('Could not get selected text:', e);
		}
		
		return context;
	}

	// ============================================
	// UI RENDERING
	// ============================================
	function renderChatList() {
		var html = '';
		state.chats.forEach(function(chat) {
			var isActive = chat.id === state.currentChatId;
			html += '<div class="chat-item' + (isActive ? ' active' : '') + '" data-chat-id="' + chat.id + '">';
			html += '<span class="chat-item-title">' + escapeHtml(chat.title) + '</span>';
			html += '<button class="chat-item-delete" data-chat-id="' + chat.id + '" title="Delete chat">';
			html += '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';
			html += '</button>';
			html += '</div>';
		});
		elements.chatList.innerHTML = html;
	}

	function renderMessages() {
		var chat = getCurrentChat();
		if (!chat) return;
		
		var html = '';
		chat.messages.forEach(function(msg) {
			html += renderMessage(msg);
		});
		elements.messages.innerHTML = html;
		scrollToBottom();
	}

	function renderMessage(msg) {
		var html = '';
		
		if (msg.role === 'user') {
			html += '<div class="message user">';
			html += '<div class="message-content">' + escapeHtml(msg.content) + '</div>';
			html += '</div>';
		} else if (msg.role === 'assistant') {
			html += '<div class="message assistant">';
			
			// Render metadata items (thinking, tool calls)
			if (msg.metadata && msg.metadata.items) {
				msg.metadata.items.forEach(function(item) {
					html += renderMetadataItem(item);
				});
			}
			
			// Render text content
			if (msg.content) {
				html += '<div class="message-content">' + parseMarkdown(msg.content) + '</div>';
			}
			
			html += '</div>';
		} else if (msg.role === 'error') {
			html += '<div class="error-message">' + escapeHtml(msg.content) + '</div>';
		}
		
		return html;
	}

	function renderMetadataItem(item) {
		var html = '';
		
		if (item.type === 'thinking') {
			html += '<div class="thinking-block">';
			html += '<div class="thinking-header">';
			html += '<span class="thinking-icon">💭</span>';
			html += '<span class="thinking-label">Thinking...</span>';
			html += '<span class="thinking-toggle">▼</span>';
			html += '</div>';
			html += '<div class="thinking-content">' + escapeHtml(item.content) + '</div>';
			html += '</div>';
		} else if (item.type === 'tool_call' || item.type === 'tool_result') {
			var status = item.status || 'running';
			var statusText = status === 'success' ? '✓ Done' : (status === 'error' ? '✗ Error' : 'Running...');
			
			html += '<div class="tool-call" data-status="' + status + '">';
			html += '<div class="tool-call-header">';
			html += '<div class="tool-call-left">';
			html += '<span class="tool-icon">' + getToolIcon(item.name) + '</span>';
			html += '<span class="tool-name">' + escapeHtml(item.name) + '</span>';
			html += '<span class="tool-status ' + status + '">';
			if (status === 'running') {
				html += '<span class="spinner"></span>';
			}
			html += statusText;
			html += '</span>';
			html += '</div>';
			html += '<span class="tool-toggle">▶</span>';
			html += '</div>';
			html += '<div class="tool-call-body">';
			html += '<div class="tool-section-label">Parameters</div>';
			html += '<pre class="tool-json">' + formatJson(item.params || {}) + '</pre>';
			if (item.result !== undefined) {
				html += '<div class="tool-result-section">';
				html += '<div class="tool-section-label">Result</div>';
				html += '<pre class="tool-json">' + escapeHtml(typeof item.result === 'string' ? item.result : formatJson(item.result)) + '</pre>';
				html += '</div>';
			}
			html += '</div>';
			html += '</div>';
		} else if (item.type === 'checkpoint') {
			html += '<div class="checkpoint">';
			html += '<div class="checkpoint-line"></div>';
			html += '<div class="checkpoint-actions">';
			html += '<button class="checkpoint-btn apply">Apply Changes</button>';
			html += '<button class="checkpoint-btn revert">Revert</button>';
			html += '</div>';
			html += '</div>';
		}
		
		return html;
	}

	function appendMessageElement(html) {
		elements.messages.insertAdjacentHTML('beforeend', html);
		scrollToBottom();
	}

	function updateLastAssistantContent(content) {
		var assistantMessages = elements.messages.querySelectorAll('.message.assistant');
		if (assistantMessages.length > 0) {
			var lastMsg = assistantMessages[assistantMessages.length - 1];
			var contentEl = lastMsg.querySelector('.message-content');
			if (contentEl) {
				contentEl.innerHTML = parseMarkdown(content);
			} else {
				lastMsg.insertAdjacentHTML('beforeend', '<div class="message-content">' + parseMarkdown(content) + '</div>');
			}
		}
		scrollToBottom();
	}

	function showWelcomeScreen(show) {
		if (show) {
			elements.welcomeScreen.classList.remove('hidden');
		} else {
			elements.welcomeScreen.classList.add('hidden');
		}
	}

	function updateChatTitle(title) {
		elements.chatTitle.textContent = title;
	}

	function scrollToBottom() {
		elements.messagesContainer.scrollTop = elements.messagesContainer.scrollHeight;
	}

	function setGenerating(isGenerating) {
		state.isGenerating = isGenerating;
		elements.sendBtn.classList.toggle('hidden', isGenerating);
		elements.stopBtn.classList.toggle('hidden', !isGenerating);
		elements.chatInput.disabled = isGenerating;
	}

	// ============================================
	// MESSAGE SENDING & RESPONSE HANDLING
	// ============================================
	function sendMessage() {
		var text = elements.chatInput.value.trim();
		if (!text || state.isGenerating) return;
		
		// Clear input
		elements.chatInput.value = '';
		autoResizeTextarea();
		
		// Hide welcome screen
		showWelcomeScreen(false);
		
		// Add user message
		var userMsg = addMessage('user', text);
		appendMessageElement(renderMessage(userMsg));
		
		// Generate response
		if (Config.USE_DUMMY) {
			generateDummyResponse(text);
		} else {
			generateBackendResponse(text);
		}
	}

	// ============================================
	// REAL BACKEND INTEGRATION
	// ============================================
	async function generateBackendResponse(userText) {
		setGenerating(true);
		
		// Create abort controller for cancellation
		state.abortController = new AbortController();
		
		try {
			// Gather context
			var context = await gatherDocumentContext();
			
			// Prepare request
			var requestData = {
				session_id: state.currentChatId,
				message: userText,
				mode: state.mode,
				context: context,
				conversation_history: getConversationHistory()
			};
			
			// Create assistant message container
			appendMessageElement('<div class="message assistant"></div>');
			var msgContainer = elements.messages.querySelector('.message.assistant:last-child');
			
			// Track response data for saving
			var items = [];
			var textContent = '';
			
			// Make SSE request
			var response = await fetch(Config.BACKEND_URL + '/api/copilot/chat', {
				method: 'POST',
				headers: {
					'Content-Type': 'application/json',
					'Accept': 'text/event-stream'
				},
				body: JSON.stringify(requestData),
				signal: state.abortController.signal
			});
			
			if (!response.ok) {
				throw new Error('Backend error: ' + response.status);
			}
			
			// Read SSE stream
			var reader = response.body.getReader();
			var decoder = new TextDecoder();
			var buffer = '';
			var currentEvent = '';
			var hasContent = false;
			
			while (true) {
				var result = await reader.read();
				if (result.done) break;
				
				buffer += decoder.decode(result.value, { stream: true });
				var lines = buffer.split('\n');
				buffer = lines.pop(); // Keep incomplete line in buffer
				
				for (var i = 0; i < lines.length; i++) {
					var line = lines[i].trim();
					
					if (line.startsWith('event:')) {
						currentEvent = line.substring(6).trim();
					} else if (line.startsWith('data:')) {
						var dataStr = line.substring(5).trim();
						if (dataStr) {
							try {
								var data = JSON.parse(dataStr);
								await handleSSEEvent(currentEvent, data, msgContainer, items, function(content) {
									textContent += content;
									hasContent = true;
								});
							} catch (e) {
								console.warn('Failed to parse SSE data:', dataStr, e);
							}
						}
					}
				}
			}
			
			// Save message
			if (hasContent || items.length > 0) {
				addMessage('assistant', textContent, { items: items });
			}
			
		} catch (error) {
			if (error.name === 'AbortError') {
				console.log('Request aborted');
			} else {
				console.error('Backend error:', error);
				addMessage('error', 'Connection error: ' + error.message);
				appendMessageElement('<div class="error-message">Connection error: ' + escapeHtml(error.message) + '</div>');
			}
		} finally {
			setGenerating(false);
			state.abortController = null;
		}
	}

	async function handleSSEEvent(eventType, data, msgContainer, items, onContent) {
		switch (eventType) {
			case 'thinking':
				var thinkingItem = { type: 'thinking', content: data.content };
				items.push(thinkingItem);
				msgContainer.insertAdjacentHTML('beforeend', renderMetadataItem(thinkingItem));
				scrollToBottom();
				break;
			
			case 'tool_result_request':
				// Backend is requesting frontend to execute a tool
				var toolItem = { 
					type: 'tool_call', 
					name: data.name, 
					params: data.params, 
					status: 'running' 
				};
				items.push(toolItem);
				msgContainer.insertAdjacentHTML('beforeend', renderMetadataItem(toolItem));
				scrollToBottom();
				
				// Execute the tool locally
				var toolResult = await ToolExecutor.execute(data.name, data.params);
				
				// Send result back to backend
				try {
					await fetch(Config.BACKEND_URL + '/api/copilot/tool_result', {
						method: 'POST',
						headers: { 'Content-Type': 'application/json' },
						body: JSON.stringify({
							session_id: state.currentChatId,
							tool_call_id: data.id,
							result: toolResult.result,
							success: toolResult.success,
							error: toolResult.error
						})
					});
				} catch (e) {
					console.warn('Failed to send tool result:', e);
				}
				
				// Update UI with result
				updateToolCallResult(msgContainer, data.name, toolResult);
				
				// Update stored item
				for (var j = items.length - 1; j >= 0; j--) {
					if (items[j].name === data.name && items[j].status === 'running') {
						items[j].status = toolResult.success ? 'success' : 'error';
						items[j].result = toolResult.result;
						break;
					}
				}
				break;
			
			case 'tool_call':
				// Tool was executed (for display)
				var displayToolItem = { 
					type: 'tool_call', 
					name: data.name, 
					params: data.params, 
					status: data.status || 'success',
					result: data.result
				};
				
				// Check if we already have this tool call running
				var existingToolCall = msgContainer.querySelector('.tool-call[data-status="running"]');
				if (existingToolCall) {
					// Update existing
					updateToolCallResult(msgContainer, data.name, { 
						success: data.status === 'success', 
						result: data.result 
					});
				} else {
					items.push(displayToolItem);
					msgContainer.insertAdjacentHTML('beforeend', renderMetadataItem(displayToolItem));
				}
				scrollToBottom();
				break;
			
			case 'content':
				// Streaming text content
				var contentEl = msgContainer.querySelector('.message-content');
				if (!contentEl) {
					msgContainer.insertAdjacentHTML('beforeend', '<div class="message-content streaming"></div>');
					contentEl = msgContainer.querySelector('.message-content');
				}
				
				onContent(data.delta);
				
				// Get current text and append delta
				var currentText = contentEl.getAttribute('data-raw') || '';
				currentText += data.delta;
				contentEl.setAttribute('data-raw', currentText);
				contentEl.innerHTML = parseMarkdown(currentText);
				scrollToBottom();
				break;
			
			case 'checkpoint':
				var checkpointItem = { type: 'checkpoint' };
				items.push(checkpointItem);
				msgContainer.insertAdjacentHTML('beforeend', renderMetadataItem(checkpointItem));
				scrollToBottom();
				break;
			
			case 'error':
				console.error('Backend error:', data.message);
				msgContainer.insertAdjacentHTML('beforeend', 
					'<div class="error-message">' + escapeHtml(data.message) + '</div>');
				scrollToBottom();
				break;
			
			case 'done':
				// Remove streaming class
				var streamingEl = msgContainer.querySelector('.streaming');
				if (streamingEl) {
					streamingEl.classList.remove('streaming');
				}
				break;
		}
	}

	function updateToolCallResult(msgContainer, toolName, result) {
		var toolCalls = msgContainer.querySelectorAll('.tool-call');
		for (var i = toolCalls.length - 1; i >= 0; i--) {
			var toolCall = toolCalls[i];
			var nameEl = toolCall.querySelector('.tool-name');
			if (nameEl && nameEl.textContent === toolName && toolCall.getAttribute('data-status') === 'running') {
				var status = result.success ? 'success' : 'error';
				toolCall.setAttribute('data-status', status);
				
				var statusEl = toolCall.querySelector('.tool-status');
				statusEl.className = 'tool-status ' + status;
				statusEl.innerHTML = status === 'success' ? '✓ Done' : '✗ Error';
				
				var bodyEl = toolCall.querySelector('.tool-call-body');
				var resultStr = typeof result.result === 'string' ? result.result : formatJson(result.result);
				bodyEl.insertAdjacentHTML('beforeend', 
					'<div class="tool-result-section"><div class="tool-section-label">Result</div>' +
					'<pre class="tool-json">' + escapeHtml(resultStr) + '</pre></div>'
				);
				break;
			}
		}
	}

	// ============================================
	// DUMMY RESPONSE FLOW (for testing)
	// ============================================
	async function generateDummyResponse(userText) {
		setGenerating(true);
		
		// Simulate delay
		await sleep(500);
		
		// Get response based on mode
		var response;
		if (state.mode === 'ask') {
			response = DummyResponses.askMode[DummyResponses.currentIndex.ask % DummyResponses.askMode.length];
			DummyResponses.currentIndex.ask++;
			await streamTextResponse(response.content);
		} else {
			var responseSequence = DummyResponses.agentMode[DummyResponses.currentIndex.agent % DummyResponses.agentMode.length];
			DummyResponses.currentIndex.agent++;
			await streamAgentResponse(responseSequence);
		}
		
		setGenerating(false);
	}

	async function streamTextResponse(text) {
		// Create assistant message container
		appendMessageElement('<div class="message assistant"><div class="message-content streaming"></div></div>');
		
		var contentEl = elements.messages.querySelector('.message.assistant:last-child .message-content');
		var displayedText = '';
		
		// Stream character by character
		for (var i = 0; i < text.length; i++) {
			if (!state.isGenerating) break;
			
			displayedText += text[i];
			contentEl.innerHTML = parseMarkdown(displayedText);
			scrollToBottom();
			
			// Variable delay for more natural feel
			var delay = text[i] === ' ' ? 20 : (text[i] === '.' || text[i] === '\n' ? 80 : 15);
			await sleep(delay);
		}
		
		// Remove streaming cursor
		contentEl.classList.remove('streaming');
		
		// Save message
		addMessage('assistant', text);
	}

	async function streamAgentResponse(sequence) {
		var items = [];
		var textContent = '';
		
		// Create assistant message container
		appendMessageElement('<div class="message assistant"></div>');
		var msgContainer = elements.messages.querySelector('.message.assistant:last-child');
		
		for (var i = 0; i < sequence.length; i++) {
			if (!state.isGenerating) break;
			
			var item = sequence[i];
			
			if (item.type === 'thinking') {
				items.push(item);
				msgContainer.insertAdjacentHTML('beforeend', renderMetadataItem(item));
				scrollToBottom();
				await sleep(800);
			} else if (item.type === 'tool_call') {
				var toolItem = { type: 'tool_call', name: item.name, params: item.params, status: 'running' };
				items.push(toolItem);
				msgContainer.insertAdjacentHTML('beforeend', renderMetadataItem(toolItem));
				scrollToBottom();
				await sleep(1000);
			} else if (item.type === 'tool_result') {
				// Update the last tool call with result
				var toolCalls = msgContainer.querySelectorAll('.tool-call');
				if (toolCalls.length > 0) {
					var lastToolCall = toolCalls[toolCalls.length - 1];
					lastToolCall.setAttribute('data-status', item.status);
					
					var statusEl = lastToolCall.querySelector('.tool-status');
					statusEl.className = 'tool-status ' + item.status;
					statusEl.innerHTML = item.status === 'success' ? '✓ Done' : '✗ Error';
					
					var bodyEl = lastToolCall.querySelector('.tool-call-body');
					bodyEl.insertAdjacentHTML('beforeend', 
						'<div class="tool-result-section"><div class="tool-section-label">Result</div>' +
						'<pre class="tool-json">' + escapeHtml(typeof item.result === 'string' ? item.result : formatJson(item.result)) + '</pre></div>'
					);
				}
				
				// Update stored item
				for (var j = items.length - 1; j >= 0; j--) {
					if (items[j].type === 'tool_call' && items[j].name === item.name) {
						items[j].status = item.status;
						items[j].result = item.result;
						break;
					}
				}
				
				scrollToBottom();
				await sleep(500);
			} else if (item.type === 'text') {
				textContent = item.content;
				msgContainer.insertAdjacentHTML('beforeend', '<div class="message-content streaming"></div>');
				var contentEl = msgContainer.querySelector('.message-content');
				
				var displayedText = '';
				for (var c = 0; c < item.content.length; c++) {
					if (!state.isGenerating) break;
					
					displayedText += item.content[c];
					contentEl.innerHTML = parseMarkdown(displayedText);
					scrollToBottom();
					
					var delay = item.content[c] === ' ' ? 15 : (item.content[c] === '.' || item.content[c] === '\n' ? 60 : 10);
					await sleep(delay);
				}
				
				contentEl.classList.remove('streaming');
			} else if (item.type === 'checkpoint') {
				items.push(item);
				msgContainer.insertAdjacentHTML('beforeend', renderMetadataItem(item));
				scrollToBottom();
			}
		}
		
		// Save message
		addMessage('assistant', textContent, { items: items });
	}

	function stopGeneration() {
		state.isGenerating = false;
		
		// Abort fetch request if in progress
		if (state.abortController) {
			state.abortController.abort();
		}
		
		// Also tell backend to stop
		if (!Config.USE_DUMMY) {
			fetch(Config.BACKEND_URL + '/api/copilot/stop', {
				method: 'POST',
				headers: { 'Content-Type': 'application/json' },
				body: JSON.stringify({ session_id: state.currentChatId })
			}).catch(function(e) {
				console.warn('Failed to send stop request:', e);
			});
		}
		
		setGenerating(false);
		
		// Remove streaming class from any streaming content
		var streamingEls = elements.messages.querySelectorAll('.streaming');
		streamingEls.forEach(function(el) {
			el.classList.remove('streaming');
		});
	}

	function sleep(ms) {
		return new Promise(function(resolve) {
			setTimeout(resolve, ms);
		});
	}

	// ============================================
	// MODE SWITCHING
	// ============================================
	function setMode(mode) {
		state.mode = mode;
		
		elements.askModeBtn.classList.toggle('active', mode === 'ask');
		elements.agentModeBtn.classList.toggle('active', mode === 'agent');
		
		elements.modeIndicator.textContent = mode === 'ask' ? 'Ask mode' : 'Agent mode';
		elements.chatInput.placeholder = mode === 'ask' 
			? 'Ask anything about your document...' 
			: 'Tell me what to do with the document...';
	}

	// ============================================
	// SIDEBAR
	// ============================================
	function toggleSidebar() {
		state.sidebarCollapsed = !state.sidebarCollapsed;
		elements.sidebar.classList.toggle('collapsed', state.sidebarCollapsed);
	}

	// ============================================
	// TEXT AREA AUTO-RESIZE
	// ============================================
	function autoResizeTextarea() {
		var el = elements.chatInput;
		el.style.height = 'auto';
		el.style.height = Math.min(el.scrollHeight, 120) + 'px';
	}

	// ============================================
	// EVENT HANDLERS
	// ============================================
	function setupEventListeners() {
		// Send message
		elements.sendBtn.addEventListener('click', sendMessage);
		elements.chatInput.addEventListener('keydown', function(e) {
			if (e.key === 'Enter' && !e.shiftKey) {
				e.preventDefault();
				sendMessage();
			}
		});
		elements.chatInput.addEventListener('input', autoResizeTextarea);
		
		// Stop generation
		elements.stopBtn.addEventListener('click', stopGeneration);
		
		// Mode switching
		elements.askModeBtn.addEventListener('click', function() { setMode('ask'); });
		elements.agentModeBtn.addEventListener('click', function() { setMode('agent'); });
		
		// Sidebar
		elements.sidebarToggle.addEventListener('click', toggleSidebar);
		elements.mobileSidebarToggle.addEventListener('click', toggleSidebar);
		
		// New chat
		elements.newChatBtn.addEventListener('click', createNewChat);
		
		// Clear chat
		elements.clearChatBtn.addEventListener('click', clearCurrentChat);
		
		// Chat list clicks
		elements.chatList.addEventListener('click', function(e) {
			var deleteBtn = e.target.closest('.chat-item-delete');
			if (deleteBtn) {
				e.stopPropagation();
				var chatId = deleteBtn.getAttribute('data-chat-id');
				deleteChat(chatId);
				return;
			}
			
			var chatItem = e.target.closest('.chat-item');
			if (chatItem) {
				var chatId = chatItem.getAttribute('data-chat-id');
				loadChat(chatId);
			}
		});
		
		// Quick actions
		elements.welcomeScreen.addEventListener('click', function(e) {
			var quickAction = e.target.closest('.quick-action');
			if (quickAction) {
				var prompt = quickAction.getAttribute('data-prompt');
				elements.chatInput.value = prompt;
				sendMessage();
			}
		});
		
		// Thinking block toggle
		elements.messages.addEventListener('click', function(e) {
			var thinkingHeader = e.target.closest('.thinking-header');
			if (thinkingHeader) {
				var block = thinkingHeader.closest('.thinking-block');
				block.classList.toggle('collapsed');
			}
			
			var toolHeader = e.target.closest('.tool-call-header');
			if (toolHeader) {
				var card = toolHeader.closest('.tool-call');
				card.classList.toggle('expanded');
			}
		});
	}

	// ============================================
	// INITIALIZATION
	// ============================================
	function init() {
		// Cache DOM elements
		elements = {
			sidebar: document.getElementById('sidebar'),
			sidebarToggle: document.getElementById('sidebarToggle'),
			mobileSidebarToggle: document.getElementById('mobileSidebarToggle'),
			newChatBtn: document.getElementById('newChatBtn'),
			chatList: document.getElementById('chatList'),
			chatTitle: document.getElementById('chatTitle'),
			clearChatBtn: document.getElementById('clearChatBtn'),
			messagesContainer: document.getElementById('messagesContainer'),
			welcomeScreen: document.getElementById('welcomeScreen'),
			messages: document.getElementById('messages'),
			chatInput: document.getElementById('chatInput'),
			sendBtn: document.getElementById('sendBtn'),
			stopBtn: document.getElementById('stopBtn'),
			askModeBtn: document.getElementById('askModeBtn'),
			agentModeBtn: document.getElementById('agentModeBtn'),
			modeIndicator: document.getElementById('modeIndicator')
		};
		
		// Load chats from storage
		state.chats = Storage.load();
		
		// Create new chat if none exist
		if (state.chats.length === 0) {
			createNewChat();
		} else {
			state.currentChatId = state.chats[0].id;
			renderChatList();
			renderMessages();
			var chat = getCurrentChat();
			showWelcomeScreen(chat && chat.messages.length === 0);
			updateChatTitle(chat ? chat.title : 'New Chat');
		}
		
		// Setup event listeners
		setupEventListeners();
		
		// Set initial mode
		setMode('ask');
		
		// Log configuration
		console.log('Copilot initialized. Backend:', Config.BACKEND_URL, 'Use dummy:', Config.USE_DUMMY);
	}

	// ============================================
	// ONLYOFFICE PLUGIN INTEGRATION
	// ============================================
	window.Asc = window.Asc || {};
	window.Asc.plugin = window.Asc.plugin || {};

	window.Asc.plugin.init = function() {
		init();
	};

	window.Asc.plugin.onThemeChanged = function(theme) {
		// Theme is handled by CSS variables
	};

	// Fallback init if plugin system doesn't call init
	if (document.readyState === 'loading') {
		document.addEventListener('DOMContentLoaded', function() {
			setTimeout(function() {
				if (!elements.sidebar) {
					init();
				}
			}, 100);
		});
	} else {
		setTimeout(function() {
			if (!elements.sidebar) {
				init();
			}
		}, 100);
	}

})(window);
