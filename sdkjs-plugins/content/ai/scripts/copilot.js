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
		// Document indexing API is now part of the main backend
		BACKEND_URL: 'http://98.83.138.45:8000',
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
		abortController: null, // For cancelling fetch requests
		// Document indexing state
		editorDocId: null, // Current document ID for indexing
		indexedDocs: [], // List of indexed documents
		mentionQuery: '', // Current @ mention search query
		mentionStartPos: -1, // Position where @ was typed
		selectedMentionIndex: 0 // Currently selected item in mention dropdown
	};

	// ============================================
	// DOM ELEMENTS
	// ============================================
	var elements = {};

	// ============================================
	// STORAGE - Document-specific chat storage
	// ============================================
	var Storage = {
		BASE_KEY: 'copilot_chats',
		
		// Get the storage key for the current document
		getKey: function() {
			// Use document-specific storage if available
			var docId = sessionStorage.getItem('copilot_doc_specific_id');
			if (docId) {
				return this.BASE_KEY + '_' + docId;
			}
			// Fallback to generic key (for backwards compatibility or when not in OnlyOffice)
			return this.BASE_KEY;
		},
		
		load: function() {
			try {
				var data = localStorage.getItem(this.getKey());
				return data ? JSON.parse(data) : [];
			} catch (e) {
				return [];
			}
		},
		
		save: function(chats) {
			try {
				localStorage.setItem(this.getKey(), JSON.stringify(chats));
			} catch (e) {
				console.warn('Could not save chats');
			}
		}
	};

	// ============================================
	// TOOL EXECUTOR (Frontend Tools)
	// ============================================
	var ToolExecutor = {
		mmToEmu: function(mm) {
			// 1 inch = 25.4 mm; 1 inch = 914400 EMU
			return Math.round((mm * 914400) / 25.4);
		},
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
							var maxLen = (params && params.max_length) ? params.max_length : 50000;
							var format = (params && params.format) ? params.format : 'markdown';
							if (format === 'plain') {
								var text = doc.GetText ? doc.GetText({ ParaSeparator: '\n', NewLineSeparator: '\n' }) : '';
								return (text || '').substring(0, maxLen);
							}
							// Prefer markdown: preserves headings/tables better for agents.
							var md = doc.ToMarkdown ? doc.ToMarkdown(
								!!(params && params.html_headings),
								!!(params && params.base64img),
								!!(params && params.demote_headings),
								!!(params && params.render_html_tags)
							) : '';
							return (md || '').substring(0, maxLen);
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
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var query = params && params.query ? String(params.query) : '';
							var maxResults = (params && params.max_results) ? params.max_results : 10;
							var matchCase = !!(params && params.match_case);
							if (!query) {
								return { total: 0, results: [] };
							}
							var doc = Api.GetDocument();
							var ranges = doc.Search ? doc.Search(query, matchCase) : [];
							var results = [];
							var count = Math.min(maxResults, ranges.length);
							for (var i = 0; i < count; i++) {
								var r = ranges[i];
								var fullText = r && r.GetText ? r.GetText({ ParaSeparator: '\n', NewLineSeparator: '\n' }) : '';
								results.push({
									index: i,
									text: (fullText || '').substring(0, 400),
									length: (fullText || '').length
								});
							}
							return { total: ranges.length, results: results };
						}, false, false, resolve);
					});
				}
				return { total: 0, results: [], error: 'Not in OnlyOffice environment' };
			},

			'select_search_result': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var query = params && params.query ? String(params.query) : '';
							var idx = (params && typeof params.result_index === 'number') ? params.result_index : 0;
							var matchCase = !!(params && params.match_case);
							if (!query) {
								return { success: false, error: 'query is required' };
							}
							var doc = Api.GetDocument();
							var ranges = doc.Search ? doc.Search(query, matchCase) : [];
							if (!ranges || idx < 0 || idx >= ranges.length) {
								return { success: false, error: 'result_index out of range', total: ranges ? ranges.length : 0 };
							}
							var r = ranges[idx];
							if (r && r.Select) r.Select(true);
							if (r && r.MoveCursorToPos) r.MoveCursorToPos(0);
							var selected = r && r.GetText ? r.GetText({ ParaSeparator: '\n', NewLineSeparator: '\n' }) : '';
							return { success: true, selected_text: selected, total: ranges.length, result_index: idx };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'get_page_info': async function() {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var pageIndex = doc.GetCurrentPage ? doc.GetCurrentPage() : 0; // 0-based
							var pageCount = doc.GetPageCount ? doc.GetPageCount() : 0;
							return {
								page_index: pageIndex,
								page_number: pageIndex + 1,
								page_count: pageCount
							};
						}, false, false, resolve);
					});
				}
				return { page_index: 0, page_number: 1, page_count: 0, error: 'Not in OnlyOffice environment' };
			},

			'get_current_paragraph': async function() {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var para = doc.GetCurrentParagraph ? doc.GetCurrentParagraph() : null;
							var text = para && para.GetText ? para.GetText({ NewLineSeparator: '\n' }) : '';
							var styleName = '';
							try {
								var style = para && para.GetStyle ? para.GetStyle() : null;
								styleName = style && style.GetName ? style.GetName() : '';
							} catch (e) {}
							return { text: text, style: styleName };
						}, false, false, resolve);
					});
				}
				return { text: '', error: 'Not in OnlyOffice environment' };
			},

			'go_to_page': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var pageNumber = (params && params.page_number) ? params.page_number : 1;
							var idx = Math.max(0, Number(pageNumber) - 1);
							var doc = Api.GetDocument();
							var ok = doc.GoToPage ? doc.GoToPage(idx) : false;
							return { success: !!ok, page_number: idx + 1 };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'go_to_bookmark': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var name = params && params.name ? String(params.name) : '';
							if (!name) return { success: false, error: 'name is required' };
							var doc = Api.GetDocument();
							var ok = doc.GoToBookmark ? doc.GoToBookmark(name) : false;
							return { success: !!ok, name: name };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'insert_page_break': async function() {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var para = doc.GetCurrentParagraph ? doc.GetCurrentParagraph() : null;
							if (para && para.AddPageBreak) {
								para.AddPageBreak();
								return { success: true };
							}
							return { success: false, error: 'AddPageBreak not available' };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'insert_table': async function(params) {
				// Generate HTML table and use PasteHtml for reliable insertion at cursor
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					var rows = params && params.rows ? Number(params.rows) : 1;
					var cols = params && params.cols ? Number(params.cols) : 1;
					var data = (params && params.data) ? params.data : null;
					var headerRow = !!(params && params.header_row);
					
					if (rows <= 0 || cols <= 0) {
						return { success: false, error: 'rows and cols must be positive' };
					}
					
					// If data is provided, adjust dimensions
					if (data && Array.isArray(data) && data.length > 0) {
						rows = data.length;
						if (data[0] && Array.isArray(data[0])) {
							cols = Math.max(cols, data[0].length);
						}
					}
					
					// Build HTML table
					var html = '<table style="border-collapse: collapse; width: 100%;">';
					for (var r = 0; r < rows; r++) {
						html += '<tr>';
						for (var c = 0; c < cols; c++) {
							var cellValue = '';
							if (data && data[r] && typeof data[r][c] !== 'undefined') {
								cellValue = String(data[r][c]);
							}
							// Use th for header row
							var tag = (headerRow && r === 0) ? 'th' : 'td';
							var style = 'border: 1px solid #000; padding: 5px;';
							if (headerRow && r === 0) {
								style += ' font-weight: bold; background-color: #f0f0f0;';
							}
							html += '<' + tag + ' style="' + style + '">' + escapeHtmlForTable(cellValue) + '</' + tag + '>';
						}
						html += '</tr>';
					}
					html += '</table>';
					
					return new Promise(function(resolve) {
						window.Asc.plugin.executeMethod('PasteHtml', [html], function(result) {
							resolve({ success: true, rows: rows, cols: cols, message: 'Table inserted' });
						});
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'insert_image': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					var self = this;
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var src = params && params.image_src ? String(params.image_src) : '';
							if (!src) return { success: false, error: 'image_src is required' };
							var wmm = (params && params.width_mm) ? Number(params.width_mm) : 50;
							var hmm = (params && params.height_mm) ? Number(params.height_mm) : 50;
							var wEmu = self.mmToEmu(wmm);
							var hEmu = self.mmToEmu(hmm);
							var img = Api.CreateImage ? Api.CreateImage(src, wEmu, hEmu) : null;
							if (!img) return { success: false, error: 'CreateImage failed' };

							var doc = Api.GetDocument();
							var para = doc.GetCurrentParagraph ? doc.GetCurrentParagraph() : null;
							if (para && para.Push) {
								para.Push(img);
							} else {
								return { success: false, error: 'Could not insert image at cursor' };
							}

							if (params && params.add_caption) {
								var captionText = params.caption_text ? String(params.caption_text) : '';
								var captionLabel = params.caption_label ? String(params.caption_label) : 'Figure';
								try {
									var after = doc.GetCurrentParagraph ? doc.GetCurrentParagraph() : null;
									if (after && after.AddCaption) {
										after.AddCaption(captionText, captionLabel, false, 'Arabic', false, 0, 'hyphen');
									}
								} catch (e) {}
							}
							return { success: true };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'set_header_text': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var text = params && typeof params.text !== 'undefined' ? String(params.text) : '';
							var type = params && params.type ? String(params.type) : 'default';
							var sectionIndex = (params && typeof params.section_index === 'number') ? params.section_index : 0;
							var overwrite = (params && typeof params.overwrite === 'boolean') ? params.overwrite : true;

							var doc = Api.GetDocument();
							var sections = doc.GetSections ? doc.GetSections() : [];
							if (!sections || sectionIndex < 0 || sectionIndex >= sections.length) {
								return { success: false, error: 'section_index out of range', section_count: sections ? sections.length : 0 };
							}
							var section = sections[sectionIndex];
							var header = section.GetHeader ? section.GetHeader(type, true) : null;
							if (!header) return { success: false, error: 'GetHeader failed' };
							if (overwrite && header.RemoveAllElements) header.RemoveAllElements();
							var p = Api.CreateParagraph ? Api.CreateParagraph() : null;
							if (p && p.AddText) p.AddText(text);
							if (header.Push && p) header.Push(p);
							return { success: true, type: type, section_index: sectionIndex };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'set_footer_text': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var text = params && typeof params.text !== 'undefined' ? String(params.text) : '';
							var type = params && params.type ? String(params.type) : 'default';
							var sectionIndex = (params && typeof params.section_index === 'number') ? params.section_index : 0;
							var overwrite = (params && typeof params.overwrite === 'boolean') ? params.overwrite : true;

							var doc = Api.GetDocument();
							var sections = doc.GetSections ? doc.GetSections() : [];
							if (!sections || sectionIndex < 0 || sectionIndex >= sections.length) {
								return { success: false, error: 'section_index out of range', section_count: sections ? sections.length : 0 };
							}
							var section = sections[sectionIndex];
							var footer = section.GetFooter ? section.GetFooter(type, true) : null;
							if (!footer) return { success: false, error: 'GetFooter failed' };
							if (overwrite && footer.RemoveAllElements) footer.RemoveAllElements();
							var p = Api.CreateParagraph ? Api.CreateParagraph() : null;
							if (p && p.AddText) p.AddText(text);
							if (footer.Push && p) footer.Push(p);
							return { success: true, type: type, section_index: sectionIndex };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'add_table_of_contents': async function() {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var ok = doc.AddTableOfContents ? doc.AddTableOfContents({}) : false;
							return { success: !!ok };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'add_table_of_figures': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var buildFrom = params && params.build_from ? String(params.build_from) : 'Figure';
							var replace = !!(params && params.replace);
							var ok = doc.AddTableOfFigures ? doc.AddTableOfFigures({ BuildFrom: buildFrom }, replace) : false;
							return { success: !!ok, build_from: buildFrom };
						}, true, false, resolve);
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'get_document_snapshot': async function(params) {
				var includeMarkdown = !!(params && params.include_markdown);
				var includeOutline = (params && typeof params.include_outline === 'boolean') ? params.include_outline : true;
				var includeHeadersFooters = !!(params && params.include_headers_footers);
				var maxMd = (params && params.max_markdown_chars) ? params.max_markdown_chars : 20000;

				// Selected text is easiest via executeMethod.
				var selectedText = '';
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					selectedText = await new Promise(function(resolve) {
						window.Asc.plugin.executeMethod('GetSelectedText', [], function(result) {
							resolve(result || '');
						});
					});
				}

				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var pageIndex = doc.GetCurrentPage ? doc.GetCurrentPage() : 0;
							var pageCount = doc.GetPageCount ? doc.GetPageCount() : 0;

							var para = doc.GetCurrentParagraph ? doc.GetCurrentParagraph() : null;
							var paraText = para && para.GetText ? para.GetText({ NewLineSeparator: '\n' }) : '';

							var outline = [];
							if (includeOutline) {
								var count = doc.GetElementsCount ? doc.GetElementsCount() : 0;
								for (var i = 0; i < count; i++) {
									var elem = doc.GetElement(i);
									if (elem && elem.GetClassType && elem.GetClassType() === 'paragraph') {
										var style = elem.GetStyle ? elem.GetStyle() : null;
										if (style) {
											var styleName = style.GetName ? style.GetName() : '';
											if (styleName && styleName.indexOf('Heading') === 0) {
												var level = parseInt(styleName.replace('Heading', '').trim(), 10) || 1;
												outline.push({ text: elem.GetText ? elem.GetText({ NewLineSeparator: '\n' }) : '', level: level });
											}
										}
									}
								}
							}

							var snapshot = {
								selected_text: selectedText,
								page_index: pageIndex,
								page_number: pageIndex + 1,
								page_count: pageCount,
								current_paragraph: paraText
							};

							if (includeOutline) snapshot.document_outline = outline;

							if (includeMarkdown && doc.ToMarkdown) {
								var md = doc.ToMarkdown(true, false, false, false);
								snapshot.document_markdown = (md || '').substring(0, maxMd);
							}

							if (includeHeadersFooters && doc.GetSections) {
								var sections = doc.GetSections();
								if (sections && sections.length > 0) {
									var s0 = sections[0];
									var h = s0.GetHeader ? s0.GetHeader('default', false) : null;
									var f = s0.GetFooter ? s0.GetFooter('default', false) : null;
									snapshot.header_default = h && h.GetText ? h.GetText({ ParaSeparator: '\n', NewLineSeparator: '\n' }) : '';
									snapshot.footer_default = f && f.GetText ? f.GetText({ ParaSeparator: '\n', NewLineSeparator: '\n' }) : '';
								}
							}

							return snapshot;
						}, false, false, resolve);
					});
				}

				return {
					selected_text: selectedText,
					page_index: 0,
					page_number: 1,
					page_count: 0,
					current_paragraph: '',
					error: 'Not in OnlyOffice environment'
				};
			},
			
		'insert_text': async function(params) {
			console.log('insert_text called with:', params);
			if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
				var text = params.text;
				var format = params.format || 'plain';
				
				if (format === 'html') {
					return new Promise(function(resolve) {
						window.Asc.plugin.executeMethod('PasteHtml', [text], function(result) {
							console.log('PasteHtml result:', result);
							resolve({ success: true, message: 'HTML inserted' });
						});
					});
				} else if (format === 'markdown') {
					// Convert markdown to HTML
					var html = markdownToHtml(text);
					return new Promise(function(resolve) {
						window.Asc.plugin.executeMethod('PasteHtml', [html], function(result) {
							console.log('PasteHtml (markdown) result:', result);
							resolve({ success: true, message: 'Markdown inserted as HTML' });
						});
					});
				} else {
					// Plain text
					return new Promise(function(resolve) {
						window.Asc.plugin.executeMethod('PasteText', [text], function(result) {
							console.log('PasteText result:', result);
							resolve({ success: true, message: 'Text inserted' });
						});
					});
				}
			}
			console.log('Not in OnlyOffice environment');
			return { success: false, error: 'Not in OnlyOffice environment' };
		},
			
			'replace_selection': async function(params) {
				console.log('replace_selection called with:', params);
				var text = params.text || '';
				
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					return new Promise(function(resolve) {
						// Use InputText which properly handles selection replacement
						// When there's a selection, InputText replaces it with the new text
						// If text is empty, this effectively deletes the selection
						window.Asc.plugin.executeMethod('InputText', [text], function(result) {
							console.log('InputText result:', result);
							if (!text || text === '') {
								resolve({ success: true, message: 'Selection deleted' });
							} else {
								resolve({ success: true, message: 'Selection replaced' });
							}
						});
					});
				}
				
				// Fallback to insert_text for non-deletion cases
				if (text && text !== '') {
					return ToolExecutor.tools['insert_text'](params);
				}
				
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			'delete_selection': async function() {
				console.log('delete_selection called');
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					return new Promise(function(resolve) {
						// InputText with empty string deletes the current selection
						window.Asc.plugin.executeMethod('InputText', [''], function(result) {
							console.log('delete_selection InputText result:', result);
							resolve({ success: true, message: 'Selection deleted' });
						});
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
		'get_content_controls': async function() {
			console.log('get_content_controls called');
			if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
				return new Promise(function(resolve) {
					window.Asc.plugin.callCommand(function() {
						var doc = Api.GetDocument();
						var controls = doc.GetAllContentControls();
						console.log('GetAllContentControls returned ' + controls.length + ' controls');
						var result = [];
						for (var i = 0; i < controls.length; i++) {
							var cc = controls[i];
							var tag = cc.GetTag ? cc.GetTag() : '';
							var title = cc.GetLabel ? cc.GetLabel() : '';
							var text = '';
							
							// Try to get text content
							try {
								if (cc.GetElementsCount) {
									var count = cc.GetElementsCount();
									for (var j = 0; j < count; j++) {
										var elem = cc.GetElement(j);
										if (elem && elem.GetText) {
											text += elem.GetText();
										}
									}
								}
							} catch (e) {
								console.log('Error getting text:', e);
							}
							
							result.push({
								tag: tag,
								title: title,
								value: text
							});
							console.log('Control ' + i + ':', JSON.stringify(result[i]));
						}
						return result;
					}, false, false, resolve);
				});
			}
			console.log('Not in OnlyOffice environment');
			return [];
		},
			
		'fill_content_control': async function(params) {
			console.log('fill_content_control called with:', params);
			if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
				return new Promise(function(resolve) {
					window.Asc.plugin.callCommand(function() {
						var doc = Api.GetDocument();
						var controls = doc.GetAllContentControls();
						console.log('Found ' + controls.length + ' content controls');
						
						for (var i = 0; i < controls.length; i++) {
							var cc = controls[i];
							var tag = cc.GetTag();
							var title = cc.GetLabel ? cc.GetLabel() : '';
							console.log('Control ' + i + ': tag=' + tag + ', title=' + title);
							
							// Match by tag or title
							if (tag === params.tag || title === params.tag) {
								// Clear existing content and add new text
								var range = cc.GetRange(0, 0);
								if (range) {
									range.Delete();
								}
								
								// Get the paragraph inside the content control and add text
								var count = cc.GetElementsCount ? cc.GetElementsCount() : 0;
								if (count > 0) {
									var para = cc.GetElement(0);
									if (para && para.AddText) {
										para.AddText(params.value);
									}
								} else {
									// Fallback: try SetText if available
									if (cc.SetText) {
										cc.SetText(params.value);
									}
								}
								
								return { success: true, message: 'Content control "' + params.tag + '" filled with: ' + params.value };
							}
						}
						return { success: false, error: 'Content control "' + params.tag + '" not found. Available tags: ' + controls.map(function(c) { return c.GetTag(); }).join(', ') };
					}, true, false, resolve);  // Changed to true for recalculation
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

	// Helper to escape HTML for table cells
	function escapeHtmlForTable(text) {
		if (!text) return '';
		return String(text)
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;')
			.replace(/'/g, '&#39;');
	}

	// Simple markdown to HTML converter for insert_text
	function markdownToHtml(md) {
		if (!md) return '';
		
		// First, handle markdown tables before other transformations
		md = convertMarkdownTables(md);
		
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
		// Line breaks (but not inside tables which are already converted)
		html = html.replace(/\n/g, '<br>');
		return html;
	}

	// Convert markdown tables to HTML tables
	function convertMarkdownTables(md) {
		if (!md) return '';
		
		var lines = md.split('\n');
		var result = [];
		var tableLines = [];
		var inTable = false;
		
		for (var i = 0; i < lines.length; i++) {
			var line = lines[i].trim();
			
			// Check if this line looks like a table row (has | separators)
			var isTableRow = line.indexOf('|') !== -1 && 
				(line.match(/\|/g) || []).length >= 2;
			
			// Check if this is a separator line (|---|---|)
			var isSeparator = /^\|?[\s\-:]+\|[\s\-:|]+\|?$/.test(line);
			
			if (isTableRow || isSeparator) {
				if (!inTable) {
					inTable = true;
					tableLines = [];
				}
				if (!isSeparator) {
					tableLines.push(line);
				}
			} else {
				if (inTable && tableLines.length > 0) {
					// Convert accumulated table lines to HTML
					result.push(buildHtmlTable(tableLines));
					tableLines = [];
					inTable = false;
				}
				result.push(lines[i]);
			}
		}
		
		// Handle table at end of content
		if (inTable && tableLines.length > 0) {
			result.push(buildHtmlTable(tableLines));
		}
		
		return result.join('\n');
	}

	// Build HTML table from markdown table rows
	function buildHtmlTable(tableLines) {
		if (tableLines.length === 0) return '';
		
		var html = '<table style="border-collapse: collapse; width: 100%;">';
		
		for (var i = 0; i < tableLines.length; i++) {
			var line = tableLines[i];
			var cells = parseTableRow(line);
			var isHeader = (i === 0);
			
			html += '<tr>';
			for (var j = 0; j < cells.length; j++) {
				var tag = isHeader ? 'th' : 'td';
				var style = 'border: 1px solid #000; padding: 5px;';
				if (isHeader) {
					style += ' font-weight: bold; background-color: #f0f0f0;';
				}
				html += '<' + tag + ' style="' + style + '">' + escapeHtmlForTable(cells[j]) + '</' + tag + '>';
			}
			html += '</tr>';
		}
		
		html += '</table>';
		return html;
	}

	// Parse a markdown table row into cells
	function parseTableRow(line) {
		// Remove leading/trailing pipes and split by |
		var trimmed = line.trim();
		if (trimmed.charAt(0) === '|') trimmed = trimmed.substring(1);
		if (trimmed.charAt(trimmed.length - 1) === '|') trimmed = trimmed.substring(0, trimmed.length - 1);
		
		var cells = trimmed.split('|').map(function(cell) {
			return cell.trim();
		});
		
		return cells;
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
			'delete_selection': '🗑️',
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

	// Simple markdown parser with proper link support
	function parseMarkdown(text) {
		if (!text) return '';
		
		// Process markdown links BEFORE escaping HTML to preserve them
		// Match [text](url) pattern and extract parts
		var linkPlaceholders = [];
		var processedText = text.replace(/\[([^\]]+)\]\s*\(([^)]+)\)/g, function(match, linkText, url) {
			var placeholder = '___MDLINK_' + linkPlaceholders.length + '___';
			linkPlaceholders.push({ text: linkText.trim(), url: url.trim() });
			return placeholder;
		});
		
		// Escape HTML on the text with placeholders
		var html = escapeHtml(processedText);
		
		// Restore links as proper HTML anchor tags with beautiful styling
		for (var i = 0; i < linkPlaceholders.length; i++) {
			var link = linkPlaceholders[i];
			var linkHtml = '<a href="' + escapeHtml(link.url) + '" target="_blank" rel="noopener noreferrer" class="md-link">' + 
				escapeHtml(link.text) + 
				'<svg class="md-link-icon" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg></a>';
			html = html.replace('___MDLINK_' + i + '___', linkHtml);
		}
		
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
			// Minimal thinking indicator - collapsed by default, no emoji
			html += '<div class="thinking-block collapsed">';
			html += '<div class="thinking-header">';
			html += '<span class="thinking-icon"></span>';
			html += '<span class="thinking-label">' + escapeHtml(item.content || 'Processing...') + '</span>';
			html += '<span class="thinking-toggle">›</span>';
			html += '</div>';
			html += '<div class="thinking-content">' + escapeHtml(item.content) + '</div>';
			html += '</div>';
		} else if (item.type === 'tool_call' || item.type === 'tool_result') {
			var status = item.status || 'running';
			var toolIdAttr = item.id ? (' data-tool-call-id="' + escapeHtml(item.id) + '"') : '';
			
			// Check if this is a regulatory_search tool with results - render rich UI
			if (item.name === 'regulatory_search' && item.result && item.result.data && item.result.data.success) {
				html += renderRegulatorySearchResult(item);
			} else {
				// Minimal inline tool indicator for other tools
				html += '<div class="tool-call tool-call--minimal" data-status="' + status + '"' + toolIdAttr + '>';
				html += '<div class="tool-call-header" role="button" tabindex="0" aria-expanded="false">';
				html += '<div class="tool-call-left">';
				html += '<span class="tool-dot"></span>';
				html += '<span class="tool-name">' + escapeHtml(formatToolName(item.name)) + '</span>';
				html += '<span class="tool-status ' + status + '">';
				if (status === 'running') {
					html += '<span class="spinner" aria-hidden="true"></span>';
				}
				html += '</span>';
				html += '</div>';
				html += '<span class="tool-toggle" aria-hidden="true">›</span>';
				html += '</div>';
				// Details hidden by default, shown on expand
				html += '<div class="tool-call-body">';
				if (item.params && Object.keys(item.params).length > 0) {
					html += '<div class="tool-section-label">Params</div>';
					html += '<pre class="tool-json">' + escapeHtml(formatJsonTruncated(item.params, 1500)) + '</pre>';
				}
				if (item.result !== undefined) {
					html += '<div class="tool-result-section">';
					html += '<div class="tool-section-label">Result</div>';
					html += '<pre class="tool-json">' + escapeHtml(typeof item.result === 'string' ? truncateResult(item.result, 500) : formatJsonTruncated(item.result, 2000)) + '</pre>';
					html += '</div>';
				}
				html += '</div>';
				html += '</div>';
			}
		} else if (item.type === 'checkpoint') {
			// Simplified checkpoint - just a subtle divider
			html += '<div class="checkpoint">';
			html += '<div class="checkpoint-line"></div>';
			html += '</div>';
		}
		
		return html;
	}

	// Render regulatory search results with rich UI (Ritivel-style)
	function renderRegulatorySearchResult(item) {
		var data = item.result.data;
		var query = data.query || item.params.query || '';
		var answer = data.answer || '';
		var sources = data.sources || [];
		
		var html = '<div class="regulatory-search-result">';
		
		// Two-column layout container
		html += '<div class="reg-search-layout">';
		
		// Left column: Progress panel (simplified for results view)
		html += '<div class="reg-search-progress-panel">';
		html += '<div class="reg-progress-header">';
		html += '<div class="reg-progress-icon-wrapper">';
		html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>';
		html += '</div>';
		html += '<span class="reg-progress-title">Search Complete</span>';
		html += '</div>';
		
		// Progress steps (all completed)
		var steps = [
			{ icon: 'brain', label: 'Query Analyzed', status: 'complete' },
			{ icon: 'list', label: 'Query Decomposed', status: 'complete' },
			{ icon: 'search', label: 'Search & Rerank', status: 'complete' },
			{ icon: 'pen', label: 'Answer Synthesized', status: 'complete' }
		];
		
		html += '<div class="reg-progress-steps">';
		for (var s = 0; s < steps.length; s++) {
			var step = steps[s];
			var isLast = s === steps.length - 1;
			html += '<div class="reg-step">';
			html += '<div class="reg-step-indicator">';
			html += '<div class="reg-step-icon ' + step.status + '">';
			html += '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>';
			html += '</div>';
			if (!isLast) {
				html += '<div class="reg-step-line complete"></div>';
			}
			html += '</div>';
			html += '<div class="reg-step-content">';
			html += '<span class="reg-step-label complete">' + step.label + '</span>';
			html += '</div>';
			html += '</div>';
		}
		html += '</div>';
		
		// Sources summary
		if (sources.length > 0) {
			html += '<div class="reg-sources-summary">';
			html += '<div class="reg-sources-summary-label">Sources Found</div>';
			html += '<div class="reg-sources-summary-chips">';
			var typeCount = {};
			for (var t = 0; t < sources.length; t++) {
				var stype = (sources[t].source_type || 'doc').toUpperCase();
				typeCount[stype] = (typeCount[stype] || 0) + 1;
			}
			for (var type in typeCount) {
				html += '<span class="reg-type-chip reg-type-' + type.toLowerCase() + '">' + type + ' (' + typeCount[type] + ')</span>';
			}
			html += '</div>';
			html += '</div>';
		}
		
		html += '</div>'; // End progress panel
		
		// Right column: Results
		html += '<div class="reg-search-results-panel">';
		
		// Answer section (shown first, above sources)
		if (answer) {
			html += '<div class="reg-answer-section">';
			html += '<div class="reg-answer-header">';
			html += '<div class="reg-answer-icon">';
			html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 19l7-7 3 3-7 7-3-3z"/><path d="M18 13l-1.5-7.5L2 2l3.5 14.5L13 18l5-5z"/><path d="M2 2l7.586 7.586"/></svg>';
			html += '</div>';
			html += '<span class="reg-answer-title">Answer</span>';
			html += '</div>';
			html += '<div class="reg-answer-content">' + formatAnswerWithCitations(answer) + '</div>';
			html += '</div>';
		}
		
		// Sources section
		if (sources.length > 0) {
			html += '<div class="reg-sources-section">';
			html += '<div class="reg-sources-header">';
			html += '<div class="reg-sources-header-left">';
			html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>';
			html += '<span>Sources</span>';
			html += '<span class="reg-sources-count">' + sources.length + '</span>';
			html += '</div>';
			html += '<div class="reg-sources-header-right">';
			html += '<span class="reg-sources-dbs">ICH • FDA • PSG</span>';
			html += '</div>';
			html += '</div>';
			
			html += '<div class="reg-sources-list">';
			for (var i = 0; i < sources.length; i++) {
				var src = sources[i];
				html += renderSourceCard(src, i);
			}
			html += '</div>';
			html += '</div>';
		}
		
		html += '</div>'; // End results panel
		html += '</div>'; // End layout
		html += '</div>'; // End regulatory-search-result
		
		return html;
	}
	
	// Render individual source card (Ritivel-style expandable)
	function renderSourceCard(src, index) {
		var hasUrl = src.url && src.url.length > 0;
		var sourceType = (src.source_type || 'doc').toLowerCase();
		var uniqueId = 'source-' + index + '-' + Date.now();
		
		var html = '<div class="reg-source-card" data-source-id="' + uniqueId + '" data-expanded="false">';
		
		// Card header (always visible, clickable to expand)
		html += '<button class="reg-source-card-header" onclick="toggleSourceCard(\'' + uniqueId + '\')">';
		
		// Citation badge
		html += '<div class="reg-citation-badge">' + (index + 1) + '</div>';
		
		// Title and meta
		html += '<div class="reg-source-info">';
		html += '<div class="reg-source-title-row">';
		html += '<h4 class="reg-source-title">' + escapeHtml(src.title || 'Source ' + (index + 1)) + '</h4>';
		html += '<svg class="reg-source-chevron" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>';
		html += '</div>';
		
		// Meta row: code, type badge, pages
		html += '<div class="reg-source-meta-row">';
		if (src.code) {
			html += '<span class="reg-source-code">' + escapeHtml(src.code) + '</span>';
		}
		html += '<span class="reg-source-type-badge reg-type-' + sourceType + '">' + sourceType.toUpperCase() + '</span>';
		if (src.page_numbers) {
			html += '<span class="reg-source-pages">' + escapeHtml(src.page_numbers) + '</span>';
		}
		html += '</div>';
		
		// Snippet (truncated)
		if (src.snippet) {
			html += '<p class="reg-source-snippet-preview">' + escapeHtml(truncateText(src.snippet, 120)) + '</p>';
		}
		
		html += '</div>'; // End source info
		html += '</button>'; // End header
		
		// Expandable content
		html += '<div class="reg-source-expanded">';
		
		// Header path breadcrumb
		if (src.header_path) {
			html += '<div class="reg-source-breadcrumb">';
			html += '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>';
			html += '<span>' + escapeHtml(src.header_path) + '</span>';
			html += '</div>';
		}
		
		// Full text
		html += '<div class="reg-source-full-text">';
		html += '<p>' + escapeHtml(src.snippet || src.full_text || 'No content available.') + '</p>';
		html += '</div>';
		
		// Footer with relevance and link
		html += '<div class="reg-source-footer">';
		if (src.relevance_score && src.relevance_score > 0) {
			var relevancePercent = Math.min(100, Math.round(src.relevance_score * 100));
			html += '<div class="reg-source-relevance">';
			html += '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>';
			html += '<span>Relevance: ' + relevancePercent + '%</span>';
			html += '</div>';
		}
		if (hasUrl) {
			html += '<a href="' + escapeHtml(src.url) + '" target="_blank" rel="noopener noreferrer" class="reg-source-link-btn">';
			html += '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>';
			html += '<span>View Source</span>';
			html += '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>';
			html += '</a>';
		}
		html += '</div>';
		
		html += '</div>'; // End expanded
		html += '</div>'; // End card
		
		return html;
	}
	
	// Format answer text with clickable citation badges and proper links
	function formatAnswerWithCitations(text) {
		if (!text) return '';
		
		// Process markdown links BEFORE escaping HTML to preserve them
		// Match [text](url) pattern and extract parts
		var linkPlaceholders = [];
		var processedText = text.replace(/\[([^\]]+)\]\s*\(([^)]+)\)/g, function(match, linkText, url) {
			var placeholder = '___LINK_' + linkPlaceholders.length + '___';
			linkPlaceholders.push({ text: linkText.trim(), url: url.trim() });
			return placeholder;
		});
		
		// Now escape HTML on the text with placeholders
		var html = escapeHtml(processedText);
		
		// Restore links as proper HTML anchor tags with beautiful styling
		for (var i = 0; i < linkPlaceholders.length; i++) {
			var link = linkPlaceholders[i];
			var linkHtml = '<a href="' + escapeHtml(link.url) + '" target="_blank" rel="noopener noreferrer" class="reg-answer-link">' + 
				escapeHtml(link.text) + 
				'<svg class="reg-link-icon" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg></a>';
			html = html.replace('___LINK_' + i + '___', linkHtml);
		}
		
		// Bold text: **text** -> <strong>text</strong>
		html = html.replace(/\*\*([^*]+)\*\*/g, '<strong class="reg-answer-bold">$1</strong>');
		
		// Citations: [1], [2] -> clickable badges
		html = html.replace(/\[(\d+)\]/g, '<span class="reg-citation-inline" onclick="scrollToSource($1)" title="View source $1">$1</span>');
		
		// Handle bullet points/lists
		html = html.replace(/^\s*[-•]\s+(.+)$/gm, '<li>$1</li>');
		html = html.replace(/(<li>.*<\/li>\n?)+/g, '<ul class="reg-answer-list">$&</ul>');
		
		// Line breaks
		html = html.replace(/\n\n/g, '</p><p>');
		html = html.replace(/\n/g, '<br>');
		
		// Wrap in paragraph
		if (!html.startsWith('<p>')) {
			html = '<p>' + html + '</p>';
		}
		
		return html;
	}
	
	// Truncate text helper
	function truncateText(text, maxLen) {
		if (!text || text.length <= maxLen) return text;
		return text.substring(0, maxLen).trim() + '...';
	}
	
	// Global function to toggle source card expansion
	window.toggleSourceCard = function(sourceId) {
		var card = document.querySelector('[data-source-id="' + sourceId + '"]');
		if (card) {
			var isExpanded = card.getAttribute('data-expanded') === 'true';
			card.setAttribute('data-expanded', !isExpanded);
		}
	};
	
	// Global function to scroll to source
	window.scrollToSource = function(sourceNum) {
		var cards = document.querySelectorAll('.reg-source-card');
		var targetCard = cards[sourceNum - 1];
		if (targetCard) {
			targetCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
			targetCard.setAttribute('data-expanded', 'true');
			// Brief highlight
			targetCard.classList.add('reg-source-highlight');
			setTimeout(function() {
				targetCard.classList.remove('reg-source-highlight');
			}, 2000);
		}
	};

	// Format tool name for display (more readable)
	function formatToolName(name) {
		if (!name) return 'unknown';
		// Convert snake_case to readable: get_document_text -> Get document text
		return name.replace(/_/g, ' ');
	}

	// Truncate long result strings for cleaner display
	function truncateResult(str, maxLen) {
		if (!str || str.length <= maxLen) return str;
		return str.substring(0, maxLen) + '...';
	}

	function formatJsonTruncated(obj, maxLen) {
		var s = formatJson(obj);
		if (!maxLen || s.length <= maxLen) return s;
		return s.substring(0, maxLen) + '\n… (truncated)';
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
				conversation_history: getConversationHistory(),
				editor_doc_id: DocIndex.getEditorDocId()
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
					id: data.id,
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
				updateToolCallResult(msgContainer, data.id, data.name, toolResult);
				
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
				id: data.id,
				name: data.name, 
				params: data.params, 
				status: data.status || 'success',
				result: data.result
			};
			
			// Check if we already have this tool call (by ID first, then by running status)
			// Note: We check regardless of status because the frontend may have already
			// updated the tool call to 'success' before the backend sends the final event
			var existingToolCall = data.id ? msgContainer.querySelector('.tool-call[data-tool-call-id="' + data.id + '"]') : null;
			if (!existingToolCall) {
				// Fallback: look for a running tool call (for backwards compatibility)
				existingToolCall = msgContainer.querySelector('.tool-call[data-status="running"]');
			}
			if (existingToolCall) {
				// Update existing only if still running
				if (existingToolCall.getAttribute('data-status') === 'running') {
					updateToolCallResult(msgContainer, data.id, data.name, { 
						success: data.status === 'success', 
						result: data.result 
					});
				}
				// If already completed, skip (avoid duplicate updates)
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

	function updateToolCallResult(msgContainer, toolCallId, toolName, result) {
		var byId = null;
		if (toolCallId) {
			byId = msgContainer.querySelector('.tool-call[data-tool-call-id="' + toolCallId + '"]');
		}
		if (byId && byId.getAttribute('data-status') === 'running') {
			applyToolCallResultToElement(byId, result);
			return;
		}

		var toolCalls = msgContainer.querySelectorAll('.tool-call');
		for (var i = toolCalls.length - 1; i >= 0; i--) {
			var toolCall = toolCalls[i];
			var nameEl = toolCall.querySelector('.tool-name');
			if (nameEl && nameEl.textContent === toolName && toolCall.getAttribute('data-status') === 'running') {
				applyToolCallResultToElement(toolCall, result);
				break;
			}
		}
	}

	function applyToolCallResultToElement(toolCallEl, result) {
				var status = result.success ? 'success' : 'error';
		toolCallEl.setAttribute('data-status', status);
				
		var statusEl = toolCallEl.querySelector('.tool-status');
				statusEl.className = 'tool-status ' + status;

		var statusTextEl = toolCallEl.querySelector('.tool-status-text');
		if (statusTextEl) statusTextEl.textContent = status === 'success' ? 'Done' : 'Error';

		// Remove spinner if present
		var spinner = statusEl.querySelector('.spinner');
		if (spinner) spinner.remove();

		var bodyEl = toolCallEl.querySelector('.tool-call-body');
		// Avoid duplicating result section if already added (e.g., both tool_result_request and tool_call events arrive)
		if (bodyEl && !bodyEl.querySelector('.tool-result-section')) {
			var resultStr = typeof result.result === 'string' ? result.result : formatJsonTruncated(result.result, 8000);
				bodyEl.insertAdjacentHTML('beforeend', 
					'<div class="tool-result-section"><div class="tool-section-label">Result</div>' +
					'<pre class="tool-json">' + escapeHtml(resultStr) + '</pre></div>'
				);
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
				var toolItem = { type: 'tool_call', id: item.id, name: item.name, params: item.params, status: 'running' };
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
					var statusTextEl = lastToolCall.querySelector('.tool-status-text');
					if (statusTextEl) statusTextEl.textContent = item.status === 'success' ? 'Done' : 'Error';
					var spinner = statusEl.querySelector('.spinner');
					if (spinner) spinner.remove();
					
					var bodyEl = lastToolCall.querySelector('.tool-call-body');
					bodyEl.insertAdjacentHTML('beforeend', 
						'<div class="tool-result-section"><div class="tool-section-label">Result</div>' +
						'<pre class="tool-json">' + escapeHtml(typeof item.result === 'string' ? item.result : formatJsonTruncated(item.result, 8000)) + '</pre></div>'
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
			// Check for @ mention keyboard navigation first
			if (MentionAutocomplete.handleKeyDown(e, elements.chatInput)) {
				return;
			}
			if (e.key === 'Enter' && !e.shiftKey) {
				e.preventDefault();
				sendMessage();
			}
		});
		elements.chatInput.addEventListener('input', function() {
			autoResizeTextarea();
			// Check for @ mention
			MentionAutocomplete.checkForMention(elements.chatInput);
		});
		
		// Hide mention dropdown when clicking outside
		document.addEventListener('click', function(e) {
			if (!e.target.closest('#mentionDropdown') && !e.target.closest('#chatInput')) {
				MentionAutocomplete.hide();
			}
		});
		
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

				// Keep aria state in sync for keyboard/screen readers
				var expanded = card.classList.contains('expanded');
				toolHeader.setAttribute('aria-expanded', expanded ? 'true' : 'false');
			}
		});
	}

	// ============================================
	// DOCUMENT INDEXING MODULE
	// ============================================
	var DocIndex = {
		// Get or generate editor document ID - tied to the specific Word document
		getEditorDocId: function() {
			if (state.editorDocId) return state.editorDocId;
			
			// Use the document-specific ID if we have one (set during initialization)
			// This ties each document to its own chat and reference document session
			var docSpecificId = sessionStorage.getItem('copilot_doc_specific_id');
			if (docSpecificId) {
				state.editorDocId = docSpecificId;
				return docSpecificId;
			}
			
			// Fallback to localStorage-based ID (for backwards compatibility)
			var stored = localStorage.getItem('copilot_editor_doc_id');
			if (stored) {
				state.editorDocId = stored;
				return stored;
			}
			
			// Generate a new ID (shouldn't reach here normally)
			state.editorDocId = 'doc_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
			localStorage.setItem('copilot_editor_doc_id', state.editorDocId);
			return state.editorDocId;
		},
		
		// Initialize document-specific ID based on document content
		// Call this once when copilot loads to tie session to the specific document
		// IMPORTANT: Detects document changes and clears old session when switching documents
		initDocumentId: async function() {
			try {
				// Get a fingerprint of the document to create a unique ID
				var docFingerprint = await this.getDocumentFingerprint();
				if (docFingerprint) {
					var newDocId = 'doc_' + docFingerprint;
					var previousDocId = sessionStorage.getItem('copilot_doc_specific_id');
					
					// Check if document has changed (different fingerprint)
					if (previousDocId && previousDocId !== newDocId) {
						console.log('Copilot: Document changed! Old:', previousDocId, 'New:', newDocId);
						// Clear old session data - new document means fresh start
						this.clearSessionData();
					}
					
					// Set the new document ID
					sessionStorage.setItem('copilot_doc_specific_id', newDocId);
					state.editorDocId = newDocId;
					console.log('Copilot: Document-specific ID initialized:', newDocId);
					return newDocId;
				}
			} catch (e) {
				console.warn('Copilot: Could not get document fingerprint, using fallback:', e);
			}
			
			// Fallback to getEditorDocId
			return this.getEditorDocId();
		},
		
		// Clear session data when switching to a new document
		clearSessionData: function() {
			console.log('Copilot: Clearing session data for new document');
			// Clear the indexed documents list
			state.indexedDocs = [];
			// Clear chat state
			state.chats = [];
			state.currentChatId = null;
			state.editorDocId = null;
			
			// Update UI to reflect empty state
			if (elements.indexBadge) {
				elements.indexBadge.textContent = '0';
				elements.indexBadge.style.display = 'none';
			}
			
			// Note: We don't clear localStorage here because different documents
			// have their own chat history stored with document-specific keys
		},
		
		// Get a fingerprint of the document for identification
		getDocumentFingerprint: async function() {
			if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
				return new Promise(function(resolve) {
					window.Asc.plugin.callCommand(function() {
						var doc = Api.GetDocument();
						// Get document stats for fingerprinting
						var allText = '';
						var elemCount = doc.GetElementsCount ? doc.GetElementsCount() : 0;
						
						// Get text from first several paragraphs for fingerprinting
						for (var i = 0; i < Math.min(elemCount, 10); i++) {
							var elem = doc.GetElement(i);
							if (elem && elem.GetText) {
								var text = elem.GetText();
								if (text && text.trim()) {
									allText += text.substring(0, 200);
								}
							}
						}
						
						// For empty documents, use a special marker
						if (!allText.trim() || elemCount === 0) {
							// Empty document - generate unique ID based on current time
							// This ensures each new empty document gets its own session
							return 'empty_' + Date.now().toString(36);
						}
						
						// Create fingerprint from content + element count
						var fingerprint = allText.substring(0, 500) + '_elems_' + elemCount;
						
						// Simple hash function (djb2)
						var hash = 5381;
						for (var j = 0; j < fingerprint.length; j++) {
							var char = fingerprint.charCodeAt(j);
							hash = ((hash << 5) + hash) + char;
							hash = hash & hash; // Convert to 32bit integer
						}
						return Math.abs(hash).toString(36);
					}, false, false, resolve);
				});
			}
			// Not in OnlyOffice environment - generate unique ID
			return 'fallback_' + Date.now().toString(36);
		},
		
		// Fetch list of indexed documents
		fetchDocuments: async function() {
			try {
				var docId = this.getEditorDocId();
				var response = await fetch(Config.BACKEND_URL + '/api/docindex/' + docId + '/documents');
				if (!response.ok) throw new Error('Failed to fetch documents');
				var data = await response.json();
				state.indexedDocs = data.documents || [];
				this.updateBadge();
				return state.indexedDocs;
			} catch (e) {
				console.error('Failed to fetch indexed documents:', e);
				return [];
			}
		},
		
		// Upload and index a document
		uploadDocument: async function(file, onProgress) {
			var docId = this.getEditorDocId();
			var formData = new FormData();
			formData.append('file', file);
			
			try {
				onProgress && onProgress('Uploading...', 10);
				
				var response = await fetch(Config.BACKEND_URL + '/api/docindex/' + docId + '/upload', {
					method: 'POST',
					body: formData
				});
				
				onProgress && onProgress('Processing...', 50);
				
				if (!response.ok) throw new Error('Upload failed');
				var result = await response.json();
				
				if (result.success) {
					onProgress && onProgress('Indexed!', 100);
					await this.fetchDocuments();
					return result;
				} else {
					throw new Error(result.error || 'Indexing failed');
				}
			} catch (e) {
				console.error('Upload error:', e);
				throw e;
			}
		},
		
		// Delete an indexed document
		deleteDocument: async function(docId) {
			var editorDocId = this.getEditorDocId();
			try {
				var response = await fetch(
					Config.BACKEND_URL + '/api/docindex/' + editorDocId + '/documents/' + docId,
					{ method: 'DELETE' }
				);
				if (!response.ok) throw new Error('Delete failed');
				await this.fetchDocuments();
				return true;
			} catch (e) {
				console.error('Delete error:', e);
				return false;
			}
		},
		
		// Search indexed documents
		search: async function(query, docNames) {
			var editorDocId = this.getEditorDocId();
			try {
				var response = await fetch(
					Config.BACKEND_URL + '/api/docindex/' + editorDocId + '/search',
					{
						method: 'POST',
						headers: { 'Content-Type': 'application/json' },
						body: JSON.stringify({
							query: query,
							document_names: docNames || null,
							max_results: 5
						})
					}
				);
				if (!response.ok) throw new Error('Search failed');
				return await response.json();
			} catch (e) {
				console.error('Search error:', e);
				return { success: false, results: [] };
			}
		},
		
		// Update the badge count
		updateBadge: function() {
			var badge = document.getElementById('indexBadge');
			if (badge) {
				var count = state.indexedDocs.length;
				badge.textContent = count;
				badge.classList.toggle('hidden', count === 0);
			}
		},
		
		// Render the documents list in the modal
		renderDocsList: function() {
			var listEl = document.getElementById('docsList');
			var emptyEl = document.getElementById('docsEmpty');
			var countEl = document.getElementById('docsCount');
			
			if (!listEl) return;
			
			// Update count
			if (countEl) {
				countEl.textContent = state.indexedDocs.length + ' document' + (state.indexedDocs.length !== 1 ? 's' : '');
			}
			
			// Clear list (keep empty message)
			Array.from(listEl.children).forEach(function(child) {
				if (child.id !== 'docsEmpty') {
					child.remove();
				}
			});
			
			// Show/hide empty message
			if (emptyEl) {
				emptyEl.classList.toggle('hidden', state.indexedDocs.length > 0);
			}
			
			// Render documents
			state.indexedDocs.forEach(function(doc) {
				var item = document.createElement('div');
				item.className = 'doc-item';
				item.innerHTML = 
					'<div class="doc-icon">' +
						'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
							'<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
							'<polyline points="14 2 14 8 20 8"/>' +
						'</svg>' +
					'</div>' +
					'<div class="doc-info">' +
						'<div class="doc-name">' + escapeHtml(doc.filename) + '</div>' +
						'<div class="doc-meta">' + doc.chunk_count + ' chunks</div>' +
					'</div>' +
					'<button class="doc-delete" data-doc-id="' + doc.doc_id + '" title="Remove">' +
						'<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
							'<line x1="18" y1="6" x2="6" y2="18"/>' +
							'<line x1="6" y1="6" x2="18" y2="18"/>' +
						'</svg>' +
					'</button>';
				listEl.appendChild(item);
			});
		},
		
		// Show the modal
		showModal: function() {
			var overlay = document.getElementById('indexModalOverlay');
			if (overlay) {
				overlay.classList.remove('hidden');
				this.fetchDocuments().then(function() {
					DocIndex.renderDocsList();
				});
			}
		},
		
		// Hide the modal
		hideModal: function() {
			var overlay = document.getElementById('indexModalOverlay');
			if (overlay) {
				overlay.classList.add('hidden');
			}
		},
		
		// Setup modal event listeners
		setupModalEvents: function() {
			var self = this;
			
			// Open modal button
			var openBtn = document.getElementById('indexDocsBtn');
			if (openBtn) {
				openBtn.addEventListener('click', function() {
					self.showModal();
				});
			}
			
			// Close modal button
			var closeBtn = document.getElementById('indexModalClose');
			if (closeBtn) {
				closeBtn.addEventListener('click', function() {
					self.hideModal();
				});
			}
			
			// Close on overlay click
			var overlay = document.getElementById('indexModalOverlay');
			if (overlay) {
				overlay.addEventListener('click', function(e) {
					if (e.target === overlay) {
						self.hideModal();
					}
				});
			}
			
			// Upload zone
			var uploadZone = document.getElementById('uploadZone');
			var fileInput = document.getElementById('fileInput');
			
			if (uploadZone && fileInput) {
				uploadZone.addEventListener('click', function() {
					fileInput.click();
				});
				
				uploadZone.addEventListener('dragover', function(e) {
					e.preventDefault();
					uploadZone.classList.add('drag-over');
				});
				
				uploadZone.addEventListener('dragleave', function() {
					uploadZone.classList.remove('drag-over');
				});
				
				uploadZone.addEventListener('drop', function(e) {
					e.preventDefault();
					uploadZone.classList.remove('drag-over');
					var files = e.dataTransfer.files;
					if (files.length > 0) {
						self.handleFileUpload(files);
					}
				});
				
				fileInput.addEventListener('change', function() {
					if (fileInput.files.length > 0) {
						self.handleFileUpload(fileInput.files);
					}
				});
			}
			
			// Delete button clicks (event delegation)
			var docsList = document.getElementById('docsList');
			if (docsList) {
				docsList.addEventListener('click', function(e) {
					var deleteBtn = e.target.closest('.doc-delete');
					if (deleteBtn) {
						var docId = deleteBtn.getAttribute('data-doc-id');
						if (docId && confirm('Remove this document from the index?')) {
							self.deleteDocument(docId).then(function() {
								self.renderDocsList();
							});
						}
					}
				});
			}
		},
		
		// Handle file upload
		handleFileUpload: async function(files) {
			var self = this;
			var progressEl = document.getElementById('uploadProgress');
			var progressFill = document.getElementById('progressFill');
			var progressText = document.getElementById('progressText');
			var uploadZone = document.getElementById('uploadZone');
			
			if (progressEl) progressEl.classList.remove('hidden');
			if (uploadZone) uploadZone.style.display = 'none';
			
			for (var i = 0; i < files.length; i++) {
				var file = files[i];
				try {
					await self.uploadDocument(file, function(text, pct) {
						if (progressText) progressText.textContent = file.name + ': ' + text;
						if (progressFill) progressFill.style.width = pct + '%';
					});
				} catch (e) {
					if (progressText) progressText.textContent = 'Error: ' + e.message;
				}
			}
			
			// Reset UI after a delay
			setTimeout(function() {
				if (progressEl) progressEl.classList.add('hidden');
				if (progressFill) progressFill.style.width = '0%';
				if (uploadZone) uploadZone.style.display = '';
				self.renderDocsList();
			}, 1000);
		}
	};
	
	// ============================================
	// @ MENTION AUTOCOMPLETE
	// ============================================
	var MentionAutocomplete = {
		// Check if we're in a mention context
		checkForMention: function(inputEl) {
			var value = inputEl.value;
			var cursorPos = inputEl.selectionStart;
			
			// Look for @ before cursor
			var textBeforeCursor = value.substring(0, cursorPos);
			var lastAtIndex = textBeforeCursor.lastIndexOf('@');
			
			if (lastAtIndex === -1) {
				this.hide();
				return;
			}
			
			// Check if there's a space between @ and cursor (would mean mention ended)
			var textAfterAt = textBeforeCursor.substring(lastAtIndex + 1);
			if (textAfterAt.includes(' ') || textAfterAt.includes('\n')) {
				this.hide();
				return;
			}
			
			// We have an active mention
			state.mentionQuery = textAfterAt.toLowerCase();
			state.mentionStartPos = lastAtIndex;
			state.selectedMentionIndex = 0;
			
			this.show(inputEl);
		},
		
		// Show the dropdown
		show: function(inputEl) {
			var dropdown = document.getElementById('mentionDropdown');
			var list = document.getElementById('mentionList');
			if (!dropdown || !list) return;
			
			// Filter documents
			var matches = state.indexedDocs.filter(function(doc) {
				return doc.filename.toLowerCase().includes(state.mentionQuery);
			});
			
			// Build list HTML
			var html = '';
			
			if (matches.length > 0) {
				matches.forEach(function(doc, idx) {
					html += '<div class="mention-item' + (idx === state.selectedMentionIndex ? ' selected' : '') + '" data-doc-name="' + escapeHtml(doc.filename) + '">' +
						'<div class="mention-item-icon">' +
							'<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
								'<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
								'<polyline points="14 2 14 8 20 8"/>' +
							'</svg>' +
						'</div>' +
						'<span class="mention-item-name">' + escapeHtml(doc.filename) + '</span>' +
					'</div>';
				});
			} else if (state.mentionQuery.length > 0) {
				html = '<div class="mention-empty">No documents match "' + escapeHtml(state.mentionQuery) + '"</div>';
			}
			
			// Add option to index new
			html += '<div class="mention-add-new" id="mentionAddNew">' +
				'<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
					'<line x1="12" y1="5" x2="12" y2="19"/>' +
					'<line x1="5" y1="12" x2="19" y2="12"/>' +
				'</svg>' +
				'<span>Add reference document...</span>' +
			'</div>';
			
			list.innerHTML = html;
			
			// Position dropdown near the input
			var inputRect = inputEl.getBoundingClientRect();
			dropdown.style.left = inputRect.left + 'px';
			dropdown.style.bottom = (window.innerHeight - inputRect.top + 8) + 'px';
			
			dropdown.classList.remove('hidden');
			
			// Add click listeners
			var items = list.querySelectorAll('.mention-item');
			items.forEach(function(item) {
				item.addEventListener('click', function() {
					var docName = item.getAttribute('data-doc-name');
					MentionAutocomplete.selectDocument(docName, inputEl);
				});
			});
			
			var addNew = document.getElementById('mentionAddNew');
			if (addNew) {
				addNew.addEventListener('click', function() {
					MentionAutocomplete.hide();
					DocIndex.showModal();
				});
			}
		},
		
		// Hide the dropdown
		hide: function() {
			var dropdown = document.getElementById('mentionDropdown');
			if (dropdown) {
				dropdown.classList.add('hidden');
			}
			state.mentionStartPos = -1;
			state.mentionQuery = '';
		},
		
		// Select a document
		selectDocument: function(docName, inputEl) {
			if (state.mentionStartPos === -1) return;
			
			var value = inputEl.value;
			var before = value.substring(0, state.mentionStartPos);
			var after = value.substring(inputEl.selectionStart);
			
			// Insert the document reference
			inputEl.value = before + '@' + docName + ' ' + after;
			
			// Move cursor after the inserted text
			var newPos = state.mentionStartPos + docName.length + 2;
			inputEl.setSelectionRange(newPos, newPos);
			inputEl.focus();
			
			this.hide();
		},
		
		// Handle keyboard navigation
		handleKeyDown: function(e, inputEl) {
			var dropdown = document.getElementById('mentionDropdown');
			if (!dropdown || dropdown.classList.contains('hidden')) return false;
			
			var items = dropdown.querySelectorAll('.mention-item');
			
			if (e.key === 'ArrowDown') {
				e.preventDefault();
				state.selectedMentionIndex = Math.min(state.selectedMentionIndex + 1, items.length - 1);
				this.updateSelection(items);
				return true;
			}
			
			if (e.key === 'ArrowUp') {
				e.preventDefault();
				state.selectedMentionIndex = Math.max(state.selectedMentionIndex - 1, 0);
				this.updateSelection(items);
				return true;
			}
			
			if (e.key === 'Enter' || e.key === 'Tab') {
				if (items.length > 0 && items[state.selectedMentionIndex]) {
					e.preventDefault();
					var docName = items[state.selectedMentionIndex].getAttribute('data-doc-name');
					this.selectDocument(docName, inputEl);
					return true;
				}
			}
			
			if (e.key === 'Escape') {
				e.preventDefault();
				this.hide();
				return true;
			}
			
			return false;
		},
		
		// Update visual selection
		updateSelection: function(items) {
			items.forEach(function(item, idx) {
				item.classList.toggle('selected', idx === state.selectedMentionIndex);
			});
		}
	};

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
			modeIndicator: document.getElementById('modeIndicator'),
			// Document indexing elements
			indexDocsBtn: document.getElementById('indexDocsBtn'),
			indexBadge: document.getElementById('indexBadge'),
			indexModalOverlay: document.getElementById('indexModalOverlay'),
			mentionDropdown: document.getElementById('mentionDropdown')
		};
		
		// Setup event listeners (can be done before document ID is ready)
		setupEventListeners();
		
		// Setup document indexing modal events
		DocIndex.setupModalEvents();
		
		// Initialize document-specific ID first, then load chats and documents
		// This ensures chats and reference docs are tied to the specific Word document
		DocIndex.initDocumentId().then(function(docId) {
			console.log('Copilot: Document session initialized for:', docId);
			
			// Now load chats for this specific document
		state.chats = Storage.load();
		
			// Create new chat if none exist for this document
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
		
			// Load indexed docs for this specific document
			DocIndex.fetchDocuments();
		}).catch(function(err) {
			console.warn('Copilot: Document ID init failed, using fallback:', err);
			// Fallback: load chats with generic storage
			state.chats = Storage.load();
			if (state.chats.length === 0) {
				createNewChat();
			} else {
				state.currentChatId = state.chats[0].id;
				renderChatList();
				renderMessages();
			}
			DocIndex.fetchDocuments();
		});
		
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
