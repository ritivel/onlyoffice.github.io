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
		// Backend URL - dynamically detected based on current hostname
		// All AI requests go through agents-backend for consistent architecture
		BACKEND_URL: (function() {
			try {
				var hostname = window.parent.location.hostname || window.location.hostname;
				var url = 'http://' + hostname + ':8000';
				console.log('[Copilot Config] Detected hostname:', hostname);
				console.log('[Copilot Config] Backend URL:', url);
				return url;
			} catch (e) {
				console.log('[Copilot Config] Error detecting hostname:', e);
				console.log('[Copilot Config] Using fallback: http://localhost:8000');
				return 'http://localhost:8000';
			}
		})(),
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
		selectedAgent: 'csr', // Selected agent ID for edit mode
		isGenerating: false,
		sidebarCollapsed: true,
		abortController: null, // For cancelling fetch requests
		// Document indexing state
		editorDocId: null, // Current document ID for indexing
		indexedDocs: [], // List of indexed documents
		mentionQuery: '', // Current @ mention search query
		mentionStartPos: -1, // Position where @ was typed
		selectedMentionIndex: 0, // Currently selected item in mention dropdown
		// Progress stages state
		currentStage: null,
		completedStages: [],
		// Task planning state (for complex multi-step tasks)
		activePlanId: null,
		progressContainerId: null,
		// Sources collected from tool calls for current message (for numbered citations)
		currentMessageSources: []
	};

	// ============================================
	// AGENT DEFINITIONS
	// ============================================
	var AGENTS = {
		csr: {
			id: 'csr',
			name: 'CSR Agent',
			icon: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>',
			description: 'Clinical Study Report drafting specialist',
			capabilities: [
				'End-to-end CSR drafting based on internal documents and ICH guidelines',
				'Template filling and completion',
				'Section-wise or end-to-end editing',
				'Format-aware writing (headings, fonts, colors, styles)',
				'Automatic template placeholder removal'
			],
			systemPrompt: `You are a Clinical Study Report (CSR) Agent specialized in medical/regulatory document writing.

CONTEXT AWARENESS (READ THIS FIRST):
You receive detailed context about the document including:
- position_summary: Human-readable description of cursor location
- current_context: Content type (heading_1, paragraph, list_item, etc.), formatting info
- current_section: Which section the cursor is in, parent heading
- document_outline/section_map: Full document structure with heading hierarchy

ALWAYS check position_summary and current_context BEFORE making any edits to understand:
- Are you in a heading or paragraph?
- What section are you in?
- Is there selected text to work with?
- What's the current formatting/alignment?

YOUR CORE CAPABILITIES:
1. **End-to-End CSR Drafting**: Draft complete CSR sections based on internal documents and ICH E3 guidelines with proper formatting
2. **Template Completion**: When a template is loaded, identify placeholders (like [STUDY TITLE], <<INSERT>>, {PLACEHOLDER}) and fill them with appropriate content
3. **Partial Document Completion**: When working with partially filled CSRs, identify incomplete sections and help complete them
4. **Section-wise or Full Editing**: Support both targeted section editing and comprehensive document-wide updates
5. **Template Cleanup**: Remove template instructions, placeholder text, and guidance notes when inserting actual content

CURSOR POSITIONING RULES (CRITICAL):
1. BEFORE inserting content, navigate to the correct location using insert_at_heading or search tools
2. NEVER insert heading content when cursor is in a paragraph - move first!
3. NEVER insert paragraph content when cursor is in a heading - move to content area first!
4. Use the section_map to find the exact heading to insert under
5. When adding new sections, position AFTER the previous section's content

FORMATTING PRIORITIES (CRITICAL):
- Preserve and apply proper heading styles (Heading 1, Heading 2, etc.)
- Maintain consistent font styling (size, color, weight) matching the document
- Use appropriate paragraph spacing and indentation
- Apply table formatting with proper borders and cell styling
- Respect existing document styles and extend them consistently
- Use bold, italic, underline appropriately for emphasis
- Ensure numbered lists and bullet points maintain proper formatting

CSR SECTION CONTENT GUIDELINES:
- Title Page: Study title, protocol number, sponsor info - keep formatting exact
- Synopsis: 2-3 pages summarizing entire study, structured with sub-headings
- Introduction: 1-2 pages, disease background, rationale - paragraph format
- Study Objectives: Clear numbered/bulleted lists of primary/secondary objectives
- Study Design: Detailed description, often 3-5 pages with subsections
- Study Population: Inclusion/exclusion criteria - typically formatted lists
- Efficacy/Safety Results: Data-heavy sections with tables and figures
- Discussion/Conclusions: Paragraph format, 2-4 pages synthesizing results

WHEN RESPONDING:
- Always analyze the current document structure first (use get_document_map if needed)
- Check position_summary to understand cursor location before ANY insertion
- Identify whether you're working with a template, partial CSR, or blank document
- Reference ICH E3 guidelines for proper CSR structure and content requirements
- When filling placeholders, completely remove the placeholder markers
- Maintain regulatory compliance language and terminology
- Preserve the professional tone expected in regulatory submissions

DOCUMENT ANALYSIS:
- Detect template markers: [BRACKETS], <<ANGLE BRACKETS>>, {CURLY BRACES}, CAPS_PLACEHOLDERS
- Identify incomplete sections by looking for: "TBD", "To be determined", "Insert", "Complete this section"
- Recognize CSR structure: Title Page, Synopsis, Table of Contents, Ethics, Investigators, Study Plan, etc.`,
			placeholder: 'Describe what you want to do with your CSR document...'
		},
		general: {
			id: 'general',
			name: 'General Editor',
			icon: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>',
			description: 'Format-aware document editing with proper structure',
			capabilities: [
				'Document structure and outline awareness',
				'Heading hierarchy management (H1 → H2 → H3)',
				'Proper cursor positioning before insertions',
				'Content length guidelines per section type',
				'Style-consistent formatting'
			],
			systemPrompt: `You are a format-aware document editing assistant with deep understanding of document structure.

CONTEXT AWARENESS (READ THIS FIRST):
You receive detailed context about the document. ALWAYS check these fields BEFORE any edit:

1. **position_summary**: Human-readable description like "Page 2 of 5 | Cursor is in a Heading 2 | Under section: Introduction"
2. **current_context**: 
   - content_type: "heading_1", "heading_2", "paragraph", "list_item", "title", etc.
   - is_heading: true/false
   - heading_level: 1-6 if in heading
   - is_empty: true if empty paragraph
   - formatting: {alignment, indent_left, spacing}
3. **current_section**: Which section you're in, parent heading
4. **section_map**: Full document structure - use this to navigate!

CRITICAL: CURSOR POSITIONING & DOCUMENT FLOW
Before inserting ANY content, you MUST:
1. READ position_summary and current_context FIRST
2. If current_context.is_heading is true, you're IN A HEADING - don't insert paragraph content there!
3. If current_context.content_type is "paragraph", don't insert heading-style content!
4. Use insert_at_heading tool to navigate to correct sections before inserting
5. Use section_map to find where content should go

DOCUMENT STRUCTURE AWARENESS:
- Always analyze the document outline (section_map) before making changes
- Understand the heading hierarchy: H1 (main sections) → H2 (subsections) → H3 (sub-subsections)
- NEVER insert heading-level content in the middle of a paragraph
- NEVER insert paragraph content where a heading is expected
- Respect existing section boundaries

STYLE DETECTION & CONSISTENCY:
Before writing, check current_context.formatting and match:
- Current paragraph style (from current_paragraph_style field)
- Alignment (from formatting.alignment)
- Indentation level (from formatting.indent_left)
- Whether you're in a list context (content_type contains "list")

CONTENT LENGTH GUIDELINES BY SECTION TYPE:
- Title/Document Header: 1 line, concise
- Heading 1 (Main Section): 3-10 words, descriptive
- Heading 2 (Subsection): 3-8 words
- Heading 3 (Sub-subsection): 2-6 words
- Introduction paragraphs: 3-5 sentences setting context
- Body paragraphs: 4-8 sentences with one main idea each
- Conclusion paragraphs: 2-4 sentences summarizing
- List items: 1-2 sentences each, parallel structure
- Table cells: Brief, data-focused content

INSERTION WORKFLOW:
1. Check position_summary - where is cursor now?
2. Check current_context - what type of content is cursor in?
3. Use section_map to find target location
4. Navigate to correct position using insert_at_heading or search tools
5. Verify you're in the right content type
6. Insert content matching the target format

FORMATTING COMMANDS:
When the user says:
- "Add a section about X" → Navigate to end of previous section, create Heading + intro paragraph
- "Write about X" → Check current_context, if in paragraph area, add paragraphs; if not, navigate first
- "Add bullet points" → Navigate to content area (not heading), create properly formatted list
- "Expand this" → Add content AFTER the referenced text, matching its style
- "Insert before/after" → Use search to find location, navigate precisely, then insert

NEVER:
- Insert content without first checking position_summary and current_context
- Insert content at random cursor positions
- Mix heading styles within the same hierarchy level
- Break existing paragraph flow with misplaced headings
- Ignore the document's established formatting patterns
- Insert raw text without proper paragraph/style formatting`,
			placeholder: 'Tell me what to do with the document...'
		}
	};

	// ============================================
	// PROGRESS STAGES DEFINITION
	// ============================================
	var AGENT_STAGES = [
		{ 
			id: 'understanding', 
			icon: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2a8 8 0 0 1 8 8c0 2.5-1.2 4.8-3 6.2V19a2 2 0 0 1-2 2h-6a2 2 0 0 1-2-2v-2.8C5.2 14.8 4 12.5 4 10a8 8 0 0 1 8-8z"/><path d="M12 2v4"/><path d="M9 21v-2"/><path d="M15 21v-2"/></svg>',
			label: 'Understanding Query', 
			description: 'Analyzing your question...' 
		},
		{ 
			id: 'searching', 
			icon: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/></svg>',
			label: 'Searching Knowledge', 
			description: 'Querying regulatory databases...' 
		},
		{ 
			id: 'sources', 
			icon: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/><path d="M8 7h8"/><path d="M8 11h6"/></svg>',
			label: 'Processing Sources', 
			description: 'Reviewing relevant documents...' 
		},
		{ 
			id: 'synthesizing', 
			icon: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 19l7-7 3 3-7 7-3-3z"/><path d="M18 13l-1.5-7.5L2 2l3.5 14.5L13 18l5-5z"/><path d="M2 2l7.586 7.586"/></svg>',
			label: 'Synthesizing Answer', 
			description: 'Generating comprehensive response...' 
		},
		{ 
			id: 'complete', 
			icon: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
			label: 'Complete', 
			description: 'Response ready' 
		}
	];

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
		
		// AI Authorship helper - uses the SDK's SetAssistantTrackRevisions API
		setAIAuthor: function() {
			return new Promise(function(resolve) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					console.log('[AI Authorship] Calling SetAssistantTrackRevisions(true, "Riti-AI")');
					window.Asc.plugin.callCommand(function() {
						// DO NOT add console.log here - this function runs in document context
						return Api.GetDocument().SetAssistantTrackRevisions(true, "Riti-AI");
					}, false, true, function(result) {
						console.log('[AI Authorship] Set AI author completed, result:', result);
						resolve();
					});
				} else {
					console.log('[AI Authorship] Asc.plugin.callCommand not available');
					resolve();
				}
			});
		},
		
		// Restore original user as author after AI edits
		restoreUserAuthor: function() {
			return new Promise(function(resolve) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					// Add delay to ensure the edit operation fully completes
					console.log('[AI Authorship] Scheduling restore in 200ms');
					setTimeout(function() {
						console.log('[AI Authorship] Calling SetAssistantTrackRevisions(false)');
						window.Asc.plugin.callCommand(function() {
							// DO NOT add console.log here - this function runs in document context
							return Api.GetDocument().SetAssistantTrackRevisions(false);
						}, false, true, function(result) {
							console.log('[AI Authorship] Restore completed, result:', result);
							resolve();
						});
					}, 200);
				} else {
					console.log('[AI Authorship] Asc.plugin.callCommand not available for restore');
					resolve();
				}
			});
		},
		
		// Wrapper to execute an edit operation with AI authorship
		executeWithAIAuthor: async function(editFn) {
			console.log('[AI Authorship] === Starting AI edit operation ===');
			await this.setAIAuthor();
			console.log('[AI Authorship] AI author set, now executing edit function');
			var result;
			var error;
			try {
				result = await editFn();
				console.log('[AI Authorship] Edit function completed successfully');
			} catch (e) {
				error = e;
				console.error('[AI Authorship] Error during edit:', e);
			}
			// Always restore user, even after errors
			console.log('[AI Authorship] About to restore user');
			await this.restoreUserAuthor();
			console.log('[AI Authorship] === AI edit operation complete ===');
			if (error) {
				throw error;
			}
			return result;
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

			// Layout constraints tool - helps agent understand spatial constraints
			// Critical for proper formatting, table sizing, and content layout
			'get_layout_constraints': async function() {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var sections = doc.GetSections ? doc.GetSections() : [];
							
							// Default values for Letter paper
							var result = {
								page_width_inches: 8.5,
								page_height_inches: 11,
								text_width_inches: 6.5,  // After margins
								chars_per_line: 80,      // Approx at 12pt
								orientation: 'portrait',
								margins: {
									top: 1.0,
									bottom: 1.0,
									left: 1.0,
									right: 1.0
								},
								recommended_table_cols: 5,  // Max readable columns
								formatting_tips: [
									'Keep table columns under 6 for readability',
									'Column headers should be 15 chars or less',
									'Use abbreviations in tables when appropriate',
									'Long text should be in paragraph form, not tables',
									'Consider using bullet lists for short items'
								]
							};
							
							// Try to get actual page dimensions from first section
							if (sections && sections.length > 0) {
								var section = sections[0];
								try {
									// GetPageSize returns {Width, Height} in twips (1/1440 inch)
									if (section.GetPageSize) {
										var size = section.GetPageSize();
										if (size && size.Width && size.Height) {
											result.page_width_inches = Math.round(size.Width / 1440 * 10) / 10;
											result.page_height_inches = Math.round(size.Height / 1440 * 10) / 10;
											result.orientation = size.Width > size.Height ? 'landscape' : 'portrait';
										}
									}
									// Get margins
									if (section.GetPageMargins) {
										var margins = section.GetPageMargins();
										if (margins) {
											result.margins = {
												top: Math.round(margins.Top / 1440 * 10) / 10,
												bottom: Math.round(margins.Bottom / 1440 * 10) / 10,
												left: Math.round(margins.Left / 1440 * 10) / 10,
												right: Math.round(margins.Right / 1440 * 10) / 10
											};
											result.text_width_inches = result.page_width_inches - result.margins.left - result.margins.right;
											result.text_width_inches = Math.round(result.text_width_inches * 10) / 10;
										}
									}
								} catch (e) {
									// Use defaults on error
								}
								
								// Calculate chars per line and recommended table columns
								// Assumes ~10 chars per inch at 12pt font
								result.chars_per_line = Math.round(result.text_width_inches * 12);
								result.recommended_table_cols = Math.min(6, Math.floor(result.text_width_inches / 1.2));
							}
							
							return result;
						}, false, false, resolve);
					});
				}
				// Return sensible defaults when not in OnlyOffice
				return {
					page_width_inches: 8.5,
					text_width_inches: 6.5,
					chars_per_line: 80,
					orientation: 'portrait',
					margins: { top: 1, bottom: 1, left: 1, right: 1 },
					recommended_table_cols: 5,
					formatting_tips: [
						'Keep table columns under 6 for readability',
						'Column headers should be 15 chars or less'
					],
					error: 'Not in OnlyOffice environment - using defaults'
				};
			},

			'get_current_paragraph': async function() {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var para = doc.GetCurrentParagraph ? doc.GetCurrentParagraph() : null;
							var text = para && para.GetText ? para.GetText({ NewLineSeparator: '\n' }) : '';
							var styleName = '';
							var isHeading = false;
							var headingLevel = 0;
							var indentLeft = 0;
							var spacingBefore = 0;
							var spacingAfter = 0;
							var alignment = 'left';
							
							try {
								var style = para && para.GetStyle ? para.GetStyle() : null;
								styleName = style && style.GetName ? style.GetName() : '';
								
								// Check if it's a heading style
								if (styleName.indexOf('Heading') === 0 || styleName.indexOf('heading') === 0) {
									isHeading = true;
									var levelMatch = styleName.match(/\d+/);
									headingLevel = levelMatch ? parseInt(levelMatch[0]) : 1;
								}
								
								// Get paragraph formatting
								if (para) {
									try {
										var indent = para.GetIndLeft ? para.GetIndLeft() : 0;
										indentLeft = indent || 0;
									} catch (e) {}
									
									try {
										var before = para.GetSpacingBefore ? para.GetSpacingBefore() : 0;
										spacingBefore = before || 0;
									} catch (e) {}
									
									try {
										var after = para.GetSpacingAfter ? para.GetSpacingAfter() : 0;
										spacingAfter = after || 0;
									} catch (e) {}
									
									try {
										var jc = para.GetJc ? para.GetJc() : 'left';
										alignment = jc || 'left';
									} catch (e) {}
								}
							} catch (e) {}
							
							// Determine content type
							var contentType = 'paragraph';
							if (isHeading) {
								contentType = 'heading_' + headingLevel;
							} else if (styleName === 'Title') {
								contentType = 'title';
							} else if (styleName === 'Subtitle') {
								contentType = 'subtitle';
							} else if (styleName.indexOf('List') !== -1) {
								contentType = 'list_item';
							} else if (styleName.indexOf('TOC') !== -1) {
								contentType = 'table_of_contents';
							}
							
							return { 
								text: text, 
								style: styleName,
								content_type: contentType,
								is_heading: isHeading,
								heading_level: headingLevel,
								formatting: {
									indent_left: indentLeft,
									spacing_before: spacingBefore,
									spacing_after: spacingAfter,
									alignment: alignment
								},
								char_count: text.length,
								is_empty: text.trim().length === 0
							};
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
					
					// Wrap edit with AI authorship
					return ToolExecutor.executeWithAIAuthor(async function() {
						return new Promise(function(resolve) {
							window.Asc.plugin.executeMethod('PasteHtml', [html], function(result) {
								resolve({ success: true, rows: rows, cols: cols, message: 'Table inserted' });
							});
						});
					});
				}
				return { success: false, error: 'Not in OnlyOffice environment' };
			},

			'insert_image': async function(params) {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					var self = this;
					// Wrap edit with AI authorship
					return ToolExecutor.executeWithAIAuthor(async function() {
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
				
				// Wrap edit with AI authorship
				return ToolExecutor.executeWithAIAuthor(async function() {
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
				});
			}
			console.log('Not in OnlyOffice environment');
			return { success: false, error: 'Not in OnlyOffice environment' };
		},
		
		// Add an endnote (reference) at cursor position OR after specific anchor text
		'add_endnote': async function(params) {
			console.log('add_endnote called with:', JSON.stringify(params));
			if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
				// Handle both reference_text and text parameter names
				var referenceText = params.reference_text || params.text || '';
				var anchorText = params.anchor_text || ''; // Optional: text to position after
				console.log('add_endnote referenceText:', referenceText, 'anchorText:', anchorText);
				
				if (!referenceText) {
					return { success: false, error: 'reference_text is required' };
				}
				
				// Use Asc.scope to pass data to callCommand (critical for serialization)
				window.Asc.scope.endnoteText = String(referenceText);
				window.Asc.scope.anchorText = String(anchorText);
				
				// Wrap edit with AI authorship
				return ToolExecutor.executeWithAIAuthor(async function() {
					return new Promise(function(resolve) {
						// isNoCalc=false ensures document recalculates endnote numbers
						// isNoUndo=false creates undo point for each endnote
						window.Asc.plugin.callCommand(function() {
							var textToAdd = Asc.scope.endnoteText;
							var anchor = Asc.scope.anchorText;
							var doc = Api.GetDocument();
							var logicDoc = doc.Document; // Get internal document for SetDocPosType
							
							// IMPORTANT: Ensure we're in the main document body before searching
							// docpostype_Content = 0x00
							if (logicDoc && logicDoc.SetDocPosType) {
								logicDoc.SetDocPosType(0x00); // docpostype_Content
							}
							
							// If anchor_text provided, find it and position cursor after it
							var anchorFound = false;
							if (anchor && anchor.length > 0) {
								var ranges = doc.Search ? doc.Search(anchor, false) : [];
								if (ranges && ranges.length > 0) {
									// Select the first match and move cursor to end of it
									var range = ranges[0];
									if (range && range.Select) {
										range.Select(true);
										anchorFound = true;
									}
									// Move cursor to the end of the selected range
									if (range && range.MoveCursorToPos) {
										range.MoveCursorToPos(false); // false = move to end
									}
								}
							}
							
							// 1. Insert the endnote marker at cursor position
							doc.AddEndnote();
							
							// 2. Get all first paragraphs of endnotes
							var endnoteParas = doc.GetEndNotesFirstParagraphs();
							
							if (endnoteParas && endnoteParas.length > 0) {
								// 3. Get the LAST one - it's the one we just added
								var newestEndnote = endnoteParas[endnoteParas.length - 1];
								
								// 4. Add the reference text to this paragraph
								if (newestEndnote && newestEndnote.AddText) {
									newestEndnote.AddText(textToAdd);
								}
							}
							
							// CRITICAL: Move cursor back to main document body after adding endnote
							// This ensures the next add_endnote call can search in the main document
							if (logicDoc && logicDoc.SetDocPosType) {
								logicDoc.SetDocPosType(0x00); // docpostype_Content
								if (logicDoc.RemoveSelection) {
									logicDoc.RemoveSelection();
								}
							}
							
							return { success: true, message: 'Endnote added with reference text', anchor_found: anchorFound, anchor_used: anchor || 'cursor position' };
						}, false, false, function(result) {
							console.log('add_endnote result:', result);
							resolve(result || { success: true });
						});
					});
				});
			}
			return { success: false, error: 'Not in OnlyOffice environment' };
		},
		
		// Add a footnote at cursor position OR after specific anchor text
		'add_footnote': async function(params) {
			console.log('add_footnote called with:', JSON.stringify(params));
			if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
				// Handle both reference_text and text parameter names
				var referenceText = params.reference_text || params.text || '';
				var anchorText = params.anchor_text || ''; // Optional: text to position after
				console.log('add_footnote referenceText:', referenceText, 'anchorText:', anchorText);
				
				if (!referenceText) {
					return { success: false, error: 'reference_text is required' };
				}
				
				// Use Asc.scope to pass data to callCommand (critical for serialization)
				window.Asc.scope.footnoteText = String(referenceText);
				window.Asc.scope.anchorText = String(anchorText);
				
				// Wrap edit with AI authorship
				return ToolExecutor.executeWithAIAuthor(async function() {
					return new Promise(function(resolve) {
						// isNoCalc=false ensures document recalculates footnote numbers
						// isNoUndo=false creates undo point for each footnote
						window.Asc.plugin.callCommand(function() {
							var textToAdd = Asc.scope.footnoteText;
							var anchor = Asc.scope.anchorText;
							var doc = Api.GetDocument();
							var logicDoc = doc.Document; // Get internal document for SetDocPosType
							
							// IMPORTANT: Ensure we're in the main document body before searching
							// docpostype_Content = 0x00
							if (logicDoc && logicDoc.SetDocPosType) {
								logicDoc.SetDocPosType(0x00); // docpostype_Content
							}
							
							// If anchor_text provided, find it and position cursor after it
							var anchorFound = false;
							if (anchor && anchor.length > 0) {
								var ranges = doc.Search ? doc.Search(anchor, false) : [];
								if (ranges && ranges.length > 0) {
									// Select the first match and move cursor to end of it
									var range = ranges[0];
									if (range && range.Select) {
										range.Select(true);
										anchorFound = true;
									}
									// Move cursor to the end of the selected range
									if (range && range.MoveCursorToPos) {
										range.MoveCursorToPos(false); // false = move to end
									}
								}
							}
							
							// 1. Insert the footnote marker at cursor position
							doc.AddFootnote();
							
							// 2. Get all first paragraphs of footnotes
							var footnoteParas = doc.GetFootnotesFirstParagraphs();
							
							if (footnoteParas && footnoteParas.length > 0) {
								// 3. Get the LAST one - it's the one we just added
								var newestFootnote = footnoteParas[footnoteParas.length - 1];
								
								// 4. Add the reference text to this paragraph
								if (newestFootnote && newestFootnote.AddText) {
									newestFootnote.AddText(textToAdd);
								}
							}
							
							// CRITICAL: Move cursor back to main document body after adding footnote
							// This ensures the next add_footnote call can search in the main document
							if (logicDoc && logicDoc.SetDocPosType) {
								logicDoc.SetDocPosType(0x00); // docpostype_Content
								if (logicDoc.RemoveSelection) {
									logicDoc.RemoveSelection();
								}
							}
							
							return { success: true, message: 'Footnote added with reference text', anchor_found: anchorFound, anchor_used: anchor || 'cursor position' };
						}, false, false, function(result) {
							console.log('add_footnote result:', result);
							resolve(result || { success: true });
						});
					});
				});
			}
			return { success: false, error: 'Not in OnlyOffice environment' };
		},
			
			'replace_selection': async function(params) {
				console.log('replace_selection called with:', params);
				var text = params.text || '';
				
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					// Wrap edit with AI authorship
					return ToolExecutor.executeWithAIAuthor(async function() {
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
					// Wrap edit with AI authorship
					return ToolExecutor.executeWithAIAuthor(async function() {
						return new Promise(function(resolve) {
							// InputText with empty string deletes the current selection
							window.Asc.plugin.executeMethod('InputText', [''], function(result) {
								console.log('delete_selection InputText result:', result);
								resolve({ success: true, message: 'Selection deleted' });
							});
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
			},
			
			// =========================================================================
			// EXECUTE_SDKJS: Execute arbitrary JavaScript code in document context
			// This is the primary tool for complex document manipulation
			// =========================================================================
			'execute_sdkjs': async function(params) {
				console.log('execute_sdkjs called with:', params);
				
				if (!params || !params.code) {
					return { success: false, error: 'code parameter is required' };
				}
				
				var code = params.code;
				var needsRecalc = params.needs_recalc !== false; // Default true
				var needsExecuteMethod = !!params.needs_execute_method;
				
				// If using executeMethod (for PasteHtml, PasteText, etc.)
				if (needsExecuteMethod && params.execute_method_name) {
					if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
						return new Promise(function(resolve) {
							var methodName = params.execute_method_name;
							var methodArgs = params.execute_method_args || [];
							console.log('Executing method:', methodName, 'with args:', methodArgs);
							window.Asc.plugin.executeMethod(methodName, methodArgs, function(result) {
								console.log('executeMethod result:', result);
								resolve({ success: true, result: result, method: methodName });
							});
						});
					}
					return { success: false, error: 'executeMethod not available' };
				}
				
				// Execute arbitrary SDKJS code via callCommand
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					return new Promise(function(resolve) {
						try {
							// Wrap the code in a function and execute it
							// The code should return a value which becomes the result
							var wrappedCode = '(function() {\n' +
								'try {\n' +
								code + '\n' +
								'} catch (e) {\n' +
								'  return { success: false, error: e.message || String(e) };\n' +
								'}\n' +
								'})()';
							
							console.log('Executing SDKJS code:', wrappedCode.substring(0, 500) + '...');
							
							// Create a function from the wrapped code
							var execFn = new Function('return ' + wrappedCode);
							
							window.Asc.plugin.callCommand(execFn, needsRecalc, false, function(result) {
								console.log('callCommand result:', result);
								// If result is already an object with success field, use it
								if (result && typeof result === 'object' && 'success' in result) {
									resolve(result);
								} else {
									resolve({ success: true, result: result });
								}
							});
						} catch (e) {
							console.error('execute_sdkjs error:', e);
							resolve({ success: false, error: e.message || String(e) });
						}
					});
				}
				
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			// =========================================================================
			// HIGH-LEVEL SECTION-AWARE EDITING TOOLS
			// These provide semantic, position-aware document editing
			// =========================================================================
			
			'insert_at_heading': async function(params) {
				console.log('insert_at_heading called with:', params);
				
				if (!params || !params.heading_text || !params.content) {
					return { success: false, error: 'heading_text and content are required' };
				}
				
				var headingText = params.heading_text;
				var content = params.content;
				var position = params.position || 'after_heading';
				var format = params.format || 'markdown';
				
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					// First, find the heading and get document structure
					var headingInfo = await new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var count = doc.GetElementsCount();
							var targetIndex = -1;
							var nextHeadingIndex = -1;
							var targetLevel = 0;
							
							// Find the target heading
							for (var i = 0; i < count; i++) {
								var elem = doc.GetElement(i);
								if (elem.GetClassType && elem.GetClassType() === 'paragraph') {
									var style = elem.GetStyle();
									if (style) {
										var styleName = style.GetName ? style.GetName() : '';
										if (styleName.indexOf('Heading') === 0 || styleName.indexOf('heading') === 0) {
											var text = elem.GetText ? elem.GetText() : '';
											// Check if this heading matches (partial match)
											if (text.toLowerCase().indexOf(Asc.scope.headingText.toLowerCase()) !== -1 ||
												Asc.scope.headingText.toLowerCase().indexOf(text.toLowerCase()) !== -1) {
												targetIndex = i;
												// Extract level from style name
												var levelMatch = styleName.match(/\d+/);
												targetLevel = levelMatch ? parseInt(levelMatch[0]) : 1;
												break;
											}
										}
									}
								}
							}
							
							if (targetIndex === -1) {
								return { found: false, error: 'Heading not found: ' + Asc.scope.headingText };
							}
							
							// Find next heading of same or higher level
							for (var j = targetIndex + 1; j < count; j++) {
								var elem2 = doc.GetElement(j);
								if (elem2.GetClassType && elem2.GetClassType() === 'paragraph') {
									var style2 = elem2.GetStyle();
									if (style2) {
										var styleName2 = style2.GetName ? style2.GetName() : '';
										if (styleName2.indexOf('Heading') === 0 || styleName2.indexOf('heading') === 0) {
											var levelMatch2 = styleName2.match(/\d+/);
											var level2 = levelMatch2 ? parseInt(levelMatch2[0]) : 1;
											if (level2 <= targetLevel) {
												nextHeadingIndex = j;
												break;
											}
										}
									}
								}
							}
							
							return {
								found: true,
								targetIndex: targetIndex,
								nextHeadingIndex: nextHeadingIndex,
								targetLevel: targetLevel,
								elementCount: count
							};
						}, false, false, resolve);
					});
					
					// Pass parameters via Asc.scope
					window.Asc.scope.headingText = headingText;
					
					if (!headingInfo || !headingInfo.found) {
						return { success: false, error: headingInfo ? headingInfo.error : 'Failed to find heading' };
					}
					
					// Now insert content at the appropriate position
					// We'll use PasteHtml for formatted content
					var htmlContent = content;
					if (format === 'markdown') {
						htmlContent = markdownToHtml(content);
					}
					
					// Navigate to the correct position based on position parameter
					window.Asc.scope.targetIndex = headingInfo.targetIndex;
					window.Asc.scope.nextHeadingIndex = headingInfo.nextHeadingIndex;
					window.Asc.scope.position = position;
					
					await new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var pos = Asc.scope.position;
							var targetIdx = Asc.scope.targetIndex;
							var nextIdx = Asc.scope.nextHeadingIndex;
							
							// Get the target element to position cursor
							if (pos === 'after_heading') {
								// Position right after the heading
								var headingElem = doc.GetElement(targetIdx);
								if (headingElem && headingElem.SetCursorPos) {
									headingElem.SetCursorPos();
								}
							} else if (pos === 'end_of_section') {
								// Position at end of section (before next heading or end of doc)
								var endIdx = nextIdx > 0 ? nextIdx - 1 : doc.GetElementsCount() - 1;
								var endElem = doc.GetElement(endIdx);
								if (endElem && endElem.SetCursorPos) {
									endElem.SetCursorPos();
								}
							} else if (pos === 'before_heading') {
								// Position before the heading
								if (targetIdx > 0) {
									var prevElem = doc.GetElement(targetIdx - 1);
									if (prevElem && prevElem.SetCursorPos) {
										prevElem.SetCursorPos();
									}
								}
							}
							return { positioned: true };
						}, false, false, resolve);
					});
					
					// Now paste the content with AI authorship
					return ToolExecutor.executeWithAIAuthor(async function() {
						return new Promise(function(resolve) {
							window.Asc.plugin.executeMethod('PasteHtml', ['<br>' + htmlContent], function(result) {
								console.log('insert_at_heading PasteHtml result:', result);
								resolve({ 
									success: true, 
									message: 'Content inserted ' + position + ' "' + headingText + '"',
									heading: headingText,
									position: position
								});
							});
						});
					});
				}
				
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			// Read content by page range - for very long documents
			'read_pages': async function(params) {
				console.log('read_pages called with:', params);
				
				var startPage = (params && params.start_page) ? params.start_page : 1;
				var endPage = (params && params.end_page) ? params.end_page : startPage;
				var maxChars = (params && params.max_chars) ? params.max_chars : 10000;
				
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					window.Asc.scope.startPage = startPage;
					window.Asc.scope.endPage = endPage;
					window.Asc.scope.maxChars = maxChars;
					
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var pageCount = doc.GetPageCount ? doc.GetPageCount() : 0;
							
							// Validate page range
							var start = Math.max(1, Math.min(Asc.scope.startPage, pageCount));
							var end = Math.max(start, Math.min(Asc.scope.endPage, pageCount));
							
							// Get document as markdown and extract page range
							// This is approximate since we don't have exact page boundaries
							var md = doc.ToMarkdown ? doc.ToMarkdown(true, false, false, false) : '';
							
							// Estimate chars per page
							var charsPerPage = md.length / Math.max(1, pageCount);
							var startChar = Math.floor((start - 1) * charsPerPage);
							var endChar = Math.min(md.length, Math.floor(end * charsPerPage));
							
							// Extract content
							var content = md.substring(startChar, Math.min(endChar, startChar + Asc.scope.maxChars));
							
							// Find section boundaries if possible
							if (startChar > 0) {
								// Try to start at a heading
								var headingMatch = content.match(/^[^#]*?(#{1,6}\s)/);
								if (headingMatch && headingMatch.index < 200) {
									content = content.substring(headingMatch.index);
								}
							}
							
							return {
								success: true,
								start_page: start,
								end_page: end,
								total_pages: pageCount,
								content: content,
								content_length: content.length,
								truncated: (endChar - startChar) > Asc.scope.maxChars
							};
						}, false, false, resolve);
					});
				}
				
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			// Get hierarchical structure of a specific section
			'get_subsections': async function(params) {
				console.log('get_subsections called with:', params);
				
				if (!params || !params.parent_heading) {
					return { success: false, error: 'parent_heading is required' };
				}
				
				var parentHeading = params.parent_heading;
				var maxDepth = params.max_depth || 3;
				
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					window.Asc.scope.parentHeading = parentHeading;
					window.Asc.scope.maxDepth = maxDepth;
					
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var count = doc.GetElementsCount();
							var subsections = [];
							var inSection = false;
							var parentLevel = 0;
							var parentCharCount = 0;
							
							for (var i = 0; i < count; i++) {
								var elem = doc.GetElement(i);
								if (elem.GetClassType && elem.GetClassType() === 'paragraph') {
									var style = elem.GetStyle();
									var styleName = style ? (style.GetName ? style.GetName() : '') : '';
									var text = elem.GetText ? elem.GetText() : '';
									
									if (styleName.indexOf('Heading') === 0 || styleName.indexOf('heading') === 0) {
										var levelMatch = styleName.match(/\d+/);
										var level = levelMatch ? parseInt(levelMatch[0]) : 1;
										
										// Check if this is our parent
										if (text.toLowerCase().indexOf(Asc.scope.parentHeading.toLowerCase()) !== -1) {
											inSection = true;
											parentLevel = level;
											continue;
										}
										
										// Check if we've exited the section
										if (inSection && level <= parentLevel) {
											break; // Done with this section
										}
										
										// Record subsection if within depth limit
										if (inSection && level <= parentLevel + Asc.scope.maxDepth) {
											subsections.push({
												heading: text.trim(),
												level: level,
												relative_level: level - parentLevel,
												element_index: i,
												char_count: 0
											});
										}
									} else if (inSection && subsections.length > 0 && text.trim()) {
										// Add char count to most recent subsection
										subsections[subsections.length - 1].char_count += text.length;
									} else if (inSection && subsections.length === 0 && text.trim()) {
										// Content before first subsection
										parentCharCount += text.length;
									}
								}
							}
							
							return {
								success: true,
								parent_heading: Asc.scope.parentHeading,
								parent_content_chars: parentCharCount,
								subsections: subsections,
								total_subsections: subsections.length
							};
						}, false, false, resolve);
					});
				}
				
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			'get_section_content': async function(params) {
				console.log('get_section_content called with:', params);
				
				if (!params || !params.heading_text) {
					return { success: false, error: 'heading_text is required' };
				}
				
				var headingText = params.heading_text;
				var includeSubsections = params.include_subsections || false;
				var maxChars = params.max_chars || 5000;
				
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					window.Asc.scope.headingText = headingText;
					window.Asc.scope.includeSubsections = includeSubsections;
					window.Asc.scope.maxChars = maxChars;
					
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var count = doc.GetElementsCount();
							var targetIndex = -1;
							var targetLevel = 0;
							var endIndex = count;
							var content = [];
							
							// Find the target heading
							for (var i = 0; i < count; i++) {
								var elem = doc.GetElement(i);
								if (elem.GetClassType && elem.GetClassType() === 'paragraph') {
									var style = elem.GetStyle();
									if (style) {
										var styleName = style.GetName ? style.GetName() : '';
										if (styleName.indexOf('Heading') === 0 || styleName.indexOf('heading') === 0) {
											var text = elem.GetText ? elem.GetText() : '';
											if (text.toLowerCase().indexOf(Asc.scope.headingText.toLowerCase()) !== -1) {
												targetIndex = i;
												var levelMatch = styleName.match(/\d+/);
												targetLevel = levelMatch ? parseInt(levelMatch[0]) : 1;
												break;
											}
										}
									}
								}
							}
							
							if (targetIndex === -1) {
								return { success: false, error: 'Section not found: ' + Asc.scope.headingText };
							}
							
							// Find end of section
							for (var j = targetIndex + 1; j < count; j++) {
								var elem2 = doc.GetElement(j);
								if (elem2.GetClassType && elem2.GetClassType() === 'paragraph') {
									var style2 = elem2.GetStyle();
									if (style2) {
										var styleName2 = style2.GetName ? style2.GetName() : '';
										if (styleName2.indexOf('Heading') === 0 || styleName2.indexOf('heading') === 0) {
											var levelMatch2 = styleName2.match(/\d+/);
											var level2 = levelMatch2 ? parseInt(levelMatch2[0]) : 1;
											// Stop at same or higher level heading (unless including subsections)
											if (level2 <= targetLevel || (!Asc.scope.includeSubsections && level2 > targetLevel)) {
												if (!Asc.scope.includeSubsections || level2 <= targetLevel) {
													endIndex = j;
													break;
												}
											}
										}
									}
								}
							}
							
							// Collect content
							var totalChars = 0;
							for (var k = targetIndex; k < endIndex && totalChars < Asc.scope.maxChars; k++) {
								var elem3 = doc.GetElement(k);
								var text3 = elem3.GetText ? elem3.GetText() : '';
								content.push(text3);
								totalChars += text3.length;
							}
							
							return {
								success: true,
								heading: doc.GetElement(targetIndex).GetText(),
								level: targetLevel,
								content: content.join('\n\n'),
								char_count: totalChars,
								truncated: totalChars >= Asc.scope.maxChars
							};
						}, false, false, resolve);
					});
				}
				
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			'find_and_replace': async function(params) {
				console.log('find_and_replace called with:', params);
				
				if (!params || !params.find_text) {
					return { success: false, error: 'find_text is required' };
				}
				
				var findText = params.find_text;
				var replaceWith = params.replace_with || '';
				var inSection = params.in_section || null;
				var matchCase = params.match_case || false;
				var replaceAll = params.replace_all !== false; // default true
				
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					window.Asc.scope.findText = findText;
					window.Asc.scope.replaceWith = replaceWith;
					window.Asc.scope.matchCase = matchCase;
					window.Asc.scope.replaceAll = replaceAll;
					window.Asc.scope.inSection = inSection;
					
					return ToolExecutor.executeWithAIAuthor(async function() {
						return new Promise(function(resolve) {
							window.Asc.plugin.callCommand(function() {
								var doc = Api.GetDocument();
								var find = Asc.scope.findText;
								var replace = Asc.scope.replaceWith;
								var mc = Asc.scope.matchCase;
								var all = Asc.scope.replaceAll;
								
								var ranges = doc.Search(find, mc);
								var replacedCount = 0;
								
								if (ranges && ranges.length > 0) {
									// Replace either first or all
									var limit = all ? ranges.length : 1;
									for (var i = 0; i < limit; i++) {
										var range = ranges[i];
										if (range) {
											range.Select(true);
											// Delete and insert replacement
											range.Delete();
											if (replace) {
												var para = doc.GetCurrentParagraph();
												if (para) {
													para.AddText(replace);
												}
											}
											replacedCount++;
										}
									}
								}
								
								return {
									success: true,
									replaced_count: replacedCount,
									find_text: find,
									replace_with: replace
								};
							}, true, false, resolve);
						});
					});
				}
				
				return { success: false, error: 'Not in OnlyOffice environment' };
			},
			
			// Enhanced document map with hierarchical structure for long documents
			'get_document_map': async function(params) {
				console.log('get_document_map called with:', params);
				
				var maxDepth = (params && params.max_depth) ? params.max_depth : 6;
				var includePageNumbers = (params && params.include_page_numbers) ? params.include_page_numbers : false;
				var parentHeading = (params && params.parent_heading) ? params.parent_heading : null;
				
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand) {
					window.Asc.scope.maxDepth = maxDepth;
					window.Asc.scope.parentHeading = parentHeading;
					
					return new Promise(function(resolve) {
						window.Asc.plugin.callCommand(function() {
							var doc = Api.GetDocument();
							var count = doc.GetElementsCount();
							var sections = [];
							var currentSection = null;
							var sectionCounter = 0;
							var inTargetSection = !Asc.scope.parentHeading; // If no parent, include all
							var targetLevel = 0;
							
							for (var i = 0; i < count; i++) {
								var elem = doc.GetElement(i);
								if (elem.GetClassType && elem.GetClassType() === 'paragraph') {
									var style = elem.GetStyle();
									var styleName = style ? (style.GetName ? style.GetName() : '') : '';
									var text = elem.GetText ? elem.GetText() : '';
									
									if (styleName.indexOf('Heading') === 0 || styleName.indexOf('heading') === 0) {
										var levelMatch = styleName.match(/\d+/);
										var level = levelMatch ? parseInt(levelMatch[0]) : 1;
										
										// If filtering by parent heading
										if (Asc.scope.parentHeading) {
											if (text.toLowerCase().indexOf(Asc.scope.parentHeading.toLowerCase()) !== -1) {
												inTargetSection = true;
												targetLevel = level;
												continue; // Skip the parent itself
											}
											if (inTargetSection && level <= targetLevel) {
												inTargetSection = false; // Exited the section
											}
											if (!inTargetSection) continue;
										}
										
										// Only include up to maxDepth
										if (level > Asc.scope.maxDepth) continue;
										
										// Save previous section
										if (currentSection) {
											sections.push(currentSection);
										}
										
										sectionCounter++;
										currentSection = {
											id: 'sec_' + sectionCounter,
											heading: text.trim(),
											level: level,
											char_count: 0,
											element_index: i,
											has_content: false,
											subsection_count: 0
										};
									} else if (currentSection && text.trim()) {
										// Content paragraph
										currentSection.char_count += text.length;
										currentSection.has_content = true;
									}
								}
							}
							
							// Add last section
							if (currentSection) {
								sections.push(currentSection);
							}
							
							return {
								success: true,
								sections: sections,
								total_sections: sections.length
							};
						}, false, false, resolve);
					});
				}
				
				return { success: false, error: 'Not in OnlyOffice environment' };
			}
		},
		
		execute: async function(toolName, params) {
			// Check both this.tools[toolName] and this[toolName] due to structure variations
			var toolFn = this.tools[toolName] || this[toolName];
			if (!toolFn || typeof toolFn !== 'function') {
				console.error('Unknown tool:', toolName, 'Available in tools:', Object.keys(this.tools), 'Direct props:', Object.keys(this).filter(function(k) { return typeof this[k] === 'function' && k !== 'execute'; }.bind(this)));
				return { success: false, error: 'Unknown tool: ' + toolName };
			}
			try {
				// Use .call(this, ...) to preserve context
				var result = await toolFn.call(this, params || {});
				return { success: true, result: result };
			} catch (e) {
				console.error('Tool execution error for', toolName, ':', e);
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

	// ============================================
	// TRACEABILITY: Wrap source-derived content with metadata
	// ============================================
	
	
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
			'get_layout_constraints': '📐',
			'get_page_info': '📄',
			'get_document_map': '🗺️',
			'get_section_content': '📖',
			'get_subsections': '🔽',
			'read_pages': '📑',
			'search_document': '🔍',
			'get_content_controls': '📝',
			'insert_text': '✏️',
			'replace_selection': '🔄',
			'delete_selection': '🗑️',
			'fill_content_control': '📝',
			'add_comment': '💬',
			'insert_at_heading': '📍',
			'find_and_replace': '🔄',
			'regulatory_search': '⚖️',
			'search_reference_documents': '📚'
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

	// =================================================================
	// Task Plan UI Functions
	// =================================================================
	
	// Inject task plan styles if not already present
	(function injectTaskPlanStyles() {
		if (document.getElementById('task-plan-styles')) return;
		var style = document.createElement('style');
		style.id = 'task-plan-styles';
		style.textContent = `
			.task-plan {
				background: linear-gradient(135deg, rgba(155, 92, 255, 0.05) 0%, rgba(100, 60, 200, 0.03) 100%);
				border: 1px solid rgba(155, 92, 255, 0.2);
				border-radius: 12px;
				padding: 16px;
				margin: 12px 0;
				font-family: inherit;
			}
			.task-plan-header {
				margin-bottom: 16px;
			}
			.task-plan-title {
				font-size: 16px;
				font-weight: 600;
				color: #333;
				display: flex;
				align-items: center;
				gap: 8px;
			}
			.plan-icon { font-size: 18px; }
			.task-plan-goal {
				font-size: 13px;
				color: #666;
				margin-top: 6px;
				line-height: 1.4;
			}
			.task-plan-progress {
				margin-top: 12px;
			}
			.progress-bar {
				height: 6px;
				background: rgba(155, 92, 255, 0.1);
				border-radius: 3px;
				overflow: hidden;
			}
			.progress-fill {
				height: 100%;
				background: linear-gradient(90deg, #9B5CFF, #7C3AED);
				border-radius: 3px;
				transition: width 0.3s ease;
			}
			.progress-text {
				font-size: 11px;
				color: #888;
				margin-top: 6px;
				text-align: right;
			}
			.task-list {
				display: flex;
				flex-direction: column;
				gap: 8px;
			}
			.task-item {
				display: flex;
				align-items: flex-start;
				gap: 10px;
				padding: 10px 12px;
				background: white;
				border-radius: 8px;
				border: 1px solid #eee;
				transition: all 0.2s ease;
			}
			.task-item.task-in_progress {
				border-color: rgba(155, 92, 255, 0.4);
				background: rgba(155, 92, 255, 0.03);
			}
			.task-item.task-completed {
				border-color: rgba(34, 197, 94, 0.3);
				background: rgba(34, 197, 94, 0.03);
			}
			.task-item.task-failed {
				border-color: rgba(239, 68, 68, 0.3);
				background: rgba(239, 68, 68, 0.03);
			}
			.task-checkbox {
				flex-shrink: 0;
				width: 20px;
				height: 20px;
				display: flex;
				align-items: center;
				justify-content: center;
			}
			.status-icon {
				font-size: 14px;
				font-weight: bold;
			}
			.status-icon.pending { color: #ccc; }
			.status-icon.in-progress { color: #9B5CFF; }
			.status-icon.completed { color: #22C55E; }
			.status-icon.failed { color: #EF4444; }
			.status-icon.skipped { color: #888; }
			.spinner-small {
				width: 14px;
				height: 14px;
				border: 2px solid rgba(155, 92, 255, 0.2);
				border-top-color: #9B5CFF;
				border-radius: 50%;
				animation: spin 0.8s linear infinite;
			}
			@keyframes spin {
				to { transform: rotate(360deg); }
			}
			.task-content {
				flex: 1;
				min-width: 0;
			}
			.task-title {
				font-size: 13px;
				font-weight: 500;
				color: #333;
				line-height: 1.3;
			}
			.task-section-badge {
				display: inline-block;
				font-size: 10px;
				padding: 2px 6px;
				background: rgba(155, 92, 255, 0.1);
				color: #7C3AED;
				border-radius: 4px;
				margin-top: 4px;
			}
			.task-result {
				font-size: 11px;
				max-width: 150px;
			}
			.task-success {
				color: #22C55E;
			}
			.task-error {
				color: #EF4444;
			}
			.plan-complete-summary {
				margin-top: 12px;
				padding: 12px;
				background: rgba(34, 197, 94, 0.1);
				border-radius: 8px;
				color: #16A34A;
				font-size: 13px;
			}
			.plan-completed .task-plan-header {
				opacity: 0.8;
			}
		`;
		document.head.appendChild(style);
	})();
	
	/**
	 * Render a task plan with checkboxes for each task.
	 * Inspired by Cursor's planning UI.
	 */
	function renderTaskPlan(planData) {
		var tasksHtml = planData.tasks.map(function(task, index) {
			var statusIcon = getTaskStatusIcon(task.status);
			var sectionBadge = task.target_section 
				? '<span class="task-section-badge">' + escapeHtml(task.target_section) + '</span>'
				: '';
			
			return '<div class="task-item" id="task-' + task.id + '" data-task-id="' + task.id + '">' +
				'<div class="task-checkbox">' + statusIcon + '</div>' +
				'<div class="task-content">' +
					'<div class="task-title">' + escapeHtml(task.title) + '</div>' +
					sectionBadge +
				'</div>' +
				'<div class="task-result"></div>' +
			'</div>';
		}).join('');
		
		return '<div class="task-plan" id="task-plan-' + planData.plan_id + '">' +
			'<div class="task-plan-header">' +
				'<div class="task-plan-title">' +
					'<span class="plan-icon">📋</span> ' + escapeHtml(planData.title) +
				'</div>' +
				'<div class="task-plan-goal">' + escapeHtml(planData.goal) + '</div>' +
				'<div class="task-plan-progress">' +
					'<div class="progress-bar">' +
						'<div class="progress-fill" style="width: 0%"></div>' +
					'</div>' +
					'<div class="progress-text">0 / ' + planData.tasks.length + ' tasks</div>' +
				'</div>' +
			'</div>' +
			'<div class="task-list">' + tasksHtml + '</div>' +
		'</div>';
	}
	
	/**
	 * Get the appropriate icon for a task status.
	 */
	function getTaskStatusIcon(status) {
		switch (status) {
			case 'pending':
				return '<span class="status-icon pending">○</span>';
			case 'in_progress':
				return '<span class="status-icon in-progress"><span class="spinner-small"></span></span>';
			case 'completed':
				return '<span class="status-icon completed">✓</span>';
			case 'failed':
				return '<span class="status-icon failed">✗</span>';
			case 'skipped':
				return '<span class="status-icon skipped">–</span>';
			default:
				return '<span class="status-icon">○</span>';
		}
	}
	
	/**
	 * Update the status of a task in the UI.
	 */
	function updateTaskStatus(taskId, status, result) {
		var taskEl = document.getElementById('task-' + taskId);
		if (!taskEl) return;
		
		// Update checkbox
		var checkboxEl = taskEl.querySelector('.task-checkbox');
		if (checkboxEl) {
			checkboxEl.innerHTML = getTaskStatusIcon(status);
		}
		
		// Update class for styling
		taskEl.className = 'task-item task-' + status;
		
		// Show result/error if provided
		if (result) {
			var resultEl = taskEl.querySelector('.task-result');
			if (resultEl) {
				var resultClass = status === 'failed' ? 'task-error' : 'task-success';
				resultEl.innerHTML = '<div class="' + resultClass + '">' + 
					escapeHtml(result.substring(0, 100)) + 
					(result.length > 100 ? '...' : '') + '</div>';
			}
		}
	}
	
	/**
	 * Update the plan's overall progress.
	 */
	function updatePlanProgress(progressData) {
		var planEl = document.querySelector('.task-plan');
		if (!planEl) return;
		
		var percent = progressData.percent || 0;
		var completed = progressData.completed || 0;
		var total = progressData.total || 0;
		
		// Update progress bar
		var fillEl = planEl.querySelector('.progress-fill');
		if (fillEl) {
			fillEl.style.width = percent + '%';
		}
		
		// Update text
		var textEl = planEl.querySelector('.progress-text');
		if (textEl) {
			textEl.textContent = completed + ' / ' + total + ' tasks (' + percent + '%)';
		}
	}
	
	/**
	 * Update the plan's overall status.
	 */
	function updatePlanStatus(planId, status) {
		var planEl = document.getElementById('task-plan-' + planId);
		if (!planEl) return;
		
		planEl.className = 'task-plan plan-' + status;
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
		
		// Remove any other XML-like internal tags (be careful not to remove legitimate HTML-like content)
		// Only target specific internal tags with underscores that are clearly not HTML
		text = text.replace(/<\/?[a-z_]+_[a-z_]+>/gi, '');
		
		// Clean up multiple newlines left behind
		text = text.replace(/\n{3,}/g, '\n\n');
		
		// Trim leading/trailing whitespace
		text = text.trim();
		
		return text;
	}

	// Simple markdown parser with proper link support
	function parseMarkdown(text) {
		if (!text) return '';
		
		// First, strip any internal processing tags from the backend
		text = stripInternalTags(text);
		
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
		
		// Numbered citations: [1], [2] -> clickable badges that scroll to source cards
		html = html.replace(/\[(\d+)\]/g, '<span class="reg-citation-inline" onclick="scrollToSource($1)" title="View source $1">$1</span>');
		
		// Headers
		html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
		html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
		html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');
		
		// Lists
		html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
		html = html.replace(/^(\d+)\. (.+)$/gm, '<li>$2</li>');
		
		// Wrap consecutive <li> in <ul> and remove newlines between list items
		html = html.replace(/(<li>.*<\/li>\n?)+/g, function(match) {
			// Remove newlines between list items to prevent <br> insertion
			var cleanedMatch = match.replace(/\n/g, '');
			return '<ul>' + cleanedMatch + '</ul>';
		});
		
		// Line breaks (only outside of lists now)
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

	// Gather comprehensive document context for backend
	// Based on Anthropic's context engineering best practices:
	// - Provide rich, high-signal context about the working environment
	// - Include position, structure, and formatting awareness
	// - Enable agent to understand spatial constraints
	async function gatherDocumentContext() {
		var context = {
			// Selection & cursor
			selected_text: null,
			has_selection: false,
			cursor_position: null,
			
			// Document structure
			document_outline: null,
			section_map: null,        // Sections with char counts
			total_sections: 0,
			
			// Page/layout awareness (critical for formatting decisions)
			page_info: null,
			document_stats: null,
			
			// Current location context
			current_section: null,
			current_paragraph_style: null,
			surrounding_context: null  // ~500 chars before/after cursor
		};
		
		try {
			// 1. Get selected text (limit to 2000 chars to prevent context overflow)
			var selectedResult = await ToolExecutor.execute('get_selected_text');
			if (selectedResult.success && selectedResult.result) {
				var selectedText = selectedResult.result;
				// Truncate large selections to prevent rate limits
				if (selectedText.length > 2000) {
					context.selected_text = selectedText.substring(0, 1500) + 
						'\n\n[... selection truncated, ' + (selectedText.length - 1500) + ' more chars ...]\n\n' +
						selectedText.substring(selectedText.length - 300);
				} else {
					context.selected_text = selectedText;
				}
				context.has_selection = true;
			}
		} catch (e) {
			console.warn('Could not get selected text:', e);
		}
		
		try {
			// 2. Get page info (position awareness)
			var pageResult = await ToolExecutor.execute('get_page_info');
			if (pageResult && !pageResult.error) {
				context.page_info = {
					current_page: pageResult.page_number,
					total_pages: pageResult.page_count,
					position_in_doc: pageResult.page_count > 0 
						? Math.round((pageResult.page_number / pageResult.page_count) * 100) + '%'
						: 'unknown'
				};
			}
		} catch (e) {
			console.warn('Could not get page info:', e);
		}
		
		try {
			// 3. Get document outline/structure
			var outlineResult = await ToolExecutor.execute('get_document_outline');
			if (outlineResult && Array.isArray(outlineResult)) {
				context.document_outline = outlineResult;
				context.total_sections = outlineResult.length;
				
				// Build section map with abbreviated info (limit to first 50 sections)
				var sectionsToInclude = outlineResult.slice(0, 50);
				context.section_map = sectionsToInclude.map(function(h, idx) {
					return {
						index: idx,
						level: h.level || 1,
						heading: (h.text || '').substring(0, 60),
						char_count: h.char_count || null
					};
				});
				
				// Note if there are more sections
				if (outlineResult.length > 50) {
					context.section_map.push({
						index: -1,
						note: '... and ' + (outlineResult.length - 50) + ' more sections (use get_document_map for full list)'
					});
				}
			}
		} catch (e) {
			console.warn('Could not get document outline:', e);
		}
		
		try {
			// 4. Get current paragraph context (style awareness)
			var paraResult = await ToolExecutor.execute('get_current_paragraph');
			if (paraResult && !paraResult.error) {
				context.current_paragraph_style = paraResult.style || 'Normal';
				
				// Enhanced style context for format-aware editing
				context.current_context = {
					content_type: paraResult.content_type || 'paragraph',
					is_heading: paraResult.is_heading || false,
					heading_level: paraResult.heading_level || 0,
					is_empty: paraResult.is_empty || false,
					char_count: paraResult.char_count || 0,
					formatting: paraResult.formatting || {}
				};
				
				// Get surrounding context - helps agent understand where cursor is
				var paraText = paraResult.text || '';
				if (paraText.length > 0) {
					context.surrounding_context = paraText.substring(0, 500) + 
						(paraText.length > 500 ? '...' : '');
				}
				
				// Determine which section cursor is in based on document outline
				if (context.document_outline && context.document_outline.length > 0) {
					var currentHeading = null;
					var prevHeading = null;
					
					// If cursor is in a heading, that's the current section
					if (paraResult.is_heading) {
						var cursorHeadingText = paraText.trim().substring(0, 60);
						for (var i = 0; i < context.document_outline.length; i++) {
							var h = context.document_outline[i];
							if (h.text && h.text.substring(0, 60) === cursorHeadingText) {
								currentHeading = h;
								if (i > 0) prevHeading = context.document_outline[i - 1];
								break;
							}
						}
					}
					
					context.current_section = {
						in_heading: paraResult.is_heading,
						heading_text: currentHeading ? currentHeading.text : null,
						heading_level: currentHeading ? currentHeading.level : (prevHeading ? prevHeading.level : null),
						parent_section: prevHeading ? prevHeading.text : null,
						position_hint: paraResult.is_heading ? 'at_heading' : 
							(paraResult.is_empty ? 'empty_paragraph' : 'in_content')
					};
				}
			}
		} catch (e) {
			console.warn('Could not get current paragraph:', e);
		}
		
		try {
			// 5. Get document stats via snapshot (efficient single call)
			var snapshotResult = await ToolExecutor.execute('get_document_snapshot', {
				include_outline: false,
				include_markdown: false,
				include_headers_footers: false
			});
			if (snapshotResult && !snapshotResult.error) {
				context.document_stats = {
					element_count: snapshotResult.element_count,
					page_count: snapshotResult.page_count
				};
				context.cursor_position = {
					page: snapshotResult.page_number,
					paragraph_preview: (snapshotResult.current_paragraph || '').substring(0, 100)
				};
			}
		} catch (e) {
			console.warn('Could not get document snapshot:', e);
		}
		
		// 6. Generate position summary for the agent (human-readable context)
		context.position_summary = generatePositionSummary(context);
		
		return context;
	}
	
	// Generate a human-readable summary of cursor position for the agent
	function generatePositionSummary(context) {
		var parts = [];
		
		// Page position
		if (context.page_info) {
			parts.push('Page ' + context.page_info.current_page + ' of ' + context.page_info.total_pages);
		}
		
		// Current context type
		if (context.current_context) {
			var ctx = context.current_context;
			if (ctx.is_heading) {
				parts.push('Cursor is in a Heading ' + ctx.heading_level);
			} else if (ctx.content_type === 'title') {
				parts.push('Cursor is in the document Title');
			} else if (ctx.content_type === 'list_item') {
				parts.push('Cursor is in a list item');
			} else if (ctx.is_empty) {
				parts.push('Cursor is in an empty paragraph');
			} else {
				parts.push('Cursor is in a ' + (context.current_paragraph_style || 'Normal') + ' paragraph');
			}
		}
		
		// Current section
		if (context.current_section && context.current_section.parent_section) {
			parts.push('Under section: "' + context.current_section.parent_section + '"');
		}
		
		// Selection status
		if (context.has_selection && context.selected_text) {
			var selLen = context.selected_text.length;
			parts.push('Has selection (' + selLen + ' chars)');
		}
		
		// Formatting hints
		if (context.current_context && context.current_context.formatting) {
			var fmt = context.current_context.formatting;
			if (fmt.alignment && fmt.alignment !== 'left') {
				parts.push('Alignment: ' + fmt.alignment);
			}
			if (fmt.indent_left > 0) {
				parts.push('Indented');
			}
		}
		
		return parts.length > 0 ? parts.join(' | ') : 'Document position unknown';
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
				// Separate tool_calls from other items for grouping
				var toolCalls = [];
				var otherItems = [];
				
				msg.metadata.items.forEach(function(item) {
					if (item.type === 'tool_call' || item.type === 'tool_result') {
						toolCalls.push(item);
					} else {
						otherItems.push(item);
					}
				});
				
				// Render non-tool items first (thinking blocks, etc.)
				otherItems.forEach(function(item) {
					html += renderMetadataItem(item);
				});
				
				// Render tool calls in a collapsible container
				if (toolCalls.length > 0) {
					html += renderToolsUsedContainer(toolCalls, true);
				}
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

	// Render the collapsible "Tools Used" container
	function renderToolsUsedContainer(toolCalls, collapsed) {
		if (!toolCalls || toolCalls.length === 0) return '';
		
		var hasRunning = toolCalls.some(function(t) { return t.status === 'running'; });
		var expandedClass = collapsed ? '' : ' expanded';
		var hasRunningAttr = hasRunning ? ' data-has-running="true"' : '';
		
		var html = '<div class="tools-used-container' + expandedClass + '"' + hasRunningAttr + '>';
		
		// Header bar
		html += '<div class="tools-used-header" role="button" tabindex="0" aria-expanded="' + (!collapsed) + '">';
		html += '<span class="tools-used-icon">';
		html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">';
		html += '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>';
		html += '</svg>';
		html += '</span>';
		html += '<span class="tools-used-label">Tools Used</span>';
		html += '<span class="tools-used-count">' + toolCalls.length + '</span>';
		
		// Show spinner if any tool is running
		if (hasRunning) {
			html += '<span class="tools-used-status"><span class="spinner" aria-hidden="true"></span></span>';
		}
		
		html += '<span class="tools-used-toggle" aria-hidden="true">›</span>';
		html += '</div>';
		
		// List of tool calls (hidden by default)
		html += '<div class="tools-used-list">';
		toolCalls.forEach(function(item) {
			html += renderMetadataItem(item);
		});
		html += '</div>';
		
		html += '</div>';
		
		return html;
	}

	// Get or create the tools-used container in a message during streaming
	function getOrCreateToolsContainer(msgContainer) {
		var container = msgContainer.querySelector('.tools-used-container');
		if (!container) {
			// Create container expanded during generation, will collapse on 'done'
			var html = '<div class="tools-used-container expanded" data-generating="true">';
			html += '<div class="tools-used-header" role="button" tabindex="0" aria-expanded="true">';
			html += '<span class="tools-used-icon">';
			html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">';
			html += '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>';
			html += '</svg>';
			html += '</span>';
			html += '<span class="tools-used-label">Tools Used</span>';
			html += '<span class="tools-used-count">0</span>';
			html += '<span class="tools-used-status"><span class="spinner" aria-hidden="true"></span></span>';
			html += '<span class="tools-used-toggle" aria-hidden="true">›</span>';
			html += '</div>';
			html += '<div class="tools-used-list"></div>';
			html += '</div>';
			msgContainer.insertAdjacentHTML('beforeend', html);
			container = msgContainer.querySelector('.tools-used-container');
		}
		return container;
	}

	// Add a tool call to the tools container
	function addToolToContainer(msgContainer, toolItem) {
		var container = getOrCreateToolsContainer(msgContainer);
		var list = container.querySelector('.tools-used-list');
		var countEl = container.querySelector('.tools-used-count');
		var statusEl = container.querySelector('.tools-used-status');
		
		// Add the tool to the list
		list.insertAdjacentHTML('beforeend', renderMetadataItem(toolItem));
		
		// Get the newly added tool element
		var toolCalls = list.querySelectorAll('.tool-call');
		var newToolEl = toolCalls[toolCalls.length - 1];
		
		// Expand the tool during generation so context is visible
		newToolEl.classList.add('expanded');
		var toolHeader = newToolEl.querySelector('.tool-call-header');
		if (toolHeader) {
			toolHeader.setAttribute('aria-expanded', 'true');
		}
		
		// Capture any intermediate content and associate it with this tool
		captureIntermediateContent(msgContainer, newToolEl);
		
		// Update count
		var toolCount = toolCalls.length;
		countEl.textContent = toolCount;
		
		// Show spinner if this tool is running
		if (toolItem.status === 'running') {
			statusEl.innerHTML = '<span class="spinner" aria-hidden="true"></span>';
			container.setAttribute('data-has-running', 'true');
		}
	}

	// Update the tools container status (check for running tools)
	function updateToolsContainerStatus(msgContainer) {
		var container = msgContainer.querySelector('.tools-used-container');
		if (!container) return;
		
		var runningTools = container.querySelectorAll('.tool-call[data-status="running"]');
		var isGenerating = container.getAttribute('data-generating') === 'true';
		var statusEl = container.querySelector('.tools-used-status');
		
		if (runningTools.length > 0) {
			statusEl.innerHTML = '<span class="spinner" aria-hidden="true"></span>';
			container.setAttribute('data-has-running', 'true');
		} else if (isGenerating) {
			// Keep spinner while agent is still generating response
			statusEl.innerHTML = '<span class="spinner" aria-hidden="true"></span>';
			container.removeAttribute('data-has-running');
		} else {
			statusEl.innerHTML = '';
			container.removeAttribute('data-has-running');
		}
	}

	// Finalize the tools container when generation is complete
	function finalizeToolsContainer(msgContainer) {
		var container = msgContainer.querySelector('.tools-used-container');
		if (!container) return;
		
		// Remove generating state
		container.removeAttribute('data-generating');
		
		// Collapse the tools-used container
		container.classList.remove('expanded');
		var containerHeader = container.querySelector('.tools-used-header');
		if (containerHeader) {
			containerHeader.setAttribute('aria-expanded', 'false');
		}
		
		// Collapse all individual tool chips
		var toolCalls = container.querySelectorAll('.tool-call');
		toolCalls.forEach(function(toolCall) {
			toolCall.classList.remove('expanded');
			var toolHeader = toolCall.querySelector('.tool-call-header');
			if (toolHeader) {
				toolHeader.setAttribute('aria-expanded', 'false');
			}
		});
		
		// Update status (will remove spinner since no longer generating)
		updateToolsContainerStatus(msgContainer);
	}

	// Capture intermediate content and associate it with a tool
	function captureIntermediateContent(msgContainer, toolCallEl) {
		var intermediateEl = msgContainer.querySelector('.intermediate-content');
		if (!intermediateEl) return;
		
		var text = intermediateEl.getAttribute('data-raw') || '';
		if (!text.trim()) {
			// No content to capture, just remove the element
			intermediateEl.remove();
			return;
		}
		
		// Store the context text in the tool's expandable section
		var bodyEl = toolCallEl.querySelector('.tool-call-body');
		if (bodyEl) {
			var contextHtml = '<div class="tool-context-section">';
			contextHtml += '<div class="tool-section-label">Context</div>';
			contextHtml += '<div class="tool-context-text">' + parseMarkdown(text) + '</div>';
			contextHtml += '</div>';
			bodyEl.insertAdjacentHTML('afterbegin', contextHtml);
		}
		
		// Clear and remove the intermediate content element
		intermediateEl.remove();
	}

	// Get or create intermediate content element for displaying text during tool execution
	function getOrCreateIntermediateContent(msgContainer) {
		var el = msgContainer.querySelector('.intermediate-content');
		if (!el) {
			msgContainer.insertAdjacentHTML('beforeend', '<div class="intermediate-content streaming"></div>');
			el = msgContainer.querySelector('.intermediate-content');
		}
		return el;
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
	
	// Render sources section for agent response (displays source chips below the text)
	function renderSourcesSection(sources) {
		if (!sources || sources.length === 0) return '';
		
		var html = '<div class="agent-sources-section">';
		
		// Compact chips row - one chip per source
		html += '<div class="agent-sources-chips">';
		for (var i = 0; i < sources.length; i++) {
			html += renderSourceChip(sources[i], i);
		}
		html += '</div>';
		
		// Expandable detail panel (hidden by default, shown when chip is clicked)
		html += '<div class="agent-source-detail" id="agent-source-detail"></div>';
		
		html += '</div>';
		
		return html;
	}
	
	// Render individual source chip (small oval pill)
	function renderSourceChip(src, index) {
		var sourceType = (src.source_type || 'doc').toLowerCase();
		var title = src.title || src.code || 'Source ' + (index + 1);
		// Truncate title for chip display
		var shortTitle = title.length > 30 ? title.substring(0, 27) + '...' : title;
		var uniqueId = 'source-chip-' + index;
		
		var html = '<button class="source-chip" data-source-index="' + index + '" id="' + uniqueId + '" onclick="toggleSourceDetail(' + index + ')" title="' + escapeHtml(title) + '">';
		html += '<span class="source-chip-badge">[' + (index + 1) + ']</span>';
		html += '<span class="source-chip-title">' + escapeHtml(shortTitle) + '</span>';
		if (src.code && src.code !== title) {
			html += '<span class="source-chip-code">' + escapeHtml(src.code) + '</span>';
		}
		html += '<span class="source-chip-type source-chip-type--' + sourceType + '">' + sourceType.toUpperCase() + '</span>';
		html += '</button>';
		
		return html;
	}
	
	// Store sources globally for detail panel access
	var _currentSources = [];
	
	// Toggle source detail panel
	window.toggleSourceDetail = function(index) {
		var detailPanel = document.getElementById('agent-source-detail');
		var chips = document.querySelectorAll('.source-chip');
		
		if (!detailPanel || !_currentSources[index]) return;
		
		var src = _currentSources[index];
		var isAlreadyShowing = detailPanel.getAttribute('data-showing-index') === String(index);
		
		// Remove active state from all chips
		chips.forEach(function(chip) {
			chip.classList.remove('source-chip--active');
		});
		
		if (isAlreadyShowing) {
			// Hide panel
			detailPanel.innerHTML = '';
			detailPanel.classList.remove('visible');
			detailPanel.removeAttribute('data-showing-index');
		} else {
			// Show panel with source details
			var chip = document.querySelector('.source-chip[data-source-index="' + index + '"]');
			if (chip) chip.classList.add('source-chip--active');
			
			var sourceType = (src.source_type || 'doc').toLowerCase();
			var html = '<div class="source-detail-content">';
			html += '<div class="source-detail-header">';
			html += '<span class="source-detail-badge">[' + (index + 1) + ']</span>';
			html += '<span class="source-detail-title">' + escapeHtml(src.title || 'Source ' + (index + 1)) + '</span>';
			if (src.code) {
				html += '<span class="source-detail-code">' + escapeHtml(src.code) + '</span>';
			}
			html += '<span class="source-detail-type source-chip-type--' + sourceType + '">' + sourceType.toUpperCase() + '</span>';
			html += '<button class="source-detail-close" onclick="toggleSourceDetail(' + index + ')">&times;</button>';
			html += '</div>';
			
			if (src.header_path) {
				html += '<div class="source-detail-path">' + escapeHtml(src.header_path) + '</div>';
			}
			
			html += '<div class="source-detail-text">';
			html += '<p>' + escapeHtml(src.snippet || src.full_text || src.context || 'No content available.') + '</p>';
			html += '</div>';
			
			if (src.url) {
				html += '<a href="' + escapeHtml(src.url) + '" target="_blank" rel="noopener noreferrer" class="source-detail-link">';
				html += 'View Original';
				html += '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>';
				html += '</a>';
			}
			
			html += '</div>';
			
			detailPanel.innerHTML = html;
			detailPanel.classList.add('visible');
			detailPanel.setAttribute('data-showing-index', index);
			
			// Scroll detail into view
			detailPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
		}
	};
	
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
		
		// First, strip any internal processing tags from the backend
		text = stripInternalTags(text);
		
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

	// ============================================
	// PROGRESS CARDS RENDERING
	// ============================================
	
	// Initialize progress tracking for a new message
	function initProgressTracking(msgContainer) {
		var containerId = 'progress-' + Date.now();
		state.progressContainerId = containerId;
		state.currentStage = 'understanding';
		state.completedStages = [];
		
		// Insert progress cards container at the start of message
		var progressHtml = renderProgressCards(containerId);
		msgContainer.insertAdjacentHTML('afterbegin', progressHtml);
		
		// Start with understanding stage active
		updateProgressStage('understanding');
	}
	
	// Render the progress cards HTML
	function renderProgressCards(containerId) {
		var html = '<div class="agent-progress-container" id="' + containerId + '">';
		html += '<div class="agent-progress-header">';
		html += '<span class="agent-progress-title">Processing your request</span>';
		html += '<button class="agent-progress-toggle" onclick="toggleProgressCards(\'' + containerId + '\')" title="Collapse">';
		html += '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="18 15 12 9 6 15"/></svg>';
		html += '</button>';
		html += '</div>';
		html += '<div class="agent-progress-cards">';
		
		for (var i = 0; i < AGENT_STAGES.length; i++) {
			var stage = AGENT_STAGES[i];
			var statusClass = i === 0 ? 'active' : 'pending';
			
			html += '<div class="progress-card" data-stage="' + stage.id + '" data-status="' + statusClass + '">';
			html += '<div class="progress-card-icon">' + stage.icon + '</div>';
			html += '<div class="progress-card-content">';
			html += '<h4 class="progress-card-label">' + stage.label + '</h4>';
			html += '<p class="progress-card-description">' + stage.description + '</p>';
			html += '</div>';
			html += '<div class="progress-card-status">';
			html += '<div class="progress-spinner"></div>';
			html += '<svg class="progress-check" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>';
			html += '</div>';
			html += '</div>';
		}
		
		html += '</div></div>';
		return html;
	}
	
	// Update progress to a specific stage
	function updateProgressStage(stageId) {
		if (!state.progressContainerId) return;
		
		var container = document.getElementById(state.progressContainerId);
		if (!container) return;
		
		var stageIndex = -1;
		for (var i = 0; i < AGENT_STAGES.length; i++) {
			if (AGENT_STAGES[i].id === stageId) {
				stageIndex = i;
				break;
			}
		}
		
		if (stageIndex === -1) return;
		
		// Mark all previous stages as complete
		for (var j = 0; j < stageIndex; j++) {
			var prevStageId = AGENT_STAGES[j].id;
			if (state.completedStages.indexOf(prevStageId) === -1) {
				state.completedStages.push(prevStageId);
			}
			var prevCard = container.querySelector('[data-stage="' + prevStageId + '"]');
			if (prevCard) {
				prevCard.setAttribute('data-status', 'complete');
			}
		}
		
		// Set current stage as active
		state.currentStage = stageId;
		var currentCard = container.querySelector('[data-stage="' + stageId + '"]');
		if (currentCard) {
			currentCard.setAttribute('data-status', 'active');
		}
		
		// Reset all FUTURE stages back to pending (important for when stages go backwards)
		for (var k = stageIndex + 1; k < AGENT_STAGES.length; k++) {
			var futureStageId = AGENT_STAGES[k].id;
			// Remove from completed stages if it was there
			var completedIndex = state.completedStages.indexOf(futureStageId);
			if (completedIndex !== -1) {
				state.completedStages.splice(completedIndex, 1);
			}
			var futureCard = container.querySelector('[data-stage="' + futureStageId + '"]');
			if (futureCard) {
				futureCard.setAttribute('data-status', 'pending');
			}
		}
		
		// Update header text based on stage
		var headerTitle = container.querySelector('.agent-progress-title');
		if (headerTitle) {
			var stageLabels = {
				'understanding': 'Understanding your query...',
				'searching': 'Searching knowledge base...',
				'sources': 'Processing sources...',
				'synthesizing': 'Generating response...',
				'complete': 'Response complete'
			};
			headerTitle.textContent = stageLabels[stageId] || 'Processing...';
		}
	}
	
	// Complete all progress stages
	function completeProgressStages() {
		updateProgressStage('complete');
		
		if (!state.progressContainerId) return;
		
		var container = document.getElementById(state.progressContainerId);
		if (!container) return;
		
		// Mark the complete stage as complete too
		var completeCard = container.querySelector('[data-stage="complete"]');
		if (completeCard) {
			completeCard.setAttribute('data-status', 'complete');
		}
		
		// Add completed class to container for styling
		container.classList.add('completed');
		
		// Auto-collapse after a short delay
		setTimeout(function() {
			if (container && !container.classList.contains('manually-expanded')) {
				container.classList.add('collapsed');
			}
		}, 2000);
	}
	
	// Reset progress tracking
	function resetProgressTracking() {
		state.currentStage = null;
		state.completedStages = [];
		state.progressContainerId = null;
	}
	
	// Global function to toggle progress cards visibility
	window.toggleProgressCards = function(containerId) {
		var container = document.getElementById(containerId);
		if (container) {
			container.classList.toggle('collapsed');
			container.classList.toggle('manually-expanded');
		}
	};

	// Update stage based on SSE event type and data
	function updateStageFromSSEEvent(eventType, data) {
		switch (eventType) {
			case 'tool_call':
				if (data.name === 'regulatory_search' || data.name === 'search_reference_documents') {
					if (data.status === 'running') {
						updateProgressStage('searching');
					} else if (data.status === 'success') {
						// Check if we have sources
						if (data.result && data.result.sources && data.result.sources.length > 0) {
							updateProgressStage('sources');
						} else {
							updateProgressStage('synthesizing');
						}
					}
				}
				break;
				
			case 'tool_result_request':
				// A frontend tool is being requested
				updateProgressStage('searching');
				break;
				
			case 'content':
				// Content is streaming - we're in synthesizing phase
				if (state.currentStage !== 'synthesizing' && state.currentStage !== 'complete') {
					updateProgressStage('synthesizing');
				}
				break;
				
			case 'sources':
				// Direct sources event
				updateProgressStage('sources');
				break;
				
			case 'step':
				// Step events from regulatory search API
				if (data.step === 'analyze' && data.status === 'active') {
					updateProgressStage('understanding');
				} else if (data.step === 'search' && data.status === 'active') {
					updateProgressStage('searching');
				} else if (data.step === 'synthesize' && data.status === 'active') {
					updateProgressStage('synthesizing');
				}
				break;
				
			case 'done':
				completeProgressStages();
				break;
				
			case 'error':
				// On error, mark current stage as error (could add error styling)
				if (state.progressContainerId) {
					var container = document.getElementById(state.progressContainerId);
					if (container) {
						container.classList.add('has-error');
					}
				}
				break;
			
			case 'status':
				// Status events (rate limiting, retries) - no progress change needed
				break;
			
			// Task plan events - no stage tracking needed
			case 'plan_created':
			case 'plan_start':
			case 'task_start':
			case 'task_complete':
			case 'plan_progress':
			case 'plan_complete':
				// These have their own UI tracking
				break;
		}
	}
	
	// Global function to toggle source card expansion
	window.toggleSourceCard = function(sourceId) {
		var card = document.querySelector('[data-source-id="' + sourceId + '"]');
		if (card) {
			var isExpanded = card.getAttribute('data-expanded') === 'true';
			card.setAttribute('data-expanded', !isExpanded);
		}
	};
	
	// Global function to scroll to source - works with both chip UI and card UI
	window.scrollToSource = function(sourceNum) {
		var index = sourceNum - 1; // Convert to 0-based index
		
		// First, try the new chip UI
		var chip = document.querySelector('.source-chip[data-source-index="' + index + '"]');
		if (chip) {
			// Scroll to chip
			chip.scrollIntoView({ behavior: 'smooth', block: 'center' });
			// Highlight chip briefly
			chip.classList.add('source-chip--highlight');
			setTimeout(function() {
				chip.classList.remove('source-chip--highlight');
			}, 1500);
			// Expand the detail panel
			if (typeof toggleSourceDetail === 'function') {
				toggleSourceDetail(index);
			}
			return;
		}
		
		// Fallback to old card UI
		var cards = document.querySelectorAll('.reg-source-card');
		var targetCard = cards[index];
		if (targetCard) {
			targetCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
			targetCard.setAttribute('data-expanded', 'true');
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
		
		// Reset sources for this new message
		state.currentMessageSources = [];
		
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
			
			// Include agent info for edit mode
			if (state.mode === 'agent') {
				var agentInfo = getSelectedAgentInfo();
				if (agentInfo) {
					requestData.agent = agentInfo;
				}
			}
			
			// Create assistant message container
			appendMessageElement('<div class="message assistant"></div>');
			var msgContainer = elements.messages.querySelector('.message.assistant:last-child');
			
			// Initialize progress tracking for this message
			initProgressTracking(msgContainer);
			
			// Track response data for saving
			var items = [];
			var textContent = '';
			
			// Make SSE request
			var requestUrl = Config.BACKEND_URL + '/api/copilot/chat';
			console.log('[Copilot] Making request to:', requestUrl);
			console.log('[Copilot] Request data:', JSON.stringify(requestData, null, 2));
			
			var response = await fetch(requestUrl, {
				method: 'POST',
				headers: {
					'Content-Type': 'application/json',
					'Accept': 'text/event-stream'
				},
				body: JSON.stringify(requestData),
				signal: state.abortController.signal
			});
			
			console.log('[Copilot] Response status:', response.status, response.statusText);
			
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
			// Finalize tools container on error/abort
			if (msgContainer) {
				finalizeToolsContainer(msgContainer);
			}
			if (error.name === 'AbortError') {
				console.log('[Copilot] Request aborted by user');
			} else {
				console.error('[Copilot] Backend error:', error);
				console.error('[Copilot] Error name:', error.name);
				console.error('[Copilot] Error message:', error.message);
				console.error('[Copilot] Attempted URL was:', Config.BACKEND_URL + '/api/copilot/chat');
				addMessage('error', 'Connection error: ' + error.message);
				appendMessageElement('<div class="error-message">Connection error: ' + escapeHtml(error.message) + '</div>');
			}
		} finally {
			setGenerating(false);
			state.abortController = null;
			// Reset progress tracking for next message
			resetProgressTracking();
		}
	}

	async function handleSSEEvent(eventType, data, msgContainer, items, onContent) {
		// Update progress stage based on event
		updateStageFromSSEEvent(eventType, data);
		
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
				addToolToContainer(msgContainer, toolItem);
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
			
			// For search tools, skip rendering the tool indicator - sources will be shown as chips at the end
			var isSearchTool = (data.name === 'regulatory_search' || data.name === 'search_reference_documents');
			
			if (!isSearchTool) {
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
					addToolToContainer(msgContainer, displayToolItem);
				}
			}
			
			// Collect sources from search tools for numbered citations
			if (data.status === 'success' && data.result) {
				var toolSources = [];
				if (data.name === 'regulatory_search' && data.result.sources) {
					console.log('[Copilot] Collecting sources from regulatory_search:', data.result.sources.length, 'sources, start_index:', data.result.source_start_index);
					toolSources = data.result.sources.map(function(s) {
						return {
							title: s.title,
							code: s.code,
							snippet: s.snippet,
							full_text: s.full_text,
							source_type: s.source_type || 'regulatory',
							url: s.url,
							page_numbers: s.page_numbers,
							header_path: s.header_path
						};
					});
				} else if (data.name === 'search_reference_documents' && data.result.results) {
					console.log('[Copilot] Collecting sources from search_reference_documents:', data.result.results.length, 'sources, start_index:', data.result.source_start_index);
					toolSources = data.result.results.map(function(r) {
						return {
							title: r.source,
							snippet: r.text,
							full_text: r.text,
							context: r.context,
							source_type: 'reference',
							relevance_score: r.relevance_score
						};
					});
				}
				// Append to current message sources (maintaining order for citation numbers)
				if (toolSources.length > 0) {
					state.currentMessageSources = state.currentMessageSources.concat(toolSources);
					console.log('[Copilot] Total sources collected so far:', state.currentMessageSources.length);
				}
			}
			
			scrollToBottom();
			break;
			
			case 'content':
				// Streaming text content
				if (data.delta) {
					onContent(data.delta);
					
					// During generation, ALL content goes to intermediate-content
					// It will be captured when tools start, or converted to message-content on 'done'
					var intermediateEl = getOrCreateIntermediateContent(msgContainer);
					var currentText = intermediateEl.getAttribute('data-raw') || '';
					currentText += data.delta;
					intermediateEl.setAttribute('data-raw', currentText);
					intermediateEl.innerHTML = parseMarkdown(currentText);
					
					scrollToBottom();
					
					// Yield to browser to allow repaint (enables visible streaming)
					await new Promise(function(resolve) { requestAnimationFrame(resolve); });
				}
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
			
			case 'status':
				// Status messages (rate limiting, retries, etc)
				console.log('Status:', data.message);
				// Show as a subtle status indicator
				var statusHtml = '<div class="status-message" style="color: #666; font-style: italic; padding: 4px 0; font-size: 12px;">' + 
					escapeHtml(data.message || '') + '</div>';
				msgContainer.insertAdjacentHTML('beforeend', statusHtml);
				scrollToBottom();
				// Auto-remove after 10 seconds
				setTimeout(function() {
					var statusEl = msgContainer.querySelector('.status-message:last-of-type');
					if (statusEl) statusEl.remove();
				}, 10000);
				break;
			
			// =================================================================
			// Task Planning Events (for complex multi-step tasks like CSR)
			// =================================================================
			
			case 'plan_created':
				// A task plan has been created - show the plan UI
				console.log('[TaskPlan] Plan created:', data);
				var planHtml = renderTaskPlan(data);
				msgContainer.insertAdjacentHTML('beforeend', planHtml);
				state.activePlanId = data.plan_id;
				scrollToBottom();
				break;
			
			case 'plan_start':
				// Plan execution starting
				console.log('[TaskPlan] Plan starting:', data);
				updatePlanStatus(data.plan_id, 'in_progress');
				break;
			
			case 'task_start':
				// Individual task starting
				console.log('[TaskPlan] Task starting:', data.task_id, data.title);
				updateTaskStatus(data.task_id, 'in_progress');
				break;
			
			case 'task_complete':
				// Individual task completed
				console.log('[TaskPlan] Task complete:', data.task_id, data.status);
				updateTaskStatus(data.task_id, data.status, data.result || data.error);
				break;
			
			case 'plan_progress':
				// Progress update
				console.log('[TaskPlan] Progress:', data);
				updatePlanProgress(data);
				break;
			
			case 'plan_complete':
				// Plan finished
				console.log('[TaskPlan] Plan complete:', data);
				updatePlanStatus(data.plan_id, 'completed');
				// Add completion summary
				var summaryHtml = '<div class="plan-complete-summary">' +
					'<strong>✓ Plan Complete</strong><br>' +
					'Completed ' + data.progress.completed + ' of ' + data.progress.total + ' tasks' +
					(data.progress.failed > 0 ? ' (' + data.progress.failed + ' failed)' : '') +
					'</div>';
				var planEl = document.getElementById('task-plan-' + data.plan_id);
				if (planEl) {
					planEl.insertAdjacentHTML('beforeend', summaryHtml);
				}
				break;
			
			case 'done':
				// Convert intermediate-content to message-content (the final answer)
				var intermediateEl = msgContainer.querySelector('.intermediate-content');
				if (intermediateEl) {
					// This is the final answer - convert to message-content
					intermediateEl.classList.remove('intermediate-content', 'streaming');
					intermediateEl.classList.add('message-content');
				}
				
				// Remove streaming class from any remaining elements
				var streamingEl = msgContainer.querySelector('.streaming');
				if (streamingEl) {
					streamingEl.classList.remove('streaming');
				}
				
				// Finalize tools container (stop spinner)
				finalizeToolsContainer(msgContainer);
				
				// Render source chips if we collected any from tool calls
				console.log('[Copilot] Done event - collected sources:', state.currentMessageSources.length);
				if (state.currentMessageSources.length > 0) {
					console.log('[Copilot] Rendering sources section with', state.currentMessageSources.length, 'sources');
					// Store sources globally for detail panel access
					_currentSources = state.currentMessageSources.slice();
					var sourcesHtml = renderSourcesSection(state.currentMessageSources);
					msgContainer.insertAdjacentHTML('beforeend', sourcesHtml);
				// Reset for next message
				state.currentMessageSources = [];
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
		
		// Update tools container status (remove spinner if no more running tools)
		var msgContainer = toolCallEl.closest('.message');
		if (msgContainer) {
			updateToolsContainerStatus(msgContainer);
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
				addToolToContainer(msgContainer, toolItem);
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
					
					// Update tools container status
					updateToolsContainerStatus(msgContainer);
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
		
		// Finalize tools container (stop spinner)
		finalizeToolsContainer(msgContainer);
		
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
		
		// Finalize any active tools containers (stop spinners)
		var generatingContainers = elements.messages.querySelectorAll('.tools-used-container[data-generating="true"]');
		generatingContainers.forEach(function(container) {
			var msgContainer = container.closest('.message');
			if (msgContainer) {
				finalizeToolsContainer(msgContainer);
			}
		});
		
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
		
		elements.modeIndicator.textContent = mode === 'ask' ? 'Ask mode' : 'Edit mode';
		
		// Show/hide agent selector based on mode
		if (elements.agentSelector) {
			elements.agentSelector.style.display = mode === 'agent' ? 'flex' : 'none';
		}
		
		// Set placeholder based on mode and selected agent
		if (mode === 'ask') {
			elements.chatInput.placeholder = 'Ask anything about your document...';
		} else {
			var agent = AGENTS[state.selectedAgent];
			elements.chatInput.placeholder = agent ? agent.placeholder : 'Tell me what to do with the document...';
		}
	}
	
	// ============================================
	// AGENT SELECTION
	// ============================================
	function selectAgent(agentId) {
		if (!AGENTS[agentId]) return;
		
		state.selectedAgent = agentId;
		var agent = AGENTS[agentId];
		
		// Update UI
		if (elements.selectedAgentName) {
			elements.selectedAgentName.textContent = agent.name;
		}
		if (elements.selectedAgentIcon) {
			elements.selectedAgentIcon.innerHTML = agent.icon;
		}
		
		// Update dropdown items
		if (elements.agentDropdown) {
			var items = elements.agentDropdown.querySelectorAll('.agent-option');
			items.forEach(function(item) {
				item.classList.toggle('active', item.dataset.agentId === agentId);
			});
		}
		
		// Update placeholder
		elements.chatInput.placeholder = agent.placeholder;
		
		// Close dropdown
		closeAgentDropdown();
		
		console.log('[Copilot] Selected agent:', agent.name);
	}
	
	function toggleAgentDropdown() {
		if (elements.agentDropdown) {
			var isVisible = elements.agentDropdown.classList.contains('visible');
			if (isVisible) {
				closeAgentDropdown();
			} else {
				elements.agentDropdown.classList.add('visible');
				// Add click outside listener
				setTimeout(function() {
					document.addEventListener('click', handleAgentDropdownOutsideClick);
				}, 0);
			}
		}
	}
	
	function closeAgentDropdown() {
		if (elements.agentDropdown) {
			elements.agentDropdown.classList.remove('visible');
			document.removeEventListener('click', handleAgentDropdownOutsideClick);
		}
	}
	
	function handleAgentDropdownOutsideClick(e) {
		if (elements.agentSelector && !elements.agentSelector.contains(e.target)) {
			closeAgentDropdown();
		}
	}
	
	function renderAgentDropdown() {
		if (!elements.agentDropdown) return;
		
		var html = '';
		Object.keys(AGENTS).forEach(function(agentId) {
			var agent = AGENTS[agentId];
			var isActive = state.selectedAgent === agentId;
			
			html += '<div class="agent-option' + (isActive ? ' active' : '') + '" data-agent-id="' + agentId + '">';
			html += '<div class="agent-option-icon">' + agent.icon + '</div>';
			html += '<div class="agent-option-content">';
			html += '<div class="agent-option-name">' + agent.name + '</div>';
			html += '<div class="agent-option-desc">' + agent.description + '</div>';
			html += '</div>';
			if (isActive) {
				html += '<svg class="agent-option-check" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>';
			}
			html += '</div>';
		});
		
		elements.agentDropdown.innerHTML = html;
		
		// Add click handlers
		var options = elements.agentDropdown.querySelectorAll('.agent-option');
		options.forEach(function(option) {
			option.addEventListener('click', function(e) {
				e.stopPropagation();
				selectAgent(option.dataset.agentId);
			});
		});
	}
	
	function getSelectedAgentInfo() {
		var agent = AGENTS[state.selectedAgent];
		if (!agent) return null;
		
		return {
			id: agent.id,
			name: agent.name,
			systemPrompt: agent.systemPrompt,
			capabilities: agent.capabilities
		};
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
		
		// Agent selector
		if (elements.agentSelectorBtn) {
			elements.agentSelectorBtn.addEventListener('click', function(e) {
				e.stopPropagation();
				toggleAgentDropdown();
			});
		}
		
		// Initialize agent dropdown options
		renderAgentDropdown();
		
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
			
			// Tools Used container toggle
			var toolsUsedHeader = e.target.closest('.tools-used-header');
			if (toolsUsedHeader) {
				var container = toolsUsedHeader.closest('.tools-used-container');
				container.classList.toggle('expanded');

				// Keep aria state in sync
				var expanded = container.classList.contains('expanded');
				toolsUsedHeader.setAttribute('aria-expanded', expanded ? 'true' : 'false');
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
			// Agent selector elements (for edit mode)
			agentSelector: document.getElementById('agentSelector'),
			agentSelectorBtn: document.getElementById('agentSelectorBtn'),
			agentDropdown: document.getElementById('agentDropdown'),
			selectedAgentName: document.getElementById('selectedAgentName'),
			selectedAgentIcon: document.getElementById('selectedAgentIcon'),
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
		console.log('='.repeat(60));
		console.log('[Copilot] INITIALIZATION COMPLETE');
		console.log('[Copilot] Backend URL:', Config.BACKEND_URL);
		console.log('[Copilot] Use dummy mode:', Config.USE_DUMMY);
		console.log('[Copilot] Window location:', window.location.href);
		try {
			console.log('[Copilot] Parent location:', window.parent.location.href);
		} catch (e) {
			console.log('[Copilot] Parent location: (cross-origin, cannot access)');
		}
		console.log('='.repeat(60));
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
