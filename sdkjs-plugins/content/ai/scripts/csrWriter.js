/*
 * (c) Copyright Ascensio System SIA 2010-2025
 *
 * CSR Writer Agent Script
 */

(function(window, undefined) {
	'use strict';

	// State management
	var state = {
		protocolFile: null,
		sapFile: null,
		sessionId: null,
		isIndexing: false,
		isIndexed: false,
		stats: null,
		indexedFiles: null,
		// Query state (Phase 5)
		isQuerying: false,
		currentQuery: null,
		lastResponse: null,
		lastSources: null
	};

	// Allowed file types
	var ALLOWED_TYPES = [
		'application/pdf',
		'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		'text/markdown',
		'text/x-markdown'
	];
	var ALLOWED_EXTENSIONS = ['.pdf', '.docx', '.md', '.mmd'];
	var MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB

	// API Configuration
	var API_CONFIG = {
		// Base URL will be determined dynamically based on the editor's location
		getBaseUrl: function() {
			// Try to get the server URL from the editor
			var baseUrl = '';
			try {
				// For plugin windows, we need to get the parent editor's origin
				// or use the URL that loaded the editor
				var editorUrl = window.location.origin;
				
				// Check if we're in an iframe and try to get parent origin
				try {
					if (window.parent && window.parent.location && window.parent.location.origin) {
						editorUrl = window.parent.location.origin;
					}
				} catch (e) {
					// Cross-origin restriction, use window.location
				}
				
				// In production, the API is on the same server
				// In development, we might need to adjust
				if (editorUrl.includes('localhost') || editorUrl.includes('127.0.0.1')) {
					baseUrl = 'http://localhost:8000';
				} else {
					// Use the same origin for API calls (works for ngrok)
					baseUrl = editorUrl;
				}
				
				console.log('CSR Writer: API baseUrl determined:', baseUrl, 'from:', editorUrl);
			} catch (e) {
				console.warn('CSR Writer: Could not determine base URL, using relative paths', e);
				baseUrl = '';
			}
			return baseUrl;
		},
		endpoints: {
			index: '/api/csr-writer/index',
			query: '/api/csr-writer/query',
			status: '/api/csr-writer/status/'
		}
	};

	// Guard to prevent double initialization
	var isInitialized = false;

	// Initialize when DOM is ready
	function initCSRWriter() {
		// Prevent double initialization which causes duplicate event handlers
		if (isInitialized) {
			console.log('CSR Writer: Already initialized, skipping');
			return;
		}
		isInitialized = true;
		
		console.log('CSR Writer: initCSRWriter called');
		
		// Apply translations
		applyTranslations();

		// Initialize theme
		if (window.initTheme) {
			window.initTheme();
		}

		// Setup file upload handlers
		setupFileHandlers();

		// Setup query handlers (Phase 5)
		setupQueryHandlers();

		// Load saved session if exists
		loadSession();

		console.log('CSR Writer: Initialization complete');
	}

	// Initialize plugin window (for plugin system)
	window.Asc.plugin.init = function() {
		console.log('CSR Writer: plugin.init called');
		initCSRWriter();
		
		// Send init event to parent
		window.Asc.plugin.sendToPlugin('event_onInit');
	};

	// Also initialize on DOMContentLoaded as fallback
	if (document.readyState === 'loading') {
		document.addEventListener('DOMContentLoaded', function() {
			console.log('CSR Writer: DOMContentLoaded - initializing');
			// Small delay to ensure plugin system is ready
			setTimeout(initCSRWriter, 100);
		});
	} else {
		// DOM already ready
		console.log('CSR Writer: DOM already ready - initializing');
		setTimeout(initCSRWriter, 100);
	}

	// Handle theme changes
	window.Asc.plugin.attachEvent('onThemeChanged', function(theme) {
		if (window.onThemeChange) {
			window.onThemeChange(theme);
		}
	});

	// Handle window close
	window.Asc.plugin.button = function(id) {
		window.Asc.plugin.executeMethod('CloseWindow');
	};

	// Apply i18n translations
	function applyTranslations() {
		var elements = document.querySelectorAll('.i18n');
		elements.forEach(function(el) {
			var text = el.textContent.trim();
			if (text && window.Asc.plugin.tr) {
				el.textContent = window.Asc.plugin.tr(text);
			}
		});
	}

	// Setup file input handlers
	function setupFileHandlers() {
		var protocolInput = document.getElementById('protocol-file');
		var sapInput = document.getElementById('sap-file');
		var processBtn = document.getElementById('process-btn');
		var reuploadBtn = document.getElementById('reupload-btn');

		console.log('CSR Writer: Setting up file handlers', {
			protocolInput: !!protocolInput,
			sapInput: !!sapInput,
			processBtn: !!processBtn
		});

		if (protocolInput) {
			protocolInput.addEventListener('change', function(e) {
				console.log('CSR Writer: Protocol file change event', e.target.files);
				handleFileSelect(e.target.files[0], 'protocol');
			});
		}

		if (sapInput) {
			sapInput.addEventListener('change', function(e) {
				console.log('CSR Writer: SAP file change event', e.target.files);
				handleFileSelect(e.target.files[0], 'sap');
			});
		}

		if (processBtn) {
			processBtn.addEventListener('click', function() {
				processAndIndex();
			});
		}

		if (reuploadBtn) {
			reuploadBtn.addEventListener('click', function() {
				resetUpload();
			});
		}
	}

	// Check if extension is allowed
	function isAllowedExtension(ext) {
		for (var i = 0; i < ALLOWED_EXTENSIONS.length; i++) {
			if (ALLOWED_EXTENSIONS[i] === ext) {
				return true;
			}
		}
		return false;
	}

	// Handle file selection
	function handleFileSelect(file, type) {
		console.log('CSR Writer: handleFileSelect called', { file: file, type: type });

		var statusEl = document.getElementById(type + '-status');
		var errorEl = document.getElementById(type + '-error');
		var fileNameEl = statusEl ? statusEl.querySelector('.csr-file-name') : null;

		console.log('CSR Writer: UI elements', { statusEl: !!statusEl, errorEl: !!errorEl, fileNameEl: !!fileNameEl });

		// Clear previous error
		if (errorEl) {
			errorEl.classList.add('hidden');
			errorEl.textContent = '';
		}

		if (!file) {
			console.log('CSR Writer: No file provided');
			if (fileNameEl) {
				fileNameEl.textContent = 'No file selected';
				fileNameEl.classList.remove('has-file');
			}
			if (type === 'protocol') {
				state.protocolFile = null;
			} else {
				state.sapFile = null;
			}
			updateProcessButton();
			return;
		}

		// Validate file type
		var extension = '.' + file.name.split('.').pop().toLowerCase();
		console.log('CSR Writer: File extension:', extension);
		
		if (!isAllowedExtension(extension)) {
			console.log('CSR Writer: Invalid extension');
			showFileError(type, 'Invalid file type. Please select a PDF, DOCX, MD, or MMD file.');
			return;
		}

		// Validate file size
		if (file.size > MAX_FILE_SIZE) {
			console.log('CSR Writer: File too large');
			showFileError(type, 'File is too large. Maximum size is 50MB.');
			return;
		}

		console.log('CSR Writer: File validated successfully');

		// Store file reference
		if (type === 'protocol') {
			state.protocolFile = file;
			console.log('CSR Writer: Protocol file stored:', file.name);
		} else {
			state.sapFile = file;
			console.log('CSR Writer: SAP file stored:', file.name);
		}

		// Update UI
		if (fileNameEl) {
			fileNameEl.textContent = file.name;
			fileNameEl.classList.add('has-file');
			console.log('CSR Writer: Updated file name display to:', file.name);
		}

		console.log('CSR Writer: Current state:', {
			protocolFile: state.protocolFile ? state.protocolFile.name : null,
			sapFile: state.sapFile ? state.sapFile.name : null
		});

		updateProcessButton();
	}

	// Show file error
	function showFileError(type, message) {
		var errorEl = document.getElementById(type + '-error');
		var statusEl = document.getElementById(type + '-status');
		var fileNameEl = statusEl ? statusEl.querySelector('.csr-file-name') : null;
		var inputEl = document.getElementById(type + '-file');

		if (errorEl) {
			errorEl.textContent = message;
			errorEl.classList.remove('hidden');
		}

		if (fileNameEl) {
			fileNameEl.textContent = 'No file selected';
			fileNameEl.classList.remove('has-file');
		}

		// Clear the input
		if (inputEl) {
			inputEl.value = '';
		}

		// Clear state
		if (type === 'protocol') {
			state.protocolFile = null;
		} else {
			state.sapFile = null;
		}

		updateProcessButton();
	}

	// Update process button state
	function updateProcessButton() {
		var processBtn = document.getElementById('process-btn');
		if (processBtn) {
			var hasProtocol = !!state.protocolFile;
			var hasSAP = !!state.sapFile;
			var hasFiles = hasProtocol && hasSAP;
			var shouldDisable = !hasFiles || state.isIndexing;
			
			console.log('CSR Writer: updateProcessButton', {
				hasProtocol: hasProtocol,
				hasSAP: hasSAP,
				hasFiles: hasFiles,
				isIndexing: state.isIndexing,
				shouldDisable: shouldDisable
			});
			
			processBtn.disabled = shouldDisable;
		}
	}

	// Process and index files - Phase 4: Real API integration
	function processAndIndex() {
		if (!state.protocolFile || !state.sapFile) {
			return;
		}

		state.isIndexing = true;
		updateProcessButton();

		// Show progress
		showProgress();

		// Call the backend API to index files
		indexFiles();
	}

	// Show progress UI
	function showProgress() {
		var progressContainer = document.getElementById('progress-container');
		var uploadActions = document.querySelector('.csr-upload-actions');
		var protocolUpload = document.getElementById('protocol-upload');
		var sapUpload = document.getElementById('sap-upload');

		if (uploadActions) uploadActions.classList.add('hidden');
		if (protocolUpload) protocolUpload.style.opacity = '0.5';
		if (sapUpload) sapUpload.style.opacity = '0.5';
		if (progressContainer) progressContainer.classList.remove('hidden');

		// Animate progress
		animateProgress(0);
	}

	// Animate progress bar
	var indexingStartTime = null;
	
	function animateProgress(percent, message) {
		var progressFill = document.getElementById('progress-fill');
		var progressText = document.getElementById('progress-text');

		if (progressFill) {
			progressFill.style.width = percent + '%';
		}

		if (progressText) {
			// If a custom message is provided, use it
			if (message) {
				progressText.textContent = message;
			} else {
				// Calculate elapsed time
				var elapsed = indexingStartTime ? Math.floor((Date.now() - indexingStartTime) / 1000) : 0;
				var elapsedStr = elapsed > 0 ? ' (' + formatTime(elapsed) + ')' : '';
				
				if (percent < 15) {
					progressText.textContent = 'Uploading documents...' + elapsedStr;
				} else if (percent < 30) {
					progressText.textContent = 'Documents received, starting analysis...' + elapsedStr;
				} else if (percent < 50) {
					progressText.textContent = 'Converting documents to markdown...' + elapsedStr;
				} else if (percent < 70) {
					progressText.textContent = 'Generating semantic embeddings...' + elapsedStr;
				} else if (percent < 85) {
					progressText.textContent = 'Building search index...' + elapsedStr;
				} else {
					progressText.textContent = 'Finalizing indexing...' + elapsedStr;
				}
			}
		}
	}
	
	// Format time in mm:ss
	function formatTime(seconds) {
		var mins = Math.floor(seconds / 60);
		var secs = seconds % 60;
		if (mins > 0) {
			return mins + 'm ' + secs + 's';
		}
		return secs + 's';
	}

	// Polling configuration
	var POLLING_CONFIG = {
		INDEX_MAX_WAIT_TIME: 600000,     // 10 minutes max wait for indexing
		INDEX_POLL_INTERVAL: 5000,       // Poll every 5 seconds
		INDEX_MAX_POLL_TIME: 600000,     // 10 minutes max polling
		QUERY_INITIAL_TIMEOUT: 60000,    // 60 seconds for initial query request
		QUERY_POLL_INTERVAL: 3000,       // Poll every 3 seconds
		QUERY_MAX_RETRIES: 5,            // Max query retries
		QUERY_RETRY_DELAY: 5000          // Delay between retries
	};

	// Check session status via polling
	function checkSessionStatus(sessionId) {
		var baseUrl = API_CONFIG.getBaseUrl();
		var statusUrl = baseUrl + API_CONFIG.endpoints.status + sessionId;
		
		console.log('CSR Writer: Checking session status:', statusUrl);
		
		return fetch(statusUrl, {
			method: 'GET',
			headers: { 'Accept': 'application/json' }
		})
		.then(function(response) {
			if (!response.ok) {
				throw new Error('Status check failed: ' + response.status);
			}
			return response.json();
		})
		.then(function(data) {
			console.log('CSR Writer: Session status:', data);
			return data;
		});
	}

	// Poll session status until indexed or timeout
	function pollIndexingStatus(sessionId, progressInterval) {
		var pollStartTime = Date.now();
		var totalElapsed = indexingStartTime ? (Date.now() - indexingStartTime) : 0;
		
		function poll() {
			var pollElapsed = Date.now() - pollStartTime;
			var totalTime = indexingStartTime ? Math.floor((Date.now() - indexingStartTime) / 1000) : Math.floor(pollElapsed / 1000);
			
			if (pollElapsed > POLLING_CONFIG.INDEX_MAX_POLL_TIME) {
				clearInterval(progressInterval);
				onIndexingError('Indexing is taking longer than 10 minutes. The server may be overloaded. Please try again later.');
				return;
			}
			
			checkSessionStatus(sessionId)
				.then(function(data) {
					if (data.success && data.session) {
						var status = data.session.status;
						console.log('CSR Writer: Poll status:', status, 'total elapsed:', formatTime(totalTime));
						
						if (status === 'indexed') {
							// Indexing complete!
							clearInterval(progressInterval);
							animateProgress(100, 'Indexing complete!');
							
							var stats = data.session.stats || {};
							var files = data.session.files || {};
							
							console.log('CSR Writer: Indexing completed via polling in', formatTime(totalTime));
							
							setTimeout(function() {
								onIndexingComplete(sessionId, stats, files);
							}, 500);
						} else if (status === 'error') {
							clearInterval(progressInterval);
							onIndexingError(data.session.error || 'Indexing failed on server');
						} else if (status === 'indexing') {
							// Still indexing, continue polling with progress message
							var progressMsg = 'Server is indexing... (' + formatTime(totalTime) + ')';
							if (totalTime > 120) {
								progressMsg = 'Still processing (this is normal for large documents)... (' + formatTime(totalTime) + ')';
							}
							if (totalTime > 300) {
								progressMsg = 'Almost there... (' + formatTime(totalTime) + ')';
							}
							animateProgress(null, progressMsg);
							setTimeout(poll, POLLING_CONFIG.INDEX_POLL_INTERVAL);
						} else {
							// Unknown status, continue polling
							console.log('CSR Writer: Unknown status:', status);
							setTimeout(poll, POLLING_CONFIG.INDEX_POLL_INTERVAL);
						}
					} else {
						// Unexpected response, retry
						console.log('CSR Writer: Unexpected poll response, retrying...');
						setTimeout(poll, POLLING_CONFIG.INDEX_POLL_INTERVAL);
					}
				})
				.catch(function(error) {
					console.error('CSR Writer: Status poll error:', error);
					// Continue polling on error - server might be busy
					animateProgress(null, 'Checking server... (' + formatTime(totalTime) + ')');
					setTimeout(poll, POLLING_CONFIG.INDEX_POLL_INTERVAL);
				});
		}
		
		// Start polling immediately
		poll();
	}

	// Index files via backend API - Phase 4/7 (Real CSR Agent with long polling)
	function indexFiles() {
		var baseUrl = API_CONFIG.getBaseUrl();
		var indexUrl = baseUrl + API_CONFIG.endpoints.index;
		
		console.log('CSR Writer: Starting file indexing (Real CSR Agent - no timeout)', {
			url: indexUrl,
			protocol: state.protocolFile.name,
			sap: state.sapFile.name
		});

		// Create FormData with both files
		var formData = new FormData();
		formData.append('protocol', state.protocolFile, state.protocolFile.name);
		formData.append('sap', state.sapFile, state.sapFile.name);

		// Track start time for elapsed display
		indexingStartTime = Date.now();
		
		// Start progress animation with elapsed time updates
		var progressInterval = startProgressAnimation();
		
		// Also start an elapsed time display updater
		var elapsedInterval = setInterval(function() {
			var elapsed = Math.floor((Date.now() - indexingStartTime) / 1000);
			var progressText = document.getElementById('progress-text');
			if (progressText && state.isIndexing) {
				var currentText = progressText.textContent;
				// Update time in the message
				if (elapsed > 60) {
					progressText.textContent = currentText.replace(/\([^)]*\)/, '(' + formatTime(elapsed) + ')');
				}
			}
			
			// Safety check - if indexing takes more than 10 minutes, something is wrong
			if (elapsed > 600 && state.isIndexing) {
				clearInterval(elapsedInterval);
				clearInterval(progressInterval);
				onIndexingError('Indexing is taking unusually long. The server may be overloaded. Please try again later.');
			}
		}, 1000);

		// Make the API request WITHOUT a timeout - let it run for as long as needed
		// The backend proxy has a 10-minute timeout
		fetch(indexUrl, {
			method: 'POST',
			body: formData
			// No AbortController - let the request complete naturally
		})
		.then(function(response) {
			clearInterval(elapsedInterval);
			console.log('CSR Writer: Index response status:', response.status);
			if (!response.ok) {
				return response.json().then(function(data) {
					throw new Error(data.error || 'Indexing failed with status ' + response.status);
				});
			}
			return response.json();
		})
		.then(function(data) {
			console.log('CSR Writer: Index response:', data);
			clearInterval(elapsedInterval);
			
			if (data.success && data.sessionId) {
				// Check if already indexed or need to poll
				if (data.stats && data.stats.totalChunks > 0) {
					// Indexing already complete
					clearInterval(progressInterval);
					animateProgress(100, 'Indexing complete!');
					
					var totalTime = Math.floor((Date.now() - indexingStartTime) / 1000);
					console.log('CSR Writer: Indexing completed in', formatTime(totalTime));
					
					setTimeout(function() {
						onIndexingComplete(data.sessionId, data.stats, data.files);
					}, 500);
				} else {
					// Indexing in progress, start polling
					console.log('CSR Writer: Indexing started, polling for completion...');
					state.sessionId = data.sessionId;
					saveSession(); // Save early so we can recover
					pollIndexingStatus(data.sessionId, progressInterval);
				}
			} else {
				throw new Error(data.error || 'Indexing failed - no session ID returned');
			}
		})
		.catch(function(error) {
			clearInterval(elapsedInterval);
			clearInterval(progressInterval);
			console.error('CSR Writer: Index error:', error);
			
			// Check if error message indicates timeout
			var errorMsg = error.message || '';
			if (errorMsg.includes('timeout') || errorMsg.includes('Timeout') || errorMsg.includes('TIMEOUT')) {
				// Server or proxy timeout - try polling
				console.log('CSR Writer: Request likely timed out, attempting status poll recovery...');
				
				// Try to get any existing session from localStorage and poll
				var savedSessionId = null;
				try {
					var saved = localStorage.getItem('csr_writer_session');
					if (saved) {
						var parsedData = JSON.parse(saved);
						savedSessionId = parsedData.sessionId;
					}
				} catch (e) {}
				
				if (savedSessionId) {
					console.log('CSR Writer: Found saved session, polling:', savedSessionId);
					var recoveryProgressInterval = startProgressAnimation();
					animateProgress(null, 'Checking indexing status...');
					pollIndexingStatus(savedSessionId, recoveryProgressInterval);
				} else {
					onIndexingError('Request timed out. Please try uploading again - subsequent uploads of same files are faster.');
				}
			} else {
				onIndexingError(error.message || 'Failed to connect to server');
			}
		});
	}

	// Start progress animation that increases gradually (optimized for real CSR Agent)
	// Indexing can take 1-10 minutes for first-time documents
	function startProgressAnimation() {
		var progress = 0;
		var tickCount = 0;
		
		var interval = setInterval(function() {
			tickCount++;
			var elapsedSecs = Math.floor((Date.now() - indexingStartTime) / 1000);
			
			// Slow down as we approach 90% (never reach 100% until complete)
			// Each tick is 500ms, so 120 ticks = 1 minute
			// We want to reach ~50% at 2 minutes, ~70% at 5 minutes, ~85% at 8 minutes
			if (progress < 15) {
				// First 15%: quick upload phase (~15 seconds)
				progress += Math.random() * 2 + 1;
			} else if (progress < 30) {
				// 15-30%: conversion phase (~30 seconds)
				progress += Math.random() * 0.8 + 0.4;
			} else if (progress < 50) {
				// 30-50%: embedding generation (~2 minutes)
				progress += Math.random() * 0.15 + 0.08;
			} else if (progress < 70) {
				// 50-70%: indexing phase (~3 minutes)
				progress += Math.random() * 0.1 + 0.04;
			} else if (progress < 85) {
				// 70-85%: finalization (~2 minutes)
				progress += Math.random() * 0.05 + 0.02;
			} else if (progress < 90) {
				// 85-90%: very slow crawl
				progress += Math.random() * 0.01;
			}
			
			// Cap at 90% until we get the actual response
			var cappedProgress = Math.min(progress, 90);
			
			// Generate message based on progress and elapsed time
			var message = null;
			if (elapsedSecs > 10) {
				var timeStr = formatTime(elapsedSecs);
				if (cappedProgress < 20) {
					message = 'Uploading to server... (' + timeStr + ')';
				} else if (cappedProgress < 40) {
					message = 'Converting documents... (' + timeStr + ')';
				} else if (cappedProgress < 60) {
					message = 'Generating embeddings... (' + timeStr + ')';
				} else if (cappedProgress < 80) {
					message = 'Building search index... (' + timeStr + ')';
				} else {
					message = 'Finalizing... (' + timeStr + ')';
				}
			}
			
			animateProgress(cappedProgress, message);
		}, 500);
		
		return interval;
	}

	// Handle indexing error
	function onIndexingError(message) {
		state.isIndexing = false;
		updateProcessButton();

		// Hide progress
		var progressContainer = document.getElementById('progress-container');
		if (progressContainer) progressContainer.classList.add('hidden');

		// Reset upload areas opacity
		var protocolUpload = document.getElementById('protocol-upload');
		var sapUpload = document.getElementById('sap-upload');
		if (protocolUpload) protocolUpload.style.opacity = '1';
		if (sapUpload) sapUpload.style.opacity = '1';

		// Show upload actions again
		var uploadActions = document.querySelector('.csr-upload-actions');
		if (uploadActions) uploadActions.classList.remove('hidden');

		// Show error message
		showIndexError(message);
	}

	// Show index error in UI
	function showIndexError(message) {
		// Check if error container exists, create if not
		var errorContainer = document.getElementById('index-error');
		if (!errorContainer) {
			var uploadSection = document.querySelector('.csr-section-content');
			if (uploadSection) {
				errorContainer = document.createElement('div');
				errorContainer.id = 'index-error';
				errorContainer.className = 'csr-index-error';
				uploadSection.appendChild(errorContainer);
			}
		}

		if (errorContainer) {
			errorContainer.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>' +
				'<span>' + escapeHtml(message) + '</span>' +
				'<button class="csr-btn csr-btn-text" onclick="document.getElementById(\'index-error\').classList.add(\'hidden\')">Dismiss</button>';
			errorContainer.classList.remove('hidden');
		}
	}

	// Escape HTML to prevent XSS
	function escapeHtml(text) {
		var div = document.createElement('div');
		div.textContent = text;
		return div.innerHTML;
	}

	// Handle indexing completion
	function onIndexingComplete(sessionId, stats, files) {
		state.isIndexing = false;
		state.isIndexed = true;
		state.sessionId = sessionId;
		state.stats = stats || {};
		state.indexedFiles = files || {};

		// Save session
		saveSession();

		// Hide progress, show success
		var progressContainer = document.getElementById('progress-container');
		var successMessage = document.getElementById('success-message');
		var querySection = document.getElementById('query-section');
		var uploadBadge = document.getElementById('upload-status-badge');
		var uploadActions = document.querySelector('.csr-upload-actions');

		if (progressContainer) progressContainer.classList.add('hidden');
		if (uploadActions) uploadActions.classList.add('hidden');
		if (successMessage) {
			successMessage.classList.remove('hidden');
			// Update success message with stats
			updateSuccessMessage(stats, files);
		}
		if (querySection) querySection.classList.remove('hidden');
		if (uploadBadge) {
			uploadBadge.textContent = 'Ready';
			uploadBadge.className = 'csr-section-badge csr-badge-success';
		}

		console.log('CSR Writer: Documents indexed with session:', sessionId, 'Stats:', stats);
	}

	// Update success message with indexing stats
	function updateSuccessMessage(stats, files) {
		var successMessage = document.getElementById('success-message');
		if (!successMessage) return;

		var statsHtml = '';
		if (stats) {
			statsHtml = '<div class="csr-index-stats">';
			if (stats.totalChunks) {
				statsHtml += '<span class="csr-stat"><strong>' + stats.totalChunks + '</strong> chunks indexed</span>';
			}
			if (stats.indexingTimeMs) {
				var timeSeconds = (stats.indexingTimeMs / 1000).toFixed(1);
				statsHtml += '<span class="csr-stat">in <strong>' + timeSeconds + 's</strong></span>';
			}
			statsHtml += '</div>';
		}

		var filesHtml = '';
		if (files) {
			filesHtml = '<div class="csr-indexed-files">';
			if (files.protocol) {
				filesHtml += '<span class="csr-indexed-file">📄 ' + escapeHtml(files.protocol) + '</span>';
			}
			if (files.sap) {
				filesHtml += '<span class="csr-indexed-file">📊 ' + escapeHtml(files.sap) + '</span>';
			}
			filesHtml += '</div>';
		}

		// Find or create stats container
		var statsContainer = successMessage.querySelector('.csr-success-details');
		if (!statsContainer) {
			statsContainer = document.createElement('div');
			statsContainer.className = 'csr-success-details';
			// Insert after the success icon and text
			var reuploadBtn = successMessage.querySelector('#reupload-btn');
			if (reuploadBtn) {
				successMessage.insertBefore(statsContainer, reuploadBtn);
			} else {
				successMessage.appendChild(statsContainer);
			}
		}
		statsContainer.innerHTML = filesHtml + statsHtml;
	}

	// Reset upload state
	function resetUpload() {
		state.protocolFile = null;
		state.sapFile = null;
		state.sessionId = null;
		state.isIndexed = false;
		state.isIndexing = false;
		state.stats = null;
		state.indexedFiles = null;

		// Clear saved session
		clearSession();

		// Reset UI
		var protocolInput = document.getElementById('protocol-file');
		var sapInput = document.getElementById('sap-file');
		var protocolStatus = document.getElementById('protocol-status');
		var sapStatus = document.getElementById('sap-status');
		var successMessage = document.getElementById('success-message');
		var querySection = document.getElementById('query-section');
		var uploadActions = document.querySelector('.csr-upload-actions');
		var protocolUpload = document.getElementById('protocol-upload');
		var sapUpload = document.getElementById('sap-upload');
		var uploadBadge = document.getElementById('upload-status-badge');

		if (protocolInput) protocolInput.value = '';
		if (sapInput) sapInput.value = '';
		
		if (protocolStatus) {
			var pName = protocolStatus.querySelector('.csr-file-name');
			if (pName) {
				pName.textContent = 'No file selected';
				pName.classList.remove('has-file');
			}
		}
		
		if (sapStatus) {
			var sName = sapStatus.querySelector('.csr-file-name');
			if (sName) {
				sName.textContent = 'No file selected';
				sName.classList.remove('has-file');
			}
		}

		if (successMessage) successMessage.classList.add('hidden');
		if (querySection) querySection.classList.add('hidden');
		if (uploadActions) uploadActions.classList.remove('hidden');
		if (protocolUpload) protocolUpload.style.opacity = '1';
		if (sapUpload) sapUpload.style.opacity = '1';
		if (uploadBadge) uploadBadge.textContent = '';

		updateProcessButton();
	}

	// Save session to localStorage
	function saveSession() {
		if (state.sessionId) {
			try {
				localStorage.setItem('csr_writer_session', JSON.stringify({
					sessionId: state.sessionId,
					protocolName: state.protocolFile ? state.protocolFile.name : (state.indexedFiles ? state.indexedFiles.protocol : null),
					sapName: state.sapFile ? state.sapFile.name : (state.indexedFiles ? state.indexedFiles.sap : null),
					stats: state.stats || null,
					files: state.indexedFiles || null,
					timestamp: Date.now()
				}));
				console.log('CSR Writer: Session saved to localStorage');
			} catch (e) {
				console.error('Failed to save session:', e);
			}
		}
	}

	// Load session from localStorage
	function loadSession() {
		try {
			var saved = localStorage.getItem('csr_writer_session');
			if (saved) {
				var data = JSON.parse(saved);
				// Check if session is not too old (24 hours)
				if (Date.now() - data.timestamp < 24 * 60 * 60 * 1000) {
					state.sessionId = data.sessionId;
					state.isIndexed = true;
					state.stats = data.stats || {};
					state.indexedFiles = data.files || {
						protocol: data.protocolName,
						sap: data.sapName
					};

					// Show indexed state
					var successMessage = document.getElementById('success-message');
					var querySection = document.getElementById('query-section');
					var uploadActions = document.querySelector('.csr-upload-actions');
					var uploadBadge = document.getElementById('upload-status-badge');
					var protocolUpload = document.getElementById('protocol-upload');
					var sapUpload = document.getElementById('sap-upload');

					if (successMessage) {
						successMessage.classList.remove('hidden');
						// Update success message with saved stats
						updateSuccessMessage(state.stats, state.indexedFiles);
					}
					if (querySection) querySection.classList.remove('hidden');
					if (uploadActions) uploadActions.classList.add('hidden');
					if (uploadBadge) {
						uploadBadge.textContent = 'Ready';
						uploadBadge.className = 'csr-section-badge csr-badge-success';
					}
					if (protocolUpload) protocolUpload.style.opacity = '0.5';
					if (sapUpload) sapUpload.style.opacity = '0.5';

					// Update query button state after session restore
					updateQueryButton();

					console.log('CSR Writer: Restored session:', data.sessionId, 'Stats:', state.stats, 'isIndexed:', state.isIndexed);
				} else {
					console.log('CSR Writer: Session expired, clearing');
					clearSession();
				}
			}
		} catch (e) {
			console.error('Failed to load session:', e);
		}
	}

	// Clear session from localStorage
	function clearSession() {
		try {
			localStorage.removeItem('csr_writer_session');
		} catch (e) {
			console.error('Failed to clear session:', e);
		}
	}

	// ============================================
	// Phase 5: Query Handling with SSE Streaming
	// ============================================

	// Setup query input handlers
	function setupQueryHandlers() {
		var queryInput = document.getElementById('query-input');
		var queryBtn = document.getElementById('query-btn');
		var copyBtn = document.getElementById('copy-btn');
		var newQueryBtn = document.getElementById('new-query-btn');
		var retryBtn = document.getElementById('retry-query-btn');
		var suggestions = document.querySelectorAll('.csr-suggestion-chip');

		console.log('CSR Writer: Setting up query handlers', {
			queryInput: !!queryInput,
			queryBtn: !!queryBtn,
			suggestions: suggestions.length
		});

		// Query input - auto-resize and enable button
		if (queryInput) {
			queryInput.addEventListener('input', function() {
				// Auto-resize textarea
				this.style.height = 'auto';
				this.style.height = Math.min(this.scrollHeight, 120) + 'px';
				// Update button state
				updateQueryButton();
			});

			// Submit on Enter (without Shift)
			queryInput.addEventListener('keydown', function(e) {
				if (e.key === 'Enter' && !e.shiftKey) {
					e.preventDefault();
					if (!state.isQuerying && queryInput.value.trim()) {
						submitQuery();
					}
				}
			});
		}

		// Query submit button
		if (queryBtn) {
			queryBtn.addEventListener('click', function() {
				console.log('CSR Writer: Query button clicked!', {
					isQuerying: state.isQuerying,
					isIndexed: state.isIndexed,
					sessionId: state.sessionId,
					disabled: queryBtn.disabled
				});
				submitQuery();
			});
			console.log('CSR Writer: Query button event listener attached');
		} else {
			console.error('CSR Writer: Query button not found!');
		}

		// Suggestion chips
		suggestions.forEach(function(chip) {
			chip.addEventListener('click', function() {
				var query = this.getAttribute('data-query');
				if (query && queryInput) {
					queryInput.value = query;
					queryInput.dispatchEvent(new Event('input'));
					submitQuery();
				}
			});
		});

		// Copy button
		if (copyBtn) {
			copyBtn.addEventListener('click', function() {
				copyResponseToClipboard();
			});
		}

		// New query button
		if (newQueryBtn) {
			newQueryBtn.addEventListener('click', function() {
				resetQueryUI();
			});
		}

		// Insert at cursor button
		var insertBtn = document.getElementById('insert-btn');
		if (insertBtn) {
			insertBtn.addEventListener('click', function() {
				insertAtCursor();
			});
		}

		// Replace selection button
		var replaceBtn = document.getElementById('replace-btn');
		if (replaceBtn) {
			replaceBtn.addEventListener('click', function() {
				replaceSelection();
			});
		}

		// Retry button
		if (retryBtn) {
			retryBtn.addEventListener('click', function() {
				if (state.currentQuery) {
					var queryInput = document.getElementById('query-input');
					if (queryInput) queryInput.value = state.currentQuery;
					submitQuery();
				}
			});
		}

		// Set initial button state
		updateQueryButton();
		console.log('CSR Writer: Query handlers setup complete');
	}

	// Update query button state
	function updateQueryButton() {
		var queryInput = document.getElementById('query-input');
		var queryBtn = document.getElementById('query-btn');
		
		if (queryBtn && queryInput) {
			var hasQuery = queryInput.value.trim().length > 0;
			queryBtn.disabled = !hasQuery || state.isQuerying || !state.isIndexed;
		}
	}

	// Submit query to backend
	function submitQuery() {
		console.log('CSR Writer: submitQuery called');
		
		var queryInput = document.getElementById('query-input');
		if (!queryInput) {
			console.error('CSR Writer: Query input element not found');
			return;
		}

		var query = queryInput.value.trim();
		
		// Try to get session from localStorage if not in state
		if (!state.sessionId) {
			console.log('CSR Writer: sessionId not in state, checking localStorage');
			try {
				var saved = localStorage.getItem('csr_writer_session');
				if (saved) {
					var data = JSON.parse(saved);
					if (data.sessionId) {
						state.sessionId = data.sessionId;
						state.isIndexed = true;
						console.log('CSR Writer: Recovered sessionId from localStorage:', state.sessionId);
					}
				}
			} catch (e) {
				console.error('CSR Writer: Error reading localStorage:', e);
			}
		}
		
		console.log('CSR Writer: Query validation', {
			query: query,
			hasQuery: !!query,
			isQuerying: state.isQuerying,
			sessionId: state.sessionId,
			isIndexed: state.isIndexed
		});
		
		if (!query) {
			console.log('CSR Writer: Cannot submit - no query text');
			return;
		}
		
		if (state.isQuerying) {
			console.log('CSR Writer: Cannot submit - already querying');
			return;
		}
		
		if (!state.sessionId) {
			console.error('CSR Writer: Cannot submit - no session ID. Please re-index documents.');
			showQueryError('No active session. Please upload and index your documents first.');
			return;
		}

		console.log('CSR Writer: Submitting query:', query);

		state.isQuerying = true;
		state.currentQuery = query;
		state.lastResponse = '';
		state.lastSources = [];

		updateQueryButton();
		showQueryLoading();
		hideQueryError();

		// Call the query API with SSE
		streamQuery(query);
	}

	// Stream query via SSE with retry logic
	var queryRetryCount = 0;
	var queryStartTime = null;
	
	function streamQuery(query, retryAttempt) {
		retryAttempt = retryAttempt || 0;
		var baseUrl = API_CONFIG.getBaseUrl();
		var queryUrl = baseUrl + API_CONFIG.endpoints.query;

		console.log('CSR Writer: Starting SSE query (attempt ' + (retryAttempt + 1) + ')', { 
			url: queryUrl, 
			query: query,
			sessionId: state.sessionId 
		});

		if (retryAttempt === 0) {
			queryStartTime = Date.now();
		}

		// Create an AbortController for timeout handling
		var controller = new AbortController();
		var timeoutId = setTimeout(function() {
			controller.abort();
		}, POLLING_CONFIG.QUERY_INITIAL_TIMEOUT);
		
		var receivedData = false;

		// Use fetch with streaming response
		fetch(queryUrl, {
			method: 'POST',
			headers: {
				'Content-Type': 'application/json'
			},
			body: JSON.stringify({
				sessionId: state.sessionId,
				query: query
			}),
			signal: controller.signal
		})
		.then(function(response) {
			clearTimeout(timeoutId);
			console.log('CSR Writer: Query response received', {
				status: response.status,
				ok: response.ok,
				contentType: response.headers.get('content-type')
			});
			
			if (!response.ok) {
				return response.text().then(function(text) {
					console.error('CSR Writer: Query response error body:', text);
					throw new Error('Query failed with status ' + response.status + ': ' + text);
				});
			}
			
			// Check if response.body is available (ReadableStream)
			if (!response.body) {
				console.error('CSR Writer: ReadableStream not supported');
				return response.text().then(function(text) {
					console.log('CSR Writer: Full response text:', text);
					receivedData = true;
					// Parse as SSE manually
					var lines = text.split('\n');
					lines.forEach(function(line) {
						if (line.startsWith('data: ')) {
							var data = line.slice(6);
							try {
								var parsed = JSON.parse(data);
								handleSSEMessage(parsed);
							} catch (e) {
								console.warn('CSR Writer: Failed to parse SSE line:', data);
							}
						}
					});
					onQueryComplete();
				});
			}
			
			var reader = response.body.getReader();
			var decoder = new TextDecoder();
			var buffer = '';

			function processStream() {
				return reader.read().then(function(result) {
					if (result.done) {
						console.log('CSR Writer: Stream complete');
						// Process any remaining buffer
						if (buffer.trim()) {
							var lines = buffer.split('\n');
							lines.forEach(function(line) {
								if (line.startsWith('data: ')) {
									var data = line.slice(6);
									try {
										var parsed = JSON.parse(data);
										handleSSEMessage(parsed);
									} catch (e) {
										console.warn('CSR Writer: Failed to parse final SSE data:', data);
									}
								}
							});
						}
						onQueryComplete();
						return;
					}

					receivedData = true;
					buffer += decoder.decode(result.value, { stream: true });
					
					// Process complete SSE messages
					var lines = buffer.split('\n');
					buffer = lines.pop() || ''; // Keep incomplete line in buffer

					lines.forEach(function(line) {
						if (line.startsWith('data: ')) {
							var data = line.slice(6);
							try {
								var parsed = JSON.parse(data);
								handleSSEMessage(parsed);
							} catch (e) {
								console.warn('CSR Writer: Failed to parse SSE data:', data, e);
							}
						}
					});

					return processStream();
				}).catch(function(readError) {
					console.error('CSR Writer: Stream read error:', readError);
					// Try to retry on stream errors
					handleQueryRetry(query, retryAttempt, 'Stream error: ' + readError.message);
				});
			}

			return processStream();
		})
		.catch(function(error) {
			clearTimeout(timeoutId);
			console.error('CSR Writer: Query fetch error:', error);
			
			// Handle timeout with retry logic
			if (error.name === 'AbortError') {
				handleQueryTimeout(query, retryAttempt);
			} else {
				handleQueryRetry(query, retryAttempt, error.message || 'Failed to connect to server');
			}
		});
	}
	
	// Handle query timeout - check session status and retry
	function handleQueryTimeout(query, retryAttempt) {
		var elapsed = Date.now() - queryStartTime;
		console.log('CSR Writer: Query timeout after', formatTime(Math.floor(elapsed / 1000)));
		
		// Check if we've exceeded max polling time (8 minutes)
		if (elapsed > POLLING_CONFIG.INDEX_MAX_POLL_TIME) {
			onQueryError('Query timed out after ' + formatTime(Math.floor(elapsed / 1000)) + '. The server may be overloaded.');
			return;
		}
		
		// Update loading text
		updateLoadingText('Checking server status... (' + formatTime(Math.floor(elapsed / 1000)) + ')');
		
		// Check session status before retrying
		checkSessionStatus(state.sessionId)
			.then(function(data) {
				if (data.success && data.session) {
					var status = data.session.status;
					console.log('CSR Writer: Session status during query timeout:', status);
					
					if (status === 'indexed') {
						// Session is healthy, retry the query
						if (retryAttempt < POLLING_CONFIG.QUERY_MAX_RETRIES) {
							updateLoadingText('Retrying query... (attempt ' + (retryAttempt + 2) + ')');
							setTimeout(function() {
								streamQuery(query, retryAttempt + 1);
							}, POLLING_CONFIG.QUERY_RETRY_DELAY);
						} else {
							onQueryError('Query timed out after multiple retries. Please try a simpler question.');
						}
					} else if (status === 'indexing') {
						// Still indexing, wait and retry
						updateLoadingText('Server still processing documents... (' + formatTime(Math.floor(elapsed / 1000)) + ')');
						setTimeout(function() {
							handleQueryTimeout(query, retryAttempt);
						}, POLLING_CONFIG.QUERY_POLL_INTERVAL);
					} else {
						onQueryError('Session is in an unexpected state: ' + status);
					}
				} else {
					onQueryError('Could not verify session status. Please re-upload documents.');
				}
			})
			.catch(function(error) {
				console.error('CSR Writer: Status check failed during query timeout:', error);
				// Retry anyway if we haven't exceeded retries
				if (retryAttempt < POLLING_CONFIG.QUERY_MAX_RETRIES) {
					updateLoadingText('Connection issue, retrying... (attempt ' + (retryAttempt + 2) + ')');
					setTimeout(function() {
						streamQuery(query, retryAttempt + 1);
					}, POLLING_CONFIG.QUERY_RETRY_DELAY);
				} else {
					onQueryError('CSR Agent request timeout. Please check your connection and try again.');
				}
			});
	}
	
	// Handle query retry for non-timeout errors
	function handleQueryRetry(query, retryAttempt, errorMessage) {
		console.log('CSR Writer: Query error, attempt', retryAttempt + 1, ':', errorMessage);
		
		if (retryAttempt < POLLING_CONFIG.QUERY_MAX_RETRIES) {
			updateLoadingText('Retrying... (attempt ' + (retryAttempt + 2) + ')');
			setTimeout(function() {
				streamQuery(query, retryAttempt + 1);
			}, POLLING_CONFIG.QUERY_RETRY_DELAY);
		} else {
			onQueryError(errorMessage);
		}
	}

	// Handle individual SSE message
	function handleSSEMessage(data) {
		console.log('CSR Writer: SSE message:', data.type);

		switch (data.type) {
			case 'status':
				updateLoadingText(data.message);
				break;

			case 'chunk':
				hideQueryLoading();
				showResponseContainer();
				appendResponseText(data.text);
				break;

			case 'sources':
				if (data.sources && data.sources.length > 0) {
					state.lastSources = data.sources;
					displaySources(data.sources);
				}
				break;

			case 'done':
				onQueryComplete();
				break;

			case 'error':
				onQueryError(data.message || 'Query failed');
				break;
		}
	}

	// Show loading state
	function showQueryLoading() {
		var responseContainer = document.getElementById('response-container');
		var loadingEl = document.getElementById('response-loading');
		var contentEl = document.getElementById('response-content');
		var sourcesEl = document.getElementById('sources-container');
		var actionsEl = document.getElementById('response-actions');

		if (responseContainer) responseContainer.classList.remove('hidden');
		if (loadingEl) loadingEl.classList.remove('hidden');
		if (contentEl) {
			var textEl = contentEl.querySelector('#response-text');
			if (textEl) textEl.textContent = '';
		}
		if (sourcesEl) sourcesEl.classList.add('hidden');
		if (actionsEl) actionsEl.classList.add('hidden');
	}

	// Hide loading state
	function hideQueryLoading() {
		var loadingEl = document.getElementById('response-loading');
		if (loadingEl) loadingEl.classList.add('hidden');
	}

	// Update loading text
	function updateLoadingText(text) {
		var loadingText = document.getElementById('loading-text');
		if (loadingText) {
			loadingText.textContent = text;
		}
	}

	// Show response container
	function showResponseContainer() {
		var responseContainer = document.getElementById('response-container');
		var responseText = document.getElementById('response-text');
		
		if (responseContainer) responseContainer.classList.remove('hidden');
		if (responseText) responseText.classList.add('typing');
	}

	// Append text to response (streaming effect)
	function appendResponseText(text) {
		var responseText = document.getElementById('response-text');
		if (responseText) {
			state.lastResponse += text;
			responseText.textContent = state.lastResponse;
			
			// Auto-scroll to bottom
			var container = responseText.closest('.csr-response-container');
			if (container) {
				container.scrollTop = container.scrollHeight;
			}
		}
	}

	// Display sources
	function displaySources(sources) {
		var sourcesContainer = document.getElementById('sources-container');
		var sourcesList = document.getElementById('sources-list');

		if (!sourcesContainer || !sourcesList) return;

		sourcesList.innerHTML = '';

		sources.forEach(function(source) {
			var iconClass = source.type === 'protocol' ? 'protocol' : 'sap';
			var iconSvg = source.type === 'protocol' 
				? '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>'
				: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>';

			var card = document.createElement('div');
			card.className = 'csr-source-card';
			card.innerHTML = 
				'<div class="csr-source-icon ' + iconClass + '">' + iconSvg + '</div>' +
				'<div class="csr-source-info">' +
					'<div class="csr-source-title">' + escapeHtml(source.title) + '</div>' +
					'<div class="csr-source-detail">' + escapeHtml(source.section || '') + (source.page ? ' • Page ' + source.page : '') + '</div>' +
				'</div>';

			sourcesList.appendChild(card);
		});

		sourcesContainer.classList.remove('hidden');
	}

	// Query complete
	function onQueryComplete() {
		state.isQuerying = false;
		updateQueryButton();

		// Remove typing cursor
		var responseText = document.getElementById('response-text');
		if (responseText) {
			responseText.classList.remove('typing');
		}

		// Show actions
		var actionsEl = document.getElementById('response-actions');
		if (actionsEl) actionsEl.classList.remove('hidden');

		console.log('CSR Writer: Query complete, response length:', state.lastResponse.length);
	}

	// Query error
	function onQueryError(message) {
		state.isQuerying = false;
		updateQueryButton();
		hideQueryLoading();

		// Hide response container if empty
		if (!state.lastResponse) {
			var responseContainer = document.getElementById('response-container');
			if (responseContainer) responseContainer.classList.add('hidden');
		}

		showQueryError(message);
	}

	// Show query error
	function showQueryError(message) {
		var errorEl = document.getElementById('query-error');
		var errorText = document.getElementById('query-error-text');

		if (errorEl && errorText) {
			errorText.textContent = message;
			errorEl.classList.remove('hidden');
		}
	}

	// Hide query error
	function hideQueryError() {
		var errorEl = document.getElementById('query-error');
		if (errorEl) {
			errorEl.classList.add('hidden');
		}
	}

	// Copy response to clipboard
	function copyResponseToClipboard() {
		if (!state.lastResponse) return;

		if (navigator.clipboard && navigator.clipboard.writeText) {
			navigator.clipboard.writeText(state.lastResponse)
				.then(function() {
					showCopyFeedback('Copied!');
				})
				.catch(function() {
					fallbackCopy();
				});
		} else {
			fallbackCopy();
		}
	}

	// Fallback copy method
	function fallbackCopy() {
		var textArea = document.createElement('textarea');
		textArea.value = state.lastResponse;
		textArea.style.position = 'fixed';
		textArea.style.left = '-9999px';
		document.body.appendChild(textArea);
		textArea.select();
		
		try {
			document.execCommand('copy');
			showCopyFeedback('Copied!');
		} catch (e) {
			showCopyFeedback('Failed to copy');
		}
		
		document.body.removeChild(textArea);
	}

	// Show copy feedback
	function showCopyFeedback(message) {
		var copyBtn = document.getElementById('copy-btn');
		if (copyBtn) {
			var originalText = copyBtn.innerHTML;
			copyBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg><span>' + message + '</span>';
			
			setTimeout(function() {
				copyBtn.innerHTML = originalText;
			}, 2000);
		}
	}

	// Insert response at cursor position in document
	function insertAtCursor() {
		if (!state.lastResponse) {
			console.log('CSR Writer: No response to insert');
			return;
		}

		console.log('CSR Writer: Inserting at cursor, length:', state.lastResponse.length);

		try {
			// Send to parent plugin window to insert into document
			window.Asc.plugin.sendToPlugin('event_onCSRInsert', {
				type: 'insert',
				content: state.lastResponse
			});
			
			showActionFeedback('Inserted!');
		} catch (e) {
			console.error('CSR Writer: Insert error:', e);
			// Fallback: try direct method
			try {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					// Convert markdown-style formatting to HTML for better rendering
					var htmlContent = convertToHtml(state.lastResponse);
					window.Asc.plugin.executeMethod('PasteHtml', [htmlContent]);
					showActionFeedback('Inserted!');
				}
			} catch (e2) {
				console.error('CSR Writer: Fallback insert error:', e2);
				showActionFeedback('Insert failed');
			}
		}
	}

	// Replace selected text in document
	function replaceSelection() {
		if (!state.lastResponse) {
			console.log('CSR Writer: No response to replace with');
			return;
		}

		console.log('CSR Writer: Replacing selection, length:', state.lastResponse.length);

		try {
			// Send to parent plugin window to replace selection
			window.Asc.plugin.sendToPlugin('event_onCSRInsert', {
				type: 'replace',
				content: state.lastResponse
			});
			
			showActionFeedback('Replaced!');
		} catch (e) {
			console.error('CSR Writer: Replace error:', e);
			// Fallback: try direct method - same as insert since it replaces selection
			try {
				if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
					var htmlContent = convertToHtml(state.lastResponse);
					window.Asc.plugin.executeMethod('PasteHtml', [htmlContent]);
					showActionFeedback('Replaced!');
				}
			} catch (e2) {
				console.error('CSR Writer: Fallback replace error:', e2);
				showActionFeedback('Replace failed');
			}
		}
	}

	// Convert markdown-style text to simple HTML
	function convertToHtml(text) {
		if (!text) return '';
		
		var html = text
			// Escape HTML entities first
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			// Bold: **text** or __text__
			.replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
			.replace(/__(.+?)__/g, '<b>$1</b>')
			// Italic: *text* or _text_
			.replace(/\*(.+?)\*/g, '<i>$1</i>')
			.replace(/_(.+?)_/g, '<i>$1</i>')
			// Headers
			.replace(/^### (.+)$/gm, '<h3>$1</h3>')
			.replace(/^## (.+)$/gm, '<h2>$1</h2>')
			.replace(/^# (.+)$/gm, '<h1>$1</h1>')
			// Lists
			.replace(/^- (.+)$/gm, '• $1')
			.replace(/^\d+\. (.+)$/gm, function(match, p1, offset, string) {
				return match; // Keep numbered lists as-is
			})
			// Line breaks
			.replace(/\n\n/g, '</p><p>')
			.replace(/\n/g, '<br>');
		
		return '<p>' + html + '</p>';
	}

	// Show feedback on action buttons
	function showActionFeedback(message) {
		var actionsEl = document.getElementById('response-actions');
		if (!actionsEl) return;

		// Remove existing feedback
		var existingFeedback = actionsEl.querySelector('.csr-action-feedback');
		if (existingFeedback) {
			existingFeedback.remove();
		}

		// Add new feedback
		var feedback = document.createElement('span');
		feedback.className = 'csr-action-feedback';
		feedback.innerHTML = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>' + escapeHtml(message);
		actionsEl.appendChild(feedback);

		// Remove after delay
		setTimeout(function() {
			if (feedback.parentNode) {
				feedback.remove();
			}
		}, 3000);
	}

	// Reset query UI for new query
	function resetQueryUI() {
		var queryInput = document.getElementById('query-input');
		var responseContainer = document.getElementById('response-container');
		var responseText = document.getElementById('response-text');
		var sourcesContainer = document.getElementById('sources-container');
		var actionsEl = document.getElementById('response-actions');

		// Clear input
		if (queryInput) {
			queryInput.value = '';
			queryInput.style.height = 'auto';
			queryInput.focus();
		}

		// Reset state
		state.lastResponse = '';
		state.lastSources = [];
		state.currentQuery = null;

		// Hide response
		if (responseContainer) responseContainer.classList.add('hidden');
		if (responseText) {
			responseText.textContent = '';
			responseText.classList.remove('typing');
		}
		if (sourcesContainer) sourcesContainer.classList.add('hidden');
		if (actionsEl) actionsEl.classList.add('hidden');

		hideQueryError();
		updateQueryButton();
	}

	// Expose CSRWriter API for future phases
	window.CSRWriter = {
		getState: function() {
			return Object.assign({}, state);
		},
		
		getSessionId: function() {
			return state.sessionId;
		},

		isReady: function() {
			return state.isIndexed && state.sessionId;
		},

		reset: resetUpload,

		// Phase 5: Query methods
		getLastResponse: function() {
			return state.lastResponse;
		},

		getLastSources: function() {
			return state.lastSources ? state.lastSources.slice() : [];
		},

		submitQuery: submitQuery,
		resetQuery: resetQueryUI,

		// Phase 6: Document insertion methods
		insertAtCursor: insertAtCursor,
		replaceSelection: replaceSelection,
		copyResponse: copyResponseToClipboard
	};

})(window);
