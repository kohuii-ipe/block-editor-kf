/**
 * ========================================
 * Block Editor KF - Main Script
 * ========================================
 *
 * 目次 (Table of Contents):
 * 1. DOM要素のキャッシュ
 * 2. グローバル変数とステート管理
 * 3. ユーティリティ関数
 * 4. セキュリティ関連（サニタイズ・バリデーション）
 * 5. データ永続化（LocalStorage管理）
 * 6. UI構築関数
 * 7. パターン管理
 * 8. タグボタン管理
 * 9. 入力エリア管理
 * 10. 変換ロジック
 * 11. イベントハンドラ
 * 12. 初期化
 *
 * ========================================
 */

document.addEventListener("DOMContentLoaded", function () {
	// ========================================
	// 1. DOM要素のキャッシュ（パフォーマンス最適化）
	// ========================================
	// すべての頻繁にアクセスされる要素を一度だけ取得して保持
	const DOM = {
		// 出力関連
		convertAllButton: document.getElementById("convertToAllButton"),
		clearOutputButton: document.getElementById("clearOutputButton"),
		copyOutputButton: document.getElementById("copyOutputButton"),
		htmlCodeOutput: document.getElementById("htmlCodeOutput"),

		// コンテナ
		container: document.getElementById("container"),
		clearAllInputsButton: document.getElementById("clearAllInputsButton"),
		insertionPointSelector: document.getElementById("insertionPointSelector"),

		// パターン管理
		patternSelectorContainer: document.getElementById("patternSelectorContainer"),
		addPatternButton: document.getElementById("addPatternButton"),
		renamePatternButton: document.getElementById("renamePatternButton"),
		deletePatternButton: document.getElementById("deletePatternButton"),
		tagButtonsContainer: document.getElementById("tagButtonsContainer"),

		// タグ管理
		addNewTagButton: document.getElementById("addNewTagButton"),
		tagButtonList: document.getElementById("tagButtonList"),
		tagModal: document.getElementById("tagModal"),
		modalTitle: document.getElementById("modalTitle"),
		tagModalForm: document.getElementById("tagModalForm"),
		tagIdInput: document.getElementById("tagIdInput"),
		tagNameInput: document.getElementById("tagNameInput"),
		tagTypeSelector: document.getElementById("tagTypeSelector"),
		tagTemplateInput: document.getElementById("tagTemplateInput"),
		linkItemTemplateSection: document.getElementById("linkItemTemplateSection"),
		linkItemTemplateInput: document.getElementById("linkItemTemplateInput"),
		removeLastBrInput: document.getElementById("removeLastBrInput"),
		closeModalSpan: null, // 後で初期化

		// 書式マッピング管理
		formattingList: document.getElementById("formattingList")
	};

	// closeModalSpanは依存関係があるため後で設定
	DOM.closeModalSpan = DOM.tagModal ? DOM.tagModal.querySelector(".close") : null;

	// 後方互換性のため、個別の定数も維持（既存コードとの互換性）
	const {
		convertAllButton, clearOutputButton, copyOutputButton, container,
		clearAllInputsButton, insertionPointSelector, htmlCodeOutput, patternSelectorContainer,
		addPatternButton, renamePatternButton, deletePatternButton,
		tagButtonsContainer, addNewTagButton, tagButtonList, tagModal,
		modalTitle, tagModalForm, tagIdInput, tagNameInput, tagTypeSelector,
		tagTemplateInput, linkItemTemplateSection, linkItemTemplateInput,
		removeLastBrInput, closeModalSpan, formattingList
	} = DOM;

	// ========================================
	// 2. グローバル変数とステート管理
	// ========================================

	// --- 定数定義 (マジックナンバー・文字列の集約) ---
	const CONSTANTS = {
		// ストレージ関連
		STORAGE_KEYS: {
			PATTERNS: "customTagPatternsV2",
			SETTINGS: "editorSettingsV2",
			INPUT_AREAS: "inputAreasV2"
		},

		// サイズ制限
		MAX_STORAGE_SIZE: 5 * 1024 * 1024, // 5MB
		MAX_FILENAME_LENGTH: 255,

		// タイミング
		DEBOUNCE_DELAY: 500, // ミリ秒
		MODAL_FOCUS_DELAY: 100, // ミリ秒

		// UI文字列
		MESSAGES: {
			PATTERN_DELETE_CONFIRM: (name) => `パターン「${name}」を削除してもよろしいですか？`,
			TAG_DELETE_CONFIRM: (name) => `タグ「${name}」を削除してもよろしいですか？`,
			STORAGE_QUOTA_EXCEEDED: 'ストレージの容量が不足しています。\n一部のデータを削除するか、ブラウザのキャッシュをクリアしてください。',
			PRIVATE_BROWSING: 'プライベートブラウジングモードでは保存できません。\n通常モードで開き直してください。',
			DATA_TOO_LARGE: 'データが大きすぎて保存できません。一部のパターンを削除してください。',
			CORRUPT_DATA: '保存されたデータが破損しています。デフォルトパターンを使用します。',
			NO_COPY_DATA: 'コピーするHTMLコードがありません。',
			COPY_SUCCESS: 'コピーしました！',
			COPY_SUCCESS_FALLBACK: 'コピーしました！(FB)',
		},

		// プレースホルダー
		PLACEHOLDERS: {
			TEXT: '[TEXT]',
			URL: '[URL]',
			TEXT_P: '[TEXT_P]',
			TEXT_P_NUMBERED: (n) => `[TEXT_P_${n}]`,
			TEXT_LIST: '[TEXT_LIST]',
			LINK_LIST: '[LINK_LIST]',
		},

		// スタイル
		STYLES: {
			DRAG_BORDER: '3px solid #4CAF50',
			DRAG_BORDER_TOP: '2px solid #4CAF50',
		},

		// クラス名
		CSS_CLASSES: {
			DRAGGING: 'dragging',
			DRAGGING_BUTTON: 'dragging-button',
			DRAGGING_INPUT_AREA: 'dragging-input-area',
		}
	};

	let areaIndex = 0;
	let maxPatternId = 0;
	let maxTagId = 0;
	let isLoadingInputAreas = false; // 入力エリア読み込み中フラグ

	// ========================================
	// 3. ユーティリティ関数
	// ========================================

	/**
	 * デバウンス関数：連続した呼び出しを遅延させ、最後の呼び出しのみ実行
	 * @param {Function} func - 実行する関数
	 * @param {number} wait - 待機時間（ミリ秒）
	 * @returns {Function} デバウンスされた関数
	 */
	const debounce = (func, wait) => {
		let timeout;
		return function executedFunction(...args) {
			const later = () => {
				clearTimeout(timeout);
				func(...args);
			};
			clearTimeout(timeout);
			timeout = setTimeout(later, wait);
		};
	};

	/**
	 * ドラッグ＆ドロップのイベントリスナーを設定する共通関数
	 * コードの重複を削減するための抽象化
	 * @param {HTMLElement} element - ドラッグ対象の要素
	 * @param {Object} options - オプション設定
	 * @param {string} options.draggingClass - ドラッグ中のCSSクラス名
	 * @param {Function} options.onDrop - ドロップ時のコールバック関数
	 * @param {AbortSignal} options.signal - イベントリスナー管理用のAbortSignal
	 * @param {string} options.borderStyle - ドロップターゲットの境界線スタイル (default: '3px solid #4CAF50')
	 * @param {string} options.cursorGrabbing - ドラッグ中のカーソルスタイル (default: 'grabbing')
	 * @param {string} options.cursorGrab - 通常時のカーソルスタイル (default: 'grab')
	 */
	const setupDragAndDrop = (element, options) => {
		const {
			draggingClass,
			onDrop,
			signal,
			borderStyle = '3px solid #4CAF50',
			cursorGrabbing = 'grabbing',
			cursorGrab = 'grab'
		} = options;

		// dragstart
		element.addEventListener('dragstart', (e) => {
			e.dataTransfer.effectAllowed = 'move';
			e.dataTransfer.setData('text/html', e.target.innerHTML);
			element.classList.add(draggingClass);
			if (element.style) element.style.cursor = cursorGrabbing;
		}, { signal });

		// dragend
		element.addEventListener('dragend', (e) => {
			element.classList.remove(draggingClass);
			if (element.style) element.style.cursor = cursorGrab;
		}, { signal });

		// dragover
		element.addEventListener('dragover', (e) => {
			e.preventDefault();
			e.dataTransfer.dropEffect = 'move';

			const draggingElement = document.querySelector(`.${draggingClass}`);
			if (draggingElement && draggingElement !== element) {
				element.style.borderLeft = borderStyle;
			}
		}, { signal });

		// dragleave
		element.addEventListener('dragleave', (e) => {
			element.style.borderLeft = '';
		}, { signal });

		// drop
		element.addEventListener('drop', (e) => {
			e.preventDefault();
			element.style.borderLeft = '';

			const draggingElement = document.querySelector(`.${draggingClass}`);
			if (!draggingElement || draggingElement === element) return;

			if (onDrop) {
				onDrop(draggingElement, element);
			}
		}, { signal });
	};

	// ========================================
	// 4. セキュリティ関連（サニタイズ・バリデーション）
	// ========================================

	/**
	 * テンプレート文字列を検証する関数
	 * @param {string} template - 検証するテンプレート
	 * @param {string} tagType - タグタイプ (multi, p-list, link-list など)
	 * @returns {Object} { valid: boolean, error: string }
	 */
	const validateTemplate = (template, tagType) => {
		// 空のテンプレートチェック（静的モードも含む）
		if (!template || typeof template !== 'string' || template.trim() === '') {
			return { valid: false, error: 'テンプレートが空です。雛形を入力してください。' };
		}

		// 静的モードの場合、プレースホルダーチェックや危険パターンチェックは不要
		if (tagType === 'static') {
			return { valid: true, error: '' };
		}

		// 必須プレースホルダーのチェック
		const requiredPlaceholders = {
			'multi': ['[TEXT]'],
			'single': ['[TEXT]'],
			'list': ['[TEXT]'],
			'link': ['[TEXT]', '[URL]'],
			'p-list': ['[TEXT_LIST]'],
			'link-list': ['[LINK_LIST]'],
			'static': [] // 静的モードはプレースホルダー不要
		};

		const required = requiredPlaceholders[tagType] || [];
		for (const placeholder of required) {
			if (!template.includes(placeholder)) {
				return {
					valid: false,
					error: `テンプレートに必須のプレースホルダー「${placeholder}」が含まれていません。`
				};
			}
		}

		// 危険なパターンのチェック
		const dangerousPatterns = [
			{ pattern: /<script[\s\S]*?>[\s\S]*?<\/script>/gi, name: 'scriptタグ' },
			{ pattern: /javascript:/gi, name: 'javascriptプロトコル' },
			{ pattern: /on\w+\s*=/gi, name: 'イベントハンドラ (onclick等)' },
			{ pattern: /<iframe/gi, name: 'iframeタグ' },
			{ pattern: /<embed/gi, name: 'embedタグ' },
			{ pattern: /<object/gi, name: 'objectタグ' }
		];

		for (const { pattern, name } of dangerousPatterns) {
			if (pattern.test(template)) {
				return {
					valid: false,
					error: `テンプレートに危険な要素（${name}）が含まれています。`
				};
			}
		}

		// 基本的なHTMLバランスチェック（開始タグと終了タグの数）
		const openTags = (template.match(/<[^/][^>]*>/g) || [])
			.filter(tag => !tag.match(/<(br|hr|img|input|meta|link)\b/i)); // 自己閉じタグを除外
		const closeTags = template.match(/<\/[^>]+>/g) || [];

		// プレースホルダーを除いた場合の開始/終了タグの数をチェック
		const templateWithoutPlaceholders = template.replace(/\[[A-Z_0-9]+\]/g, '');
		const openCount = (templateWithoutPlaceholders.match(/<[^/][^>]*>/g) || [])
			.filter(tag => !tag.match(/<(br|hr|img|input|meta|link)\b/i)).length;
		const closeCount = (templateWithoutPlaceholders.match(/<\/[^>]+>/g) || []).length;

		// 警告レベル: タグのバランスが取れていない場合
		if (openCount !== closeCount) {
			// これは警告のみで、エラーとはしない（柔軟性のため）
			console.warn('テンプレートのHTMLタグのバランスが取れていない可能性があります:', template);
		}

		return { valid: true, error: '' };
	};

	/**
	 * HTMLを安全にサニタイズする関数
	 * 許可されたタグと属性のみを保持し、危険なコンテンツを除去します
	 */
	const sanitizeHTML = (html) => {
		if (!html || typeof html !== 'string') return '';

		// 許可するタグと属性のホワイトリスト
		const allowedTags = ['p', 'div', 'span', 'strong', 'b', 'em', 'i', 'u', 'mark', 'a', 'br', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'code', 'pre'];
		const allowedAttributes = {
			'a': ['href', 'title'],
			'*': ['class', 'id', 'data-highlight', 'data-tag-id', 'data-placeholder']
		};

		// 危険なプロトコルをブロック
		const dangerousProtocols = /^(javascript|data|vbscript):/i;

		const tempDiv = document.createElement('div');
		tempDiv.innerHTML = html;

		const sanitizeNode = (node) => {
			// テキストノードはそのまま返す
			if (node.nodeType === Node.TEXT_NODE) {
				return node.cloneNode(false);
			}

			// 要素ノードのみ処理
			if (node.nodeType !== Node.ELEMENT_NODE) {
				return null;
			}

			const tagName = node.tagName.toLowerCase();

			// 許可されていないタグは子ノードのみ保持
			if (!allowedTags.includes(tagName)) {
				const fragment = document.createDocumentFragment();
				for (let child of node.childNodes) {
					const sanitizedChild = sanitizeNode(child);
					if (sanitizedChild) {
						fragment.appendChild(sanitizedChild);
					}
				}
				return fragment;
			}

			// 許可されたタグの場合、新しい要素を作成
			const newElement = document.createElement(tagName);

			// 属性をフィルタリング
			const tagAllowedAttrs = allowedAttributes[tagName] || [];
			const globalAllowedAttrs = allowedAttributes['*'] || [];
			const combinedAllowedAttrs = [...tagAllowedAttrs, ...globalAllowedAttrs];

			for (let attr of node.attributes) {
				const attrName = attr.name.toLowerCase();

				// 許可された属性のみコピー
				if (combinedAllowedAttrs.includes(attrName)) {
					let attrValue = attr.value;

					// hrefの場合、危険なプロトコルをチェック
					if (attrName === 'href' && dangerousProtocols.test(attrValue)) {
						continue; // 危険なhrefはスキップ
					}

					// on* イベントハンドラを除外
					if (attrName.startsWith('on')) {
						continue;
					}

					newElement.setAttribute(attrName, attrValue);
				}
			}

			// 子ノードを再帰的にサニタイズ
			for (let child of node.childNodes) {
				const sanitizedChild = sanitizeNode(child);
				if (sanitizedChild) {
					newElement.appendChild(sanitizedChild);
				}
			}

			return newElement;
		};

		const sanitizedFragment = document.createDocumentFragment();
		for (let child of tempDiv.childNodes) {
			const sanitizedChild = sanitizeNode(child);
			if (sanitizedChild) {
				sanitizedFragment.appendChild(sanitizedChild);
			}
		}

		const resultDiv = document.createElement('div');
		resultDiv.appendChild(sanitizedFragment);
		return resultDiv.innerHTML;
	};

	// ========================================
	// 5. データ永続化（LocalStorage管理）
	// ========================================

	// --- デフォルトデータ定義 ---

	// デフォルトの書式マッピング設定（Word貼り付け用）
	const defaultFormattingMap = {
		bold: { template: "<strong>[TEXT]</strong>", displayName: "太字 (Bold)" },
		highlight: { template: "<mark>[TEXT]</mark>", displayName: "ハイライト (Highlight)" },
	};

	const defaultTagButtons = [
		// id, name, template, tagType
		// tagType: multi (マルチラインモード), p-list (段落+リストモード), link-list (リンクリストモード)
		{ id: "p", name: "pタグ", template: "<p>[TEXT]</p>", tagType: "multi" },
		{ id: "h2", name: "h2タグ", template: "<h2>[TEXT]</h2>", tagType: "multi" },
		{ id: "box", name: "枠", template: '<div class="box">[TEXT]</div>', tagType: "multi" },
		{
			id: "box-p-ul",
			name: "箇条書きリスト",
			template: '<div class="box"><p>[TEXT_P_1]</p><p>[TEXT_P_2]</p><ul>\n[TEXT_LIST]\n</ul></div>',
			tagType: "p-list",
		},
		{
			id: "link-list-sample",
			name: "リンクリスト",
			template: '<div class="link-wrapper">\n<ul>\n[LINK_LIST]\n</ul>\n</div>',
			tagType: "link-list",
			linkItemTemplate: '<li><a href="[URL]">[TEXT]</a></li>',
		},
	];

	const defaultPatterns = {
		pattern1: {
			name: "ソニー (デフォルト)",
			buttons: JSON.parse(JSON.stringify(defaultTagButtons)), // ディープコピー
			formattingMap: JSON.parse(JSON.stringify(defaultFormattingMap)),
			inputAreas: [], // 各パターンごとに入力エリアを保持
		},
		pattern2: {
			name: "カスタム例",
			buttons: [
				{
					id: "p-red",
					name: "赤文字P",
					template: '<p style="color:red;">[TEXT]</p>',
					tagType: "single",
				},
				{
					id: "h2-border",
					name: "下線H2",
					template: '<h2 style="border-bottom: 2px solid black;">[TEXT]</h2>',
					tagType: "single",
				},
				{
					id: "note",
					name: "注意枠",
					template:
						'<div class="note" style="border: 1px solid red; padding: 10px;">\n<p style="font-weight: bold;">注意：</p>\n[TEXT]\n</div>',
					tagType: "single",
				},
			],
			formattingMap: JSON.parse(JSON.stringify(defaultFormattingMap)),
			inputAreas: [], // 各パターンごとに入力エリアを保持
		},
	};

	let patterns = {}; // すべてのパターン設定を保持するグローバルオブジェクト
	let formattingMap = {};

	const saveAllPatterns = () => {
		try {
			const selectedId = getSelectedPatternId();

			// データサイズをチェック
			const dataToSave = JSON.stringify(patterns);
			const dataSize = new Blob([dataToSave]).size;

			if (dataSize > CONSTANTS.MAX_STORAGE_SIZE) {
				console.error('保存データが大きすぎます:', dataSize, 'bytes');
				alert(CONSTANTS.MESSAGES.DATA_TOO_LARGE);
				return;
			}

			localStorage.setItem(CONSTANTS.STORAGE_KEYS.PATTERNS, dataToSave);

			const settings = {
				selectedPattern: selectedId,
			};
			localStorage.setItem(CONSTANTS.STORAGE_KEYS.SETTINGS, JSON.stringify(settings));
		} catch (error) {
			console.error('パターンの保存に失敗しました:', error);

			// QuotaExceededError の場合
			if (error.name === 'QuotaExceededError' || error.code === 22) {
				alert(CONSTANTS.MESSAGES.STORAGE_QUOTA_EXCEEDED);
			} else if (error.name === 'SecurityError') {
				alert(CONSTANTS.MESSAGES.PRIVATE_BROWSING);
			} else {
				alert('データの保存に失敗しました。\nエラー: ' + error.message);
			}
		}
	};

	const saveFormattingMap = () => {
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];

		if (currentPattern) {
			currentPattern.formattingMap = formattingMap;
			saveAllPatterns();
		}
	};

	const loadFormattingMap = () => {
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];

		if (currentPattern && currentPattern.formattingMap) {
			formattingMap = currentPattern.formattingMap;
			// 古いパターンから italic と underline を削除
			delete formattingMap.italic;
			delete formattingMap.underline;
		} else {
			// フォールバック：パターンに書式マッピングがない場合はデフォルトを設定
			formattingMap = JSON.parse(JSON.stringify(defaultFormattingMap));
			if (currentPattern) {
				currentPattern.formattingMap = formattingMap;
			}
		}
	};

	// 入力エリアの保存と読み込み
	const saveInputAreas = () => {
		// 読み込み中は保存しない
		if (isLoadingInputAreas) return;

		try {
			const areas = [];
			const allGroups = container.querySelectorAll(".input-area-group");

			allGroups.forEach((group) => {
				const textarea = group.querySelector(".input-text-area");
				if (textarea) {
					areas.push({
						tagId: textarea.getAttribute("data-tag-id"),
						content: textarea.innerHTML,
						groupId: group.id
					});
				}
			});

			// 現在のパターンに入力エリアを保存
			const patternId = getSelectedPatternId();
			const currentPattern = patterns[patternId];
			if (currentPattern) {
				currentPattern.inputAreas = areas;
			}

			// パターン全体を保存
			const dataToSave = JSON.stringify(patterns);
			const dataSize = new Blob([dataToSave]).size;

			if (dataSize > CONSTANTS.MAX_STORAGE_SIZE) {
				console.warn('パターンデータが大きすぎます:', dataSize, 'bytes');
				// 警告のみで、保存を試みる（古いデータを保持するため）
			}

			localStorage.setItem(CONSTANTS.STORAGE_KEYS.PATTERNS, dataToSave);
		} catch (error) {
			console.error('入力エリアの保存に失敗しました:', error);

			// エラーは通知するが、作業は続行可能
			if (error.name === 'QuotaExceededError' || error.code === 22) {
				console.warn('ストレージの容量が不足しています。入力内容が保存されない可能性があります。');
			} else if (error.name === 'SecurityError') {
				console.warn('プライベートブラウジングモードでは保存できません。');
			}
		}
	};

	// デバウンスされた保存関数（入力中のパフォーマンス向上）
	const debouncedSaveInputAreas = debounce(saveInputAreas, CONSTANTS.DEBOUNCE_DELAY);

	const loadInputAreas = () => {
		try {
			// 現在のパターンから入力エリアを読み込む
			const patternId = getSelectedPatternId();
			const currentPattern = patterns[patternId];

			if (!currentPattern || !currentPattern.inputAreas) {
				return;
			}

			try {
				isLoadingInputAreas = true; // 読み込み開始
				const areas = currentPattern.inputAreas;

				if (!Array.isArray(areas)) {
					console.error('入力エリアのデータ形式が不正です');
					return;
				}

				// 既存の入力エリアをクリア
				while (container.firstChild) {
					const child = container.firstChild;
					if (child._cleanup) {
						child._cleanup();
					}
					container.removeChild(child);
				}

				// 入力エリアを復元
				areas.forEach((area) => {
					try {
						const tagInfo = getCustomTagInfo(area.tagId);
						if (tagInfo) {
							createTextarea(area.tagId, tagInfo.tagType, tagInfo.name, area.content);
						} else {
							console.warn(`タグID "${area.tagId}" が見つかりません。スキップします。`);
						}
					} catch (areaError) {
						console.error('個別の入力エリア復元に失敗しました:', areaError);
						// 一部のエリアが失敗しても続行
					}
				});
			} catch (parseError) {
				console.error("入力エリアの解析に失敗しました:", parseError);
				alert('入力エリアのデータが破損しています。新規作成してください。');
			} finally {
				isLoadingInputAreas = false; // 読み込み終了
			}
		} catch (error) {
			console.error("入力エリアの読み込みに失敗しました:", error);
			isLoadingInputAreas = false;
		}
	};

	// ========================================
	// 6. UI構築関数
	// ========================================

	/**
	 * 書式マッピング管理UIを構築
	 */
	const buildFormattingUI = () => {
		if (!formattingList) return;

		formattingList.innerHTML = "";

		const types = Object.keys(formattingMap);

		types.forEach((type) => {
			// italic と underline は表示しない
			if (type === 'italic' || type === 'underline') {
				return;
			}

			const formatting = formattingMap[type];
			const item = document.createElement("div");
			item.className = "formatting-item";

			const label = document.createElement("label");
			label.textContent = formatting.displayName;

			const input = document.createElement("input");
			input.type = "text";
			input.value = formatting.template || formatting.tag || ""; // 後方互換性のため tag もチェック
			input.placeholder = "テンプレートを入力 (例: <strong>[TEXT]</strong>)";

			input.addEventListener("input", (e) => {
				const newTemplate = e.target.value.trim();

				// 空のテンプレートは許可（デフォルトに戻すため）
				if (newTemplate) {
					// 入力検証: [TEXT]プレースホルダーが含まれているかチェック
					if (!newTemplate.includes('[TEXT]')) {
						console.warn(`書式マッピング「${formatting.displayName}」に [TEXT] プレースホルダーが含まれていません`);
						// 警告のみで保存は続行（柔軟性のため）
					}

					// 危険なパターンをチェック
					const dangerousPatterns = [
						/<script/gi,
						/javascript:/gi,
						/on\w+\s*=/gi
					];

					const hasDangerousContent = dangerousPatterns.some(pattern => pattern.test(newTemplate));
					if (hasDangerousContent) {
						alert(`書式マッピング「${formatting.displayName}」に危険な要素が含まれています。保存できません。`);
						e.target.value = formattingMap[type].template || formatting.template || '';
						return;
					}
				}

				formattingMap[type].template = newTemplate;
				// 後方互換性のため tag プロパティも削除
				delete formattingMap[type].tag;
				saveFormattingMap();
			});

			item.appendChild(label);
			item.appendChild(input);
			formattingList.appendChild(item);
		});
	};

	const loadAllPatterns = () => {
		try {
			const savedPatterns = localStorage.getItem(CONSTANTS.STORAGE_KEYS.PATTERNS);
			if (savedPatterns) {
				try {
					patterns = JSON.parse(savedPatterns);
					if (Object.keys(patterns).length === 0) {
						console.warn('保存されたパターンが空です。デフォルトパターンを使用します。');
						patterns = JSON.parse(JSON.stringify(defaultPatterns));
					}
				} catch (parseError) {
					console.error('保存データの解析に失敗しました:', parseError);
					alert(CONSTANTS.MESSAGES.CORRUPT_DATA);
					patterns = JSON.parse(JSON.stringify(defaultPatterns));
				}
			} else {
				patterns = JSON.parse(JSON.stringify(defaultPatterns));
			}
		} catch (error) {
			console.error('パターンの読み込みに失敗しました:', error);
			alert('データの読み込みに失敗しました。デフォルトパターンを使用します。\nエラー: ' + error.message);
			patterns = JSON.parse(JSON.stringify(defaultPatterns));
		}

		// IDの最大値を更新（ローカルストレージとデフォルトパターンの両方に対応）
		Object.keys(patterns).forEach((id) => {
			const num = parseInt(id.replace("pattern", ""));
			if (!isNaN(num)) {
				maxPatternId = Math.max(maxPatternId, num);
			}
			// タグIDの最大値も更新
			patterns[id].buttons.forEach((button) => {
				if (button.id.startsWith("custom")) {
					const tagNum = parseInt(button.id.replace("custom", ""));
					if (!isNaN(tagNum)) {
						maxTagId = Math.max(maxTagId, tagNum);
					}
				}
			});
			// 後方互換性：書式マッピングがない古いパターンにデフォルト書式マッピングを追加
			if (!patterns[id].formattingMap) {
				patterns[id].formattingMap = JSON.parse(JSON.stringify(defaultFormattingMap));
			}
			// 後方互換性：古い tag 形式を新しい template 形式に変換
			if (patterns[id].formattingMap) {
				Object.keys(patterns[id].formattingMap).forEach((type) => {
					const formatting = patterns[id].formattingMap[type];
					if (formatting.tag && !formatting.template) {
						// tag を template に変換
						formatting.template = `<${formatting.tag}>[TEXT]</${formatting.tag}>`;
						delete formatting.tag;
					}
				});
				// 古いパターンから italic と underline を削除
				delete patterns[id].formattingMap.italic;
				delete patterns[id].formattingMap.underline;
			}
			// 後方互換性：入力エリアがない古いパターンに空の配列を追加
			if (!patterns[id].inputAreas) {
				patterns[id].inputAreas = [];
			}
		});

		let initialSelectedPattern = "pattern1";
		try {
			const savedSettings = localStorage.getItem(CONSTANTS.STORAGE_KEYS.SETTINGS);
			if (savedSettings) {
				try {
					const settings = JSON.parse(savedSettings);
					initialSelectedPattern = settings.selectedPattern || "pattern1";
				} catch (parseError) {
					console.warn('設定の解析に失敗しました:', parseError);
					initialSelectedPattern = "pattern1";
				}
			}
		} catch (error) {
			console.warn('設定の読み込みに失敗しました:', error);
			initialSelectedPattern = "pattern1";
		}

		// 後方互換性：古いグローバル入力エリアを現在のパターンに移行
		try {
			const oldInputAreas = localStorage.getItem(CONSTANTS.STORAGE_KEYS.INPUT_AREAS);
			if (oldInputAreas) {
				const areas = JSON.parse(oldInputAreas);
				if (Array.isArray(areas) && areas.length > 0) {
					// 選択されたパターンに入力エリアがない場合のみ移行
					const currentPattern = patterns[initialSelectedPattern];
					if (currentPattern && (!currentPattern.inputAreas || currentPattern.inputAreas.length === 0)) {
						currentPattern.inputAreas = areas;
						console.log('古い入力エリアを現在のパターンに移行しました');
					}
				}
				// 古いストレージキーを削除
				localStorage.removeItem(CONSTANTS.STORAGE_KEYS.INPUT_AREAS);
			}
		} catch (migrationError) {
			console.warn('入力エリアの移行に失敗しました:', migrationError);
		}

		rebuildPatternUI(initialSelectedPattern);
		rebuildTagButtons(); // タグボタン群も初期化時に生成
		updateInsertionPoints(); // 入力エリアのリストを更新
	};

	// ========================================
	// 7. パターン管理
	// ========================================

	/**
	 * 現在選択されているパターンIDを取得
	 * @returns {string} パターンID
	 */
	const getSelectedPatternId = () => {
		const checked = document.querySelector('input[name="conversionPattern"]:checked');
		return checked ? checked.value : Object.keys(patterns)[0];
	};

	/**
	 * パターン選択UIを再構築する関数
	 * パフォーマンス最適化：DocumentFragmentを使用
	 */
	const rebuildPatternUI = (initialSelectedId) => {
		patternSelectorContainer.innerHTML = "";

		let firstId = null;
		let selectedId = initialSelectedId;

		// DocumentFragmentを使用してDOM操作を最適化
		const fragment = document.createDocumentFragment();

		Object.keys(patterns).forEach((id, index) => {
			if (index === 0) firstId = id;
			const pattern = patterns[id];

			// 1. ラジオボタンの生成
			const label = document.createElement("label");
			const input = document.createElement("input");
			input.type = "radio";
			input.name = "conversionPattern";
			input.id = id;
			input.value = id;
			input.addEventListener("change", () => {
				saveInputAreas(); // 現在のパターンの入力エリアを保存してから切り替え
				rebuildTagButtons(); // パターン切り替え時、ボタンを再生成
				rebuildTagButtonList(); // タグボタン管理一覧を更新
				loadFormattingMap(); // 書式マッピングを再読み込み
				buildFormattingUI(); // 書式マッピングUIを再構築
				loadInputAreas(); // 新しいパターンの入力エリアを読み込み
				updateInsertionPoints(); // 挿入位置セレクタを更新
				saveAllPatterns();
			});

			const nameSpan = document.createElement("span");
			nameSpan.id = `label-name-${id}`;
			nameSpan.textContent = pattern.name;

			label.appendChild(input);
			label.appendChild(nameSpan);
			fragment.appendChild(label);
		});

		// 一度にすべてのパターンを追加（リフローを最小化）
		patternSelectorContainer.appendChild(fragment);

		// 3. 初期選択と表示状態の設定
		const finalSelectedId = patterns[selectedId] ? selectedId : firstId;
		if (finalSelectedId) {
			const selectedRadio = document.getElementById(finalSelectedId);
			if (selectedRadio) {
				selectedRadio.checked = true;
			}
		}

		updateDeleteButtonState();
	};

	// 削除ボタンの disabled 状態を更新
	const updateDeleteButtonState = () => {
		const count = Object.keys(patterns).length;
		deletePatternButton.disabled = count <= 1;
	};

	// --- パターン管理ボタンの機能 ---

	addPatternButton.addEventListener("click", () => {
		maxPatternId++;
		const newId = `pattern${maxPatternId}`;

		// ★★★ 修正箇所 ★★★
		// 既存のパターン名から「新しいパターン X」の番号を抽出し、最小の空き番号を見つける
		const existingNumbers = Object.values(patterns)
			.map(p => {
				const match = p.name.match(/^新しいパターン\s+(\d+)$/);
				return match ? parseInt(match[1]) : null;
			})
			.filter(n => n !== null);

		// 最小の空き番号を見つける（1から順番にチェック）
		let nameNumber = 1;
		while (existingNumbers.includes(nameNumber)) {
			nameNumber++;
		}

		const newName = `新しいパターン ${nameNumber}`;

		// 新しいパターンはデフォルト設定で開始
		patterns[newId] = {
			name: newName,
			// デフォルトのタグボタンを使用
			buttons: JSON.parse(JSON.stringify(defaultTagButtons)),
			// デフォルトの書式マッピングを使用
			formattingMap: JSON.parse(JSON.stringify(defaultFormattingMap)),
			// 入力エリアは空で開始（パターンごとに独立）
			inputAreas: [],
		};

		rebuildPatternUI(newId);
		rebuildTagButtons();
		rebuildTagButtonList();
		loadFormattingMap(); // 新しいパターンの書式マッピングを読み込み
		buildFormattingUI(); // 書式マッピングUIを再構築
		loadInputAreas(); // 新しいパターンの入力エリアを読み込み（空）
		updateInsertionPoints(); // 挿入位置セレクタを更新
		saveAllPatterns();
	});

	renamePatternButton.addEventListener("click", () => {
		const selectedId = getSelectedPatternId();
		const currentPattern = patterns[selectedId];
		if (!currentPattern) return;

		const newName = prompt("新しいパターン名を入力してください:", currentPattern.name);
		if (newName !== null && newName.trim() !== "") {
			currentPattern.name = newName.trim();
			document.getElementById(`label-name-${selectedId}`).textContent = newName.trim();
			saveAllPatterns();
		}
	});

	deletePatternButton.addEventListener("click", () => {
		const selectedId = getSelectedPatternId();
		if (Object.keys(patterns).length <= 1) {
			console.warn("パターンは最低1つ必要です。");
			return;
		}

		const currentPattern = patterns[selectedId];
		if (confirm(CONSTANTS.MESSAGES.PATTERN_DELETE_CONFIRM(currentPattern.name))) {
			console.log(`パターン「${currentPattern.name}」を削除しました。`);
			delete patterns[selectedId];

			const firstId = Object.keys(patterns)[0];
			rebuildPatternUI(firstId);
			rebuildTagButtons();
			rebuildTagButtonList();
			loadFormattingMap(); // 新しいパターンの書式マッピングを読み込み
			buildFormattingUI(); // 書式マッピングUIを再構築
			loadInputAreas(); // 削除後、残ったパターンの入力エリアを読み込み
			updateInsertionPoints(); // 挿入位置セレクタを更新
			saveAllPatterns();
		}
	});

	// ========================================
	// 8. タグボタン管理
	// ========================================

	/**
	 * タグ追加/編集モーダルを開く
	 * @param {Object|null} button - 編集する場合はボタンオブジェクト、新規追加の場合はnull
	 */
	const openTagModal = (button = null) => {
		tagModal.style.display = "block";
		tagModal.setAttribute('aria-hidden', 'false');

		// フォーカスをモーダル内の最初の入力欄に移動（アクセシビリティ向上）
		setTimeout(() => {
			tagNameInput.focus();
		}, CONSTANTS.MODAL_FOCUS_DELAY);

		if (button) {
			// 編集モード
			modalTitle.textContent = "タグボタンの編集";
			tagIdInput.value = button.id;
			tagNameInput.value = button.name;
			tagTypeSelector.value = button.tagType;
			tagTemplateInput.value = button.template;
			linkItemTemplateInput.value = button.linkItemTemplate || '<li class="rtoc-item"><a href="[URL]">[TEXT]</a></li>';
			removeLastBrInput.checked = button.removeLastBr || false;
		} else {
			// 新規追加モード
			modalTitle.textContent = "新しいタグボタンの追加";
			tagIdInput.value = "";
			tagNameInput.value = "";
			tagTypeSelector.value = "single";
			tagTemplateInput.value = "";
			linkItemTemplateInput.value = '<li class="rtoc-item"><a href="[URL]">[TEXT]</a></li>';
			removeLastBrInput.checked = false;
		}

		// モーダル表示時にプレースホルダーを更新
		updateTemplatePlaceholder();
	};

	// モーダルの閉じる処理
	closeModalSpan.onclick = function () {
		tagModal.style.display = "none";
		tagModal.setAttribute('aria-hidden', 'true');
	};

	// キーボードでモーダルを閉じる（Enter/Spaceキー）
	closeModalSpan.addEventListener('keydown', function(e) {
		if (e.key === 'Enter' || e.key === ' ') {
			e.preventDefault();
			tagModal.style.display = "none";
			tagModal.setAttribute('aria-hidden', 'true');
		}
	});

	// 「新しいタグボタンを追加」ボタン
	addNewTagButton.addEventListener("click", () => openTagModal());

	// 雛形テキストエリアのプレースホルダーを更新する関数
	const updateTemplatePlaceholder = () => {
		const selectedType = tagTypeSelector.value;
		let placeholder = "";

		// link-list用のフィールドの表示/非表示を切り替え
		if (selectedType === "link-list") {
			linkItemTemplateSection.style.display = "block";
		} else {
			linkItemTemplateSection.style.display = "none";
		}

		switch (selectedType) {
			case "multi":
				placeholder = '例: <p class="item">[TEXT]</p>\n\n【マルチラインモード】\n※各行がそれぞれ個別のタグとして出力されます\n※[TEXT]が各行のテキストに置き換わります';
				break;
			case "p-list":
				placeholder = '例:\n<div>\n<p>[TEXT_P_1]</p>\n<p>[TEXT_P_2]</p>\n<ul>\n[TEXT_LIST]\n</ul>\n</div>\n\n【段落+リストモード】\n※通常の行: [TEXT_P_1], [TEXT_P_2]として段落に変換\n※「-」または「*」で始まる行: [TEXT_LIST]として箇条書きに変換';
				break;
			case "link-list":
				placeholder = '例:\n<div class="wrapper">\n<ul>\n[LINK_LIST]\n</ul>\n</div>\n\n【リンクリストモード】\n※[LINK_LIST]の位置に複数のリンク項目が挿入されます\n※下の「リンク項目の雛形」で各リンクの形式を指定します';
				break;
			case "static":
				placeholder = '【静的モード】\n※テンプレート欄に入力したHTMLがそのまま出力されます\n※入力エリアは無効化されます（入力不可）\n※固定のHTMLブロックや定型文を出力するのに便利です\n\n例:\n<div class="alert">\n<p>このメッセージは固定です</p>\n</div>';
				break;
			default:
				placeholder = "[TEXT] を使って雛形を入力";
		}
		tagTemplateInput.placeholder = placeholder;
	};

	// タグタイプセレクタが変更されたら、プレースホルダーを更新
	tagTypeSelector.addEventListener("change", updateTemplatePlaceholder);

	// タグ保存 (追加/編集) 処理
	tagModalForm.addEventListener("submit", (e) => {
		e.preventDefault();

		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];
		if (!currentPattern) return;

		const existingId = tagIdInput.value;
		const newName = tagNameInput.value.trim();
		const newType = tagTypeSelector.value;
		const newTemplate = tagTemplateInput.value;
		const newLinkItemTemplate = linkItemTemplateInput.value;
		const newRemoveLastBr = removeLastBrInput.checked;

		// 入力検証: ボタン名が空でないかチェック
		if (!newName) {
			alert('ボタン名を入力してください。');
			return;
		}

		// 入力検証: テンプレートの検証
		const templateValidation = validateTemplate(newTemplate, newType);
		if (!templateValidation.valid) {
			alert('テンプレートエラー:\n' + templateValidation.error);
			return;
		}

		// link-listモードの場合、linkItemTemplateも検証
		if (newType === 'link-list' && newLinkItemTemplate) {
			const linkItemValidation = validateTemplate(newLinkItemTemplate, 'link');
			if (!linkItemValidation.valid) {
				alert('リンク項目の雛形エラー:\n' + linkItemValidation.error);
				return;
			}
		}

		if (existingId) {
			// 編集
			const button = currentPattern.buttons.find((b) => b.id === existingId);
			if (button) {
				button.name = newName;
				button.tagType = newType;
				button.template = newTemplate;
				button.linkItemTemplate = newLinkItemTemplate;
				button.removeLastBr = newRemoveLastBr;
			}
		} else {
			// 新規追加
			maxTagId++;
			const newId = `custom${maxTagId}`;
			currentPattern.buttons.push({
				id: newId,
				name: newName,
				template: newTemplate,
				tagType: newType,
				linkItemTemplate: newLinkItemTemplate,
				removeLastBr: newRemoveLastBr,
			});
		}

		tagModal.style.display = "none";
		rebuildTagButtons();
		rebuildTagButtonList();
		saveAllPatterns();
	});

	// タグボタン一覧の構築
	/**
	 * タグボタン管理リストを再構築する関数
	 * パフォーマンス最適化：DocumentFragmentを使用してバッチDOM操作を実行
	 */
	const rebuildTagButtonList = () => {
		tagButtonList.innerHTML = "";
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];

		if (!currentPattern || currentPattern.buttons.length === 0) {
			tagButtonList.textContent = "タグボタンがありません。";
			return;
		}

		// DocumentFragmentを使用してDOM操作を最適化
		const fragment = document.createDocumentFragment();

		currentPattern.buttons.forEach((button, index) => {
			const item = document.createElement("div");
			item.className = "tag-item";

			const infoSpan = document.createElement("span");
			infoSpan.textContent = `${button.name} (${button.tagType})`;

			// ドラッグ可能にする
			item.draggable = true;
			item.dataset.index = index;

			// ドラッグハンドル
			const dragHandle = document.createElement("span");
			dragHandle.textContent = "⋮⋮";
			dragHandle.style.cursor = "grab";
			dragHandle.style.marginRight = "8px";
			dragHandle.style.color = "#999";
			dragHandle.title = "ドラッグして順序を変更";

			// AbortControllerでイベントリスナーを管理（メモリリーク対策）
			const abortController = new AbortController();
			const signal = abortController.signal;

			// ドラッグイベント（signalオプションでメモリリーク防止）
			item.addEventListener("dragstart", (e) => {
				e.dataTransfer.effectAllowed = "move";
				e.dataTransfer.setData("text/html", e.target.innerHTML);
				item.classList.add("dragging");
			}, { signal });

			item.addEventListener("dragend", (e) => {
				item.classList.remove("dragging");
			}, { signal });

			item.addEventListener("dragover", (e) => {
				e.preventDefault();
				e.dataTransfer.dropEffect = "move";

				const draggingItem = document.querySelector(".dragging");
				if (draggingItem && draggingItem !== item) {
					const rect = item.getBoundingClientRect();
					const midpoint = rect.top + rect.height / 2;

					if (e.clientY < midpoint) {
						item.style.borderTop = "2px solid #4CAF50";
						item.style.borderBottom = "";
					} else {
						item.style.borderTop = "";
						item.style.borderBottom = "2px solid #4CAF50";
					}
				}
			}, { signal });

			item.addEventListener("dragleave", (e) => {
				item.style.borderTop = "";
				item.style.borderBottom = "";
			}, { signal });

			item.addEventListener("drop", (e) => {
				e.preventDefault();
				item.style.borderTop = "";
				item.style.borderBottom = "";

				const draggingItem = document.querySelector(".dragging");
				if (!draggingItem || draggingItem === item) return;

				const fromIndex = parseInt(draggingItem.dataset.index);
				const toIndex = parseInt(item.dataset.index);

				const [movedButton] = currentPattern.buttons.splice(fromIndex, 1);
				currentPattern.buttons.splice(toIndex, 0, movedButton);

				rebuildTagButtons();
				rebuildTagButtonList();
				saveAllPatterns();
			}, { signal });

			const editButton = document.createElement("button");
			editButton.textContent = "編集";
			editButton.addEventListener("click", () => openTagModal(button), { signal });

			const deleteButton = document.createElement("button");
			deleteButton.textContent = "削除";
			deleteButton.style.color = "red";
			deleteButton.addEventListener("click", () => {
				if (confirm(CONSTANTS.MESSAGES.TAG_DELETE_CONFIRM(button.name))) {
					console.log(`タグ「${button.name}」を削除しました。`);
					currentPattern.buttons = currentPattern.buttons.filter((b) => b.id !== button.id);
					rebuildTagButtons();
					rebuildTagButtonList();
					saveAllPatterns();
				}
			}, { signal });

			item.appendChild(dragHandle);
			item.appendChild(infoSpan);
			item.appendChild(editButton);
			item.appendChild(deleteButton);
			fragment.appendChild(item);
		});

		// 一度にすべてのアイテムを追加（リフローを最小化）
		tagButtonList.appendChild(fragment);
	};

	// --- タグボタン群の動的生成 ---

	/**
	 * タグボタンを再構築する関数
	 * パフォーマンス最適化：DocumentFragmentを使用してバッチDOM操作を実行
	 * 注意：ドラッグ＆ドロップの順序変更やパターン切り替えでは全体の再構築が必要です
	 */
	const rebuildTagButtons = () => {
		tagButtonsContainer.innerHTML = "";
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];
		if (!currentPattern) return;

		// DocumentFragmentを使用してDOM操作を最適化（リフローを1回のみに削減）
		const fragment = document.createDocumentFragment();

		currentPattern.buttons.forEach((button, index) => {
			const btn = document.createElement("button");
			btn.id = `addTagButton-${button.id}`;
			btn.textContent = button.name;
			btn.draggable = true;
			btn.dataset.index = index;
			btn.style.cursor = "grab";

			// AbortControllerでイベントリスナーを管理（メモリリーク対策）
			const abortController = new AbortController();
			const signal = abortController.signal;

			// クリックイベント
			btn.addEventListener("click", () =>
				createTextarea(button.id, button.tagType, button.name)
			, { signal });

			// ドラッグイベント（signalオプションでメモリリーク防止）
			btn.addEventListener("dragstart", (e) => {
				e.dataTransfer.effectAllowed = "move";
				e.dataTransfer.setData("text/html", e.target.innerHTML);
				btn.classList.add("dragging-button");
				btn.style.cursor = "grabbing";
			}, { signal });

			btn.addEventListener("dragend", (e) => {
				btn.classList.remove("dragging-button");
				btn.style.cursor = "grab";
			}, { signal });

			btn.addEventListener("dragover", (e) => {
				e.preventDefault();
				e.dataTransfer.dropEffect = "move";

				const draggingButton = document.querySelector(".dragging-button");
				if (draggingButton && draggingButton !== btn) {
					btn.style.borderLeft = "3px solid #4CAF50";
				}
			}, { signal });

			btn.addEventListener("dragleave", (e) => {
				btn.style.borderLeft = "";
			}, { signal });

			btn.addEventListener("drop", (e) => {
				e.preventDefault();
				btn.style.borderLeft = "";

				const draggingButton = document.querySelector(".dragging-button");
				if (!draggingButton || draggingButton === btn) return;

				const fromIndex = parseInt(draggingButton.dataset.index);
				const toIndex = parseInt(btn.dataset.index);

				// 配列を並び替え
				const [movedButton] = currentPattern.buttons.splice(fromIndex, 1);
				currentPattern.buttons.splice(toIndex, 0, movedButton);

				rebuildTagButtons();
				rebuildTagButtonList();
				saveAllPatterns();
			}, { signal });

			fragment.appendChild(btn);
		});

		// 一度にすべてのボタンを追加（リフローを最小化）
		tagButtonsContainer.appendChild(fragment);
	};

	// ========================================
	// 9. 入力エリア管理
	// ========================================

	/**
	 * 選択されたパターンIDとタグIDに基づいてカスタムテンプレートとタイプを取得する関数
	 * @param {string} tagId - タグID
	 * @returns {Object|null} タグ情報オブジェクト
	 */
	const getCustomTagInfo = (tagId) => {
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];
		if (!currentPattern) return null;

		return currentPattern.buttons.find((b) => b.id === tagId);
	};

	// contenteditable div から改行を保持したテキストを取得するヘルパー関数
	// formatting タグ（strong, mark）は保持し、block要素は改行に変換
	function getTextWithLineBreaks(htmlContent) {
		// XSS対策: 入力HTMLをサニタイズ
		const sanitizedContent = sanitizeHTML(htmlContent);
		const tempDiv = document.createElement("div");
		tempDiv.innerHTML = sanitizedContent;

		const processNode = (node) => {
			if (node.nodeType === Node.TEXT_NODE) {
				return node.textContent;
			}

			if (node.nodeType === Node.ELEMENT_NODE) {
				const tagName = node.tagName.toLowerCase();

				// br は改行に変換
				if (tagName === "br") {
					return "\n";
				}

				// block 要素は改行で区切る
				const isBlock = ["div", "p"].includes(tagName);
				// formatting タグは保持（リンクも含む）
				const isFormatting = ["strong", "mark", "b", "em", "u", "span", "a"].includes(tagName);

				let result = "";

				// formatting タグの開始タグを追加（属性も含めて）
				if (isFormatting) {
					let openTag = `<${tagName}`;
					// 属性があれば追加（不要な属性は除外）
					if (node.attributes && node.attributes.length > 0) {
						// リンクの場合は href のみ保持
						const allowedAttrs = tagName === "a" ? ["href"] : ["class", "id", "style"];

						for (let i = 0; i < node.attributes.length; i++) {
							const attr = node.attributes[i];
							// data- で始まる属性や許可されていない属性を除外
							if (allowedAttrs.includes(attr.name) && !attr.name.startsWith("data-")) {
								openTag += ` ${attr.name}="${attr.value}"`;
							}
						}
					}
					openTag += `>`;
					result += openTag;
				}

				// 子ノードを処理
				for (let child of node.childNodes) {
					result += processNode(child);
				}

				// formatting タグの終了タグを追加
				if (isFormatting) {
					result += `</${tagName}>`;
				}

				// block 要素の後に改行を追加
				if (isBlock) {
					result += "\n";
				}

				return result;
			}

			return "";
		};

		let text = "";
		for (let child of tempDiv.childNodes) {
			text += processNode(child);
		}

		// 末尾の余分な改行を削除
		return text.replace(/\n+$/, "");
	}

	// contenteditable div から改行を保持したプレーンテキストを取得するヘルパー関数
	// link-list モード用：すべてのHTMLタグを除去し、純粋なテキストのみを抽出
	function getPlainTextWithLineBreaks(htmlContent) {
		// XSS対策: 入力HTMLをサニタイズ
		const sanitizedContent = sanitizeHTML(htmlContent);
		const tempDiv = document.createElement("div");
		tempDiv.innerHTML = sanitizedContent;

		// innerText を使用すると、ブラウザが自動的に<br>やblock要素を改行に変換してくれる
		let text = tempDiv.innerText || tempDiv.textContent || "";

		// 末尾の余分な改行を削除
		return text.replace(/\n+$/, "");
	}

	// HTML出力をクリーンアップする関数（重複タグの統合と空タグの削除）
	function cleanupHTML(html) {
		let cleaned = html;

		// 0. 視覚効果用の data-highlight 属性を削除
		cleaned = cleaned.replace(/\s*data-highlight="true"/g, '');

		// 1. 連続する同じタグを統合（例: </strong><strong> を削除）
		const tagsToMerge = ['strong', 'mark', 'b', 'em', 'u', 'span', 'i'];
		tagsToMerge.forEach(tag => {
			// </tag><tag> パターンを削除（属性なしの場合）
			const pattern = new RegExp(`</${tag}><${tag}>`, 'g');
			cleaned = cleaned.replace(pattern, '');

			// 属性付きの場合も対応（同じ属性の連続タグを統合）
			// 例: </strong><strong class="x"> の場合は統合しない（属性が異なる可能性）
		});

		// 2. 空のタグを削除（例: <strong></strong>）
		const allTags = ['strong', 'mark', 'b', 'em', 'u', 'span', 'i', 'a', 'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
		allTags.forEach(tag => {
			// 属性なしの空タグ
			const emptyPattern = new RegExp(`<${tag}></${tag}>`, 'g');
			cleaned = cleaned.replace(emptyPattern, '');

			// 属性付きの空タグ（スペースや改行のみを含む場合も）
			const emptyWithAttrsPattern = new RegExp(`<${tag}[^>]*>\\s*</${tag}>`, 'g');
			cleaned = cleaned.replace(emptyWithAttrsPattern, '');
		});

		// 複数回実行して、入れ子の空タグも削除
		let previousCleaned = '';
		let iterations = 0;
		while (previousCleaned !== cleaned && iterations < 5) {
			previousCleaned = cleaned;
			allTags.forEach(tag => {
				const emptyPattern = new RegExp(`<${tag}></${tag}>`, 'g');
				cleaned = cleaned.replace(emptyPattern, '');
				const emptyWithAttrsPattern = new RegExp(`<${tag}[^>]*>\\s*</${tag}>`, 'g');
				cleaned = cleaned.replace(emptyWithAttrsPattern, '');
			});
			iterations++;
		}

		return cleaned;
	}

	// ========================================
	// 10. 変換ロジック
	// ========================================

	/**
	 * テキストエリアの内容をHTMLに変換する共通関数
	 * タグタイプに応じて適切な変換処理を実行
	 * @param {HTMLElement} textareaElement - 変換対象のテキストエリア要素
	 * @returns {string} 変換されたHTML文字列
	 */
	function convertTextToHtmlString(textareaElement) {
		const htmlContent = textareaElement.innerHTML;

		const tagId = textareaElement.getAttribute("data-tag-id");
		const tagInfo = getCustomTagInfo(tagId);

		// タグ情報がない場合は、HTMLコメントとして警告を返す
		if (!tagInfo) return `<!-- 警告: 不明なタグID (${tagId}) のためスキップされました -->\n`;

		const tagType = tagInfo.tagType;
		const templateString = tagInfo.template;

		// 静的モード: テンプレートをそのまま出力（入力テキストは無視、空でもOK）
		if (tagType === "static") {
			if (templateString.trim() === "") {
				return `<!-- 警告: タグ (${tagInfo.name}) の雛形が空のためスキップされました -->\n`;
			}
			const output = templateString.trim();
			// 最後に改行を加えておく (追記時に扱いやすいように)
			return output + "\n";
		}

		// 静的モード以外では、入力が空の場合は空文字を返す
		if (!htmlContent || htmlContent.trim() === "") return "";

		// link-list モードの場合はプレーンテキストを取得、それ以外は書式を保持
		const text = (tagInfo.tagType === "link-list")
			? getPlainTextWithLineBreaks(htmlContent)
			: getTextWithLineBreaks(htmlContent);

		let output = "";

		// テンプレートが空文字列の場合、警告を返す
		if (templateString.trim() === "") {
			return `<!-- 警告: タグ (${tagInfo.name}) の雛形が空のためスキップされました -->\n`;
		}

		if (tagType === "link") {
			// リンクタグ系の特別処理 (link)
			const lines = text.trim().split("\n").map((line) => line.trim());
			const linkText = lines[0] || "リンク";
			const url = lines[1] || "";

			if (url) {
				output = templateString.replace(/\[TEXT\]/g, linkText).replace(/\[URL\]/g, url);
				// 最後の項目から <br> を削除するオプション
				if (tagInfo.removeLastBr) {
					output = output.replace(/<br\s*\/?\s*>\s*$/i, '').trimEnd();
				}
			} else {
				output = `<!-- 警告: リンクエリアですがURLが2行目に指定されていません (${linkText}) -->\n`;
			}
		} else if (tagType === "p-list") {
			// P+リストタグの処理 (p-list) - 複数段落対応版
			const lines = text.trim().split("\n");
			const paragraphLines = [];
			const listLines = [];

			// 行を段落とリストに分類
			lines.forEach(line => {
				const trimmedLine = line.trim();
				if (trimmedLine === "") return; // 空行はスキップ

				// - または * で始まる行はリスト項目
				if (trimmedLine.startsWith('-') || trimmedLine.startsWith('*')) {
					// 先頭の - または * と空白を除去
					const listItem = trimmedLine.replace(/^[-*]\s*/, '').trim();
					listLines.push(listItem);
				} else {
					// それ以外は段落
					paragraphLines.push(trimmedLine);
				}
			});

			// リストコンテンツの構築
			let listContent = "";
			if (listLines.length > 0) {
				const listItems = listLines.map((line) => {
					// すでにliタグが含まれていたらそのまま、そうでなければliタグで囲む
					return line.match(/^\s*<li/i) ? line : `<li>${line}</li>`;
				});
				// 最後の項目から <br> を削除するオプション
				if (tagInfo.removeLastBr && listItems.length > 0) {
					const lastIndex = listItems.length - 1;
					listItems[lastIndex] = listItems[lastIndex].replace(/<br\s*\/?\s*>\s*$/i, '').trimEnd();
				}
				listContent = listItems.join("\n");
			} else {
				listContent = `<li>リスト項目がありません</li>`;
			}

			// テンプレートから開始
			output = templateString;

			// [TEXT_LIST] の置換
			output = output.replace(/\[TEXT_LIST\]/g, listContent);

			// [TEXT_P_1], [TEXT_P_2], ... の置換（番号付き段落プレースホルダー）
			paragraphLines.forEach((pLine, index) => {
				const placeholder = `[TEXT_P_${index + 1}]`;
				// 正規表現のエスケープ処理をして置換
				const escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
				output = output.replace(new RegExp(escapedPlaceholder, 'g'), pLine);
			});

			// 後方互換性: [TEXT_P] は最初の段落で置換（または全段落を結合）
			if (paragraphLines.length > 0) {
				// 最初の段落のみを使用
				output = output.replace(/\[TEXT_P\]/g, paragraphLines[0]);
			} else {
				output = output.replace(/\[TEXT_P\]/g, '段落がありません');
			}
		} else if (tagType === "list") {
			// リストタグの処理 (list)
			const listItems = text.trim().split("\n")
				.filter((line) => line.trim() !== "")
				.map((line) => {
					// すでにliタグが含まれていたらそのまま、そうでなければliタグで囲む
					return line.match(/^\s*<li/i) ? line : `<li>${line.trim()}</li>`;
				});

			// 最後の項目から <br> を削除するオプション
			if (tagInfo.removeLastBr && listItems.length > 0) {
				const lastIndex = listItems.length - 1;
				listItems[lastIndex] = listItems[lastIndex].replace(/<br\s*\/?\s*>\s*$/i, '').trimEnd();
			}

			let listContent = listItems.join("\n");
			if (listContent.trim() === "") listContent = `<li>リスト項目がありません</li>`;

			output = templateString.replace(/\[TEXT\]/g, listContent);
		} else if (tagType === "single") {
			// 単一タグの処理 (single)
			let content = text.trim().replace(/\n/g, "<br>");
			// 最後の <br> を削除するオプション
			if (tagInfo.removeLastBr) {
				content = content.replace(/<br\s*\/?\s*>\s*$/i, '').trimEnd();
			}
			output = templateString.replace(/\[TEXT\]/g, content);
		} else if (tagType === "multi") {
			// マルチラインタグの処理 (multi) - 各行が個別のタグになる
			const lines = text.trim().split("\n").filter((line) => line.trim() !== "");

			if (lines.length > 0) {
				const lineItems = lines.map((line) => {
					return templateString.replace(/\[TEXT\]/g, line.trim());
				});
				// 最後の項目から <br> を削除するオプション
				if (tagInfo.removeLastBr && lineItems.length > 0) {
					const lastIndex = lineItems.length - 1;
					lineItems[lastIndex] = lineItems[lastIndex].replace(/<br\s*\/?\s*>\s*$/i, '').trimEnd();
				}
				output = lineItems.join("\n");
			} else {
				output = `<!-- 警告: テキストが入力されていません -->\n`;
			}
		} else if (tagType === "link-list") {
			// リンクリストの処理 (link-list)
			const lines = text.trim().split("\n").map((line) => line.trim()).filter((line) => line !== "");
			const linkItems = [];

			// カスタムリンク項目テンプレートを取得（デフォルト値を設定）
			const linkItemTemplate = tagInfo.linkItemTemplate || '<li class="rtoc-item"><a href="[URL]">[TEXT]</a></li>';

			// 2行ずつペアにして処理 (奇数行:TEXT, 偶数行:URL)
			for (let i = 0; i < lines.length; i += 2) {
				const linkText = lines[i] || "";
				const linkUrl = lines[i + 1] || "";

				if (linkText && linkUrl) {
					// カスタムテンプレートを使用して変換
					const itemHtml = linkItemTemplate.replace(/\[TEXT\]/g, linkText).replace(/\[URL\]/g, linkUrl);
					linkItems.push(itemHtml);
				} else if (linkText) {
					// URLがない場合は警告
					linkItems.push(`<!-- 警告: "${linkText}" のURLが指定されていません -->`);
				}
			}

			// 最後の項目から <br> または <br /> を削除するオプション
			if (tagInfo.removeLastBr && linkItems.length > 0) {
				const lastIndex = linkItems.length - 1;
				// 末尾の <br> または <br /> とその後の空白を削除（大文字小文字・スペースの有無を考慮）
				linkItems[lastIndex] = linkItems[lastIndex].replace(/<br\s*\/?\s*>\s*$/i, '').trimEnd();
			}

			let linkListContent = "";
			if (linkItems.length > 0) {
				linkListContent = linkItems.join("\n");
			} else {
				linkListContent = `<li class="rtoc-item">リンク項目がありません</li>`;
			}

			output = templateString.replace(/\[LINK_LIST\]/g, linkListContent.trim());
		}

		// HTML出力をクリーンアップ（重複タグの統合と空タグの削除）
		output = cleanupHTML(output);

		// 最後に改行を加えておく (追記時に扱いやすいように)
		return output + "\n";
	}

	// 挿入位置セレクタを更新する関数 (変更なし)
	const updateInsertionPoints = () => {
		const selectedValue = insertionPointSelector.value;

		insertionPointSelector.innerHTML = "";

		const allGroups = container.querySelectorAll(".input-area-group");

		const optionStart = document.createElement("option");
		optionStart.value = "start";
		optionStart.textContent = "リストの先頭 (#1 として追加)";
		insertionPointSelector.appendChild(optionStart);

		allGroups.forEach((group, index) => {
			const tagId = group.querySelector(".input-text-area").getAttribute("data-tag-id");
			// getCustomTagInfo を使って最新のタグ名を取得
			const tagInfo = getCustomTagInfo(tagId) || { name: "不明なタグ" };

			const currentNumber = index + 1;
			const numberPrefix = group.querySelector(".number-prefix");
			if (numberPrefix) {
				numberPrefix.textContent = `#${currentNumber}`;
			}

			const option = document.createElement("option");
			option.value = group.id;
			option.textContent = `#${currentNumber} ${tagInfo.name} の直後`;

			insertionPointSelector.appendChild(option);
		});

		const optionEnd = document.createElement("option");
		optionEnd.value = "end";
		optionEnd.textContent = `リストの末尾 (#${allGroups.length + 1} として追加)`;
		insertionPointSelector.appendChild(optionEnd);

		if (selectedValue && Array.from(insertionPointSelector.options).some((opt) => opt.value === selectedValue)) {
			insertionPointSelector.value = selectedValue;
		} else {
			insertionPointSelector.value = allGroups.length === 0 ? "start" : "end";
		}
	};

	// Word書式を保持した貼り付けの処理
	const handleFormattedPaste = (event, editableDiv) => {
		const clipboardData = event.clipboardData || window.clipboardData;
		if (!clipboardData) return;

		// HTMLデータを取得
		const htmlData = clipboardData.getData("text/html");
		if (!htmlData) return; // HTMLデータがない場合は通常の貼り付けを許可

		// WordやOfficeアプリからのHTMLかどうかを検出
		// Word/Office HTMLには特定のマーカーが含まれている
		const isFromWord = htmlData.includes("urn:schemas-microsoft-com:office") ||
			htmlData.includes("class=\"Mso") ||
			htmlData.includes("class='Mso") ||
			htmlData.match(/<meta\s+name=["']?Generator["']?\s+content=["']?Microsoft/i);

		// Word/Office以外からのHTML（内部コピーなど）の場合は通常の貼り付けを許可
		if (!isFromWord) return;

		event.preventDefault(); // デフォルトの貼り付けをキャンセル

		// HTMLをパースして変換
		const convertedHTML = convertWordHTMLToTags(htmlData);

		// XSS対策: 変換されたHTMLをサニタイズ
		const sanitizedHTML = sanitizeHTML(convertedHTML);

		// contenteditable の場合、カーソル位置にHTMLを挿入
		const selection = window.getSelection();
		if (!selection.rangeCount) return;

		const range = selection.getRangeAt(0);
		range.deleteContents();

		// HTMLフラグメントを作成して挿入
		const tempDiv = document.createElement("div");
		tempDiv.innerHTML = sanitizedHTML;
		const fragment = document.createDocumentFragment();
		while (tempDiv.firstChild) {
			fragment.appendChild(tempDiv.firstChild);
		}
		range.insertNode(fragment);

		// カーソルを挿入した内容の最後に移動
		range.collapse(false);
		selection.removeAllRanges();
		selection.addRange(range);

		// 保存をトリガー
		saveInputAreas();
	};

	// テンプレートから開始タグと終了タグを抽出するヘルパー関数
	const extractTagsFromTemplate = (template) => {
		if (!template || typeof template !== 'string') {
			return { openTag: "", closeTag: "" };
		}
		// [TEXT] で分割して開始タグと終了タグを取得
		const parts = template.split("[TEXT]");
		return {
			openTag: parts[0] || "",
			closeTag: parts[1] || ""
		};
	};

	// Word HTMLを設定されたタグに変換
	const convertWordHTMLToTags = (html) => {
		// 一時的なdiv要素でHTMLをパース
		const tempDiv = document.createElement("div");
		tempDiv.innerHTML = html;

		// style、script、コメントノードを削除
		const removeUnwantedNodes = (parent) => {
			const nodesToRemove = [];

			for (let node of parent.childNodes) {
				// コメントノード、style、scriptタグを削除対象にマーク
				if (node.nodeType === Node.COMMENT_NODE) {
					nodesToRemove.push(node);
				} else if (node.nodeType === Node.ELEMENT_NODE) {
					const tagName = node.tagName.toLowerCase();
					if (tagName === "style" || tagName === "script" || tagName === "meta" || tagName === "link") {
						nodesToRemove.push(node);
					} else if (tagName === "p") {
						// Wordが出力する不要なCSSコードを含むp要素をフィルタ
						const textContent = node.textContent.trim();
						// CSSコード、コメントマーカー、@font-face、mso-などを含むp要素を削除
						if (
							textContent.startsWith("<!--") ||
							textContent.startsWith("-->") ||
							textContent.startsWith("/*") ||
							textContent.startsWith("*/") ||
							textContent.includes("@font-face") ||
							textContent.includes("@page") ||
							textContent.includes("mso-") ||
							textContent.includes("panose-") ||
							textContent.match(/^\s*\{/) ||
							textContent.match(/^\s*\}/) ||
							textContent.includes("Style Definitions") ||
							textContent.includes("Font Definitions") ||
							textContent.includes("Page Definitions")
						) {
							nodesToRemove.push(node);
						} else {
							// 通常のp要素は子要素もチェック
							removeUnwantedNodes(node);
						}
					} else {
						// 再帰的に子要素もチェック
						removeUnwantedNodes(node);
					}
				}
			}

			// マークしたノードを削除
			nodesToRemove.forEach(node => node.remove());
		};

		removeUnwantedNodes(tempDiv);

		// テキストノードと書式を再帰的に処理
		const processNode = (node) => {
			if (node.nodeType === Node.TEXT_NODE) {
				// タブを通常のスペースに変換し、同一行内の連続スペースを1つにまとめる
				// 改行は保持する
				return node.textContent
					.replace(/\t/g, " ")  // タブをスペースに変換
					.replace(/[^\S\r\n]+/g, " ");  // 改行以外の連続空白文字を1つのスペースに
			}

			if (node.nodeType === Node.ELEMENT_NODE) {
				const tagName = node.tagName.toLowerCase();
				let result = "";
				let openTag = "";
				let closeTag = "";

				// br タグや p タグの改行を処理
				if (tagName === "br") {
					return "\n";
				} else if (tagName === "p" || tagName === "div") {
					// p や div タグの前後に改行を追加（末尾のみ）
					closeTag = "\n";
				}

				// Word/ブラウザの書式タグをカスタムタグにマッピング
				// 太字、マーカー、リンクを保持し、他の書式は無視する
				if ((tagName === "b" || tagName === "strong") && formattingMap.bold) {
					// テンプレートから開始タグと終了タグを抽出（後方互換性のため tag もサポート）
					if (formattingMap.bold.template) {
						const tags = extractTagsFromTemplate(formattingMap.bold.template);
						openTag = tags.openTag;
						closeTag = tags.closeTag;
					} else if (formattingMap.bold.tag) {
						openTag = `<${formattingMap.bold.tag}>`;
						closeTag = `</${formattingMap.bold.tag}>`;
					}
				} else if ((tagName === "mark" || (tagName === "span" && node.style.backgroundColor)) && formattingMap.highlight) {
					// ハイライト（背景色がある場合も含む）
					// Word のハイライトは span の background-color として来ることがある
					if (formattingMap.highlight.template) {
						const tags = extractTagsFromTemplate(formattingMap.highlight.template);
						// data-highlight 属性を追加して視覚効果を統一
						openTag = tags.openTag.replace(/>$/, ' data-highlight="true">');
						closeTag = tags.closeTag;
					} else if (formattingMap.highlight.tag) {
						openTag = `<${formattingMap.highlight.tag} data-highlight="true">`;
						closeTag = `</${formattingMap.highlight.tag}>`;
					}
				} else if (tagName === "a") {
					// リンクタグを保持（href属性を含む）
					const href = node.getAttribute("href");
					if (href) {
						openTag = `<a href="${href}">`;
						closeTag = `</a>`;
					}
				}
				// 斜体、下線などの他の書式は無視（タグを付けずに子ノードのみ処理）

				// 子ノードを処理
				for (let child of node.childNodes) {
					result += processNode(child);
				}

				return openTag + result + closeTag;
			}

			return "";
		};

		// ルート要素の子ノードを処理
		let convertedText = "";
		for (let child of tempDiv.childNodes) {
			convertedText += processNode(child);
		}

		// 最終的な整形：前後の空白を削除、改行を保持しつつ同一行内の連続スペースを1つに
		convertedText = convertedText
			.trim()
			.replace(/[^\S\r\n]+/g, " ")  // 改行以外の連続空白を1つのスペースに
			.replace(/\n+/g, "\n");  // 連続する改行を1つに統一

		return convertedText;
	};

	// テキストエリアを作成し、選択された位置に挿入する関数 (引数を変更)
	const createTextarea = (tagId, tagType, tagName, initialContent = "") => {
		areaIndex++;
		const groupDiv = document.createElement("div");
		groupDiv.className = "input-area-group";
		groupDiv.id = `area-group-${areaIndex}`;
		groupDiv.setAttribute('role', 'listitem');
		groupDiv.setAttribute('aria-label', `入力エリア: ${tagName}`);

		// AbortControllerでイベントリスナーを管理（メモリリーク対策）
		const abortController = new AbortController();
		const signal = abortController.signal;

		// 要素削除時にイベントリスナーをクリーンアップする関数を保存
		groupDiv._cleanup = () => abortController.abort();

		const tagLabel = document.createElement("span");
		tagLabel.className = "input-label-tag";
		tagLabel.draggable = true;  // ラベルのみドラッグ可能に
		tagLabel.style.cursor = "grab";
		tagLabel.title = "ドラッグして順序を変更";

		const numberPrefix = document.createElement("span");
		numberPrefix.className = "number-prefix";
		tagLabel.appendChild(numberPrefix);

		const tagTextSpan = document.createElement("span");
		tagTextSpan.textContent = tagName;
		tagLabel.appendChild(tagTextSpan);

		// ドラッグイベント（ラベルのみ）（signalオプションでメモリリーク防止）
		tagLabel.addEventListener("dragstart", (e) => {
			e.dataTransfer.effectAllowed = "move";
			e.dataTransfer.setData("text/html", groupDiv.id);
			groupDiv.classList.add("dragging-input-area");
			tagLabel.style.cursor = "grabbing";
		}, { signal });

		tagLabel.addEventListener("dragend", (e) => {
			groupDiv.classList.remove("dragging-input-area");
			tagLabel.style.cursor = "grab";
		}, { signal });

		groupDiv.addEventListener("dragover", (e) => {
			e.preventDefault();
			e.dataTransfer.dropEffect = "move";

			const draggingGroup = document.querySelector(".dragging-input-area");
			if (draggingGroup && draggingGroup !== groupDiv) {
				const rect = groupDiv.getBoundingClientRect();
				const midpoint = rect.top + rect.height / 2;

				if (e.clientY < midpoint) {
					groupDiv.style.borderTop = "3px solid #4CAF50";
					groupDiv.style.borderBottom = "";
				} else {
					groupDiv.style.borderTop = "";
					groupDiv.style.borderBottom = "3px solid #4CAF50";
				}
			}
		}, { signal });

		groupDiv.addEventListener("dragleave", (e) => {
			groupDiv.style.borderTop = "";
			groupDiv.style.borderBottom = "";
		}, { signal });

		groupDiv.addEventListener("drop", (e) => {
			e.preventDefault();
			groupDiv.style.borderTop = "";
			groupDiv.style.borderBottom = "";

			const draggingGroup = document.querySelector(".dragging-input-area");
			if (!draggingGroup || draggingGroup === groupDiv) return;

			// 位置を判定
			const rect = groupDiv.getBoundingClientRect();
			const midpoint = rect.top + rect.height / 2;

			if (e.clientY < midpoint) {
				// 上に挿入
				container.insertBefore(draggingGroup, groupDiv);
			} else {
				// 下に挿入
				container.insertBefore(draggingGroup, groupDiv.nextSibling);
			}

			updateInsertionPoints();
			saveInputAreas();
		}, { signal });

		const newTextarea = document.createElement("div");
		newTextarea.className = `input-text-area input-${tagId}-area`;
		newTextarea.id = `text-area-${tagId}-${areaIndex}`;

		// 静的モードの場合は編集不可にする
		if (tagType === "static") {
			newTextarea.setAttribute("contenteditable", "false");
			newTextarea.style.backgroundColor = "#f0f0f0";
			newTextarea.style.cursor = "not-allowed";
			newTextarea.style.opacity = "0.7";
		} else {
			newTextarea.setAttribute("contenteditable", "true");
		}

		newTextarea.setAttribute("role", "textbox");
		newTextarea.setAttribute("aria-multiline", "true");
		newTextarea.setAttribute("aria-label", `${tagName}のテキスト入力エリア`);
		newTextarea.setAttribute("data-tag-id", tagId); // タグIDをデータ属性として保持
		// 初期コンテンツを設定（空の場合は完全に空にしてプレースホルダーを表示）
		if (initialContent && initialContent.trim() !== "") {
			// XSS対策: 保存されたコンテンツをサニタイズしてから設定
			newTextarea.innerHTML = sanitizeHTML(initialContent);
		}

		// テキストエリアの内容が変更されたら保存（デバウンス版を使用してパフォーマンス向上）
		newTextarea.addEventListener("input", () => {
			debouncedSaveInputAreas();
		}, { signal });

		// 貼り付けイベント：Word書式を保持
		newTextarea.addEventListener("paste", (e) => {
			handleFormattedPaste(e, newTextarea);
		}, { signal });

		let placeholderText = `${tagName}タグ用のテキストエリア`;

		if (tagType === "multi") {
			placeholderText = `【${tagName}】マルチラインモード：各行が個別のタグになります`;
		} else if (tagType === "p-list") {
			placeholderText = `【${tagName}】段落+リストモード：通常の行は段落、「-」または「*」で始まる行はリスト項目になります`;
		} else if (tagType === "link-list") {
			placeholderText = `【${tagName}】リンクリストモード：奇数行にリンクテキスト、偶数行にURLを交互に入力してください`;
		} else if (tagType === "static") {
			placeholderText = `【${tagName}】静的モード：テンプレートで設定した内容がそのまま出力されます（入力不可）`;
		} else {
			// 古いモード (single, list, link) のための後方互換性
			placeholderText = `${tagName}エリア。テキストを入力してください。`;
		}
		newTextarea.setAttribute("data-placeholder", placeholderText);

		const convertButton = document.createElement("button");
		convertButton.textContent = "変換";
		convertButton.className = "convert-input-button";
		convertButton.setAttribute('aria-label', `${tagName}入力エリアを個別に変換`);
		convertButton.style.backgroundColor = "#4CAF50";
		convertButton.style.color = "white";
		convertButton.style.marginRight = "5px";

		convertButton.addEventListener("click", () => {
			// 変換直前に設定を保存し、最新のテンプレートを取得できるように保証
			saveAllPatterns();

			// この入力エリアのみを変換
			const outputHtmlString = convertTextToHtmlString(newTextarea);

			// 新しい内容が空でなければ追記
			if (outputHtmlString.trim().length > 0) {
				const existingContent = htmlCodeOutput.textContent.trim();
				if (existingContent.length > 0) {
					htmlCodeOutput.textContent += "\n"; // 既存の内容があれば改行を追加
				}
				htmlCodeOutput.textContent += outputHtmlString;
			}
		}, { signal });

		const deleteButton = document.createElement("button");
		deleteButton.textContent = "削除";
		deleteButton.className = "delete-input-button";
		deleteButton.setAttribute('aria-label', `${tagName}入力エリアを削除`);

		deleteButton.addEventListener("click", () => {
			// イベントリスナーをクリーンアップしてメモリリークを防止
			if (groupDiv._cleanup) {
				groupDiv._cleanup();
			}
			groupDiv.remove();
			updateInsertionPoints();
			saveInputAreas(); // 削除後に保存
		}, { signal });

		groupDiv.appendChild(tagLabel);
		groupDiv.appendChild(newTextarea);
		groupDiv.appendChild(convertButton);
		groupDiv.appendChild(deleteButton);

		const insertPoint = insertionPointSelector.value;

		if (insertPoint === "start") {
			container.prepend(groupDiv);
		} else if (insertPoint === "end") {
			container.appendChild(groupDiv);
		} else {
			const selectedElement = document.getElementById(insertPoint);
			if (selectedElement) {
				container.insertBefore(groupDiv, selectedElement.nextElementSibling);
			} else {
				container.appendChild(groupDiv);
			}
		}

		updateInsertionPoints();
		insertionPointSelector.value = groupDiv.id;
		saveInputAreas(); // 作成後に保存
	};

	// ========================================
	// 11. イベントハンドラとその他の機能
	// ========================================

	// 選択範囲をタグで囲むコア関数 (変更なし)
	window.wrapSelectionWithTag = (tagName) => {
		const textarea = document.activeElement;
		if (!textarea || textarea.tagName !== "TEXTAREA" || !textarea.classList.contains("input-text-area")) {
			return;
		}
		const start = textarea.selectionStart;
		const end = textarea.selectionEnd;
		const selectedText = textarea.value.substring(start, end);
		if (selectedText.length === 0) {
			return;
		}

		let openTag;
		let closeTag;

		if (tagName === "a") {
			openTag = `<a href="">`;
			closeTag = `</a>`;
		} else {
			openTag = `<${tagName}>`;
			closeTag = `</${tagName}>`;
		}

		const newText = textarea.value.substring(0, start) + openTag + selectedText + closeTag + textarea.value.substring(end);

		textarea.value = newText;

		let newCursorPos;
		if (tagName === "a") {
			// カーソルを href="" のダブルクォートの内側 (a href="|") に設定
			newCursorPos = start + openTag.length - 1;
		} else {
			newCursorPos = start + openTag.length + selectedText.length + closeTag.length;
		}

		textarea.selectionStart = newCursorPos;
		textarea.selectionEnd = newCursorPos;
		textarea.focus();
	};

	// 出力エリアクリア機能 (変更なし)
	if (clearOutputButton) {
		clearOutputButton.addEventListener("click", function () {
			htmlCodeOutput.textContent = "";
		});
	}

	// すべての入力エリアをクリア機能
	if (clearAllInputsButton) {
		clearAllInputsButton.addEventListener("click", function () {
			// 確認ダイアログを表示
			if (!confirm("すべての入力エリアを削除してもよろしいですか？")) {
				return;
			}

			// すべての子要素を削除
			while (container.firstChild) {
				const child = container.firstChild;
				// クリーンアップ関数が存在する場合は実行（メモリリーク防止）
				if (child._cleanup) {
					child._cleanup();
				}
				container.removeChild(child);
			}

			// 挿入位置セレクタを更新
			updateInsertionPoints();

			// 状態を保存
			saveInputAreas();
		});
	}

	// コードコピー機能 (変更なし)
	if (copyOutputButton) {
		copyOutputButton.addEventListener("click", async function () {
			const textToCopy = htmlCodeOutput.textContent;
			if (textToCopy.trim().length === 0) {
				console.warn("コピーするHTMLコードがありません。");
				return;
			}

			try {
				await navigator.clipboard.writeText(textToCopy);
				const originalText = copyOutputButton.textContent;
				copyOutputButton.textContent = CONSTANTS.MESSAGES.COPY_SUCCESS;
				setTimeout(() => {
					copyOutputButton.textContent = originalText;
				}, 1500);
			} catch (err) {
				console.error("クリップボードへの書き込みに失敗しました:", err);
				// navigator.clipboard が使えない環境（httpなど）のためのフォールバック
				try {
					const tempTextarea = document.createElement("textarea");
					tempTextarea.value = textToCopy;
					document.body.appendChild(tempTextarea);
					tempTextarea.select();
					document.execCommand("copy");
					document.body.removeChild(tempTextarea);

					const originalText = copyOutputButton.textContent;
					copyOutputButton.textContent = CONSTANTS.MESSAGES.COPY_SUCCESS_FALLBACK;
					setTimeout(() => {
						copyOutputButton.textContent = originalText;
					}, 1500);
				} catch (copyErr) {
					console.error("フォールバックコピーにも失敗しました:", copyErr);
				}
			}
		});
	}

	// ========================================
	// 12. 初期化
	// ========================================

	// ページ読み込み時にデータを復元してUIを初期化
	loadAllPatterns();
	loadFormattingMap();
	buildFormattingUI(); // 書式マッピング管理UIを初期表示
	rebuildTagButtonList(); // タグボタン管理一覧を初期表示
	loadInputAreas(); // 入力エリアを復元

	// 一括変換機能
	if (convertAllButton) {
		convertAllButton.addEventListener("click", function () {
			// 変換直前に設定を保存し、最新のテンプレートを取得できるように保証します
			saveAllPatterns();

			const allTextareas = document.querySelectorAll(".input-text-area");
			let outputHtmlString = "";

			if (allTextareas.length === 0) {
				outputHtmlString = "<!-- 警告: 入力されたテキストエリアがありません -->\n";
			} else {
				allTextareas.forEach((textarea) => {
					outputHtmlString += convertTextToHtmlString(textarea);
				});
			}

			// 新しい内容が空でなければ追記
			if (outputHtmlString.trim().length > 0) {
				const existingContent = htmlCodeOutput.textContent.trim();
				if (existingContent.length > 0) {
					htmlCodeOutput.textContent += "\n"; // 既存の内容があれば改行を追加
				}
				// 末尾の不要な改行を削除して追記
				htmlCodeOutput.textContent += outputHtmlString.trimEnd();
			}
		});
	}
});
