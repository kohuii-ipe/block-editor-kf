document.addEventListener("DOMContentLoaded", function () {
	// --- 要素の取得 ---
	const convertAllButton = document.getElementById("convertToAllButton");
	const clearOutputButton = document.getElementById("clearOutputButton");
	const copyOutputButton = document.getElementById("copyOutputButton");
	const container = document.getElementById("container");
	const insertionPointSelector = document.getElementById("insertionPointSelector");
	const htmlCodeOutput = document.getElementById("htmlCodeOutput");

	const patternSelectorContainer = document.getElementById("patternSelectorContainer");
	const addPatternButton = document.getElementById("addPatternButton");
	const renamePatternButton = document.getElementById("renamePatternButton");
	const deletePatternButton = document.getElementById("deletePatternButton");
	const tagButtonsContainer = document.getElementById("tagButtonsContainer"); // 新しいボタンコンテナ

	// タグ管理関連
	const addNewTagButton = document.getElementById("addNewTagButton");
	const tagButtonList = document.getElementById("tagButtonList");
	const tagModal = document.getElementById("tagModal");
	const modalTitle = document.getElementById("modalTitle");
	const tagModalForm = document.getElementById("tagModalForm");
	const tagIdInput = document.getElementById("tagIdInput");
	const tagNameInput = document.getElementById("tagNameInput");
	const tagTypeSelector = document.getElementById("tagTypeSelector");
	const tagTemplateInput = document.getElementById("tagTemplateInput");
	const linkItemTemplateSection = document.getElementById("linkItemTemplateSection");
	const linkItemTemplateInput = document.getElementById("linkItemTemplateInput");
	const closeModalSpan = tagModal.querySelector(".close");

	// ショートカット管理関連
	const shortcutList = document.getElementById("shortcutList");
	const addNewShortcutButton = document.getElementById("addNewShortcutButton");
	const shortcutModal = document.getElementById("shortcutModal");
	const shortcutModalForm = document.getElementById("shortcutModalForm");
	const shortcutKeyInput = document.getElementById("shortcutKeyInput");
	const shortcutTagInput = document.getElementById("shortcutTagInput");
	const closeShortcutModalSpan = document.getElementById("closeShortcutModal");

	let areaIndex = 0;
	let maxPatternId = 0;
	let maxTagId = 0;

	// --- 雛形設定のデフォルトデータ（構造変更） ---
	const defaultTagButtons = [
		// id, name, template, tagType (single, list, link, p-list)
		{ id: "p", name: "pタグ", template: "<p>[TEXT]</p>", tagType: "single" },
		{ id: "h2", name: "h2タグ", template: "<h2>[TEXT]</h2>", tagType: "single" },
		{
			id: "a-int",
			name: "内部リンク",
			template: '<a href="[URL]">[TEXT]</a>',
			tagType: "link",
		},
		{
			id: "a-ext",
			name: "外部リンク",
			template: '<a href="[URL]" target="_blank">[TEXT]</a>',
			tagType: "link",
		},
		{ id: "ul", name: "ULタグ", template: "<ul>\n[TEXT]\n</ul>", tagType: "list" },
		{ id: "box", name: "枠", template: '<div class="box">\n[TEXT]\n</div>', tagType: "single" },
		{
			id: "box-p-ul",
			name: "箇条書きリスト",
			template: '<div class="box"><p>[TEXT_P]</p><ul>\n[TEXT_LIST]\n</ul></div>',
			tagType: "p-list",
		},
	];

	const defaultPatterns = {
		pattern1: {
			name: "ソニー (デフォルト)",
			buttons: JSON.parse(JSON.stringify(defaultTagButtons)), // ディープコピー
			shortcuts: {
				B: { tag: "strong", displayName: "Ctrl/Cmd+B" },
				M: { tag: "mark", displayName: "Ctrl/Cmd+M" },
				U: { tag: "sup", displayName: "Ctrl/Cmd+U" },
				L: { tag: "li", displayName: "Ctrl/Cmd+L" },
				K: { tag: "a", displayName: "Ctrl/Cmd+K" },
			},
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
			shortcuts: {
				B: { tag: "strong", displayName: "Ctrl/Cmd+B" },
				M: { tag: "mark", displayName: "Ctrl/Cmd+M" },
				U: { tag: "sup", displayName: "Ctrl/Cmd+U" },
				L: { tag: "li", displayName: "Ctrl/Cmd+L" },
				K: { tag: "a", displayName: "Ctrl/Cmd+K" },
			},
		},
	};

	let patterns = {}; // すべてのパターン設定を保持するグローバルオブジェクト

	// --- 永続化と初期化 ---
	const STORAGE_KEY = "customTagPatternsV2";
	const SETTINGS_KEY = "editorSettingsV2";
	const SHORTCUTS_KEY = "keyboardShortcutsV2";

	// デフォルトのショートカット設定
	const defaultShortcuts = {
		B: { tag: "strong", displayName: "Ctrl/Cmd+B" },
		M: { tag: "mark", displayName: "Ctrl/Cmd+M" },
		U: { tag: "sup", displayName: "Ctrl/Cmd+U" },
		L: { tag: "li", displayName: "Ctrl/Cmd+L" },
		K: { tag: "a", displayName: "Ctrl/Cmd+K" },
	};

	let keyboardShortcuts = {};

	const saveAllPatterns = () => {
		const selectedId = getSelectedPatternId();

		localStorage.setItem(STORAGE_KEY, JSON.stringify(patterns));

		const settings = {
			selectedPattern: selectedId,
		};
		localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
	};

	const saveShortcuts = () => {
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];

		if (currentPattern) {
			currentPattern.shortcuts = keyboardShortcuts;
			saveAllPatterns(); // パターン全体を保存
		}
	};

	const loadShortcuts = () => {
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];

		if (currentPattern && currentPattern.shortcuts) {
			keyboardShortcuts = currentPattern.shortcuts;
		} else {
			// フォールバック：パターンにショートカットがない場合はデフォルトを設定
			keyboardShortcuts = JSON.parse(JSON.stringify(defaultShortcuts));
			if (currentPattern) {
				currentPattern.shortcuts = keyboardShortcuts;
			}
		}
	};

	// ショートカット管理UIの構築
	const buildShortcutUI = () => {
		const shortcutList = document.getElementById("shortcutList");
		if (!shortcutList) return;

		shortcutList.innerHTML = "";

		Object.keys(keyboardShortcuts).forEach((key) => {
			const shortcut = keyboardShortcuts[key];
			const item = document.createElement("div");
			item.className = "shortcut-item";

			const label = document.createElement("label");
			label.textContent = shortcut.displayName;

			const input = document.createElement("input");
			input.type = "text";
			input.value = shortcut.tag;
			input.placeholder = "タグ名を入力 (例: strong, mark, code)";

			input.addEventListener("input", (e) => {
				keyboardShortcuts[key].tag = e.target.value.trim();
				saveShortcuts();
			});

			const deleteButton = document.createElement("button");
			deleteButton.textContent = "削除";
			deleteButton.style.color = "red";
			deleteButton.addEventListener("click", () => {
				if (confirm(`ショートカット「${shortcut.displayName}」を削除してもよろしいですか？`)) {
					console.log(`ショートカット「${shortcut.displayName}」を削除しました。`);
					delete keyboardShortcuts[key];
					buildShortcutUI();
					saveShortcuts();
				}
			});

			item.appendChild(label);
			item.appendChild(input);
			item.appendChild(deleteButton);
			shortcutList.appendChild(item);
		});
	};

	const loadAllPatterns = () => {
		const savedPatterns = localStorage.getItem(STORAGE_KEY);
		if (savedPatterns) {
			try {
				patterns = JSON.parse(savedPatterns);
				if (Object.keys(patterns).length === 0) {
					patterns = defaultPatterns;
				}

				// IDの最大値を更新
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
					// 後方互換性：ショートカットがない古いパターンにデフォルトショートカットを追加
					if (!patterns[id].shortcuts) {
						patterns[id].shortcuts = JSON.parse(JSON.stringify(defaultShortcuts));
					}
				});
			} catch (e) {
				patterns = defaultPatterns;
			}
		} else {
			patterns = defaultPatterns;
		}

		const savedSettings = localStorage.getItem(SETTINGS_KEY);
		let initialSelectedPattern = "pattern1";
		if (savedSettings) {
			try {
				const settings = JSON.parse(savedSettings);
				initialSelectedPattern = settings.selectedPattern || "pattern1";
			} catch (e) {
				// do nothing
			}
		}

		rebuildPatternUI(initialSelectedPattern);
		rebuildTagButtons(); // タグボタン群も初期化時に生成
		updateInsertionPoints(); // 入力エリアのリストを更新
	};

	// --- UI構築とパターン切り替え機能 ---
	const getSelectedPatternId = () => {
		const checked = document.querySelector('input[name="conversionPattern"]:checked');
		return checked ? checked.value : Object.keys(patterns)[0];
	};

	const rebuildPatternUI = (initialSelectedId) => {
		patternSelectorContainer.innerHTML = "";

		let firstId = null;
		let selectedId = initialSelectedId;

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
				rebuildTagButtons(); // パターン切り替え時、ボタンを再生成
				rebuildTagButtonList(); // タグボタン管理一覧を更新
				loadShortcuts(); // ショートカットを再読み込み
				buildShortcutUI(); // ショートカットUIを再構築
				saveAllPatterns();
			});

			const nameSpan = document.createElement("span");
			nameSpan.id = `label-name-${id}`;
			nameSpan.textContent = pattern.name;

			label.appendChild(input);
			label.appendChild(nameSpan);
			patternSelectorContainer.appendChild(label);
		});

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
		const newName = `新しいパターン ${maxPatternId}`;

		// ★★★ 修正箇所 ★★★
		// 既存のパターンのボタン定義をディープコピー（テンプレートもそのままコピー）
		const sourcePattern = patterns[getSelectedPatternId()] || patterns[Object.keys(patterns)[0]];

		patterns[newId] = {
			name: newName,
			// テンプレート(template)もそのままコピーするように修正
			buttons: JSON.parse(
				JSON.stringify(
					sourcePattern.buttons.map((b) => ({
						id: b.id,
						name: b.name,
						template: b.template, // 修正: '' ではなく b.template をコピー
						tagType: b.tagType,
						linkItemTemplate: b.linkItemTemplate || "", // linkItemTemplateもコピー
					}))
				)
			),
			// ショートカットもコピー
			shortcuts: JSON.parse(JSON.stringify(sourcePattern.shortcuts || defaultShortcuts)),
		};

		rebuildPatternUI(newId);
		rebuildTagButtons();
		rebuildTagButtonList();
		loadShortcuts(); // 新しいパターンのショートカットを読み込み
		buildShortcutUI(); // ショートカットUIを再構築
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
			// alert("パターンは最低1つ必要です。");
			// アラートの代わりにカスタムUIやコンソールログを使用
			console.warn("パターンは最低1つ必要です。");
			return;
		}

		const currentPattern = patterns[selectedId];
		if (confirm(`パターン「${currentPattern.name}」を削除してもよろしいですか？`)) {
			console.log(`パターン「${currentPattern.name}」を削除しました。`);
			delete patterns[selectedId];

			const firstId = Object.keys(patterns)[0];
			rebuildPatternUI(firstId);
			rebuildTagButtons();
			rebuildTagButtonList();
			loadShortcuts(); // 新しいパターンのショートカットを読み込み
			buildShortcutUI(); // ショートカットUIを再構築
			saveAllPatterns();
		}
	});

	// --- タグボタン管理機能 ---

	// モーダルの表示 (追加/編集)
	const openTagModal = (button = null) => {
		tagModal.style.display = "block";
		if (button) {
			// 編集モード
			modalTitle.textContent = "タグボタンの編集";
			tagIdInput.value = button.id;
			tagNameInput.value = button.name;
			tagTypeSelector.value = button.tagType;
			tagTemplateInput.value = button.template;
			linkItemTemplateInput.value = button.linkItemTemplate || '<li class="rtoc-item"><a href="[URL]">[TEXT]</a></li>';
		} else {
			// 新規追加モード
			modalTitle.textContent = "新しいタグボタンの追加";
			tagIdInput.value = "";
			tagNameInput.value = "";
			tagTypeSelector.value = "single";
			tagTemplateInput.value = "";
			linkItemTemplateInput.value = '<li class="rtoc-item"><a href="[URL]">[TEXT]</a></li>';
		}

		// ★★★ 修正箇所 ★★★
		// モーダル表示時にプレースホルダーを更新
		updateTemplatePlaceholder();
		// ★★★ 修正箇所ここまで ★★★
	};

	// モーダルの閉じる処理
	closeModalSpan.onclick = function () {
		tagModal.style.display = "none";
	};
	// 外側クリックでモーダルを閉じる機能は無効化（誤操作防止のため）
	// window.onclick = function (event) {
	// 	if (event.target == tagModal) {
	// 		tagModal.style.display = "none";
	// 	}
	// };

	// 「新しいタグボタンを追加」ボタン
	addNewTagButton.addEventListener("click", () => openTagModal());

	// --- ショートカット追加機能 ---

	// ショートカットモーダルを開く
	addNewShortcutButton.addEventListener("click", () => {
		shortcutModal.style.display = "block";
		shortcutKeyInput.value = "";
		shortcutTagInput.value = "";
	});

	// ショートカットモーダルを閉じる
	closeShortcutModalSpan.onclick = function () {
		shortcutModal.style.display = "none";
	};

	// ショートカット追加フォームの送信
	shortcutModalForm.addEventListener("submit", (e) => {
		e.preventDefault();

		const key = shortcutKeyInput.value.trim().toUpperCase();
		const tag = shortcutTagInput.value.trim();

		// バリデーション
		if (!/^[A-Z]$/.test(key)) {
			alert("A-Zの英字1文字を入力してください。");
			return;
		}

		if (!tag) {
			alert("タグ名を入力してください。");
			return;
		}

		// 既存のショートカットがある場合は上書き確認
		if (keyboardShortcuts[key]) {
			if (!confirm(`Ctrl/Cmd+${key} は既に使用されています。上書きしますか？`)) {
				return;
			}
		}

		// ショートカットを追加
		keyboardShortcuts[key] = {
			tag: tag,
			displayName: `Ctrl/Cmd+${key}`,
		};

		shortcutModal.style.display = "none";
		buildShortcutUI();
		saveShortcuts();
	});

	// ★★★ ここから修正箇所 ★★★

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
			case "single":
				placeholder = '例: <p class="custom">[TEXT]</p>';
				break;
			case "list":
				placeholder = "例:\n<ul>\n[TEXT]\n</ul>\n\n※[TEXT]は各行が<li>...</li>に置換されます";
				break;
			case "link":
				placeholder = '例: <a href="[URL]">[TEXT]</a>';
				break;
			case "p-list":
				placeholder = "例:\n<div>\n<p>[TEXT_P]</p>\n<ul>\n[TEXT_LIST]\n</ul>\n</div>";
				break;
			case "link-list":
				placeholder = '例:\n<div class="wrapper">\n<ul>\n[LINK_LIST]\n</ul>\n</div>\n\n※[LINK_LIST]の位置に繰り返されるリンク項目が挿入されます';
				break;
			default:
				placeholder = "[TEXT] を使って雛形を入力";
		}
		tagTemplateInput.placeholder = placeholder;
	};

	// タグタイプセレクタが変更されたら、プレースホルダーを更新
	tagTypeSelector.addEventListener("change", updateTemplatePlaceholder);

	// ★★★ 修正箇所ここまで ★★★

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

		if (existingId) {
			// 編集
			const button = currentPattern.buttons.find((b) => b.id === existingId);
			if (button) {
				button.name = newName;
				button.tagType = newType;
				button.template = newTemplate;
				button.linkItemTemplate = newLinkItemTemplate;
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
			});
		}

		tagModal.style.display = "none";
		rebuildTagButtons();
		rebuildTagButtonList();
		saveAllPatterns();
	});

	// タグボタン一覧の構築
	const rebuildTagButtonList = () => {
		tagButtonList.innerHTML = "";
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];

		if (!currentPattern || currentPattern.buttons.length === 0) {
			tagButtonList.textContent = "タグボタンがありません。";
			return;
		}

		currentPattern.buttons.forEach((button) => {
			const item = document.createElement("div");
			item.className = "tag-item";

			const infoSpan = document.createElement("span");
			infoSpan.textContent = `${button.name} (${button.tagType})`;

			const editButton = document.createElement("button");
			editButton.textContent = "編集";
			editButton.addEventListener("click", () => openTagModal(button));

			const deleteButton = document.createElement("button");
			deleteButton.textContent = "削除";
			deleteButton.style.color = "red";
			deleteButton.addEventListener("click", () => {
				if (confirm(`タグ「${button.name}」を削除してもよろしいですか？`)) {
					console.log(`タグ「${button.name}」を削除しました。`);
					currentPattern.buttons = currentPattern.buttons.filter((b) => b.id !== button.id);
					rebuildTagButtons();
					rebuildTagButtonList();
					saveAllPatterns();
				}
			});

			item.appendChild(infoSpan);
			item.appendChild(editButton);
			item.appendChild(deleteButton);
			tagButtonList.appendChild(item);
		});
	};

	// --- タグボタン群の動的生成 ---

	const rebuildTagButtons = () => {
		tagButtonsContainer.innerHTML = "";
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];
		if (!currentPattern) return;

		currentPattern.buttons.forEach((button) => {
			const btn = document.createElement("button");
			btn.id = `addTagButton-${button.id}`;
			btn.textContent = button.name;

			// イベントリスナーを動的に割り当て
			btn.addEventListener("click", () =>
				createTextarea(button.id, button.tagType, button.name)
			);

			tagButtonsContainer.appendChild(btn);
		});
	};

	// --- 変換ロジックの修正 ---

	/**
	 * 選択されたパターンIDとタグIDに基づいてカスタムテンプレートとタイプを取得する関数
	 */
	const getCustomTagInfo = (tagId) => {
		const patternId = getSelectedPatternId();
		const currentPattern = patterns[patternId];
		if (!currentPattern) return null;

		return currentPattern.buttons.find((b) => b.id === tagId);
	};

	// 共通の変換ロジック (タグタイプに応じて処理を振り分けるように変更)
	function convertTextToHtmlString(textareaElement) {
		const text = textareaElement.value;
		if (!text || text.trim() === "") return ""; // 空白/空の場合は空文字を返す

		const tagId = textareaElement.getAttribute("data-tag-id");
		const tagInfo = getCustomTagInfo(tagId);

		// タグ情報がない場合は、HTMLコメントとして警告を返す
		if (!tagInfo) return `<!-- 警告: 不明なタグID (${tagId}) のためスキップされました -->\n`;

		// ★★★
		// テンプレートが空文字列の場合、警告を返すように修正
		// (修正しない場合、空文字列が返され、trim()で消えてしまうため)
		const templateString = tagInfo.template;
		if (templateString.trim() === "") {
			return `<!-- 警告: タグ (${tagInfo.name}) の雛形が空のためスキップされました -->\n`;
		}
		// ★★★

		const tagType = tagInfo.tagType;
		let output = "";

		if (tagType === "link") {
			// リンクタグ系の特別処理 (link)
			const lines = text.trim().split("\n").map((line) => line.trim());
			const linkText = lines[0] || "リンク";
			const url = lines[1] || "";

			if (url) {
				output = templateString.replace(/\[TEXT\]/g, linkText).replace(/\[URL\]/g, url);
			} else {
				output = `<!-- 警告: リンクエリアですがURLが2行目に指定されていません (${linkText}) -->\n`;
			}
		} else if (tagType === "p-list") {
			// P+リストタグの処理 (p-list)
			const lines = text.trim().split("\n");
			const pText = lines[0] ? lines[0].trim().replace(/\n/g, "<br>") : "リストのタイトル";
			const listLines = lines.slice(1).filter((line) => line.trim() !== "");

			let listContent = "";
			if (listLines.length > 0) {
				listContent = listLines.map((line) => `<li>${line.trim()}</li>`).join("\n");
			} else {
				listContent = `<li>リスト項目がありません</li>`;
			}

			output = templateString.replace(/\[TEXT_P\]/g, pText).replace(/\[TEXT_LIST\]/g, listContent);
		} else if (tagType === "list") {
			// リストタグの処理 (list)
			let listContent = text.trim().split("\n")
				.filter((line) => line.trim() !== "")
				.map((line) => {
					// すでにliタグが含まれていたらそのまま、そうでなければliタグで囲む
					return line.match(/^\s*<li/i) ? line : `<li>${line.trim()}</li>`;
				}).join("\n");

			if (listContent.trim() === "") listContent = `<li>リスト項目がありません</li>`;

			output = templateString.replace(/\[TEXT\]/g, listContent);
		} else if (tagType === "single") {
			// 単一タグの処理 (single)
			let content = text.trim().replace(/\n/g, "<br>");
			output = templateString.replace(/\[TEXT\]/g, content);
		} else if (tagType === "link-list") {
			// リンクリストの処理 (link-list)
			const lines = text.trim().split("\n").map((line) => line.trim()).filter((line) => line !== "");
			let linkListContent = "";

			// カスタムリンク項目テンプレートを取得（デフォルト値を設定）
			const linkItemTemplate = tagInfo.linkItemTemplate || '<li class="rtoc-item"><a href="[URL]">[TEXT]</a></li>';

			// 2行ずつペアにして処理 (奇数行:TEXT, 偶数行:URL)
			for (let i = 0; i < lines.length; i += 2) {
				const linkText = lines[i] || "";
				const linkUrl = lines[i + 1] || "";

				if (linkText && linkUrl) {
					// カスタムテンプレートを使用して変換
					const itemHtml = linkItemTemplate.replace(/\[TEXT\]/g, linkText).replace(/\[URL\]/g, linkUrl);
					linkListContent += itemHtml + "\n";
				} else if (linkText) {
					// URLがない場合は警告
					linkListContent += `<!-- 警告: "${linkText}" のURLが指定されていません -->\n`;
				}
			}

			if (linkListContent.trim() === "") {
				linkListContent = `<li class="rtoc-item">リンク項目がありません</li>`;
			}

			output = templateString.replace(/\[LINK_LIST\]/g, linkListContent.trim());
		}

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

	// テキストエリアを作成し、選択された位置に挿入する関数 (引数を変更)
	const createTextarea = (tagId, tagType, tagName) => {
		areaIndex++;
		const groupDiv = document.createElement("div");
		groupDiv.className = "input-area-group";
		groupDiv.id = `area-group-${areaIndex}`;

		const tagLabel = document.createElement("span");
		tagLabel.className = "input-label-tag";

		const numberPrefix = document.createElement("span");
		numberPrefix.className = "number-prefix";
		tagLabel.appendChild(numberPrefix);

		const tagTextSpan = document.createElement("span");
		tagTextSpan.textContent = tagName;
		tagLabel.appendChild(tagTextSpan);

		const newTextarea = document.createElement("textarea");
		newTextarea.className = `input-text-area input-${tagId}-area`;
		newTextarea.id = `text-area-${tagId}-${areaIndex}`;
		newTextarea.setAttribute("data-tag-id", tagId); // タグIDをデータ属性として保持

		newTextarea.rows = 3;
		newTextarea.cols = 50;
		newTextarea.style.resize = "vertical";

		let placeholderText = `${tagName}タグ用のテキストエリア`;

		if (tagType === "link") {
			placeholderText = `${tagName}エリア。1行目: 表示テキスト、2行目: リンクURL`;
		} else if (tagType === "p-list") {
			placeholderText = `${tagName}エリア。1行目: Pタグの内容、2行目以降: リスト項目`;
		} else if (tagType === "single") {
			placeholderText = `${tagName}エリア。テキストを入力してください。改行は<br>になります。`;
		} else if (tagType === "list") {
			placeholderText = `${tagName}エリア。各行がliタグに変換されます。`;
		} else if (tagType === "link-list") {
			placeholderText = `${tagName}エリア。奇数行: リンクテキスト、偶数行: URL を交互に入力`;
		}
		newTextarea.placeholder = placeholderText;

		const deleteButton = document.createElement("button");
		deleteButton.textContent = "削除";
		deleteButton.className = "delete-input-button";

		deleteButton.addEventListener("click", () => {
			groupDiv.remove();
			updateInsertionPoints();
		});

		groupDiv.appendChild(tagLabel);
		groupDiv.appendChild(newTextarea);
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
	};

	// --- ページ初期化 ---
	loadAllPatterns();
	loadShortcuts();
	buildShortcutUI(); // ショートカット管理UIを初期表示
	rebuildTagButtonList(); // タグボタン管理一覧を初期表示

	// --- その他の機能 (変更なし) ---

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

	// キーボードショートカットイベントリスナー (カスタマイズ対応版)
	document.addEventListener("keydown", function (event) {
		const isModifierKey = event.ctrlKey || event.metaKey;
		if (isModifierKey) {
			const key = event.key.toUpperCase();

			// カスタマイズされたショートカット設定を使用
			if (keyboardShortcuts[key]) {
				const tagName = keyboardShortcuts[key].tag;
				if (tagName) {
					event.preventDefault();
					window.wrapSelectionWithTag(tagName);
				}
			}
		}
	});

	// 出力エリアクリア機能 (変更なし)
	if (clearOutputButton) {
		clearOutputButton.addEventListener("click", function () {
			htmlCodeOutput.textContent = "";
		});
	}

	// コードコピー機能 (変更なし)
	if (copyOutputButton) {
		copyOutputButton.addEventListener("click", async function () {
			const textToCopy = htmlCodeOutput.textContent;
			if (textToCopy.trim().length === 0) {
				// alert("コピーするHTMLコードがありません。");
				console.warn("コピーするHTMLコードがありません。");
				return;
			}

			try {
				await navigator.clipboard.writeText(textToCopy);
				const originalText = copyOutputButton.textContent;
				copyOutputButton.textContent = "コピーしました！";
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
					copyOutputButton.textContent = "コピーしました！(FB)";
					setTimeout(() => {
						copyOutputButton.textContent = originalText;
					}, 1500);
				} catch (copyErr) {
					console.error("フォールバックコピーにも失敗しました:", copyErr);
					// alert は使用しない
				}
			}
		});
	}

	// 【修正】一括変換機能
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
