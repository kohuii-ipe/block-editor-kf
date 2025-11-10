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

	// 書式マッピング管理関連
	const formattingList = document.getElementById("formattingList");

	let areaIndex = 0;
	let maxPatternId = 0;
	let maxTagId = 0;
	let isLoadingInputAreas = false; // 入力エリア読み込み中フラグ

	// --- 雛形設定のデフォルトデータ（構造変更） ---

	// デフォルトの書式マッピング設定（Word貼り付け用）
	const defaultFormattingMap = {
		bold: { tag: "strong", displayName: "太字 (Bold)" },
		italic: { tag: "em", displayName: "斜体 (Italic)" },
		underline: { tag: "u", displayName: "下線 (Underline)" },
		highlight: { tag: "mark", displayName: "ハイライト (Highlight)" },
	};

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
			template: '<div class="box"><p>[TEXT_P_1]</p><p>[TEXT_P_2]</p><ul>\n[TEXT_LIST]\n</ul></div>',
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
			formattingMap: JSON.parse(JSON.stringify(defaultFormattingMap)),
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
			formattingMap: JSON.parse(JSON.stringify(defaultFormattingMap)),
		},
	};

	let patterns = {}; // すべてのパターン設定を保持するグローバルオブジェクト

	// --- 永続化と初期化 ---
	const STORAGE_KEY = "customTagPatternsV2";
	const SETTINGS_KEY = "editorSettingsV2";
	const SHORTCUTS_KEY = "keyboardShortcutsV2";
	const INPUT_AREAS_KEY = "inputAreasV2";

	// デフォルトのショートカット設定
	const defaultShortcuts = {
		B: { tag: "strong", displayName: "Ctrl/Cmd+B" },
		M: { tag: "mark", displayName: "Ctrl/Cmd+M" },
		U: { tag: "sup", displayName: "Ctrl/Cmd+U" },
		L: { tag: "li", displayName: "Ctrl/Cmd+L" },
		K: { tag: "a", displayName: "Ctrl/Cmd+K" },
	};

	let keyboardShortcuts = {};
	let formattingMap = {};

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

		localStorage.setItem(INPUT_AREAS_KEY, JSON.stringify(areas));
	};

	const loadInputAreas = () => {
		const saved = localStorage.getItem(INPUT_AREAS_KEY);
		if (!saved) return;

		try {
			isLoadingInputAreas = true; // 読み込み開始
			const areas = JSON.parse(saved);
			areas.forEach((area) => {
				const tagInfo = getCustomTagInfo(area.tagId);
				if (tagInfo) {
					createTextarea(area.tagId, tagInfo.tagType, tagInfo.name, area.content);
				}
			});
		} catch (e) {
			console.error("入力エリアの読み込みに失敗しました:", e);
		} finally {
			isLoadingInputAreas = false; // 読み込み終了
		}
	};

	// ショートカット管理UIの構築
	const buildShortcutUI = () => {
		const shortcutList = document.getElementById("shortcutList");
		if (!shortcutList) return;

		shortcutList.innerHTML = "";

		const keys = Object.keys(keyboardShortcuts);

		keys.forEach((key, index) => {
			const shortcut = keyboardShortcuts[key];
			const item = document.createElement("div");
			item.className = "shortcut-item";
			item.draggable = true;
			item.dataset.key = key;
			item.dataset.index = index;

			// ドラッグハンドル
			const dragHandle = document.createElement("span");
			dragHandle.textContent = "⋮⋮";
			dragHandle.style.cursor = "grab";
			dragHandle.style.marginRight = "8px";
			dragHandle.style.color = "#999";
			dragHandle.style.fontSize = "0.9em";
			dragHandle.title = "ドラッグして順序を変更";

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

			// ドラッグイベント
			item.addEventListener("dragstart", (e) => {
				e.dataTransfer.effectAllowed = "move";
				e.dataTransfer.setData("text/html", e.target.innerHTML);
				item.classList.add("dragging-shortcut");
			});

			item.addEventListener("dragend", (e) => {
				item.classList.remove("dragging-shortcut");
			});

			item.addEventListener("dragover", (e) => {
				e.preventDefault();
				e.dataTransfer.dropEffect = "move";

				const draggingItem = document.querySelector(".dragging-shortcut");
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
			});

			item.addEventListener("dragleave", (e) => {
				item.style.borderTop = "";
				item.style.borderBottom = "";
			});

			item.addEventListener("drop", (e) => {
				e.preventDefault();
				item.style.borderTop = "";
				item.style.borderBottom = "";

				const draggingItem = document.querySelector(".dragging-shortcut");
				if (!draggingItem || draggingItem === item) return;

				const fromIndex = parseInt(draggingItem.dataset.index);
				const toIndex = parseInt(item.dataset.index);

				// オブジェクトを配列に変換して並び替え
				const entries = Object.entries(keyboardShortcuts);
				const [movedEntry] = entries.splice(fromIndex, 1);
				entries.splice(toIndex, 0, movedEntry);

				// 新しいオブジェクトを再構築
				keyboardShortcuts = {};
				entries.forEach(([k, v]) => {
					keyboardShortcuts[k] = v;
				});

				buildShortcutUI();
				saveShortcuts();
			});

			item.appendChild(dragHandle);
			item.appendChild(label);
			item.appendChild(input);
			item.appendChild(deleteButton);
			shortcutList.appendChild(item);
		});
	};

	// 書式マッピング管理UIの構築
	const buildFormattingUI = () => {
		if (!formattingList) return;

		formattingList.innerHTML = "";

		const types = Object.keys(formattingMap);

		types.forEach((type) => {
			const formatting = formattingMap[type];
			const item = document.createElement("div");
			item.className = "shortcut-item"; // 同じスタイルを使用

			const label = document.createElement("label");
			label.textContent = formatting.displayName;

			const input = document.createElement("input");
			input.type = "text";
			input.value = formatting.tag;
			input.placeholder = "タグ名を入力 (例: strong, em, mark)";

			input.addEventListener("input", (e) => {
				formattingMap[type].tag = e.target.value.trim();
				saveFormattingMap();
			});

			item.appendChild(label);
			item.appendChild(input);
			formattingList.appendChild(item);
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
			} catch (e) {
				patterns = defaultPatterns;
			}
		} else {
			patterns = defaultPatterns;
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
			// 後方互換性：ショートカットがない古いパターンにデフォルトショートカットを追加
			if (!patterns[id].shortcuts) {
				patterns[id].shortcuts = JSON.parse(JSON.stringify(defaultShortcuts));
			}
			// 後方互換性：書式マッピングがない古いパターンにデフォルト書式マッピングを追加
			if (!patterns[id].formattingMap) {
				patterns[id].formattingMap = JSON.parse(JSON.stringify(defaultFormattingMap));
			}
		});

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
				loadFormattingMap(); // 書式マッピングを再読み込み
				buildFormattingUI(); // 書式マッピングUIを再構築
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
			// 書式マッピングもコピー
			formattingMap: JSON.parse(JSON.stringify(sourcePattern.formattingMap || defaultFormattingMap)),
		};

		rebuildPatternUI(newId);
		rebuildTagButtons();
		rebuildTagButtonList();
		loadShortcuts(); // 新しいパターンのショートカットを読み込み
		buildShortcutUI(); // ショートカットUIを再構築
		loadFormattingMap(); // 新しいパターンの書式マッピングを読み込み
		buildFormattingUI(); // 書式マッピングUIを再構築
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
			loadFormattingMap(); // 新しいパターンの書式マッピングを読み込み
			buildFormattingUI(); // 書式マッピングUIを再構築
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
			case "multi":
				placeholder = '例: <p class="item">[TEXT]</p>\n\n※各行が個別のタグとして出力されます';
				break;
			case "list":
				placeholder = "例:\n<ul>\n[TEXT]\n</ul>\n\n※[TEXT]は各行が<li>...</li>に置換されます";
				break;
			case "link":
				placeholder = '例: <a href="[URL]">[TEXT]</a>';
				break;
			case "p-list":
				placeholder = "例:\n<div>\n<p>[TEXT_P_1]</p>\n<p>[TEXT_P_2]</p>\n<ul>\n[TEXT_LIST]\n</ul>\n</div>\n\n※[TEXT_P_1], [TEXT_P_2]...で複数段落、[TEXT_LIST]でリスト項目を配置";
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

			// ドラッグイベント
			item.addEventListener("dragstart", (e) => {
				e.dataTransfer.effectAllowed = "move";
				e.dataTransfer.setData("text/html", e.target.innerHTML);
				item.classList.add("dragging");
			});

			item.addEventListener("dragend", (e) => {
				item.classList.remove("dragging");
			});

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
			});

			item.addEventListener("dragleave", (e) => {
				item.style.borderTop = "";
				item.style.borderBottom = "";
			});

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
			});

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

			item.appendChild(dragHandle);
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

		currentPattern.buttons.forEach((button, index) => {
			const btn = document.createElement("button");
			btn.id = `addTagButton-${button.id}`;
			btn.textContent = button.name;
			btn.draggable = true;
			btn.dataset.index = index;
			btn.style.cursor = "grab";

			// クリックイベント
			btn.addEventListener("click", () =>
				createTextarea(button.id, button.tagType, button.name)
			);

			// ドラッグイベント
			btn.addEventListener("dragstart", (e) => {
				e.dataTransfer.effectAllowed = "move";
				e.dataTransfer.setData("text/html", e.target.innerHTML);
				btn.classList.add("dragging-button");
				btn.style.cursor = "grabbing";
			});

			btn.addEventListener("dragend", (e) => {
				btn.classList.remove("dragging-button");
				btn.style.cursor = "grab";
			});

			btn.addEventListener("dragover", (e) => {
				e.preventDefault();
				e.dataTransfer.dropEffect = "move";

				const draggingButton = document.querySelector(".dragging-button");
				if (draggingButton && draggingButton !== btn) {
					btn.style.borderLeft = "3px solid #4CAF50";
				}
			});

			btn.addEventListener("dragleave", (e) => {
				btn.style.borderLeft = "";
			});

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
			});

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

	// contenteditable div から改行を保持したテキストを取得するヘルパー関数
	// formatting タグ（strong, mark）は保持し、block要素は改行に変換
	function getTextWithLineBreaks(htmlContent) {
		const tempDiv = document.createElement("div");
		tempDiv.innerHTML = htmlContent;

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
				// formatting タグは保持
				const isFormatting = ["strong", "mark", "b", "em", "u", "span"].includes(tagName);

				let result = "";

				// formatting タグの開始タグを追加
				if (isFormatting) {
					result += `<${tagName}>`;
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

	// 共通の変換ロジック (タグタイプに応じて処理を振り分けるように変更)
	function convertTextToHtmlString(textareaElement) {
		const htmlContent = textareaElement.innerHTML;
		if (!htmlContent || htmlContent.trim() === "") return ""; // 空白/空の場合は空文字を返す

		// HTMLから改行を保持したテキストを取得
		const text = getTextWithLineBreaks(htmlContent);

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
				listContent = listLines.map((line) => {
				// すでにliタグが含まれていたらそのまま、そうでなければliタグで囲む
				return line.match(/^\s*<li/i) ? line : `<li>${line}</li>`;
			}).join("\n");
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
		} else if (tagType === "multi") {
			// マルチラインタグの処理 (multi) - 各行が個別のタグになる
			const lines = text.trim().split("\n").filter((line) => line.trim() !== "");

			if (lines.length > 0) {
				output = lines.map((line) => {
					return templateString.replace(/\[TEXT\]/g, line.trim());
				}).join("\n");
			} else {
				output = `<!-- 警告: テキストが入力されていません -->\n`;
			}
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

	// Word書式を保持した貼り付けの処理
	const handleFormattedPaste = (event, editableDiv) => {
		const clipboardData = event.clipboardData || window.clipboardData;
		if (!clipboardData) return;

		// HTMLデータを取得
		const htmlData = clipboardData.getData("text/html");
		if (!htmlData) return; // HTMLデータがない場合は通常の貼り付けを許可

		event.preventDefault(); // デフォルトの貼り付けをキャンセル

		// HTMLをパースして変換
		const convertedHTML = convertWordHTMLToTags(htmlData);

		// contenteditable の場合、カーソル位置にHTMLを挿入
		const selection = window.getSelection();
		if (!selection.rangeCount) return;

		const range = selection.getRangeAt(0);
		range.deleteContents();

		// HTMLフラグメントを作成して挿入
		const tempDiv = document.createElement("div");
		tempDiv.innerHTML = convertedHTML;
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
				// 太字とマーカーのみを保持し、他の書式は無視する
				if ((tagName === "b" || tagName === "strong") && formattingMap.bold) {
					openTag = `<${formattingMap.bold.tag}>`;
					closeTag = `</${formattingMap.bold.tag}>`;
				} else if ((tagName === "mark" || (tagName === "span" && node.style.backgroundColor)) && formattingMap.highlight) {
					// ハイライト（背景色がある場合も含む）
					// Word のハイライトは span の background-color として来ることがある
					openTag = `<${formattingMap.highlight.tag}>`;
					closeTag = `</${formattingMap.highlight.tag}>`;
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

		// ドラッグイベント（ラベルのみ）
		tagLabel.addEventListener("dragstart", (e) => {
			e.dataTransfer.effectAllowed = "move";
			e.dataTransfer.setData("text/html", groupDiv.id);
			groupDiv.classList.add("dragging-input-area");
			tagLabel.style.cursor = "grabbing";
		});

		tagLabel.addEventListener("dragend", (e) => {
			groupDiv.classList.remove("dragging-input-area");
			tagLabel.style.cursor = "grab";
		});

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
		});

		groupDiv.addEventListener("dragleave", (e) => {
			groupDiv.style.borderTop = "";
			groupDiv.style.borderBottom = "";
		});

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
		});

		const newTextarea = document.createElement("div");
		newTextarea.className = `input-text-area input-${tagId}-area`;
		newTextarea.id = `text-area-${tagId}-${areaIndex}`;
		newTextarea.setAttribute("contenteditable", "true");
		newTextarea.setAttribute("data-tag-id", tagId); // タグIDをデータ属性として保持
		// 初期コンテンツを設定（空の場合は完全に空にしてプレースホルダーを表示）
		if (initialContent && initialContent.trim() !== "") {
			newTextarea.innerHTML = initialContent;
		}

		// テキストエリアの内容が変更されたら保存
		newTextarea.addEventListener("input", () => {
			saveInputAreas();
		});

		// 貼り付けイベント：Word書式を保持
		newTextarea.addEventListener("paste", (e) => {
			handleFormattedPaste(e, newTextarea);
		});

		let placeholderText = `${tagName}タグ用のテキストエリア`;

		if (tagType === "link") {
			placeholderText = `${tagName}エリア。1行目: 表示テキスト、2行目: リンクURL`;
		} else if (tagType === "p-list") {
			placeholderText = `${tagName}エリア。通常の行: 段落、「-」または「*」で始まる行: リスト項目`;
		} else if (tagType === "single") {
			placeholderText = `${tagName}エリア。テキストを入力してください。改行は<br>になります。`;
		} else if (tagType === "multi") {
			placeholderText = `${tagName}エリア。各行が個別のタグとして出力されます。`;
		} else if (tagType === "list") {
			placeholderText = `${tagName}エリア。各行がliタグに変換されます。`;
		} else if (tagType === "link-list") {
			placeholderText = `${tagName}エリア。奇数行: リンクテキスト、偶数行: URL を交互に入力`;
		}
		newTextarea.setAttribute("data-placeholder", placeholderText);

		const deleteButton = document.createElement("button");
		deleteButton.textContent = "削除";
		deleteButton.className = "delete-input-button";

		deleteButton.addEventListener("click", () => {
			groupDiv.remove();
			updateInsertionPoints();
			saveInputAreas(); // 削除後に保存
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
		saveInputAreas(); // 作成後に保存
	};

	// --- ページ初期化 ---
	loadAllPatterns();
	loadShortcuts();
	buildShortcutUI(); // ショートカット管理UIを初期表示
	loadFormattingMap();
	buildFormattingUI(); // 書式マッピング管理UIを初期表示
	rebuildTagButtonList(); // タグボタン管理一覧を初期表示
	loadInputAreas(); // 入力エリアを復元

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
