# Block Editor KF - Refactoring Documentation

## ⚠️ STATUS: REVERTED

**The modularization has been reverted as of 2025-12-05.**

The application has been restored to use the original single `script.js` file due to issues encountered with the modular architecture. All modular JS files have been removed, and the application now uses the original monolithic structure.

---

## Overview (Historical - Modularization Attempt)

The monolithic `script.js` file (~2900 lines, 121KB) was temporarily refactored into a modular architecture with a shared global namespace `window.BlockEditor`, but this has been reverted.

## Module Structure

### Core Modules (Fully Modularized)

1. **config.js** (94 lines)
   - Application constants (CONSTANTS object)
   - Storage keys, size limits, messages, placeholders, styles, CSS classes
   - All magic numbers and strings centralized

2. **state.js** (55 lines)
   - Global application state management
   - Variables: areaIndex, maxPatternId, maxTagId, isLoadingInputAreas, patterns, formattingMap

3. **dom.js** (95 lines)
   - DOM element caching for performance
   - `initDOM()` function initializes and caches all frequently accessed elements

4. **utils.js** (94 lines)
   - Utility functions
   - `debounce()` - Debounce function for performance
   - `setupDragAndDrop()` - Common drag & drop event handler setup

5. **validation.js** (206 lines)
   - Template validation and HTML sanitization
   - `validateTemplate()` - Validates tag templates with security checks
   - `sanitizeHTML()` - Sanitizes HTML content to prevent XSS

6. **storage.js** (290 lines)
   - LocalStorage data persistence
   - Default patterns and formatting maps
   - Save/load functions for patterns, formatting, and input areas
   - Error handling for storage quota and private browsing

### Legacy Code

7. **legacy-functions.js** (~2657 lines)
   - Wrapped in `document.addEventListener("DOMContentLoaded", function () { ... });`
   - Contains remaining code not yet fully modularized:
     - UI building functions (section 6)
     - Pattern management (section 7)
     - Tag button management (section 8)
     - Input area management (section 9)
     - Conversion logic (section 10)
     - Event handlers (section 11)
     - Data sharing/export/import (section 13)
   - **Important**: Must load AFTER main.js to ensure globals are available

8. **main.js** (84 lines)
   - Application initialization
   - Sets up global namespace and backward compatibility bridges
   - Coordinates module loading

## Loading Order (index.html)

```html
<!-- Core modules (load first) -->
<script src="js/config.js"></script>
<script src="js/state.js"></script>
<script src="js/dom.js"></script>
<script src="js/utils.js"></script>
<script src="js/validation.js"></script>
<script src="js/storage.js"></script>

<!-- Main initialization (sets up globals) -->
<script src="js/main.js"></script>

<!-- Legacy functions (executes after globals are set up) -->
<script src="js/legacy-functions.js"></script>
```

**IMPORTANT**: The loading order is critical:
1. Core modules load and export to `window.BlockEditor` namespace
2. **main.js** loads and registers a `DOMContentLoaded` handler (registers FIRST)
3. **legacy-functions.js** loads and registers a `DOMContentLoaded` handler (registers SECOND)
4. When DOM loads, handlers execute in registration order:
   - main.js runs first → sets up global variables and backward compatibility aliases
   - legacy-functions.js runs second → uses the globals that were just set up

## Namespace Structure

All modules export to `window.BlockEditor`:

- `window.BlockEditor.CONSTANTS` - Configuration constants
- `window.BlockEditor.state` - Application state
- `window.BlockEditor.DOM` - Cached DOM elements
- `window.BlockEditor.utils` - Utility functions
- `window.BlockEditor.validation` - Validation functions
- `window.BlockEditor.storage` - Storage functions
- `window.BlockEditor.initDOM()` - DOM initialization function

## Backward Compatibility

The `main.js` file creates global aliases for backward compatibility with legacy code:
- All CONSTANTS, state variables, and functions are exposed to `window`
- This allows legacy-functions.js to work without modification

## Benefits of Refactoring

1. **Better Organization** - Code is logically separated by concern
2. **Improved Maintainability** - Easier to find and modify specific functionality
3. **Reusability** - Modules can be used independently
4. **Testing** - Individual modules can be tested in isolation
5. **Performance** - Only load what's needed (future optimization)
6. **Scalability** - Easier to add new features

## Future Refactoring Tasks

The following sections in `legacy-functions.js` should be extracted to separate modules:

1. **ui.js** - UI building functions
   - `buildFormattingUI()`
   - `updateExistingInputAreas()`
   - `loadAllPatterns()`

2. **patterns.js** - Pattern management
   - `getSelectedPatternId()`
   - `rebuildPatternUI()`
   - `updateDeleteButtonState()`
   - Pattern CRUD operations

3. **tags.js** - Tag button management
   - `openTagModal()`
   - `rebuildTagButtonList()`
   - `rebuildTagButtons()`
   - Tag CRUD operations

4. **inputAreas.js** - Input area management
   - `createTextarea()`
   - `updateInsertionPoints()`
   - Drag & drop for input areas

5. **converter.js** - Text to HTML conversion
   - `convertTextToHtmlString()`
   - `getTextWithLineBreaks()`
   - `cleanupHTML()`

6. **dataSharing.js** - Export/Import functionality
   - Encryption/decryption
   - Export modal handling
   - Import modal handling

## Testing Checklist

- [ ] All patterns load correctly
- [ ] Pattern creation/deletion works
- [ ] Tag buttons are functional
- [ ] Input areas can be created and edited
- [ ] Conversion logic works correctly
- [ ] Export/import functionality works
- [ ] Formatting maps are applied correctly
- [ ] LocalStorage persistence works
- [ ] No console errors on page load
- [ ] Drag & drop works for input areas and tags

## Backup

The original monolithic script has been backed up to `script.js.backup` (121KB, 3392 lines).

## Notes

- This is a hybrid approach: core functionality is modularized, while complex interdependent code remains in legacy-functions.js
- The architecture allows for gradual refactoring - functions can be moved from legacy-functions.js to dedicated modules one at a time
- All functionality from the original script.js is preserved
