# Wordlister Rescoring Tool — Assessment

## What works well
- The core interaction loop is sound: keyboard-driven, one-word-at-a-time, with undo — this is the right paradigm for bulk rescoring.
- In-memory batch saves are a good design choice — not hammering disk on every keystroke.
- The ticker providing visual feedback of recent actions is a nice touch.

## Architecture & Code Quality Issues

### 1. Pandas is the wrong data store
DataFrames are used as glorified dictionaries. Every `update_personal_in_memory` and `update_tracker_in_memory` call does a linear scan (`word in self.personal_df['word'].values`), then a `.loc` filter, then potentially a `pd.concat` to append a single row. This is O(n) per keystroke. A plain `dict[str, int]` would be O(1) and dramatically simpler. Pandas only makes sense for the initial CSV load and final write-out.

### 2. Slow done_words rebuild
The done_words rebuild (line 222-223) iterates every row of the tracker DataFrame on startup and after every settings change. With a large wordlist this is painfully slow.

### 3. Massive code duplication
`open_settings_dialog` (lines 581-619) is a copy-paste of `__init__` lines 186-223. All the settings-reading, file-loading, filter-applying, done_words-rebuilding logic is duplicated. Should be one `reload()` method called from both places.

### 4. Hardcoded Windows path as fallback
`r"C:\Users\Dennis\OneDrive\XwiWordList.txt"` — a personal dev path baked into the code. Anyone else running this hits a confusing error.

### 5. Font size setting is completely broken
```python
self.main_word_font_size = int(self.settings.value("main_word_font_size", 32))
self.main_word_font_size = 64  # immediately overwritten
```
The user can configure the font size in settings, it gets saved, loaded... and then hard-overridden to 64. Appears in two places.

### 6. Confusing score bucket model
`score_buckets = [61, 60, 50, 25, 0]` — the values are arbitrary and the descending order makes the index math unintuitive. The 61 vs 60 distinction is unclear. Consider named tiers or constants.

### 7. Slow apply_personal_scores
Uses `.apply()` with a lambda over every row — the slowest possible way to do a column update in pandas. A `.map()` or merge would be faster and cleaner.

### 8. No error handling on file I/O
If the semicolon-delimited file has a malformed line, the whole app crashes. No try/except around any `pd.read_csv` calls.

### 9. Fragile scoring_in_progress lock
Set to `True` in `rescore_word`, then `False` in `show_next_word` via `QTimer.singleShot`. If anything goes wrong in that chain, the app locks up permanently — keyboard input is silently ignored with no feedback.

## UI/UX Issues

### 10. PyQt5 is abandoned
No meaningful release since 2021. No security patches or bug fixes. If sticking with Qt, PyQt6 or PySide6 is the path.

### 11. A web UI would be a better fit
The entire interaction is: show a word, receive a keypress, show the next word. This is a perfect fit for a minimal web app (Flask/FastAPI + vanilla HTML/JS). Benefits:
- Cross-platform immediately (no Windows-only `QApplication.setStyle("windows")`)
- No dependency on Qt binaries
- Easier to style (CSS vs Qt stylesheets)
- Could be deployed for others to use
- Simpler keyboard handling

### 12. Non-obvious keybindings
D=increase, A=decrease maps to right/left spatially, but Q/E for "double" is non-obvious. No visual legend showing the bucket scale, so users have no mental model of where a word sits or where it's going.

### 13. No skip/already-done filtering
The app shows words linearly by incrementing `current_index`. Already-rescored words still get shown — the user has to press Space to skip. `done_words` is only used for the progress counter, not for filtering.

### 14. Competing timers
After scoring, the word stays visible for 200ms (disappear delay), then vanishes. But the new score flash also has a hardcoded 200ms timer. These race — the word can disappear at the same time as the score flash, making feedback feel janky.

### 15. No confirmation on exit without save
Closing via the X button doesn't trigger `export_and_exit`. All unsaved work is silently lost. No `closeEvent` override.

### 16. Hardcoded geometry
1200x900 is not responsive. No consideration for different screen sizes or DPI.

## Data Integrity Issues

### 17. Undo doesn't fully restore state
When undoing, the tracker is set back to `rescored=0`, but if the word was already in the personal wordlist before this session, undo sets its score to `old_score` — which might be the bucket-mapped score, not the original personal score. The bucket mapping in `apply_filters` is lossy (48 becomes 50), so undo can corrupt data.

### 18. Shuffle on every reload
The shuffle happens on every settings change. Combined with `current_index = 0`, the user re-reviews words they already scored this session. The shuffle should happen once, and the index should be preserved across settings changes.

## Recommended Path Forward

1. **Replace pandas with plain dicts** for all in-memory state. Use pandas only for CSV I/O.
2. **Skip already-scored words** in `show_next_word` instead of showing them.
3. **Move to a web frontend** — Flask + a single HTML page with keyboard listeners. The server holds state, the browser handles display.
4. **Add a `closeEvent`** override that prompts to save.
5. **Fix the font size override**, the hardcoded Windows path, and the dual-timer race.
6. **Extract the duplicated settings/reload logic** into one method.
7. **Add basic error handling** around file I/O.

The bones are good. The idea, the interaction model, the data flow — all sound. It's the implementation that needs a rewrite more than a refactor.
