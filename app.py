"""Wordlister Rescoring Tool — Web UI."""

import csv
import configparser
import os
import random
from pathlib import Path
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)

CONFIG_PATH = Path(__file__).parent / "config.ini"
DEFAULT_CONFIG = {
    "master_wordlist_file": "master_wordlist.txt",
    "personal_wordlist_file": "personal_wordlist.txt",
    "rescore_tracker_file": "rescore_tracker.txt",
    "length_min": "3",
    "length_max": "15",
    "score_min": "25",
    "score_max": "60",
}

SCORE_BUCKETS = [61, 60, 50, 25, 0]
BUCKET_COLORS = {
    0: {"color": "#e74c3c", "label": "Terrible"},
    25: {"color": "#f39c12", "label": "Poor"},
    50: {"color": "#3498db", "label": "Fair"},
    60: {"color": "#2ecc71", "label": "Good"},
    61: {"color": "#9b59b6", "label": "Great"},
}


def map_score_to_bucket(score: int) -> int:
    return min(SCORE_BUCKETS, key=lambda b: abs(b - score))


def load_config() -> dict:
    config = configparser.ConfigParser()
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG_PATH.exists():
        config.read(CONFIG_PATH)
        if "General" in config:
            cfg.update(config["General"])
    return cfg


def save_config(cfg: dict):
    config = configparser.ConfigParser()
    config["General"] = cfg
    with open(CONFIG_PATH, "w") as f:
        config.write(f)


def load_wordlist(path: str) -> dict[str, int]:
    """Load a semicolon-delimited word;score file into a dict."""
    words = {}
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return words
    try:
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            for row in reader:
                if len(row) >= 2:
                    try:
                        words[row[0]] = int(row[1])
                    except ValueError:
                        continue
    except (IOError, OSError):
        pass
    return words


def load_tracker(path: str) -> dict[str, bool]:
    """Load tracker file into a dict of word -> rescored."""
    tracker = {}
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return tracker
    try:
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            for row in reader:
                if len(row) >= 2:
                    try:
                        tracker[row[0]] = int(row[1]) == 1
                    except ValueError:
                        continue
    except (IOError, OSError):
        pass
    return tracker


def save_wordlist(path: str, words: dict[str, int]):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        for word, score in words.items():
            writer.writerow([word, score])


def save_tracker(path: str, tracker: dict[str, bool]):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        for word, rescored in tracker.items():
            writer.writerow([word, 1 if rescored else 0])


class AppState:
    """Holds all in-memory state for the rescoring session."""

    def __init__(self):
        self.config: dict = {}
        self.master: dict[str, int] = {}
        self.personal: dict[str, int] = {}
        self.tracker: dict[str, bool] = {}
        self.queue: list[dict] = []  # [{word, score}] — shuffled, filtered, unscored
        self.current_index: int = 0
        self.history: list[dict] = []  # [{word, old_score, new_score, original_personal}]
        self.total_filtered: int = 0
        self.reload()

    def reload(self):
        """Load/reload all data from disk and rebuild the queue."""
        self.config = load_config()
        master_path = self.config["master_wordlist_file"]
        personal_path = self.config["personal_wordlist_file"]
        tracker_path = self.config["rescore_tracker_file"]

        self.master = load_wordlist(master_path)
        self.personal = load_wordlist(personal_path)
        self.tracker = load_tracker(tracker_path)
        self.history = []
        self._build_queue()

    def _build_queue(self):
        """Filter, bucket-map, remove already-scored, shuffle."""
        cfg = self.config
        len_min = int(cfg["length_min"])
        len_max = int(cfg["length_max"])
        score_min = int(cfg["score_min"])
        score_max = int(cfg["score_max"])

        filtered = []
        for word, score in self.master.items():
            # Apply personal score overlay
            effective_score = self.personal.get(word, score)
            if len_min <= len(word) <= len_max and score_min <= effective_score <= score_max:
                filtered.append({"word": word, "score": map_score_to_bucket(effective_score)})

        self.total_filtered = len(filtered)

        # Remove already-rescored words
        self.queue = [w for w in filtered if not self.tracker.get(w["word"], False)]
        random.shuffle(self.queue)
        self.current_index = 0

    def done_count(self) -> int:
        """Count how many of the filtered words have been rescored."""
        return self.total_filtered - len(self.queue) + self.current_index

    def current_word(self) -> dict | None:
        if self.current_index < len(self.queue):
            return self.queue[self.current_index]
        return None

    def rescore(self, action: str) -> dict | None:
        """Apply a rescoring action. Returns the result or None if queue exhausted."""
        current = self.current_word()
        if current is None:
            return None

        word = current["word"]
        old_score = current["score"]
        new_score = self._compute_new_score(old_score, action)

        # Store original personal score for proper undo
        original_personal = self.personal.get(word)

        # Update in-memory state
        self.personal[word] = new_score
        self.tracker[word] = True

        # Push to history (capped at 20)
        self.history.append({
            "word": word,
            "old_score": old_score,
            "new_score": new_score,
            "original_personal": original_personal,
            "queue_index": self.current_index,
        })
        if len(self.history) > 20:
            self.history.pop(0)

        self.current_index += 1

        return {
            "word": word,
            "old_score": old_score,
            "new_score": new_score,
            "action": action,
        }

    def undo(self) -> dict | None:
        """Undo the last rescoring action. Returns restored word or None."""
        if not self.history:
            return None

        entry = self.history.pop()
        word = entry["word"]

        # Restore personal score to what it was before this action
        if entry["original_personal"] is None:
            self.personal.pop(word, None)
        else:
            self.personal[word] = entry["original_personal"]

        # Restore tracker
        self.tracker[word] = False

        # Move index back
        self.current_index = entry["queue_index"]
        # Restore the score in the queue to old_score
        self.queue[self.current_index]["score"] = entry["old_score"]

        return {"word": word, "score": entry["old_score"]}

    def save_to_disk(self):
        save_wordlist(self.config["personal_wordlist_file"], self.personal)
        save_tracker(self.config["rescore_tracker_file"], self.tracker)

    @staticmethod
    def _compute_new_score(old_score: int, action: str) -> int:
        idx = SCORE_BUCKETS.index(old_score)
        if action == "increase":
            idx = max(idx - 1, 0)
        elif action == "decrease":
            idx = min(idx + 1, len(SCORE_BUCKETS) - 1)
        elif action == "increase_double":
            idx = max(idx - 2, 0)
        elif action == "decrease_double":
            idx = min(idx + 2, len(SCORE_BUCKETS) - 1)
        return SCORE_BUCKETS[idx]


state = AppState()


# --- Routes ---

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/state")
def get_state():
    """Return current word, progress, and bucket info."""
    current = state.current_word()
    return jsonify({
        "current": current,
        "done": state.done_count(),
        "total": state.total_filtered,
        "remaining": len(state.queue) - state.current_index,
        "history_size": len(state.history),
        "buckets": {str(k): v for k, v in BUCKET_COLORS.items()},
        "bucket_order": SCORE_BUCKETS,
    })


@app.route("/api/rescore", methods=["POST"])
def rescore():
    action = request.json.get("action")
    if action not in ("increase", "decrease", "keep", "increase_double", "decrease_double"):
        return jsonify({"error": "Invalid action"}), 400
    result = state.rescore(action)
    if result is None:
        return jsonify({"done": True, "total": state.total_filtered})
    current = state.current_word()
    return jsonify({
        "result": result,
        "next": current,
        "done_count": state.done_count(),
        "total": state.total_filtered,
        "remaining": len(state.queue) - state.current_index,
        "history_size": len(state.history),
    })


@app.route("/api/undo", methods=["POST"])
def undo():
    result = state.undo()
    if result is None:
        return jsonify({"error": "Nothing to undo"}), 400
    return jsonify({
        "restored": result,
        "done_count": state.done_count(),
        "total": state.total_filtered,
        "remaining": len(state.queue) - state.current_index,
        "history_size": len(state.history),
    })


@app.route("/api/save", methods=["POST"])
def save():
    try:
        state.save_to_disk()
    except (IOError, OSError) as e:
        return jsonify({"error": str(e)}), 500
    return jsonify({"ok": True})


@app.route("/api/settings", methods=["GET"])
def get_settings():
    return jsonify(state.config)


@app.route("/api/settings", methods=["POST"])
def update_settings():
    new_cfg = request.json
    # Validate required keys
    for key in DEFAULT_CONFIG:
        if key in new_cfg:
            state.config[key] = str(new_cfg[key])
    save_config(state.config)
    state.reload()
    return jsonify({"ok": True, "total": state.total_filtered, "remaining": len(state.queue)})


if __name__ == "__main__":
    app.run(debug=True, port=5000)
