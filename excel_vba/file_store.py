"""Single-writer JSON storage for the Excel client service."""

import json
import os
import tempfile
import threading
from contextlib import contextmanager
from copy import deepcopy
from pathlib import Path

COLLECTIONS = ("groups", "products", "variants", "departments", "people", "documents",
               "movements", "versions", "counts", "users", "periods", "audit")


class FileStore:
    def __init__(self, directory):
        self.path = Path(directory).resolve() / "records.json"
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.guard = threading.RLock()
        self.lock = self.path.with_name("writer.lock").open("a+b")
        try:
            if os.name == "nt":
                import msvcrt

                self.lock.seek(0)
                msvcrt.locking(self.lock.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(self.lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
            if self.path.exists():
                with self.path.open(encoding="utf-8") as stream:
                    self.state = json.load(stream)
                if self.state.get("schema") != 1 or any(key not in self.state for key in COLLECTIONS):
                    raise ValueError("Kayıt dosyası sürümü geçersiz.")
            else:
                self.state = {"schema": 1, "sequences": {}, **{key: [] for key in COLLECTIONS}}
                self._write(self.state)
        except BaseException:
            self.lock.close()
            raise

    def close(self):
        self.lock.close()

    def read(self):
        with self.guard:
            return deepcopy(self.state)

    @contextmanager
    def update(self):
        with self.guard:
            draft = deepcopy(self.state)
            yield draft
            self._write(draft)
            self.state = draft

    def _write(self, state):
        descriptor, temporary = tempfile.mkstemp(prefix=".records-", suffix=".tmp", dir=self.path.parent)
        try:
            with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as stream:
                json.dump(state, stream, ensure_ascii=False, separators=(",", ":"))
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.path)
        finally:
            Path(temporary).unlink(missing_ok=True)


def insert(state, table, /, **fields):
    sequences = state["sequences"]
    next_id = sequences[table] = sequences.get(table, 0) + 1
    row = {"id": next_id, **fields}
    state[table].append(row)
    return row
