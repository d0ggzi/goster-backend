import os
import uuid
from pathlib import Path

STORAGE_DIR = Path("storage")


class LocalStorageService:
    def __init__(self):
        STORAGE_DIR.mkdir(exist_ok=True)

    def upload_original(self, user_id: str, filename: str, content: bytes) -> str:
        key = f"{user_id}/original/{uuid.uuid4()}_{filename}"
        path = STORAGE_DIR / key
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)
        return key

    def upload_fixed(self, user_id: str, filename: str, content: bytes) -> str:
        key = f"{user_id}/fixed/{uuid.uuid4()}_{filename}"
        path = STORAGE_DIR / key
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)
        return key

    def download_file(self, key: str) -> bytes:
        path = STORAGE_DIR / key
        return path.read_bytes()

    def delete_file(self, key: str) -> None:
        path = STORAGE_DIR / key
        if path.exists():
            path.unlink()


storage_service = LocalStorageService()
