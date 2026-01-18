import enum


class FileStatus(enum.Enum):
    NEW = "NEW"
    PROCESSING = "PROCESSING"
    FAILED = "FAILED"
    SUCCESS = "SUCCESS"
