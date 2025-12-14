from datetime import datetime, timezone

from sqlalchemy import DateTime, Column, func


class TimestampMixin:
    created_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())
    deleted_at = Column(DateTime(timezone=True), nullable=True)

    def set_deleted(self):
        self.deleted_at = datetime.now(timezone.utc)
