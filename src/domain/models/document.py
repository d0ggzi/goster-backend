import uuid
from datetime import datetime

from sqlalchemy import String, DateTime, ForeignKey, func
from sqlalchemy.dialects.postgresql import UUID as PG_UUID
from sqlalchemy.orm import mapped_column, Mapped, relationship

from src.domain.base import BaseModel


class Document(BaseModel):
    __tablename__ = "document"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    user_id: Mapped[str] = mapped_column(ForeignKey("user.id"), nullable=False)

    filename: Mapped[str] = mapped_column(String(255), nullable=False)
    original_s3_key: Mapped[str] = mapped_column(String(500), nullable=False)
    fixed_s3_key: Mapped[str] = mapped_column(String(500), nullable=False)

    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now(), nullable=False)
