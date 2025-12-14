from sqlalchemy import UUID, Uuid, String
from sqlalchemy.orm import mapped_column, Mapped
import uuid

from src.domain.base import BaseModel
from src.domain.choices import FileStatus
from src.domain.models.mixins import TimestampMixin


class File(TimestampMixin, BaseModel):
    __tablename__ = "file"

    uuid: Mapped[UUID] = mapped_column(Uuid(as_uuid=True), primary_key=True, default=uuid.uuid4, unique=True)
    status: Mapped[FileStatus] = mapped_column(String(50), nullable=False)
    original_s3_link: Mapped[String] = mapped_column(String, nullable=False)
    processed_s3_link: Mapped[String] = mapped_column(String, nullable=True, default=None)
