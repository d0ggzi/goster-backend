import tempfile
import os
from pathlib import Path
from typing import BinaryIO

from sqlalchemy.ext.asyncio import AsyncSession

from goster.core import DocumentContext, ValidationPipeline
from goster.rules import (
    FontRule,
    AlignmentRule,
    IndentRule,
    SpacingRule,
    ParagraphSpacingRule,
    TextColorRule,
    NoUnderlineRule,
    ListFormatRule,
    HeadingNumberingRule,
    HeadingCaseRule,
    HeadingAlignmentRule,
    HeadingPunctuationRule,
    HeadingBoldRule,
    HeadingNewPageRule,
    SubheadingCaseRule,
    RequiredSectionsRule,
    TableReferenceRule,
    FigureReferenceRule,
    TableCaptionRule,
    FigureCaptionRule,
    FormulaReferenceRule,
    CitationFormatRule,
    MarginsRule,
    PageNumberingRule,
    AppendixRule,
)
from src.domain.models import Document
from src.service.s3 import s3_service


class ProcessingService:
    def __init__(self, session: AsyncSession):
        self.session = session

    def _create_pipeline(self) -> ValidationPipeline:
        pipeline = ValidationPipeline()
        pipeline.add_rules(
            MarginsRule(),
            FontRule(),
            TextColorRule(),
            AlignmentRule(),
            IndentRule(),
            SpacingRule(),
            ParagraphSpacingRule(),
            NoUnderlineRule(),
            ListFormatRule(),
            HeadingNumberingRule(),
            HeadingCaseRule(),
            HeadingAlignmentRule(),
            HeadingPunctuationRule(),
            HeadingBoldRule(),
            HeadingNewPageRule(),
            SubheadingCaseRule(),
            RequiredSectionsRule(),
            TableReferenceRule(),
            FigureReferenceRule(),
            TableCaptionRule(),
            FigureCaptionRule(),
            FormulaReferenceRule(),
            CitationFormatRule(),
            AppendixRule(),
            PageNumberingRule(),
        )
        return pipeline

    async def process(self, file: BinaryIO, filename: str, user_id: str) -> tuple[bytes, Document]:
        original_content = file.read()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(original_content)
            tmp_path = tmp.name

        output_path = tmp_path.replace(".docx", "_fixed.docx")

        try:
            ctx = DocumentContext(tmp_path)
            pipeline = self._create_pipeline()
            pipeline.validate_and_fix(ctx)
            ctx.save(Path(output_path))

            with open(output_path, "rb") as f:
                fixed_content = f.read()

            original_key = s3_service.upload_original(user_id, filename, original_content)
            fixed_key = s3_service.upload_fixed(user_id, filename, fixed_content)

            document = Document(
                user_id=user_id,
                filename=filename,
                original_s3_key=original_key,
                fixed_s3_key=fixed_key,
            )
            self.session.add(document)
            await self.session.commit()
            await self.session.refresh(document)

            return fixed_content, document
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
            if os.path.exists(output_path):
                os.unlink(output_path)

    async def get_user_documents(self, user_id: str) -> list[Document]:
        from sqlalchemy import select
        result = await self.session.execute(
            select(Document).where(Document.user_id == user_id).order_by(Document.created_at.desc())
        )
        return list(result.scalars().all())

    async def get_document(self, document_id: str, user_id: str) -> Document | None:
        from sqlalchemy import select
        result = await self.session.execute(
            select(Document).where(Document.id == document_id, Document.user_id == user_id)
        )
        return result.scalar_one_or_none()

    def download_original(self, document: Document) -> bytes:
        return s3_service.download_file(document.original_s3_key)

    def download_fixed(self, document: Document) -> bytes:
        return s3_service.download_file(document.fixed_s3_key)
