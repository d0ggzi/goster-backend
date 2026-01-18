from typing import BinaryIO

from sqlalchemy.orm import Session
from src.domain.models import File
from src.domain.choices import FileStatus


class OldProcessingService:
    def __init__(self, session: Session):
        self.session = session

    async def execute(self, file: BinaryIO):
        self.session.add(File(status=FileStatus.NEW, original_s3_link=""))
        self.session.commit()
        print(file.readline(10))
