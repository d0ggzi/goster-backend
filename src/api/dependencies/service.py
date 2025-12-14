from fastapi import Depends
from sqlalchemy.orm import Session

from src.domain.base import get_session
from src.service.old.old_processing_service import OldProcessingService


async def get_processing_service(session: Session = Depends(get_session)):
    return OldProcessingService(session)
