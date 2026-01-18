from fastapi import Depends
from sqlalchemy.ext.asyncio import AsyncSession

from src.domain.base import get_async_session
from src.service.processing_service import ProcessingService


async def get_processing_service(session: AsyncSession = Depends(get_async_session)):
    return ProcessingService(session)
