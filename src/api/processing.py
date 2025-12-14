from fastapi import APIRouter, UploadFile, Depends

from src.api.dependencies.service import get_processing_service
from src.service.old.old_processing_service import OldProcessingService

processing_router = APIRouter(prefix="/api/processing", tags=["processing"])


@processing_router.post("/")
async def list_services(file: UploadFile, processing_service: OldProcessingService = Depends(get_processing_service)):
    return await processing_service.execute(file.file)
