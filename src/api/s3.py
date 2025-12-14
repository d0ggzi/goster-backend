from fastapi import APIRouter, Depends

from src.api.dependencies.s3 import get_s3_service
from src.service.s3 import S3Service

s3_router = APIRouter(prefix="/api/s3", tags=["s3"])


@s3_router.get("/presigned-url")
def get_presigned_url(filename: str, content_type: str, s3_service: S3Service = Depends(get_s3_service)):
    presigned_url, public_url = s3_service.get_presigned_url(filename, content_type)

    return {
        "upload_url": presigned_url,
        "public_url": public_url
    }
