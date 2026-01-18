from src.service.s3 import S3Service, s3_service


async def get_s3_service() -> S3Service:
    return s3_service
