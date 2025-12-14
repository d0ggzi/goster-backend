from datetime import datetime
from typing import BinaryIO

import boto3
from botocore.client import BaseClient
from botocore.exceptions import ClientError

from src.settings.config import settings


class S3Service:
    def __init__(self):
        self.s3: BaseClient = boto3.client(
            's3',
            endpoint_url=settings.YANDEX_S3_ENDPOINT,
            aws_access_key_id=settings.YANDEX_S3_ACCESS_KEY,
            aws_secret_access_key=settings.YANDEX_S3_SECRET_KEY,
        )

    def upload_file(self, file: BinaryIO):
        ...

    def get_presigned_url(self, filename: str, content_type: str) -> tuple[str, str]:
        object_name = f"files/{datetime.utcnow().timestamp()}_{filename}"
        presigned_url = self.s3.generate_presigned_url(
            'get_object',
            Params={
                'Bucket': settings.YANDEX_S3_BUCKET_NAME,
                'Key': object_name,
                'ContentType': content_type
            },
            ExpiresIn=300
        )
        public_url = f"{settings.YANDEX_S3_ENDPOINT}/{settings.YANDEX_S3_BUCKET_NAME}/{object_name}"
        return presigned_url, public_url

    def delete_file(self, file: str) -> None:
        key = file.split(settings.YANDEX_S3_BUCKET_NAME)[1][1:]
        try:
            self.s3.delete_object(Bucket=settings.YANDEX_S3_BUCKET_NAME, Key=key)
        except ClientError:
            raise

s3_service = S3Service()
