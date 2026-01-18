import uuid
from typing import BinaryIO

import boto3
from botocore.client import BaseClient

from src.settings.config import settings


class S3Service:
    def __init__(self):
        self.s3: BaseClient = boto3.client(
            "s3",
            endpoint_url=settings.S3_ENDPOINT or None,
            aws_access_key_id=settings.S3_ACCESS_KEY or None,
            aws_secret_access_key=settings.S3_SECRET_KEY or None,
        )
        self.bucket = settings.S3_BUCKET_NAME

    def upload_file(self, file_content: bytes, key: str) -> str:
        self.s3.put_object(
            Bucket=self.bucket,
            Key=key,
            Body=file_content,
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        return key

    def upload_original(self, user_id: str, filename: str, content: bytes) -> str:
        key = f"{user_id}/original/{uuid.uuid4()}_{filename}"
        return self.upload_file(content, key)

    def upload_fixed(self, user_id: str, filename: str, content: bytes) -> str:
        key = f"{user_id}/fixed/{uuid.uuid4()}_{filename}"
        return self.upload_file(content, key)

    def download_file(self, key: str) -> bytes:
        response = self.s3.get_object(Bucket=self.bucket, Key=key)
        return response["Body"].read()

    def get_presigned_url(self, key: str, expires_in: int = 3600) -> str:
        return self.s3.generate_presigned_url(
            "get_object",
            Params={"Bucket": self.bucket, "Key": key},
            ExpiresIn=expires_in,
        )

    def delete_file(self, key: str) -> None:
        self.s3.delete_object(Bucket=self.bucket, Key=key)


s3_service = S3Service()
