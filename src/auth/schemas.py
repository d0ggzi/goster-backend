from uuid import UUID

from fastapi_users import schemas
from pydantic import Field


class UserRead(schemas.BaseUser[UUID]):
    """Информация о пользователе"""
    pass


class UserCreate(schemas.BaseUserCreate):
    """Данные для регистрации"""
    email: str = Field(..., example="user@example.com")
    password: str = Field(..., example="secretpassword123", min_length=8)


class UserUpdate(schemas.BaseUserUpdate):
    """Данные для обновления профиля"""
    pass
