from fastapi_users.db import SQLAlchemyBaseUserTableUUID

from src.domain.base import BaseModel


class User(SQLAlchemyBaseUserTableUUID, BaseModel):
    __tablename__ = "user"
