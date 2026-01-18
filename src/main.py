from contextlib import asynccontextmanager

import fastapi
from fastapi.middleware.cors import CORSMiddleware

from src.api.processing import processing_router
from src.auth import auth_backend, fastapi_users
from src.auth.schemas import UserCreate, UserRead
from src.domain.base import create_db_and_tables
from src.utils.middleware import catch_exceptions_middleware


@asynccontextmanager
async def lifespan(app: fastapi.FastAPI):
    await create_db_and_tables()
    yield


app = fastapi.FastAPI(
    title="GOST Document Validator API",
    description="""
API для валидации и исправления документов по ГОСТ 7.32-2017.

## Возможности

* **Регистрация и авторизация** — JWT токены
* **Обработка документов** — загрузка .docx, получение исправленного файла
* **История документов** — все обработанные документы сохраняются
* **Скачивание** — оригинал и исправленный документ доступны для скачивания

## Использование

1. Зарегистрируйтесь через `POST /auth/register`
2. Получите токен через `POST /auth/jwt/login`
3. Используйте токен в заголовке `Authorization: Bearer <token>`
4. Отправьте документ на `POST /api/documents/process`
5. Получите историю через `GET /api/documents/`
    """,
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.middleware("http")(catch_exceptions_middleware)

app.include_router(
    fastapi_users.get_auth_router(auth_backend),
    prefix="/auth/jwt",
    tags=["Auth"],
)
app.include_router(
    fastapi_users.get_register_router(UserRead, UserCreate),
    prefix="/auth",
    tags=["Auth"],
)
app.include_router(
    fastapi_users.get_users_router(UserRead, UserRead),
    prefix="/users",
    tags=["Users"],
)
app.include_router(processing_router)
