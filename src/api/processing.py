from urllib.parse import quote

from fastapi import APIRouter, UploadFile, Depends, File, HTTPException
from fastapi.responses import Response

from src.api.dependencies.service import get_processing_service
from src.api.schemas.document import DocumentResponse, DocumentListResponse
from src.auth import current_active_user
from src.domain.models import User
from src.service.processing_service import ProcessingService

processing_router = APIRouter(prefix="/api/documents", tags=["Documents"])

CYRILLIC_TO_LATIN = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "kh", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "shch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
    "А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Е": "E", "Ё": "Yo",
    "Ж": "Zh", "З": "Z", "И": "I", "Й": "Y", "К": "K", "Л": "L", "М": "M",
    "Н": "N", "О": "O", "П": "P", "Р": "R", "С": "S", "Т": "T", "У": "U",
    "Ф": "F", "Х": "Kh", "Ц": "Ts", "Ч": "Ch", "Ш": "Sh", "Щ": "Shch",
    "Ъ": "", "Ы": "Y", "Ь": "", "Э": "E", "Ю": "Yu", "Я": "Ya",
}


def transliterate(text: str) -> str:
    return "".join(CYRILLIC_TO_LATIN.get(c, c) for c in text)


def make_content_disposition(filename: str, prefix: str = "") -> str:
    full_name = f"{prefix}{filename}"
    ascii_name = transliterate(full_name)
    utf8_name = quote(full_name)
    return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{utf8_name}"


@processing_router.post(
    "/process",
    summary="Обработать документ",
    description="""
Загрузите .docx файл для проверки и исправления по ГОСТ 7.32-2017.

Документ сохраняется в историю и доступен для скачивания через `/api/documents/{id}/fixed`.
    """,
    response_model=DocumentResponse,
    responses={
        401: {"description": "Не авторизован"},
    },
)
async def process_document(
    file: UploadFile = File(..., description="Документ .docx для обработки"),
    user: User = Depends(current_active_user),
    service: ProcessingService = Depends(get_processing_service),
):
    filename = file.filename or "document.docx"
    _, document = await service.process(file.file, filename, str(user.id))
    return DocumentResponse.model_validate(document)


@processing_router.get(
    "/",
    summary="Список документов",
    description="Получить список всех обработанных документов пользователя",
    response_model=DocumentListResponse,
)
async def list_documents(
    user: User = Depends(current_active_user),
    service: ProcessingService = Depends(get_processing_service),
):
    documents = await service.get_user_documents(str(user.id))
    return DocumentListResponse(
        documents=[DocumentResponse.model_validate(doc) for doc in documents],
        total=len(documents),
    )


@processing_router.get(
    "/{document_id}",
    summary="Информация о документе",
    response_model=DocumentResponse,
)
async def get_document(
    document_id: str,
    user: User = Depends(current_active_user),
    service: ProcessingService = Depends(get_processing_service),
):
    document = await service.get_document(document_id, str(user.id))
    if not document:
        raise HTTPException(status_code=404, detail="Документ не найден")
    return DocumentResponse.model_validate(document)


@processing_router.get(
    "/{document_id}/original",
    summary="Скачать оригинал",
    description="Скачать оригинальный загруженный документ",
    responses={
        200: {
            "content": {"application/vnd.openxmlformats-officedocument.wordprocessingml.document": {}},
        },
    },
)
async def download_original(
    document_id: str,
    user: User = Depends(current_active_user),
    service: ProcessingService = Depends(get_processing_service),
):
    document = await service.get_document(document_id, str(user.id))
    if not document:
        raise HTTPException(status_code=404, detail="Документ не найден")

    content = service.download_original(document)
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": make_content_disposition(document.filename),
        },
    )


@processing_router.get(
    "/{document_id}/fixed",
    summary="Скачать исправленный",
    description="Скачать исправленный документ",
    responses={
        200: {
            "content": {"application/vnd.openxmlformats-officedocument.wordprocessingml.document": {}},
        },
    },
)
async def download_fixed(
    document_id: str,
    user: User = Depends(current_active_user),
    service: ProcessingService = Depends(get_processing_service),
):
    document = await service.get_document(document_id, str(user.id))
    if not document:
        raise HTTPException(status_code=404, detail="Документ не найден")

    content = service.download_fixed(document)
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": make_content_disposition(document.filename, "fixed_"),
        },
    )
