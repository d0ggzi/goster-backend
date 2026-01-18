from uuid import UUID

from fastapi_users import FastAPIUsers

from src.auth.backend import auth_backend
from src.auth.user_manager import get_user_manager
from src.domain.models import User

fastapi_users = FastAPIUsers[User, UUID](
    get_user_manager,
    [auth_backend],
)
