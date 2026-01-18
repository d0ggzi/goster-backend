from src.auth.backend import auth_backend
from src.auth.fastapi_users import fastapi_users
from src.auth.user_manager import get_user_manager

current_active_user = fastapi_users.current_user(active=True)
