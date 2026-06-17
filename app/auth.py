"""Authentication, sessions, and bcrypt-backed user helpers."""

from .legacy import (
    SESSIONS,
    authenticate_user,
    create_user,
    current_user,
    get_user_by_email,
    init_user_db,
    require_role,
    require_user,
)

__all__ = [
    "SESSIONS",
    "authenticate_user",
    "create_user",
    "current_user",
    "get_user_by_email",
    "init_user_db",
    "require_role",
    "require_user",
]

