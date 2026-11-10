"""MSAL / Entra ID token validation + directory role authorization."""

from __future__ import annotations
import logging
import os
import re
import time
import httpx
import jwt as pyjwt
from jwt import PyJWKClient, InvalidTokenError
from fastapi import Depends, HTTPException
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from config import settings

logger = logging.getLogger("auth")

# In dev/mock mode with no tenant configured we allow an unauthenticated "demo user"
# so local demos work without wiring up Entra ID. This is GATED on tenant_id being
# empty AND DEV_AUTH_BYPASS=true - it will never kick in for a real deployment.
_DEV_AUTH_BYPASS = (
    not settings.azure_tenant_id
    and os.getenv("DEV_AUTH_BYPASS", "false").lower() in ("1", "true", "yes")
)

bearer_scheme = HTTPBearer(auto_error=not _DEV_AUTH_BYPASS)

# JWKS client - handles caching and rotation internally
_jwks_client: PyJWKClient | None = None

# Entra directory role template IDs
GLOBAL_ADMIN_ROLE = "62e90394-69f5-4237-9190-012177145e10"
AGENT_ID_ADMIN_ROLE = "dd485484-4e16-4e10-a23f-3e0f31fed964"

APPROVER_ROLE_IDS = {GLOBAL_ADMIN_ROLE, AGENT_ID_ADMIN_ROLE}

# Role cache with TTL (5 min)
_role_cache: dict[str, tuple[set[str], float]] = {}
_ROLE_CACHE_TTL = 300

# OID format validation
_OID_PATTERN = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE)


async def _get_jwks(force_refresh: bool = False) -> dict:
    """Deprecated shim kept for any callers; prefer PyJWKClient."""
    return {}


def _get_jwks_client() -> PyJWKClient:
    global _jwks_client
    if _jwks_client is None:
        url = (
            f"https://login.microsoftonline.com/{settings.azure_tenant_id}"
            "/discovery/v2.0/keys"
        )
        _jwks_client = PyJWKClient(url, cache_keys=True, lifespan=3600)
    return _jwks_client


async def _find_signing_key(token: str):
    """Kept for backwards compat - returns the signing key via PyJWKClient."""
    return _get_jwks_client().get_signing_key_from_jwt(token).key


async def _get_user_directory_roles(user_oid: str, token: str) -> set[str]:
    """Look up the user's Entra directory role assignments via Graph."""
    now = time.time()
    cached = _role_cache.get(user_oid)
    if cached and (now - cached[1]) < _ROLE_CACHE_TTL:
        return cached[0]

    if not _OID_PATTERN.match(user_oid):
        logger.warning("Invalid OID format: %s", user_oid)
        return set()

    try:
        async with httpx.AsyncClient() as client:
            resp = await client.get(
                f"https://graph.microsoft.com/v1.0/users/{user_oid}/memberOf",
                params={"$select": "id,roleTemplateId", "$top": "100"},
                headers={"Authorization": f"Bearer {token}"},
            )
            if resp.status_code == 200:
                data = resp.json()
                role_ids = {
                    m.get("roleTemplateId")
                    for m in data.get("value", [])
                    if m.get("@odata.type") == "#microsoft.graph.directoryRole"
                    and m.get("roleTemplateId")
                }
                _role_cache[user_oid] = (role_ids, now)
                return role_ids
            else:
                logger.error("Graph memberOf returned %d for user %s", resp.status_code, user_oid)
    except Exception as exc:
        logger.error("Failed to fetch directory roles for %s: %s", user_oid, exc)

    return set()


def _get_graph_token_for_role_check() -> str | None:
    """Acquire an app-only Graph token for directory-role lookups.

    Delegates to ``services.credentials.get_graph_token`` which honors the
    MI -> Cert -> Key Vault secret -> inline secret priority.
    """
    try:
        from services.credentials import get_graph_token
        return get_graph_token()
    except Exception as exc:
        logger.error("Failed to acquire Graph token for role check: %s", exc)
        return None


async def get_current_user(
    cred: HTTPAuthorizationCredentials = Depends(bearer_scheme),
) -> dict:
    """Validate the bearer access token and return claims."""
    # Local-dev bypass (opt-in, only when no tenant configured).
    if _DEV_AUTH_BYPASS and cred is None:
        logger.warning("DEV_AUTH_BYPASS active - returning demo user. DO NOT USE IN PRODUCTION.")
        return {
            "preferred_username": "demo@example.com",
            "oid": "00000000-0000-0000-0000-000000000000",
            "is_approver": True,
            "_directory_roles": set(),
        }

    if cred is None:
        raise HTTPException(status_code=401, detail="Missing bearer token")

    token = cred.credentials
    try:
        signing_key = _get_jwks_client().get_signing_key_from_jwt(token).key

        issuer = f"https://login.microsoftonline.com/{settings.azure_tenant_id}/v2.0"
        valid_audiences = [
            f"api://{settings.azure_client_id}",
            settings.azure_client_id,
        ]

        # PyJWT validates aud, iss, exp, iat, nbf, and signature in one call.
        claims = pyjwt.decode(
            token,
            signing_key,
            algorithms=["RS256"],
            audience=valid_audiences,
            issuer=issuer,
        )

        if "preferred_username" not in claims and "upn" in claims:
            claims["preferred_username"] = claims["upn"]

        # Look up directory roles
        user_oid = claims.get("oid")
        if user_oid:
            graph_token = _get_graph_token_for_role_check()
            if graph_token:
                dir_roles = await _get_user_directory_roles(user_oid, graph_token)
                claims["_directory_roles"] = dir_roles
                claims["is_approver"] = bool(dir_roles & APPROVER_ROLE_IDS)

        return claims

    except InvalidTokenError as exc:
        logger.warning("JWT validation failed: %s", exc)
        raise HTTPException(status_code=401, detail="Token validation failed")


def require_approver(user: dict = Depends(get_current_user)) -> dict:
    """Ensure the caller has Global Admin or Agent ID Administrator directory role."""
    if not user.get("is_approver", False):
        logger.warning("Access denied for %s — missing approver role", user.get("preferred_username", "unknown"))
        raise HTTPException(
            status_code=403,
            detail="This action requires Global Administrator or Agent ID Administrator role.",
        )
    return user
