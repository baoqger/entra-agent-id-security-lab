from pydantic_settings import BaseSettings, SettingsConfigDict
from typing import Literal


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # Mode toggle
    provisioning_mode: Literal["mock", "live"] = "mock"

    # Local dev only (see auth.py, seed.py, main.py)
    dev_auth_bypass: bool = False
    seed_demo_data: bool = False
    enable_docs: bool = False

    # Entra ID
    azure_tenant_id: str = ""
    azure_client_id: str = ""
    frontend_client_id: str = ""

    # --- Credential selection (pick ONE, in order of recommendation) ---
    # 1) Managed Identity (recommended when deployed to Azure).
    use_managed_identity: bool = False
    # Optional: user-assigned MI client ID. Leave blank for system-assigned.
    managed_identity_client_id: str = ""

    # 2) Certificate auth (recommended for non-Azure hosts).
    # Path to the PEM-encoded private key. The corresponding public cert must
    # be uploaded to the app registration.
    azure_client_certificate_path: str = ""
    azure_client_certificate_thumbprint: str = ""

    # 3) Client secret (least preferred - rotate frequently, store in Key Vault).
    azure_client_secret: str = ""
    # Optional: Key Vault reference, e.g. "https://<vault>.vault.azure.net/secrets/app-secret"
    # When set and use_managed_identity=true, the secret is fetched at startup.
    azure_client_secret_keyvault_uri: str = ""

    # Graph
    graph_scope: str = "https://graph.microsoft.com/.default"

    # Email
    smtp_host: str = "smtp.office365.com"
    smtp_port: int = 587
    smtp_username: str = ""
    smtp_password: str = ""
    approval_link_base_url: str = "http://localhost:3000"

    # Teams
    teams_webhook_url: str = ""

    # Database
    database_url: str = "sqlite:///./agent_portal.db"

    # CORS (comma-separated list supported)
    frontend_origin: str = "http://localhost:3000"


settings = Settings()
