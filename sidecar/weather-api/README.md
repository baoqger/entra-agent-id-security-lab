# Weather API — Shared Downstream API

A minimal Flask app that plays the role of *"the downstream API your agent calls"* in every sample in this repo. It validates incoming Agent Identity tokens and returns real weather data from [Open-Meteo](https://open-meteo.com) (no key required).

This service is intentionally boring. The interesting part is the token validation — it proves that Entra Agent ID tokens work cryptographically end-to-end.

## What it does

- Exposes `GET /weather?city=<name>` and `GET /healthz`.
- On every request to `/weather`, validates the `Authorization: Bearer <token>` header:
  - **Signature** — RS256 verified against the JWKS at `https://login.microsoftonline.com/<tenant>/discovery/v2.0/keys`
  - **Issuer** — must match `https://sts.windows.net/<tenant>/` or `https://login.microsoftonline.com/<tenant>/v2.0`
  - **Audience** — `EXPECTED_AUDIENCE` (e.g. `api://<weather-api-app-id>` from `setup-weather-api-resource.ps1`)
  - **Agent-identity marker** — the `xms_par_app_azp` claim (present in Agent tokens, absent in plain app-only tokens) is logged so you can see which Blueprint minted the call
- Returns real weather on success, HTTP 401 with a reason on any validation failure.

## How the samples use it

```
┌─────────────┐                  ┌──────────────────┐
│ Agent       │  Bearer <T2>     │   weather-api    │
│ (dev/aws)   ├─────────────────▶│                  │
└─────────────┘                  │  - verify token  │
                                 │  - call          │
                                 │    Open-Meteo    │
                                 └──────────────────┘
```

Both the [`dev/`](../dev/README.md) (Ollama) and [`aws/`](../aws/README.md) (Bedrock) sidecar samples call this exact same container — the cross-cloud story only works because the downstream API is identical.

## Run it standalone

```bash
# From sidecar/weather-api
docker build -t weather-api:local .
docker run --rm -p 8080:8080 \
  -e TENANT_ID=<your-tenant-id> \
  -e EXPECTED_AUDIENCE=api://<weather-api-app-id> \
  weather-api:local
```

Then, with an Agent Identity token from [`../../scripts/README.md`](../../scripts/README.md):

```bash
curl -H "Authorization: Bearer $TOKEN" \
     "http://localhost:8080/weather?city=Dallas"
```

Expected response (real data):

```json
{
  "city": "Dallas",
  "temperature": 61,
  "temperature_unit": "F",
  "condition": "Overcast",
  "humidity": 93,
  "wind_speed": 8,
  "is_agent_identity": true,
  "agent_app_id": "<agent-app-id from xms_par_app_azp>",
  "validated_by": "Agent Identity Token",
  "data_source": "Open-Meteo API (Real-time)"
}
```

## Environment variables

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `TENANT_ID` | yes | — | Entra tenant — used to build JWKS URL and issuer check |
| `EXPECTED_AUDIENCE` | recommended | — (aud not checked if unset) | Must match token `aud` (e.g. `api://<app-id>`). Graph-scoped OBO tokens are rejected until OBO uses the same custom resource. |
| `PORT` | no | `8080` | HTTP port |

## Files

| File | Purpose |
|---|---|
| `app.py` | Flask app — route handlers, token validation, Open-Meteo client |
| `Dockerfile` | `python:3.13-slim` base, runs `gunicorn app:app -b 0.0.0.0:8080` |
| `requirements.txt` | `flask`, `pyjwt[crypto]`, `cryptography`, `requests`, `gunicorn` |

## Why this is in the repo

Token validation libraries differ between ecosystems (Python, Node, .NET). Having a *known-good* validator that the `dev` and `aws` samples both hit makes issues much easier to isolate — if both samples fail against `weather-api`, the problem is the sidecar or Entra config, not the downstream API.

For the conceptual overview of how tokens flow from agent → sidecar → API, see [`../README.md`](../README.md).
