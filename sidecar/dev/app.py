"""
LLM Agent with Weather Tool - Uses LangChain + Agent Identity for secure API calls
This agent demonstrates how an AI agent uses tools with Agent Identity tokens.
"""

import os
import json
import base64
import requests
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS

# LangChain imports - using try/except for graceful fallback
LANGCHAIN_AVAILABLE = False
ChatOllama = None
tool = None

try:
    from langchain_ollama import ChatOllama
    from langchain_core.tools import tool
    from langchain.agents import create_agent
    LANGCHAIN_AVAILABLE = True
    print("LangChain loaded successfully (LangGraph ReAct)")
except ImportError as e:
    print(f"LangChain not available: {e}")
    print("Running in direct mode only")

app = Flask(__name__)
CORS(app)  # ⚠️  DEMO ONLY — allows all origins. Restrict in production.

# Configuration
SIDECAR_URL = os.environ.get('SIDECAR_URL', 'http://sidecar:5000')
WEATHER_API_URL = os.environ.get('WEATHER_API_URL', 'http://weather-api:8080')
AGENT_APP_ID = os.environ.get('AGENT_APP_ID', '')
BLUEPRINT_APP_ID = os.environ.get('BLUEPRINT_APP_ID', '')
TENANT_ID = os.environ.get('TENANT_ID', '')
CLIENT_SPA_APP_ID = os.environ.get('CLIENT_SPA_APP_ID', '')
OLLAMA_URL = os.environ.get('OLLAMA_URL', 'http://ollama:11434')
OLLAMA_MODEL = os.environ.get('OLLAMA_MODEL', 'llama3.2')

# ⚠️  DEMO ONLY — global state. Not thread-safe; suitable for single-user demos only.
debug_logs = []


def log_debug(step, message, data=None):
    """Log debug information for UI display"""
    entry = {
        "step": step,
        "message": message,
        "data": data
    }
    debug_logs.append(entry)
    print(f"[{step}] {message}")
    if data:
        print(f"    Data: {json.dumps(data, indent=2)[:500]}")


def clear_debug():
    """Clear debug logs for new request"""
    global debug_logs
    debug_logs = []


def decode_jwt_payload(token):
    """Decode JWT payload (without verification) to display claims"""
    try:
        if token.startswith('Bearer '):
            token = token[7:]
        parts = token.split('.')
        if len(parts) != 3:
            return None
        payload = parts[1]
        padding = 4 - len(payload) % 4
        if padding != 4:
            payload += '=' * padding
        decoded = base64.urlsafe_b64decode(payload)
        return json.loads(decoded)
    except Exception:
        return None


def get_agent_token():
    """Get Agent Identity token from sidecar (autonomous / app-only)"""
    log_debug("2.A TOKEN REQUEST", f"Requesting token for Agent: {AGENT_APP_ID}")
    
    try:
        url = f"{SIDECAR_URL}/AuthorizationHeaderUnauthenticated/weather-app?AgentIdentity={AGENT_APP_ID}"
        log_debug("2.B SIDECAR CALL", f"Sidecar URL: {url}")
        
        response = requests.get(url, timeout=30, headers={"Host": "localhost"})
        response.raise_for_status()
        
        result = response.json()
        auth_header = result.get('authorizationHeader', '')
        
        if auth_header:
            claims = decode_jwt_payload(auth_header)
            if claims:
                global _last_tr_claims
                _last_tr_claims = claims
                log_debug("2.C TOKEN RECEIVED", "Got Agent Identity token (TR) from sidecar", {
                    "_jwt_token": {
                        "type": "tr",
                        "title": "\U0001f512 TR \u2014 Autonomous Agent Token (App-Only / Client Credentials)",
                        "css": "tr",
                        "hl": "highlight-purple",
                        "claims": claims
                    }
                })
        
        return auth_header
    except Exception as e:
        log_debug("2. TOKEN ERROR", f"Failed to get token: {str(e)}")
        return None


def get_agent_token_obo(user_token=None):
    """Get Agent Identity token via OBO (On-Behalf-Of) flow."""
    log_debug("OBO 2.A TOKEN REQUEST", f"Requesting OBO token for Agent: {AGENT_APP_ID}", {
        "endpoint": "/AuthorizationHeader/graph (authenticated)",
        "flow": "User Token (Tc) \u2192 Sidecar \u2192 T1 (Blueprint) \u2192 OBO Exchange \u2192 TR (Interactive Agent)",
    })
    
    try:
        url = f"{SIDECAR_URL}/AuthorizationHeader/weather?AgentIdentity={AGENT_APP_ID}"
        tc_snippet = ""
        if user_token:
            raw = user_token.replace("Bearer ", "") if user_token.startswith("Bearer ") else user_token
            tc_snippet = raw[:32] + "..." + raw[-16:] if len(raw) > 52 else raw
        log_debug("OBO 2.B ENDPOINT", f"Authenticated sidecar URL: {url}", {
            "authorization_header": f"Bearer {tc_snippet}" if tc_snippet else "(none)",
            "note": "Unlike /AuthorizationHeaderUnauthenticated, this endpoint REQUIRES a Bearer token (Tc)",
        })
        
        headers = {"Host": "localhost"}
        if user_token:
            if not user_token.startswith("Bearer "):
                headers["Authorization"] = f"Bearer {user_token}"
            else:
                headers["Authorization"] = user_token
        
        response = requests.get(url, timeout=30, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            auth_header = result.get('authorizationHeader', '')
            
            if auth_header:
                claims = decode_jwt_payload(auth_header)
                if claims:
                    global _last_tr_claims
                    _last_tr_claims = claims
                    log_debug("OBO 2.D TOKEN RECEIVED", "Got interactive agent token (TR) via OBO exchange", {
                        "_jwt_token": {
                            "type": "tr",
                            "title": "\U0001f4aa TR \u2014 Agent OBO Token (Interactive)",
                            "css": "tr",
                            "hl": "highlight-green",
                            "claims": claims
                        }
                    })
            
            return auth_header
        else:
            error_text = response.text[:500]
            log_debug("OBO 2.D ENDPOINT RESPONSE", f"Sidecar returned HTTP {response.status_code}", {
                "status_code": response.status_code,
                "response": error_text,
            })
            return None
    except Exception as e:
        log_debug("OBO 2. ERROR", f"Failed to get OBO token: {str(e)}")
        return None


def get_t1_token_claims():
    """Get the app-only (T1) token claims for display purposes."""
    try:
        url = f"{SIDECAR_URL}/AuthorizationHeaderUnauthenticated/weather-app?AgentIdentity={AGENT_APP_ID}"
        response = requests.get(url, timeout=30, headers={"Host": "localhost"})
        response.raise_for_status()
        result = response.json()
        auth_header = result.get('authorizationHeader', '')
        if auth_header:
            return decode_jwt_payload(auth_header)
    except Exception:
        pass
    return None


def call_weather_api(city: str, token: str, token_label: str = "TR", is_obo: bool = False):
    """Call Weather API with Agent Identity token"""
    log_debug("3.A API CALL", f"Calling Weather API for: {city}")
    
    try:
        url = f"{WEATHER_API_URL}/weather?city={city}"
        headers = {"Authorization": token}
        
        raw = token.replace("Bearer ", "") if token.startswith("Bearer ") else token
        snippet = raw[:32] + "..." + raw[-16:] if len(raw) > 52 else raw
        if is_obo:
            token_desc = "TR \u2014 Interactive Agent Token (acts on behalf of user via OBO)"
        else:
            token_desc = "TR \u2014 Autonomous Agent Token (app-only, no user context)"
        
        log_debug("3.B API URL", f"URL: {url}", {
            "token_sent": token_label,
            "token_description": token_desc,
            "authorization_header": f"Authorization: Bearer {snippet}",
        })
        
        # Show what the Weather API validates on the token
        tr_claims = decode_jwt_payload(raw) or {}
        iss = tr_claims.get('iss', '?')
        exp = tr_claims.get('exp', '?')
        aud = tr_claims.get('aud', '?')
        appid = tr_claims.get('appid') or tr_claims.get('azp') or '?'
        xms_frd = tr_claims.get('xms_frd', '')
        xms_par = tr_claims.get('xms_par_app_azp', '')
        is_agent = xms_frd == "FederatedAgent" or bool(xms_par)
        flow_type = "Interactive Agent (OBO)" if is_obo else "Autonomous Agent"
        
        checks = {
            "1_signature": "RS256 via JWKS (login.microsoftonline.com/.well-known/openid-configuration)",
            "2_issuer": f"{'PASS' if 'sts.windows.net' in iss or 'login.microsoftonline.com' in iss else 'CHECK'} \u2014 iss: {iss}",
            "3_expiry": f"PASS \u2014 exp: {exp}",
            "4_agent_identity": f"{'PASS' if is_agent else 'NONE'} \u2014 xms_frd={xms_frd or '(absent)'}, xms_par_app_azp={xms_par or '(absent)'}",
            "5_flow_type": flow_type,
            "6_app_id": appid,
            "7_audience": aud,
        }
        log_debug("3.C TOKEN VALIDATION", f"Weather API validates TR token ({flow_type})", checks)
        
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        weather_data = response.json()
        log_debug("3.D API RESPONSE", "Got weather data from API", weather_data)
        
        return weather_data
    except Exception as e:
        log_debug("3. WEATHER ERROR", f"API call failed: {str(e)}")
        return None


# ============================================
# Weather Function (works with or without LangChain)
# ============================================
# Global holder for user_token when OBO mode is active (set per-request)
_current_user_token = None
# Store last TR (result token) claims for display
_last_tr_claims = None


def get_weather_data(city: str, user_token=None) -> str:
    """
    Get the current weather for a city.
    Uses Agent Identity to securely authenticate with the Weather API.
    If user_token is provided, uses OBO flow via authenticated sidecar endpoint.
    """
    is_obo = user_token is not None
    flow_label = "OBO" if is_obo else "Autonomous"
    log_debug("1.B TOOL CALL", f"Weather function called for city: {city} (flow: {flow_label})")
    
    # Step 1: Get Agent Identity token from sidecar
    if is_obo:
        token = get_agent_token_obo(user_token=user_token)
    else:
        token = get_agent_token()
    
    if not token:
        return f"Error: Could not authenticate with Agent Identity ({flow_label}). The sidecar may not be running."
    
    # Step 2: Call Weather API with the token
    token_label = "TR"
    weather = call_weather_api(city, token, token_label=token_label, is_obo=is_obo)
    if not weather:
        return f"Error: Could not get weather data for {city}. The API may have rejected the token."
    
    # Step 3: Format response
    result = f"""Weather for {weather.get('city', city)}:
- Temperature: {weather.get('temperature', 'N/A')}\u00b0{weather.get('temperature_unit', 'F')}
- Condition: {weather.get('condition', 'N/A')}
- Humidity: {weather.get('humidity', 'N/A')}%
- Wind Speed: {weather.get('wind_speed', 'N/A')} {weather.get('wind_unit', 'mph')}
- Timestamp: {weather.get('timestamp', 'N/A')} ({weather.get('timezone', 'UTC')})
- Data Source: {weather.get('data_source', 'Weather API')}
- Authentication: Validated by {weather.get('validated_by', 'Agent Identity Token')}
- Agent App ID: {weather.get('agent_app_id', 'N/A')}
- Token Flow: {flow_label}"""
    
    log_debug("4. TOOL RESULT", f"Weather data retrieved ({flow_label})", {"result": result})
    return result


# Create LangChain tool wrapper if available
if LANGCHAIN_AVAILABLE and tool is not None:
    @tool
    def get_weather(city: str) -> str:
        """
        Get the current weather for a city. Use this tool when the user asks about weather.
        This tool uses Agent Identity to securely authenticate with the Weather API.
        
        Args:
            city: The name of the city to get weather for (e.g., "Seattle", "New York", "London")
        
        Returns:
            Weather information including temperature, condition, and humidity.
        """
        return get_weather_data(city, user_token=_current_user_token)


# ============================================
# LangChain Agent Setup
# ============================================
def create_weather_agent():
    """Create LangChain agent with weather tool using LangGraph ReAct pattern"""
    
    # Initialize Ollama LLM with extended timeout for first request
    llm = ChatOllama(
        model=OLLAMA_MODEL,
        base_url=OLLAMA_URL,
        temperature=0.7,
        timeout=120,  # 2 min timeout for first request (model loading)
    )
    
    # Define tools
    tools = [get_weather]
    
    # Use LangGraph ReAct agent (same pattern as aws/gcp sidecars)
    agent = create_agent(llm, tools)
    return agent


def process_with_langchain(user_query: str):
    """Process query using LangChain agent with tools"""
    log_debug("0.A START", f"User query: {user_query}")
    log_debug("0.B LANGCHAIN", "Sending query to LangChain agent (LangGraph ReAct)")
    
    try:
        agent = create_weather_agent()
        log_debug("0.C AGENT READY", f"LangChain agent created with Ollama ({OLLAMA_MODEL})")
        
        # LangGraph ReAct agent — keep prompt minimal; small models (qwen2.5:1.5b)
        # are sensitive to extra instructions and may skip the tool call.
        result = agent.invoke(
            {"messages": [("human", user_query)]},
            {"recursion_limit": 10}  # Max 10 steps to prevent loops
        )
        
        # Extract final message
        output = result.get("messages", [])[-1].content if result.get("messages") else "No response"
        
        log_debug("5. COMPLETE", "LangChain agent finished processing")
        
        return {
            "response": output,
            "debug": debug_logs,
            "success": True,
            "agent_type": "langchain"
        }
    except Exception as e:
        log_debug("ERROR", f"LangChain agent failed: {str(e)}")
        return {
            "response": f"Agent error: {str(e)}",
            "debug": debug_logs,
            "success": False,
            "agent_type": "langchain"
        }


def process_without_llm(user_query: str, user_token=None):
    """Fallback: Process query without LLM (direct tool call).
    If user_token is provided, uses OBO flow."""
    is_obo = user_token is not None
    flow_label = "OBO" if is_obo else "Autonomous"
    log_debug("0.A START", f"Processing query (Direct + {flow_label}): {user_query}")
    
    if is_obo:
        log_debug("0.B OBO MODE", "User token provided \u2014 will use authenticated sidecar endpoint", {
            "endpoint": "/AuthorizationHeader/graph (requires Bearer token)",
        })
    
    # Extract city from query
    import re
    
    city = None
    clean_query = user_query.strip().rstrip('?').rstrip('.')
    
    match = re.search(r'\bin\s+([A-Za-z][A-Za-z\s]*?)$', clean_query, re.IGNORECASE)
    if match:
        city = match.group(1).strip()
    
    if not city:
        match = re.search(r'\bfor\s+([A-Za-z][A-Za-z\s]*?)$', clean_query, re.IGNORECASE)
        if match:
            city = match.group(1).strip()
    
    if not city:
        words = clean_query.split()
        if words:
            last_word = words[-1]
            common_words = {'weather', 'what', 'is', 'the', 'how', 'today', 'now', 'like'}
            if last_word.lower() not in common_words:
                city = last_word
    
    if not city:
        city = "Seattle"
    
    log_debug("1.A DIRECT CALL", f"Calling weather function directly for: {city} (flow: {flow_label})")
    weather_result = get_weather_data(city, user_token=user_token)
    
    flow_badge = "\U0001f504 OBO" if is_obo else "\u26a1 Autonomous"
    response = f"""Here's what I found:

{weather_result}

\u2705 *Securely retrieved using Agent Identity ({flow_badge})*"""
    
    log_debug("5. COMPLETE", f"Query processed (Direct + {flow_label})")
    
    return {
        "response": response,
        "debug": debug_logs,
        "success": True,
        "agent_type": "direct",
        "token_flow": "obo" if is_obo else "autonomous"
    }


def check_ollama_available():
    """Check if Ollama is running and has the model"""
    try:
        response = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
        if response.status_code == 200:
            models = response.json().get("models", [])
            model_names = [m.get("name", "").split(":")[0] for m in models]
            return OLLAMA_MODEL.split(":")[0] in model_names
    except:
        pass
    return False


# ============================================
# Flask Routes
# ============================================
@app.route('/')
def index():
    """Serve the chat UI"""
    return render_template('index.html')


@app.route('/api/chat', methods=['POST'])
def chat():
    """Handle chat messages.
    
    Accepts:
        message: str - user query
        use_langchain: bool - whether to use Ollama LLM
        token_flow: 'autonomous' | 'obo' - token acquisition flow
        user_token: str | null - MSAL user access token (required for OBO)
    """
    global _current_user_token, _last_tr_claims
    data = request.json
    user_message = data.get('message', '')
    use_langchain = data.get('use_langchain', True)
    token_flow = data.get('token_flow', 'autonomous')
    user_token = data.get('user_token', None)
    
    if not user_message:
        return jsonify({"error": "No message provided"}), 400
    
    # For OBO flow, user_token is required
    if token_flow == 'obo' and not user_token:
        return jsonify({"error": "OBO flow requires a user token. Please sign in first."}), 400
    
    # Set global token for LangChain tool access
    _current_user_token = user_token if token_flow == 'obo' else None
    _last_tr_claims = None
    clear_debug()
    
    # Decode Tc (user token) claims for display
    if user_token and token_flow == 'obo':
        tc_claims = decode_jwt_payload(user_token)
        if tc_claims:
            log_debug("OBO 0.A USER TOKEN (Tc)", "Decoded user access token from MSAL sign-in", {
                "_jwt_token": {
                    "type": "tc",
                    "title": "\U0001f511 Tc \u2014 User Token (from MSAL sign-in)",
                    "css": "tc",
                    "hl": "highlight",
                    "claims": tc_claims
                }
            })
    
    try:
        if use_langchain and LANGCHAIN_AVAILABLE and check_ollama_available():
            result = process_with_langchain(user_message)
            result['token_flow'] = token_flow
        else:
            result = process_without_llm(user_message, user_token=_current_user_token)
    finally:
        _current_user_token = None
    
    # For OBO, fetch T1 and insert before TR in debug log
    if token_flow == 'obo':
        t1_claims = get_t1_token_claims()
        if t1_claims:
            t1_entry = {
                "step": "OBO 2.C T1 (Blueprint)",
                "message": "Blueprint app-only token used as client_assertion in OBO exchange",
                "data": {
                    "_jwt_token": {
                        "type": "t1",
                        "title": "\U0001f4dc T1 \u2014 Blueprint Token (App-Only / Client Credentials)",
                        "css": "t1",
                        "hl": "highlight-purple",
                        "claims": t1_claims
                    }
                }
            }
            tr_idx = next((i for i, e in enumerate(result['debug']) if 'OBO 2.D' in e.get('step', '')), None)
            if tr_idx is not None:
                result['debug'].insert(tr_idx, t1_entry)
            else:
                result['debug'].append(t1_entry)
    _last_tr_claims = None
    
    # Add doc links at end
    result['debug'].append({
        "step": "DOCS",
        "message": "_doc_links",
        "data": None
    })
    
    return jsonify(result)


@app.route('/api/status', methods=['GET'])
def status():
    """Check service status"""
    ollama_ready = check_ollama_available()
    return jsonify({
        "ollama_available": ollama_ready,
        "ollama_url": OLLAMA_URL,
        "ollama_model": OLLAMA_MODEL,
        "sidecar_url": SIDECAR_URL,
        "agent_app_id": AGENT_APP_ID[:8] + "..." if AGENT_APP_ID else "not set"
    })


@app.route('/api/config', methods=['GET'])
def config():
    """Return MSAL configuration for browser-side OBO sign-in."""
    scheme = request.headers.get('X-Forwarded-Proto', request.scheme)
    host = request.headers.get('X-Forwarded-Host', request.host)
    redirect_uri = f"{scheme}://{host}"
    return jsonify({
        "tenant_id": TENANT_ID,
        "blueprint_app_id": BLUEPRINT_APP_ID,
        "client_spa_app_id": CLIENT_SPA_APP_ID,
        "agent_app_id": AGENT_APP_ID,
        "obo_scopes": [f"api://{BLUEPRINT_APP_ID}/access_as_user"] if BLUEPRINT_APP_ID else [],
        "authority": f"https://login.microsoftonline.com/{TENANT_ID}" if TENANT_ID else "",
        "redirect_uri": redirect_uri,
    })


@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({
        "status": "healthy",
        "service": "LLM Weather Agent (LangChain)",
        "agent_app_id": AGENT_APP_ID[:8] + "..." if AGENT_APP_ID else "not set"
    })


# ============================================


if __name__ == '__main__':
    print("=" * 60)
    print("  3P Agent Identity Demo")
    print("=" * 60)
    print(f"  Sidecar URL: {SIDECAR_URL}")
    print(f"  Weather API: {WEATHER_API_URL}")
    print(f"  Agent App ID: {AGENT_APP_ID[:8]}..." if AGENT_APP_ID else "  Agent App ID: NOT SET")
    print(f"  Ollama URL: {OLLAMA_URL}")
    print(f"  Ollama Model: {OLLAMA_MODEL}")
    print("=" * 60)
    print("  Open http://localhost:3000 in your browser")
    print("=" * 60)
    
    app.run(host='0.0.0.0', port=3000, debug=True)
