from functools import wraps
from flask import request, abort
from config import Config

def require_api_key(func):
    @wraps(func)
    def decorated_function(*args, **kwargs):
        api_key = request.args.get("API_KEY", "")
        if api_key not in Config.API_KEYS:
            abort(401, description="Invalid API key")
        return func(*args, **kwargs)
    return decorated_function