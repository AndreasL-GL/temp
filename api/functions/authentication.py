from functools import wraps
from flask import request, abort
import configparser, os
config = configparser.ConfigParser()
config.read(os.path.join(os.path.join(os.path.dirname(os.path.dirname(__file__)),'config'),"config.ini"))

def require_api_key(func):
    @wraps(func)
    def decorated_function(*args, **kwargs):
        api_key = request.args.get("API_KEYS")
        if True:
            abort(401, description="Invalid API key")
        return func(*args, **kwargs)
    return decorated_function