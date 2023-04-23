import pandas as pd
from flask import jsonify, request
import json, os


class SharepointColumns():
    def get(body):
        print(body, type(body))
        with open('./logs/sharepointlog.json', 'w') as f:
            json.dump(body,f,indent=4)
        return jsonify(body)