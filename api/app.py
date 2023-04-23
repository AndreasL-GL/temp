from flask import Flask, render_template, request, redirect, url_for, flash, Response,\
    make_response, json, send_file, jsonify, abort
from config import Config
from functions.SQL_commands import Sql
from image_api_blueprints import image_resize
from flow_tools_blueprints import flow_starting_page, get_sharepoint_columns_


app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
app.register_blueprint(image_resize)
app.register_blueprint(flow_starting_page)
app.register_blueprint(get_sharepoint_columns_)

@app.route("/", methods=['GET', 'POST'])
def Home():
    Sql.initialize_db()
    return render_template('home.html')

@app.route("/help")
def help():
    
    return jsonify(Config.API_KEYS)

@app.before_request
def limit_remote_addr():
    if any(Config.ACCEPT_CONNECTIONS_FROM):
        if any(Config.ACCEPT_CONNECTIONS_FROM) and request.remote_addr not in Config.ACCEPT_CONNECTIONS_FROM:
            abort(403)  # Forbidden
        
        
@app.route("/api")
def api():
    return render_template("api_description.html")


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=3000)
