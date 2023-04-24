
from flask import Blueprint, render_template, request
from functions.get_sharepoint_columns import SharepointColumns
from functions.authentication import require_api_key


flow_starting_page = Blueprint('flow_starting_page', __name__)
get_sharepoint_columns_ = Blueprint('get_sharepoint_columns_', __name__)


@flow_starting_page.route("/api/flows")
def flows_starting_page():
    return render_template("flows_starting_template.html")


@get_sharepoint_columns_.route("/api/flows/get_sharepoint_columns", methods=['POST'])
@require_api_key
def get_sharepoint_columns():
    js = request.json

    result = SharepointColumns.get(js)
    return result