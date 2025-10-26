import azure.functions as func
import json
from services.process_field_contrato import ProcessFieldContrato

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

@app.route(route="http_trigger", methods=["POST"])
def build_contract(req: func.HttpRequest) -> func.HttpResponse:
    try:
        data = req.get_json()
    except ValueError:
        return func.HttpResponse("Body inválido. Envie JSON.", status_code=400)

    arquivo_final = ProcessFieldContrato().preencher_contrato(data_body=data)

    return func.HttpResponse(
        json.dumps({"status": "ok", "arquivo": arquivo_final}, ensure_ascii=False),
        mimetype="application/json",
        status_code=200
    )
