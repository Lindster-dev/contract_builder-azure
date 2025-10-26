from contrato_service import ContratoService
from get_data_sharepoint import GetDataSharepoint
from send_data_sharepoint import SendDataSharepoint


class ProcessFieldContrato:
    def __init__(self):
        self.get_data = GetDataSharepoint()
        self.contrato_service = ContratoService()
        self.send_data = SendDataSharepoint()

    def preencher_contrato(self, data_body):
        id_email = data_body["id_email"]
        data = self.get_data.get_item_by_field(field_name="id_email", value=id_email)
        self.contrato_service.processar_contrato(data)
        self.send_data.create_folder_and_upload(
            library_name="Shared Documents",
            folder_name=f"Output/{id_email}",
            file_name="contrato_gerado.docx",
            local_file_path="result/contrato_gerado.docx",
        )
        return data


