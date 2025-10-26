from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.files.file import File
import os
from decouple import config


class SendDataSharepoint:
    """
    Classe para operações com SharePoint usando App Registration (Client ID + Secret).
    """

    def __init__(self):
        self.client_id = config("CLIENT_ID")
        self.client_secret = config("CLIENT_SECRET")
        self.site_url = config("SITE_URL")
        self.ctx = None  # Lazy init

    def _connect(self):
        """Cria uma conexão autenticada caso ainda não exista."""
        if not self.ctx:
            credentials = ClientCredential(self.client_id, self.client_secret)
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
        return self.ctx

    def create_folder_and_upload(self, library_name: str, folder_name: str, local_file_path: str, file_name: str):
        """
        Cria subpastas dentro de uma Document Library e envia um arquivo para ela.

        :param library_name: Nome da biblioteca (ex: "Shared Documents")
        :param folder_name: Caminho das pastas (ex: "Output/Contratos/ID")
        :param local_file_path: Caminho completo do arquivo local
        :param file_name: Nome do arquivo ao enviar
        """
        ctx = self._connect()
        root_folder = ctx.web.get_folder_by_server_relative_path(library_name).get().execute_query()

        # Quebra caminho em pastas: "Output/ID" → ["Output", "ID"]
        folder_parts = folder_name.split("/")
        current_folder = root_folder

        for part in folder_parts:
            try:
                # Tenta acessar a pasta
                current_folder = current_folder.folders.get_by_url(part).get().execute_query()
            except ClientRequestException:
                # Se não existir, cria
                current_folder = current_folder.folders.add(part).execute_query()

        # Upload
        with open(local_file_path, "rb") as file_stream:
            uploaded_file = current_folder.upload_file(file_name, file_stream.read()).execute_query()

        print(f"✅ Arquivo enviado em: {uploaded_file.serverRelativeUrl}")
        return uploaded_file.serverRelativeUrl

