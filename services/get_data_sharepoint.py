from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.listitems.caml.query import CamlQuery
from decouple import config
import json


class GetDataSharepoint:
    """
    Classe para busca de dados em listas do SharePoint usando App Registration (Client ID + Secret).
    """

    def __init__(self):
        self.client_id = config("CLIENT_ID")
        self.client_secret = config("CLIENT_SECRET")
        self.site_url = config("SITE_URL")
        self.list_name = config("LIST_NAME")
        self.ctx = None  # Conexão será criada sob demanda (lazy)

    def _connect(self):
        """
        Cria e cacheia o contexto de conexão se ainda não existir.
        """
        if not self.ctx:
            credentials = ClientCredential(self.client_id, self.client_secret)
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
        return self.ctx

    def get_item_by_field(self, field_name: str, value: str) -> list[dict]:
        """
        Busca por um campo customizado na lista (ex: ID_EMAIL, CPF, CNPJ, etc).

        :param field_name: Nome interno do campo SharePoint (não o display name).
        :param value: Valor a ser buscado.
        :return: lista de dicts com os resultados encontrados.
        """
        ctx = self._connect()
        lista = ctx.web.lists.get_by_title(self.list_name)

        query = CamlQuery.parse(f"""
        <View>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name='{field_name}' />
                        <Value Type='Text'>{value}</Value>
                    </Eq>
                </Where>
            </Query>
        </View>
        """)

        try:
            items = lista.get_items(query).execute_query()
            items_list = [i.properties for i in items]
            dados_str = items_list[0].get("dados")  # vem como string
            json_data = json.loads(dados_str)
            data_fields = json_data["data_fields"]
            return data_fields
        except ClientRequestException as e:
            print(f"⚠️ Erro ao buscar campo {field_name}={value}: {e}")
            return []



