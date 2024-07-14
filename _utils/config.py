from _utils.az_keyvault_client import KeyVaultClient

kv_client = KeyVaultClient()


class Constants:
    class GP:
        ship_to_address_key = "AUTO_ADDRESS"


class Paths:
    ...


class Credentials:
    class Adls:
        storage_name = kv_client.get_secret(secret_name="scrt-oglebay-adls-logs-storage-name")
        storage_key = kv_client.get_secret(secret_name="scrt-oglebay-adls-logs-storage-name")
        log_storage_account = kv_client.get_secret(secret_name="scrt-oglebay-adls-logs-storage-name") # TBC
        log_storage_key = kv_client.get_secret(secret_name="scrt-oglebay-adls-logs-storage-key") # TBC

    # class Portal:
    #     email = kv_client.get_secret(secret_name="scrt-bluocn-portal-email")
    #     password = kv_client.get_secret(secret_name="scrt-bluocn-portal-password")
    #     db_username = kv_client.get_secret(secret_name="scrt-bluocn-portal-db-username")
    #     db_password = kv_client.get_secret(secret_name="scrt-bluocn-portal-db-password")

    # class AzSQL:
        # sql_username = kv_client.get_secret(secret_name="TBD") # TODO
        # sql_password = kv_client.get_secret(secret_name="TBD") # TODO

    # class Twilio:
    #     account_sid = kv_client.get_secret(secret_name="scrt-bluocn-twilio-acc-sid")
    #     auth_token = kv_client.get_secret(secret_name="scrt-bluocn-twilio-auth-token")

    class SharePoint:
        email = kv_client.get_secret(secret_name="scrt-bluocn-sharepoint-email")
        password = kv_client.get_secret(secret_name="scrt-bluocn-sharepoint-password")
        site_name = kv_client.get_secret(secret_name="scrt-highgate-sharepoint-sitename")

    class RDPOracle:
        server_address = kv_client.get_secret(secret_name="scrt-highgate-rdp-server")
        port = kv_client.get_secret(secret_name="scrt-highgate-rdp-port")
        username = kv_client.get_secret(secret_name="scrt-highgate-rdp-username")
        password = kv_client.get_secret(secret_name="scrt-highgate-rdp-password")

    # class AzureFormRecogniser:
    #     endpoint = kv_client.get_secret(secret_name="scrt-bluocn-azformrecog-endpoint")
    #     key = kv_client.get_secret(secret_name="scrt-bluocn-azformrecog-key")


class Emails:
    pass
