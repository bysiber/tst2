"""
This function should serve as the interface to retrieve credentials from Azure KeyVault.
It is separated from the other packages to:
    - allow independent usage.
    - prevent circle-referencing while defining the project config.
"""

from azure.identity import ClientSecretCredential
from azure.keyvault.secrets import SecretClient, KeyVaultSecret

# It's a wrapper around the Azure Key Vault SDK for Python
class KeyVaultClient:
    def __init__(self):
        """
        It creates a client object that can be used to access the Azure Key Vault
        """
        self.__tenant_id=""
        self.__client_id="" # App: PythonKvReader
        self.__client_secret="" # App: PythonKvReader
        self.__vault_url=""
        self._credential = ClientSecretCredential(
            self.__tenant_id, 
            self.__client_id, 
            self.__client_secret
        )
        self._client = SecretClient(vault_url=self.__vault_url, credential=self._credential)

    def get_secret(self, secret_name):
        """
        It takes a secret name as input and returns the secret value
        
        :param secret_name: The name of the secret to retrieve
        :return: The secret value
        """
        secret = self._client.get_secret(secret_name)
        return str(secret.value)

    def set_secret(self, secret_name, secret_value):
        """
        It sets the secret value in the key vault.
        
        :param secret_name: The name of the secret to be set
        :param secret_value: The value of the secret
        """
        self._client.set_secret(secret_name, secret_value)

    def delete_secret(self, secret_name):
        """
        This function deletes a secret from the Azure Key Vault
        
        :param secret_name: The name of the secret to be deleted
        """
        self._client.delete_secret(secret_name)

    def update_secret(self, secret_name, secret_value):
        """
        It updates the secret value of the secret with the name `secret_name` to the value `secret_value`
        
        :param secret_name: The name of the secret to update
        :param secret_value: The value of the secret
        """
        self._client.update_secret(secret_name, secret_value)

    def list_secrets(self):
        """
        It returns a list of all the secrets in the vault
        :return: A list of secrets
        """
        secrets = self._client.list_properties_of_secrets()
        return [secret.name for secret in secrets]