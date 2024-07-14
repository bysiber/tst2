import os, json, requests
import time
from datetime import datetime
from pathlib import Path
from threading import Lock
from urllib import parse

from azure.identity import ClientSecretCredential
from azure.mgmt.datafactory import DataFactoryManagementClient
from azure.storage.blob import BlobSasPermissions, generate_blob_sas
from azure.storage.blob._models import ContentSettings
from azure.storage.filedatalake import DataLakeServiceClient
from azure.storage.fileshare import ShareDirectoryClient

from _utils.config import Credentials

try:
    import pyodbc
except:
    print('pyodbc couldn\'t be imported functions using pyodbc will not work')


class AdlsHandler(object):
    def __init__(self, adls_name: str = Credentials.Adls.storage_name, accnt_key: str = Credentials.Adls.storage_key):
        self.__adls_name = adls_name
        self.__accnt_key = accnt_key
        self.__adls_fileshare_name = 'fs-hrdywrld-bluocn'
        self.conn = self.__getConnection()

    def __getConnection(self):
        self.__service_client = DataLakeServiceClient(
            account_url="{}://{}.dfs.core.windows.net".format("https", self.__adls_name),
            credential=self.__accnt_key
        )
        self.__fs_client = self.__service_client.get_file_system_client(file_system="root")

        return self.__fs_client

    def upload(self, local_path: Path, fname: str, adls_path: Path) -> bool:
        
        directory_client = self.conn.get_directory_client(f'{adls_path}')

        file_client = directory_client.get_file_client(fname)
        local_file = open(local_path / fname, 'rb')
        file_contents = local_file.read()
        file_client.upload_data(file_contents, overwrite=True)

        return True

    def download(self, local_path: Path, fname: str, adls_path: Path) -> Path:

        try:
            directory_client = self.conn.get_directory_client(f'{adls_path}')
            file_client = directory_client.get_file_client(fname)
            download = file_client.download_file()
            downloaded_bytes = download.readall()
            local_file = open(local_path / fname,'wb')
            local_file.write(downloaded_bytes)
            local_file.close()
            return local_path / fname
        except:
            return None
    
    def delete(self, fname: str, adls_path: Path):

        directory_client = self.conn.get_directory_client(f'{adls_path}')
        file_client = directory_client.get_file_client(fname)
        file_client.delete_file()

        return True

    def move(self, source: Path, target: Path):
        
        blob_service_client = self.__service_client._blob_service_client
        source_blob = f"https://{self.__adls_name}.blob.core.windows.net/root/{parse.quote(str(source))}"
        target_blob = blob_service_client.get_blob_client('root', str(target))

        target_blob.start_copy_from_url(source_blob)

        remove_blob = blob_service_client.get_blob_client('root', str(source))
        remove_blob.delete_blob()   

    def list_directory_contents(self, adls_path: Path):

        try:
            paths = self.conn.get_paths(path=adls_path, recursive=False, max_results=1)
            if paths:
                return [i.name for i in paths]
            else:
                return None

        except Exception as e:
            print(e)
            return None
    
    def download_directory_files(self, local_path: Path, adls_path: Path):

        dir_contents = self.list_directory_contents(adls_path)
        # extract filenames from full adls paths
        dir_contents = [Path(path).name for path in dir_contents]
        for filename in dir_contents:
            self.download(local_path, filename, adls_path)

    def send_directory_files(self, local_path: Path, adls_path: Path):

        for filename in os.listdir(local_path):
            self.upload(local_path, filename, adls_path)
    
    def delete_directory_files(self, adls_path: Path):

        dir_contents = self.list_directory_contents(adls_path)
        # extract filenames from full adls paths
        dir_contents = [Path(path).name for path in dir_contents]
        for filename in dir_contents:
            self.delete(filename, adls_path)

    def create_adls_dir(self, adls_path: Path):

        self.conn.create_directory(adls_path)

    def convert_to_local_path(self, adls_path: Path):
        """
        This function converts given adls path to local path
        """

        return Path.cwd() / 'root' / adls_path

    def get_file_metadata(self, fname: str, adls_path: Path) -> dict:

        try:
            directory_client = self.conn.get_directory_client(f'{adls_path}')
            file_client = directory_client.get_file_client(fname)
            file_metadata = file_client.get_file_properties()
            return file_metadata
        except:
            return None

    def create_fileshare_dir(self, adls_path):
        try:
            dir_lst = adls_path.split("/")
            parent_dir = dir_lst[0]
            dir_client = ShareDirectoryClient(account_url=f"https://{self.__adls_name}.file.core.windows.net/", share_name= f"{self.__adls_fileshare_name}", directory_path= f"{parent_dir}", credential=f"{self.__accnt_key}")
            dir_lst.pop(0)
            sub_dir = '/'.join(dir_lst)
            dir_client.create_directory()
            dir_client.create_subdirectory(sub_dir)
        except Exception as e:
            print('Resource Exists: ', str(e))

    def set_http_headers(self, adls_path: Path, fname:str, headers: ContentSettings):

        directory_client = self.conn.get_directory_client(f'{adls_path}')
        file_client = directory_client.get_file_client(fname)

        file_client.set_http_headers(headers)

    def generate_sas_token(self, blob_name: str, expiry_date: datetime, content_type: str):

        sas = generate_blob_sas(account_name=self.__adls_name,
                                account_key=self.__accnt_key,
                                container_name='root',
                                blob_name=blob_name,
                                permission=BlobSasPermissions(read=True),
                                expiry=expiry_date,
                                content_type = content_type
                            )

        sas_url = f'https://{self.__adls_name}.blob.core.windows.net/root/{parse.quote(blob_name)}?{sas}'

        return sas_url

class SqlDBConnect(object):
    
    '''
        Class to connect to SQL DB and run the SQL queries and get JSON output
    '''
    
    def __init__(self, name):    
        self.__instance = None
        self.__connection = None
        self.__lock = Lock()
        self.__conn_str = name
    
    def __getConnection(self):
        if (self.__connection == None):
            # application_name = ";APP={0}".format(socket.gethostname())
            self.__connection = pyodbc.connect("{}".format(self.__conn_str))                  
        
        return self.__connection

    def __removeConnection(self):
        self.__connection = None

    # @retry(stop=stop_after_attempt(3), wait=wait_fixed(10), retry=retry_if_exception_type(pyodbc.OperationalError), after=after_log(app.logger, logging.DEBUG))
    def execure_query(self, query):
        result = {}  
        try:
            # conn = self.__getConnection()
            conn = pyodbc.connect("{}".format(self.__conn_str))
            crsr = conn.cursor()
            
            crsr.execute(query)

            result = crsr.fetchall()

            crsr.commit()
        except pyodbc.OperationalError as e:            
            # app.logger.error(f"{e.args[1]}")
            if e.args[0] == "08S01":
                # If there is a "Communication Link Failure" error, 
                # then connection must be removed
                # as it will be in an invalid state
                self.__removeConnection() 
                raise                        
        finally:
            crsr.close()
                         
        return result


class ADFConnect(object):
    def __init__(self):
        # This data should be stored in an Azure Vault instead of Robocorp
        "put here"

    def create_adf_client(self):
        
        credentials = ClientSecretCredential(
            tenant_id = self.__tenant_id,
            client_id = self.__client_id,
            client_secret = self.__client_secret
        )
        return DataFactoryManagementClient(credentials, self.__subscription_id)

    def trigger_pipeline(self, pipeline_name: str) -> str:
        
        run_response = self.__adf_client.pipelines.create_run(
            self.__resource_group_name, self.__datafactory_name, pipeline_name
        )

        return run_response.run_id
    
    def wait_for_pipeline_to_finish(self, run_id: str, timeout: int = 300):

        waited_seconds = 0
        while True:
            time.sleep(30)
            run_response = self.__adf_client.pipeline_runs.get(
                self.__resource_group_name, self.__datafactory_name, run_id
            )
            if run_response.status == 'Succeeded':
                print('Pipeline run succeeded')
                break
            elif run_response.status in ['Failed', 'Canceling', 'Cancelled']:
                raise Exception('Pipeline failed')

            waited_seconds += 30
            if waited_seconds > timeout:
                raise Exception('Pipeline timed out')

class AdlsLogHandler(AdlsHandler):

    def __init__(self, adls_name: str = Credentials.Adls.log_storage_account, accnt_key: str = Credentials.Adls.log_storage_key):
        super().__init__(adls_name, accnt_key)
        self.flow_url = "https://prod-163.westus.logic.azure.com:443/workflows/fb905ee39c844f2eae6944de0d8e369e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=52BmCrWj9UFuZahdAfwZipZf0PR0pnl1TUIzDlGSosg"
        self.headers = {
            "Content-Type" : "application/json"
        }

    def save_run_logs(self, entity: dict) -> None:
        """
        It takes a dictionary as an argument, converts it to a JSON string, and then posts it to a URL
        
        :param entity: dict
        :type entity: dict
        :return: The status code of the response.
        """
        self.body = json.dumps(entity)
        response = requests.post(url=self.flow_url, headers=self.headers, data=self.body)
        return response.status_code
    
    def upload(self, local_path: Path, fname: str, adls_path: Path) -> bool:
        """
        This function uploads a file from a local path to an Azure Data Lake Storage path with a specified
        content type.
        
        :param local_path: The local directory path where the file to be uploaded is located
        :type local_path: Path
        :param fname: The name of the file to be uploaded to the Azure Data Lake Storage (ADLS) account
        :type fname: str
        :param adls_path: The path to the Azure Data Lake Storage directory where the file will be uploaded
        :type adls_path: Path
        :return: a boolean value of `True`.
        """
        directory_client = self.conn.get_directory_client(f'{adls_path}')
        file_client = directory_client.get_file_client(fname)
        local_file = open(local_path / fname, 'rb')
        file_contents = local_file.read()
        if fname == "automation.log":
            # Set the content type to "text/plain"
            content_settings = ContentSettings(content_type='text/plain')
            # Upload the file with the specified content type
            file_client.upload_data(file_contents, overwrite=True, content_settings=content_settings)
        else:
            file_client.upload_data(file_contents, overwrite=True)


        return True
