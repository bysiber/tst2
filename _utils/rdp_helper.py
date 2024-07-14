import paramiko
from _utils.logger import logger
from _utils.config import Credentials


class RDPConnection:
    """
    Usage example:

    local_path = 'path/to/your/local/file.txt'
    remote_path = '/path/on/server/file.txt'

    with RDPConnection() as conn:
        conn.put(local_path, remote_path)
    """
    default_retry = 3

    def __init__(self,
                 server_address: str = Credentials.RDPOracle.server_address,
                 port: int = int(Credentials.RDPOracle.port),
                 username: str = Credentials.RDPOracle.username,
                 password: str = Credentials.RDPOracle.password):
        self.ssh_client = paramiko.SSHClient()
        self.ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        self.server_address = server_address
        self.port = port
        self.username = username
        self.password = password

        self._connect()

    def _connect(self) -> None:
        try:
            self.ssh_client.connect(self.server_address, port=self.port, username=self.username, password=self.password)
            logger.info("SSH connection successfully established.")
        except Exception as e:
            logger.error(f"Failed to establish SSH connection: {e}")
            raise

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._close()

    def _close(self) -> None:
        if self.ssh_client:
            self.ssh_client.close()
            logger.info("SSH connection closed.")

    def put(self, local_filepath: str, target_filepath: str) -> None:
        last_exception = None
        for i in range(self.default_retry):
            try:
                with self.ssh_client.open_sftp() as sftp:
                    sftp.put(local_filepath, target_filepath)
                    logger.info(f"File '{local_filepath}' successfully uploaded to '{target_filepath}'.")
                    return
            except Exception as e:
                last_exception = e
                logger.error(f"Attempt {i + 1}: Failed to upload file. Reconnecting. Error: {e}")
                self._connect()
        logger.error(f"Failed to upload file after {self.default_retry} attempts. Last error: {last_exception}")
        raise last_exception

    def _clean_dir(self, dir_to_clean: str) -> None:
        """
        For internal and development usage only.

        dir_to_clean: str, example: '/TEST'
        """
        last_exception = None
        for i in range(self.default_retry):
            try:
                with self.ssh_client.open_sftp() as sftp:
                    files = sftp.listdir(dir_to_clean)
                    for file in files:
                        filepath = f'{dir_to_clean}/{file}'
                        sftp.remove(filepath)
                        logger.info(f"File '{filepath}' was removed.")
                    return

            except Exception as e:
                last_exception = e
                logger.error(f"Attempt {i + 1}: Failed to remove file. Reconnecting. Error: {e}")
                self._connect()
        logger.error(f"Failed to remove file after {self.default_retry} attempts. Last error: {last_exception}")
        raise last_exception

    def list_dir(self, remote_path: str) -> list:
        """List directory contents on the remote server."""
        try:
            with self.ssh_client.open_sftp() as sftp:
                return sftp.listdir(remote_path)
        except Exception as e:
            logger.error(f"Failed to list directory '{remote_path}': {e}")
            raise

    def download_file(self, remote_filepath: str, local_filepath: str) -> None:
        """Download a file from the remote server."""
        last_exception = None
        for i in range(self.default_retry):
            try:
                with self.ssh_client.open_sftp() as sftp:
                    sftp.get(remote_filepath, local_filepath)
                    logger.info(f"File '{remote_filepath}' successfully downloaded to '{local_filepath}'.")
                    return
            except Exception as e:
                last_exception = e
                logger.error(f"Attempt {i + 1}: Failed to download file. Reconnecting. Error: {e}")
                self._connect()
        logger.error(f"Failed to download file after {self.default_retry} attempts. Last error: {last_exception}")
        raise last_exception
