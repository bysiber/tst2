from pathlib import Path

from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.sharing.sharing_link_kind import SharingLinkKind

from _utils.config import Credentials
from _utils.logger import logger


class SharePointException(Exception):
    """Custom SharePoint exception"""


class SharePoint:
    """SharePoint utility class to use the SharePoint REST API

    :type email: str
    :param email: Email address for the user account
    :type password: str
    :param password: Password for the user account
    :type site_name: str
    :param site_name: Site name
        e.g. If site name is Development, the site url will be
        https://blueoceansp.sharepoint.com/sites/Development
    """

    def __init__(
        self,  email: str = Credentials.SharePoint.email, password: str = Credentials.SharePoint.password,
        site_name: str = Credentials.SharePoint.site_name
    ) -> None:

        self._email = email
        self._password = password
        self._site_name = site_name
        self._site_url = f"https://blueoceansp.sharepoint.com/sites/{self._site_name}"
        self._documents = f"/sites/{self._site_name}/Shared Documents"
        self._client_context = None

        self._connect()

    def _connect(self) -> None:
        """Authenticate with user credentials and obtain ClientContext for the site url"""

        user_credentials = UserCredential(self._email, self._password)
        self._client_context = ClientContext(self._site_url).with_credentials(
            user_credentials
        )

    def list_contents(self, relative_folder_path: Path) -> dict:
        """Get files and folders under the given path

        Raises SharePointException if the given folder path does not exist.

        :type relative_folder_path: Path
        :param relative_folder_path: Relative folder path on server
        :rtype: dict
        :returns: Dictionary of files and folders
        """

        logger.info(f"Getting contents from {relative_folder_path}")

        if not self._folder_exists(relative_folder_path):
            raise SharePointException(f"Folder path[{relative_folder_path}] does not exist!")

        root_folder = self._client_context.web.get_folder_by_server_relative_path(
            f"{self._documents}/{relative_folder_path}"
        )

        root_folder.expand(["Files", "Folders"]).get().execute_query()

        contents = {"files": [], "folders": []}

        for folder in sorted(root_folder.folders, key=lambda x: x.name):
            contents["folders"].append(folder.name)

        for file in sorted(root_folder.files, key=lambda x: x.name):
            contents["files"].append(file.name)

        return contents

    def get_file_properties(self, relative_filepath: Path) -> dict:
        """Get properties of the given file

        Raises SharePointException if an error occurs.

        The returned dictionary contains the following fields

        {
            "CheckInComment": "",
            "CheckOutType": 2,
            "ContentTag": "{3616457A-76C0-4CE6-9071-7B9CF6CFAA08},47,88",
            "CustomizedPageStatus": 0,
            "ETag": "\"{3616457A-76C0-4CE6-9071-7B9CF6CFAA08},47\"",
            "Exists": true,
            "IrmEnabled": false,
            "Length": "23669",
            "Level": 1,
            "LinkingUri": "https://blueoceansp.sharepoint.com/sites/Development/Shared%20Documents/...",
            "LinkingUrl": "https://blueoceansp.sharepoint.com/sites/Development/Shared Documents/...",
            "MajorVersion": 6,
            "MinorVersion": 0,
            "Name": "Working with Blue Ocean.docx",
            "ServerRelativeUrl": "/sites/Development/Shared Documents/Projects/Blue Ocean/Working with Blue Ocean.docx",
            "TimeCreated": "2021-11-10T21:30:36Z",
            "TimeLastModified": "2021-11-11T01:43:14Z",
            "Title": "",
            "UIVersion": 3072,
            "UIVersionLabel": "6.0",
            "UniqueId": "3616457a-76c0-4ce6-9071-7b9cf6cfaa08"
        }

        :type relative_filepath: Path
        :param relative_filepath: Relative file path on server
        :rtype: dict
        :returns: File properties dictionary
        """

        file_url = f"{self._documents}/{relative_filepath}"

        try:
            file = (
                self._client_context.web.get_file_by_server_relative_url(file_url)
                .get()
                .execute_query()
            )

            return file.properties

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(
                f"Failed to get file properties for {relative_filepath}!"
            ) from exp

    def download(self, relative_filepath: Path, local_filepath: Path) -> None:
        """Download file from the given path to local path

        Raises SharePointException if the given file path does not exists.

        :type relative_filepath: Path
        :param relative_filepath: Relative file path on server
        :type local_filepath: Path
        :param local_filepath: Local file path to save downloaded file
        """

        download_path = f"{self._documents}/{relative_filepath}"

        logger.info(f"Downloading {download_path}")

        try:
            with open(local_filepath, "wb") as local_file:
                file = self._client_context.web.get_file_by_server_relative_path(download_path)
                file.download(local_file).execute_query()

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"{download_path} does not exists!") from exp

        logger.info(f"File downloaded to {local_filepath}")

    def upload(self, local_filepath: Path, target_filepath: Path) -> None:
        """Upload file to the target path on server

        Raises SharePointException if the given paths do not exist
        or upload operation fails.

        :type local_filepath: Path
        :param local_filepath: Local file path to be uploaded
        :type target_filepath: Path
        :param target_filepath: Relative target file path on server
        """

        if not local_filepath.exists():
            raise SharePointException(f"{local_filepath} does not exists!")

        target_folder_path = target_filepath.parent

        if not self._folder_exists(target_folder_path):
            raise SharePointException(f"Target folder[{target_folder_path}] does not exist!")

        logger.info(f"Uploading {local_filepath.name} to {target_filepath}")

        with open(local_filepath, "rb") as local_file:
            file_content = local_file.read()

        try:
            target_folder = self._client_context.web.get_folder_by_server_relative_path(
                f"{self._documents}/{target_folder_path}"
            )
            target_file = target_folder.upload_file(target_filepath.name, file_content).execute_query()

        except ClientRequestException as exp:
            logger.debug(exp)

            message = f"Failed to upload {local_filepath} to {target_filepath}"

            if "SPFileLockException" in exp.code:
                message += "\nFile is being edited by someone else and locked!"

            raise SharePointException(message) from exp

        logger.info(f"File has been uploaded to {target_file.serverRelativeUrl}")

    def move(self, source_filepath: str, target_filepath: Path) -> None:
        """Move the given file to the given path.

        Raises SharePointException if delete operation fails.

        :type relative_filepath: Path
        :param relative_filepath: Relative file path on server
        """
        # TODO: Add folder creation if target is not existing

        # target_folder_path = target_filepath.parent
        # source_folder_path = source_filepath.parent
        # if not self._folder_exists(source_folder_path):
        #     raise SharePointException(f"Source folder [{source_folder_path}] does not exist!")

        # if not self._folder_exists(target_folder_path):
        #     raise SharePointException(f"Target folder [{target_folder_path}] does not exist!")

        # logger.info(f"Moving {source_filepath} to {target_filepath}")
        # try:
        #     source_file_url = f"{self._documents}/{source_filepath}"
        #     source_file = (
        #         self._client_context.web.get_file_by_server_relative_url(source_file_url).get().execute_query())
        #     target_folder_url = f"{self._documents}/{target_folder_path}"
        #     target_folder = (
        #         self._client_context.web.get_folder_by_server_relative_path(target_folder_url).get().execute_query())
        #     target_file_url = f'{target_folder.serverRelativeUrl}/{target_filepath.name}'

        #     source_file.moveto(target_file_url, 1).execute_query()

        # except ClientRequestException as exp:
        #     logger.debug(exp)
        #     raise SharePointException(f"Failed to move {source_filepath} to {target_filepath}") from exp

        # logger.info("File moved successfully")

        source_file_url = f"{self._documents}/{source_filepath}"
        target_file_url = f"{self._documents}/{target_filepath}"

        logger.info(f"Moving file from {source_filepath} to {target_filepath}")

        try:
            file = self._client_context.web.get_file_by_server_relative_url(source_file_url)
            file.moveto(target_file_url, 1)
            self._client_context.execute_query()

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"Failed to move file from {source_filepath} to {target_filepath}") from exp

        logger.info(f"Moved file from {source_filepath} to {target_filepath}")

    def delete(self, relative_filepath: Path) -> None:
        """Delete the given file from server

        Raises SharePointException if delete operation fails.

        :type relative_filepath: Path
        :param relative_filepath: Relative file path on server
        """

        file_url = f"{self._documents}/{relative_filepath}"

        logger.info(f"Deleting {file_url}")

        try:
            file = self._client_context.web.get_file_by_server_relative_url(file_url)
            file.delete_object().execute_query()

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"Failed to delete {file_url}") from exp

        logger.info(f"Deleted {file_url}")

    def share(self, relative_filepath: Path, permission: SharingLinkKind) -> str:
        """Share the file with the given permissions

        Raises SharePointException if file sharing fails.

        :type relative_filepath: Path
        :param relative_filepath: Relative file path to share
        :type permission: SharingLinkKind
        :param permission: Permission level for the shared file
        :rtype: str
        :returns: Url of shared file
        """

        file_url = f"{self._documents}/{relative_filepath}"

        logger.info(f"Sharing {file_url}")

        try:
            target_file = self._client_context.web.get_file_by_server_relative_url(file_url)
            result = target_file.share_link(permission).execute_query()
            link_url = result.value.sharingLinkInfo.Url

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"Failed to share {file_url}!") from exp

        logger.info(f"Share link: {link_url}")

        return link_url

    def unshare(self, relative_filepath: Path, permission: SharingLinkKind) -> bool:
        """Unshare the given file

        Raises SharePointException if file unsharing fails.

        :type relative_filepath: Path
        :param relative_filepath: Relative file path to unshare
        :type permission: SharingLinkKind
        :param permission: Permission level for the shared file
        :rtype: bool
        :returns: True if file was unshared successfully
        """

        file_url = f"{self._documents}/{relative_filepath}"

        logger.info(f"Unsharing {file_url}")

        try:
            target_file = self._client_context.web.get_file_by_server_relative_url(file_url)
            target_file.unshare_link(permission).execute_query()

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"Failed to unshare {file_url}!") from exp

        logger.info(f"Unshared {file_url}")

        return True

    def create_folder(self, target_folder_path: Path) -> None:
        """Create folder on server

        Nested folders can also be created.

        e.g. sharepoint_client.create_folder("Projects/Folder1/Folder2")

        Raises SharePointException if folder creation fails.

        :type target_folder_path: Path
        :param target_folder_path: Relative target folder path to create on server
        """

        target_folder_url = f"/{self._documents.split('/')[-1]}/{target_folder_path}"

        try:
            target_folder = self._client_context.web.ensure_folder_path(
                target_folder_url
            ).execute_query()

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"Failed to create {target_folder_path}") from exp

        logger.info(f"Created {target_folder.serverRelativeUrl}")

    def delete_folder(self, target_folder_path: Path) -> None:
        """Delete the given folder from server

        Raises SharePointException if folder does not exist on server
        or deletion fails.

        :type target_folder_path: Path
        :param target_folder_path: Relative target folder path to delete from server
        """

        folder_path = f"{self._documents}/{target_folder_path}"

        logger.info(f"Deleting {folder_path}")

        if not self._folder_exists(target_folder_path):
            raise SharePointException(f"Target folder[{target_folder_path}] does not exist!")

        try:
            folder = self._client_context.web.get_folder_by_server_relative_path(folder_path)
            folder.delete_object().execute_query()

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"Failed to delete {folder_path}") from exp

        logger.info(f"Deleted {folder_path}")

    def share_folder(self, target_folder_path: Path, permission: SharingLinkKind) -> str:
        """Share the folder with the given permissions

        Raises SharePointException if folder sharing fails or folder does not exist.

        :type target_folder_path: Path
        :param target_folder_path: Relative target folder path to share
        :type permission: SharingLinkKind
        :param permission: Permission level for the shared file
        :rtype: str
        :returns: Url of shared folder
        """

        target_folder_url = f"{self._documents}/{target_folder_path}"

        logger.info(f"Sharing {target_folder_url}")

        if not self._folder_exists(target_folder_path):
            raise SharePointException(f"Target folder[{target_folder_path}] does not exist!")

        try:
            target_folder = self._client_context.web.get_folder_by_server_relative_url(
                target_folder_url
            )
            result = target_folder.share_link(permission).execute_query()
            link_url = result.value.sharingLinkInfo.Url

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"Failed to share {target_folder_url}!") from exp

        logger.info(f"Share link: {link_url}")

        return link_url

    def unshare_folder(self, target_folder_path: Path, permission: SharingLinkKind) -> bool:
        """Unshare the given folder

        Raises SharePointException if folder sharing fails or folder does not exist.

        :type target_folder_path: Path
        :param target_folder_path: Relative target folder path to share
        :type permission: SharingLinkKind
        :param permission: Permission level for the shared folder
        :rtype: bool
        :returns: True if folder was unshared successfully
        """

        target_folder_url = f"{self._documents}/{target_folder_path}"

        logger.info(f"Unsharing {target_folder_url}")

        if not self._folder_exists(target_folder_path):
            raise SharePointException(f"Target folder[{target_folder_path}] does not exist!")

        try:
            target_folder = self._client_context.web.get_folder_by_server_relative_url(
                target_folder_url
            )
            target_folder.unshare_link(permission).execute_query()

        except ClientRequestException as exp:
            logger.debug(exp)

            raise SharePointException(f"Failed to unshare {target_folder_url}!") from exp

        logger.info(f"Unshared {target_folder_url}")

        return True

    def _folder_exists(self, relative_folder_path: Path) -> bool:
        """Check if the given folder path exists on server

        :type relative_folder_path: Path
        :param relative_folder_path: Relative folder path on server
        :rtype: bool
        :returns: Whether given folder exists
        """

        folder_url = f"{self._documents}/{relative_folder_path}"

        try:
            self._client_context.web.get_folder_by_server_relative_url(
                folder_url
            ).get().execute_query()

        except ClientRequestException as exp:
            if exp.response.status_code == 404:
                return False

        return True
