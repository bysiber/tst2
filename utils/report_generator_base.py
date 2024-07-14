from abc import ABC, abstractmethod


class ReportGeneratorBase(ABC):
    """An abstract class for report generation classes"""

    @abstractmethod
    def create_reports(self):
        """
        Main method for report generation.

        The input file must be downloaded locally before.
        """
        pass

    @abstractmethod
    def upload_reports(self, download_directory: str, processed_directory: str):
        """
        Upload all generated reports to the Share Point
        """
        pass

    @abstractmethod
    def download_master_file(self):
        """
        Download Master File from the Share Point
        """
        pass
    
    @abstractmethod
    def send_report_emails(self):
        """
        Send all the corresponding reports to the VPOS
        """
        pass
