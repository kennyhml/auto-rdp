import requests
import zipfile
import subprocess
import ctypes
import os
from glob import glob


class Installer:
    """Automatic RDPWrap installer
    -------------------------------
    Automatically installs and sets up stascorp's rdpwrapper.

    The class is not compatible with other rdp wrappers.\n
    Creating a new windows user and managing permissions is still
    end-users responsibility.
    """

    zip_link = "https://github.com/stascorp/rdpwrap/releases/download/v1.6.2/RDPWrap-v1.6.2.zip"
    updater_link = "https://github.com/asmtron/rdpwrap/raw/master/autoupdate.zip"
    rdp_dir = "C:\\Program Files\\RDP Wrapper"

    def check_admin_perms(self) -> bool:
        """Check if the script is being ran as admin"""
        print("Checking for admin perms...")
        try:
            is_admin = os.getuid() == 0
        except AttributeError:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        if is_admin:
            print("Verified admin permissions...")
            return True

        ctypes.windll.user32.MessageBoxW(
            0,
            "Please run the bot as administrator to run the installer!",
            "Mising permissions!",
            0,
        )

    def already_installed(self) -> bool:
        """Checks if an rdp wrapper is already installed"""
        folders = glob("C:\\Program Files\\*")

        if not "C:\\Program Files\\RDP Wrapper" in folders:
            print("Verified that the file is not already installed...")
            return False

        override = (
            not (
                ctypes.windll.user32.MessageBoxW(
                    0,
                    "The program is already installed on your system, running the "
                    "intaller again may not be successful.\n\n Would you like to "
                    "run the installer again anyways?",
                    "RDP Wrapper already installed, overwrite current version?",
                    4,
                    0x40000,
                )
            )
            == 6
        )
        return override

    def download(self, url: str, target_dir: str) -> None:
        """Download zip from URL and manifest to the target directory"""
        print(f"Started downloading zip...")

        # downloading the files
        req = requests.get(url)
        filename = target_dir + url.split("/")[-1]

        # writing file to target directory
        with open(filename, "wb") as output_file:
            output_file.write(req.content)
        print("Download completed!")

    def unpack(self, filepath: str, target_dir: str) -> None:
        """Unpacks a zip to a target directory"""
        print(f"Unpacking {filepath} to {target_dir}")

        # unpack the zip to our targets folder
        with zipfile.ZipFile(filepath, "r") as zip_ref:
            zip_ref.extractall(target_dir)

    def run_batch(self, bat_path: str):
        """Runs a bat file by path, closes automatically"""
        path = self.rdp_dir + bat_path
        print(f"Executing {path}")
        p = subprocess.Popen(path)
        p.terminate()
        p.wait()

        print(f"Executed {path} successfully!")

    def setup_files(self) -> None:
        """Installs and unpacks the zip files"""
        print("Setting up the required files...")

        self.download(self.zip_link, "images/temp/")
        self.download(self.updater_link, "images/temp/")
        self.unpack("images/tempautoupdate.zip", self.rdp_dir)
        self.unpack("images/temp/RDPWrap-v1.6.2.zip", self.rdp_dir)

        print("Files should all be in place!")

    def run_installers(self) -> None:
        """Runs the bat files"""
        print("Running installs...")

        self.run_batch("\\helper\\autoupdate__enable_autorun_on_startup.bat")
        self.run_batch("\\autoupdate.bat")

    def run(self):
        """Sets up everything"""
        # check for admin perms (Program Files requires admin permission)
        # and that the file is not already on the system (avoid misclicks)
        if self.already_installed() or not self.check_admin_perms():
            return

        print("Permission and directory in place, running installers...")
        self.setup_files()
        self.run_installers()
        print(
            "Installation complete, the hardest part should be done!\n"
            "Please refer to the guide in discord for the last steps."
        )

if __name__ == "__main__":
    installer = Installer()
    installer.run()
