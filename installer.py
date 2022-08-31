import requests
import zipfile
import subprocess

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
    rdp_dir = "C:\\Program Files\\RDP Wrapper test"

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

    def setup_files(self):
        """Installs and unpacks the zip files"""
        print("Setting up the required files...")

        self.download(self.zip_link, "images/temp/")
        self.download(self.updater_link, "images/temp/")
        self.unpack("images/tempautoupdate.zip", self.rdp_dir)
        self.unpack("images/temp/RDPWrap-v1.6.2.zip", self.rdp_dir)

        print("Files should all be in place!")

    def run_installers(self):
        """Runs the bat files"""
        print("Running installs...")

        self.run_batch("\\helper\\autoupdate__enable_autorun_on_startup.bat")
        self.run_batch("\\autoupdate.bat")

    def run(self):
        print("Setting up your RDPWrapper...")

        self.setup_files()
        self.run_installers()
        print(
            "Installation complete, the hardest part should be done!\n"
            "Please refer to the guide in discord for the last steps."
        )