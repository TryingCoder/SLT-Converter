import platform
from setuptools import setup, find_packages
from setuptools.command.install import install


class PostInstall(install):
    """Custom install step to run PATH prompt on Windows."""
    def run(self):
        install.run(self)  # normal install first
        if platform.system() == "Windows":
            try:
                from src.slt_converter import path_entry
                path_entry.main()
            except Exception as e:
                print(f"[!] PATH setup skipped: {e}")


setup(
    name="slt-converter",
    version="1.0.0",
    author="Your Name",
    author_email="you@example.com",
    description="Convert SLT files using LibreOffice CLI tools",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.7",
    install_requires=[],
    entry_points={
        "console_scripts": [
            "slt-convert=slt_converter.converter:main",  # CLI command
        ],
    },
    cmdclass={
        "install": PostInstall,
    },
)
