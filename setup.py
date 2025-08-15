from setuptools import setup, find_packages
import sys
import os

# Post-install script for Windows PATH addition
def post_install():
    if os.name == "nt":
        try:
            import slt_converter.path_entry as path_entry
            path_entry.add_scripts_to_path()
        except Exception:
            pass

setup(
    name="slt-converter",
    version="1.0.0",
    description="AIO converter in progress. Currently supports .qpw to .xlsx conversion.",
    author="Your Name",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "tqdm",
        "pandas",
        "openpyxl"
    ],
    entry_points={
        "console_scripts": [
            "slt-converter=slt_converter.converter:main",
        ],
    },
    python_requires=">=3.6",
)

if __name__ == "__main__" and "install" in sys.argv[1:]:
    post_install()
