from setuptools import setup, find_packages
import os
import sys

# Post-install hook to prompt user about adding Scripts folder to PATH
def post_install():
    try:
        from src.slt_converter import path_entry
        path_entry.ensure_path()
    except Exception as e:
        print(f"Could not run post-install PATH check: {e}")

setup(
    name="slt-converter",
    version="1.0.0",
    author="Ethan Brand",
    author_email="brandet251@gmail.com",
    description="File converter currently only supports converting QPW files to XLSX format.",
    long_description=open("README.md", "r", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/TryingCoder/SLT-Converter",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "tqdm",
        "pandas",
        "openpyxl",
    ],
    entry_points={
        "console_scripts": [
            "slt-converter=src.slt_converter.converter:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)

# Run post-install hook if this is a direct install
if 'install' in sys.argv:
    post_install()
