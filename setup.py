from setuptools import setup, find_packages

setup(
    name="slt-converter",
    version="1.0.0",
    description="Convert QPW files to XLSX using LibreOffice in headless mode",
    author="TryingCoder",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "pandas>=2.0",
        "openpyxl>=3.0",
        "tqdm>=4.60"
    ],
    entry_points={
        "console_scripts": [
            "slt = slt_converter.__main__:main",
        ],
    },
    python_requires=">=3.10",
    zip_safe=False,
)
