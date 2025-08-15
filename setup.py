from setuptools import setup, find_packages

setup(
    name="slt-converter",
    version="1.0.0",
    description="Batch convert QPW files to XLSX using LibreOffice CLI",
    author="TryingCoder",
    packages=find_packages("src"),
    package_dir={"": "src"},
    install_requires=[
        "tqdm",
        "pandas",
        "openpyxl"
    ],
    entry_points={
        "console_scripts": [
            "slt = slt_converter.__main__:main"
        ]
    },
    python_requires=">=3.6",
    include_package_data=True,
    zip_safe=False,
)
