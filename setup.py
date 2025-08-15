from setuptools import setup, find_packages

setup(
    name="slt-converter",
    version="0.1.1",
    description="BETA | QPW to XLSX Converter - More features coming soon!",
    author="TryingCoder",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "pandas>=2.0",
        "openpyxl>=3.0",
        "tqdm>=4.60",
        "requests>=2.31.0",
    ],
    entry_points={
        "console_scripts": [
            "slt = slt_converter.__main__:main",
        ],
    },
    python_requires=">=3.10",
    include_package_data=True,
    zip_safe=False,
    url="https://github.com/TryingCoder/SLT-Converter",
)
