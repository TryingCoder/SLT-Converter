from setuptools import setup, find_packages

setup(
    name="slt-converter",
    version="0.1.0",
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
            "slt-converter=converter:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)
