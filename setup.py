from setuptools import setup, find_packages

setup(
    name="slt-converter",
    version="0.1.0 (beta)",
    description="QPW to XLSX converter using LibreOffice",
    author="Your Name",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.10",
    install_requires=[
        "tqdm>=4.65.0",
        "requests>=2.31.0",
        "beautifulsoup4>=4.12.2",
    ],
    entry_points={
        "console_scripts": [
            "slt_converter=converter:main"
        ],
    },
    include_package_data=True,
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
)
