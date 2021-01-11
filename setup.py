import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name="TableauDesktopPy-bpewyllie", # Replace with your own username
    version="0.0.1",
    author="Brady Wyllie",
    author_email="bpewyllie@gmail.com",
    description="Tools for extracting metadata from Tableau Desktop workbook files.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/bpewyllie/tableaudesktoppy",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)