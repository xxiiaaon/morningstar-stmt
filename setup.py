import setuptools

with open("README.md", "r") as f:
    long_description = f.read()

setuptools.setup(
    name="morningstar_stmt",
    version="0.0.2",
    author="xxiiaaon",
    author_email="xxiiaaon@gmail.com",
    description="Morningstar Financials Statement Downloader",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/xxiiaaon/morningstar-stmt",
    platforms=['Any'],
    packages=['morningstar_stmt'],
    install_requires=[
        "selenium>=3.0.0",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)
