from setuptools import setup, find_packages

setup(
    name="orange2df2excel",                           # The name of your package
    version="0.2",                                   # Initial version
    author="maxchergik",                              # Replace with your name
    author_email="maxchergik@gmail.com",           # Replace with your email
    description="Tools for working with DataFrames and saving to Excel",
    long_description=open("README.md").read(),       # Reads the content of README.md
    long_description_content_type="text/markdown",   # README.md format
    url="https://github.com/yourusername/orange2df2xcel",  # Replace with your GitHub repo URL
    packages=find_packages(),                        # Finds the sub-packages automatically
    install_requires=[                               # Dependencies for your package
        "pandas",
        "openpyxl",
        "koboextractor",
        "io"
    ],
    classifiers=[                                    # Optional metadata for package indexing
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",                         # Minimum Python version requirement
)
