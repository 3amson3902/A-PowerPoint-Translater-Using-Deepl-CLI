from setuptools import setup, find_packages

setup(
    name="A PowerPoint Translator Using DeepL",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        "python-pptx",
        "deepl",
        "tkinter"
    ],
    entry_points={
        "console_scripts": [
            "ppt_translator = ppt_translator.main:main_run",
        ],
    },
    author="3amson3902",
    author_email="samson3902@gmail.com",
    description="A GUI application to translate text within PowerPoint presentations using the DeepL translation service.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/3amson3902/MSTranslator",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)