from setuptools import setup

setup(
    name="smartsheet_tools",
    version="0.0.2",
    description="A collection of convenience functions to aid with transitioning from simple-smartsheet to the SDK API and common tasks",
    author="Ashton Pooley",
    author_email="Ashton@ashi.digital",
    install_requires=[
        "smartsheet-python-sdk>=3.1.0",
    ],
)
