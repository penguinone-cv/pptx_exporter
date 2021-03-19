from setuptools import setup, find_packages

with open('requirements.txt') as requirements_file:
    install_requirements = requirements_file.read().splitlines()

setup(
    name="pptx_exporter",
    version="0.0.1",
    description="Export .pptx file that summarizes log(s)",
    author="tomita",
    install_requires=install_requirements,
    license="MIT",
    classifiers=[
        "Development Status :: 4 - Beta"
    ])