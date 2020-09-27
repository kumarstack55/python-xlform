# -*- coding: utf-8 -*-
from setuptools import setup  # type: ignore

packages = ["xlform", "xlform.engine"]

package_data = {"": ["*"]}

install_requires = ["openpyxl>=3.0.3,<4.0.0"]

setup_kwargs = {
    "name": "xlform",
    "version": "0.1.0",
    "description": "",
    "long_description": None,
    "author": "kumarstack55",
    "author_email": "kumarstack55@gmail.com",
    "maintainer": None,
    "maintainer_email": None,
    "url": None,
    "packages": packages,
    "package_data": package_data,
    "install_requires": install_requires,
    "python_requires": ">=3.6.1,<4.0.0",
}


setup(**setup_kwargs)
