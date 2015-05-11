# -*- coding: utf-8 -*-

from distutils.core import setup 
import py2exe
import xml.etree.ElementTree
 
setup(name="Conversor", 
 version="1.0", 
 description="Conversor de xls a txt", 
 author="David Diz", 
 author_email="daviddiz@gmail.com", 
 scripts=["conversor_xls.py"], 
 console=["conversor_xls.py"], 
 options={"py2exe": 		{"packages":["encodings"],
                 "includes":["xlrd","datetime", "decimal"],
                 "bundle_files":2,
                 "optimize":2},}, 
 zipfile=None,
)