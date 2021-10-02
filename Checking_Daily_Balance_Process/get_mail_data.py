# -*- coding: utf-8 -*-
"""
Created on Fri Apr 16 12:31:33 2021

@author: rikushwa
"""

import xml.etree.ElementTree as ET
import os

tree = ET.parse(r'{}\mail_config.xml'.format(os.getcwd()))

def getData(Data):
    root = tree.getroot()
    for ele in root.findall('add'):
        key = ele.get('key')
        value = ele.get('value')
        if(Data == key):
            result = value
    
    return result