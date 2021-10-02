# -*- coding: utf-8 -*-
"""
Created on Fri Apr 16 11:39:46 2021

@author: rikushwa
"""

import xml.etree.ElementTree as ET
import os

tree = ET.parse(r'{}\db2_config.xml'.format(os.getcwd()))

def getData(Data):
    root = tree.getroot()
    for ele in root.findall('add'):
        key = ele.get('key')
        value = ele.get('value')
        if(Data == key):
            result = value
    
    return result