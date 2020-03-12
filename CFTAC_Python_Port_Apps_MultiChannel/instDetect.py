# -*- coding: utf-8 -*-
"""
Created on Mon Jan 22 19:53:28 2018

@author: beschanz
"""

import visa

def detect_inst(keyword):
    rm = visa.ResourceManager()
    instrumentlist = rm.list_resources()
    
    print(instrumentlist)
    
    instID = {'Name': "", 'Resource': ""}
    
    for instStr in instrumentlist:
        if "ASRL" in instStr:
            continue
        if "redstub" in instStr.lower():
            continue
        try:
          #print("-------LAST_INST_STR----" + instStr + "\n")
          inst = rm.open_resource(instStr)
          try:
              instName = str(inst.query("*IDN?"))
              print(instName + "\n")
              if keyword in instName:
                  instID['Name'] = instName
                  instID['Resource'] = inst
                  break
          except:
              print("Error in querying: " + instStr)
        except:
          print("Error in opening: " + instStr)
    
    if instID['Name'] is "":
        raise ValueError('inst ' + str(keyword) + " not identified in given visa resources")
        return
    print(instID)
    
    return instID