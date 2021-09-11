# -*- coding: utf-8 -*-
"""
@author: BMONISH

Make DB Connection and Splunk Connection
"""

import sys
import os

#Dir containing Splunk SDK Lib files.
__file__ = file_path

#Identifies the Splunk Library Path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "lib")) 

## splunk lib
import splunklib.client as client
import splunklib.results as results

import pandas as pd

def db_splunk_connection():
    '''This function connects and logs in to a Splunk instance.'''
    
    host=hotname
    app=applicationname
    port=8089
    
    username = username
    password = pswd
    
    con = client.connect(host = host,port=port,username=username,password=password,app=app)
    
    return con

def oneshot_search(query,param):
    
    con = db_splunk_connection()
    response = con.jobs.oneshot(query,**param)
    result = results.ResultsReader(response)

    temp_list = []
    #traversing through the splunk result
    for line in result:
        line_dict = dict(line)
        temp_list.append(line_dict)

    #create dataframe to store all ctq details
    #df = pd.DataFrame(temp_list)

    return temp_list