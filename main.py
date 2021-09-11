# -*- coding: utf-8 -*-
"""
@author: BMONISH

Automatically pulls the data from Splunk and updates the Business Deck Powerpoint slides.

Slides updated :
    1. FEDEBOM CTQ Performance
    2. FEDEBOM User Metrics
    3. Monthly BOM Usage by Role 
"""

''' Main method to connect to other Python Modules'''
from getpass import getpass


if __name__ == '__main__':
    
    from FEDEBOMCTQPerformance import update_slide
    from BOMUsageByRole import bom_usage
    from FEDEBOMUserMetrics import user_metrics
    
    
    presentation_file = 'Eng BOM Prod Sup And CTQ Review 08-13-2021.pptx'    
     
    update_slide(presentation_file)   #slide2
    bom_usage(presentation_file)      #slide4
    #user_metrics(presentation_file)   #slide3