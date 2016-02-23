# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        ia__methods.py
# Purpose:     store custom methods for wrapper class of IA_ Interface, the Base 
#              interface for all analysis windows.  
# Licence:     MIT License
#-------------------------------------------------------------------------------
"""Store custom methods for wrapper class of IA_ Interface.
   name := repr(zos_obj).split()[0].split('.')[-1].lower() + '_methods.py' 
"""
from __future__ import print_function
from __future__ import division
from win32com.client import CastTo as _CastTo, constants as _constants
from pyzos.zosutils import (wrapped_zos_object as _wrapped_zos_object,
                            replicate_methods as _replicate_methods)

# Overridden methods
# ------------------
def GetResults(self):
    return _wrapped_zos_object(self._ia_.GetResults())

def GetSettings(self):
    # the IA_.GetSettings() returns IAS_ which is the base class of all other settings 
    # objects such as IAS_FftMtf, IAS_FftMap, etc. The base-class IAS_ objects needs to be 
    # "specialized" for a particular analysis using the _CastTo function. When we apply 
    # CastTo(), the specialized objects "gain" the specific analysis functions and properties.
    # before returning, when we invoke _wrapped_zos_object(), the base-class methods and 
    # properties also gets patched. 
    
    settings_base = self._ia_.GetSettings()
    
    # CastTo the specialized class to add properties 
    if self._ia_.AnalysisType == _constants.AnalysisIDM_FftMtf: 
        settings = _CastTo(settings_base, 'IAS_FftMtf')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftMtfMap: 
        settings = _CastTo(settings_base, 'IAS_FftMtfMap')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftMtfvsField: 
        settings = _CastTo(settings_base, 'IAS_FftMtfvsField')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftSurfaceMtf: 
        settings = _CastTo(settings_base, 'IAS_FftSurfaceMtf')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftThroughFocusMtf: 
        settings = _CastTo(settings_base, 'IAS_FftThroughFocusMtf')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricMtf: 
        settings = _CastTo(settings_base, 'IAS_GeometricMtf')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricMtfMap: 
        settings = _CastTo(settings_base, 'IAS_GeometricMtfMap')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricMtfvsField: 
        settings = _CastTo(settings_base, 'IAS_GeometricMtfvsField')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricThroughFocusMtf: 
        settings = _CastTo(settings_base, 'IAS_GeometricThroughFocusMtf')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensMtf: 
        settings = _CastTo(settings_base, 'IAS_HuygensMtf')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensMtfvsField: 
        settings = _CastTo(settings_base, 'IAS_HuygensMtfvsField')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensSurfaceMtf: 
        settings = _CastTo(settings_base, 'IAS_HuygensSurfaceMtf')
    elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensThroughFocusMtf: 
        settings = _CastTo(settings_base, 'IAS_HuygensThroughFocusMtf')
        
    # create the settings object 
    settings = _wrapped_zos_object(settings)

    return settings

# Extra methods
# -------------