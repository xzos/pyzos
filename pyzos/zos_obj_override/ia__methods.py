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
    # the GetSettings() method of IA_ interface returns IAS_ interface object 
    # which is the base class of all other settings interface. IAS_ contains a 
    # set of methods intended to be inherited by the specialized settings classes
    # such as IAS_FftMtf, IAS_FftMap, etc. 
    # If we just return the IA_ interface only the common set of methods are available
    # to the caller. In order to use the properties specific to a particular analysis
    # the caller needs to cast the interface.
    # If we merely cast the IAS_ object to the specialized interface in order to access
    # the properties, then the common set of methods from the base class (IAS_) are no
    # longer available. Thefore, we do the following to ensure that the returned 
    # settings object has both the common set of methods and the specific properties.
    
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
    
    # add the common set of methods of the base settings interface  
    _replicate_methods(srcObj=settings_base, dstObj=settings)
    return settings

# Extra methods
# -------------