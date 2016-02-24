# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        i_analyses_methods.py
# Purpose:     store custom methods for wrapper class of I_Analyses
# Licence:     MIT License
#-------------------------------------------------------------------------------
"""Store custom methods for wrapper class of I_Analyses.
   name := repr(zos_obj).split()[0].split('.')[-1].lower() + '_methods.py' 
"""
from __future__ import print_function
from __future__ import division
from win32com.client import CastTo as _CastTo, constants as _constants
from pyzos.zosutils import wrapped_zos_object as _wrapped_zos_object

# overridden methods
# ------------------
# def Get_AnalysisAtIndex(self, index):
#     """Retrieve the specified analysis."""
#     return _wrapped_zos_object(self._i_analyses.Get_AnalysisAtIndex(index))

# def New_Analysis(self, analysis_type):
#     """Create a new analysis of the specified type.
#     @analysis_type is one of AnalysisIDM (enum)
#     """
#     return _wrapped_zos_object(self._i_analyses.New_Analysis(analysis_type))
    
# def New_ConfigurationMatrixSpot(self):
#     return _wrapped_zos_object(self._i_analyses.New_ConfigurationMatrixSpot())

# def New_DiffractionEncircledEnergy(self):
#     return _wrapped_zos_object(self._i_analyses.New_DiffractionEncircledEnergy())

# def New_FftMtf(self):
#     """Open a new FFT MTF window"""
#     return _wrapped_zos_object(self._i_analyses.New_FftMtf())

# def New_FftMtfMap(self):
#     return _wrapped_zos_object(self._i_analyses.New_FftMtfMap())

# Custom methods
# --------------