# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        imeritfunctioneditor_methods.py
# Purpose:     store custom methods for wrapper class of IFields Interface
# Licence:     MIT License
#-------------------------------------------------------------------------------
"""Store custom methods for wrapper class of IMeritFunctionEditor Interface, 
   which interface defines all properties and methods needed to interact with 
   the MRit Function Editor. 
   name := repr(zos_obj).split()[0].split('.')[-1].lower() + '_methods.py' 
"""
from __future__ import print_function
from __future__ import division
from win32com.client import CastTo as _CastTo, constants as _constants
from pyzos.zosutils import wrapped_zos_object as _wrapped_zos_object


# Overridden methods
# ------------------
def AddRow(self):
    """Adds a new row at the end of the editor."""
    base_mfe = _CastTo(self._imeritfunctioneditor, 'IEditor')
    return _wrapped_zos_object(base_mfe.AddRow())


def GetRowAt(self, pos):
    """Gets the row at the specified index (0 to NumberOfRows-1).
    @pos : The row index.
    """
    base_mfe = _CastTo(self._imeritfunctioneditor, 'IEditor')
    return _wrapped_zos_object(base_mfe.GetRowAt(pos))


# Overridden properties
# ---------------------
@property
def pSEQOptimizationWizard(self):
    """Get the Sequential Optimization Wizard """
    return _wrapped_zos_object(self._imeritfunctioneditor.SEQOptimizationWizard)

#@property
#def pNSCOptimizationWizard(self):
#    """Get the Non-Sequential Optimization Wizard"""
#    return _wrapped_zos_object(self._imeritfunctioneditor.NSCOptimizationWizard)


