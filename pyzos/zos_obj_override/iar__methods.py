# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        iar__methods.py
# Purpose:     store custom methods for wrapper class of IAR_ Interface
# Licence:     MIT License
#-------------------------------------------------------------------------------
"""Store custom methods for wrapper class of IAR_ Interface.
   name := repr(zos_obj).split()[0].split('.')[-1].lower() + '_methods.py' 
"""
from __future__ import print_function
from __future__ import division
from win32com.client import CastTo as _CastTo, constants as _constants
from pyzos.zosutils import wrapped_zos_object as _wrapped_zos_object

# Overridden methods
# ------------------
#def GetDataGrid(self, index):
#    return _wrapped_zos_object(self._iar_.GetDataGrid(index))

# Extra methods
# -------------