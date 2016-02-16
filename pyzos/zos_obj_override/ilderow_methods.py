# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        ilderow_methods.py
# Purpose:     store custom methods for wrapper class of ILDERow Interface
# Licence:     MIT License
#-------------------------------------------------------------------------------
"""Store custom methods for wrapper class of ILDERow Interface, which contains
   all data for a LDE surface. This interface can be accessed via the 
   ILensDataEditor interface. 
   name := repr(zos_obj).split()[0].split('.')[-1].lower() + '_methods.py' 
"""
from __future__ import print_function
from __future__ import division
from win32com.client import CastTo as _CastTo, constants as _constants
from pyzos.zosutils import wrapped_zos_object as _wrapped_zos_object
