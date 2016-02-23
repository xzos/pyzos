# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        ilocaloptimization_methods.py
# Purpose:     store custom methods for wrapper class of ILocalOptimization Interface
# Licence:     MIT License
#-------------------------------------------------------------------------------
"""Store custom methods for wrapper class of ILocalOptimization Interface, which
   contains methods to run various system-wide tools.
   name := repr(zos_obj).split()[0].split('.')[-1].lower() + '_methods.py' 
"""
from __future__ import print_function
from __future__ import division
from win32com.client import CastTo as _CastTo, constants as _constants
from pyzos.zosutils import wrapped_zos_object as _wrapped_zos_object


# Overridden methods
# ------------------
