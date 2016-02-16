# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        zos.py
# Purpose:     COM interface class for easy ZOS-API
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
"""Helper functions for accessing Zemax ZOS API from Python in standalone mode. 
"""
from __future__ import division, print_function

import win32com.client as _comclient
import warnings as _warnings
from pyzos.zosutils import (ZOSPropMapper as _ZOSPropMapper, 
                            replicate_methods as _replicate_methods,
                            wrapped_zos_object as wrapped_zos_object)

#%%
Const = None  # Constants (placeholder)
#%%
class _ConnectionError(Exception): pass 
class _InitializationError(Exception): pass
class _ZOSSystemError(Exception): pass

#%%
class _PyZOSApp(object):
    app = None
    connect = None
    
    def __new__(cls):
        global Const
        if not cls.app:
            #TODO: Add relavent exceptions
            # ensure win32com support files for ZOSAPI_Interfaces are available,
            # generate if necessary.
            _comclient.gencache.EnsureModule('ZOSAPI_Interfaces', 0, 1, 0)
            edispatch = _comclient.gencache.EnsureDispatch
            cls.connect = edispatch('ZOSAPI.ZOSAPI_Connection')
            cls.app = cls.connect.CreateNewApplication()
            Const = type('Const', (), _comclient.constants.__dicts__[0]) # Constants class
        return cls.app

#%%
class OpticalSystem(object):
    """Wrapper for IOpticalSystem interface
    """
    _instantiated = False
    _pyzosapp = None
    
    # Managed properties (prefix with 'p' to ease identification)
    pIsNonAxial = _ZOSPropMapper('_iopticalsystem', 'IsNonAxial')
    pMode = _ZOSPropMapper('_iopticalsystem', 'Mode')
    pNeedsSave = _ZOSPropMapper('_iopticalsystem', 'NeedsSave')
    pSystemFile = _ZOSPropMapper('_iopticalsystem', 'SystemFile')
    pSystemID = _ZOSPropMapper('_iopticalsystem', 'SystemID')
    pSystemName = _ZOSPropMapper('_iopticalsystem', 'SystemName', setter=True)
    
    def __init__(self, mode=0):
        """ mode : sequential (0) or non-sequential (1)
        """
        
        if OpticalSystem._instantiated:
            self._iopticalsystem = OpticalSystem._pyzosapp.CreateNewSystem(mode) # wrapped object
        else:
            OpticalSystem._pyzosapp = _PyZOSApp()               # wrapped object
            self._iopticalsystem = OpticalSystem._pyzosapp.GetSystemAt(0) # PrimarySystem
            if mode == 1:
                self._iopticalsystem.MakeNonSequential()
            OpticalSystem._instantiated = True
        
        ## patch methods from IOpticalSystem to the instance
        _replicate_methods(self._iopticalsystem, self)
            
    def __getattr__(self, attrname):
        """handle any method calls that were not accounted for by 
           `_replicate_methods` (shouldn't happen generally), or 
           property calls that are not managed.
        """
        #print('__getattr__ in OpticalSystem called for', attrname) ##TODO: remove print later
        return getattr(self._iopticalsystem, attrname)
    
    @property
    def pConnectIsAlive(self):
        """ZOS-API connection active/inactive status"""
        return _PyZOSApp.connect.IsAlive
    
    @property
    def pAnalyses(self):
        """Gets the analyses for the current system (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.Analyses)
    
    @property
    def pLDE(self):
        """Gets the lens data editor interface (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.LDE)
    
    @property
    def pMCE(self):
        """Gets the Gets the multi-configuration editor interface (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.MCE)
    
    @property
    def pMFE(self):
        """Gets the Gets the multi-function editor interface (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.MFE)

    @property
    def pNCE(self):
        """Gets the Gets the Non-Sequential Component editor interface (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.NCE)
    
    @property
    def pSystemData(self):
        """Gets the System Explorer interface (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.SystemData)

    @property
    def pTDE(self):
        """Gets the tolerance data editor interface (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.TDE)

    @property
    def pTheApplication(self):
        """The ZOSAPI_Application (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.TheApplication)
    
    @property
    def pTools(self):
        """Gets an interface used to run various tools on the optical system (wrapped)"""
        return wrapped_zos_object(self._iopticalsystem.Tools)
    
    # Extra added helper methods 
    def zDummyOpticalSystemsMethod(self):
        #TODO: remove this method in production code
        print('Dummy Optical System method!')