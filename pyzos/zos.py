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
import os as _os
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
    

    #%% Overridden Methods
    def SaveAs(self, filename):
        """Saves the current system to the specified file. 

        @param filename: absolute path (string)
        @return: None
        @raise: ValueError if path (excluding the zemax file name) is not valid

        All future calls to `Save()`  will use the same file.
        """
        directory = _os.path.split(filename)[0]
        if not _os.path.exists(directory):
            raise ValueError('{} is not valid.'.format(directory))
        else:
            self._iopticalsystem.SaveAs(filename)

    #%% Overridden Properties
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
    def zInsertNewSurfaceAt(self, surfNum):
        if self.pMode == 0:
            lde = self.pLDE
            lde.InsertNewSurfaceAt(surfNum)
        else:
            raise NotImplementedError('Function not implemented for non-sequential mode')

    def zSetSurfaceData(self, surfNum, radius=None, thick=None, material=None, semidia=None, conic=None, comment=None):
        if self.pMode == 0: # sequential mode
            surf = self.pLDE.GetSurfaceAt(surfNum)
            if radius is not None:
                surf.pRadius = radius
            if thick is not None:
                surf.pThickness = thick
            if material is not None:
                surf.pMaterial = material
            if semidia is not None:
                surf.pSemiDiameter = semidia
            if conic is not None:
                surf.pConic = conic
            if comment is not None:
                surf.pComment = comment
        else:
            raise NotImplementedError('Function not implemented for non-sequential mode')

    def zSetDefaultMeritFunctionSEQ(self, ofType=0, ofData=0, ofRef=0, pupilInteg=0, rings=0,
                                    arms=0, obscuration=0, grid=0, delVignetted=False, useGlass=False, 
                                    glassMin=0, glassMax=1000, glassEdge=0, useAir=False, airMin=0, 
                                    airMax=1000, airEdge=0, axialSymm=True, ignoreLatCol=False, 
                                    addFavOper=False, startAt=1, relativeXWgt=1.0, overallWgt=1.0, 
                                    configNum=0):
        """Sets the default merit function for Sequential Merit Function Editor

        Parameters
        ----------
        ofType : integer
            optimization function type (0=RMS, ...)
        ofData : integer 
            optimization function data (0=Wavefront, 1=Spot Radius, ...)
        ofRef : integer
            optimization function reference (0=Centroid, ...)
        pupilInteg : integer
            pupil integration method (0=Gaussian Quadrature, 1=Rectangular Array)
        rings : integer
            rings (0=1, 1=2, 2=3, 3=4, ...)
        arms : integer 
            arms (0=6, 1=8, 2=10, 3=12)
        obscuration : real
            obscuration
        delVignetted : boolean 
            delete vignetted ?
        useGlass : boolean 
            whether to use Glass settings for thickness boundary
        glassMin : real
            glass mininum thickness 
        glassMax : real 
            glass maximum thickness
        glassEdge : real
            glass edge thickness
        useAir : boolean
            whether to use Air settings for thickness boundary
        airMin : real
            air minimum thickness      
        airMax : real 
            air maximum thickness
        airEdge : real 
            air edge thickness
        axialSymm : boolean 
            assume axial symmetry 
        ignoreLatCol : boolean
            ignore latent color
        addFavOper : boolean
            add favorite color
        configNum : integer
            configuration number (0=All)
        startAt : integer 
            start at
        relativeXWgt : real 
            relative X weight
        overallWgt : real
            overall weight
        """
        mfe = self.pMFE
        wizard = mfe.pSEQOptimizationWizard
        wizard.pType = ofType
        wizard.pData = ofData
        wizard.pReference = ofRef
        wizard.pPupilIntegrationMethod = pupilInteg 
        wizard.pRing = rings
        wizard.pArm = arms
        wizard.pObscuration = obscuration
        wizard.pGrid = grid
        wizard.pIsDeleteVignetteUsed =  delVignetted
        wizard.pIsGlassUsed = useGlass 
        wizard.pGlassMin = glassMin
        wizard.pGlassMax = glassMax
        wizard.pGlassEdge = glassEdge
        wizard.pIsAirUsed = useAir
        wizard.pAirMin = airMin
        wizard.pAirMax = airMax 
        wizard.pAirEdge = airEdge 
        wizard.pIsAssumeAxialSymmetryUsed = axialSymm
        wizard.pIsIgnoreLateralColorUsed = ignoreLatCol
        wizard.pConfiguration = configNum 
        wizard.pIsAddFavoriteOperandsUsed = addFavOper
        wizard.pStartAt = startAt
        wizard.pRelativeXWeight = relativeXWgt
        wizard.pOverallWeight = overallWgt
        wizard.CommonSettings.OK() # Settings are set, perform the wizardry. 
