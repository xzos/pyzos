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
import sys as _sys
import collections as _co
import win32com.client as _comclient
import tempfile as _tempfile
import time as _time
import warnings as _warnings
from pyzos.zosutils import (ZOSPropMapper as _ZOSPropMapper, 
                            replicate_methods as _replicate_methods,
                            inheritance_dict as _inheritance_dict,
                            wrapped_zos_object as wrapped_zos_object)
import pyzos.ddeclient as _dde


#%% Global variables
Const = None  # Constants (placeholder)

#%% Custom Exceptions
class _ConnectionError(Exception): pass 
class _InitializationError(Exception): pass
class _ZOSSystemError(Exception): pass


#%% Module helper functions
def _get_python_version():
    return _sys.version_info[0]
    
def _get_constants_dict():
    const_dict = _comclient.constants.__dicts__[0]
    if _get_python_version() == 2:
        return const_dict
    else:
        return dict(const_dict)
    
def _get_sync_ui_filename():
    temp_dir = _tempfile.gettempdir()
    temp_file = 'pyzos_ui_sync_file_{}.zmx'.format(_os.getpid())
    return _os.path.join(temp_dir, temp_file)

def _get_new_dde_link():
    ln = _PyZDDE()
    ln.zDDEInit()
    return ln
    
def _delete_file(fileName, n=10):
    """Cleanly deletes a file in `n` attempts (if necessary)"""
    status = False
    count = 0
    while not status and count < n:
        try:
            _os.remove(fileName)
        except OSError:
            count += 1
            _time.sleep(0.2)
        else:
            status = True
    return status


#%% _PyZDDE class (stripped down)
class _PyZDDE(object):
    """Class for communicating with Zemax using DDE"""
    chNum = -1   
    liveCh = 0  
    server = 0    
    
    def __init__(self):
        _PyZDDE.chNum += 1   
        self.appName = "ZEMAX" + str(_PyZDDE.chNum) if _PyZDDE.chNum > 0 else "ZEMAX"
        self.connection = False  

    def zDDEInit(self):
        """Initiates link with OpticStudio DDE server"""
        self.pyver = _get_python_version()
        # do this only one time or when there is no channel
        if _PyZDDE.liveCh==0:
            try:
                _PyZDDE.server = _dde.CreateServer()
                _PyZDDE.server.Create("ZCLIENT")   
            except Exception as err:
                _sys.stderr.write("{}: DDE server may be in use!".format(str(err)))
                return -1
        # Try to create individual conversations for each ZEMAX application.
        self.conversation = _dde.CreateConversation(_PyZDDE.server)
        try:
            self.conversation.ConnectTo(self.appName, " ")
        except Exception as err:
            _sys.stderr.write("{}.\nOpticStudio UI may not be running!\n".format(str(err)))
            # should close the DDE server if it exist
            self.zDDEClose()
            return -1
        else:
            _PyZDDE.liveCh += 1 
            self.connection = True
            return 0

    def zDDEClose(self):
        """Close the DDE link with Zemax server"""
        if _PyZDDE.server and not _PyZDDE.liveCh:
            _PyZDDE.server.Shutdown(self.conversation)
            _PyZDDE.server = 0
        elif _PyZDDE.server and self.connection and _PyZDDE.liveCh == 1:
            _PyZDDE.server.Shutdown(self.conversation)
            self.connection = False
            self.appName = ''
            _PyZDDE.liveCh -= 1  
            _PyZDDE.server = 0  
        elif self.connection:  
            _PyZDDE.server.Shutdown(self.conversation)
            self.connection = False
            self.appName = ''
            _PyZDDE.liveCh -= 1
        return 0

    def setTimeout(self, time):
        """Set global timeout value, in seconds, for all DDE calls"""
        self.conversation.SetDDETimeout(round(time))
        return self.conversation.GetDDETimeout()

    def _sendDDEcommand(self, cmd, timeout=None):
        """Send command to DDE client"""
        reply = self.conversation.Request(cmd, timeout)
        if self.pyver > 2:
            reply = reply.decode('ascii').rstrip()
        return reply

    def __del__(self):
        self.zDDEClose()
        
    def zGetFile(self):
        """Returns the full name of the lens file in DDE server"""
        reply = self._sendDDEcommand('GetFile')
        return reply.rstrip()
               
    def zGetRefresh(self):
        """Copy lens data from the LDE into the Zemax DDE server"""
        reply = None
        reply = self._sendDDEcommand('GetRefresh')
        if reply:
            return int(reply) #Note: Zemax returns -1 if GetRefresh fails.
        else:
            return -998

    def zGetUpdate(self):
        """Update the lens"""
        status,ret = -998, None
        ret = self._sendDDEcommand("GetUpdate")
        if ret != None:
            status = int(ret)  #Note: Zemax returns -1 if GetUpdate fails.
        return status
            
    def zGetVersion(self):
        """Get the version of Zemax """
        return int(self._sendDDEcommand("GetVersion"))
        
    def zLoadFile(self, fileName, append=None):
        """Loads a zmx file into the DDE server"""
        reply = None
        if append:
            cmd = "LoadFile,{},{}".format(fileName, append)
        else:
            cmd = "LoadFile,{}".format(fileName)
        reply = self._sendDDEcommand(cmd)
        if reply:
            return int(reply) #Note: Zemax returns -999 if update fails.
        else:
            return -998
        
    def zPushLens(self, update=None, timeout=None):
        """Copy lens in the Zemax DDE server into LDE"""
        reply = None
        if update == 1:
            reply = self._sendDDEcommand('PushLens,1', timeout)
        elif update == 0 or update is None:
            reply = self._sendDDEcommand('PushLens,0', timeout)
        else:
            raise ValueError('Invalid value for flag')
        if reply:
            return int(reply)   # Note: Zemax returns -999 if push lens fails
        else:
            return -998

    def zPushLensPermission(self):
        status = None
        status = self._sendDDEcommand('PushLensPermission')
        return int(status)

    def zSaveFile(self, fileName):
        """Saves the lens currently loaded in the server to a Zemax file """
        cmd = "SaveFile,{}".format(fileName)
        reply = self._sendDDEcommand(cmd)
        return int(float(reply.rstrip()))


#%% ZOS API Application Class
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
            Const = type('Const', (), _get_constants_dict()) # Constants class
        return cls.app

#%% Optical System Class
class OpticalSystem(object):
    """Wrapper for IOpticalSystem interface
    """
    _instantiated = False
    _pyzosapp = None
    _dde_link = None

    # Patch managed properties of IOpticalSystem's base classes
    # skipped for now ... IOpticalSystem doesn't have any base class (currently)
        
    # Patch managed properties of ZOS IOpticalSystem
    pIsNonAxial = _ZOSPropMapper('_iopticalsystem', 'IsNonAxial')
    pMode = _ZOSPropMapper('_iopticalsystem', 'Mode')
    pNeedsSave = _ZOSPropMapper('_iopticalsystem', 'NeedsSave')
    pSystemFile = _ZOSPropMapper('_iopticalsystem', 'SystemFile')
    pSystemID = _ZOSPropMapper('_iopticalsystem', 'SystemID')
    pSystemName = _ZOSPropMapper('_iopticalsystem', 'SystemName', setter=True)
    
    def __init__(self, sync_ui=False, mode=0):
        """Returns an instance of Optical System
        @param mode : sequential (0) or non-sequential (1)
        """
        if OpticalSystem._instantiated:
            self._iopticalsystem = OpticalSystem._pyzosapp.CreateNewSystem(mode) # wrapped object
        else:
            OpticalSystem._pyzosapp = _PyZOSApp()                         # wrapped object
            self._iopticalsystem = OpticalSystem._pyzosapp.GetSystemAt(0) # PrimarySystem
            if mode == 1:
                self._iopticalsystem.MakeNonSequential()
            OpticalSystem._instantiated = True

        ## Store ZOS IOpticalSystem's base class(es)
        self._base_cls_list = _inheritance_dict.get('IOpticalSystem', None)
            
        ## activate PyZDDE if sync_ui requested
        if sync_ui and not OpticalSystem._dde_link:
            OpticalSystem._dde_link = _get_new_dde_link()
        self._sync_ui_file = _get_sync_ui_filename() if sync_ui else None
        self._file_to_save_on_Save = None
        
        ## patch methods from base class of IOpticalSystem to the wrapped object
        if self._base_cls_list:
            for base_cls_name in self._base_cls_list:
                _replicate_methods(_comclient.CastTo(self._iopticalsystem, base_cls_name), self)

        ## patch methods from ZOS IOpticalSystem to the wrapped object
        _replicate_methods(self._iopticalsystem, self)

    # Provide a way to make property calls without the prefix p, but don't try to wrap the returned object            
    def __getattr__(self, attrname):
        return getattr(self._iopticalsystem, attrname)
    
    def __del__(self):
        if self._sync_ui_file:
            ext_dict = ['.zmx', '.ZMX', '.CFG', '.SES', '.ZDA']
            filename_bar_ext = self._sync_ui_file.rsplit('.')[0]
            for ext in ext_dict:
                if _os.path.exists(filename_bar_ext + ext):
                    _delete_file(filename_bar_ext + ext)
        if OpticalSystem._dde_link:
            OpticalSystem._dde_link.zDDEClose()  ##TODO: FIX should probably have a reference count???
        
    #%% UI sync machinery
    def zPushLens(self, update=None):
        """Push lens in ZOS COM server to UI"""
        self.SaveAs(self._sync_ui_file)
        OpticalSystem._dde_link.zLoadFile(self._sync_ui_file)
        OpticalSystem._dde_link.zPushLens(update)
        
    def zGetRefresh(self):
        """Copy lens in UI to headless ZOS COM server"""
        OpticalSystem._dde_link.zGetRefresh()
        OpticalSystem._dde_link.zSaveFile(self._sync_ui_file)
        self._iopticalsystem.LoadFile (self._sync_ui_file, False)
    
    #%% Overridden Methods
    def SaveAs(self, filename):
        """Saves the current system to the specified file. 

        @param filename: absolute path (string)
        @return: None
        @raise: ValueError if path (excluding the zemax file name) is not valid

        All future calls to `Save()`  will use the same file.
        """
        directory, zfile = _os.path.split(filename)
        if zfile.startswith('pyzos_ui_sync_file'):
            self._iopticalsystem.SaveAs(filename)
        else: # regular file
            if not _os.path.exists(directory):
                raise ValueError('{} is not valid.'.format(directory))
            else:
                self._file_to_save_on_Save = filename   # store to use in Save()
                self._iopticalsystem.SaveAs(filename)
            
    def Save(self):
        """Saves the current system"""
        # This method is intercepted to allow ui_sync
        if self._file_to_save_on_Save:
            self._iopticalsystem.SaveAs(self._file_to_save_on_Save)
        else:
            self._iopticalsystem.Save()

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
    
    #%% Extra / Custom methods 
    def zGetSurfaceData(self, surfNum):
        """Return surface data"""
        if self.pMode == 0: # Sequential mode
            surf_data = _co.namedtuple('surface_data', ['radius', 'thick', 'material', 'semidia', 
                                                        'conic', 'comment'])
            surf = self.pLDE.GetSurfaceAt(surfNum)
            return surf_data(surf.pRadius, surf.pThickness, surf.pMaterial, surf.pSemiDiameter,
                             surf.pConic, surf.pComment)
        else:
            raise NotImplementedError('Function not implemented for non-sequential mode')

    def zInsertNewSurfaceAt(self, surfNum):
        if self.pMode == 0:
            lde = self.pLDE
            lde.InsertNewSurfaceAt(surfNum)
        else:
            raise NotImplementedError('Function not implemented for non-sequential mode')

    def zSetSurfaceData(self, surfNum, radius=None, thick=None, material=None, semidia=None, 
                        conic=None, comment=None):
        """Sets surface data"""
        if self.pMode == 0: # Sequential mode
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
