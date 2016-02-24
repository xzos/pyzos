# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        zosutils.py
# Purpose:     Utilities for pyzos
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
from __future__ import division, print_function
import sys as _sys
import warnings as _warnings
from win32com.client import CastTo as _CastTo

#%% Module Global variables
_NO_MODULE_WARNING = False   # Tempory no-module warning (for development)

def get_callable_method_dict(obj):
    """Returns a dictionary of callable methods of object `obj`.

    @param obj: ZOS API Python COM object
    @return: a dictionary of callable methods
    
    Notes: 
    the function only returns the callable attributes that are listed by dir() 
    function. Properties are not returned.
    """
    methodDict = {}
    for methodStr in dir(obj):
        method = getattr(obj, methodStr, 'none')
        if callable(method) and not methodStr.startswith('_'):
            methodDict[methodStr] = method
    return methodDict

def replicate_methods(srcObj, dstObj):
    """Replicate callable methods from a `srcObj` to `dstObj` (generally a wrapper object). 
    
    @param srcObj: source object
    @param dstObj: destination object of the same type.
    @return : none
    
    Implementer notes: 
    1. Once the methods are mapped from the `srcObj` to the `dstObj`, the method calls will 
       not get "routed" through `__getattr__` method (if implemented) in `type(dstObj)` class.
    2. An example of what a 'key' and 'value' look like:
       key: MakeSequential
       value: <bound method IOpticalSystem.MakeSequential of 
              <win32com.gen_py.ZOSAPI_Interfaces.IOpticalSystem instance at 0x77183968>>
    """
    # prevent methods that we intend to specialize from being mapped. The specialized 
    # (overridden) methods are methods with the same name as the corresponding method in 
    # the source ZOS API COM object written for each ZOS API COM object in an associated 
    # python script such as i_analyses_methods.py for I_Analyses
    overridden_methods = get_callable_method_dict(type(dstObj)).keys()
    #overridden_attrs = [each for each in type(dstObj).__dict__.keys() if not each.startswith('_')]
    #print('overridden_methods:', overridden_methods)
    for key, value in get_callable_method_dict(srcObj).items():
        if key not in overridden_methods:
            #setattr(dstObj, key, value)
            print('\n>> Replicating method:')
            print('key:', key)
            print('value:', value)
            setattr(dstObj, key, wrapped_zos_object(value))
        
def get_properties(zos_obj):
    """Returns a lists of properties bound to the object `zos_obj`

    @param zos_obj: ZOS API Python COM object
    @return prop_get: list of properties that are only getters
    @return prop_set: list of properties that are both getters and setters
    """
    prop_get = set(zos_obj._prop_map_get_.keys())
    prop_set = set(zos_obj._prop_map_put_.keys())
    if prop_set.issubset(prop_get):
        prop_get = prop_get.difference(prop_set)
    else:
        msg = 'Assumption all getters are also setters is incorrect!'
        raise NotImplementedError(msg)
    return list(prop_get), list(prop_set)

#%%
class ZOSPropMapper(object):
    """Descriptor for mapping ZOS object properties to corresponding wrapper classes
    """
    def __init__(self, zos_interface_attr, property_name, setter=False, cast_to=None):
        """
        @param zos_interface_attr : attribute used to dispatch method/property calls to 
        the zos_object (it hold the zos_object)
        @param propname : string, like 'SystemName' for IOpticalSystem
        @param setter : if False, a read-only data descriptor is created
        @param cast_to : Name of class (generally the base class) whose property to call
        """
        self.property_name = property_name  # property_name is a string like 'SystemName' for IOpticalSystem
        self.zos_interface_attr = zos_interface_attr  
        self.setter = setter
        self.cast_to = cast_to

    def __get__(self, obj, objtype):
        if self.cast_to:   
            return wrapped_zos_object(getattr(_CastTo(obj.__dict__[self.zos_interface_attr], self.cast_to), self.property_name))
        else:
            return wrapped_zos_object(getattr(obj.__dict__[self.zos_interface_attr], self.property_name))
    
    def __set__(self, obj, value):
        if self.setter:
            if self.cast_to:
                setattr(_CastTo(obj.__dict__[self.zos_interface_attr], self.cast_to), self.property_name, value)
            else:
                setattr(obj.__dict__[self.zos_interface_attr], self.property_name, value)
        else:
            raise AttributeError("Can't set {}".format(self.property_name))
            

def managed_wrapper_class_factory(zos_obj):
    """Creates and returns a wrapper class of a ZOS object, exposing the ZOS objects 
    methods and propertis, and patching custom specialized attributes

    @param zos_obj: ZOS API Python COM object
    """
    cls_name = repr(zos_obj).split()[0].split('.')[-1]  
    dispatch_attr = '_' + cls_name.lower()  # protocol to be followed to store the ZOS COM object
    
    cdict = {}  # class dictionary

    # patch the properties of the base objects 
    base_cls_list = inheritance_dict.get(cls_name, None)
    if base_cls_list:
        for base_cls_name in base_cls_list:
            getters, setters = get_properties(_CastTo(zos_obj, base_cls_name))
            for each in getters:
                exec("p{} = ZOSPropMapper('{}', '{}', cast_to='{}')".format(each, dispatch_attr, each, base_cls_name), globals(), cdict)
            for each in setters:
                exec("p{} = ZOSPropMapper('{}', '{}', setter=True, cast_to='{}')".format(each, dispatch_attr, each, base_cls_name), globals(), cdict)

    # patch the property attributes of the given ZOS object
    getters, setters = get_properties(zos_obj)
    for each in getters:
        exec("p{} = ZOSPropMapper('{}', '{}')".format(each, dispatch_attr, each), globals(), cdict)
    for each in setters:
        exec("p{} = ZOSPropMapper('{}', '{}', setter=True)".format(each, dispatch_attr, each), globals(), cdict)
    
    def __init__(self, zos_obj):
        
        # dispatcher attribute
        cls_name = repr(zos_obj).split()[0].split('.')[-1] 
        dispatch_attr = '_' + cls_name.lower()    # protocol to be followed to store the ZOS COM object
        self.__dict__[dispatch_attr] = zos_obj
        self._dispatch_attr_value = dispatch_attr # used in __getattr__
        
        # Store base class object 
        self._base_cls_list = inheritance_dict.get(cls_name, None)

        # patch the methods of the base class(s) of the given ZOS object
        if self._base_cls_list:
            for base_cls_name in self._base_cls_list:
                replicate_methods(_CastTo(zos_obj, base_cls_name), self)

        # patch the methods of given ZOS object 
        replicate_methods(zos_obj, self)

        # mark object as wrapped to prevent it from being wrapped subsequently
        self._wrapped = True
    
    # Provide a way to make property calls without the prefix p
    def __getattr__(self, attrname):
        return wrapped_zos_object(getattr(self.__dict__[self._dispatch_attr_value], attrname))
        
    cdict['__init__'] = __init__
    cdict['__getattr__'] = __getattr__
    
    # patch custom methods from python files imported as modules
    module_import_str = """
try: 
    from pyzos.zos_obj_override.{module:} import *
except ImportError:
    if _NO_MODULE_WARNING:
        _warnings.warn('No module {module:} found', UserWarning, 2)
""".format(module=cls_name.lower() + '_methods')
    exec(module_import_str, globals(), cdict)

    _ = cdict.pop('print_function', None)
    _ = cdict.pop('division', None)
    
    return type(cls_name, (), cdict) 

def wrapped_zos_object(zos_obj):
    """Helper function to wrap ZOS API COM objects. 

    @param zos_obj : ZOS API Python COM object
    @return: instance of the wrapped ZOS API class. If the input object is not a ZOS-API
             COM object or if it is already wrapped, then the object is returned without
             wrapping.

    Notes:
    The function dynamically creates a wrapped class with all the provided methods, 
    properties, and custom methods monkey patched; and returns an instance of it.
    """
    print('... Wrappped_zos_object called for object:')
    print(repr(zos_obj))
    print("Has CLSID ?", 'CLSID' in dir(zos_obj))
    if hasattr(zos_obj, '_wrapped') or ('CLSID' not in dir(zos_obj)):
        return zos_obj
    else:
        Class = managed_wrapper_class_factory(zos_obj)   
        return Class(zos_obj)

#%% ZOS object inheritance relationships dictionary
# Unfortunately this dict is created manually following the ZOS-API documentation. There
# is no way to know this relationship querying the pythoncom objects.
# Rules (and assumptions made by functions using this dict):
#  1. The base class hierarchy is encoded as lists (i.e. elements are ordered) in the value fields of the dict 
#  2. The dict only contain those ZOS objects that have one or more parent classes. i.e. empty lists are not
#     allowed. 
#  3. The order of super classes in each list: [immediate-base-cls, next-level-base-cls, ..., top-most-base-cls]
inheritance_dict = {
    ## IEditor Interface - base interface for all 5 editors
    'ILensDataEditor' : ['IEditor',],
    'IMultiConfigEditor' : ['IEditor',],
    'IMeritFunctionEditor' : ['IEditor',],
    'INonSeqEditor' : ['IEditor',],
    'IToleranceDataEditor' : ['IEditor',],
    ## IAS_ Interface - base class for all analysis settings interfaces
    # Aberrations interface settings
    'IAS_FieldCurvatureAndDistortion' : ['IAS_',],
    'IAS_FocalShiftDiagram' : ['IAS_',],
    'IAS_GridDistortion' : ['IAS_',],
    'IAS_LateralColor' : ['IAS_',],
    'IAS_LongitudinalAberration' : ['IAS_',],
    'IAS_RayTrace' : ['IAS_',],
    'IAS_SeidelCoefficients' : ['IAS_',],
    'IAS_SeidelDiagram' : ['IAS_',],
    'IAS_ZernikeAnnularCoefficients' : ['IAS_',],
    'IAS_ZernikeCoefficientsVsField' : ['IAS_',],
    'IAS_ZernikeFringeCoefficients' : ['IAS_',],
    'IAS_ZernikeStandardCoefficients' : ['IAS_',],
    # EncircledEnergy interface settings
    'IAS_DiffractionEncircledEnergy' : ['IAS_',],
    'IAS_ExtendedSourceEncircledEnergy' : ['IAS_',],
    'IAS_GeometricEncircledEnergy' : ['IAS_',],
    'IAS_GeometricLineEdgeSpread' : ['IAS_',],
    # Fans interface settings
    'IAS_Fan' : ['IAS_',],
    # Mtf interface settings
    'IAS_FftMtf' : ['IAS_',],
    'IAS_FftMtfMap' : ['IAS_',],
    'IAS_FftMtfvsField' : ['IAS_',],
    'IAS_FftSurfaceMtf' : ['IAS_',],
    'IAS_FftThroughFocusMtf' : ['IAS_',],
    'IAS_GeometricMtf' : ['IAS_',],
    'IAS_GeometricMtfMap' : ['IAS_',],
    'IAS_GeometricMtfvsField' : ['IAS_',],
    'IAS_GeometricThroughFocusMtf' : ['IAS_',],
    'IAS_HuygensMtf' : ['IAS_',],
    'IAS_HuygensMtfvsField' : ['IAS_',],
    'IAS_HuygensSurfaceMtf' : ['IAS_',],
    'IAS_HuygensThroughFocusMtf' : ['IAS_',],
    # Psf interface settings
    'IAS_FftPsf' : ['IAS_',],
    'IAS_FftPsfCrossSection' : ['IAS_',],
    'IAS_FftPsfLineEdgeSpread' : ['IAS_',],
    'IAS_HuygensPsf' : ['IAS_',],
    'IAS_HuygensPsfCrossSection' : ['IAS_',],
    # RayTracing interface settings 
    'IAS_DetectorViewer' : ['IAS_',],
    # RMS interface settings
    'IAS_RMSField' : ['IAS_',],
    'IAS_RMSFieldMap' : ['IAS_',],
    'IAS_RMSFocus' : ['IAS_',],
    'IAS_RMSLambdaDiagram' : ['IAS_',],
    # Spot interface settings
    'IAS_Spot' : ['IAS_',],
    # Surface interface settings
    'IAS_SurfaceCurvature' : ['IAS_',],
    'IAS_SurfaceCurvatureCross' : ['IAS_',],
    'IAS_SurfacePhase' : ['IAS_',],
    'IAS_SurfacePhaseCross' : ['IAS_',],
    'IAS_SurfaceSag' : ['IAS_',],
    'IAS_SurfaceSagCross' : ['IAS_',],
    # Wavefront interface settings
    'IAS_Foucault' : ['IAS_',],
    ## IOpticalSystemTools Interface - base class for all system tools
    'IBatchRayTrace' : ['ISystemTool',],
    'IConvertToNSCGroup' : ['ISystemTool',],
    'ICreateArchive' : ['ISystemTool',],
    'IExportCAD' : ['ISystemTool',],
    'IGlobalOptimization' : ['ISystemTool',],
    'IHammerOptimization' : ['ISystemTool',],
    'ILensCatalogs' : ['ISystemTool',],
    'ILightningTrace' : ['ISystemTool',],
    'ILocalOptimization' : ['ISystemTool',],
    'IMFCalculator' : ['ISystemTool',],
    'INSCRayTrace' : ['ISystemTool',],
    'IQuickAdjust' : ['ISystemTool',],
    'IQuickFocus' : ['ISystemTool',],
    'IRestoreArchive' : ['ISystemTool',],
    'IScale' : ['ISystemTool',],
    'ITolerancing' : ['ISystemTool',],    
    ## IWizard Interface - base interface for all wizards
    'INSCWizard' : ['IWizard',],
    'INSCBitmapWizard' : ['INSCWizard', 'IWizard',],   
    'INSCOptimizationWizard' : ['INSCWizard', 'IWizard',], 
    'INSCRoadwayLightingWizard' : ['IWizard',],
    'IToleranceWizard' : ['IWizard',], 
    'INSCToleranceWizard': ['IToleranceWizard', 'IWizard',], 
    'ISEQToleranceWizard' : ['IToleranceWizard', 'IWizard',], 
    'ISEQOptimizationWizard' : ['IWizard',],
}
# Ensure Rule #2 of inheritance_dict.
for each in inheritance_dict.values():
    assert len(each), 'Empty base class list not allowed in inheritance_dict'