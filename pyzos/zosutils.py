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
    # prevent methods that we intend to specialize from being mapped.
    # the specialized (overridden) methods are methods with the same
    # name as the corresponding method in the source ZOS API COM object
    # written for each ZOS API COM object in an associated python script
    # such as i_analyses_methods.py for I_Analyses
    overridden_methods = get_callable_method_dict(type(dstObj)).keys()
    #overridden_attrs = [each for each in type(dstObj).__dict__.keys() if not each.startswith('_')]
    #print('overridden_methods:', overridden_methods)
    for key, value in get_callable_method_dict(srcObj).items():
        if key not in overridden_methods:
            setattr(dstObj, key, value)
        
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
    """Descriptor for mapping ZOS object getter and setter properties to 
    corresponding wrapper classes
    """
    def __init__(self, zos_interface_attr, property_name, setter=False):
        """
        @zos_interface_attr is the attribute used to dispatch method/property
         calls to the zos_object (it hold the zos_object)
        @propname is a string like 'SystemName' for IOpticalSystem
        if `setter` is False, a read-only data descriptor is created
        """
        self.property_name = property_name  # property_name is a string like 'SystemName' 
                                            # for IOpticalSystem
        self.zos_interface_attr = zos_interface_attr  # TODO!! This one probably needs to  
                                        # be a class attribute with weakrefkey dict
                                        # otherwise a second system may interfere
        self.setter = setter

    def __get__(self, obj, objtype):
        return getattr(obj.__dict__[self.zos_interface_attr], self.property_name)
        
    def __set__(self, obj, value):
        if self.setter:
            setattr(obj.__dict__[self.zos_interface_attr], self.property_name, value)
        else:
            raise AttributeError("Can't set {}".format(self.property_name))
            

#TODO## Review the class several times to ensure that there is nothing extra bit 
# of code (especially vestigial code) lingering around.

def managed_wrapper_class_factory(zos_obj):
    """Creates and returns a wrapper class of a ZOS object, exposing the ZOS objects 
    methods and propertis, and patching custom specialized attributes

    @param zos_obj: ZOS API Python COM object
    """
    cls_name = repr(zos_obj).split()[0].split('.')[-1]  # protocol to be followed
    dispatch_attr = '_' + cls_name.lower()   # protocol to be followed
    
    cdict = {}  # class dictionary
    
    # patch the property attributes
    getters, setters = get_properties(zos_obj)
    for each in getters:
        exec("p{} = ZOSPropMapper('{}', '{}')".format(each, dispatch_attr, each), globals(), cdict)
    for each in setters:
        exec("p{} = ZOSPropMapper('{}', '{}', True)".format(each, dispatch_attr, each), globals(), cdict)
    
    def __init__(self, zos_obj):
        
        # dispatcher attribute
        cls_name = repr(zos_obj).split()[0].split('.')[-1] 
        dispatch_attr = '_' + cls_name.lower() # protocol to be followed
        self.__dict__[dispatch_attr] = zos_obj
        self._dispatch_attr_value = dispatch_attr 
        
        # patch the methods
        replicate_methods(zos_obj, self)
        
    def __getattr__(self, attrname):
        #print('__getattr__ in managed instance called for', attrname)
        return getattr(self.__dict__[self._dispatch_attr_value], attrname)
    
    
    cdict['__init__'] = __init__
    cdict['__getattr__'] = __getattr__
    
    # patch custom methods from python files imported as modules
    module_import_str = """
try: 
    from pyzos.zos_obj_override.{module:} import *
except ImportError: 
    _warnings.warn('No module {module:} found', UserWarning, 2)
""".format(module=cls_name.lower() + '_methods')
    exec(module_import_str, globals(), cdict)

    _ = cdict.pop('print_function', None)
    _ = cdict.pop('division', None)
    
    return type(cls_name, (), cdict) 

def wrapped_zos_object(zos_obj):
    """Helper function to wrap ZOS API COM objects. 

    @param zos_obj : ZOS API Python COM object
    @return: instance of the wrapped ZOS API class.

    Notes:
    The function dynamically creates a wrapped class with all the provided methods, 
    properties, and custom methods monkey patched; and returns an instance of it.
    """
    Class = managed_wrapper_class_factory(zos_obj)
    return Class(zos_obj)