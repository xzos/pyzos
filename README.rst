..  image:: https://raw.githubusercontent.com/pyzos/pyzos/master/Doc/images/pyzos_banner_small.png

Python Zemax OpticStudio API 
----------------------------

Current revision
''''''''''''''''
The current code base (version 0.0.2) is a working PROTOTYPE! 

Philosophy / Design Goals
'''''''''''''''''''''''''
The ZOS-API Interface is already an excellent tool. However, interfacing the ZOS-API with 
Python using PyWin32 creates some problems. For e.g., the large set of *property* attributes 
are not introspectable, several objects require appropriate type casting before use, and 
the interface is quite complex (albeit flexible) that require a significant amount of 
coding to do even very simple tasks. 

The philosophy behind PyZOS is to make ZOS-API easier to use with Python by:

1. provide functions to interact with a live OpticStuido UI
2. providing better introspection  
3. reduce complexity by
  * providing a set of helper methods that encapsulate common tasks
  * allowing helper methods to be easily added to ZOS objects for custom tasks
  * taking care of type casting of ZOS objects
4. do all the above without limiting or obscuring the ZOS-API in any way. 

These enhancements to ZOS-API using PyZOS library are documented in this (work in progress) 
`Jupyter notebook <http://nbviewer.jupyter.org/github/pyzos/pyzos/blob/master/Examples/jupyter_notebooks/00_Enhancing_the_ZOS_API_Interface.ipynb>`__.   


Example usage
'''''''''''''    
.. code:: python

    import pyzos.zos as zos   
    osys = zos.OpticalSystem()
    sdata = osys.pSystemData
    sdata.pAperture.pApertureValue = 40
    sdata.pFields.AddField(0, 2.0, 1.0)
    wave = zos.Const.WavelengthPreset_d_0p587
    sdata.pWavelengths.SelectWavelengthPreset(wave)
    ...
 


Dependencies
''''''''''''

The core PyZOS library only depends on the standard Python Library. 

1. Python 3.3 and above / Python 2.7; 32/64 bit version
2. `PyWin32 <http://sourceforge.net/projects/pywin32/>`__

All the dependencies can be installed by using the Anaconda Python distribution.

License
'''''''

The code is under the `MIT License <http://opensource.org/licenses/MIT>`__.


