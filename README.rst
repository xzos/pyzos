..  image:: https://raw.githubusercontent.com/pyzos/pyzos/master/Doc/images/pyzos_banner_small.png

Python Zemax OpticStudio API 
----------------------------

Current revision
''''''''''''''''
The current code (version 0.0.2) is a working PROTOTYPE! 

Philosophy / Design Goals
'''''''''''''''''''''''''

Problems
~~~~~~~~

The ZOS-API is an excellent interface for OpticStudio. However, using the ZOS COM API in 
Python through PyWin32 creates some problems: 

* the large set of *property* attributes of the ZOS objects are not introspectable, 
* several interface objects require appropriate type casting before use, and 
* the interface is quite complex (albeit flexible) requiring a significant amount of coding.
* ZOS-API only works in standalone (headless) mode. This prevents one to interact with a 
  running OpticStudio user-interface and observe changes made to the design instantly.   

Solutions
~~~~~~~~~

The philosophy behind PyZOS is to make ZOS-API easier to use in Python by:

1. enabling interactivity with a running OpticStudio user-interface (`see demo <https://www.youtube.com/watch?v=ot5CrjMXc_w>`__)
2. providing better introspection of objects  
3. reducing complexity by
  * providing a set of helper methods that encapsulates common tasks
  * allowing helper methods to be easily coupled to existing ZOS objects for custom functions
  * managing appropriate type casting of ZOS objects
4. do all the above without limiting or obscuring the ZOS-API in any way. 

These *enhancements* to ZOS-API using PyZOS library are documented in this (work in progress) 
`Jupyter notebook <http://nbviewer.jupyter.org/github/pyzos/pyzos/blob/master/Examples/jupyter_notebooks/00_Enhancing_the_ZOS_API_Interface.ipynb>`__.   


Example usage
'''''''''''''    
.. code:: python

    import pyzos.zos as zos   
    osys = zos.OpticalSystem(sync_ui=True) # enable interactivity with a running UI
    sdata = osys.pSystemData
    sdata.pAperture.pApertureValue = 40
    sdata.pFields.AddField(0, 2.0, 1.0)
    wave = zos.Const.WavelengthPreset_d_0p587
    sdata.pWavelengths.SelectWavelengthPreset(wave)
    ...
    osys.zPushLens(1)  # copy lens from ZOS COM server to the visible UI app
    ...
    osys.zGetRefresh() # copy changes from the visible UI app to the ZOS COM server
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


