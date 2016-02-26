..  image:: https://raw.githubusercontent.com/pyzos/pyzos/master/Doc/images/pyzos_banner_small.png

Python Zemax OpticStudio API 
----------------------------

Current revision
''''''''''''''''
0.0.8 (Last significant update on 02/26/2016) 

Philosophy / Design Goals
'''''''''''''''''''''''''

Problems
~~~~~~~~

The ZOS-API is an excellent interface for OpticStudio. However, using the ZOS COM API in 
Python directly through PyWin32 is not conducive and feels very unpythonic for the following
reasons: 

* the large set of *property* attributes of the ZOS objects are not introspectable, 
* several interface objects require appropriate type casting before use, and 
* the interface is quite complex (albeit flexible) requiring a significant amount of coding.
* ZOS-API only works in standalone (headless) mode. This prevents one to interact with a 
  running OpticStudio user-interface and observe changes made to the design instantly.   

Solutions
~~~~~~~~~

The philosophy behind PyZOS is to make ZOS-API easier to use in Python by:

* enabling interactivity with a running OpticStudio user-interface (`see demo <https://www.youtube.com/watch?v=ot5CrjMXc_w>`__)
* providing better introspection of objects' properties and methods 
* reducing complexity by

  - providing a set of custom helper-methods that accomplishes common tasks in single or fewest possible steps
  - providing a framework for easily adding custom methods that seamlessly couple with existing ZOS objects
  - managing appropriate type casting of ZOS objects so that we can concentrate on solving optical design problem

* do all the above without limiting or obscuring the ZOS-API in any way
* do all the above with as minimum coding as possible (i.e. do a lot by doing very little)

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
 
See more example use of PyZOS used within Jupyter notebooks `here <https://github.com/pyzos/pyzos/tree/master/Examples/jupyter_notebooks>`__.


Install PyZOS from PyPI
''''''''''''''''''''''''

Use the following command from the command line to install PyZOS from PyPI:

.. code:: python

  pip install pyzos



Interested in contributing?
'''''''''''''''''''''''''''
You can contribute in may ways:

1. using the library and giving feedback, suggestions and reporting bugs 
2. adding custom functions you wrote for your project that others can use
3. helping to write the unit-test, adding test cases
4. letting others know about PyZOS (if you found it useful)


Dependencies
''''''''''''

The core PyZOS library only depends on the standard Python Library. 

1. Python 3.3 and above / Python 2.7; 32/64 bit version
2. `PyWin32 <http://sourceforge.net/projects/pywin32/>`__

All the dependencies can be installed by using the Anaconda Python distribution.

License
'''''''

The code is under the `MIT License <http://opensource.org/licenses/MIT>`__.
