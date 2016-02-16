
Python ZOS API 
----------------

Current revision
''''''''''''''''
The current code base (version 0.0.1) is a working PROTOTYPE! 

Philosophy / Design Goals
'''''''''''''''''''''''''
The ZOS API Interface is already excellent tool. However, using the ZOS API in with Python
is not as easy. The philosophy of PyZOS is to ease the use ZOS API with Python and at the 
same time without limiting or obscuring the ZOS API in any way. In addition, PyZOS aims to
provide a framework that is easily extendable. 



Example usage
'''''''''''''    
.. code:: python

    from pyzos import zos    
    optical_system = zos.OptSys()
             
 


Dependencies
''''''''''''

The core PyZOS library only depends on the standard Python Library. 

1. Python 2.7 / Python 3.3 and above; 32/64 bit version
2. [PyWin32](http://sourceforge.net/projects/pywin32/)

All the dependencies can be installed by using the Anaconda Python distribution.

License
'''''''

The code is under the `MIT License <http://opensource.org/licenses/MIT>`__.


