# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        ilensdataeditor_methods.py
# Purpose:     store custom methods for wrapper class of ILensDataEditor
# Licence:     MIT License
#-------------------------------------------------------------------------------
"""Store custom methods for wrapper class of ILensDataEditor, which defines 
   all properties and methods needed to interact with the Lens Data Editor. 
   This interface can be accessed via the IOpticalSystem interface. 
   name := repr(zos_obj).split()[0].split('.')[-1].lower() + '_methods.py' 
"""
from __future__ import print_function
from __future__ import division
import collections as _co
from win32com.client import CastTo as _CastTo, constants as _constants
from pyzos.zosutils import wrapped_zos_object as _wrapped_zos_object

# Overridden methods
# ------------------
def AddRow(self):
    """Adds a new row at the end of the editor."""
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return _wrapped_zos_object(base_lde.AddRow())


def DeleteAllRows(self):
    """Deletes all rows from the current editor"""
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return base_lde.DeleteAllRows()


def DeleteRowAt(self, pos):
    """Deletes a row at the specified index (0 to NumberOfRows-1)
    @pos : The index of the first row to remove
    """
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return base_lde.DeleteRowAt(pos)

def DeleteRowsAt(self, pos, numOfRows):
    """Deletes one or more rows at the specified index (0 to NumberOfRows-1)
    @pos : The index of the first row to remove
    @numOfRows : The number of rows to remove
    """
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return base_lde.DeleteRowsAt(pos, numOfRows)

def GetPupil(self):
    """Retrieve pupil data
    """
    pupil_data = _co.namedtuple('pupil_data', ['ZemaxApertureType',
                                               'ApertureValue',
                                               'entrancePupilDiameter',
                                               'entrancePupilPosition',
                                               'exitPupilDiameter',
                                               'exitPupilPosition',
                                               'ApodizationType',
                                               'ApodizationFactor'])
    data = self._ilensdataeditor.GetPupil()
    return pupil_data(*data)

def GetRowAt(self, pos):
    """Gets the row at the specified index (0 to NumberOfRows-1).
    @pos : The row index.
    """
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return _wrapped_zos_object(base_lde.GetRowAt(pos))

def GetSurfaceAt(self, surfNum):
    """Gets the data for the specified surface."""
    return _wrapped_zos_object(self._ilensdataeditor.GetSurfaceAt(surfNum))


# Overridden properties
# ---------------------
@property
def pEditor(self):
    """Gets the type of editor instance"""
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return base_lde.Editor

@property
def pMaxColumn(self):
    """The maximum column index that can be used in calls to GetCellAt()"""
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return base_lde.MaxColumn

@property
def pMinColumn(self):
    """The minimum column index that can be used in calls to GetCellAt()"""
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return base_lde.MinColumn

@property
def pNumberOfRows(self):
    """Gets the number of rows in this editor"""
    base_lde = _CastTo(self._ilensdataeditor, 'IEditor')
    return base_lde.NumberOfRows


# Extra methods
# -------------
