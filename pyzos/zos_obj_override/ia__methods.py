# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        ia__methods.py
# Purpose:     store custom methods for wrapper class of IA_ Interface, the Base 
#              interface for all analysis windows.  
# Licence:     MIT License
#-------------------------------------------------------------------------------
"""Store custom methods for wrapper class of IA_ Interface.
   name := repr(zos_obj).split()[0].split('.')[-1].lower() + '_methods.py' 
"""
from __future__ import print_function
from __future__ import division
import warnings as _warnings
from win32com.client import CastTo as _CastTo, constants as _constants
from pyzos.zosutils import (wrapped_zos_object as _wrapped_zos_object,
                            replicate_methods as _replicate_methods)

# Overridden methods
# ------------------

def GetSettings(self):
    # the IA_.GetSettings() returns IAS_ which is the base class of all other settings 
    # objects such as IAS_FftMtf, IAS_FftMap, etc. The base-class IAS_ objects needs to be 
    # "specialized" for a particular analysis using the _CastTo function. When we apply 
    # CastTo(), the specialized objects "gain" the specific analysis functions and properties.
    # before returning, when we invoke _wrapped_zos_object(), the base-class methods and 
    # properties also gets patched. 
    
    settings_base = self._ia_.GetSettings()
    
    # CastTo the specialized class to add properties if they corresponding 
    # interface class is available in the library, else return generic IAS_
    try:
        if self._ia_.AnalysisType == _constants.AnalysisIDM_RayFan:
            settings = _CastTo(settings_base, 'IAS_RayFan')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RayTrace:
            settings = _CastTo(settings_base, 'IAS_RayTrace')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_OpticalPathFan:
            settings = _CastTo(settings_base, 'IAS_OpticalPathFan')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PupilAberrationFan:
            settings = _CastTo(settings_base, 'IAS_PupilAberrationFan')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FieldCurvatureAndDistortion:
            settings = _CastTo(settings_base, 'IAS_FieldCurvatureAndDistortion')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FocalShiftDiagram:
            settings = _CastTo(settings_base, 'IAS_FocalShiftDiagram')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GridDistortion:
            settings = _CastTo(settings_base, 'IAS_GridDistortion')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_LateralColor:
            settings = _CastTo(settings_base, 'IAS_LateralColor')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_LongitudinalAberration:
            settings = _CastTo(settings_base, 'IAS_LongitudinalAberration')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SeidelCoefficients:
            settings = _CastTo(settings_base, 'IAS_SeidelCoefficients')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SeidelDiagram:
            settings = _CastTo(settings_base, 'IAS_SeidelDiagram')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ZernikeAnnularCoefficients:
            settings = _CastTo(settings_base, 'IAS_ZernikeAnnularCoefficients')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ZernikeCoefficientsVsField:
            settings = _CastTo(settings_base, 'IAS_ZernikeCoefficientsVsField')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ZernikeFringeCoefficients:
            settings = _CastTo(settings_base, 'IAS_ZernikeFringeCoefficients')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ZernikeStandardCoefficients:
            settings = _CastTo(settings_base, 'IAS_ZernikeStandardCoefficients')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftMtf:
            settings = _CastTo(settings_base, 'IAS_FftMtf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftMtfMap: 
            settings = _CastTo(settings_base, 'IAS_FftMtfMap')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftMtfvsField: 
            settings = _CastTo(settings_base, 'IAS_FftMtfvsField')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftSurfaceMtf: 
            settings = _CastTo(settings_base, 'IAS_FftSurfaceMtf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftThroughFocusMtf: 
            settings = _CastTo(settings_base, 'IAS_FftThroughFocusMtf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricMtf: 
            settings = _CastTo(settings_base, 'IAS_GeometricMtf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricMtfMap: 
            settings = _CastTo(settings_base, 'IAS_GeometricMtfMap')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricMtfvsField: 
            settings = _CastTo(settings_base, 'IAS_GeometricMtfvsField')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricThroughFocusMtf: 
            settings = _CastTo(settings_base, 'IAS_GeometricThroughFocusMtf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensMtf: 
            settings = _CastTo(settings_base, 'IAS_HuygensMtf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensMtfvsField: 
            settings = _CastTo(settings_base, 'IAS_HuygensMtfvsField')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensSurfaceMtf: 
            settings = _CastTo(settings_base, 'IAS_HuygensSurfaceMtf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensThroughFocusMtf: 
            settings = _CastTo(settings_base, 'IAS_HuygensThroughFocusMtf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftPsf: 
            settings = _CastTo(settings_base, 'IAS_FftPsf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftPsfCrossSection: 
            settings = _CastTo(settings_base, 'IAS_FftPsfCrossSection')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FftPsfLineEdgeSpread: 
            settings = _CastTo(settings_base, 'IAS_FftPsfLineEdgeSpread')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensPsfCrossSection:
            settings = _CastTo(settings_base, 'IAS_HuygensPsfCrossSection')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_HuygensPsf:
            settings = _CastTo(settings_base, 'IAS_HuygensPsf')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_DiffractionEncircledEnergy:
            settings = _CastTo(settings_base, 'IAS_DiffractionEncircledEnergy')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricEncircledEnergy:
            settings = _CastTo(settings_base, 'IAS_GeometricEncircledEnergy')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricLineEdgeSpread:
            settings = _CastTo(settings_base, 'IAS_GeometricLineEdgeSpread')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ExtendedSourceEncircledEnergy:
            settings = _CastTo(settings_base, 'IAS_ExtendedSourceEncircledEnergy')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SurfaceCurvatureCross:
            settings = _CastTo(settings_base, 'IAS_SurfaceCurvatureCross')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SurfacePhaseCross:
            settings = _CastTo(settings_base, 'IAS_SurfacePhaseCross')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SurfaceSagCross:
            settings = _CastTo(settings_base, 'IAS_SurfaceSagCross')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SurfaceCurvature:
            settings = _CastTo(settings_base, 'IAS_SurfaceCurvature')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SurfacePhase:
            settings = _CastTo(settings_base, 'IAS_SurfacePhase')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SurfaceSag:
            settings = _CastTo(settings_base, 'IAS_SurfaceSag')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_StandardSpot:
            settings = _CastTo(settings_base, 'IAS_StandardSpot')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ThroughFocusSpot:
            settings = _CastTo(settings_base, 'IAS_ThroughFocusSpot')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FullFieldSpot:
            settings = _CastTo(settings_base, 'IAS_FullFieldSpot')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_MatrixSpot:
            settings = _CastTo(settings_base, 'IAS_MatrixSpot')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ConfigurationMatrixSpot:
            settings = _CastTo(settings_base, 'IAS_ConfigurationMatrixSpot')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RMSField:
            settings = _CastTo(settings_base, 'IAS_RMSField')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RMSFieldMap:
            settings = _CastTo(settings_base, 'IAS_RMSFieldMap')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RMSLambdaDiagram:
            settings = _CastTo(settings_base, 'IAS_RMSLambdaDiagram')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RMSFocus:
            settings = _CastTo(settings_base, 'IAS_RMSFocus')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_Foucault:
            settings = _CastTo(settings_base, 'IAS_Foucault')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_Interferogram:
            settings = _CastTo(settings_base, 'IAS_Interferogram')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_WavefrontMap:
            settings = _CastTo(settings_base, 'IAS_WavefrontMap')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_DetectorViewer:
            settings = _CastTo(settings_base, 'IAS_DetectorViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_Draw2D:
            settings = _CastTo(settings_base, 'IAS_Draw2D')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_Draw3D:
            settings = _CastTo(settings_base, 'IAS_Draw3D')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ImageSimulation:
            settings = _CastTo(settings_base, 'IAS_ImageSimulation')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricImageAnalysis:
            settings = _CastTo(settings_base, 'IAS_GeometricImageAnalysis')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_IMABIMFileViewer:
            settings = _CastTo(settings_base, 'IAS_IMABIMFileViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GeometricBitmapImageAnalysis:
            settings = _CastTo(settings_base, 'IAS_GeometricBitmapImageAnalysis')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_BitmapFileViewer:
            settings = _CastTo(settings_base, 'IAS_BitmapFileViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_LightSourceAnalysis:
            settings = _CastTo(settings_base, 'IAS_LightSourceAnalysis')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PartiallyCoherentImageAnalysis:
            settings = _CastTo(settings_base, 'IAS_PartiallyCoherentImageAnalysis')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ExtendedDiffractionImageAnalysis:
            settings = _CastTo(settings_base, 'IAS_ExtendedDiffractionImageAnalysis')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_BiocularFieldOfViewAnalysis:
            settings = _CastTo(settings_base, 'IAS_BiocularFieldOfViewAnalysis')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_BiocularDipvergenceConvergence:
            settings = _CastTo(settings_base, 'IAS_BiocularDipvergenceConvergence')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RelativeIllumination:
            settings = _CastTo(settings_base, 'IAS_RelativeIllumination')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_VignettingDiagramSettings:
            settings = _CastTo(settings_base, 'IAS_VignettingDiagramSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FootprintSettings:
            settings = _CastTo(settings_base, 'IAS_FootprintSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_YYbarDiagram:
            settings = _CastTo(settings_base, 'IAS_YYbarDiagram')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PowerFieldMapSettings:
            settings = _CastTo(settings_base, 'IAS_PowerFieldMapSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PowerPupilMapSettings:
            settings = _CastTo(settings_base, 'IAS_PowerPupilMapSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_IncidentAnglevsImageHeight:
            settings = _CastTo(settings_base, 'IAS_IncidentAnglevsImageHeight')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FiberCouplingSettings:
            settings = _CastTo(settings_base, 'IAS_FiberCouplingSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_YNIContributions:
            settings = _CastTo(settings_base, 'IAS_YNIContributions')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SagTable:
            settings = _CastTo(settings_base, 'IAS_SagTable')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_CardinalPoints:
            settings = _CastTo(settings_base, 'IAS_CardinalPoints')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_DispersionDiagram:
            settings = _CastTo(settings_base, 'IAS_DispersionDiagram')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GlassMap:
            settings = _CastTo(settings_base, 'IAS_GlassMap')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_AthermalGlassMap:
            settings = _CastTo(settings_base, 'IAS_AthermalGlassMap')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_InternalTransmissionvsWavelength:
            settings = _CastTo(settings_base, 'IAS_InternalTransmissionvsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_DispersionvsWavelength:
            settings = _CastTo(settings_base, 'IAS_DispersionvsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GrinProfile:
            settings = _CastTo(settings_base, 'IAS_GrinProfile')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_GradiumProfile:
            settings = _CastTo(settings_base, 'IAS_GradiumProfile')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_UniversalPlot1D:
            settings = _CastTo(settings_base, 'IAS_UniversalPlot1D')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_UniversalPlot2D:
            settings = _CastTo(settings_base, 'IAS_UniversalPlot2D')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PolarizationRayTrace:
            settings = _CastTo(settings_base, 'IAS_PolarizationRayTrace')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PolarizationPupilMap:
            settings = _CastTo(settings_base, 'IAS_PolarizationPupilMap')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_Transmission:
            settings = _CastTo(settings_base, 'IAS_Transmission')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PhaseAberration:
            settings = _CastTo(settings_base, 'IAS_PhaseAberration')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_TransmissionFan:
            settings = _CastTo(settings_base, 'IAS_TransmissionFan')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ParaxialGaussianBeam:
            settings = _CastTo(settings_base, 'IAS_ParaxialGaussianBeam')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SkewGaussianBeam:
            settings = _CastTo(settings_base, 'IAS_SkewGaussianBeam')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PhysicalOpticsPropagation:
            settings = _CastTo(settings_base, 'IAS_PhysicalOpticsPropagation')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_BeamFileViewer:
            settings = _CastTo(settings_base, 'IAS_BeamFileViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ReflectionvsAngle:
            settings = _CastTo(settings_base, 'IAS_ReflectionvsAngle')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_TransmissionvsAngle:
            settings = _CastTo(settings_base, 'IAS_TransmissionvsAngle')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_AbsorptionvsAngle:
            settings = _CastTo(settings_base, 'IAS_AbsorptionvsAngle')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_DiattenuationvsAngle:
            settings = _CastTo(settings_base, 'IAS_DiattenuationvsAngle')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PhasevsAngle:
            settings = _CastTo(settings_base, 'IAS_PhasevsAngle')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RetardancevsAngle:
            settings = _CastTo(settings_base, 'IAS_RetardancevsAngle')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ReflectionvsWavelength:
            settings = _CastTo(settings_base, 'IAS_ReflectionvsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_TransmissionvsWavelength:
            settings = _CastTo(settings_base, 'IAS_TransmissionvsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_AbsorptionvsWavelength:
            settings = _CastTo(settings_base, 'IAS_AbsorptionvsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_DiattenuationvsWavelength:
            settings = _CastTo(settings_base, 'IAS_DiattenuationvsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PhasevsWavelength:
            settings = _CastTo(settings_base, 'IAS_PhasevsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RetardancevsWavelength:
            settings = _CastTo(settings_base, 'IAS_RetardancevsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_DirectivityPlot:
            settings = _CastTo(settings_base, 'IAS_DirectivityPlot')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SourcePolarViewer:
            settings = _CastTo(settings_base, 'IAS_SourcePolarViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PhotoluminscenceViewer:
            settings = _CastTo(settings_base, 'IAS_PhotoluminscenceViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SourceSpectrumViewer:
            settings = _CastTo(settings_base, 'IAS_SourceSpectrumViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RadiantSourceModelViewerSettings:
            settings = _CastTo(settings_base, 'IAS_RadiantSourceModelViewerSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SurfaceDataSettings:
            settings = _CastTo(settings_base, 'IAS_SurfaceDataSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PrescriptionDataSettings:
            settings = _CastTo(settings_base, 'IAS_PrescriptionDataSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FileComparatorSettings:
            settings = _CastTo(settings_base, 'IAS_FileComparatorSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PartViewer:
            settings = _CastTo(settings_base, 'IAS_PartViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ReverseRadianceAnalysis:
            settings = _CastTo(settings_base, 'IAS_ReverseRadianceAnalysis') 
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PathAnalysis:
            settings = _CastTo(settings_base, 'IAS_PathAnalysis')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_FluxvsWavelength:
            settings = _CastTo(settings_base, 'IAS_FluxvsWavelength')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RoadwayLighting:
            settings = _CastTo(settings_base, 'IAS_RoadwayLighting')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SourceIlluminationMap:
            settings = _CastTo(settings_base, 'IAS_SourceIlluminationMap')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ScatterFunctionViewer:
            settings = _CastTo(settings_base, 'IAS_ScatterFunctionViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ScatterPolarPlotSettings:
            settings = _CastTo(settings_base, 'IAS_ScatterPolarPlotSettings')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ZemaxElementDrawing:
            settings = _CastTo(settings_base, 'IAS_ZemaxElementDrawing')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ShadedModel:
            settings = _CastTo(settings_base, 'IAS_ShadedModel')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_NSCShadedModel:
            settings = _CastTo(settings_base, 'IAS_NSCShadedModel')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_NSC3DLayout:
            settings = _CastTo(settings_base, 'IAS_NSC3DLayout')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_NSCObjectViewer:
            settings = _CastTo(settings_base, 'IAS_NSCObjectViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_RayDatabaseViewer:
            settings = _CastTo(settings_base, 'IAS_RayDatabaseViewer')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_ISOElementDrawing:
            settings = _CastTo(settings_base, 'IAS_ISOElementDrawing')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SystemData:
            settings = _CastTo(settings_base, 'IAS_SystemData')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_TestPlateList:
            settings = _CastTo(settings_base, 'IAS_TestPlateList')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SourceColorChart1931:
            settings = _CastTo(settings_base, 'IAS_SourceColorChart1931')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_SourceColorChart1976:
            settings = _CastTo(settings_base, 'IAS_SourceColorChart1976')
        elif self._ia_.AnalysisType == _constants.AnalysisIDM_PrescriptionGraphic:
            settings = _CastTo(settings_base, 'IAS_PrescriptionGraphic')
    except ValueError as err:
        _warnings.warn("Couldn't find and cast to specialized analysis settings.", stacklevel=2)
        settings = settings_base

    # create the settings object 
    settings = _wrapped_zos_object(settings)

    return settings

# Extra methods
# -------------