library(RDCOMClient)

# c:/Duncan/R-2.1.0/library/SWinTypeLibs/runtime/common.S
# Used in generating R code to interface to Type Library definitions
# and also at run-time for the generated code.

createTypeVarName = 
# Map the given names in var to a unique and legitimate
# R variable name for the given class.
# 
function(className, var, quote = TRUE)
{
  if(is(className, "ClassDefinition"))
    className = className@className
	
  ans = paste("COM", className, var, sep = ".") 
  if(quote) {
    ans = paste("'", ans, "'", sep = "")
  }
  names(ans) = var

  ans
}


library(RDCOMClient)

setClass('_Application', contains = 'CompiledCOMIDispatch')
setClass('Workbooks', contains = 'COMTypedNamedList', prototype = list(name = '_Workbook'))
'ConstantsEnum' = c(
   'xlAll' = -4104,
   'xlAutomatic' = -4105,
   'xlBoth' = 1,
   'xlCenter' = -4108,
   'xlChecker' = 9,
   'xlCircle' = 8,
   'xlCorner' = 2,
   'xlCrissCross' = 16,
   'xlCross' = 4,
   'xlDiamond' = 2,
   'xlDistributed' = -4117,
   'xlDoubleAccounting' = 5,
   'xlFixedValue' = 1,
   'xlFormats' = -4122,
   'xlGray16' = 17,
   'xlGray8' = 18,
   'xlGrid' = 15,
   'xlHigh' = -4127,
   'xlInside' = 2,
   'xlJustify' = -4130,
   'xlLightDown' = 13,
   'xlLightHorizontal' = 11,
   'xlLightUp' = 14,
   'xlLightVertical' = 12,
   'xlLow' = -4134,
   'xlManual' = -4135,
   'xlMinusValues' = 3,
   'xlModule' = -4141,
   'xlNextToAxis' = 4,
   'xlNone' = -4142,
   'xlNotes' = -4144,
   'xlOff' = -4146,
   'xlOn' = 1,
   'xlPercent' = 2,
   'xlPlus' = 9,
   'xlPlusValues' = 2,
   'xlSemiGray75' = 10,
   'xlShowLabel' = 4,
   'xlShowLabelAndPercent' = 5,
   'xlShowPercent' = 3,
   'xlShowValue' = 2,
   'xlSimple' = -4154,
   'xlSingle' = 2,
   'xlSingleAccounting' = 4,
   'xlSolid' = 1,
   'xlSquare' = 1,
   'xlStar' = 5,
   'xlStError' = 4,
   'xlToolbarButton' = 2,
   'xlTriangle' = 3,
   'xlGray25' = -4124,
   'xlGray50' = -4125,
   'xlGray75' = -4126,
   'xlBottom' = -4107,
   'xlLeft' = -4131,
   'xlRight' = -4152,
   'xlTop' = -4160,
   'xl3DBar' = -4099,
   'xl3DSurface' = -4103,
   'xlBar' = 2,
   'xlColumn' = 3,
   'xlCombination' = -4111,
   'xlCustom' = -4114,
   'xlDefaultAutoFormat' = -1,
   'xlMaximum' = 2,
   'xlMinimum' = 4,
   'xlOpaque' = 3,
   'xlTransparent' = 2,
   'xlBidi' = -5000,
   'xlLatin' = -5001,
   'xlContext' = -5002,
   'xlLTR' = -5003,
   'xlRTL' = -5004,
   'xlFullScript' = 1,
   'xlPartialScript' = 2,
   'xlMixedScript' = 3,
   'xlMixedAuthorizedScript' = 4,
   'xlVisualCursor' = 2,
   'xlLogicalCursor' = 1,
   'xlSystem' = 1,
   'xlPartial' = 3,
   'xlHindiNumerals' = 3,
   'xlBidiCalendar' = 3,
   'xlGregorian' = 2,
   'xlComplete' = 4,
   'xlScale' = 3,
   'xlClosed' = 3,
   'xlColor1' = 7,
   'xlColor2' = 8,
   'xlColor3' = 9,
   'xlConstants' = 2,
   'xlContents' = 2,
   'xlBelow' = 1,
   'xlCascade' = 7,
   'xlCenterAcrossSelection' = 7,
   'xlChart4' = 2,
   'xlChartSeries' = 17,
   'xlChartShort' = 6,
   'xlChartTitles' = 18,
   'xlClassic1' = 1,
   'xlClassic2' = 2,
   'xlClassic3' = 3,
   'xl3DEffects1' = 13,
   'xl3DEffects2' = 14,
   'xlAbove' = 0,
   'xlAccounting1' = 4,
   'xlAccounting2' = 5,
   'xlAccounting3' = 6,
   'xlAccounting4' = 17,
   'xlAdd' = 2,
   'xlDebugCodePane' = 13,
   'xlDesktop' = 9,
   'xlDirect' = 1,
   'xlDivide' = 5,
   'xlDoubleClosed' = 5,
   'xlDoubleOpen' = 4,
   'xlDoubleQuote' = 1,
   'xlEntireChart' = 20,
   'xlExcelMenus' = 1,
   'xlExtended' = 3,
   'xlFill' = 5,
   'xlFirst' = 0,
   'xlFloating' = 5,
   'xlFormula' = 5,
   'xlGeneral' = 1,
   'xlGridline' = 22,
   'xlIcons' = 1,
   'xlImmediatePane' = 12,
   'xlInteger' = 2,
   'xlLast' = 1,
   'xlLastCell' = 11,
   'xlList1' = 10,
   'xlList2' = 11,
   'xlList3' = 12,
   'xlLocalFormat1' = 15,
   'xlLocalFormat2' = 16,
   'xlLong' = 3,
   'xlLotusHelp' = 2,
   'xlMacrosheetCell' = 7,
   'xlMixed' = 2,
   'xlMultiply' = 4,
   'xlNarrow' = 1,
   'xlNoDocuments' = 3,
   'xlOpen' = 2,
   'xlOutside' = 3,
   'xlReference' = 4,
   'xlSemiautomatic' = 2,
   'xlShort' = 1,
   'xlSingleQuote' = 2,
   'xlStrict' = 2,
   'xlSubtract' = 3,
   'xlTextBox' = 16,
   'xlTiled' = 1,
   'xlTitleBar' = 8,
   'xlToolbar' = 1,
   'xlVisible' = 12,
   'xlWatchPane' = 11,
   'xlWide' = 3,
   'xlWorkbookTab' = 6,
   'xlWorksheet4' = 1,
   'xlWorksheetCell' = 3,
   'xlWorksheetShort' = 5,
   'xlAllExceptBorders' = 7,
   'xlLeftToRight' = 2,
   'xlTopToBottom' = 1,
   'xlVeryHidden' = 2,
   'xlDrawingObject' = 14 
 )
storage.mode( ConstantsEnum ) = 'integer'
'XlCreatorEnum' = c(
   'xlCreatorCode' = 1480803660 
 )
storage.mode( XlCreatorEnum ) = 'integer'
'XlChartGalleryEnum' = c(
   'xlBuiltIn' = 21,
   'xlUserDefined' = 22,
   'xlAnyGallery' = 23 
 )
storage.mode( XlChartGalleryEnum ) = 'integer'
'XlColorIndexEnum' = c(
   'xlColorIndexAutomatic' = -4105,
   'xlColorIndexNone' = -4142 
 )
storage.mode( XlColorIndexEnum ) = 'integer'
'XlEndStyleCapEnum' = c(
   'xlCap' = 1,
   'xlNoCap' = 2 
 )
storage.mode( XlEndStyleCapEnum ) = 'integer'
'XlRowColEnum' = c(
   'xlColumns' = 2,
   'xlRows' = 1 
 )
storage.mode( XlRowColEnum ) = 'integer'
'XlScaleTypeEnum' = c(
   'xlScaleLinear' = -4132,
   'xlScaleLogarithmic' = -4133 
 )
storage.mode( XlScaleTypeEnum ) = 'integer'
'XlDataSeriesTypeEnum' = c(
   'xlAutoFill' = 4,
   'xlChronological' = 3,
   'xlGrowth' = 2,
   'xlDataSeriesLinear' = -4132 
 )
storage.mode( XlDataSeriesTypeEnum ) = 'integer'
'XlAxisCrossesEnum' = c(
   'xlAxisCrossesAutomatic' = -4105,
   'xlAxisCrossesCustom' = -4114,
   'xlAxisCrossesMaximum' = 2,
   'xlAxisCrossesMinimum' = 4 
 )
storage.mode( XlAxisCrossesEnum ) = 'integer'
'XlAxisGroupEnum' = c(
   'xlPrimary' = 1,
   'xlSecondary' = 2 
 )
storage.mode( XlAxisGroupEnum ) = 'integer'
'XlBackgroundEnum' = c(
   'xlBackgroundAutomatic' = -4105,
   'xlBackgroundOpaque' = 3,
   'xlBackgroundTransparent' = 2 
 )
storage.mode( XlBackgroundEnum ) = 'integer'
'XlWindowStateEnum' = c(
   'xlMaximized' = -4137,
   'xlMinimized' = -4140,
   'xlNormal' = -4143 
 )
storage.mode( XlWindowStateEnum ) = 'integer'
'XlAxisTypeEnum' = c(
   'xlCategory' = 1,
   'xlSeriesAxis' = 3,
   'xlValue' = 2 
 )
storage.mode( XlAxisTypeEnum ) = 'integer'
'XlArrowHeadLengthEnum' = c(
   'xlArrowHeadLengthLong' = 3,
   'xlArrowHeadLengthMedium' = -4138,
   'xlArrowHeadLengthShort' = 1 
 )
storage.mode( XlArrowHeadLengthEnum ) = 'integer'
'XlVAlignEnum' = c(
   'xlVAlignBottom' = -4107,
   'xlVAlignCenter' = -4108,
   'xlVAlignDistributed' = -4117,
   'xlVAlignJustify' = -4130,
   'xlVAlignTop' = -4160 
 )
storage.mode( XlVAlignEnum ) = 'integer'
'XlTickMarkEnum' = c(
   'xlTickMarkCross' = 4,
   'xlTickMarkInside' = 2,
   'xlTickMarkNone' = -4142,
   'xlTickMarkOutside' = 3 
 )
storage.mode( XlTickMarkEnum ) = 'integer'
'XlErrorBarDirectionEnum' = c(
   'xlX' = -4168,
   'xlY' = 1 
 )
storage.mode( XlErrorBarDirectionEnum ) = 'integer'
'XlErrorBarIncludeEnum' = c(
   'xlErrorBarIncludeBoth' = 1,
   'xlErrorBarIncludeMinusValues' = 3,
   'xlErrorBarIncludeNone' = -4142,
   'xlErrorBarIncludePlusValues' = 2 
 )
storage.mode( XlErrorBarIncludeEnum ) = 'integer'
'XlDisplayBlanksAsEnum' = c(
   'xlInterpolated' = 3,
   'xlNotPlotted' = 1,
   'xlZero' = 2 
 )
storage.mode( XlDisplayBlanksAsEnum ) = 'integer'
'XlArrowHeadStyleEnum' = c(
   'xlArrowHeadStyleClosed' = 3,
   'xlArrowHeadStyleDoubleClosed' = 5,
   'xlArrowHeadStyleDoubleOpen' = 4,
   'xlArrowHeadStyleNone' = -4142,
   'xlArrowHeadStyleOpen' = 2 
 )
storage.mode( XlArrowHeadStyleEnum ) = 'integer'
'XlArrowHeadWidthEnum' = c(
   'xlArrowHeadWidthMedium' = -4138,
   'xlArrowHeadWidthNarrow' = 1,
   'xlArrowHeadWidthWide' = 3 
 )
storage.mode( XlArrowHeadWidthEnum ) = 'integer'
'XlHAlignEnum' = c(
   'xlHAlignCenter' = -4108,
   'xlHAlignCenterAcrossSelection' = 7,
   'xlHAlignDistributed' = -4117,
   'xlHAlignFill' = 5,
   'xlHAlignGeneral' = 1,
   'xlHAlignJustify' = -4130,
   'xlHAlignLeft' = -4131,
   'xlHAlignRight' = -4152 
 )
storage.mode( XlHAlignEnum ) = 'integer'
'XlTickLabelPositionEnum' = c(
   'xlTickLabelPositionHigh' = -4127,
   'xlTickLabelPositionLow' = -4134,
   'xlTickLabelPositionNextToAxis' = 4,
   'xlTickLabelPositionNone' = -4142 
 )
storage.mode( XlTickLabelPositionEnum ) = 'integer'
'XlLegendPositionEnum' = c(
   'xlLegendPositionBottom' = -4107,
   'xlLegendPositionCorner' = 2,
   'xlLegendPositionLeft' = -4131,
   'xlLegendPositionRight' = -4152,
   'xlLegendPositionTop' = -4160 
 )
storage.mode( XlLegendPositionEnum ) = 'integer'
'XlChartPictureTypeEnum' = c(
   'xlStackScale' = 3,
   'xlStack' = 2,
   'xlStretch' = 1 
 )
storage.mode( XlChartPictureTypeEnum ) = 'integer'
'XlChartPicturePlacementEnum' = c(
   'xlSides' = 1,
   'xlEnd' = 2,
   'xlEndSides' = 3,
   'xlFront' = 4,
   'xlFrontSides' = 5,
   'xlFrontEnd' = 6,
   'xlAllFaces' = 7 
 )
storage.mode( XlChartPicturePlacementEnum ) = 'integer'
'XlOrientationEnum' = c(
   'xlDownward' = -4170,
   'xlHorizontal' = -4128,
   'xlUpward' = -4171,
   'xlVertical' = -4166 
 )
storage.mode( XlOrientationEnum ) = 'integer'
'XlTickLabelOrientationEnum' = c(
   'xlTickLabelOrientationAutomatic' = -4105,
   'xlTickLabelOrientationDownward' = -4170,
   'xlTickLabelOrientationHorizontal' = -4128,
   'xlTickLabelOrientationUpward' = -4171,
   'xlTickLabelOrientationVertical' = -4166 
 )
storage.mode( XlTickLabelOrientationEnum ) = 'integer'
'XlBorderWeightEnum' = c(
   'xlHairline' = 1,
   'xlMedium' = -4138,
   'xlThick' = 4,
   'xlThin' = 2 
 )
storage.mode( XlBorderWeightEnum ) = 'integer'
'XlDataSeriesDateEnum' = c(
   'xlDay' = 1,
   'xlMonth' = 3,
   'xlWeekday' = 2,
   'xlYear' = 4 
 )
storage.mode( XlDataSeriesDateEnum ) = 'integer'
'XlUnderlineStyleEnum' = c(
   'xlUnderlineStyleDouble' = -4119,
   'xlUnderlineStyleDoubleAccounting' = 5,
   'xlUnderlineStyleNone' = -4142,
   'xlUnderlineStyleSingle' = 2,
   'xlUnderlineStyleSingleAccounting' = 4 
 )
storage.mode( XlUnderlineStyleEnum ) = 'integer'
'XlErrorBarTypeEnum' = c(
   'xlErrorBarTypeCustom' = -4114,
   'xlErrorBarTypeFixedValue' = 1,
   'xlErrorBarTypePercent' = 2,
   'xlErrorBarTypeStDev' = -4155,
   'xlErrorBarTypeStError' = 4 
 )
storage.mode( XlErrorBarTypeEnum ) = 'integer'
'XlTrendlineTypeEnum' = c(
   'xlExponential' = 5,
   'xlLinear' = -4132,
   'xlLogarithmic' = -4133,
   'xlMovingAvg' = 6,
   'xlPolynomial' = 3,
   'xlPower' = 4 
 )
storage.mode( XlTrendlineTypeEnum ) = 'integer'
'XlLineStyleEnum' = c(
   'xlContinuous' = 1,
   'xlDash' = -4115,
   'xlDashDot' = 4,
   'xlDashDotDot' = 5,
   'xlDot' = -4118,
   'xlDouble' = -4119,
   'xlSlantDashDot' = 13,
   'xlLineStyleNone' = -4142 
 )
storage.mode( XlLineStyleEnum ) = 'integer'
'XlDataLabelsTypeEnum' = c(
   'xlDataLabelsShowNone' = -4142,
   'xlDataLabelsShowValue' = 2,
   'xlDataLabelsShowPercent' = 3,
   'xlDataLabelsShowLabel' = 4,
   'xlDataLabelsShowLabelAndPercent' = 5,
   'xlDataLabelsShowBubbleSizes' = 6 
 )
storage.mode( XlDataLabelsTypeEnum ) = 'integer'
'XlMarkerStyleEnum' = c(
   'xlMarkerStyleAutomatic' = -4105,
   'xlMarkerStyleCircle' = 8,
   'xlMarkerStyleDash' = -4115,
   'xlMarkerStyleDiamond' = 2,
   'xlMarkerStyleDot' = -4118,
   'xlMarkerStyleNone' = -4142,
   'xlMarkerStylePicture' = -4147,
   'xlMarkerStylePlus' = 9,
   'xlMarkerStyleSquare' = 1,
   'xlMarkerStyleStar' = 5,
   'xlMarkerStyleTriangle' = 3,
   'xlMarkerStyleX' = -4168 
 )
storage.mode( XlMarkerStyleEnum ) = 'integer'
'XlPictureConvertorTypeEnum' = c(
   'xlBMP' = 1,
   'xlCGM' = 7,
   'xlDRW' = 4,
   'xlDXF' = 5,
   'xlEPS' = 8,
   'xlHGL' = 6,
   'xlPCT' = 13,
   'xlPCX' = 10,
   'xlPIC' = 11,
   'xlPLT' = 12,
   'xlTIF' = 9,
   'xlWMF' = 2,
   'xlWPG' = 3 
 )
storage.mode( XlPictureConvertorTypeEnum ) = 'integer'
'XlPatternEnum' = c(
   'xlPatternAutomatic' = -4105,
   'xlPatternChecker' = 9,
   'xlPatternCrissCross' = 16,
   'xlPatternDown' = -4121,
   'xlPatternGray16' = 17,
   'xlPatternGray25' = -4124,
   'xlPatternGray50' = -4125,
   'xlPatternGray75' = -4126,
   'xlPatternGray8' = 18,
   'xlPatternGrid' = 15,
   'xlPatternHorizontal' = -4128,
   'xlPatternLightDown' = 13,
   'xlPatternLightHorizontal' = 11,
   'xlPatternLightUp' = 14,
   'xlPatternLightVertical' = 12,
   'xlPatternNone' = -4142,
   'xlPatternSemiGray75' = 10,
   'xlPatternSolid' = 1,
   'xlPatternUp' = -4162,
   'xlPatternVertical' = -4166 
 )
storage.mode( XlPatternEnum ) = 'integer'
'XlChartSplitTypeEnum' = c(
   'xlSplitByPosition' = 1,
   'xlSplitByPercentValue' = 3,
   'xlSplitByCustomSplit' = 4,
   'xlSplitByValue' = 2 
 )
storage.mode( XlChartSplitTypeEnum ) = 'integer'
'XlDisplayUnitEnum' = c(
   'xlHundreds' = -2,
   'xlThousands' = -3,
   'xlTenThousands' = -4,
   'xlHundredThousands' = -5,
   'xlMillions' = -6,
   'xlTenMillions' = -7,
   'xlHundredMillions' = -8,
   'xlThousandMillions' = -9,
   'xlMillionMillions' = -10 
 )
storage.mode( XlDisplayUnitEnum ) = 'integer'
'XlDataLabelPositionEnum' = c(
   'xlLabelPositionCenter' = -4108,
   'xlLabelPositionAbove' = 0,
   'xlLabelPositionBelow' = 1,
   'xlLabelPositionLeft' = -4131,
   'xlLabelPositionRight' = -4152,
   'xlLabelPositionOutsideEnd' = 2,
   'xlLabelPositionInsideEnd' = 3,
   'xlLabelPositionInsideBase' = 4,
   'xlLabelPositionBestFit' = 5,
   'xlLabelPositionMixed' = 6,
   'xlLabelPositionCustom' = 7 
 )
storage.mode( XlDataLabelPositionEnum ) = 'integer'
'XlTimeUnitEnum' = c(
   'xlDays' = 0,
   'xlMonths' = 1,
   'xlYears' = 2 
 )
storage.mode( XlTimeUnitEnum ) = 'integer'
'XlCategoryTypeEnum' = c(
   'xlCategoryScale' = 2,
   'xlTimeScale' = 3,
   'xlAutomaticScale' = -4105 
 )
storage.mode( XlCategoryTypeEnum ) = 'integer'
'XlBarShapeEnum' = c(
   'xlBox' = 0,
   'xlPyramidToPoint' = 1,
   'xlPyramidToMax' = 2,
   'xlCylinder' = 3,
   'xlConeToPoint' = 4,
   'xlConeToMax' = 5 
 )
storage.mode( XlBarShapeEnum ) = 'integer'
'XlChartTypeEnum' = c(
   'xlColumnClustered' = 51,
   'xlColumnStacked' = 52,
   'xlColumnStacked100' = 53,
   'xl3DColumnClustered' = 54,
   'xl3DColumnStacked' = 55,
   'xl3DColumnStacked100' = 56,
   'xlBarClustered' = 57,
   'xlBarStacked' = 58,
   'xlBarStacked100' = 59,
   'xl3DBarClustered' = 60,
   'xl3DBarStacked' = 61,
   'xl3DBarStacked100' = 62,
   'xlLineStacked' = 63,
   'xlLineStacked100' = 64,
   'xlLineMarkers' = 65,
   'xlLineMarkersStacked' = 66,
   'xlLineMarkersStacked100' = 67,
   'xlPieOfPie' = 68,
   'xlPieExploded' = 69,
   'xl3DPieExploded' = 70,
   'xlBarOfPie' = 71,
   'xlXYScatterSmooth' = 72,
   'xlXYScatterSmoothNoMarkers' = 73,
   'xlXYScatterLines' = 74,
   'xlXYScatterLinesNoMarkers' = 75,
   'xlAreaStacked' = 76,
   'xlAreaStacked100' = 77,
   'xl3DAreaStacked' = 78,
   'xl3DAreaStacked100' = 79,
   'xlDoughnutExploded' = 80,
   'xlRadarMarkers' = 81,
   'xlRadarFilled' = 82,
   'xlSurface' = 83,
   'xlSurfaceWireframe' = 84,
   'xlSurfaceTopView' = 85,
   'xlSurfaceTopViewWireframe' = 86,
   'xlBubble' = 15,
   'xlBubble3DEffect' = 87,
   'xlStockHLC' = 88,
   'xlStockOHLC' = 89,
   'xlStockVHLC' = 90,
   'xlStockVOHLC' = 91,
   'xlCylinderColClustered' = 92,
   'xlCylinderColStacked' = 93,
   'xlCylinderColStacked100' = 94,
   'xlCylinderBarClustered' = 95,
   'xlCylinderBarStacked' = 96,
   'xlCylinderBarStacked100' = 97,
   'xlCylinderCol' = 98,
   'xlConeColClustered' = 99,
   'xlConeColStacked' = 100,
   'xlConeColStacked100' = 101,
   'xlConeBarClustered' = 102,
   'xlConeBarStacked' = 103,
   'xlConeBarStacked100' = 104,
   'xlConeCol' = 105,
   'xlPyramidColClustered' = 106,
   'xlPyramidColStacked' = 107,
   'xlPyramidColStacked100' = 108,
   'xlPyramidBarClustered' = 109,
   'xlPyramidBarStacked' = 110,
   'xlPyramidBarStacked100' = 111,
   'xlPyramidCol' = 112,
   'xl3DColumn' = -4100,
   'xlLine' = 4,
   'xl3DLine' = -4101,
   'xl3DPie' = -4102,
   'xlPie' = 5,
   'xlXYScatter' = -4169,
   'xl3DArea' = -4098,
   'xlArea' = 1,
   'xlDoughnut' = -4120,
   'xlRadar' = -4151 
 )
storage.mode( XlChartTypeEnum ) = 'integer'
'XlChartItemEnum' = c(
   'xlDataLabel' = 0,
   'xlChartArea' = 2,
   'xlSeries' = 3,
   'xlChartTitle' = 4,
   'xlWalls' = 5,
   'xlCorners' = 6,
   'xlDataTable' = 7,
   'xlTrendline' = 8,
   'xlErrorBars' = 9,
   'xlXErrorBars' = 10,
   'xlYErrorBars' = 11,
   'xlLegendEntry' = 12,
   'xlLegendKey' = 13,
   'xlShape' = 14,
   'xlMajorGridlines' = 15,
   'xlMinorGridlines' = 16,
   'xlAxisTitle' = 17,
   'xlUpBars' = 18,
   'xlPlotArea' = 19,
   'xlDownBars' = 20,
   'xlAxis' = 21,
   'xlSeriesLines' = 22,
   'xlFloor' = 23,
   'xlLegend' = 24,
   'xlHiLoLines' = 25,
   'xlDropLines' = 26,
   'xlRadarAxisLabels' = 27,
   'xlNothing' = 28,
   'xlLeaderLines' = 29,
   'xlDisplayUnitLabel' = 30,
   'xlPivotChartFieldButton' = 31,
   'xlPivotChartDropZone' = 32 
 )
storage.mode( XlChartItemEnum ) = 'integer'
'XlSizeRepresentsEnum' = c(
   'xlSizeIsWidth' = 2,
   'xlSizeIsArea' = 1 
 )
storage.mode( XlSizeRepresentsEnum ) = 'integer'
'XlInsertShiftDirectionEnum' = c(
   'xlShiftDown' = -4121,
   'xlShiftToRight' = -4161 
 )
storage.mode( XlInsertShiftDirectionEnum ) = 'integer'
'XlDeleteShiftDirectionEnum' = c(
   'xlShiftToLeft' = -4159,
   'xlShiftUp' = -4162 
 )
storage.mode( XlDeleteShiftDirectionEnum ) = 'integer'
'XlDirectionEnum' = c(
   'xlDown' = -4121,
   'xlToLeft' = -4159,
   'xlToRight' = -4161,
   'xlUp' = -4162 
 )
storage.mode( XlDirectionEnum ) = 'integer'
'XlConsolidationFunctionEnum' = c(
   'xlAverage' = -4106,
   'xlCount' = -4112,
   'xlCountNums' = -4113,
   'xlMax' = -4136,
   'xlMin' = -4139,
   'xlProduct' = -4149,
   'xlStDev' = -4155,
   'xlStDevP' = -4156,
   'xlSum' = -4157,
   'xlVar' = -4164,
   'xlVarP' = -4165,
   'xlUnknown' = 1000 
 )
storage.mode( XlConsolidationFunctionEnum ) = 'integer'
'XlSheetTypeEnum' = c(
   'xlChart' = -4109,
   'xlDialogSheet' = -4116,
   'xlExcel4IntlMacroSheet' = 4,
   'xlExcel4MacroSheet' = 3,
   'xlWorksheet' = -4167 
 )
storage.mode( XlSheetTypeEnum ) = 'integer'
'XlLocationInTableEnum' = c(
   'xlColumnHeader' = -4110,
   'xlColumnItem' = 5,
   'xlDataHeader' = 3,
   'xlDataItem' = 7,
   'xlPageHeader' = 2,
   'xlPageItem' = 6,
   'xlRowHeader' = -4153,
   'xlRowItem' = 4,
   'xlTableBody' = 8 
 )
storage.mode( XlLocationInTableEnum ) = 'integer'
'XlFindLookInEnum' = c(
   'xlFormulas' = -4123,
   'xlComments' = -4144,
   'xlValues' = -4163 
 )
storage.mode( XlFindLookInEnum ) = 'integer'
'XlWindowTypeEnum' = c(
   'xlChartAsWindow' = 5,
   'xlChartInPlace' = 4,
   'xlClipboard' = 3,
   'xlInfo' = -4129,
   'xlWorkbook' = 1 
 )
storage.mode( XlWindowTypeEnum ) = 'integer'
'XlPivotFieldDataTypeEnum' = c(
   'xlDate' = 2,
   'xlNumber' = -4145,
   'xlText' = -4158 
 )
storage.mode( XlPivotFieldDataTypeEnum ) = 'integer'
'XlCopyPictureFormatEnum' = c(
   'xlBitmap' = 2,
   'xlPicture' = -4147 
 )
storage.mode( XlCopyPictureFormatEnum ) = 'integer'
'XlPivotTableSourceTypeEnum' = c(
   'xlScenario' = 4,
   'xlConsolidation' = 3,
   'xlDatabase' = 1,
   'xlExternal' = 2,
   'xlPivotTable' = -4148 
 )
storage.mode( XlPivotTableSourceTypeEnum ) = 'integer'
'XlReferenceStyleEnum' = c(
   'xlA1' = 1,
   'xlR1C1' = -4150 
 )
storage.mode( XlReferenceStyleEnum ) = 'integer'
'XlMSApplicationEnum' = c(
   'xlMicrosoftAccess' = 4,
   'xlMicrosoftFoxPro' = 5,
   'xlMicrosoftMail' = 3,
   'xlMicrosoftPowerPoint' = 2,
   'xlMicrosoftProject' = 6,
   'xlMicrosoftSchedulePlus' = 7,
   'xlMicrosoftWord' = 1 
 )
storage.mode( XlMSApplicationEnum ) = 'integer'
'XlMouseButtonEnum' = c(
   'xlNoButton' = 0,
   'xlPrimaryButton' = 1,
   'xlSecondaryButton' = 2 
 )
storage.mode( XlMouseButtonEnum ) = 'integer'
'XlCutCopyModeEnum' = c(
   'xlCopy' = 1,
   'xlCut' = 2 
 )
storage.mode( XlCutCopyModeEnum ) = 'integer'
'XlFillWithEnum' = c(
   'xlFillWithAll' = -4104,
   'xlFillWithContents' = 2,
   'xlFillWithFormats' = -4122 
 )
storage.mode( XlFillWithEnum ) = 'integer'
'XlFilterActionEnum' = c(
   'xlFilterCopy' = 2,
   'xlFilterInPlace' = 1 
 )
storage.mode( XlFilterActionEnum ) = 'integer'
'XlOrderEnum' = c(
   'xlDownThenOver' = 1,
   'xlOverThenDown' = 2 
 )
storage.mode( XlOrderEnum ) = 'integer'
'XlLinkTypeEnum' = c(
   'xlLinkTypeExcelLinks' = 1,
   'xlLinkTypeOLELinks' = 2 
 )
storage.mode( XlLinkTypeEnum ) = 'integer'
'XlApplyNamesOrderEnum' = c(
   'xlColumnThenRow' = 2,
   'xlRowThenColumn' = 1 
 )
storage.mode( XlApplyNamesOrderEnum ) = 'integer'
'XlEnableCancelKeyEnum' = c(
   'xlDisabled' = 0,
   'xlErrorHandler' = 2,
   'xlInterrupt' = 1 
 )
storage.mode( XlEnableCancelKeyEnum ) = 'integer'
'XlPageBreakEnum' = c(
   'xlPageBreakAutomatic' = -4105,
   'xlPageBreakManual' = -4135,
   'xlPageBreakNone' = -4142 
 )
storage.mode( XlPageBreakEnum ) = 'integer'
'XlOLETypeEnum' = c(
   'xlOLEControl' = 2,
   'xlOLEEmbed' = 1,
   'xlOLELink' = 0 
 )
storage.mode( XlOLETypeEnum ) = 'integer'
'XlPageOrientationEnum' = c(
   'xlLandscape' = 2,
   'xlPortrait' = 1 
 )
storage.mode( XlPageOrientationEnum ) = 'integer'
'XlLinkInfoEnum' = c(
   'xlEditionDate' = 2,
   'xlUpdateState' = 1,
   'xlLinkInfoStatus' = 3 
 )
storage.mode( XlLinkInfoEnum ) = 'integer'
'XlCommandUnderlinesEnum' = c(
   'xlCommandUnderlinesAutomatic' = -4105,
   'xlCommandUnderlinesOff' = -4146,
   'xlCommandUnderlinesOn' = 1 
 )
storage.mode( XlCommandUnderlinesEnum ) = 'integer'
'XlOLEVerbEnum' = c(
   'xlVerbOpen' = 2,
   'xlVerbPrimary' = 1 
 )
storage.mode( XlOLEVerbEnum ) = 'integer'
'XlCalculationEnum' = c(
   'xlCalculationAutomatic' = -4105,
   'xlCalculationManual' = -4135,
   'xlCalculationSemiautomatic' = 2 
 )
storage.mode( XlCalculationEnum ) = 'integer'
'XlFileAccessEnum' = c(
   'xlReadOnly' = 3,
   'xlReadWrite' = 2 
 )
storage.mode( XlFileAccessEnum ) = 'integer'
'XlEditionTypeEnum' = c(
   'xlPublisher' = 1,
   'xlSubscriber' = 2 
 )
storage.mode( XlEditionTypeEnum ) = 'integer'
'XlObjectSizeEnum' = c(
   'xlFitToPage' = 2,
   'xlFullPage' = 3,
   'xlScreenSize' = 1 
 )
storage.mode( XlObjectSizeEnum ) = 'integer'
'XlLookAtEnum' = c(
   'xlPart' = 2,
   'xlWhole' = 1 
 )
storage.mode( XlLookAtEnum ) = 'integer'
'XlMailSystemEnum' = c(
   'xlMAPI' = 1,
   'xlNoMailSystem' = 0,
   'xlPowerTalk' = 2 
 )
storage.mode( XlMailSystemEnum ) = 'integer'
'XlLinkInfoTypeEnum' = c(
   'xlLinkInfoOLELinks' = 2,
   'xlLinkInfoPublishers' = 5,
   'xlLinkInfoSubscribers' = 6 
 )
storage.mode( XlLinkInfoTypeEnum ) = 'integer'
'XlCVErrorEnum' = c(
   'xlErrDiv0' = 2007,
   'xlErrNA' = 2042,
   'xlErrName' = 2029,
   'xlErrNull' = 2000,
   'xlErrNum' = 2036,
   'xlErrRef' = 2023,
   'xlErrValue' = 2015 
 )
storage.mode( XlCVErrorEnum ) = 'integer'
'XlEditionFormatEnum' = c(
   'xlBIFF' = 2,
   'xlPICT' = 1,
   'xlRTF' = 4,
   'xlVALU' = 8 
 )
storage.mode( XlEditionFormatEnum ) = 'integer'
'XlLinkEnum' = c(
   'xlExcelLinks' = 1,
   'xlOLELinks' = 2,
   'xlPublishers' = 5,
   'xlSubscribers' = 6 
 )
storage.mode( XlLinkEnum ) = 'integer'
'XlCellTypeEnum' = c(
   'xlCellTypeBlanks' = 4,
   'xlCellTypeConstants' = 2,
   'xlCellTypeFormulas' = -4123,
   'xlCellTypeLastCell' = 11,
   'xlCellTypeComments' = -4144,
   'xlCellTypeVisible' = 12,
   'xlCellTypeAllFormatConditions' = -4172,
   'xlCellTypeSameFormatConditions' = -4173,
   'xlCellTypeAllValidation' = -4174,
   'xlCellTypeSameValidation' = -4175 
 )
storage.mode( XlCellTypeEnum ) = 'integer'
'XlArrangeStyleEnum' = c(
   'xlArrangeStyleCascade' = 7,
   'xlArrangeStyleHorizontal' = -4128,
   'xlArrangeStyleTiled' = 1,
   'xlArrangeStyleVertical' = -4166 
 )
storage.mode( XlArrangeStyleEnum ) = 'integer'
'XlMousePointerEnum' = c(
   'xlIBeam' = 3,
   'xlDefault' = -4143,
   'xlNorthwestArrow' = 1,
   'xlWait' = 2 
 )
storage.mode( XlMousePointerEnum ) = 'integer'
'XlEditionOptionsOptionEnum' = c(
   'xlAutomaticUpdate' = 4,
   'xlCancel' = 1,
   'xlChangeAttributes' = 6,
   'xlManualUpdate' = 5,
   'xlOpenSource' = 3,
   'xlSelect' = 3,
   'xlSendPublisher' = 2,
   'xlUpdateSubscriber' = 2 
 )
storage.mode( XlEditionOptionsOptionEnum ) = 'integer'
'XlAutoFillTypeEnum' = c(
   'xlFillCopy' = 1,
   'xlFillDays' = 5,
   'xlFillDefault' = 0,
   'xlFillFormats' = 3,
   'xlFillMonths' = 7,
   'xlFillSeries' = 2,
   'xlFillValues' = 4,
   'xlFillWeekdays' = 6,
   'xlFillYears' = 8,
   'xlGrowthTrend' = 10,
   'xlLinearTrend' = 9 
 )
storage.mode( XlAutoFillTypeEnum ) = 'integer'
'XlAutoFilterOperatorEnum' = c(
   'xlAnd' = 1,
   'xlBottom10Items' = 4,
   'xlBottom10Percent' = 6,
   'xlOr' = 2,
   'xlTop10Items' = 3,
   'xlTop10Percent' = 5 
 )
storage.mode( XlAutoFilterOperatorEnum ) = 'integer'
'XlClipboardFormatEnum' = c(
   'xlClipboardFormatBIFF' = 8,
   'xlClipboardFormatBIFF2' = 18,
   'xlClipboardFormatBIFF3' = 20,
   'xlClipboardFormatBIFF4' = 30,
   'xlClipboardFormatBinary' = 15,
   'xlClipboardFormatBitmap' = 9,
   'xlClipboardFormatCGM' = 13,
   'xlClipboardFormatCSV' = 5,
   'xlClipboardFormatDIF' = 4,
   'xlClipboardFormatDspText' = 12,
   'xlClipboardFormatEmbeddedObject' = 21,
   'xlClipboardFormatEmbedSource' = 22,
   'xlClipboardFormatLink' = 11,
   'xlClipboardFormatLinkSource' = 23,
   'xlClipboardFormatLinkSourceDesc' = 32,
   'xlClipboardFormatMovie' = 24,
   'xlClipboardFormatNative' = 14,
   'xlClipboardFormatObjectDesc' = 31,
   'xlClipboardFormatObjectLink' = 19,
   'xlClipboardFormatOwnerLink' = 17,
   'xlClipboardFormatPICT' = 2,
   'xlClipboardFormatPrintPICT' = 3,
   'xlClipboardFormatRTF' = 7,
   'xlClipboardFormatScreenPICT' = 29,
   'xlClipboardFormatStandardFont' = 28,
   'xlClipboardFormatStandardScale' = 27,
   'xlClipboardFormatSYLK' = 6,
   'xlClipboardFormatTable' = 16,
   'xlClipboardFormatText' = 0,
   'xlClipboardFormatToolFace' = 25,
   'xlClipboardFormatToolFacePICT' = 26,
   'xlClipboardFormatVALU' = 1,
   'xlClipboardFormatWK1' = 10 
 )
storage.mode( XlClipboardFormatEnum ) = 'integer'
'XlFileFormatEnum' = c(
   'xlAddIn' = 18,
   'xlCSV' = 6,
   'xlCSVMac' = 22,
   'xlCSVMSDOS' = 24,
   'xlCSVWindows' = 23,
   'xlDBF2' = 7,
   'xlDBF3' = 8,
   'xlDBF4' = 11,
   'xlDIF' = 9,
   'xlExcel2' = 16,
   'xlExcel2FarEast' = 27,
   'xlExcel3' = 29,
   'xlExcel4' = 33,
   'xlExcel5' = 39,
   'xlExcel7' = 39,
   'xlExcel9795' = 43,
   'xlExcel4Workbook' = 35,
   'xlIntlAddIn' = 26,
   'xlIntlMacro' = 25,
   'xlWorkbookNormal' = -4143,
   'xlSYLK' = 2,
   'xlTemplate' = 17,
   'xlCurrentPlatformText' = -4158,
   'xlTextMac' = 19,
   'xlTextMSDOS' = 21,
   'xlTextPrinter' = 36,
   'xlTextWindows' = 20,
   'xlWJ2WD1' = 14,
   'xlWK1' = 5,
   'xlWK1ALL' = 31,
   'xlWK1FMT' = 30,
   'xlWK3' = 15,
   'xlWK4' = 38,
   'xlWK3FM3' = 32,
   'xlWKS' = 4,
   'xlWorks2FarEast' = 28,
   'xlWQ1' = 34,
   'xlWJ3' = 40,
   'xlWJ3FJ3' = 41,
   'xlUnicodeText' = 42,
   'xlHtml' = 44,
   'xlWebArchive' = 45,
   'xlXMLSpreadsheet' = 46 
 )
storage.mode( XlFileFormatEnum ) = 'integer'
'XlApplicationInternationalEnum' = c(
   'xl24HourClock' = 33,
   'xl4DigitYears' = 43,
   'xlAlternateArraySeparator' = 16,
   'xlColumnSeparator' = 14,
   'xlCountryCode' = 1,
   'xlCountrySetting' = 2,
   'xlCurrencyBefore' = 37,
   'xlCurrencyCode' = 25,
   'xlCurrencyDigits' = 27,
   'xlCurrencyLeadingZeros' = 40,
   'xlCurrencyMinusSign' = 38,
   'xlCurrencyNegative' = 28,
   'xlCurrencySpaceBefore' = 36,
   'xlCurrencyTrailingZeros' = 39,
   'xlDateOrder' = 32,
   'xlDateSeparator' = 17,
   'xlDayCode' = 21,
   'xlDayLeadingZero' = 42,
   'xlDecimalSeparator' = 3,
   'xlGeneralFormatName' = 26,
   'xlHourCode' = 22,
   'xlLeftBrace' = 12,
   'xlLeftBracket' = 10,
   'xlListSeparator' = 5,
   'xlLowerCaseColumnLetter' = 9,
   'xlLowerCaseRowLetter' = 8,
   'xlMDY' = 44,
   'xlMetric' = 35,
   'xlMinuteCode' = 23,
   'xlMonthCode' = 20,
   'xlMonthLeadingZero' = 41,
   'xlMonthNameChars' = 30,
   'xlNoncurrencyDigits' = 29,
   'xlNonEnglishFunctions' = 34,
   'xlRightBrace' = 13,
   'xlRightBracket' = 11,
   'xlRowSeparator' = 15,
   'xlSecondCode' = 24,
   'xlThousandsSeparator' = 4,
   'xlTimeLeadingZero' = 45,
   'xlTimeSeparator' = 18,
   'xlUpperCaseColumnLetter' = 7,
   'xlUpperCaseRowLetter' = 6,
   'xlWeekdayNameChars' = 31,
   'xlYearCode' = 19 
 )
storage.mode( XlApplicationInternationalEnum ) = 'integer'
'XlPageBreakExtentEnum' = c(
   'xlPageBreakFull' = 1,
   'xlPageBreakPartial' = 2 
 )
storage.mode( XlPageBreakExtentEnum ) = 'integer'
'XlCellInsertionModeEnum' = c(
   'xlOverwriteCells' = 0,
   'xlInsertDeleteCells' = 1,
   'xlInsertEntireRows' = 2 
 )
storage.mode( XlCellInsertionModeEnum ) = 'integer'
'XlFormulaLabelEnum' = c(
   'xlNoLabels' = -4142,
   'xlRowLabels' = 1,
   'xlColumnLabels' = 2,
   'xlMixedLabels' = 3 
 )
storage.mode( XlFormulaLabelEnum ) = 'integer'
'XlHighlightChangesTimeEnum' = c(
   'xlSinceMyLastSave' = 1,
   'xlAllChanges' = 2,
   'xlNotYetReviewed' = 3 
 )
storage.mode( XlHighlightChangesTimeEnum ) = 'integer'
'XlCommentDisplayModeEnum' = c(
   'xlNoIndicator' = 0,
   'xlCommentIndicatorOnly' = -1,
   'xlCommentAndIndicator' = 1 
 )
storage.mode( XlCommentDisplayModeEnum ) = 'integer'
'XlFormatConditionTypeEnum' = c(
   'xlCellValue' = 1,
   'xlExpression' = 2 
 )
storage.mode( XlFormatConditionTypeEnum ) = 'integer'
'XlFormatConditionOperatorEnum' = c(
   'xlBetween' = 1,
   'xlNotBetween' = 2,
   'xlEqual' = 3,
   'xlNotEqual' = 4,
   'xlGreater' = 5,
   'xlLess' = 6,
   'xlGreaterEqual' = 7,
   'xlLessEqual' = 8 
 )
storage.mode( XlFormatConditionOperatorEnum ) = 'integer'
'XlEnableSelectionEnum' = c(
   'xlNoRestrictions' = 0,
   'xlUnlockedCells' = 1,
   'xlNoSelection' = -4142 
 )
storage.mode( XlEnableSelectionEnum ) = 'integer'
'XlDVTypeEnum' = c(
   'xlValidateInputOnly' = 0,
   'xlValidateWholeNumber' = 1,
   'xlValidateDecimal' = 2,
   'xlValidateList' = 3,
   'xlValidateDate' = 4,
   'xlValidateTime' = 5,
   'xlValidateTextLength' = 6,
   'xlValidateCustom' = 7 
 )
storage.mode( XlDVTypeEnum ) = 'integer'
'XlIMEModeEnum' = c(
   'xlIMEModeNoControl' = 0,
   'xlIMEModeOn' = 1,
   'xlIMEModeOff' = 2,
   'xlIMEModeDisable' = 3,
   'xlIMEModeHiragana' = 4,
   'xlIMEModeKatakana' = 5,
   'xlIMEModeKatakanaHalf' = 6,
   'xlIMEModeAlphaFull' = 7,
   'xlIMEModeAlpha' = 8,
   'xlIMEModeHangulFull' = 9,
   'xlIMEModeHangul' = 10 
 )
storage.mode( XlIMEModeEnum ) = 'integer'
'XlDVAlertStyleEnum' = c(
   'xlValidAlertStop' = 1,
   'xlValidAlertWarning' = 2,
   'xlValidAlertInformation' = 3 
 )
storage.mode( XlDVAlertStyleEnum ) = 'integer'
'XlChartLocationEnum' = c(
   'xlLocationAsNewSheet' = 1,
   'xlLocationAsObject' = 2,
   'xlLocationAutomatic' = 3 
 )
storage.mode( XlChartLocationEnum ) = 'integer'
'XlPaperSizeEnum' = c(
   'xlPaper10x14' = 16,
   'xlPaper11x17' = 17,
   'xlPaperA3' = 8,
   'xlPaperA4' = 9,
   'xlPaperA4Small' = 10,
   'xlPaperA5' = 11,
   'xlPaperB4' = 12,
   'xlPaperB5' = 13,
   'xlPaperCsheet' = 24,
   'xlPaperDsheet' = 25,
   'xlPaperEnvelope10' = 20,
   'xlPaperEnvelope11' = 21,
   'xlPaperEnvelope12' = 22,
   'xlPaperEnvelope14' = 23,
   'xlPaperEnvelope9' = 19,
   'xlPaperEnvelopeB4' = 33,
   'xlPaperEnvelopeB5' = 34,
   'xlPaperEnvelopeB6' = 35,
   'xlPaperEnvelopeC3' = 29,
   'xlPaperEnvelopeC4' = 30,
   'xlPaperEnvelopeC5' = 28,
   'xlPaperEnvelopeC6' = 31,
   'xlPaperEnvelopeC65' = 32,
   'xlPaperEnvelopeDL' = 27,
   'xlPaperEnvelopeItaly' = 36,
   'xlPaperEnvelopeMonarch' = 37,
   'xlPaperEnvelopePersonal' = 38,
   'xlPaperEsheet' = 26,
   'xlPaperExecutive' = 7,
   'xlPaperFanfoldLegalGerman' = 41,
   'xlPaperFanfoldStdGerman' = 40,
   'xlPaperFanfoldUS' = 39,
   'xlPaperFolio' = 14,
   'xlPaperLedger' = 4,
   'xlPaperLegal' = 5,
   'xlPaperLetter' = 1,
   'xlPaperLetterSmall' = 2,
   'xlPaperNote' = 18,
   'xlPaperQuarto' = 15,
   'xlPaperStatement' = 6,
   'xlPaperTabloid' = 3,
   'xlPaperUser' = 256 
 )
storage.mode( XlPaperSizeEnum ) = 'integer'
'XlPasteSpecialOperationEnum' = c(
   'xlPasteSpecialOperationAdd' = 2,
   'xlPasteSpecialOperationDivide' = 5,
   'xlPasteSpecialOperationMultiply' = 4,
   'xlPasteSpecialOperationNone' = -4142,
   'xlPasteSpecialOperationSubtract' = 3 
 )
storage.mode( XlPasteSpecialOperationEnum ) = 'integer'
'XlPasteTypeEnum' = c(
   'xlPasteAll' = -4104,
   'xlPasteAllExceptBorders' = 7,
   'xlPasteFormats' = -4122,
   'xlPasteFormulas' = -4123,
   'xlPasteComments' = -4144,
   'xlPasteValues' = -4163,
   'xlPasteColumnWidths' = 8,
   'xlPasteValidation' = 6,
   'xlPasteFormulasAndNumberFormats' = 11,
   'xlPasteValuesAndNumberFormats' = 12 
 )
storage.mode( XlPasteTypeEnum ) = 'integer'
'XlPhoneticCharacterTypeEnum' = c(
   'xlKatakanaHalf' = 0,
   'xlKatakana' = 1,
   'xlHiragana' = 2,
   'xlNoConversion' = 3 
 )
storage.mode( XlPhoneticCharacterTypeEnum ) = 'integer'
'XlPhoneticAlignmentEnum' = c(
   'xlPhoneticAlignNoControl' = 0,
   'xlPhoneticAlignLeft' = 1,
   'xlPhoneticAlignCenter' = 2,
   'xlPhoneticAlignDistributed' = 3 
 )
storage.mode( XlPhoneticAlignmentEnum ) = 'integer'
'XlPictureAppearanceEnum' = c(
   'xlPrinter' = 2,
   'xlScreen' = 1 
 )
storage.mode( XlPictureAppearanceEnum ) = 'integer'
'XlPivotFieldOrientationEnum' = c(
   'xlColumnField' = 2,
   'xlDataField' = 4,
   'xlHidden' = 0,
   'xlPageField' = 3,
   'xlRowField' = 1 
 )
storage.mode( XlPivotFieldOrientationEnum ) = 'integer'
'XlPivotFieldCalculationEnum' = c(
   'xlDifferenceFrom' = 2,
   'xlIndex' = 9,
   'xlNoAdditionalCalculation' = -4143,
   'xlPercentDifferenceFrom' = 4,
   'xlPercentOf' = 3,
   'xlPercentOfColumn' = 7,
   'xlPercentOfRow' = 6,
   'xlPercentOfTotal' = 8,
   'xlRunningTotal' = 5 
 )
storage.mode( XlPivotFieldCalculationEnum ) = 'integer'
'XlPlacementEnum' = c(
   'xlFreeFloating' = 3,
   'xlMove' = 2,
   'xlMoveAndSize' = 1 
 )
storage.mode( XlPlacementEnum ) = 'integer'
'XlPlatformEnum' = c(
   'xlMacintosh' = 1,
   'xlMSDOS' = 3,
   'xlWindows' = 2 
 )
storage.mode( XlPlatformEnum ) = 'integer'
'XlPrintLocationEnum' = c(
   'xlPrintSheetEnd' = 1,
   'xlPrintInPlace' = 16,
   'xlPrintNoComments' = -4142 
 )
storage.mode( XlPrintLocationEnum ) = 'integer'
'XlPriorityEnum' = c(
   'xlPriorityHigh' = -4127,
   'xlPriorityLow' = -4134,
   'xlPriorityNormal' = -4143 
 )
storage.mode( XlPriorityEnum ) = 'integer'
'XlPTSelectionModeEnum' = c(
   'xlLabelOnly' = 1,
   'xlDataAndLabel' = 0,
   'xlDataOnly' = 2,
   'xlOrigin' = 3,
   'xlButton' = 15,
   'xlBlanks' = 4,
   'xlFirstRow' = 256 
 )
storage.mode( XlPTSelectionModeEnum ) = 'integer'
'XlRangeAutoFormatEnum' = c(
   'xlRangeAutoFormat3DEffects1' = 13,
   'xlRangeAutoFormat3DEffects2' = 14,
   'xlRangeAutoFormatAccounting1' = 4,
   'xlRangeAutoFormatAccounting2' = 5,
   'xlRangeAutoFormatAccounting3' = 6,
   'xlRangeAutoFormatAccounting4' = 17,
   'xlRangeAutoFormatClassic1' = 1,
   'xlRangeAutoFormatClassic2' = 2,
   'xlRangeAutoFormatClassic3' = 3,
   'xlRangeAutoFormatColor1' = 7,
   'xlRangeAutoFormatColor2' = 8,
   'xlRangeAutoFormatColor3' = 9,
   'xlRangeAutoFormatList1' = 10,
   'xlRangeAutoFormatList2' = 11,
   'xlRangeAutoFormatList3' = 12,
   'xlRangeAutoFormatLocalFormat1' = 15,
   'xlRangeAutoFormatLocalFormat2' = 16,
   'xlRangeAutoFormatLocalFormat3' = 19,
   'xlRangeAutoFormatLocalFormat4' = 20,
   'xlRangeAutoFormatReport1' = 21,
   'xlRangeAutoFormatReport2' = 22,
   'xlRangeAutoFormatReport3' = 23,
   'xlRangeAutoFormatReport4' = 24,
   'xlRangeAutoFormatReport5' = 25,
   'xlRangeAutoFormatReport6' = 26,
   'xlRangeAutoFormatReport7' = 27,
   'xlRangeAutoFormatReport8' = 28,
   'xlRangeAutoFormatReport9' = 29,
   'xlRangeAutoFormatReport10' = 30,
   'xlRangeAutoFormatClassicPivotTable' = 31,
   'xlRangeAutoFormatTable1' = 32,
   'xlRangeAutoFormatTable2' = 33,
   'xlRangeAutoFormatTable3' = 34,
   'xlRangeAutoFormatTable4' = 35,
   'xlRangeAutoFormatTable5' = 36,
   'xlRangeAutoFormatTable6' = 37,
   'xlRangeAutoFormatTable7' = 38,
   'xlRangeAutoFormatTable8' = 39,
   'xlRangeAutoFormatTable9' = 40,
   'xlRangeAutoFormatTable10' = 41,
   'xlRangeAutoFormatPTNone' = 42,
   'xlRangeAutoFormatNone' = -4142,
   'xlRangeAutoFormatSimple' = -4154 
 )
storage.mode( XlRangeAutoFormatEnum ) = 'integer'
'XlReferenceTypeEnum' = c(
   'xlAbsolute' = 1,
   'xlAbsRowRelColumn' = 2,
   'xlRelative' = 4,
   'xlRelRowAbsColumn' = 3 
 )
storage.mode( XlReferenceTypeEnum ) = 'integer'
'XlLayoutFormTypeEnum' = c(
   'xlTabular' = 0,
   'xlOutline' = 1 
 )
storage.mode( XlLayoutFormTypeEnum ) = 'integer'
'XlRoutingSlipDeliveryEnum' = c(
   'xlAllAtOnce' = 2,
   'xlOneAfterAnother' = 1 
 )
storage.mode( XlRoutingSlipDeliveryEnum ) = 'integer'
'XlRoutingSlipStatusEnum' = c(
   'xlNotYetRouted' = 0,
   'xlRoutingComplete' = 2,
   'xlRoutingInProgress' = 1 
 )
storage.mode( XlRoutingSlipStatusEnum ) = 'integer'
'XlRunAutoMacroEnum' = c(
   'xlAutoActivate' = 3,
   'xlAutoClose' = 2,
   'xlAutoDeactivate' = 4,
   'xlAutoOpen' = 1 
 )
storage.mode( XlRunAutoMacroEnum ) = 'integer'
'XlSaveActionEnum' = c(
   'xlDoNotSaveChanges' = 2,
   'xlSaveChanges' = 1 
 )
storage.mode( XlSaveActionEnum ) = 'integer'
'XlSaveAsAccessModeEnum' = c(
   'xlExclusive' = 3,
   'xlNoChange' = 1,
   'xlShared' = 2 
 )
storage.mode( XlSaveAsAccessModeEnum ) = 'integer'
'XlSaveConflictResolutionEnum' = c(
   'xlLocalSessionChanges' = 2,
   'xlOtherSessionChanges' = 3,
   'xlUserResolution' = 1 
 )
storage.mode( XlSaveConflictResolutionEnum ) = 'integer'
'XlSearchDirectionEnum' = c(
   'xlNext' = 1,
   'xlPrevious' = 2 
 )
storage.mode( XlSearchDirectionEnum ) = 'integer'
'XlSearchOrderEnum' = c(
   'xlByColumns' = 2,
   'xlByRows' = 1 
 )
storage.mode( XlSearchOrderEnum ) = 'integer'
'XlSheetVisibilityEnum' = c(
   'xlSheetVisible' = -1,
   'xlSheetHidden' = 0,
   'xlSheetVeryHidden' = 2 
 )
storage.mode( XlSheetVisibilityEnum ) = 'integer'
'XlSortMethodEnum' = c(
   'xlPinYin' = 1,
   'xlStroke' = 2 
 )
storage.mode( XlSortMethodEnum ) = 'integer'
'XlSortMethodOldEnum' = c(
   'xlCodePage' = 2,
   'xlSyllabary' = 1 
 )
storage.mode( XlSortMethodOldEnum ) = 'integer'
'XlSortOrderEnum' = c(
   'xlAscending' = 1,
   'xlDescending' = 2 
 )
storage.mode( XlSortOrderEnum ) = 'integer'
'XlSortOrientationEnum' = c(
   'xlSortRows' = 2,
   'xlSortColumns' = 1 
 )
storage.mode( XlSortOrientationEnum ) = 'integer'
'XlSortTypeEnum' = c(
   'xlSortLabels' = 2,
   'xlSortValues' = 1 
 )
storage.mode( XlSortTypeEnum ) = 'integer'
'XlSpecialCellsValueEnum' = c(
   'xlErrors' = 16,
   'xlLogical' = 4,
   'xlNumbers' = 1,
   'xlTextValues' = 2 
 )
storage.mode( XlSpecialCellsValueEnum ) = 'integer'
'XlSubscribeToFormatEnum' = c(
   'xlSubscribeToPicture' = -4147,
   'xlSubscribeToText' = -4158 
 )
storage.mode( XlSubscribeToFormatEnum ) = 'integer'
'XlSummaryRowEnum' = c(
   'xlSummaryAbove' = 0,
   'xlSummaryBelow' = 1 
 )
storage.mode( XlSummaryRowEnum ) = 'integer'
'XlSummaryColumnEnum' = c(
   'xlSummaryOnLeft' = -4131,
   'xlSummaryOnRight' = -4152 
 )
storage.mode( XlSummaryColumnEnum ) = 'integer'
'XlSummaryReportTypeEnum' = c(
   'xlSummaryPivotTable' = -4148,
   'xlStandardSummary' = 1 
 )
storage.mode( XlSummaryReportTypeEnum ) = 'integer'
'XlTabPositionEnum' = c(
   'xlTabPositionFirst' = 0,
   'xlTabPositionLast' = 1 
 )
storage.mode( XlTabPositionEnum ) = 'integer'
'XlTextParsingTypeEnum' = c(
   'xlDelimited' = 1,
   'xlFixedWidth' = 2 
 )
storage.mode( XlTextParsingTypeEnum ) = 'integer'
'XlTextQualifierEnum' = c(
   'xlTextQualifierDoubleQuote' = 1,
   'xlTextQualifierNone' = -4142,
   'xlTextQualifierSingleQuote' = 2 
 )
storage.mode( XlTextQualifierEnum ) = 'integer'
'XlWBATemplateEnum' = c(
   'xlWBATChart' = -4109,
   'xlWBATExcel4IntlMacroSheet' = 4,
   'xlWBATExcel4MacroSheet' = 3,
   'xlWBATWorksheet' = -4167 
 )
storage.mode( XlWBATemplateEnum ) = 'integer'
'XlWindowViewEnum' = c(
   'xlNormalView' = 1,
   'xlPageBreakPreview' = 2 
 )
storage.mode( XlWindowViewEnum ) = 'integer'
'XlXLMMacroTypeEnum' = c(
   'xlCommand' = 2,
   'xlFunction' = 1,
   'xlNotXLM' = 3 
 )
storage.mode( XlXLMMacroTypeEnum ) = 'integer'
'XlYesNoGuessEnum' = c(
   'xlGuess' = 0,
   'xlNo' = 2,
   'xlYes' = 1 
 )
storage.mode( XlYesNoGuessEnum ) = 'integer'
'XlBordersIndexEnum' = c(
   'xlInsideHorizontal' = 12,
   'xlInsideVertical' = 11,
   'xlDiagonalDown' = 5,
   'xlDiagonalUp' = 6,
   'xlEdgeBottom' = 9,
   'xlEdgeLeft' = 7,
   'xlEdgeRight' = 10,
   'xlEdgeTop' = 8 
 )
storage.mode( XlBordersIndexEnum ) = 'integer'
'XlToolbarProtectionEnum' = c(
   'xlNoButtonChanges' = 1,
   'xlNoChanges' = 4,
   'xlNoDockingChanges' = 3,
   'xlToolbarProtectionNone' = -4143,
   'xlNoShapeChanges' = 2 
 )
storage.mode( XlToolbarProtectionEnum ) = 'integer'
'XlBuiltInDialogEnum' = c(
   'xlDialogOpen' = 1,
   'xlDialogOpenLinks' = 2,
   'xlDialogSaveAs' = 5,
   'xlDialogFileDelete' = 6,
   'xlDialogPageSetup' = 7,
   'xlDialogPrint' = 8,
   'xlDialogPrinterSetup' = 9,
   'xlDialogArrangeAll' = 12,
   'xlDialogWindowSize' = 13,
   'xlDialogWindowMove' = 14,
   'xlDialogRun' = 17,
   'xlDialogSetPrintTitles' = 23,
   'xlDialogFont' = 26,
   'xlDialogDisplay' = 27,
   'xlDialogProtectDocument' = 28,
   'xlDialogCalculation' = 32,
   'xlDialogExtract' = 35,
   'xlDialogDataDelete' = 36,
   'xlDialogSort' = 39,
   'xlDialogDataSeries' = 40,
   'xlDialogTable' = 41,
   'xlDialogFormatNumber' = 42,
   'xlDialogAlignment' = 43,
   'xlDialogStyle' = 44,
   'xlDialogBorder' = 45,
   'xlDialogCellProtection' = 46,
   'xlDialogColumnWidth' = 47,
   'xlDialogClear' = 52,
   'xlDialogPasteSpecial' = 53,
   'xlDialogEditDelete' = 54,
   'xlDialogInsert' = 55,
   'xlDialogPasteNames' = 58,
   'xlDialogDefineName' = 61,
   'xlDialogCreateNames' = 62,
   'xlDialogFormulaGoto' = 63,
   'xlDialogFormulaFind' = 64,
   'xlDialogGalleryArea' = 67,
   'xlDialogGalleryBar' = 68,
   'xlDialogGalleryColumn' = 69,
   'xlDialogGalleryLine' = 70,
   'xlDialogGalleryPie' = 71,
   'xlDialogGalleryScatter' = 72,
   'xlDialogCombination' = 73,
   'xlDialogGridlines' = 76,
   'xlDialogAxes' = 78,
   'xlDialogAttachText' = 80,
   'xlDialogPatterns' = 84,
   'xlDialogMainChart' = 85,
   'xlDialogOverlay' = 86,
   'xlDialogScale' = 87,
   'xlDialogFormatLegend' = 88,
   'xlDialogFormatText' = 89,
   'xlDialogParse' = 91,
   'xlDialogUnhide' = 94,
   'xlDialogWorkspace' = 95,
   'xlDialogActivate' = 103,
   'xlDialogCopyPicture' = 108,
   'xlDialogDeleteName' = 110,
   'xlDialogDeleteFormat' = 111,
   'xlDialogNew' = 119,
   'xlDialogRowHeight' = 127,
   'xlDialogFormatMove' = 128,
   'xlDialogFormatSize' = 129,
   'xlDialogFormulaReplace' = 130,
   'xlDialogSelectSpecial' = 132,
   'xlDialogApplyNames' = 133,
   'xlDialogReplaceFont' = 134,
   'xlDialogSplit' = 137,
   'xlDialogOutline' = 142,
   'xlDialogSaveWorkbook' = 145,
   'xlDialogCopyChart' = 147,
   'xlDialogFormatFont' = 150,
   'xlDialogNote' = 154,
   'xlDialogSetUpdateStatus' = 159,
   'xlDialogColorPalette' = 161,
   'xlDialogChangeLink' = 166,
   'xlDialogAppMove' = 170,
   'xlDialogAppSize' = 171,
   'xlDialogMainChartType' = 185,
   'xlDialogOverlayChartType' = 186,
   'xlDialogOpenMail' = 188,
   'xlDialogSendMail' = 189,
   'xlDialogStandardFont' = 190,
   'xlDialogConsolidate' = 191,
   'xlDialogSortSpecial' = 192,
   'xlDialogGallery3dArea' = 193,
   'xlDialogGallery3dColumn' = 194,
   'xlDialogGallery3dLine' = 195,
   'xlDialogGallery3dPie' = 196,
   'xlDialogView3d' = 197,
   'xlDialogGoalSeek' = 198,
   'xlDialogWorkgroup' = 199,
   'xlDialogFillGroup' = 200,
   'xlDialogUpdateLink' = 201,
   'xlDialogPromote' = 202,
   'xlDialogDemote' = 203,
   'xlDialogShowDetail' = 204,
   'xlDialogObjectProperties' = 207,
   'xlDialogSaveNewObject' = 208,
   'xlDialogApplyStyle' = 212,
   'xlDialogAssignToObject' = 213,
   'xlDialogObjectProtection' = 214,
   'xlDialogCreatePublisher' = 217,
   'xlDialogSubscribeTo' = 218,
   'xlDialogShowToolbar' = 220,
   'xlDialogPrintPreview' = 222,
   'xlDialogEditColor' = 223,
   'xlDialogFormatMain' = 225,
   'xlDialogFormatOverlay' = 226,
   'xlDialogEditSeries' = 228,
   'xlDialogDefineStyle' = 229,
   'xlDialogGalleryRadar' = 249,
   'xlDialogEditionOptions' = 251,
   'xlDialogZoom' = 256,
   'xlDialogInsertObject' = 259,
   'xlDialogSize' = 261,
   'xlDialogMove' = 262,
   'xlDialogFormatAuto' = 269,
   'xlDialogGallery3dBar' = 272,
   'xlDialogGallery3dSurface' = 273,
   'xlDialogCustomizeToolbar' = 276,
   'xlDialogWorkbookAdd' = 281,
   'xlDialogWorkbookMove' = 282,
   'xlDialogWorkbookCopy' = 283,
   'xlDialogWorkbookOptions' = 284,
   'xlDialogSaveWorkspace' = 285,
   'xlDialogChartWizard' = 288,
   'xlDialogAssignToTool' = 293,
   'xlDialogPlacement' = 300,
   'xlDialogFillWorkgroup' = 301,
   'xlDialogWorkbookNew' = 302,
   'xlDialogScenarioCells' = 305,
   'xlDialogScenarioAdd' = 307,
   'xlDialogScenarioEdit' = 308,
   'xlDialogScenarioSummary' = 311,
   'xlDialogPivotTableWizard' = 312,
   'xlDialogPivotFieldProperties' = 313,
   'xlDialogOptionsCalculation' = 318,
   'xlDialogOptionsEdit' = 319,
   'xlDialogOptionsView' = 320,
   'xlDialogAddinManager' = 321,
   'xlDialogMenuEditor' = 322,
   'xlDialogAttachToolbars' = 323,
   'xlDialogOptionsChart' = 325,
   'xlDialogVbaInsertFile' = 328,
   'xlDialogVbaProcedureDefinition' = 330,
   'xlDialogRoutingSlip' = 336,
   'xlDialogMailLogon' = 339,
   'xlDialogInsertPicture' = 342,
   'xlDialogGalleryDoughnut' = 344,
   'xlDialogChartTrend' = 350,
   'xlDialogWorkbookInsert' = 354,
   'xlDialogOptionsTransition' = 355,
   'xlDialogOptionsGeneral' = 356,
   'xlDialogFilterAdvanced' = 370,
   'xlDialogMailNextLetter' = 378,
   'xlDialogDataLabel' = 379,
   'xlDialogInsertTitle' = 380,
   'xlDialogFontProperties' = 381,
   'xlDialogMacroOptions' = 382,
   'xlDialogWorkbookUnhide' = 384,
   'xlDialogWorkbookName' = 386,
   'xlDialogGalleryCustom' = 388,
   'xlDialogAddChartAutoformat' = 390,
   'xlDialogChartAddData' = 392,
   'xlDialogTabOrder' = 394,
   'xlDialogSubtotalCreate' = 398,
   'xlDialogWorkbookTabSplit' = 415,
   'xlDialogWorkbookProtect' = 417,
   'xlDialogScrollbarProperties' = 420,
   'xlDialogPivotShowPages' = 421,
   'xlDialogTextToColumns' = 422,
   'xlDialogFormatCharttype' = 423,
   'xlDialogPivotFieldGroup' = 433,
   'xlDialogPivotFieldUngroup' = 434,
   'xlDialogCheckboxProperties' = 435,
   'xlDialogLabelProperties' = 436,
   'xlDialogListboxProperties' = 437,
   'xlDialogEditboxProperties' = 438,
   'xlDialogOpenText' = 441,
   'xlDialogPushbuttonProperties' = 445,
   'xlDialogFilter' = 447,
   'xlDialogFunctionWizard' = 450,
   'xlDialogSaveCopyAs' = 456,
   'xlDialogOptionsListsAdd' = 458,
   'xlDialogSeriesAxes' = 460,
   'xlDialogSeriesX' = 461,
   'xlDialogSeriesY' = 462,
   'xlDialogErrorbarX' = 463,
   'xlDialogErrorbarY' = 464,
   'xlDialogFormatChart' = 465,
   'xlDialogSeriesOrder' = 466,
   'xlDialogMailEditMailer' = 470,
   'xlDialogStandardWidth' = 472,
   'xlDialogScenarioMerge' = 473,
   'xlDialogProperties' = 474,
   'xlDialogSummaryInfo' = 474,
   'xlDialogFindFile' = 475,
   'xlDialogActiveCellFont' = 476,
   'xlDialogVbaMakeAddin' = 478,
   'xlDialogFileSharing' = 481,
   'xlDialogAutoCorrect' = 485,
   'xlDialogCustomViews' = 493,
   'xlDialogInsertNameLabel' = 496,
   'xlDialogSeriesShape' = 504,
   'xlDialogChartOptionsDataLabels' = 505,
   'xlDialogChartOptionsDataTable' = 506,
   'xlDialogSetBackgroundPicture' = 509,
   'xlDialogDataValidation' = 525,
   'xlDialogChartType' = 526,
   'xlDialogChartLocation' = 527,
   '_xlDialogPhonetic' = 538,
   'xlDialogChartSourceData' = 540,
   '_xlDialogChartSourceData' = 541,
   'xlDialogSeriesOptions' = 557,
   'xlDialogPivotTableOptions' = 567,
   'xlDialogPivotSolveOrder' = 568,
   'xlDialogPivotCalculatedField' = 570,
   'xlDialogPivotCalculatedItem' = 572,
   'xlDialogConditionalFormatting' = 583,
   'xlDialogInsertHyperlink' = 596,
   'xlDialogProtectSharing' = 620,
   'xlDialogOptionsME' = 647,
   'xlDialogPublishAsWebPage' = 653,
   'xlDialogPhonetic' = 656,
   'xlDialogNewWebQuery' = 667,
   'xlDialogImportTextFile' = 666,
   'xlDialogExternalDataProperties' = 530,
   'xlDialogWebOptionsGeneral' = 683,
   'xlDialogWebOptionsFiles' = 684,
   'xlDialogWebOptionsPictures' = 685,
   'xlDialogWebOptionsEncoding' = 686,
   'xlDialogWebOptionsFonts' = 687,
   'xlDialogPivotClientServerSet' = 689,
   'xlDialogPropertyFields' = 754,
   'xlDialogSearch' = 731,
   'xlDialogEvaluateFormula' = 709,
   'xlDialogDataLabelMultiple' = 723,
   'xlDialogChartOptionsDataLabelMultiple' = 724,
   'xlDialogErrorChecking' = 732,
   'xlDialogWebOptionsBrowsers' = 773,
   'xlDialogCreateList' = 796,
   'xlDialogPermission' = 832,
   'xlDialogMyPermission' = 834 
 )
storage.mode( XlBuiltInDialogEnum ) = 'integer'
'XlParameterTypeEnum' = c(
   'xlPrompt' = 0,
   'xlConstant' = 1,
   'xlRange' = 2 
 )
storage.mode( XlParameterTypeEnum ) = 'integer'
'XlParameterDataTypeEnum' = c(
   'xlParamTypeUnknown' = 0,
   'xlParamTypeChar' = 1,
   'xlParamTypeNumeric' = 2,
   'xlParamTypeDecimal' = 3,
   'xlParamTypeInteger' = 4,
   'xlParamTypeSmallInt' = 5,
   'xlParamTypeFloat' = 6,
   'xlParamTypeReal' = 7,
   'xlParamTypeDouble' = 8,
   'xlParamTypeVarChar' = 12,
   'xlParamTypeDate' = 9,
   'xlParamTypeTime' = 10,
   'xlParamTypeTimestamp' = 11,
   'xlParamTypeLongVarChar' = -1,
   'xlParamTypeBinary' = -2,
   'xlParamTypeVarBinary' = -3,
   'xlParamTypeLongVarBinary' = -4,
   'xlParamTypeBigInt' = -5,
   'xlParamTypeTinyInt' = -6,
   'xlParamTypeBit' = -7,
   'xlParamTypeWChar' = -8 
 )
storage.mode( XlParameterDataTypeEnum ) = 'integer'
'XlFormControlEnum' = c(
   'xlButtonControl' = 0,
   'xlCheckBox' = 1,
   'xlDropDown' = 2,
   'xlEditBox' = 3,
   'xlGroupBox' = 4,
   'xlLabel' = 5,
   'xlListBox' = 6,
   'xlOptionButton' = 7,
   'xlScrollBar' = 8,
   'xlSpinner' = 9 
 )
storage.mode( XlFormControlEnum ) = 'integer'
'XlSourceTypeEnum' = c(
   'xlSourceWorkbook' = 0,
   'xlSourceSheet' = 1,
   'xlSourcePrintArea' = 2,
   'xlSourceAutoFilter' = 3,
   'xlSourceRange' = 4,
   'xlSourceChart' = 5,
   'xlSourcePivotTable' = 6,
   'xlSourceQuery' = 7 
 )
storage.mode( XlSourceTypeEnum ) = 'integer'
'XlHtmlTypeEnum' = c(
   'xlHtmlStatic' = 0,
   'xlHtmlCalc' = 1,
   'xlHtmlList' = 2,
   'xlHtmlChart' = 3 
 )
storage.mode( XlHtmlTypeEnum ) = 'integer'
'XlPivotFormatTypeEnum' = c(
   'xlReport1' = 0,
   'xlReport2' = 1,
   'xlReport3' = 2,
   'xlReport4' = 3,
   'xlReport5' = 4,
   'xlReport6' = 5,
   'xlReport7' = 6,
   'xlReport8' = 7,
   'xlReport9' = 8,
   'xlReport10' = 9,
   'xlTable1' = 10,
   'xlTable2' = 11,
   'xlTable3' = 12,
   'xlTable4' = 13,
   'xlTable5' = 14,
   'xlTable6' = 15,
   'xlTable7' = 16,
   'xlTable8' = 17,
   'xlTable9' = 18,
   'xlTable10' = 19,
   'xlPTClassic' = 20,
   'xlPTNone' = 21 
 )
storage.mode( XlPivotFormatTypeEnum ) = 'integer'
'XlCmdTypeEnum' = c(
   'xlCmdCube' = 1,
   'xlCmdSql' = 2,
   'xlCmdTable' = 3,
   'xlCmdDefault' = 4,
   'xlCmdList' = 5 
 )
storage.mode( XlCmdTypeEnum ) = 'integer'
'XlColumnDataTypeEnum' = c(
   'xlGeneralFormat' = 1,
   'xlTextFormat' = 2,
   'xlMDYFormat' = 3,
   'xlDMYFormat' = 4,
   'xlYMDFormat' = 5,
   'xlMYDFormat' = 6,
   'xlDYMFormat' = 7,
   'xlYDMFormat' = 8,
   'xlSkipColumn' = 9,
   'xlEMDFormat' = 10 
 )
storage.mode( XlColumnDataTypeEnum ) = 'integer'
'XlQueryTypeEnum' = c(
   'xlODBCQuery' = 1,
   'xlDAORecordset' = 2,
   'xlWebQuery' = 4,
   'xlOLEDBQuery' = 5,
   'xlTextImport' = 6,
   'xlADORecordset' = 7 
 )
storage.mode( XlQueryTypeEnum ) = 'integer'
'XlWebSelectionTypeEnum' = c(
   'xlEntirePage' = 1,
   'xlAllTables' = 2,
   'xlSpecifiedTables' = 3 
 )
storage.mode( XlWebSelectionTypeEnum ) = 'integer'
'XlCubeFieldTypeEnum' = c(
   'xlHierarchy' = 1,
   'xlMeasure' = 2,
   'xlSet' = 3 
 )
storage.mode( XlCubeFieldTypeEnum ) = 'integer'
'XlWebFormattingEnum' = c(
   'xlWebFormattingAll' = 1,
   'xlWebFormattingRTF' = 2,
   'xlWebFormattingNone' = 3 
 )
storage.mode( XlWebFormattingEnum ) = 'integer'
'XlDisplayDrawingObjectsEnum' = c(
   'xlDisplayShapes' = -4104,
   'xlHide' = 3,
   'xlPlaceholders' = 2 
 )
storage.mode( XlDisplayDrawingObjectsEnum ) = 'integer'
'XlSubtototalLocationTypeEnum' = c(
   'xlAtTop' = 1,
   'xlAtBottom' = 2 
 )
storage.mode( XlSubtototalLocationTypeEnum ) = 'integer'
'XlPivotTableVersionListEnum' = c(
   'xlPivotTableVersion2000' = 0,
   'xlPivotTableVersion10' = 1,
   'xlPivotTableVersionCurrent' = -1 
 )
storage.mode( XlPivotTableVersionListEnum ) = 'integer'
'XlPrintErrorsEnum' = c(
   'xlPrintErrorsDisplayed' = 0,
   'xlPrintErrorsBlank' = 1,
   'xlPrintErrorsDash' = 2,
   'xlPrintErrorsNA' = 3 
 )
storage.mode( XlPrintErrorsEnum ) = 'integer'
'XlPivotCellTypeEnum' = c(
   'xlPivotCellValue' = 0,
   'xlPivotCellPivotItem' = 1,
   'xlPivotCellSubtotal' = 2,
   'xlPivotCellGrandTotal' = 3,
   'xlPivotCellDataField' = 4,
   'xlPivotCellPivotField' = 5,
   'xlPivotCellPageFieldItem' = 6,
   'xlPivotCellCustomSubtotal' = 7,
   'xlPivotCellDataPivotField' = 8,
   'xlPivotCellBlankCell' = 9 
 )
storage.mode( XlPivotCellTypeEnum ) = 'integer'
'XlPivotTableMissingItemsEnum' = c(
   'xlMissingItemsDefault' = -1,
   'xlMissingItemsNone' = 0,
   'xlMissingItemsMax' = 32500 
 )
storage.mode( XlPivotTableMissingItemsEnum ) = 'integer'
'XlCalculationStateEnum' = c(
   'xlDone' = 0,
   'xlCalculating' = 1,
   'xlPending' = 2 
 )
storage.mode( XlCalculationStateEnum ) = 'integer'
'XlCalculationInterruptKeyEnum' = c(
   'xlNoKey' = 0,
   'xlEscKey' = 1,
   'xlAnyKey' = 2 
 )
storage.mode( XlCalculationInterruptKeyEnum ) = 'integer'
'XlSortDataOptionEnum' = c(
   'xlSortNormal' = 0,
   'xlSortTextAsNumbers' = 1 
 )
storage.mode( XlSortDataOptionEnum ) = 'integer'
'XlUpdateLinksEnum' = c(
   'xlUpdateLinksUserSetting' = 1,
   'xlUpdateLinksNever' = 2,
   'xlUpdateLinksAlways' = 3 
 )
storage.mode( XlUpdateLinksEnum ) = 'integer'
'XlLinkStatusEnum' = c(
   'xlLinkStatusOK' = 0,
   'xlLinkStatusMissingFile' = 1,
   'xlLinkStatusMissingSheet' = 2,
   'xlLinkStatusOld' = 3,
   'xlLinkStatusSourceNotCalculated' = 4,
   'xlLinkStatusIndeterminate' = 5,
   'xlLinkStatusNotStarted' = 6,
   'xlLinkStatusInvalidName' = 7,
   'xlLinkStatusSourceNotOpen' = 8,
   'xlLinkStatusSourceOpen' = 9,
   'xlLinkStatusCopiedValues' = 10 
 )
storage.mode( XlLinkStatusEnum ) = 'integer'
'XlSearchWithinEnum' = c(
   'xlWithinSheet' = 1,
   'xlWithinWorkbook' = 2 
 )
storage.mode( XlSearchWithinEnum ) = 'integer'
'XlCorruptLoadEnum' = c(
   'xlNormalLoad' = 0,
   'xlRepairFile' = 1,
   'xlExtractData' = 2 
 )
storage.mode( XlCorruptLoadEnum ) = 'integer'
'XlRobustConnectEnum' = c(
   'xlAsRequired' = 0,
   'xlAlways' = 1,
   'xlNever' = 2 
 )
storage.mode( XlRobustConnectEnum ) = 'integer'
'XlErrorChecksEnum' = c(
   'xlEvaluateToError' = 1,
   'xlTextDate' = 2,
   'xlNumberAsText' = 3,
   'xlInconsistentFormula' = 4,
   'xlOmittedCells' = 5,
   'xlUnlockedFormulaCells' = 6,
   'xlEmptyCellReferences' = 7,
   'xlListDataValidation' = 8 
 )
storage.mode( XlErrorChecksEnum ) = 'integer'
'XlDataLabelSeparatorEnum' = c(
   'xlDataLabelSeparatorDefault' = 1 
 )
storage.mode( XlDataLabelSeparatorEnum ) = 'integer'
'XlSmartTagDisplayModeEnum' = c(
   'xlIndicatorAndButton' = 0,
   'xlDisplayNone' = 1,
   'xlButtonOnly' = 2 
 )
storage.mode( XlSmartTagDisplayModeEnum ) = 'integer'
'XlRangeValueDataTypeEnum' = c(
   'xlRangeValueDefault' = 10,
   'xlRangeValueXMLSpreadsheet' = 11,
   'xlRangeValueMSPersistXML' = 12 
 )
storage.mode( XlRangeValueDataTypeEnum ) = 'integer'
'XlSpeakDirectionEnum' = c(
   'xlSpeakByRows' = 0,
   'xlSpeakByColumns' = 1 
 )
storage.mode( XlSpeakDirectionEnum ) = 'integer'
'XlInsertFormatOriginEnum' = c(
   'xlFormatFromLeftOrAbove' = 0,
   'xlFormatFromRightOrBelow' = 1 
 )
storage.mode( XlInsertFormatOriginEnum ) = 'integer'
'XlArabicModesEnum' = c(
   'xlArabicNone' = 0,
   'xlArabicStrictAlefHamza' = 1,
   'xlArabicStrictFinalYaa' = 2,
   'xlArabicBothStrict' = 3 
 )
storage.mode( XlArabicModesEnum ) = 'integer'
'XlImportDataAsEnum' = c(
   'xlQueryTable' = 0,
   'xlPivotTableReport' = 1 
 )
storage.mode( XlImportDataAsEnum ) = 'integer'
'XlCalculatedMemberTypeEnum' = c(
   'xlCalculatedMember' = 0,
   'xlCalculatedSet' = 1 
 )
storage.mode( XlCalculatedMemberTypeEnum ) = 'integer'
'XlHebrewModesEnum' = c(
   'xlHebrewFullScript' = 0,
   'xlHebrewPartialScript' = 1,
   'xlHebrewMixedScript' = 2,
   'xlHebrewMixedAuthorizedScript' = 3 
 )
storage.mode( XlHebrewModesEnum ) = 'integer'
'XlListObjectSourceTypeEnum' = c(
   'xlSrcExternal' = 0,
   'xlSrcRange' = 1,
   'xlSrcXml' = 2 
 )
storage.mode( XlListObjectSourceTypeEnum ) = 'integer'
'XlTextVisualLayoutTypeEnum' = c(
   'xlTextVisualLTR' = 1,
   'xlTextVisualRTL' = 2 
 )
storage.mode( XlTextVisualLayoutTypeEnum ) = 'integer'
'XlListDataTypeEnum' = c(
   'xlListDataTypeNone' = 0,
   'xlListDataTypeText' = 1,
   'xlListDataTypeMultiLineText' = 2,
   'xlListDataTypeNumber' = 3,
   'xlListDataTypeCurrency' = 4,
   'xlListDataTypeDateTime' = 5,
   'xlListDataTypeChoice' = 6,
   'xlListDataTypeChoiceMulti' = 7,
   'xlListDataTypeListLookup' = 8,
   'xlListDataTypeCheckbox' = 9,
   'xlListDataTypeHyperLink' = 10,
   'xlListDataTypeCounter' = 11,
   'xlListDataTypeMultiLineRichText' = 12 
 )
storage.mode( XlListDataTypeEnum ) = 'integer'
'XlTotalsCalculationEnum' = c(
   'xlTotalsCalculationNone' = 0,
   'xlTotalsCalculationSum' = 1,
   'xlTotalsCalculationAverage' = 2,
   'xlTotalsCalculationCount' = 3,
   'xlTotalsCalculationCountNums' = 4,
   'xlTotalsCalculationMin' = 5,
   'xlTotalsCalculationMax' = 6,
   'xlTotalsCalculationStdDev' = 7,
   'xlTotalsCalculationVar' = 8 
 )
storage.mode( XlTotalsCalculationEnum ) = 'integer'
'XlXmlLoadOptionEnum' = c(
   'xlXmlLoadPromptUser' = 0,
   'xlXmlLoadOpenXml' = 1,
   'xlXmlLoadImportToList' = 2,
   'xlXmlLoadMapXml' = 3 
 )
storage.mode( XlXmlLoadOptionEnum ) = 'integer'
'XlSmartTagControlTypeEnum' = c(
   'xlSmartTagControlSmartTag' = 1,
   'xlSmartTagControlLink' = 2,
   'xlSmartTagControlHelp' = 3,
   'xlSmartTagControlHelpURL' = 4,
   'xlSmartTagControlSeparator' = 5,
   'xlSmartTagControlButton' = 6,
   'xlSmartTagControlLabel' = 7,
   'xlSmartTagControlImage' = 8,
   'xlSmartTagControlCheckbox' = 9,
   'xlSmartTagControlTextbox' = 10,
   'xlSmartTagControlListbox' = 11,
   'xlSmartTagControlCombo' = 12,
   'xlSmartTagControlActiveX' = 13,
   'xlSmartTagControlRadioGroup' = 14 
 )
storage.mode( XlSmartTagControlTypeEnum ) = 'integer'
'XlListConflictEnum' = c(
   'xlListConflictDialog' = 0,
   'xlListConflictRetryAllConflicts' = 1,
   'xlListConflictDiscardAllConflicts' = 2,
   'xlListConflictError' = 3 
 )
storage.mode( XlListConflictEnum ) = 'integer'
'XlXmlExportResultEnum' = c(
   'xlXmlExportSuccess' = 0,
   'xlXmlExportValidationFailed' = 1 
 )
storage.mode( XlXmlExportResultEnum ) = 'integer'
'XlXmlImportResultEnum' = c(
   'xlXmlImportSuccess' = 0,
   'xlXmlImportElementsTruncated' = 1,
   'xlXmlImportValidationFailed' = 2 
 )
storage.mode( XlXmlImportResultEnum ) = 'integer'
'COM._Application.GetProperty'  = list('Application' = function() {
		 ans = .COM(x, 'Application', .dispatch = as.integer(2))
		 as(ans, 'Application')
},
'Creator' = function() {
		 ans = .COM(x, 'Creator', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'Parent' = function() {
		 ans = .COM(x, 'Parent', .dispatch = as.integer(2))
		 as(ans, 'Application')
},
'ActiveCell' = function() {
		 ans = .COM(x, 'ActiveCell', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'ActiveChart' = function() {
		 ans = .COM(x, 'ActiveChart', .dispatch = as.integer(2))
		 as(ans, 'Chart')
},
'ActiveDialog' = function() {
		 ans = .COM(x, 'ActiveDialog', .dispatch = as.integer(2))
		 as(ans, 'DialogSheet')
},
'ActiveMenuBar' = function() {
		 ans = .COM(x, 'ActiveMenuBar', .dispatch = as.integer(2))
		 as(ans, 'MenuBar')
},
'ActivePrinter' = function() {
		 ans = .COM(x, 'ActivePrinter', .dispatch = as.integer(2))
	ans
},
'ActiveSheet' = function() {
		 ans = .COM(x, 'ActiveSheet', .dispatch = as.integer(2))
		 ans
},
'ActiveWindow' = function() {
		 ans = .COM(x, 'ActiveWindow', .dispatch = as.integer(2))
		 as(ans, 'Window')
},
'ActiveWorkbook' = function() {
		 ans = .COM(x, 'ActiveWorkbook', .dispatch = as.integer(2))
		 as(ans, 'Workbook')
},
'AddIns' = function() {
		 ans = .COM(x, 'AddIns', .dispatch = as.integer(2))
		 as(ans, 'AddIns')
},
'Assistant' = function() {
		 ans = .COM(x, 'Assistant', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'Cells' = function() {
		 ans = .COM(x, 'Cells', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'Charts' = function() {
		 ans = .COM(x, 'Charts', .dispatch = as.integer(2))
		 as(ans, 'Sheets')
},
'Columns' = function() {
		 ans = .COM(x, 'Columns', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'CommandBars' = function() {
		 ans = .COM(x, 'CommandBars', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'DDEAppReturnCode' = function() {
		 ans = .COM(x, 'DDEAppReturnCode', .dispatch = as.integer(2))
	ans
},
'DialogSheets' = function() {
		 ans = .COM(x, 'DialogSheets', .dispatch = as.integer(2))
		 as(ans, 'Sheets')
},
'MenuBars' = function() {
		 ans = .COM(x, 'MenuBars', .dispatch = as.integer(2))
		 as(ans, 'MenuBars')
},
'Modules' = function() {
		 ans = .COM(x, 'Modules', .dispatch = as.integer(2))
		 as(ans, 'Modules')
},
'Names' = function() {
		 ans = .COM(x, 'Names', .dispatch = as.integer(2))
		 as(ans, 'Names')
},
'Rows' = function() {
		 ans = .COM(x, 'Rows', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'Selection' = function() {
		 ans = .COM(x, 'Selection', .dispatch = as.integer(2))
		 ans
},
'Sheets' = function() {
		 ans = .COM(x, 'Sheets', .dispatch = as.integer(2))
		 as(ans, 'Sheets')
},
'ThisWorkbook' = function() {
		 ans = .COM(x, 'ThisWorkbook', .dispatch = as.integer(2))
		 as(ans, 'Workbook')
},
'Toolbars' = function() {
		 ans = .COM(x, 'Toolbars', .dispatch = as.integer(2))
		 as(ans, 'Toolbars')
},
'Windows' = function() {
		 ans = .COM(x, 'Windows', .dispatch = as.integer(2))
		 as(ans, 'Windows')
},
'Workbooks' = function() {
		 ans = .COM(x, 'Workbooks', .dispatch = as.integer(2))
		 as(ans, 'Workbooks')
},
'WorksheetFunction' = function() {
		 ans = .COM(x, 'WorksheetFunction', .dispatch = as.integer(2))
		 as(ans, 'WorksheetFunction')
},
'Worksheets' = function() {
		 ans = .COM(x, 'Worksheets', .dispatch = as.integer(2))
		 as(ans, 'Sheets')
},
'Excel4IntlMacroSheets' = function() {
		 ans = .COM(x, 'Excel4IntlMacroSheets', .dispatch = as.integer(2))
		 as(ans, 'Sheets')
},
'Excel4MacroSheets' = function() {
		 ans = .COM(x, 'Excel4MacroSheets', .dispatch = as.integer(2))
		 as(ans, 'Sheets')
},
'AlertBeforeOverwriting' = function() {
		 ans = .COM(x, 'AlertBeforeOverwriting', .dispatch = as.integer(2))
	ans
},
'AltStartupPath' = function() {
		 ans = .COM(x, 'AltStartupPath', .dispatch = as.integer(2))
	ans
},
'AskToUpdateLinks' = function() {
		 ans = .COM(x, 'AskToUpdateLinks', .dispatch = as.integer(2))
	ans
},
'EnableAnimations' = function() {
		 ans = .COM(x, 'EnableAnimations', .dispatch = as.integer(2))
	ans
},
'AutoCorrect' = function() {
		 ans = .COM(x, 'AutoCorrect', .dispatch = as.integer(2))
		 as(ans, 'AutoCorrect')
},
'Build' = function() {
		 ans = .COM(x, 'Build', .dispatch = as.integer(2))
	ans
},
'CalculateBeforeSave' = function() {
		 ans = .COM(x, 'CalculateBeforeSave', .dispatch = as.integer(2))
	ans
},
'Calculation' = function() {
		 ans = .COM(x, 'Calculation', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'CanPlaySounds' = function() {
		 ans = .COM(x, 'CanPlaySounds', .dispatch = as.integer(2))
	ans
},
'CanRecordSounds' = function() {
		 ans = .COM(x, 'CanRecordSounds', .dispatch = as.integer(2))
	ans
},
'Caption' = function() {
		 ans = .COM(x, 'Caption', .dispatch = as.integer(2))
	ans
},
'CellDragAndDrop' = function() {
		 ans = .COM(x, 'CellDragAndDrop', .dispatch = as.integer(2))
	ans
},
'DisplayClipboardWindow' = function() {
		 ans = .COM(x, 'DisplayClipboardWindow', .dispatch = as.integer(2))
	ans
},
'ColorButtons' = function() {
		 ans = .COM(x, 'ColorButtons', .dispatch = as.integer(2))
	ans
},
'CommandUnderlines' = function() {
		 ans = .COM(x, 'CommandUnderlines', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'ConstrainNumeric' = function() {
		 ans = .COM(x, 'ConstrainNumeric', .dispatch = as.integer(2))
	ans
},
'CopyObjectsWithCells' = function() {
		 ans = .COM(x, 'CopyObjectsWithCells', .dispatch = as.integer(2))
	ans
},
'Cursor' = function() {
		 ans = .COM(x, 'Cursor', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'CustomListCount' = function() {
		 ans = .COM(x, 'CustomListCount', .dispatch = as.integer(2))
	ans
},
'CutCopyMode' = function() {
		 ans = .COM(x, 'CutCopyMode', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'DataEntryMode' = function() {
		 ans = .COM(x, 'DataEntryMode', .dispatch = as.integer(2))
	ans
},
'_Default' = function() {
		 ans = .COM(x, '_Default', .dispatch = as.integer(2))
	ans
},
'DefaultFilePath' = function() {
		 ans = .COM(x, 'DefaultFilePath', .dispatch = as.integer(2))
	ans
},
'Dialogs' = function() {
		 ans = .COM(x, 'Dialogs', .dispatch = as.integer(2))
		 as(ans, 'Dialogs')
},
'DisplayAlerts' = function() {
		 ans = .COM(x, 'DisplayAlerts', .dispatch = as.integer(2))
	ans
},
'DisplayFormulaBar' = function() {
		 ans = .COM(x, 'DisplayFormulaBar', .dispatch = as.integer(2))
	ans
},
'DisplayFullScreen' = function() {
		 ans = .COM(x, 'DisplayFullScreen', .dispatch = as.integer(2))
	ans
},
'DisplayNoteIndicator' = function() {
		 ans = .COM(x, 'DisplayNoteIndicator', .dispatch = as.integer(2))
	ans
},
'DisplayCommentIndicator' = function() {
		 ans = .COM(x, 'DisplayCommentIndicator', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'DisplayExcel4Menus' = function() {
		 ans = .COM(x, 'DisplayExcel4Menus', .dispatch = as.integer(2))
	ans
},
'DisplayRecentFiles' = function() {
		 ans = .COM(x, 'DisplayRecentFiles', .dispatch = as.integer(2))
	ans
},
'DisplayScrollBars' = function() {
		 ans = .COM(x, 'DisplayScrollBars', .dispatch = as.integer(2))
	ans
},
'DisplayStatusBar' = function() {
		 ans = .COM(x, 'DisplayStatusBar', .dispatch = as.integer(2))
	ans
},
'EditDirectlyInCell' = function() {
		 ans = .COM(x, 'EditDirectlyInCell', .dispatch = as.integer(2))
	ans
},
'EnableAutoComplete' = function() {
		 ans = .COM(x, 'EnableAutoComplete', .dispatch = as.integer(2))
	ans
},
'EnableCancelKey' = function() {
		 ans = .COM(x, 'EnableCancelKey', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'EnableSound' = function() {
		 ans = .COM(x, 'EnableSound', .dispatch = as.integer(2))
	ans
},
'EnableTipWizard' = function() {
		 ans = .COM(x, 'EnableTipWizard', .dispatch = as.integer(2))
	ans
},
'FileSearch' = function() {
		 ans = .COM(x, 'FileSearch', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'FileFind' = function() {
		 ans = .COM(x, 'FileFind', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'FixedDecimal' = function() {
		 ans = .COM(x, 'FixedDecimal', .dispatch = as.integer(2))
	ans
},
'FixedDecimalPlaces' = function() {
		 ans = .COM(x, 'FixedDecimalPlaces', .dispatch = as.integer(2))
	ans
},
'Height' = function() {
		 ans = .COM(x, 'Height', .dispatch = as.integer(2))
	ans
},
'IgnoreRemoteRequests' = function() {
		 ans = .COM(x, 'IgnoreRemoteRequests', .dispatch = as.integer(2))
	ans
},
'Interactive' = function() {
		 ans = .COM(x, 'Interactive', .dispatch = as.integer(2))
	ans
},
'Iteration' = function() {
		 ans = .COM(x, 'Iteration', .dispatch = as.integer(2))
	ans
},
'LargeButtons' = function() {
		 ans = .COM(x, 'LargeButtons', .dispatch = as.integer(2))
	ans
},
'Left' = function() {
		 ans = .COM(x, 'Left', .dispatch = as.integer(2))
	ans
},
'LibraryPath' = function() {
		 ans = .COM(x, 'LibraryPath', .dispatch = as.integer(2))
	ans
},
'MailSession' = function() {
		 ans = .COM(x, 'MailSession', .dispatch = as.integer(2))
		 ans
},
'MailSystem' = function() {
		 ans = .COM(x, 'MailSystem', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'MathCoprocessorAvailable' = function() {
		 ans = .COM(x, 'MathCoprocessorAvailable', .dispatch = as.integer(2))
	ans
},
'MaxChange' = function() {
		 ans = .COM(x, 'MaxChange', .dispatch = as.integer(2))
	ans
},
'MaxIterations' = function() {
		 ans = .COM(x, 'MaxIterations', .dispatch = as.integer(2))
	ans
},
'MemoryFree' = function() {
		 ans = .COM(x, 'MemoryFree', .dispatch = as.integer(2))
	ans
},
'MemoryTotal' = function() {
		 ans = .COM(x, 'MemoryTotal', .dispatch = as.integer(2))
	ans
},
'MemoryUsed' = function() {
		 ans = .COM(x, 'MemoryUsed', .dispatch = as.integer(2))
	ans
},
'MouseAvailable' = function() {
		 ans = .COM(x, 'MouseAvailable', .dispatch = as.integer(2))
	ans
},
'MoveAfterReturn' = function() {
		 ans = .COM(x, 'MoveAfterReturn', .dispatch = as.integer(2))
	ans
},
'MoveAfterReturnDirection' = function() {
		 ans = .COM(x, 'MoveAfterReturnDirection', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'RecentFiles' = function() {
		 ans = .COM(x, 'RecentFiles', .dispatch = as.integer(2))
		 as(ans, 'RecentFiles')
},
'Name' = function() {
		 ans = .COM(x, 'Name', .dispatch = as.integer(2))
	ans
},
'NetworkTemplatesPath' = function() {
		 ans = .COM(x, 'NetworkTemplatesPath', .dispatch = as.integer(2))
	ans
},
'ODBCErrors' = function() {
		 ans = .COM(x, 'ODBCErrors', .dispatch = as.integer(2))
		 as(ans, 'ODBCErrors')
},
'ODBCTimeout' = function() {
		 ans = .COM(x, 'ODBCTimeout', .dispatch = as.integer(2))
	ans
},
'OnCalculate' = function() {
		 ans = .COM(x, 'OnCalculate', .dispatch = as.integer(2))
	ans
},
'OnData' = function() {
		 ans = .COM(x, 'OnData', .dispatch = as.integer(2))
	ans
},
'OnDoubleClick' = function() {
		 ans = .COM(x, 'OnDoubleClick', .dispatch = as.integer(2))
	ans
},
'OnEntry' = function() {
		 ans = .COM(x, 'OnEntry', .dispatch = as.integer(2))
	ans
},
'OnSheetActivate' = function() {
		 ans = .COM(x, 'OnSheetActivate', .dispatch = as.integer(2))
	ans
},
'OnSheetDeactivate' = function() {
		 ans = .COM(x, 'OnSheetDeactivate', .dispatch = as.integer(2))
	ans
},
'OnWindow' = function() {
		 ans = .COM(x, 'OnWindow', .dispatch = as.integer(2))
	ans
},
'OperatingSystem' = function() {
		 ans = .COM(x, 'OperatingSystem', .dispatch = as.integer(2))
	ans
},
'OrganizationName' = function() {
		 ans = .COM(x, 'OrganizationName', .dispatch = as.integer(2))
	ans
},
'Path' = function() {
		 ans = .COM(x, 'Path', .dispatch = as.integer(2))
	ans
},
'PathSeparator' = function() {
		 ans = .COM(x, 'PathSeparator', .dispatch = as.integer(2))
	ans
},
'PivotTableSelection' = function() {
		 ans = .COM(x, 'PivotTableSelection', .dispatch = as.integer(2))
	ans
},
'PromptForSummaryInfo' = function() {
		 ans = .COM(x, 'PromptForSummaryInfo', .dispatch = as.integer(2))
	ans
},
'RecordRelative' = function() {
		 ans = .COM(x, 'RecordRelative', .dispatch = as.integer(2))
	ans
},
'ReferenceStyle' = function() {
		 ans = .COM(x, 'ReferenceStyle', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'RollZoom' = function() {
		 ans = .COM(x, 'RollZoom', .dispatch = as.integer(2))
	ans
},
'ScreenUpdating' = function() {
		 ans = .COM(x, 'ScreenUpdating', .dispatch = as.integer(2))
	ans
},
'SheetsInNewWorkbook' = function() {
		 ans = .COM(x, 'SheetsInNewWorkbook', .dispatch = as.integer(2))
	ans
},
'ShowChartTipNames' = function() {
		 ans = .COM(x, 'ShowChartTipNames', .dispatch = as.integer(2))
	ans
},
'ShowChartTipValues' = function() {
		 ans = .COM(x, 'ShowChartTipValues', .dispatch = as.integer(2))
	ans
},
'StandardFont' = function() {
		 ans = .COM(x, 'StandardFont', .dispatch = as.integer(2))
	ans
},
'StandardFontSize' = function() {
		 ans = .COM(x, 'StandardFontSize', .dispatch = as.integer(2))
	ans
},
'StartupPath' = function() {
		 ans = .COM(x, 'StartupPath', .dispatch = as.integer(2))
	ans
},
'StatusBar' = function() {
		 ans = .COM(x, 'StatusBar', .dispatch = as.integer(2))
		 ans
},
'TemplatesPath' = function() {
		 ans = .COM(x, 'TemplatesPath', .dispatch = as.integer(2))
	ans
},
'ShowToolTips' = function() {
		 ans = .COM(x, 'ShowToolTips', .dispatch = as.integer(2))
	ans
},
'Top' = function() {
		 ans = .COM(x, 'Top', .dispatch = as.integer(2))
	ans
},
'DefaultSaveFormat' = function() {
		 ans = .COM(x, 'DefaultSaveFormat', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'TransitionMenuKey' = function() {
		 ans = .COM(x, 'TransitionMenuKey', .dispatch = as.integer(2))
	ans
},
'TransitionMenuKeyAction' = function() {
		 ans = .COM(x, 'TransitionMenuKeyAction', .dispatch = as.integer(2))
	ans
},
'TransitionNavigKeys' = function() {
		 ans = .COM(x, 'TransitionNavigKeys', .dispatch = as.integer(2))
	ans
},
'UsableHeight' = function() {
		 ans = .COM(x, 'UsableHeight', .dispatch = as.integer(2))
	ans
},
'UsableWidth' = function() {
		 ans = .COM(x, 'UsableWidth', .dispatch = as.integer(2))
	ans
},
'UserControl' = function() {
		 ans = .COM(x, 'UserControl', .dispatch = as.integer(2))
	ans
},
'UserName' = function() {
		 ans = .COM(x, 'UserName', .dispatch = as.integer(2))
	ans
},
'Value' = function() {
		 ans = .COM(x, 'Value', .dispatch = as.integer(2))
	ans
},
'VBE' = function() {
		 ans = .COM(x, 'VBE', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'Version' = function() {
		 ans = .COM(x, 'Version', .dispatch = as.integer(2))
	ans
},
'Visible' = function() {
		 ans = .COM(x, 'Visible', .dispatch = as.integer(2))
	ans
},
'Width' = function() {
		 ans = .COM(x, 'Width', .dispatch = as.integer(2))
	ans
},
'WindowsForPens' = function() {
		 ans = .COM(x, 'WindowsForPens', .dispatch = as.integer(2))
	ans
},
'WindowState' = function() {
		 ans = .COM(x, 'WindowState', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'UILanguage' = function() {
		 ans = .COM(x, 'UILanguage', .dispatch = as.integer(2))
	ans
},
'DefaultSheetDirection' = function() {
		 ans = .COM(x, 'DefaultSheetDirection', .dispatch = as.integer(2))
	ans
},
'CursorMovement' = function() {
		 ans = .COM(x, 'CursorMovement', .dispatch = as.integer(2))
	ans
},
'ControlCharacters' = function() {
		 ans = .COM(x, 'ControlCharacters', .dispatch = as.integer(2))
	ans
},
'EnableEvents' = function() {
		 ans = .COM(x, 'EnableEvents', .dispatch = as.integer(2))
	ans
},
'DisplayInfoWindow' = function() {
		 ans = .COM(x, 'DisplayInfoWindow', .dispatch = as.integer(2))
	ans
},
'ExtendList' = function() {
		 ans = .COM(x, 'ExtendList', .dispatch = as.integer(2))
	ans
},
'OLEDBErrors' = function() {
		 ans = .COM(x, 'OLEDBErrors', .dispatch = as.integer(2))
		 as(ans, 'OLEDBErrors')
},
'COMAddIns' = function() {
		 ans = .COM(x, 'COMAddIns', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'DefaultWebOptions' = function() {
		 ans = .COM(x, 'DefaultWebOptions', .dispatch = as.integer(2))
		 as(ans, 'DefaultWebOptions')
},
'ProductCode' = function() {
		 ans = .COM(x, 'ProductCode', .dispatch = as.integer(2))
	ans
},
'UserLibraryPath' = function() {
		 ans = .COM(x, 'UserLibraryPath', .dispatch = as.integer(2))
	ans
},
'AutoPercentEntry' = function() {
		 ans = .COM(x, 'AutoPercentEntry', .dispatch = as.integer(2))
	ans
},
'LanguageSettings' = function() {
		 ans = .COM(x, 'LanguageSettings', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'Dummy101' = function() {
		 ans = .COM(x, 'Dummy101', .dispatch = as.integer(2))
		 ans
},
'AnswerWizard' = function() {
		 ans = .COM(x, 'AnswerWizard', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'CalculationVersion' = function() {
		 ans = .COM(x, 'CalculationVersion', .dispatch = as.integer(2))
	ans
},
'ShowWindowsInTaskbar' = function() {
		 ans = .COM(x, 'ShowWindowsInTaskbar', .dispatch = as.integer(2))
	ans
},
'FeatureInstall' = function() {
		 ans = .COM(x, 'FeatureInstall', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'Ready' = function() {
		 ans = .COM(x, 'Ready', .dispatch = as.integer(2))
	ans
},
'FindFormat' = function() {
		 ans = .COM(x, 'FindFormat', .dispatch = as.integer(2))
		 as(ans, 'CellFormat')
},
'ReplaceFormat' = function() {
		 ans = .COM(x, 'ReplaceFormat', .dispatch = as.integer(2))
		 as(ans, 'CellFormat')
},
'UsedObjects' = function() {
		 ans = .COM(x, 'UsedObjects', .dispatch = as.integer(2))
		 as(ans, 'UsedObjects')
},
'CalculationState' = function() {
		 ans = .COM(x, 'CalculationState', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'CalculationInterruptKey' = function() {
		 ans = .COM(x, 'CalculationInterruptKey', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'Watches' = function() {
		 ans = .COM(x, 'Watches', .dispatch = as.integer(2))
		 as(ans, 'Watches')
},
'DisplayFunctionToolTips' = function() {
		 ans = .COM(x, 'DisplayFunctionToolTips', .dispatch = as.integer(2))
	ans
},
'AutomationSecurity' = function() {
		 ans = .COM(x, 'AutomationSecurity', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'DisplayPasteOptions' = function() {
		 ans = .COM(x, 'DisplayPasteOptions', .dispatch = as.integer(2))
	ans
},
'DisplayInsertOptions' = function() {
		 ans = .COM(x, 'DisplayInsertOptions', .dispatch = as.integer(2))
	ans
},
'GenerateGetPivotData' = function() {
		 ans = .COM(x, 'GenerateGetPivotData', .dispatch = as.integer(2))
	ans
},
'AutoRecover' = function() {
		 ans = .COM(x, 'AutoRecover', .dispatch = as.integer(2))
		 as(ans, 'AutoRecover')
},
'Hwnd' = function() {
		 ans = .COM(x, 'Hwnd', .dispatch = as.integer(2))
	ans
},
'Hinstance' = function() {
		 ans = .COM(x, 'Hinstance', .dispatch = as.integer(2))
	ans
},
'ErrorCheckingOptions' = function() {
		 ans = .COM(x, 'ErrorCheckingOptions', .dispatch = as.integer(2))
		 as(ans, 'ErrorCheckingOptions')
},
'AutoFormatAsYouTypeReplaceHyperlinks' = function() {
		 ans = .COM(x, 'AutoFormatAsYouTypeReplaceHyperlinks', .dispatch = as.integer(2))
	ans
},
'SmartTagRecognizers' = function() {
		 ans = .COM(x, 'SmartTagRecognizers', .dispatch = as.integer(2))
		 as(ans, 'SmartTagRecognizers')
},
'NewWorkbook' = function() {
		 ans = .COM(x, 'NewWorkbook', .dispatch = as.integer(2))
		 as(ans, 'NA')
},
'SpellingOptions' = function() {
		 ans = .COM(x, 'SpellingOptions', .dispatch = as.integer(2))
		 as(ans, 'SpellingOptions')
},
'Speech' = function() {
		 ans = .COM(x, 'Speech', .dispatch = as.integer(2))
		 as(ans, 'Speech')
},
'MapPaperSize' = function() {
		 ans = .COM(x, 'MapPaperSize', .dispatch = as.integer(2))
	ans
},
'ShowStartupDialog' = function() {
		 ans = .COM(x, 'ShowStartupDialog', .dispatch = as.integer(2))
	ans
},
'DecimalSeparator' = function() {
		 ans = .COM(x, 'DecimalSeparator', .dispatch = as.integer(2))
	ans
},
'ThousandsSeparator' = function() {
		 ans = .COM(x, 'ThousandsSeparator', .dispatch = as.integer(2))
	ans
},
'UseSystemSeparators' = function() {
		 ans = .COM(x, 'UseSystemSeparators', .dispatch = as.integer(2))
	ans
},
'ThisCell' = function() {
		 ans = .COM(x, 'ThisCell', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'RTD' = function() {
		 ans = .COM(x, 'RTD', .dispatch = as.integer(2))
		 as(ans, 'RTD')
},
'DisplayDocumentActionTaskPane' = function() {
		 ans = .COM(x, 'DisplayDocumentActionTaskPane', .dispatch = as.integer(2))
	ans
},
'ArbitraryXMLSupportAvailable' = function() {
		 ans = .COM(x, 'ArbitraryXMLSupportAvailable', .dispatch = as.integer(2))
	ans
} )
'COM._Application.SetProperty'  = list('ActivePrinter' = function(value) {
	value = as(value, 'character')
	.COM(x, 'ActivePrinter', value, .dispatch = as.integer(4))
},
'AlertBeforeOverwriting' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'AlertBeforeOverwriting', value, .dispatch = as.integer(4))
},
'AltStartupPath' = function(value) {
	value = as(value, 'character')
	.COM(x, 'AltStartupPath', value, .dispatch = as.integer(4))
},
'AskToUpdateLinks' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'AskToUpdateLinks', value, .dispatch = as.integer(4))
},
'EnableAnimations' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'EnableAnimations', value, .dispatch = as.integer(4))
},
'CalculateBeforeSave' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'CalculateBeforeSave', value, .dispatch = as.integer(4))
},
'Calculation' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'Calculation', value, .dispatch = as.integer(4))
},
'Caption' = function(value) {
	value = as(value, 'character')
	.COM(x, 'Caption', value, .dispatch = as.integer(4))
},
'CellDragAndDrop' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'CellDragAndDrop', value, .dispatch = as.integer(4))
},
'DisplayClipboardWindow' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayClipboardWindow', value, .dispatch = as.integer(4))
},
'ColorButtons' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ColorButtons', value, .dispatch = as.integer(4))
},
'CommandUnderlines' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'CommandUnderlines', value, .dispatch = as.integer(4))
},
'ConstrainNumeric' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ConstrainNumeric', value, .dispatch = as.integer(4))
},
'CopyObjectsWithCells' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'CopyObjectsWithCells', value, .dispatch = as.integer(4))
},
'Cursor' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'Cursor', value, .dispatch = as.integer(4))
},
'CutCopyMode' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'CutCopyMode', value, .dispatch = as.integer(4))
},
'DataEntryMode' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'DataEntryMode', value, .dispatch = as.integer(4))
},
'DefaultFilePath' = function(value) {
	value = as(value, 'character')
	.COM(x, 'DefaultFilePath', value, .dispatch = as.integer(4))
},
'DisplayAlerts' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayAlerts', value, .dispatch = as.integer(4))
},
'DisplayFormulaBar' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayFormulaBar', value, .dispatch = as.integer(4))
},
'DisplayFullScreen' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayFullScreen', value, .dispatch = as.integer(4))
},
'DisplayNoteIndicator' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayNoteIndicator', value, .dispatch = as.integer(4))
},
'DisplayCommentIndicator' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'DisplayCommentIndicator', value, .dispatch = as.integer(4))
},
'DisplayExcel4Menus' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayExcel4Menus', value, .dispatch = as.integer(4))
},
'DisplayRecentFiles' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayRecentFiles', value, .dispatch = as.integer(4))
},
'DisplayScrollBars' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayScrollBars', value, .dispatch = as.integer(4))
},
'DisplayStatusBar' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayStatusBar', value, .dispatch = as.integer(4))
},
'EditDirectlyInCell' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'EditDirectlyInCell', value, .dispatch = as.integer(4))
},
'EnableAutoComplete' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'EnableAutoComplete', value, .dispatch = as.integer(4))
},
'EnableCancelKey' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'EnableCancelKey', value, .dispatch = as.integer(4))
},
'EnableSound' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'EnableSound', value, .dispatch = as.integer(4))
},
'EnableTipWizard' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'EnableTipWizard', value, .dispatch = as.integer(4))
},
'FixedDecimal' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'FixedDecimal', value, .dispatch = as.integer(4))
},
'FixedDecimalPlaces' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'FixedDecimalPlaces', value, .dispatch = as.integer(4))
},
'Height' = function(value) {
	value = as(value, 'numeric')
	.COM(x, 'Height', value, .dispatch = as.integer(4))
},
'IgnoreRemoteRequests' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'IgnoreRemoteRequests', value, .dispatch = as.integer(4))
},
'Interactive' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'Interactive', value, .dispatch = as.integer(4))
},
'Iteration' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'Iteration', value, .dispatch = as.integer(4))
},
'LargeButtons' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'LargeButtons', value, .dispatch = as.integer(4))
},
'Left' = function(value) {
	value = as(value, 'numeric')
	.COM(x, 'Left', value, .dispatch = as.integer(4))
},
'MaxChange' = function(value) {
	value = as(value, 'numeric')
	.COM(x, 'MaxChange', value, .dispatch = as.integer(4))
},
'MaxIterations' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'MaxIterations', value, .dispatch = as.integer(4))
},
'MoveAfterReturn' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'MoveAfterReturn', value, .dispatch = as.integer(4))
},
'MoveAfterReturnDirection' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'MoveAfterReturnDirection', value, .dispatch = as.integer(4))
},
'ODBCTimeout' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'ODBCTimeout', value, .dispatch = as.integer(4))
},
'OnCalculate' = function(value) {
	value = as(value, 'character')
	.COM(x, 'OnCalculate', value, .dispatch = as.integer(4))
},
'OnData' = function(value) {
	value = as(value, 'character')
	.COM(x, 'OnData', value, .dispatch = as.integer(4))
},
'OnDoubleClick' = function(value) {
	value = as(value, 'character')
	.COM(x, 'OnDoubleClick', value, .dispatch = as.integer(4))
},
'OnEntry' = function(value) {
	value = as(value, 'character')
	.COM(x, 'OnEntry', value, .dispatch = as.integer(4))
},
'OnSheetActivate' = function(value) {
	value = as(value, 'character')
	.COM(x, 'OnSheetActivate', value, .dispatch = as.integer(4))
},
'OnSheetDeactivate' = function(value) {
	value = as(value, 'character')
	.COM(x, 'OnSheetDeactivate', value, .dispatch = as.integer(4))
},
'OnWindow' = function(value) {
	value = as(value, 'character')
	.COM(x, 'OnWindow', value, .dispatch = as.integer(4))
},
'PivotTableSelection' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'PivotTableSelection', value, .dispatch = as.integer(4))
},
'PromptForSummaryInfo' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'PromptForSummaryInfo', value, .dispatch = as.integer(4))
},
'ReferenceStyle' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'ReferenceStyle', value, .dispatch = as.integer(4))
},
'RollZoom' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'RollZoom', value, .dispatch = as.integer(4))
},
'ScreenUpdating' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ScreenUpdating', value, .dispatch = as.integer(4))
},
'SheetsInNewWorkbook' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'SheetsInNewWorkbook', value, .dispatch = as.integer(4))
},
'ShowChartTipNames' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ShowChartTipNames', value, .dispatch = as.integer(4))
},
'ShowChartTipValues' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ShowChartTipValues', value, .dispatch = as.integer(4))
},
'StandardFont' = function(value) {
	value = as(value, 'character')
	.COM(x, 'StandardFont', value, .dispatch = as.integer(4))
},
'StandardFontSize' = function(value) {
	value = as(value, 'numeric')
	.COM(x, 'StandardFontSize', value, .dispatch = as.integer(4))
},
'StatusBar' = function(value) {
	
	.COM(x, 'StatusBar', value, .dispatch = as.integer(4))
},
'ShowToolTips' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ShowToolTips', value, .dispatch = as.integer(4))
},
'Top' = function(value) {
	value = as(value, 'numeric')
	.COM(x, 'Top', value, .dispatch = as.integer(4))
},
'DefaultSaveFormat' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'DefaultSaveFormat', value, .dispatch = as.integer(4))
},
'TransitionMenuKey' = function(value) {
	value = as(value, 'character')
	.COM(x, 'TransitionMenuKey', value, .dispatch = as.integer(4))
},
'TransitionMenuKeyAction' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'TransitionMenuKeyAction', value, .dispatch = as.integer(4))
},
'TransitionNavigKeys' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'TransitionNavigKeys', value, .dispatch = as.integer(4))
},
'UserControl' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'UserControl', value, .dispatch = as.integer(4))
},
'UserName' = function(value) {
	value = as(value, 'character')
	.COM(x, 'UserName', value, .dispatch = as.integer(4))
},
'Visible' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'Visible', value, .dispatch = as.integer(4))
},
'Width' = function(value) {
	value = as(value, 'numeric')
	.COM(x, 'Width', value, .dispatch = as.integer(4))
},
'WindowState' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'WindowState', value, .dispatch = as.integer(4))
},
'UILanguage' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'UILanguage', value, .dispatch = as.integer(4))
},
'DefaultSheetDirection' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'DefaultSheetDirection', value, .dispatch = as.integer(4))
},
'CursorMovement' = function(value) {
	value = as(value, 'integer')
	.COM(x, 'CursorMovement', value, .dispatch = as.integer(4))
},
'ControlCharacters' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ControlCharacters', value, .dispatch = as.integer(4))
},
'EnableEvents' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'EnableEvents', value, .dispatch = as.integer(4))
},
'DisplayInfoWindow' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayInfoWindow', value, .dispatch = as.integer(4))
},
'ExtendList' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ExtendList', value, .dispatch = as.integer(4))
},
'AutoPercentEntry' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'AutoPercentEntry', value, .dispatch = as.integer(4))
},
'ShowWindowsInTaskbar' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ShowWindowsInTaskbar', value, .dispatch = as.integer(4))
},
'FeatureInstall' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'FeatureInstall', value, .dispatch = as.integer(4))
},
'FindFormat' = function(value) {
	value = as(value, 'CellFormat')
	.COM(x, 'FindFormat', value, .dispatch = as.integer(4))
},
'ReplaceFormat' = function(value) {
	value = as(value, 'CellFormat')
	.COM(x, 'ReplaceFormat', value, .dispatch = as.integer(4))
},
'CalculationInterruptKey' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'CalculationInterruptKey', value, .dispatch = as.integer(4))
},
'DisplayFunctionToolTips' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayFunctionToolTips', value, .dispatch = as.integer(4))
},
'AutomationSecurity' = function(value) {
	value = as(value, 'Constants')
	.COM(x, 'AutomationSecurity', value, .dispatch = as.integer(4))
},
'DisplayPasteOptions' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayPasteOptions', value, .dispatch = as.integer(4))
},
'DisplayInsertOptions' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayInsertOptions', value, .dispatch = as.integer(4))
},
'GenerateGetPivotData' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'GenerateGetPivotData', value, .dispatch = as.integer(4))
},
'AutoFormatAsYouTypeReplaceHyperlinks' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'AutoFormatAsYouTypeReplaceHyperlinks', value, .dispatch = as.integer(4))
},
'MapPaperSize' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'MapPaperSize', value, .dispatch = as.integer(4))
},
'ShowStartupDialog' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'ShowStartupDialog', value, .dispatch = as.integer(4))
},
'DecimalSeparator' = function(value) {
	value = as(value, 'character')
	.COM(x, 'DecimalSeparator', value, .dispatch = as.integer(4))
},
'ThousandsSeparator' = function(value) {
	value = as(value, 'character')
	.COM(x, 'ThousandsSeparator', value, .dispatch = as.integer(4))
},
'UseSystemSeparators' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'UseSystemSeparators', value, .dispatch = as.integer(4))
},
'DisplayDocumentActionTaskPane' = function(value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayDocumentActionTaskPane', value, .dispatch = as.integer(4))
} )
'COM._Application.Methods'  = list('Calculate' = function(  ){
	ans = .COM(x, 'Calculate', .dispatch = as.integer(1), .ids =279)
		 ans
},
'DDEExecute' = function( Channel, String ){
	if( missing( Channel )||missing( String ) ) {
	stop('You must specify a value for the argument(s)  Channel, String ')
	}
	Channel = as(Channel, 'integer')
	String = as(String, 'character')
	ans = .COM(x, 'DDEExecute', Channel, String, .dispatch = as.integer(1), .ids =333)
		 ans
},
'DDEInitiate' = function( App, Topic ){
	if( missing( App )||missing( Topic ) ) {
	stop('You must specify a value for the argument(s)  App, Topic ')
	}
	App = as(App, 'character')
	Topic = as(Topic, 'character')
	ans = .COM(x, 'DDEInitiate', App, Topic, .dispatch = as.integer(1), .ids =334)
	ans
},
'DDEPoke' = function( Channel, Item, Data ){
	if( missing( Channel )||missing( Item )||missing( Data ) ) {
	stop('You must specify a value for the argument(s)  Channel, Item, Data ')
	}
	Channel = as(Channel, 'integer')
	
	
	ans = .COM(x, 'DDEPoke', Channel, Item, Data, .dispatch = as.integer(1), .ids =335)
		 ans
},
'DDERequest' = function( Channel, Item ){
	if( missing( Channel )||missing( Item ) ) {
	stop('You must specify a value for the argument(s)  Channel, Item ')
	}
	Channel = as(Channel, 'integer')
	Item = as(Item, 'character')
	ans = .COM(x, 'DDERequest', Channel, Item, .dispatch = as.integer(1), .ids =336)
		 ans
},
'DDETerminate' = function( Channel ){
	if( missing( Channel ) ) {
	stop('You must specify a value for the argument(s)  Channel ')
	}
	Channel = as(Channel, 'integer')
	ans = .COM(x, 'DDETerminate', Channel, .dispatch = as.integer(1), .ids =337)
		 ans
},
'Evaluate' = function( Name ){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	
	ans = .COM(x, 'Evaluate', Name, .dispatch = as.integer(1), .ids =1)
		 ans
},
'_Evaluate' = function( Name ){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	
	ans = .COM(x, '_Evaluate', Name, .dispatch = as.integer(1), .ids =-5)
		 ans
},
'ExecuteExcel4Macro' = function( String ){
	if( missing( String ) ) {
	stop('You must specify a value for the argument(s)  String ')
	}
	String = as(String, 'character')
	ans = .COM(x, 'ExecuteExcel4Macro', String, .dispatch = as.integer(1), .ids =350)
		 ans
},
'Intersect' = function( Arg1, Arg2, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	if( missing( Arg1 )||missing( Arg2 ) ) {
	stop('You must specify a value for the argument(s)  Arg1, Arg2 ')
	}
	Arg1 = as(Arg1, 'Range')
	Arg2 = as(Arg2, 'Range')
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Intersect', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =766)
		 as(ans, 'Range')
},
'Run' = function( Macro = NA, Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Run', Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =259)
		 ans
},
'_Run2' = function( Macro = NA, Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, '_Run2', Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =806)
		 ans
},
'SendKeys' = function( Keys, Wait = NA ){
	if( missing( Keys ) ) {
	stop('You must specify a value for the argument(s)  Keys ')
	}
	
	
	ans = .COM(x, 'SendKeys', Keys, Wait, .dispatch = as.integer(1), .ids =383)
		 ans
},
'Union' = function( Arg1, Arg2, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	if( missing( Arg1 )||missing( Arg2 ) ) {
	stop('You must specify a value for the argument(s)  Arg1, Arg2 ')
	}
	Arg1 = as(Arg1, 'Range')
	Arg2 = as(Arg2, 'Range')
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Union', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =779)
		 as(ans, 'Range')
},
'ActivateMicrosoftApp' = function( Index ){
	if( missing( Index ) ) {
	stop('You must specify a value for the argument(s)  Index ')
	}
	Index = as(Index, 'Constants')
	ans = .COM(x, 'ActivateMicrosoftApp', Index, .dispatch = as.integer(1), .ids =1095)
		 ans
},
'AddChartAutoFormat' = function( Chart, Name, Description = NA ){
	if( missing( Chart )||missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Chart, Name ')
	}
	
	Name = as(Name, 'character')
	
	ans = .COM(x, 'AddChartAutoFormat', Chart, Name, Description, .dispatch = as.integer(1), .ids =216)
		 ans
},
'AddCustomList' = function( ListArray, ByRow = NA ){
	if( missing( ListArray ) ) {
	stop('You must specify a value for the argument(s)  ListArray ')
	}
	
	
	ans = .COM(x, 'AddCustomList', ListArray, ByRow, .dispatch = as.integer(1), .ids =780)
		 ans
},
'CentimetersToPoints' = function( Centimeters ){
	if( missing( Centimeters ) ) {
	stop('You must specify a value for the argument(s)  Centimeters ')
	}
	Centimeters = as(Centimeters, 'numeric')
	ans = .COM(x, 'CentimetersToPoints', Centimeters, .dispatch = as.integer(1), .ids =1086)
	ans
},
'CheckSpelling' = function( Word, CustomDictionary = NA, IgnoreUppercase = NA ){
	if( missing( Word ) ) {
	stop('You must specify a value for the argument(s)  Word ')
	}
	Word = as(Word, 'character')
	
	
	ans = .COM(x, 'CheckSpelling', Word, CustomDictionary, IgnoreUppercase, .dispatch = as.integer(1), .ids =505)
	ans
},
'ConvertFormula' = function( Formula, FromReferenceStyle, ToReferenceStyle = NA, ToAbsolute = NA, RelativeTo = NA ){
	if( missing( Formula )||missing( FromReferenceStyle ) ) {
	stop('You must specify a value for the argument(s)  Formula, FromReferenceStyle ')
	}
	
	FromReferenceStyle = as(FromReferenceStyle, 'Constants')
	
	
	
	ans = .COM(x, 'ConvertFormula', Formula, FromReferenceStyle, ToReferenceStyle, ToAbsolute, RelativeTo, .dispatch = as.integer(1), .ids =325)
		 ans
},
'Dummy1' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA ){
	
	
	
	
	ans = .COM(x, 'Dummy1', Arg1, Arg2, Arg3, Arg4, .dispatch = as.integer(1), .ids =1782)
		 ans
},
'Dummy2' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA ){
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Dummy2', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, .dispatch = as.integer(1), .ids =1783)
		 ans
},
'Dummy3' = function(  ){
	ans = .COM(x, 'Dummy3', .dispatch = as.integer(1), .ids =1784)
		 ans
},
'Dummy4' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Dummy4', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, .dispatch = as.integer(1), .ids =1785)
		 ans
},
'Dummy5' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Dummy5', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, .dispatch = as.integer(1), .ids =1786)
		 ans
},
'Dummy6' = function(  ){
	ans = .COM(x, 'Dummy6', .dispatch = as.integer(1), .ids =1787)
		 ans
},
'Dummy7' = function(  ){
	ans = .COM(x, 'Dummy7', .dispatch = as.integer(1), .ids =1788)
		 ans
},
'Dummy8' = function( Arg1 = NA ){
	
	ans = .COM(x, 'Dummy8', Arg1, .dispatch = as.integer(1), .ids =1789)
		 ans
},
'Dummy9' = function(  ){
	ans = .COM(x, 'Dummy9', .dispatch = as.integer(1), .ids =1790)
		 ans
},
'Dummy10' = function( arg = NA ){
	
	ans = .COM(x, 'Dummy10', arg, .dispatch = as.integer(1), .ids =1791)
	ans
},
'Dummy11' = function(  ){
	ans = .COM(x, 'Dummy11', .dispatch = as.integer(1), .ids =1792)
		 ans
},
'DeleteChartAutoFormat' = function( Name ){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	Name = as(Name, 'character')
	ans = .COM(x, 'DeleteChartAutoFormat', Name, .dispatch = as.integer(1), .ids =217)
		 ans
},
'DeleteCustomList' = function( ListNum ){
	if( missing( ListNum ) ) {
	stop('You must specify a value for the argument(s)  ListNum ')
	}
	ListNum = as(ListNum, 'integer')
	ans = .COM(x, 'DeleteCustomList', ListNum, .dispatch = as.integer(1), .ids =783)
		 ans
},
'DoubleClick' = function(  ){
	ans = .COM(x, 'DoubleClick', .dispatch = as.integer(1), .ids =349)
		 ans
},
'_FindFile' = function(  ){
	ans = .COM(x, '_FindFile', .dispatch = as.integer(1), .ids =1068)
		 ans
},
'GetCustomListContents' = function( ListNum ){
	if( missing( ListNum ) ) {
	stop('You must specify a value for the argument(s)  ListNum ')
	}
	ListNum = as(ListNum, 'integer')
	ans = .COM(x, 'GetCustomListContents', ListNum, .dispatch = as.integer(1), .ids =786)
		 ans
},
'GetCustomListNum' = function( ListArray ){
	if( missing( ListArray ) ) {
	stop('You must specify a value for the argument(s)  ListArray ')
	}
	
	ans = .COM(x, 'GetCustomListNum', ListArray, .dispatch = as.integer(1), .ids =785)
	ans
},
'GetOpenFilename' = function( FileFilter = NA, FilterIndex = NA, Title = NA, ButtonText = NA, MultiSelect = NA ){
	
	
	
	
	
	ans = .COM(x, 'GetOpenFilename', FileFilter, FilterIndex, Title, ButtonText, MultiSelect, .dispatch = as.integer(1), .ids =1075)
		 ans
},
'GetSaveAsFilename' = function( InitialFilename = NA, FileFilter = NA, FilterIndex = NA, Title = NA, ButtonText = NA ){
	
	
	
	
	
	ans = .COM(x, 'GetSaveAsFilename', InitialFilename, FileFilter, FilterIndex, Title, ButtonText, .dispatch = as.integer(1), .ids =1076)
		 ans
},
'Goto' = function( Reference = NA, Scroll = NA ){
	
	
	ans = .COM(x, 'Goto', Reference, Scroll, .dispatch = as.integer(1), .ids =475)
		 ans
},
'Help' = function( HelpFile = NA, HelpContextID = NA ){
	
	
	ans = .COM(x, 'Help', HelpFile, HelpContextID, .dispatch = as.integer(1), .ids =354)
		 ans
},
'InchesToPoints' = function( Inches ){
	if( missing( Inches ) ) {
	stop('You must specify a value for the argument(s)  Inches ')
	}
	Inches = as(Inches, 'numeric')
	ans = .COM(x, 'InchesToPoints', Inches, .dispatch = as.integer(1), .ids =1087)
	ans
},
'InputBox' = function( Prompt, Title = NA, Default = NA, Left = NA, Top = NA, HelpFile = NA, HelpContextID = NA, Type = NA ){
	if( missing( Prompt ) ) {
	stop('You must specify a value for the argument(s)  Prompt ')
	}
	Prompt = as(Prompt, 'character')
	
	
	
	
	
	
	
	ans = .COM(x, 'InputBox', Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type, .dispatch = as.integer(1), .ids =357)
		 ans
},
'MacroOptions' = function( Macro = NA, Description = NA, HasMenu = NA, MenuText = NA, HasShortcutKey = NA, ShortcutKey = NA, Category = NA, StatusBar = NA, HelpContextID = NA, HelpFile = NA ){
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'MacroOptions', Macro, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey, Category, StatusBar, HelpContextID, HelpFile, .dispatch = as.integer(1), .ids =1135)
		 ans
},
'MailLogoff' = function(  ){
	ans = .COM(x, 'MailLogoff', .dispatch = as.integer(1), .ids =945)
		 ans
},
'MailLogon' = function( Name = NA, Password = NA, DownloadNewMail = NA ){
	
	
	
	ans = .COM(x, 'MailLogon', Name, Password, DownloadNewMail, .dispatch = as.integer(1), .ids =943)
		 ans
},
'NextLetter' = function(  ){
	ans = .COM(x, 'NextLetter', .dispatch = as.integer(1), .ids =972)
		 as(ans, 'Workbook')
},
'OnKey' = function( Key, Procedure = NA ){
	if( missing( Key ) ) {
	stop('You must specify a value for the argument(s)  Key ')
	}
	Key = as(Key, 'character')
	
	ans = .COM(x, 'OnKey', Key, Procedure, .dispatch = as.integer(1), .ids =626)
		 ans
},
'OnRepeat' = function( Text, Procedure ){
	if( missing( Text )||missing( Procedure ) ) {
	stop('You must specify a value for the argument(s)  Text, Procedure ')
	}
	Text = as(Text, 'character')
	Procedure = as(Procedure, 'character')
	ans = .COM(x, 'OnRepeat', Text, Procedure, .dispatch = as.integer(1), .ids =769)
		 ans
},
'OnTime' = function( EarliestTime, Procedure, LatestTime = NA, Schedule = NA ){
	if( missing( EarliestTime )||missing( Procedure ) ) {
	stop('You must specify a value for the argument(s)  EarliestTime, Procedure ')
	}
	
	Procedure = as(Procedure, 'character')
	
	
	ans = .COM(x, 'OnTime', EarliestTime, Procedure, LatestTime, Schedule, .dispatch = as.integer(1), .ids =624)
		 ans
},
'OnUndo' = function( Text, Procedure ){
	if( missing( Text )||missing( Procedure ) ) {
	stop('You must specify a value for the argument(s)  Text, Procedure ')
	}
	Text = as(Text, 'character')
	Procedure = as(Procedure, 'character')
	ans = .COM(x, 'OnUndo', Text, Procedure, .dispatch = as.integer(1), .ids =770)
		 ans
},
'Quit' = function(  ){
	ans = .COM(x, 'Quit', .dispatch = as.integer(1), .ids =302)
		 ans
},
'RecordMacro' = function( BasicCode = NA, XlmCode = NA ){
	
	
	ans = .COM(x, 'RecordMacro', BasicCode, XlmCode, .dispatch = as.integer(1), .ids =773)
		 ans
},
'RegisterXLL' = function( Filename ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	ans = .COM(x, 'RegisterXLL', Filename, .dispatch = as.integer(1), .ids =30)
	ans
},
'Repeat' = function(  ){
	ans = .COM(x, 'Repeat', .dispatch = as.integer(1), .ids =301)
		 ans
},
'ResetTipWizard' = function(  ){
	ans = .COM(x, 'ResetTipWizard', .dispatch = as.integer(1), .ids =928)
		 ans
},
'Save' = function( Filename = NA ){
	
	ans = .COM(x, 'Save', Filename, .dispatch = as.integer(1), .ids =283)
		 ans
},
'SaveWorkspace' = function( Filename = NA ){
	
	ans = .COM(x, 'SaveWorkspace', Filename, .dispatch = as.integer(1), .ids =212)
		 ans
},
'SetDefaultChart' = function( FormatName = NA, Gallery = NA ){
	
	
	ans = .COM(x, 'SetDefaultChart', FormatName, Gallery, .dispatch = as.integer(1), .ids =219)
		 ans
},
'Undo' = function(  ){
	ans = .COM(x, 'Undo', .dispatch = as.integer(1), .ids =303)
		 ans
},
'Volatile' = function( Volatile = NA ){
	
	ans = .COM(x, 'Volatile', Volatile, .dispatch = as.integer(1), .ids =788)
		 ans
},
'_Wait' = function( Time ){
	if( missing( Time ) ) {
	stop('You must specify a value for the argument(s)  Time ')
	}
	
	ans = .COM(x, '_Wait', Time, .dispatch = as.integer(1), .ids =393)
		 ans
},
'_WSFunction' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, '_WSFunction', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =169)
		 ans
},
'Wait' = function( Time ){
	if( missing( Time ) ) {
	stop('You must specify a value for the argument(s)  Time ')
	}
	
	ans = .COM(x, 'Wait', Time, .dispatch = as.integer(1), .ids =1770)
	ans
},
'GetPhonetic' = function( Text = NA ){
	
	ans = .COM(x, 'GetPhonetic', Text, .dispatch = as.integer(1), .ids =1795)
	ans
},
'Dummy12' = function( p1, p2 ){
	if( missing( p1 )||missing( p2 ) ) {
	stop('You must specify a value for the argument(s)  p1, p2 ')
	}
	p1 = as(p1, 'PivotTable')
	p2 = as(p2, 'PivotTable')
	ans = .COM(x, 'Dummy12', p1, p2, .dispatch = as.integer(1), .ids =1803)
		 ans
},
'CalculateFull' = function(  ){
	ans = .COM(x, 'CalculateFull', .dispatch = as.integer(1), .ids =1805)
		 ans
},
'FindFile' = function(  ){
	ans = .COM(x, 'FindFile', .dispatch = as.integer(1), .ids =1771)
	ans
},
'Dummy13' = function( Arg1, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	if( missing( Arg1 ) ) {
	stop('You must specify a value for the argument(s)  Arg1 ')
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Dummy13', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =1933)
		 ans
},
'Dummy14' = function(  ){
	ans = .COM(x, 'Dummy14', .dispatch = as.integer(1), .ids =1944)
		 ans
},
'CalculateFullRebuild' = function(  ){
	ans = .COM(x, 'CalculateFullRebuild', .dispatch = as.integer(1), .ids =1945)
		 ans
},
'CheckAbort' = function( KeepAbort = NA ){
	
	ans = .COM(x, 'CheckAbort', KeepAbort, .dispatch = as.integer(1), .ids =1952)
		 ans
},
'DisplayXMLSourcePane' = function( XmlMap = NA ){
	
	ans = .COM(x, 'DisplayXMLSourcePane', XmlMap, .dispatch = as.integer(1), .ids =2252)
		 ans
},
'Support' = function( Object, ID, arg = NA ){
	if( missing( Object )||missing( ID ) ) {
	stop('You must specify a value for the argument(s)  Object, ID ')
	}
	
	ID = as(ID, 'integer')
	
	ans = .COM(x, 'Support', Object, ID, arg, .dispatch = as.integer(1), .ids =2255)
		 ans
} )
'COM.Workbooks.GetProperty'  = list('Application' = function() {
		 ans = .COM(x, 'Application', .dispatch = as.integer(2))
		 as(ans, 'Application')
},
'Creator' = function() {
		 ans = .COM(x, 'Creator', .dispatch = as.integer(2))
		 as(ans, 'Constants')
},
'Parent' = function() {
		 ans = .COM(x, 'Parent', .dispatch = as.integer(2))
		 ans
},
'Count' = function() {
		 ans = .COM(x, 'Count', .dispatch = as.integer(2))
	ans
},
'_NewEnum' = function() {
		 ans = .COM(x, '_NewEnum', .dispatch = as.integer(2))
		 ans
} )
'COM.Workbooks.SetProperty'  = list( )
'COM.Workbooks.Methods'  = list('Add' = function( Template = NA ){
	
	ans = .COM(x, 'Add', Template, .dispatch = as.integer(1), .ids =181)
		 as(ans, 'Workbook')
},
'Close' = function(  ){
	ans = .COM(x, 'Close', .dispatch = as.integer(1), .ids =277)
		 ans
},
'_Open' = function( Filename, UpdateLinks = NA, ReadOnly = NA, Format = NA, Password = NA, WriteResPassword = NA, IgnoreReadOnlyRecommended = NA, Origin = NA, Delimiter = NA, Editable = NA, Notify = NA, Converter = NA, AddToMru = NA ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, '_Open', Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, .dispatch = as.integer(1), .ids =682)
		 as(ans, 'Workbook')
},
'__OpenText' = function( Filename, Origin = NA, StartRow = NA, DataType = NA, TextQualifier = NA, ConsecutiveDelimiter = NA, Tab = NA, Semicolon = NA, Comma = NA, Space = NA, Other = NA, OtherChar = NA, FieldInfo = NA, TextVisualLayout = NA ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	
	
	
	if(!missing( TextQualifier )) TextQualifier = as(TextQualifier, 'Constants')
	
	
	
	
	
	
	
	
	
	ans = .COM(x, '__OpenText', Filename, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, .dispatch = as.integer(1), .ids =683)
		 ans
},
'_OpenText' = function( Filename, Origin = NA, StartRow = NA, DataType = NA, TextQualifier = NA, ConsecutiveDelimiter = NA, Tab = NA, Semicolon = NA, Comma = NA, Space = NA, Other = NA, OtherChar = NA, FieldInfo = NA, TextVisualLayout = NA, DecimalSeparator = NA, ThousandsSeparator = NA ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	
	
	
	if(!missing( TextQualifier )) TextQualifier = as(TextQualifier, 'Constants')
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, '_OpenText', Filename, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator, .dispatch = as.integer(1), .ids =1773)
		 ans
},
'Open' = function( Filename, UpdateLinks = NA, ReadOnly = NA, Format = NA, Password = NA, WriteResPassword = NA, IgnoreReadOnlyRecommended = NA, Origin = NA, Delimiter = NA, Editable = NA, Notify = NA, Converter = NA, AddToMru = NA, Local = NA, CorruptLoad = NA ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Open', Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad, .dispatch = as.integer(1), .ids =1923)
		 as(ans, 'Workbook')
},
'OpenText' = function( Filename, Origin = NA, StartRow = NA, DataType = NA, TextQualifier = NA, ConsecutiveDelimiter = NA, Tab = NA, Semicolon = NA, Comma = NA, Space = NA, Other = NA, OtherChar = NA, FieldInfo = NA, TextVisualLayout = NA, DecimalSeparator = NA, ThousandsSeparator = NA, TrailingMinusNumbers = NA, Local = NA ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	
	
	
	if(!missing( TextQualifier )) TextQualifier = as(TextQualifier, 'Constants')
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'OpenText', Filename, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers, Local, .dispatch = as.integer(1), .ids =1924)
		 ans
},
'OpenDatabase' = function( Filename, CommandText = NA, CommandType = NA, BackgroundQuery = NA, ImportDataAs = NA ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	
	
	
	
	ans = .COM(x, 'OpenDatabase', Filename, CommandText, CommandType, BackgroundQuery, ImportDataAs, .dispatch = as.integer(1), .ids =2067)
		 as(ans, 'Workbook')
},
'CheckOut' = function( Filename ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	ans = .COM(x, 'CheckOut', Filename, .dispatch = as.integer(1), .ids =2069)
		 ans
},
'CanCheckOut' = function( Filename ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	ans = .COM(x, 'CanCheckOut', Filename, .dispatch = as.integer(1), .ids =2070)
	ans
},
'_OpenXML' = function( Filename, Stylesheets = NA ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	
	ans = .COM(x, '_OpenXML', Filename, Stylesheets, .dispatch = as.integer(1), .ids =2071)
		 as(ans, 'Workbook')
},
'OpenXML' = function( Filename, Stylesheets = NA, LoadOption = NA ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	
	
	ans = .COM(x, 'OpenXML', Filename, Stylesheets, LoadOption, .dispatch = as.integer(1), .ids =2280)
		 as(ans, 'Workbook')
} )
