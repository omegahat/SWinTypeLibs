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
setClass('_Document', contains = 'CompiledCOMIDispatch')
setClass('Documents', contains = 'COMTypedNamedList', prototype = list(name = '_Document'))
setClass('Range', contains = 'CompiledCOMIDispatch')
'WdMailSystemEnum' = c(
   'wdNoMailSystem' = 0,
   'wdMAPI' = 1,
   'wdPowerTalk' = 2,
   'wdMAPIandPowerTalk' = 3 
 )
storage.mode( WdMailSystemEnum ) = 'integer'
'WdTemplateTypeEnum' = c(
   'wdNormalTemplate' = 0,
   'wdGlobalTemplate' = 1,
   'wdAttachedTemplate' = 2 
 )
storage.mode( WdTemplateTypeEnum ) = 'integer'
'WdContinueEnum' = c(
   'wdContinueDisabled' = 0,
   'wdResetList' = 1,
   'wdContinueList' = 2 
 )
storage.mode( WdContinueEnum ) = 'integer'
'WdIMEModeEnum' = c(
   'wdIMEModeNoControl' = 0,
   'wdIMEModeOn' = 1,
   'wdIMEModeOff' = 2,
   'wdIMEModeHiragana' = 4,
   'wdIMEModeKatakana' = 5,
   'wdIMEModeKatakanaHalf' = 6,
   'wdIMEModeAlphaFull' = 7,
   'wdIMEModeAlpha' = 8,
   'wdIMEModeHangulFull' = 9,
   'wdIMEModeHangul' = 10 
 )
storage.mode( WdIMEModeEnum ) = 'integer'
'WdBaselineAlignmentEnum' = c(
   'wdBaselineAlignTop' = 0,
   'wdBaselineAlignCenter' = 1,
   'wdBaselineAlignBaseline' = 2,
   'wdBaselineAlignFarEast50' = 3,
   'wdBaselineAlignAuto' = 4 
 )
storage.mode( WdBaselineAlignmentEnum ) = 'integer'
'WdIndexFilterEnum' = c(
   'wdIndexFilterNone' = 0,
   'wdIndexFilterAiueo' = 1,
   'wdIndexFilterAkasatana' = 2,
   'wdIndexFilterChosung' = 3,
   'wdIndexFilterLow' = 4,
   'wdIndexFilterMedium' = 5,
   'wdIndexFilterFull' = 6 
 )
storage.mode( WdIndexFilterEnum ) = 'integer'
'WdIndexSortByEnum' = c(
   'wdIndexSortByStroke' = 0,
   'wdIndexSortBySyllable' = 1 
 )
storage.mode( WdIndexSortByEnum ) = 'integer'
'WdJustificationModeEnum' = c(
   'wdJustificationModeExpand' = 0,
   'wdJustificationModeCompress' = 1,
   'wdJustificationModeCompressKana' = 2 
 )
storage.mode( WdJustificationModeEnum ) = 'integer'
'WdFarEastLineBreakLevelEnum' = c(
   'wdFarEastLineBreakLevelNormal' = 0,
   'wdFarEastLineBreakLevelStrict' = 1,
   'wdFarEastLineBreakLevelCustom' = 2 
 )
storage.mode( WdFarEastLineBreakLevelEnum ) = 'integer'
'WdMultipleWordConversionsModeEnum' = c(
   'wdHangulToHanja' = 0,
   'wdHanjaToHangul' = 1 
 )
storage.mode( WdMultipleWordConversionsModeEnum ) = 'integer'
'WdColorIndexEnum' = c(
   'wdAuto' = 0,
   'wdBlack' = 1,
   'wdBlue' = 2,
   'wdTurquoise' = 3,
   'wdBrightGreen' = 4,
   'wdPink' = 5,
   'wdRed' = 6,
   'wdYellow' = 7,
   'wdWhite' = 8,
   'wdDarkBlue' = 9,
   'wdTeal' = 10,
   'wdGreen' = 11,
   'wdViolet' = 12,
   'wdDarkRed' = 13,
   'wdDarkYellow' = 14,
   'wdGray50' = 15,
   'wdGray25' = 16,
   'wdByAuthor' = -1,
   'wdNoHighlight' = 0 
 )
storage.mode( WdColorIndexEnum ) = 'integer'
'WdTextureIndexEnum' = c(
   'wdTextureNone' = 0,
   'wdTexture2Pt5Percent' = 25,
   'wdTexture5Percent' = 50,
   'wdTexture7Pt5Percent' = 75,
   'wdTexture10Percent' = 100,
   'wdTexture12Pt5Percent' = 125,
   'wdTexture15Percent' = 150,
   'wdTexture17Pt5Percent' = 175,
   'wdTexture20Percent' = 200,
   'wdTexture22Pt5Percent' = 225,
   'wdTexture25Percent' = 250,
   'wdTexture27Pt5Percent' = 275,
   'wdTexture30Percent' = 300,
   'wdTexture32Pt5Percent' = 325,
   'wdTexture35Percent' = 350,
   'wdTexture37Pt5Percent' = 375,
   'wdTexture40Percent' = 400,
   'wdTexture42Pt5Percent' = 425,
   'wdTexture45Percent' = 450,
   'wdTexture47Pt5Percent' = 475,
   'wdTexture50Percent' = 500,
   'wdTexture52Pt5Percent' = 525,
   'wdTexture55Percent' = 550,
   'wdTexture57Pt5Percent' = 575,
   'wdTexture60Percent' = 600,
   'wdTexture62Pt5Percent' = 625,
   'wdTexture65Percent' = 650,
   'wdTexture67Pt5Percent' = 675,
   'wdTexture70Percent' = 700,
   'wdTexture72Pt5Percent' = 725,
   'wdTexture75Percent' = 750,
   'wdTexture77Pt5Percent' = 775,
   'wdTexture80Percent' = 800,
   'wdTexture82Pt5Percent' = 825,
   'wdTexture85Percent' = 850,
   'wdTexture87Pt5Percent' = 875,
   'wdTexture90Percent' = 900,
   'wdTexture92Pt5Percent' = 925,
   'wdTexture95Percent' = 950,
   'wdTexture97Pt5Percent' = 975,
   'wdTextureSolid' = 1000,
   'wdTextureDarkHorizontal' = -1,
   'wdTextureDarkVertical' = -2,
   'wdTextureDarkDiagonalDown' = -3,
   'wdTextureDarkDiagonalUp' = -4,
   'wdTextureDarkCross' = -5,
   'wdTextureDarkDiagonalCross' = -6,
   'wdTextureHorizontal' = -7,
   'wdTextureVertical' = -8,
   'wdTextureDiagonalDown' = -9,
   'wdTextureDiagonalUp' = -10,
   'wdTextureCross' = -11,
   'wdTextureDiagonalCross' = -12 
 )
storage.mode( WdTextureIndexEnum ) = 'integer'
'WdUnderlineEnum' = c(
   'wdUnderlineNone' = 0,
   'wdUnderlineSingle' = 1,
   'wdUnderlineWords' = 2,
   'wdUnderlineDouble' = 3,
   'wdUnderlineDotted' = 4,
   'wdUnderlineThick' = 6,
   'wdUnderlineDash' = 7,
   'wdUnderlineDotDash' = 9,
   'wdUnderlineDotDotDash' = 10,
   'wdUnderlineWavy' = 11,
   'wdUnderlineWavyHeavy' = 27,
   'wdUnderlineDottedHeavy' = 20,
   'wdUnderlineDashHeavy' = 23,
   'wdUnderlineDotDashHeavy' = 25,
   'wdUnderlineDotDotDashHeavy' = 26,
   'wdUnderlineDashLong' = 39,
   'wdUnderlineDashLongHeavy' = 55,
   'wdUnderlineWavyDouble' = 43 
 )
storage.mode( WdUnderlineEnum ) = 'integer'
'WdEmphasisMarkEnum' = c(
   'wdEmphasisMarkNone' = 0,
   'wdEmphasisMarkOverSolidCircle' = 1,
   'wdEmphasisMarkOverComma' = 2,
   'wdEmphasisMarkOverWhiteCircle' = 3,
   'wdEmphasisMarkUnderSolidCircle' = 4 
 )
storage.mode( WdEmphasisMarkEnum ) = 'integer'
'WdInternationalIndexEnum' = c(
   'wdListSeparator' = 17,
   'wdDecimalSeparator' = 18,
   'wdThousandsSeparator' = 19,
   'wdCurrencyCode' = 20,
   'wd24HourClock' = 21,
   'wdInternationalAM' = 22,
   'wdInternationalPM' = 23,
   'wdTimeSeparator' = 24,
   'wdDateSeparator' = 25,
   'wdProductLanguageID' = 26 
 )
storage.mode( WdInternationalIndexEnum ) = 'integer'
'WdAutoMacrosEnum' = c(
   'wdAutoExec' = 0,
   'wdAutoNew' = 1,
   'wdAutoOpen' = 2,
   'wdAutoClose' = 3,
   'wdAutoExit' = 4,
   'wdAutoSync' = 5 
 )
storage.mode( WdAutoMacrosEnum ) = 'integer'
'WdCaptionPositionEnum' = c(
   'wdCaptionPositionAbove' = 0,
   'wdCaptionPositionBelow' = 1 
 )
storage.mode( WdCaptionPositionEnum ) = 'integer'
'WdCountryEnum' = c(
   'wdUS' = 1,
   'wdCanada' = 2,
   'wdLatinAmerica' = 3,
   'wdNetherlands' = 31,
   'wdFrance' = 33,
   'wdSpain' = 34,
   'wdItaly' = 39,
   'wdUK' = 44,
   'wdDenmark' = 45,
   'wdSweden' = 46,
   'wdNorway' = 47,
   'wdGermany' = 49,
   'wdPeru' = 51,
   'wdMexico' = 52,
   'wdArgentina' = 54,
   'wdBrazil' = 55,
   'wdChile' = 56,
   'wdVenezuela' = 58,
   'wdJapan' = 81,
   'wdTaiwan' = 886,
   'wdChina' = 86,
   'wdKorea' = 82,
   'wdFinland' = 358,
   'wdIceland' = 354 
 )
storage.mode( WdCountryEnum ) = 'integer'
'WdHeadingSeparatorEnum' = c(
   'wdHeadingSeparatorNone' = 0,
   'wdHeadingSeparatorBlankLine' = 1,
   'wdHeadingSeparatorLetter' = 2,
   'wdHeadingSeparatorLetterLow' = 3,
   'wdHeadingSeparatorLetterFull' = 4 
 )
storage.mode( WdHeadingSeparatorEnum ) = 'integer'
'WdSeparatorTypeEnum' = c(
   'wdSeparatorHyphen' = 0,
   'wdSeparatorPeriod' = 1,
   'wdSeparatorColon' = 2,
   'wdSeparatorEmDash' = 3,
   'wdSeparatorEnDash' = 4 
 )
storage.mode( WdSeparatorTypeEnum ) = 'integer'
'WdPageNumberAlignmentEnum' = c(
   'wdAlignPageNumberLeft' = 0,
   'wdAlignPageNumberCenter' = 1,
   'wdAlignPageNumberRight' = 2,
   'wdAlignPageNumberInside' = 3,
   'wdAlignPageNumberOutside' = 4 
 )
storage.mode( WdPageNumberAlignmentEnum ) = 'integer'
'WdBorderTypeEnum' = c(
   'wdBorderTop' = -1,
   'wdBorderLeft' = -2,
   'wdBorderBottom' = -3,
   'wdBorderRight' = -4,
   'wdBorderHorizontal' = -5,
   'wdBorderVertical' = -6,
   'wdBorderDiagonalDown' = -7,
   'wdBorderDiagonalUp' = -8 
 )
storage.mode( WdBorderTypeEnum ) = 'integer'
'WdBorderTypeHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdBorderTypeHIDEnum ) = 'integer'
'WdFramePositionEnum' = c(
   'wdFrameTop' = -999999,
   'wdFrameLeft' = -999998,
   'wdFrameBottom' = -999997,
   'wdFrameRight' = -999996,
   'wdFrameCenter' = -999995,
   'wdFrameInside' = -999994,
   'wdFrameOutside' = -999993 
 )
storage.mode( WdFramePositionEnum ) = 'integer'
'WdAnimationEnum' = c(
   'wdAnimationNone' = 0,
   'wdAnimationLasVegasLights' = 1,
   'wdAnimationBlinkingBackground' = 2,
   'wdAnimationSparkleText' = 3,
   'wdAnimationMarchingBlackAnts' = 4,
   'wdAnimationMarchingRedAnts' = 5,
   'wdAnimationShimmer' = 6 
 )
storage.mode( WdAnimationEnum ) = 'integer'
'WdCharacterCaseEnum' = c(
   'wdNextCase' = -1,
   'wdLowerCase' = 0,
   'wdUpperCase' = 1,
   'wdTitleWord' = 2,
   'wdTitleSentence' = 4,
   'wdToggleCase' = 5,
   'wdHalfWidth' = 6,
   'wdFullWidth' = 7,
   'wdKatakana' = 8,
   'wdHiragana' = 9 
 )
storage.mode( WdCharacterCaseEnum ) = 'integer'
'WdCharacterCaseHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdCharacterCaseHIDEnum ) = 'integer'
'WdSummaryModeEnum' = c(
   'wdSummaryModeHighlight' = 0,
   'wdSummaryModeHideAllButSummary' = 1,
   'wdSummaryModeInsert' = 2,
   'wdSummaryModeCreateNew' = 3 
 )
storage.mode( WdSummaryModeEnum ) = 'integer'
'WdSummaryLengthEnum' = c(
   'wd10Sentences' = -2,
   'wd20Sentences' = -3,
   'wd100Words' = -4,
   'wd500Words' = -5,
   'wd10Percent' = -6,
   'wd25Percent' = -7,
   'wd50Percent' = -8,
   'wd75Percent' = -9 
 )
storage.mode( WdSummaryLengthEnum ) = 'integer'
'WdStyleTypeEnum' = c(
   'wdStyleTypeParagraph' = 1,
   'wdStyleTypeCharacter' = 2,
   'wdStyleTypeTable' = 3,
   'wdStyleTypeList' = 4 
 )
storage.mode( WdStyleTypeEnum ) = 'integer'
'WdUnitsEnum' = c(
   'wdCharacter' = 1,
   'wdWord' = 2,
   'wdSentence' = 3,
   'wdParagraph' = 4,
   'wdLine' = 5,
   'wdStory' = 6,
   'wdScreen' = 7,
   'wdSection' = 8,
   'wdColumn' = 9,
   'wdRow' = 10,
   'wdWindow' = 11,
   'wdCell' = 12,
   'wdCharacterFormatting' = 13,
   'wdParagraphFormatting' = 14,
   'wdTable' = 15,
   'wdItem' = 16 
 )
storage.mode( WdUnitsEnum ) = 'integer'
'WdGoToItemEnum' = c(
   'wdGoToBookmark' = -1,
   'wdGoToSection' = 0,
   'wdGoToPage' = 1,
   'wdGoToTable' = 2,
   'wdGoToLine' = 3,
   'wdGoToFootnote' = 4,
   'wdGoToEndnote' = 5,
   'wdGoToComment' = 6,
   'wdGoToField' = 7,
   'wdGoToGraphic' = 8,
   'wdGoToObject' = 9,
   'wdGoToEquation' = 10,
   'wdGoToHeading' = 11,
   'wdGoToPercent' = 12,
   'wdGoToSpellingError' = 13,
   'wdGoToGrammaticalError' = 14,
   'wdGoToProofreadingError' = 15 
 )
storage.mode( WdGoToItemEnum ) = 'integer'
'WdGoToDirectionEnum' = c(
   'wdGoToFirst' = 1,
   'wdGoToLast' = -1,
   'wdGoToNext' = 2,
   'wdGoToRelative' = 2,
   'wdGoToPrevious' = 3,
   'wdGoToAbsolute' = 1 
 )
storage.mode( WdGoToDirectionEnum ) = 'integer'
'WdCollapseDirectionEnum' = c(
   'wdCollapseStart' = 1,
   'wdCollapseEnd' = 0 
 )
storage.mode( WdCollapseDirectionEnum ) = 'integer'
'WdRowHeightRuleEnum' = c(
   'wdRowHeightAuto' = 0,
   'wdRowHeightAtLeast' = 1,
   'wdRowHeightExactly' = 2 
 )
storage.mode( WdRowHeightRuleEnum ) = 'integer'
'WdFrameSizeRuleEnum' = c(
   'wdFrameAuto' = 0,
   'wdFrameAtLeast' = 1,
   'wdFrameExact' = 2 
 )
storage.mode( WdFrameSizeRuleEnum ) = 'integer'
'WdInsertCellsEnum' = c(
   'wdInsertCellsShiftRight' = 0,
   'wdInsertCellsShiftDown' = 1,
   'wdInsertCellsEntireRow' = 2,
   'wdInsertCellsEntireColumn' = 3 
 )
storage.mode( WdInsertCellsEnum ) = 'integer'
'WdDeleteCellsEnum' = c(
   'wdDeleteCellsShiftLeft' = 0,
   'wdDeleteCellsShiftUp' = 1,
   'wdDeleteCellsEntireRow' = 2,
   'wdDeleteCellsEntireColumn' = 3 
 )
storage.mode( WdDeleteCellsEnum ) = 'integer'
'WdListApplyToEnum' = c(
   'wdListApplyToWholeList' = 0,
   'wdListApplyToThisPointForward' = 1,
   'wdListApplyToSelection' = 2 
 )
storage.mode( WdListApplyToEnum ) = 'integer'
'WdAlertLevelEnum' = c(
   'wdAlertsNone' = 0,
   'wdAlertsMessageBox' = -2,
   'wdAlertsAll' = -1 
 )
storage.mode( WdAlertLevelEnum ) = 'integer'
'WdCursorTypeEnum' = c(
   'wdCursorWait' = 0,
   'wdCursorIBeam' = 1,
   'wdCursorNormal' = 2,
   'wdCursorNorthwestArrow' = 3 
 )
storage.mode( WdCursorTypeEnum ) = 'integer'
'WdEnableCancelKeyEnum' = c(
   'wdCancelDisabled' = 0,
   'wdCancelInterrupt' = 1 
 )
storage.mode( WdEnableCancelKeyEnum ) = 'integer'
'WdRulerStyleEnum' = c(
   'wdAdjustNone' = 0,
   'wdAdjustProportional' = 1,
   'wdAdjustFirstColumn' = 2,
   'wdAdjustSameWidth' = 3 
 )
storage.mode( WdRulerStyleEnum ) = 'integer'
'WdParagraphAlignmentEnum' = c(
   'wdAlignParagraphLeft' = 0,
   'wdAlignParagraphCenter' = 1,
   'wdAlignParagraphRight' = 2,
   'wdAlignParagraphJustify' = 3,
   'wdAlignParagraphDistribute' = 4,
   'wdAlignParagraphJustifyMed' = 5,
   'wdAlignParagraphJustifyHi' = 7,
   'wdAlignParagraphJustifyLow' = 8,
   'wdAlignParagraphThaiJustify' = 9 
 )
storage.mode( WdParagraphAlignmentEnum ) = 'integer'
'WdParagraphAlignmentHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdParagraphAlignmentHIDEnum ) = 'integer'
'WdListLevelAlignmentEnum' = c(
   'wdListLevelAlignLeft' = 0,
   'wdListLevelAlignCenter' = 1,
   'wdListLevelAlignRight' = 2 
 )
storage.mode( WdListLevelAlignmentEnum ) = 'integer'
'WdRowAlignmentEnum' = c(
   'wdAlignRowLeft' = 0,
   'wdAlignRowCenter' = 1,
   'wdAlignRowRight' = 2 
 )
storage.mode( WdRowAlignmentEnum ) = 'integer'
'WdTabAlignmentEnum' = c(
   'wdAlignTabLeft' = 0,
   'wdAlignTabCenter' = 1,
   'wdAlignTabRight' = 2,
   'wdAlignTabDecimal' = 3,
   'wdAlignTabBar' = 4,
   'wdAlignTabList' = 6 
 )
storage.mode( WdTabAlignmentEnum ) = 'integer'
'WdVerticalAlignmentEnum' = c(
   'wdAlignVerticalTop' = 0,
   'wdAlignVerticalCenter' = 1,
   'wdAlignVerticalJustify' = 2,
   'wdAlignVerticalBottom' = 3 
 )
storage.mode( WdVerticalAlignmentEnum ) = 'integer'
'WdCellVerticalAlignmentEnum' = c(
   'wdCellAlignVerticalTop' = 0,
   'wdCellAlignVerticalCenter' = 1,
   'wdCellAlignVerticalBottom' = 3 
 )
storage.mode( WdCellVerticalAlignmentEnum ) = 'integer'
'WdTrailingCharacterEnum' = c(
   'wdTrailingTab' = 0,
   'wdTrailingSpace' = 1,
   'wdTrailingNone' = 2 
 )
storage.mode( WdTrailingCharacterEnum ) = 'integer'
'WdListGalleryTypeEnum' = c(
   'wdBulletGallery' = 1,
   'wdNumberGallery' = 2,
   'wdOutlineNumberGallery' = 3 
 )
storage.mode( WdListGalleryTypeEnum ) = 'integer'
'WdListNumberStyleEnum' = c(
   'wdListNumberStyleArabic' = 0,
   'wdListNumberStyleUppercaseRoman' = 1,
   'wdListNumberStyleLowercaseRoman' = 2,
   'wdListNumberStyleUppercaseLetter' = 3,
   'wdListNumberStyleLowercaseLetter' = 4,
   'wdListNumberStyleOrdinal' = 5,
   'wdListNumberStyleCardinalText' = 6,
   'wdListNumberStyleOrdinalText' = 7,
   'wdListNumberStyleKanji' = 10,
   'wdListNumberStyleKanjiDigit' = 11,
   'wdListNumberStyleAiueoHalfWidth' = 12,
   'wdListNumberStyleIrohaHalfWidth' = 13,
   'wdListNumberStyleArabicFullWidth' = 14,
   'wdListNumberStyleKanjiTraditional' = 16,
   'wdListNumberStyleKanjiTraditional2' = 17,
   'wdListNumberStyleNumberInCircle' = 18,
   'wdListNumberStyleAiueo' = 20,
   'wdListNumberStyleIroha' = 21,
   'wdListNumberStyleArabicLZ' = 22,
   'wdListNumberStyleBullet' = 23,
   'wdListNumberStyleGanada' = 24,
   'wdListNumberStyleChosung' = 25,
   'wdListNumberStyleGBNum1' = 26,
   'wdListNumberStyleGBNum2' = 27,
   'wdListNumberStyleGBNum3' = 28,
   'wdListNumberStyleGBNum4' = 29,
   'wdListNumberStyleZodiac1' = 30,
   'wdListNumberStyleZodiac2' = 31,
   'wdListNumberStyleZodiac3' = 32,
   'wdListNumberStyleTradChinNum1' = 33,
   'wdListNumberStyleTradChinNum2' = 34,
   'wdListNumberStyleTradChinNum3' = 35,
   'wdListNumberStyleTradChinNum4' = 36,
   'wdListNumberStyleSimpChinNum1' = 37,
   'wdListNumberStyleSimpChinNum2' = 38,
   'wdListNumberStyleSimpChinNum3' = 39,
   'wdListNumberStyleSimpChinNum4' = 40,
   'wdListNumberStyleHanjaRead' = 41,
   'wdListNumberStyleHanjaReadDigit' = 42,
   'wdListNumberStyleHangul' = 43,
   'wdListNumberStyleHanja' = 44,
   'wdListNumberStyleHebrew1' = 45,
   'wdListNumberStyleArabic1' = 46,
   'wdListNumberStyleHebrew2' = 47,
   'wdListNumberStyleArabic2' = 48,
   'wdListNumberStyleHindiLetter1' = 49,
   'wdListNumberStyleHindiLetter2' = 50,
   'wdListNumberStyleHindiArabic' = 51,
   'wdListNumberStyleHindiCardinalText' = 52,
   'wdListNumberStyleThaiLetter' = 53,
   'wdListNumberStyleThaiArabic' = 54,
   'wdListNumberStyleThaiCardinalText' = 55,
   'wdListNumberStyleVietCardinalText' = 56,
   'wdListNumberStyleLowercaseRussian' = 58,
   'wdListNumberStyleUppercaseRussian' = 59,
   'wdListNumberStylePictureBullet' = 249,
   'wdListNumberStyleLegal' = 253,
   'wdListNumberStyleLegalLZ' = 254,
   'wdListNumberStyleNone' = 255 
 )
storage.mode( WdListNumberStyleEnum ) = 'integer'
'WdListNumberStyleHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdListNumberStyleHIDEnum ) = 'integer'
'WdNoteNumberStyleEnum' = c(
   'wdNoteNumberStyleArabic' = 0,
   'wdNoteNumberStyleUppercaseRoman' = 1,
   'wdNoteNumberStyleLowercaseRoman' = 2,
   'wdNoteNumberStyleUppercaseLetter' = 3,
   'wdNoteNumberStyleLowercaseLetter' = 4,
   'wdNoteNumberStyleSymbol' = 9,
   'wdNoteNumberStyleArabicFullWidth' = 14,
   'wdNoteNumberStyleKanji' = 10,
   'wdNoteNumberStyleKanjiDigit' = 11,
   'wdNoteNumberStyleKanjiTraditional' = 16,
   'wdNoteNumberStyleNumberInCircle' = 18,
   'wdNoteNumberStyleHanjaRead' = 41,
   'wdNoteNumberStyleHanjaReadDigit' = 42,
   'wdNoteNumberStyleTradChinNum1' = 33,
   'wdNoteNumberStyleTradChinNum2' = 34,
   'wdNoteNumberStyleSimpChinNum1' = 37,
   'wdNoteNumberStyleSimpChinNum2' = 38,
   'wdNoteNumberStyleHebrewLetter1' = 45,
   'wdNoteNumberStyleArabicLetter1' = 46,
   'wdNoteNumberStyleHebrewLetter2' = 47,
   'wdNoteNumberStyleArabicLetter2' = 48,
   'wdNoteNumberStyleHindiLetter1' = 49,
   'wdNoteNumberStyleHindiLetter2' = 50,
   'wdNoteNumberStyleHindiArabic' = 51,
   'wdNoteNumberStyleHindiCardinalText' = 52,
   'wdNoteNumberStyleThaiLetter' = 53,
   'wdNoteNumberStyleThaiArabic' = 54,
   'wdNoteNumberStyleThaiCardinalText' = 55,
   'wdNoteNumberStyleVietCardinalText' = 56 
 )
storage.mode( WdNoteNumberStyleEnum ) = 'integer'
'WdNoteNumberStyleHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdNoteNumberStyleHIDEnum ) = 'integer'
'WdCaptionNumberStyleEnum' = c(
   'wdCaptionNumberStyleArabic' = 0,
   'wdCaptionNumberStyleUppercaseRoman' = 1,
   'wdCaptionNumberStyleLowercaseRoman' = 2,
   'wdCaptionNumberStyleUppercaseLetter' = 3,
   'wdCaptionNumberStyleLowercaseLetter' = 4,
   'wdCaptionNumberStyleArabicFullWidth' = 14,
   'wdCaptionNumberStyleKanji' = 10,
   'wdCaptionNumberStyleKanjiDigit' = 11,
   'wdCaptionNumberStyleKanjiTraditional' = 16,
   'wdCaptionNumberStyleNumberInCircle' = 18,
   'wdCaptionNumberStyleGanada' = 24,
   'wdCaptionNumberStyleChosung' = 25,
   'wdCaptionNumberStyleZodiac1' = 30,
   'wdCaptionNumberStyleZodiac2' = 31,
   'wdCaptionNumberStyleHanjaRead' = 41,
   'wdCaptionNumberStyleHanjaReadDigit' = 42,
   'wdCaptionNumberStyleTradChinNum2' = 34,
   'wdCaptionNumberStyleTradChinNum3' = 35,
   'wdCaptionNumberStyleSimpChinNum2' = 38,
   'wdCaptionNumberStyleSimpChinNum3' = 39,
   'wdCaptionNumberStyleHebrewLetter1' = 45,
   'wdCaptionNumberStyleArabicLetter1' = 46,
   'wdCaptionNumberStyleHebrewLetter2' = 47,
   'wdCaptionNumberStyleArabicLetter2' = 48,
   'wdCaptionNumberStyleHindiLetter1' = 49,
   'wdCaptionNumberStyleHindiLetter2' = 50,
   'wdCaptionNumberStyleHindiArabic' = 51,
   'wdCaptionNumberStyleHindiCardinalText' = 52,
   'wdCaptionNumberStyleThaiLetter' = 53,
   'wdCaptionNumberStyleThaiArabic' = 54,
   'wdCaptionNumberStyleThaiCardinalText' = 55,
   'wdCaptionNumberStyleVietCardinalText' = 56 
 )
storage.mode( WdCaptionNumberStyleEnum ) = 'integer'
'WdCaptionNumberStyleHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdCaptionNumberStyleHIDEnum ) = 'integer'
'WdPageNumberStyleEnum' = c(
   'wdPageNumberStyleArabic' = 0,
   'wdPageNumberStyleUppercaseRoman' = 1,
   'wdPageNumberStyleLowercaseRoman' = 2,
   'wdPageNumberStyleUppercaseLetter' = 3,
   'wdPageNumberStyleLowercaseLetter' = 4,
   'wdPageNumberStyleArabicFullWidth' = 14,
   'wdPageNumberStyleKanji' = 10,
   'wdPageNumberStyleKanjiDigit' = 11,
   'wdPageNumberStyleKanjiTraditional' = 16,
   'wdPageNumberStyleNumberInCircle' = 18,
   'wdPageNumberStyleHanjaRead' = 41,
   'wdPageNumberStyleHanjaReadDigit' = 42,
   'wdPageNumberStyleTradChinNum1' = 33,
   'wdPageNumberStyleTradChinNum2' = 34,
   'wdPageNumberStyleSimpChinNum1' = 37,
   'wdPageNumberStyleSimpChinNum2' = 38,
   'wdPageNumberStyleHebrewLetter1' = 45,
   'wdPageNumberStyleArabicLetter1' = 46,
   'wdPageNumberStyleHebrewLetter2' = 47,
   'wdPageNumberStyleArabicLetter2' = 48,
   'wdPageNumberStyleHindiLetter1' = 49,
   'wdPageNumberStyleHindiLetter2' = 50,
   'wdPageNumberStyleHindiArabic' = 51,
   'wdPageNumberStyleHindiCardinalText' = 52,
   'wdPageNumberStyleThaiLetter' = 53,
   'wdPageNumberStyleThaiArabic' = 54,
   'wdPageNumberStyleThaiCardinalText' = 55,
   'wdPageNumberStyleVietCardinalText' = 56,
   'wdPageNumberStyleNumberInDash' = 57 
 )
storage.mode( WdPageNumberStyleEnum ) = 'integer'
'WdPageNumberStyleHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdPageNumberStyleHIDEnum ) = 'integer'
'WdStatisticEnum' = c(
   'wdStatisticWords' = 0,
   'wdStatisticLines' = 1,
   'wdStatisticPages' = 2,
   'wdStatisticCharacters' = 3,
   'wdStatisticParagraphs' = 4,
   'wdStatisticCharactersWithSpaces' = 5,
   'wdStatisticFarEastCharacters' = 6 
 )
storage.mode( WdStatisticEnum ) = 'integer'
'WdStatisticHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdStatisticHIDEnum ) = 'integer'
'WdBuiltInPropertyEnum' = c(
   'wdPropertyTitle' = 1,
   'wdPropertySubject' = 2,
   'wdPropertyAuthor' = 3,
   'wdPropertyKeywords' = 4,
   'wdPropertyComments' = 5,
   'wdPropertyTemplate' = 6,
   'wdPropertyLastAuthor' = 7,
   'wdPropertyRevision' = 8,
   'wdPropertyAppName' = 9,
   'wdPropertyTimeLastPrinted' = 10,
   'wdPropertyTimeCreated' = 11,
   'wdPropertyTimeLastSaved' = 12,
   'wdPropertyVBATotalEdit' = 13,
   'wdPropertyPages' = 14,
   'wdPropertyWords' = 15,
   'wdPropertyCharacters' = 16,
   'wdPropertySecurity' = 17,
   'wdPropertyCategory' = 18,
   'wdPropertyFormat' = 19,
   'wdPropertyManager' = 20,
   'wdPropertyCompany' = 21,
   'wdPropertyBytes' = 22,
   'wdPropertyLines' = 23,
   'wdPropertyParas' = 24,
   'wdPropertySlides' = 25,
   'wdPropertyNotes' = 26,
   'wdPropertyHiddenSlides' = 27,
   'wdPropertyMMClips' = 28,
   'wdPropertyHyperlinkBase' = 29,
   'wdPropertyCharsWSpaces' = 30 
 )
storage.mode( WdBuiltInPropertyEnum ) = 'integer'
'WdLineSpacingEnum' = c(
   'wdLineSpaceSingle' = 0,
   'wdLineSpace1pt5' = 1,
   'wdLineSpaceDouble' = 2,
   'wdLineSpaceAtLeast' = 3,
   'wdLineSpaceExactly' = 4,
   'wdLineSpaceMultiple' = 5 
 )
storage.mode( WdLineSpacingEnum ) = 'integer'
'WdNumberTypeEnum' = c(
   'wdNumberParagraph' = 1,
   'wdNumberListNum' = 2,
   'wdNumberAllNumbers' = 3 
 )
storage.mode( WdNumberTypeEnum ) = 'integer'
'WdListTypeEnum' = c(
   'wdListNoNumbering' = 0,
   'wdListListNumOnly' = 1,
   'wdListBullet' = 2,
   'wdListSimpleNumbering' = 3,
   'wdListOutlineNumbering' = 4,
   'wdListMixedNumbering' = 5,
   'wdListPictureBullet' = 6 
 )
storage.mode( WdListTypeEnum ) = 'integer'
'WdStoryTypeEnum' = c(
   'wdMainTextStory' = 1,
   'wdFootnotesStory' = 2,
   'wdEndnotesStory' = 3,
   'wdCommentsStory' = 4,
   'wdTextFrameStory' = 5,
   'wdEvenPagesHeaderStory' = 6,
   'wdPrimaryHeaderStory' = 7,
   'wdEvenPagesFooterStory' = 8,
   'wdPrimaryFooterStory' = 9,
   'wdFirstPageHeaderStory' = 10,
   'wdFirstPageFooterStory' = 11,
   'wdFootnoteSeparatorStory' = 12,
   'wdFootnoteContinuationSeparatorStory' = 13,
   'wdFootnoteContinuationNoticeStory' = 14,
   'wdEndnoteSeparatorStory' = 15,
   'wdEndnoteContinuationSeparatorStory' = 16,
   'wdEndnoteContinuationNoticeStory' = 17 
 )
storage.mode( WdStoryTypeEnum ) = 'integer'
'WdSaveFormatEnum' = c(
   'wdFormatDocument' = 0,
   'wdFormatTemplate' = 1,
   'wdFormatText' = 2,
   'wdFormatTextLineBreaks' = 3,
   'wdFormatDOSText' = 4,
   'wdFormatDOSTextLineBreaks' = 5,
   'wdFormatRTF' = 6,
   'wdFormatUnicodeText' = 7,
   'wdFormatEncodedText' = 7,
   'wdFormatHTML' = 8,
   'wdFormatWebArchive' = 9,
   'wdFormatFilteredHTML' = 10,
   'wdFormatXML' = 11 
 )
storage.mode( WdSaveFormatEnum ) = 'integer'
'WdOpenFormatEnum' = c(
   'wdOpenFormatAuto' = 0,
   'wdOpenFormatDocument' = 1,
   'wdOpenFormatTemplate' = 2,
   'wdOpenFormatRTF' = 3,
   'wdOpenFormatText' = 4,
   'wdOpenFormatUnicodeText' = 5,
   'wdOpenFormatEncodedText' = 5,
   'wdOpenFormatAllWord' = 6,
   'wdOpenFormatWebPages' = 7,
   'wdOpenFormatXML' = 8 
 )
storage.mode( WdOpenFormatEnum ) = 'integer'
'WdHeaderFooterIndexEnum' = c(
   'wdHeaderFooterPrimary' = 1,
   'wdHeaderFooterFirstPage' = 2,
   'wdHeaderFooterEvenPages' = 3 
 )
storage.mode( WdHeaderFooterIndexEnum ) = 'integer'
'WdTocFormatEnum' = c(
   'wdTOCTemplate' = 0,
   'wdTOCClassic' = 1,
   'wdTOCDistinctive' = 2,
   'wdTOCFancy' = 3,
   'wdTOCModern' = 4,
   'wdTOCFormal' = 5,
   'wdTOCSimple' = 6 
 )
storage.mode( WdTocFormatEnum ) = 'integer'
'WdTofFormatEnum' = c(
   'wdTOFTemplate' = 0,
   'wdTOFClassic' = 1,
   'wdTOFDistinctive' = 2,
   'wdTOFCentered' = 3,
   'wdTOFFormal' = 4,
   'wdTOFSimple' = 5 
 )
storage.mode( WdTofFormatEnum ) = 'integer'
'WdToaFormatEnum' = c(
   'wdTOATemplate' = 0,
   'wdTOAClassic' = 1,
   'wdTOADistinctive' = 2,
   'wdTOAFormal' = 3,
   'wdTOASimple' = 4 
 )
storage.mode( WdToaFormatEnum ) = 'integer'
'WdLineStyleEnum' = c(
   'wdLineStyleNone' = 0,
   'wdLineStyleSingle' = 1,
   'wdLineStyleDot' = 2,
   'wdLineStyleDashSmallGap' = 3,
   'wdLineStyleDashLargeGap' = 4,
   'wdLineStyleDashDot' = 5,
   'wdLineStyleDashDotDot' = 6,
   'wdLineStyleDouble' = 7,
   'wdLineStyleTriple' = 8,
   'wdLineStyleThinThickSmallGap' = 9,
   'wdLineStyleThickThinSmallGap' = 10,
   'wdLineStyleThinThickThinSmallGap' = 11,
   'wdLineStyleThinThickMedGap' = 12,
   'wdLineStyleThickThinMedGap' = 13,
   'wdLineStyleThinThickThinMedGap' = 14,
   'wdLineStyleThinThickLargeGap' = 15,
   'wdLineStyleThickThinLargeGap' = 16,
   'wdLineStyleThinThickThinLargeGap' = 17,
   'wdLineStyleSingleWavy' = 18,
   'wdLineStyleDoubleWavy' = 19,
   'wdLineStyleDashDotStroked' = 20,
   'wdLineStyleEmboss3D' = 21,
   'wdLineStyleEngrave3D' = 22,
   'wdLineStyleOutset' = 23,
   'wdLineStyleInset' = 24 
 )
storage.mode( WdLineStyleEnum ) = 'integer'
'WdLineWidthEnum' = c(
   'wdLineWidth025pt' = 2,
   'wdLineWidth050pt' = 4,
   'wdLineWidth075pt' = 6,
   'wdLineWidth100pt' = 8,
   'wdLineWidth150pt' = 12,
   'wdLineWidth225pt' = 18,
   'wdLineWidth300pt' = 24,
   'wdLineWidth450pt' = 36,
   'wdLineWidth600pt' = 48 
 )
storage.mode( WdLineWidthEnum ) = 'integer'
'WdBreakTypeEnum' = c(
   'wdSectionBreakNextPage' = 2,
   'wdSectionBreakContinuous' = 3,
   'wdSectionBreakEvenPage' = 4,
   'wdSectionBreakOddPage' = 5,
   'wdLineBreak' = 6,
   'wdPageBreak' = 7,
   'wdColumnBreak' = 8,
   'wdLineBreakClearLeft' = 9,
   'wdLineBreakClearRight' = 10,
   'wdTextWrappingBreak' = 11 
 )
storage.mode( WdBreakTypeEnum ) = 'integer'
'WdTabLeaderEnum' = c(
   'wdTabLeaderSpaces' = 0,
   'wdTabLeaderDots' = 1,
   'wdTabLeaderDashes' = 2,
   'wdTabLeaderLines' = 3,
   'wdTabLeaderHeavy' = 4,
   'wdTabLeaderMiddleDot' = 5 
 )
storage.mode( WdTabLeaderEnum ) = 'integer'
'WdTabLeaderHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdTabLeaderHIDEnum ) = 'integer'
'WdMeasurementUnitsEnum' = c(
   'wdInches' = 0,
   'wdCentimeters' = 1,
   'wdMillimeters' = 2,
   'wdPoints' = 3,
   'wdPicas' = 4 
 )
storage.mode( WdMeasurementUnitsEnum ) = 'integer'
'WdMeasurementUnitsHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdMeasurementUnitsHIDEnum ) = 'integer'
'WdDropPositionEnum' = c(
   'wdDropNone' = 0,
   'wdDropNormal' = 1,
   'wdDropMargin' = 2 
 )
storage.mode( WdDropPositionEnum ) = 'integer'
'WdNumberingRuleEnum' = c(
   'wdRestartContinuous' = 0,
   'wdRestartSection' = 1,
   'wdRestartPage' = 2 
 )
storage.mode( WdNumberingRuleEnum ) = 'integer'
'WdFootnoteLocationEnum' = c(
   'wdBottomOfPage' = 0,
   'wdBeneathText' = 1 
 )
storage.mode( WdFootnoteLocationEnum ) = 'integer'
'WdEndnoteLocationEnum' = c(
   'wdEndOfSection' = 0,
   'wdEndOfDocument' = 1 
 )
storage.mode( WdEndnoteLocationEnum ) = 'integer'
'WdSortSeparatorEnum' = c(
   'wdSortSeparateByTabs' = 0,
   'wdSortSeparateByCommas' = 1,
   'wdSortSeparateByDefaultTableSeparator' = 2 
 )
storage.mode( WdSortSeparatorEnum ) = 'integer'
'WdTableFieldSeparatorEnum' = c(
   'wdSeparateByParagraphs' = 0,
   'wdSeparateByTabs' = 1,
   'wdSeparateByCommas' = 2,
   'wdSeparateByDefaultListSeparator' = 3 
 )
storage.mode( WdTableFieldSeparatorEnum ) = 'integer'
'WdSortFieldTypeEnum' = c(
   'wdSortFieldAlphanumeric' = 0,
   'wdSortFieldNumeric' = 1,
   'wdSortFieldDate' = 2,
   'wdSortFieldSyllable' = 3,
   'wdSortFieldJapanJIS' = 4,
   'wdSortFieldStroke' = 5,
   'wdSortFieldKoreaKS' = 6 
 )
storage.mode( WdSortFieldTypeEnum ) = 'integer'
'WdSortFieldTypeHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdSortFieldTypeHIDEnum ) = 'integer'
'WdSortOrderEnum' = c(
   'wdSortOrderAscending' = 0,
   'wdSortOrderDescending' = 1 
 )
storage.mode( WdSortOrderEnum ) = 'integer'
'WdTableFormatEnum' = c(
   'wdTableFormatNone' = 0,
   'wdTableFormatSimple1' = 1,
   'wdTableFormatSimple2' = 2,
   'wdTableFormatSimple3' = 3,
   'wdTableFormatClassic1' = 4,
   'wdTableFormatClassic2' = 5,
   'wdTableFormatClassic3' = 6,
   'wdTableFormatClassic4' = 7,
   'wdTableFormatColorful1' = 8,
   'wdTableFormatColorful2' = 9,
   'wdTableFormatColorful3' = 10,
   'wdTableFormatColumns1' = 11,
   'wdTableFormatColumns2' = 12,
   'wdTableFormatColumns3' = 13,
   'wdTableFormatColumns4' = 14,
   'wdTableFormatColumns5' = 15,
   'wdTableFormatGrid1' = 16,
   'wdTableFormatGrid2' = 17,
   'wdTableFormatGrid3' = 18,
   'wdTableFormatGrid4' = 19,
   'wdTableFormatGrid5' = 20,
   'wdTableFormatGrid6' = 21,
   'wdTableFormatGrid7' = 22,
   'wdTableFormatGrid8' = 23,
   'wdTableFormatList1' = 24,
   'wdTableFormatList2' = 25,
   'wdTableFormatList3' = 26,
   'wdTableFormatList4' = 27,
   'wdTableFormatList5' = 28,
   'wdTableFormatList6' = 29,
   'wdTableFormatList7' = 30,
   'wdTableFormatList8' = 31,
   'wdTableFormat3DEffects1' = 32,
   'wdTableFormat3DEffects2' = 33,
   'wdTableFormat3DEffects3' = 34,
   'wdTableFormatContemporary' = 35,
   'wdTableFormatElegant' = 36,
   'wdTableFormatProfessional' = 37,
   'wdTableFormatSubtle1' = 38,
   'wdTableFormatSubtle2' = 39,
   'wdTableFormatWeb1' = 40,
   'wdTableFormatWeb2' = 41,
   'wdTableFormatWeb3' = 42 
 )
storage.mode( WdTableFormatEnum ) = 'integer'
'WdTableFormatApplyEnum' = c(
   'wdTableFormatApplyBorders' = 1,
   'wdTableFormatApplyShading' = 2,
   'wdTableFormatApplyFont' = 4,
   'wdTableFormatApplyColor' = 8,
   'wdTableFormatApplyAutoFit' = 16,
   'wdTableFormatApplyHeadingRows' = 32,
   'wdTableFormatApplyLastRow' = 64,
   'wdTableFormatApplyFirstColumn' = 128,
   'wdTableFormatApplyLastColumn' = 256 
 )
storage.mode( WdTableFormatApplyEnum ) = 'integer'
'WdLanguageIDEnum' = c(
   'wdLanguageNone' = 0,
   'wdNoProofing' = 1024,
   'wdAfrikaans' = 1078,
   'wdAlbanian' = 1052,
   'wdAmharic' = 1118,
   'wdArabicAlgeria' = 5121,
   'wdArabicBahrain' = 15361,
   'wdArabicEgypt' = 3073,
   'wdArabicIraq' = 2049,
   'wdArabicJordan' = 11265,
   'wdArabicKuwait' = 13313,
   'wdArabicLebanon' = 12289,
   'wdArabicLibya' = 4097,
   'wdArabicMorocco' = 6145,
   'wdArabicOman' = 8193,
   'wdArabicQatar' = 16385,
   'wdArabic' = 1025,
   'wdArabicSyria' = 10241,
   'wdArabicTunisia' = 7169,
   'wdArabicUAE' = 14337,
   'wdArabicYemen' = 9217,
   'wdArmenian' = 1067,
   'wdAssamese' = 1101,
   'wdAzeriCyrillic' = 2092,
   'wdAzeriLatin' = 1068,
   'wdBasque' = 1069,
   'wdByelorussian' = 1059,
   'wdBengali' = 1093,
   'wdBulgarian' = 1026,
   'wdBurmese' = 1109,
   'wdCatalan' = 1027,
   'wdCherokee' = 1116,
   'wdChineseHongKongSAR' = 3076,
   'wdChineseMacaoSAR' = 5124,
   'wdSimplifiedChinese' = 2052,
   'wdChineseSingapore' = 4100,
   'wdTraditionalChinese' = 1028,
   'wdCroatian' = 1050,
   'wdCzech' = 1029,
   'wdDanish' = 1030,
   'wdDivehi' = 1125,
   'wdBelgianDutch' = 2067,
   'wdDutch' = 1043,
   'wdDzongkhaBhutan' = 2129,
   'wdEdo' = 1126,
   'wdEnglishAUS' = 3081,
   'wdEnglishBelize' = 10249,
   'wdEnglishCanadian' = 4105,
   'wdEnglishCaribbean' = 9225,
   'wdEnglishIreland' = 6153,
   'wdEnglishJamaica' = 8201,
   'wdEnglishNewZealand' = 5129,
   'wdEnglishPhilippines' = 13321,
   'wdEnglishSouthAfrica' = 7177,
   'wdEnglishTrinidadTobago' = 11273,
   'wdEnglishUK' = 2057,
   'wdEnglishUS' = 1033,
   'wdEnglishZimbabwe' = 12297,
   'wdEnglishIndonesia' = 14345,
   'wdEstonian' = 1061,
   'wdFaeroese' = 1080,
   'wdFarsi' = 1065,
   'wdFilipino' = 1124,
   'wdFinnish' = 1035,
   'wdFulfulde' = 1127,
   'wdBelgianFrench' = 2060,
   'wdFrenchCameroon' = 11276,
   'wdFrenchCanadian' = 3084,
   'wdFrenchCotedIvoire' = 12300,
   'wdFrench' = 1036,
   'wdFrenchLuxembourg' = 5132,
   'wdFrenchMali' = 13324,
   'wdFrenchMonaco' = 6156,
   'wdFrenchReunion' = 8204,
   'wdFrenchSenegal' = 10252,
   'wdFrenchMorocco' = 14348,
   'wdFrenchHaiti' = 15372,
   'wdSwissFrench' = 4108,
   'wdFrenchWestIndies' = 7180,
   'wdFrenchZaire' = 9228,
   'wdFrisianNetherlands' = 1122,
   'wdGaelicIreland' = 2108,
   'wdGaelicScotland' = 1084,
   'wdGalician' = 1110,
   'wdGeorgian' = 1079,
   'wdGermanAustria' = 3079,
   'wdGerman' = 1031,
   'wdGermanLiechtenstein' = 5127,
   'wdGermanLuxembourg' = 4103,
   'wdSwissGerman' = 2055,
   'wdGreek' = 1032,
   'wdGuarani' = 1140,
   'wdGujarati' = 1095,
   'wdHausa' = 1128,
   'wdHawaiian' = 1141,
   'wdHebrew' = 1037,
   'wdHindi' = 1081,
   'wdHungarian' = 1038,
   'wdIbibio' = 1129,
   'wdIcelandic' = 1039,
   'wdIgbo' = 1136,
   'wdIndonesian' = 1057,
   'wdInuktitut' = 1117,
   'wdItalian' = 1040,
   'wdSwissItalian' = 2064,
   'wdJapanese' = 1041,
   'wdKannada' = 1099,
   'wdKanuri' = 1137,
   'wdKashmiri' = 1120,
   'wdKazakh' = 1087,
   'wdKhmer' = 1107,
   'wdKirghiz' = 1088,
   'wdKonkani' = 1111,
   'wdKorean' = 1042,
   'wdKyrgyz' = 1088,
   'wdLao' = 1108,
   'wdLatin' = 1142,
   'wdLatvian' = 1062,
   'wdLithuanian' = 1063,
   'wdMacedonian' = 1071,
   'wdMalaysian' = 1086,
   'wdMalayBruneiDarussalam' = 2110,
   'wdMalayalam' = 1100,
   'wdMaltese' = 1082,
   'wdManipuri' = 1112,
   'wdMarathi' = 1102,
   'wdMongolian' = 1104,
   'wdNepali' = 1121,
   'wdNorwegianBokmol' = 1044,
   'wdNorwegianNynorsk' = 2068,
   'wdOriya' = 1096,
   'wdOromo' = 1138,
   'wdPashto' = 1123,
   'wdPolish' = 1045,
   'wdBrazilianPortuguese' = 1046,
   'wdPortuguese' = 2070,
   'wdPunjabi' = 1094,
   'wdRhaetoRomanic' = 1047,
   'wdRomanianMoldova' = 2072,
   'wdRomanian' = 1048,
   'wdRussianMoldova' = 2073,
   'wdRussian' = 1049,
   'wdSamiLappish' = 1083,
   'wdSanskrit' = 1103,
   'wdSerbianCyrillic' = 3098,
   'wdSerbianLatin' = 2074,
   'wdSinhalese' = 1115,
   'wdSindhi' = 1113,
   'wdSindhiPakistan' = 2137,
   'wdSlovak' = 1051,
   'wdSlovenian' = 1060,
   'wdSomali' = 1143,
   'wdSorbian' = 1070,
   'wdSpanishArgentina' = 11274,
   'wdSpanishBolivia' = 16394,
   'wdSpanishChile' = 13322,
   'wdSpanishColombia' = 9226,
   'wdSpanishCostaRica' = 5130,
   'wdSpanishDominicanRepublic' = 7178,
   'wdSpanishEcuador' = 12298,
   'wdSpanishElSalvador' = 17418,
   'wdSpanishGuatemala' = 4106,
   'wdSpanishHonduras' = 18442,
   'wdMexicanSpanish' = 2058,
   'wdSpanishNicaragua' = 19466,
   'wdSpanishPanama' = 6154,
   'wdSpanishParaguay' = 15370,
   'wdSpanishPeru' = 10250,
   'wdSpanishPuertoRico' = 20490,
   'wdSpanishModernSort' = 3082,
   'wdSpanish' = 1034,
   'wdSpanishUruguay' = 14346,
   'wdSpanishVenezuela' = 8202,
   'wdSesotho' = 1072,
   'wdSutu' = 1072,
   'wdSwahili' = 1089,
   'wdSwedishFinland' = 2077,
   'wdSwedish' = 1053,
   'wdSyriac' = 1114,
   'wdTajik' = 1064,
   'wdTamazight' = 1119,
   'wdTamazightLatin' = 2143,
   'wdTamil' = 1097,
   'wdTatar' = 1092,
   'wdTelugu' = 1098,
   'wdThai' = 1054,
   'wdTibetan' = 1105,
   'wdTigrignaEthiopic' = 1139,
   'wdTigrignaEritrea' = 2163,
   'wdTsonga' = 1073,
   'wdTswana' = 1074,
   'wdTurkish' = 1055,
   'wdTurkmen' = 1090,
   'wdUkrainian' = 1058,
   'wdUrdu' = 1056,
   'wdUzbekCyrillic' = 2115,
   'wdUzbekLatin' = 1091,
   'wdVenda' = 1075,
   'wdVietnamese' = 1066,
   'wdWelsh' = 1106,
   'wdXhosa' = 1076,
   'wdYi' = 1144,
   'wdYiddish' = 1085,
   'wdYoruba' = 1130,
   'wdZulu' = 1077 
 )
storage.mode( WdLanguageIDEnum ) = 'integer'
'WdFieldTypeEnum' = c(
   'wdFieldEmpty' = -1,
   'wdFieldRef' = 3,
   'wdFieldIndexEntry' = 4,
   'wdFieldFootnoteRef' = 5,
   'wdFieldSet' = 6,
   'wdFieldIf' = 7,
   'wdFieldIndex' = 8,
   'wdFieldTOCEntry' = 9,
   'wdFieldStyleRef' = 10,
   'wdFieldRefDoc' = 11,
   'wdFieldSequence' = 12,
   'wdFieldTOC' = 13,
   'wdFieldInfo' = 14,
   'wdFieldTitle' = 15,
   'wdFieldSubject' = 16,
   'wdFieldAuthor' = 17,
   'wdFieldKeyWord' = 18,
   'wdFieldComments' = 19,
   'wdFieldLastSavedBy' = 20,
   'wdFieldCreateDate' = 21,
   'wdFieldSaveDate' = 22,
   'wdFieldPrintDate' = 23,
   'wdFieldRevisionNum' = 24,
   'wdFieldEditTime' = 25,
   'wdFieldNumPages' = 26,
   'wdFieldNumWords' = 27,
   'wdFieldNumChars' = 28,
   'wdFieldFileName' = 29,
   'wdFieldTemplate' = 30,
   'wdFieldDate' = 31,
   'wdFieldTime' = 32,
   'wdFieldPage' = 33,
   'wdFieldExpression' = 34,
   'wdFieldQuote' = 35,
   'wdFieldInclude' = 36,
   'wdFieldPageRef' = 37,
   'wdFieldAsk' = 38,
   'wdFieldFillIn' = 39,
   'wdFieldData' = 40,
   'wdFieldNext' = 41,
   'wdFieldNextIf' = 42,
   'wdFieldSkipIf' = 43,
   'wdFieldMergeRec' = 44,
   'wdFieldDDE' = 45,
   'wdFieldDDEAuto' = 46,
   'wdFieldGlossary' = 47,
   'wdFieldPrint' = 48,
   'wdFieldFormula' = 49,
   'wdFieldGoToButton' = 50,
   'wdFieldMacroButton' = 51,
   'wdFieldAutoNumOutline' = 52,
   'wdFieldAutoNumLegal' = 53,
   'wdFieldAutoNum' = 54,
   'wdFieldImport' = 55,
   'wdFieldLink' = 56,
   'wdFieldSymbol' = 57,
   'wdFieldEmbed' = 58,
   'wdFieldMergeField' = 59,
   'wdFieldUserName' = 60,
   'wdFieldUserInitials' = 61,
   'wdFieldUserAddress' = 62,
   'wdFieldBarCode' = 63,
   'wdFieldDocVariable' = 64,
   'wdFieldSection' = 65,
   'wdFieldSectionPages' = 66,
   'wdFieldIncludePicture' = 67,
   'wdFieldIncludeText' = 68,
   'wdFieldFileSize' = 69,
   'wdFieldFormTextInput' = 70,
   'wdFieldFormCheckBox' = 71,
   'wdFieldNoteRef' = 72,
   'wdFieldTOA' = 73,
   'wdFieldTOAEntry' = 74,
   'wdFieldMergeSeq' = 75,
   'wdFieldPrivate' = 77,
   'wdFieldDatabase' = 78,
   'wdFieldAutoText' = 79,
   'wdFieldCompare' = 80,
   'wdFieldAddin' = 81,
   'wdFieldSubscriber' = 82,
   'wdFieldFormDropDown' = 83,
   'wdFieldAdvance' = 84,
   'wdFieldDocProperty' = 85,
   'wdFieldOCX' = 87,
   'wdFieldHyperlink' = 88,
   'wdFieldAutoTextList' = 89,
   'wdFieldListNum' = 90,
   'wdFieldHTMLActiveX' = 91,
   'wdFieldBidiOutline' = 92,
   'wdFieldAddressBlock' = 93,
   'wdFieldGreetingLine' = 94,
   'wdFieldShape' = 95 
 )
storage.mode( WdFieldTypeEnum ) = 'integer'
'WdBuiltinStyleEnum' = c(
   'wdStyleNormal' = -1,
   'wdStyleEnvelopeAddress' = -37,
   'wdStyleEnvelopeReturn' = -38,
   'wdStyleBodyText' = -67,
   'wdStyleHeading1' = -2,
   'wdStyleHeading2' = -3,
   'wdStyleHeading3' = -4,
   'wdStyleHeading4' = -5,
   'wdStyleHeading5' = -6,
   'wdStyleHeading6' = -7,
   'wdStyleHeading7' = -8,
   'wdStyleHeading8' = -9,
   'wdStyleHeading9' = -10,
   'wdStyleIndex1' = -11,
   'wdStyleIndex2' = -12,
   'wdStyleIndex3' = -13,
   'wdStyleIndex4' = -14,
   'wdStyleIndex5' = -15,
   'wdStyleIndex6' = -16,
   'wdStyleIndex7' = -17,
   'wdStyleIndex8' = -18,
   'wdStyleIndex9' = -19,
   'wdStyleTOC1' = -20,
   'wdStyleTOC2' = -21,
   'wdStyleTOC3' = -22,
   'wdStyleTOC4' = -23,
   'wdStyleTOC5' = -24,
   'wdStyleTOC6' = -25,
   'wdStyleTOC7' = -26,
   'wdStyleTOC8' = -27,
   'wdStyleTOC9' = -28,
   'wdStyleNormalIndent' = -29,
   'wdStyleFootnoteText' = -30,
   'wdStyleCommentText' = -31,
   'wdStyleHeader' = -32,
   'wdStyleFooter' = -33,
   'wdStyleIndexHeading' = -34,
   'wdStyleCaption' = -35,
   'wdStyleTableOfFigures' = -36,
   'wdStyleFootnoteReference' = -39,
   'wdStyleCommentReference' = -40,
   'wdStyleLineNumber' = -41,
   'wdStylePageNumber' = -42,
   'wdStyleEndnoteReference' = -43,
   'wdStyleEndnoteText' = -44,
   'wdStyleTableOfAuthorities' = -45,
   'wdStyleMacroText' = -46,
   'wdStyleTOAHeading' = -47,
   'wdStyleList' = -48,
   'wdStyleListBullet' = -49,
   'wdStyleListNumber' = -50,
   'wdStyleList2' = -51,
   'wdStyleList3' = -52,
   'wdStyleList4' = -53,
   'wdStyleList5' = -54,
   'wdStyleListBullet2' = -55,
   'wdStyleListBullet3' = -56,
   'wdStyleListBullet4' = -57,
   'wdStyleListBullet5' = -58,
   'wdStyleListNumber2' = -59,
   'wdStyleListNumber3' = -60,
   'wdStyleListNumber4' = -61,
   'wdStyleListNumber5' = -62,
   'wdStyleTitle' = -63,
   'wdStyleClosing' = -64,
   'wdStyleSignature' = -65,
   'wdStyleDefaultParagraphFont' = -66,
   'wdStyleBodyTextIndent' = -68,
   'wdStyleListContinue' = -69,
   'wdStyleListContinue2' = -70,
   'wdStyleListContinue3' = -71,
   'wdStyleListContinue4' = -72,
   'wdStyleListContinue5' = -73,
   'wdStyleMessageHeader' = -74,
   'wdStyleSubtitle' = -75,
   'wdStyleSalutation' = -76,
   'wdStyleDate' = -77,
   'wdStyleBodyTextFirstIndent' = -78,
   'wdStyleBodyTextFirstIndent2' = -79,
   'wdStyleNoteHeading' = -80,
   'wdStyleBodyText2' = -81,
   'wdStyleBodyText3' = -82,
   'wdStyleBodyTextIndent2' = -83,
   'wdStyleBodyTextIndent3' = -84,
   'wdStyleBlockQuotation' = -85,
   'wdStyleHyperlink' = -86,
   'wdStyleHyperlinkFollowed' = -87,
   'wdStyleStrong' = -88,
   'wdStyleEmphasis' = -89,
   'wdStyleNavPane' = -90,
   'wdStylePlainText' = -91,
   'wdStyleHtmlNormal' = -95,
   'wdStyleHtmlAcronym' = -96,
   'wdStyleHtmlAddress' = -97,
   'wdStyleHtmlCite' = -98,
   'wdStyleHtmlCode' = -99,
   'wdStyleHtmlDfn' = -100,
   'wdStyleHtmlKbd' = -101,
   'wdStyleHtmlPre' = -102,
   'wdStyleHtmlSamp' = -103,
   'wdStyleHtmlTt' = -104,
   'wdStyleHtmlVar' = -105,
   'wdStyleNormalTable' = -106 
 )
storage.mode( WdBuiltinStyleEnum ) = 'integer'
'WdWordDialogTabEnum' = c(
   'wdDialogToolsOptionsTabView' = 204,
   'wdDialogToolsOptionsTabGeneral' = 203,
   'wdDialogToolsOptionsTabEdit' = 224,
   'wdDialogToolsOptionsTabPrint' = 208,
   'wdDialogToolsOptionsTabSave' = 209,
   'wdDialogToolsOptionsTabProofread' = 211,
   'wdDialogToolsOptionsTabTrackChanges' = 386,
   'wdDialogToolsOptionsTabUserInfo' = 213,
   'wdDialogToolsOptionsTabCompatibility' = 525,
   'wdDialogToolsOptionsTabTypography' = 739,
   'wdDialogToolsOptionsTabFileLocations' = 225,
   'wdDialogToolsOptionsTabFuzzy' = 790,
   'wdDialogToolsOptionsTabHangulHanjaConversion' = 786,
   'wdDialogToolsOptionsTabBidi' = 1029,
   'wdDialogToolsOptionsTabSecurity' = 1361,
   'wdDialogFilePageSetupTabMargins' = 150000,
   'wdDialogFilePageSetupTabPaper' = 150001,
   'wdDialogFilePageSetupTabLayout' = 150003,
   'wdDialogFilePageSetupTabCharsLines' = 150004,
   'wdDialogInsertSymbolTabSymbols' = 200000,
   'wdDialogInsertSymbolTabSpecialCharacters' = 200001,
   'wdDialogNoteOptionsTabAllFootnotes' = 300000,
   'wdDialogNoteOptionsTabAllEndnotes' = 300001,
   'wdDialogInsertIndexAndTablesTabIndex' = 400000,
   'wdDialogInsertIndexAndTablesTabTableOfContents' = 400001,
   'wdDialogInsertIndexAndTablesTabTableOfFigures' = 400002,
   'wdDialogInsertIndexAndTablesTabTableOfAuthorities' = 400003,
   'wdDialogOrganizerTabStyles' = 500000,
   'wdDialogOrganizerTabAutoText' = 500001,
   'wdDialogOrganizerTabCommandBars' = 500002,
   'wdDialogOrganizerTabMacros' = 500003,
   'wdDialogFormatFontTabFont' = 600000,
   'wdDialogFormatFontTabCharacterSpacing' = 600001,
   'wdDialogFormatFontTabAnimation' = 600002,
   'wdDialogFormatBordersAndShadingTabBorders' = 700000,
   'wdDialogFormatBordersAndShadingTabPageBorder' = 700001,
   'wdDialogFormatBordersAndShadingTabShading' = 700002,
   'wdDialogToolsEnvelopesAndLabelsTabEnvelopes' = 800000,
   'wdDialogToolsEnvelopesAndLabelsTabLabels' = 800001,
   'wdDialogFormatParagraphTabIndentsAndSpacing' = 1000000,
   'wdDialogFormatParagraphTabTextFlow' = 1000001,
   'wdDialogFormatParagraphTabTeisai' = 1000002,
   'wdDialogFormatDrawingObjectTabColorsAndLines' = 1200000,
   'wdDialogFormatDrawingObjectTabSize' = 1200001,
   'wdDialogFormatDrawingObjectTabPosition' = 1200002,
   'wdDialogFormatDrawingObjectTabWrapping' = 1200003,
   'wdDialogFormatDrawingObjectTabPicture' = 1200004,
   'wdDialogFormatDrawingObjectTabTextbox' = 1200005,
   'wdDialogFormatDrawingObjectTabWeb' = 1200006,
   'wdDialogFormatDrawingObjectTabHR' = 1200007,
   'wdDialogToolsAutoCorrectExceptionsTabFirstLetter' = 1400000,
   'wdDialogToolsAutoCorrectExceptionsTabInitialCaps' = 1400001,
   'wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet' = 1400002,
   'wdDialogToolsAutoCorrectExceptionsTabIac' = 1400003,
   'wdDialogFormatBulletsAndNumberingTabBulleted' = 1500000,
   'wdDialogFormatBulletsAndNumberingTabNumbered' = 1500001,
   'wdDialogFormatBulletsAndNumberingTabOutlineNumbered' = 1500002,
   'wdDialogLetterWizardTabLetterFormat' = 1600000,
   'wdDialogLetterWizardTabRecipientInfo' = 1600001,
   'wdDialogLetterWizardTabOtherElements' = 1600002,
   'wdDialogLetterWizardTabSenderInfo' = 1600003,
   'wdDialogToolsAutoManagerTabAutoCorrect' = 1700000,
   'wdDialogToolsAutoManagerTabAutoFormatAsYouType' = 1700001,
   'wdDialogToolsAutoManagerTabAutoText' = 1700002,
   'wdDialogToolsAutoManagerTabAutoFormat' = 1700003,
   'wdDialogToolsAutoManagerTabSmartTags' = 1700004,
   'wdDialogTablePropertiesTabTable' = 1800000,
   'wdDialogTablePropertiesTabRow' = 1800001,
   'wdDialogTablePropertiesTabColumn' = 1800002,
   'wdDialogTablePropertiesTabCell' = 1800003,
   'wdDialogEmailOptionsTabSignature' = 1900000,
   'wdDialogEmailOptionsTabStationary' = 1900001,
   'wdDialogEmailOptionsTabQuoting' = 1900002,
   'wdDialogWebOptionsBrowsers' = 2000000,
   'wdDialogWebOptionsGeneral' = 2000000,
   'wdDialogWebOptionsFiles' = 2000001,
   'wdDialogWebOptionsPictures' = 2000002,
   'wdDialogWebOptionsEncoding' = 2000003,
   'wdDialogWebOptionsFonts' = 2000004,
   'wdDialogToolsOptionsTabAcetate' = 1266,
   'wdDialogTemplates' = 2100000,
   'wdDialogTemplatesXMLSchema' = 2100001,
   'wdDialogTemplatesXMLExpansionPacks' = 2100002,
   'wdDialogTemplatesLinkedCSS' = 2100003 
 )
storage.mode( WdWordDialogTabEnum ) = 'integer'
'WdWordDialogTabHIDEnum' = c(
   'wdDialogFilePageSetupTabPaperSize' = 150001,
   'wdDialogFilePageSetupTabPaperSource' = 150002 
 )
storage.mode( WdWordDialogTabHIDEnum ) = 'integer'
'WdWordDialogEnum' = c(
   'wdDialogHelpAbout' = 9,
   'wdDialogHelpWordPerfectHelp' = 10,
   'wdDialogDocumentStatistics' = 78,
   'wdDialogFileNew' = 79,
   'wdDialogFileOpen' = 80,
   'wdDialogMailMergeOpenDataSource' = 81,
   'wdDialogMailMergeOpenHeaderSource' = 82,
   'wdDialogFileSaveAs' = 84,
   'wdDialogFileSummaryInfo' = 86,
   'wdDialogToolsTemplates' = 87,
   'wdDialogFilePrint' = 88,
   'wdDialogFilePrintSetup' = 97,
   'wdDialogFileFind' = 99,
   'wdDialogFormatAddrFonts' = 103,
   'wdDialogEditPasteSpecial' = 111,
   'wdDialogEditFind' = 112,
   'wdDialogEditReplace' = 117,
   'wdDialogEditStyle' = 120,
   'wdDialogEditLinks' = 124,
   'wdDialogEditObject' = 125,
   'wdDialogTableToText' = 128,
   'wdDialogTextToTable' = 127,
   'wdDialogTableInsertTable' = 129,
   'wdDialogTableInsertCells' = 130,
   'wdDialogTableInsertRow' = 131,
   'wdDialogTableDeleteCells' = 133,
   'wdDialogTableSplitCells' = 137,
   'wdDialogTableRowHeight' = 142,
   'wdDialogTableColumnWidth' = 143,
   'wdDialogToolsCustomize' = 152,
   'wdDialogInsertBreak' = 159,
   'wdDialogInsertSymbol' = 162,
   'wdDialogInsertPicture' = 163,
   'wdDialogInsertFile' = 164,
   'wdDialogInsertDateTime' = 165,
   'wdDialogInsertField' = 166,
   'wdDialogInsertMergeField' = 167,
   'wdDialogInsertBookmark' = 168,
   'wdDialogMarkIndexEntry' = 169,
   'wdDialogInsertIndex' = 170,
   'wdDialogInsertTableOfContents' = 171,
   'wdDialogInsertObject' = 172,
   'wdDialogToolsCreateEnvelope' = 173,
   'wdDialogFormatFont' = 174,
   'wdDialogFormatParagraph' = 175,
   'wdDialogFormatSectionLayout' = 176,
   'wdDialogFormatColumns' = 177,
   'wdDialogFileDocumentLayout' = 178,
   'wdDialogFilePageSetup' = 178,
   'wdDialogFormatTabs' = 179,
   'wdDialogFormatStyle' = 180,
   'wdDialogFormatDefineStyleFont' = 181,
   'wdDialogFormatDefineStylePara' = 182,
   'wdDialogFormatDefineStyleTabs' = 183,
   'wdDialogFormatDefineStyleFrame' = 184,
   'wdDialogFormatDefineStyleBorders' = 185,
   'wdDialogFormatDefineStyleLang' = 186,
   'wdDialogFormatPicture' = 187,
   'wdDialogToolsLanguage' = 188,
   'wdDialogFormatBordersAndShading' = 189,
   'wdDialogFormatFrame' = 190,
   'wdDialogToolsThesaurus' = 194,
   'wdDialogToolsHyphenation' = 195,
   'wdDialogToolsBulletsNumbers' = 196,
   'wdDialogToolsHighlightChanges' = 197,
   'wdDialogToolsRevisions' = 197,
   'wdDialogToolsCompareDocuments' = 198,
   'wdDialogTableSort' = 199,
   'wdDialogToolsOptionsGeneral' = 203,
   'wdDialogToolsOptionsView' = 204,
   'wdDialogToolsAdvancedSettings' = 206,
   'wdDialogToolsOptionsPrint' = 208,
   'wdDialogToolsOptionsSave' = 209,
   'wdDialogToolsOptionsSpellingAndGrammar' = 211,
   'wdDialogToolsOptionsUserInfo' = 213,
   'wdDialogToolsMacroRecord' = 214,
   'wdDialogToolsMacro' = 215,
   'wdDialogWindowActivate' = 220,
   'wdDialogFormatRetAddrFonts' = 221,
   'wdDialogOrganizer' = 222,
   'wdDialogToolsOptionsEdit' = 224,
   'wdDialogToolsOptionsFileLocations' = 225,
   'wdDialogToolsWordCount' = 228,
   'wdDialogControlRun' = 235,
   'wdDialogInsertPageNumbers' = 294,
   'wdDialogFormatPageNumber' = 298,
   'wdDialogCopyFile' = 300,
   'wdDialogFormatChangeCase' = 322,
   'wdDialogUpdateTOC' = 331,
   'wdDialogInsertDatabase' = 341,
   'wdDialogTableFormula' = 348,
   'wdDialogFormFieldOptions' = 353,
   'wdDialogInsertCaption' = 357,
   'wdDialogInsertCaptionNumbering' = 358,
   'wdDialogInsertAutoCaption' = 359,
   'wdDialogFormFieldHelp' = 361,
   'wdDialogInsertCrossReference' = 367,
   'wdDialogInsertFootnote' = 370,
   'wdDialogNoteOptions' = 373,
   'wdDialogToolsAutoCorrect' = 378,
   'wdDialogToolsOptionsTrackChanges' = 386,
   'wdDialogConvertObject' = 392,
   'wdDialogInsertAddCaption' = 402,
   'wdDialogConnect' = 420,
   'wdDialogToolsCustomizeKeyboard' = 432,
   'wdDialogToolsCustomizeMenus' = 433,
   'wdDialogToolsMergeDocuments' = 435,
   'wdDialogMarkTableOfContentsEntry' = 442,
   'wdDialogFileMacPageSetupGX' = 444,
   'wdDialogFilePrintOneCopy' = 445,
   'wdDialogEditFrame' = 458,
   'wdDialogMarkCitation' = 463,
   'wdDialogTableOfContentsOptions' = 470,
   'wdDialogInsertTableOfAuthorities' = 471,
   'wdDialogInsertTableOfFigures' = 472,
   'wdDialogInsertIndexAndTables' = 473,
   'wdDialogInsertFormField' = 483,
   'wdDialogFormatDropCap' = 488,
   'wdDialogToolsCreateLabels' = 489,
   'wdDialogToolsProtectDocument' = 503,
   'wdDialogFormatStyleGallery' = 505,
   'wdDialogToolsAcceptRejectChanges' = 506,
   'wdDialogHelpWordPerfectHelpOptions' = 511,
   'wdDialogToolsUnprotectDocument' = 521,
   'wdDialogToolsOptionsCompatibility' = 525,
   'wdDialogTableOfCaptionsOptions' = 551,
   'wdDialogTableAutoFormat' = 563,
   'wdDialogMailMergeFindRecord' = 569,
   'wdDialogReviewAfmtRevisions' = 570,
   'wdDialogViewZoom' = 577,
   'wdDialogToolsProtectSection' = 578,
   'wdDialogFontSubstitution' = 581,
   'wdDialogInsertSubdocument' = 583,
   'wdDialogNewToolbar' = 586,
   'wdDialogToolsEnvelopesAndLabels' = 607,
   'wdDialogFormatCallout' = 610,
   'wdDialogTableFormatCell' = 612,
   'wdDialogToolsCustomizeMenuBar' = 615,
   'wdDialogFileRoutingSlip' = 624,
   'wdDialogEditTOACategory' = 625,
   'wdDialogToolsManageFields' = 631,
   'wdDialogDrawSnapToGrid' = 633,
   'wdDialogDrawAlign' = 634,
   'wdDialogMailMergeCreateDataSource' = 642,
   'wdDialogMailMergeCreateHeaderSource' = 643,
   'wdDialogMailMerge' = 676,
   'wdDialogMailMergeCheck' = 677,
   'wdDialogMailMergeHelper' = 680,
   'wdDialogMailMergeQueryOptions' = 681,
   'wdDialogFileMacPageSetup' = 685,
   'wdDialogListCommands' = 723,
   'wdDialogEditCreatePublisher' = 732,
   'wdDialogEditSubscribeTo' = 733,
   'wdDialogEditPublishOptions' = 735,
   'wdDialogEditSubscribeOptions' = 736,
   'wdDialogFileMacCustomPageSetupGX' = 737,
   'wdDialogToolsOptionsTypography' = 739,
   'wdDialogToolsAutoCorrectExceptions' = 762,
   'wdDialogToolsOptionsAutoFormatAsYouType' = 778,
   'wdDialogMailMergeUseAddressBook' = 779,
   'wdDialogToolsHangulHanjaConversion' = 784,
   'wdDialogToolsOptionsFuzzy' = 790,
   'wdDialogEditGoToOld' = 811,
   'wdDialogInsertNumber' = 812,
   'wdDialogLetterWizard' = 821,
   'wdDialogFormatBulletsAndNumbering' = 824,
   'wdDialogToolsSpellingAndGrammar' = 828,
   'wdDialogToolsCreateDirectory' = 833,
   'wdDialogTableWrapping' = 854,
   'wdDialogFormatTheme' = 855,
   'wdDialogTableProperties' = 861,
   'wdDialogEmailOptions' = 863,
   'wdDialogCreateAutoText' = 872,
   'wdDialogToolsAutoSummarize' = 874,
   'wdDialogToolsGrammarSettings' = 885,
   'wdDialogEditGoTo' = 896,
   'wdDialogWebOptions' = 898,
   'wdDialogInsertHyperlink' = 925,
   'wdDialogToolsAutoManager' = 915,
   'wdDialogFileVersions' = 945,
   'wdDialogToolsOptionsAutoFormat' = 959,
   'wdDialogFormatDrawingObject' = 960,
   'wdDialogToolsOptions' = 974,
   'wdDialogFitText' = 983,
   'wdDialogEditAutoText' = 985,
   'wdDialogPhoneticGuide' = 986,
   'wdDialogToolsDictionary' = 989,
   'wdDialogFileSaveVersion' = 1007,
   'wdDialogToolsOptionsBidi' = 1029,
   'wdDialogFrameSetProperties' = 1074,
   'wdDialogTableTableOptions' = 1080,
   'wdDialogTableCellOptions' = 1081,
   'wdDialogIMESetDefault' = 1094,
   'wdDialogTCSCTranslator' = 1156,
   'wdDialogHorizontalInVertical' = 1160,
   'wdDialogTwoLinesInOne' = 1161,
   'wdDialogFormatEncloseCharacters' = 1162,
   'wdDialogConsistencyChecker' = 1121,
   'wdDialogToolsOptionsSmartTag' = 1395,
   'wdDialogFormatStylesCustom' = 1248,
   'wdDialogCSSLinks' = 1261,
   'wdDialogInsertWebComponent' = 1324,
   'wdDialogToolsOptionsEditCopyPaste' = 1356,
   'wdDialogToolsOptionsSecurity' = 1361,
   'wdDialogSearch' = 1363,
   'wdDialogShowRepairs' = 1381,
   'wdDialogMailMergeInsertAsk' = 4047,
   'wdDialogMailMergeInsertFillIn' = 4048,
   'wdDialogMailMergeInsertIf' = 4049,
   'wdDialogMailMergeInsertNextIf' = 4053,
   'wdDialogMailMergeInsertSet' = 4054,
   'wdDialogMailMergeInsertSkipIf' = 4055,
   'wdDialogMailMergeFieldMapping' = 1304,
   'wdDialogMailMergeInsertAddressBlock' = 1305,
   'wdDialogMailMergeInsertGreetingLine' = 1306,
   'wdDialogMailMergeInsertFields' = 1307,
   'wdDialogMailMergeRecipients' = 1308,
   'wdDialogMailMergeFindRecipient' = 1326,
   'wdDialogMailMergeSetDocumentType' = 1339,
   'wdDialogXMLElementAttributes' = 1460,
   'wdDialogSchemaLibrary' = 1417,
   'wdDialogPermission' = 1469,
   'wdDialogMyPermission' = 1437,
   'wdDialogXMLOptions' = 1425,
   'wdDialogFormattingRestrictions' = 1427 
 )
storage.mode( WdWordDialogEnum ) = 'integer'
'WdWordDialogHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdWordDialogHIDEnum ) = 'integer'
'WdFieldKindEnum' = c(
   'wdFieldKindNone' = 0,
   'wdFieldKindHot' = 1,
   'wdFieldKindWarm' = 2,
   'wdFieldKindCold' = 3 
 )
storage.mode( WdFieldKindEnum ) = 'integer'
'WdTextFormFieldTypeEnum' = c(
   'wdRegularText' = 0,
   'wdNumberText' = 1,
   'wdDateText' = 2,
   'wdCurrentDateText' = 3,
   'wdCurrentTimeText' = 4,
   'wdCalculationText' = 5 
 )
storage.mode( WdTextFormFieldTypeEnum ) = 'integer'
'WdChevronConvertRuleEnum' = c(
   'wdNeverConvert' = 0,
   'wdAlwaysConvert' = 1,
   'wdAskToNotConvert' = 2,
   'wdAskToConvert' = 3 
 )
storage.mode( WdChevronConvertRuleEnum ) = 'integer'
'WdMailMergeMainDocTypeEnum' = c(
   'wdNotAMergeDocument' = -1,
   'wdFormLetters' = 0,
   'wdMailingLabels' = 1,
   'wdEnvelopes' = 2,
   'wdCatalog' = 3,
   'wdEMail' = 4,
   'wdFax' = 5,
   'wdDirectory' = 3 
 )
storage.mode( WdMailMergeMainDocTypeEnum ) = 'integer'
'WdMailMergeStateEnum' = c(
   'wdNormalDocument' = 0,
   'wdMainDocumentOnly' = 1,
   'wdMainAndDataSource' = 2,
   'wdMainAndHeader' = 3,
   'wdMainAndSourceAndHeader' = 4,
   'wdDataSource' = 5 
 )
storage.mode( WdMailMergeStateEnum ) = 'integer'
'WdMailMergeDestinationEnum' = c(
   'wdSendToNewDocument' = 0,
   'wdSendToPrinter' = 1,
   'wdSendToEmail' = 2,
   'wdSendToFax' = 3 
 )
storage.mode( WdMailMergeDestinationEnum ) = 'integer'
'WdMailMergeActiveRecordEnum' = c(
   'wdNoActiveRecord' = -1,
   'wdNextRecord' = -2,
   'wdPreviousRecord' = -3,
   'wdFirstRecord' = -4,
   'wdLastRecord' = -5,
   'wdFirstDataSourceRecord' = -6,
   'wdLastDataSourceRecord' = -7,
   'wdNextDataSourceRecord' = -8,
   'wdPreviousDataSourceRecord' = -9 
 )
storage.mode( WdMailMergeActiveRecordEnum ) = 'integer'
'WdMailMergeDefaultRecordEnum' = c(
   'wdDefaultFirstRecord' = 1,
   'wdDefaultLastRecord' = -16 
 )
storage.mode( WdMailMergeDefaultRecordEnum ) = 'integer'
'WdMailMergeDataSourceEnum' = c(
   'wdNoMergeInfo' = -1,
   'wdMergeInfoFromWord' = 0,
   'wdMergeInfoFromAccessDDE' = 1,
   'wdMergeInfoFromExcelDDE' = 2,
   'wdMergeInfoFromMSQueryDDE' = 3,
   'wdMergeInfoFromODBC' = 4,
   'wdMergeInfoFromODSO' = 5 
 )
storage.mode( WdMailMergeDataSourceEnum ) = 'integer'
'WdMailMergeComparisonEnum' = c(
   'wdMergeIfEqual' = 0,
   'wdMergeIfNotEqual' = 1,
   'wdMergeIfLessThan' = 2,
   'wdMergeIfGreaterThan' = 3,
   'wdMergeIfLessThanOrEqual' = 4,
   'wdMergeIfGreaterThanOrEqual' = 5,
   'wdMergeIfIsBlank' = 6,
   'wdMergeIfIsNotBlank' = 7 
 )
storage.mode( WdMailMergeComparisonEnum ) = 'integer'
'WdBookmarkSortByEnum' = c(
   'wdSortByName' = 0,
   'wdSortByLocation' = 1 
 )
storage.mode( WdBookmarkSortByEnum ) = 'integer'
'WdWindowStateEnum' = c(
   'wdWindowStateNormal' = 0,
   'wdWindowStateMaximize' = 1,
   'wdWindowStateMinimize' = 2 
 )
storage.mode( WdWindowStateEnum ) = 'integer'
'WdPictureLinkTypeEnum' = c(
   'wdLinkNone' = 0,
   'wdLinkDataInDoc' = 1,
   'wdLinkDataOnDisk' = 2 
 )
storage.mode( WdPictureLinkTypeEnum ) = 'integer'
'WdLinkTypeEnum' = c(
   'wdLinkTypeOLE' = 0,
   'wdLinkTypePicture' = 1,
   'wdLinkTypeText' = 2,
   'wdLinkTypeReference' = 3,
   'wdLinkTypeInclude' = 4,
   'wdLinkTypeImport' = 5,
   'wdLinkTypeDDE' = 6,
   'wdLinkTypeDDEAuto' = 7 
 )
storage.mode( WdLinkTypeEnum ) = 'integer'
'WdWindowTypeEnum' = c(
   'wdWindowDocument' = 0,
   'wdWindowTemplate' = 1 
 )
storage.mode( WdWindowTypeEnum ) = 'integer'
'WdViewTypeEnum' = c(
   'wdNormalView' = 1,
   'wdOutlineView' = 2,
   'wdPrintView' = 3,
   'wdPrintPreview' = 4,
   'wdMasterView' = 5,
   'wdWebView' = 6,
   'wdReadingView' = 7 
 )
storage.mode( WdViewTypeEnum ) = 'integer'
'WdSeekViewEnum' = c(
   'wdSeekMainDocument' = 0,
   'wdSeekPrimaryHeader' = 1,
   'wdSeekFirstPageHeader' = 2,
   'wdSeekEvenPagesHeader' = 3,
   'wdSeekPrimaryFooter' = 4,
   'wdSeekFirstPageFooter' = 5,
   'wdSeekEvenPagesFooter' = 6,
   'wdSeekFootnotes' = 7,
   'wdSeekEndnotes' = 8,
   'wdSeekCurrentPageHeader' = 9,
   'wdSeekCurrentPageFooter' = 10 
 )
storage.mode( WdSeekViewEnum ) = 'integer'
'WdSpecialPaneEnum' = c(
   'wdPaneNone' = 0,
   'wdPanePrimaryHeader' = 1,
   'wdPaneFirstPageHeader' = 2,
   'wdPaneEvenPagesHeader' = 3,
   'wdPanePrimaryFooter' = 4,
   'wdPaneFirstPageFooter' = 5,
   'wdPaneEvenPagesFooter' = 6,
   'wdPaneFootnotes' = 7,
   'wdPaneEndnotes' = 8,
   'wdPaneFootnoteContinuationNotice' = 9,
   'wdPaneFootnoteContinuationSeparator' = 10,
   'wdPaneFootnoteSeparator' = 11,
   'wdPaneEndnoteContinuationNotice' = 12,
   'wdPaneEndnoteContinuationSeparator' = 13,
   'wdPaneEndnoteSeparator' = 14,
   'wdPaneComments' = 15,
   'wdPaneCurrentPageHeader' = 16,
   'wdPaneCurrentPageFooter' = 17,
   'wdPaneRevisions' = 18 
 )
storage.mode( WdSpecialPaneEnum ) = 'integer'
'WdPageFitEnum' = c(
   'wdPageFitNone' = 0,
   'wdPageFitFullPage' = 1,
   'wdPageFitBestFit' = 2,
   'wdPageFitTextFit' = 3 
 )
storage.mode( WdPageFitEnum ) = 'integer'
'WdBrowseTargetEnum' = c(
   'wdBrowsePage' = 1,
   'wdBrowseSection' = 2,
   'wdBrowseComment' = 3,
   'wdBrowseFootnote' = 4,
   'wdBrowseEndnote' = 5,
   'wdBrowseField' = 6,
   'wdBrowseTable' = 7,
   'wdBrowseGraphic' = 8,
   'wdBrowseHeading' = 9,
   'wdBrowseEdit' = 10,
   'wdBrowseFind' = 11,
   'wdBrowseGoTo' = 12 
 )
storage.mode( WdBrowseTargetEnum ) = 'integer'
'WdPaperTrayEnum' = c(
   'wdPrinterDefaultBin' = 0,
   'wdPrinterUpperBin' = 1,
   'wdPrinterOnlyBin' = 1,
   'wdPrinterLowerBin' = 2,
   'wdPrinterMiddleBin' = 3,
   'wdPrinterManualFeed' = 4,
   'wdPrinterEnvelopeFeed' = 5,
   'wdPrinterManualEnvelopeFeed' = 6,
   'wdPrinterAutomaticSheetFeed' = 7,
   'wdPrinterTractorFeed' = 8,
   'wdPrinterSmallFormatBin' = 9,
   'wdPrinterLargeFormatBin' = 10,
   'wdPrinterLargeCapacityBin' = 11,
   'wdPrinterPaperCassette' = 14,
   'wdPrinterFormSource' = 15 
 )
storage.mode( WdPaperTrayEnum ) = 'integer'
'WdOrientationEnum' = c(
   'wdOrientPortrait' = 0,
   'wdOrientLandscape' = 1 
 )
storage.mode( WdOrientationEnum ) = 'integer'
'WdSelectionTypeEnum' = c(
   'wdNoSelection' = 0,
   'wdSelectionIP' = 1,
   'wdSelectionNormal' = 2,
   'wdSelectionFrame' = 3,
   'wdSelectionColumn' = 4,
   'wdSelectionRow' = 5,
   'wdSelectionBlock' = 6,
   'wdSelectionInlineShape' = 7,
   'wdSelectionShape' = 8 
 )
storage.mode( WdSelectionTypeEnum ) = 'integer'
'WdCaptionLabelIDEnum' = c(
   'wdCaptionFigure' = -1,
   'wdCaptionTable' = -2,
   'wdCaptionEquation' = -3 
 )
storage.mode( WdCaptionLabelIDEnum ) = 'integer'
'WdReferenceTypeEnum' = c(
   'wdRefTypeNumberedItem' = 0,
   'wdRefTypeHeading' = 1,
   'wdRefTypeBookmark' = 2,
   'wdRefTypeFootnote' = 3,
   'wdRefTypeEndnote' = 4 
 )
storage.mode( WdReferenceTypeEnum ) = 'integer'
'WdReferenceKindEnum' = c(
   'wdContentText' = -1,
   'wdNumberRelativeContext' = -2,
   'wdNumberNoContext' = -3,
   'wdNumberFullContext' = -4,
   'wdEntireCaption' = 2,
   'wdOnlyLabelAndNumber' = 3,
   'wdOnlyCaptionText' = 4,
   'wdFootnoteNumber' = 5,
   'wdEndnoteNumber' = 6,
   'wdPageNumber' = 7,
   'wdPosition' = 15,
   'wdFootnoteNumberFormatted' = 16,
   'wdEndnoteNumberFormatted' = 17 
 )
storage.mode( WdReferenceKindEnum ) = 'integer'
'WdIndexFormatEnum' = c(
   'wdIndexTemplate' = 0,
   'wdIndexClassic' = 1,
   'wdIndexFancy' = 2,
   'wdIndexModern' = 3,
   'wdIndexBulleted' = 4,
   'wdIndexFormal' = 5,
   'wdIndexSimple' = 6 
 )
storage.mode( WdIndexFormatEnum ) = 'integer'
'WdIndexTypeEnum' = c(
   'wdIndexIndent' = 0,
   'wdIndexRunin' = 1 
 )
storage.mode( WdIndexTypeEnum ) = 'integer'
'WdRevisionsWrapEnum' = c(
   'wdWrapNever' = 0,
   'wdWrapAlways' = 1,
   'wdWrapAsk' = 2 
 )
storage.mode( WdRevisionsWrapEnum ) = 'integer'
'WdRevisionTypeEnum' = c(
   'wdNoRevision' = 0,
   'wdRevisionInsert' = 1,
   'wdRevisionDelete' = 2,
   'wdRevisionProperty' = 3,
   'wdRevisionParagraphNumber' = 4,
   'wdRevisionDisplayField' = 5,
   'wdRevisionReconcile' = 6,
   'wdRevisionConflict' = 7,
   'wdRevisionStyle' = 8,
   'wdRevisionReplace' = 9,
   'wdRevisionParagraphProperty' = 10,
   'wdRevisionTableProperty' = 11,
   'wdRevisionSectionProperty' = 12,
   'wdRevisionStyleDefinition' = 13 
 )
storage.mode( WdRevisionTypeEnum ) = 'integer'
'WdRoutingSlipDeliveryEnum' = c(
   'wdOneAfterAnother' = 0,
   'wdAllAtOnce' = 1 
 )
storage.mode( WdRoutingSlipDeliveryEnum ) = 'integer'
'WdRoutingSlipStatusEnum' = c(
   'wdNotYetRouted' = 0,
   'wdRouteInProgress' = 1,
   'wdRouteComplete' = 2 
 )
storage.mode( WdRoutingSlipStatusEnum ) = 'integer'
'WdSectionStartEnum' = c(
   'wdSectionContinuous' = 0,
   'wdSectionNewColumn' = 1,
   'wdSectionNewPage' = 2,
   'wdSectionEvenPage' = 3,
   'wdSectionOddPage' = 4 
 )
storage.mode( WdSectionStartEnum ) = 'integer'
'WdSaveOptionsEnum' = c(
   'wdDoNotSaveChanges' = 0,
   'wdSaveChanges' = -1,
   'wdPromptToSaveChanges' = -2 
 )
storage.mode( WdSaveOptionsEnum ) = 'integer'
'WdDocumentKindEnum' = c(
   'wdDocumentNotSpecified' = 0,
   'wdDocumentLetter' = 1,
   'wdDocumentEmail' = 2 
 )
storage.mode( WdDocumentKindEnum ) = 'integer'
'WdDocumentTypeEnum' = c(
   'wdTypeDocument' = 0,
   'wdTypeTemplate' = 1,
   'wdTypeFrameset' = 2 
 )
storage.mode( WdDocumentTypeEnum ) = 'integer'
'WdOriginalFormatEnum' = c(
   'wdWordDocument' = 0,
   'wdOriginalDocumentFormat' = 1,
   'wdPromptUser' = 2 
 )
storage.mode( WdOriginalFormatEnum ) = 'integer'
'WdRelocateEnum' = c(
   'wdRelocateUp' = 0,
   'wdRelocateDown' = 1 
 )
storage.mode( WdRelocateEnum ) = 'integer'
'WdInsertedTextMarkEnum' = c(
   'wdInsertedTextMarkNone' = 0,
   'wdInsertedTextMarkBold' = 1,
   'wdInsertedTextMarkItalic' = 2,
   'wdInsertedTextMarkUnderline' = 3,
   'wdInsertedTextMarkDoubleUnderline' = 4,
   'wdInsertedTextMarkColorOnly' = 5,
   'wdInsertedTextMarkStrikeThrough' = 6 
 )
storage.mode( WdInsertedTextMarkEnum ) = 'integer'
'WdRevisedLinesMarkEnum' = c(
   'wdRevisedLinesMarkNone' = 0,
   'wdRevisedLinesMarkLeftBorder' = 1,
   'wdRevisedLinesMarkRightBorder' = 2,
   'wdRevisedLinesMarkOutsideBorder' = 3 
 )
storage.mode( WdRevisedLinesMarkEnum ) = 'integer'
'WdDeletedTextMarkEnum' = c(
   'wdDeletedTextMarkHidden' = 0,
   'wdDeletedTextMarkStrikeThrough' = 1,
   'wdDeletedTextMarkCaret' = 2,
   'wdDeletedTextMarkPound' = 3,
   'wdDeletedTextMarkNone' = 4,
   'wdDeletedTextMarkBold' = 5,
   'wdDeletedTextMarkItalic' = 6,
   'wdDeletedTextMarkUnderline' = 7,
   'wdDeletedTextMarkDoubleUnderline' = 8,
   'wdDeletedTextMarkColorOnly' = 9 
 )
storage.mode( WdDeletedTextMarkEnum ) = 'integer'
'WdRevisedPropertiesMarkEnum' = c(
   'wdRevisedPropertiesMarkNone' = 0,
   'wdRevisedPropertiesMarkBold' = 1,
   'wdRevisedPropertiesMarkItalic' = 2,
   'wdRevisedPropertiesMarkUnderline' = 3,
   'wdRevisedPropertiesMarkDoubleUnderline' = 4,
   'wdRevisedPropertiesMarkColorOnly' = 5,
   'wdRevisedPropertiesMarkStrikeThrough' = 6 
 )
storage.mode( WdRevisedPropertiesMarkEnum ) = 'integer'
'WdFieldShadingEnum' = c(
   'wdFieldShadingNever' = 0,
   'wdFieldShadingAlways' = 1,
   'wdFieldShadingWhenSelected' = 2 
 )
storage.mode( WdFieldShadingEnum ) = 'integer'
'WdDefaultFilePathEnum' = c(
   'wdDocumentsPath' = 0,
   'wdPicturesPath' = 1,
   'wdUserTemplatesPath' = 2,
   'wdWorkgroupTemplatesPath' = 3,
   'wdUserOptionsPath' = 4,
   'wdAutoRecoverPath' = 5,
   'wdToolsPath' = 6,
   'wdTutorialPath' = 7,
   'wdStartupPath' = 8,
   'wdProgramPath' = 9,
   'wdGraphicsFiltersPath' = 10,
   'wdTextConvertersPath' = 11,
   'wdProofingToolsPath' = 12,
   'wdTempFilePath' = 13,
   'wdCurrentFolderPath' = 14,
   'wdStyleGalleryPath' = 15,
   'wdBorderArtPath' = 19 
 )
storage.mode( WdDefaultFilePathEnum ) = 'integer'
'WdCompatibilityEnum' = c(
   'wdNoTabHangIndent' = 1,
   'wdNoSpaceRaiseLower' = 2,
   'wdPrintColBlack' = 3,
   'wdWrapTrailSpaces' = 4,
   'wdNoColumnBalance' = 5,
   'wdConvMailMergeEsc' = 6,
   'wdSuppressSpBfAfterPgBrk' = 7,
   'wdSuppressTopSpacing' = 8,
   'wdOrigWordTableRules' = 9,
   'wdTransparentMetafiles' = 10,
   'wdShowBreaksInFrames' = 11,
   'wdSwapBordersFacingPages' = 12,
   'wdLeaveBackslashAlone' = 13,
   'wdExpandShiftReturn' = 14,
   'wdDontULTrailSpace' = 15,
   'wdDontBalanceSingleByteDoubleByteWidth' = 16,
   'wdSuppressTopSpacingMac5' = 17,
   'wdSpacingInWholePoints' = 18,
   'wdPrintBodyTextBeforeHeader' = 19,
   'wdNoLeading' = 20,
   'wdNoSpaceForUL' = 21,
   'wdMWSmallCaps' = 22,
   'wdNoExtraLineSpacing' = 23,
   'wdTruncateFontHeight' = 24,
   'wdSubFontBySize' = 25,
   'wdUsePrinterMetrics' = 26,
   'wdWW6BorderRules' = 27,
   'wdExactOnTop' = 28,
   'wdSuppressBottomSpacing' = 29,
   'wdWPSpaceWidth' = 30,
   'wdWPJustification' = 31,
   'wdLineWrapLikeWord6' = 32,
   'wdShapeLayoutLikeWW8' = 33,
   'wdFootnoteLayoutLikeWW8' = 34,
   'wdDontUseHTMLParagraphAutoSpacing' = 35,
   'wdDontAdjustLineHeightInTable' = 36,
   'wdForgetLastTabAlignment' = 37,
   'wdAutospaceLikeWW7' = 38,
   'wdAlignTablesRowByRow' = 39,
   'wdLayoutRawTableWidth' = 40,
   'wdLayoutTableRowsApart' = 41,
   'wdUseWord97LineBreakingRules' = 42,
   'wdDontBreakWrappedTables' = 43,
   'wdDontSnapTextToGridInTableWithObjects' = 44,
   'wdSelectFieldWithFirstOrLastCharacter' = 45,
   'wdApplyBreakingRules' = 46,
   'wdDontWrapTextWithPunctuation' = 47,
   'wdDontUseAsianBreakRulesInGrid' = 48,
   'wdUseWord2002TableStyleRules' = 49,
   'wdGrowAutofit' = 50 
 )
storage.mode( WdCompatibilityEnum ) = 'integer'
'WdPaperSizeEnum' = c(
   'wdPaper10x14' = 0,
   'wdPaper11x17' = 1,
   'wdPaperLetter' = 2,
   'wdPaperLetterSmall' = 3,
   'wdPaperLegal' = 4,
   'wdPaperExecutive' = 5,
   'wdPaperA3' = 6,
   'wdPaperA4' = 7,
   'wdPaperA4Small' = 8,
   'wdPaperA5' = 9,
   'wdPaperB4' = 10,
   'wdPaperB5' = 11,
   'wdPaperCSheet' = 12,
   'wdPaperDSheet' = 13,
   'wdPaperESheet' = 14,
   'wdPaperFanfoldLegalGerman' = 15,
   'wdPaperFanfoldStdGerman' = 16,
   'wdPaperFanfoldUS' = 17,
   'wdPaperFolio' = 18,
   'wdPaperLedger' = 19,
   'wdPaperNote' = 20,
   'wdPaperQuarto' = 21,
   'wdPaperStatement' = 22,
   'wdPaperTabloid' = 23,
   'wdPaperEnvelope9' = 24,
   'wdPaperEnvelope10' = 25,
   'wdPaperEnvelope11' = 26,
   'wdPaperEnvelope12' = 27,
   'wdPaperEnvelope14' = 28,
   'wdPaperEnvelopeB4' = 29,
   'wdPaperEnvelopeB5' = 30,
   'wdPaperEnvelopeB6' = 31,
   'wdPaperEnvelopeC3' = 32,
   'wdPaperEnvelopeC4' = 33,
   'wdPaperEnvelopeC5' = 34,
   'wdPaperEnvelopeC6' = 35,
   'wdPaperEnvelopeC65' = 36,
   'wdPaperEnvelopeDL' = 37,
   'wdPaperEnvelopeItaly' = 38,
   'wdPaperEnvelopeMonarch' = 39,
   'wdPaperEnvelopePersonal' = 40,
   'wdPaperCustom' = 41 
 )
storage.mode( WdPaperSizeEnum ) = 'integer'
'WdCustomLabelPageSizeEnum' = c(
   'wdCustomLabelLetter' = 0,
   'wdCustomLabelLetterLS' = 1,
   'wdCustomLabelA4' = 2,
   'wdCustomLabelA4LS' = 3,
   'wdCustomLabelA5' = 4,
   'wdCustomLabelA5LS' = 5,
   'wdCustomLabelB5' = 6,
   'wdCustomLabelMini' = 7,
   'wdCustomLabelFanfold' = 8,
   'wdCustomLabelVertHalfSheet' = 9,
   'wdCustomLabelVertHalfSheetLS' = 10,
   'wdCustomLabelHigaki' = 11,
   'wdCustomLabelHigakiLS' = 12,
   'wdCustomLabelB4JIS' = 13 
 )
storage.mode( WdCustomLabelPageSizeEnum ) = 'integer'
'WdProtectionTypeEnum' = c(
   'wdNoProtection' = -1,
   'wdAllowOnlyRevisions' = 0,
   'wdAllowOnlyComments' = 1,
   'wdAllowOnlyFormFields' = 2,
   'wdAllowOnlyReading' = 3 
 )
storage.mode( WdProtectionTypeEnum ) = 'integer'
'WdPartOfSpeechEnum' = c(
   'wdAdjective' = 0,
   'wdNoun' = 1,
   'wdAdverb' = 2,
   'wdVerb' = 3,
   'wdPronoun' = 4,
   'wdConjunction' = 5,
   'wdPreposition' = 6,
   'wdInterjection' = 7,
   'wdIdiom' = 8,
   'wdOther' = 9 
 )
storage.mode( WdPartOfSpeechEnum ) = 'integer'
'WdSubscriberFormatsEnum' = c(
   'wdSubscriberBestFormat' = 0,
   'wdSubscriberRTF' = 1,
   'wdSubscriberText' = 2,
   'wdSubscriberPict' = 4 
 )
storage.mode( WdSubscriberFormatsEnum ) = 'integer'
'WdEditionTypeEnum' = c(
   'wdPublisher' = 0,
   'wdSubscriber' = 1 
 )
storage.mode( WdEditionTypeEnum ) = 'integer'
'WdEditionOptionEnum' = c(
   'wdCancelPublisher' = 0,
   'wdSendPublisher' = 1,
   'wdSelectPublisher' = 2,
   'wdAutomaticUpdate' = 3,
   'wdManualUpdate' = 4,
   'wdChangeAttributes' = 5,
   'wdUpdateSubscriber' = 6,
   'wdOpenSource' = 7 
 )
storage.mode( WdEditionOptionEnum ) = 'integer'
'WdRelativeHorizontalPositionEnum' = c(
   'wdRelativeHorizontalPositionMargin' = 0,
   'wdRelativeHorizontalPositionPage' = 1,
   'wdRelativeHorizontalPositionColumn' = 2,
   'wdRelativeHorizontalPositionCharacter' = 3 
 )
storage.mode( WdRelativeHorizontalPositionEnum ) = 'integer'
'WdRelativeVerticalPositionEnum' = c(
   'wdRelativeVerticalPositionMargin' = 0,
   'wdRelativeVerticalPositionPage' = 1,
   'wdRelativeVerticalPositionParagraph' = 2,
   'wdRelativeVerticalPositionLine' = 3 
 )
storage.mode( WdRelativeVerticalPositionEnum ) = 'integer'
'WdHelpTypeEnum' = c(
   'wdHelp' = 0,
   'wdHelpAbout' = 1,
   'wdHelpActiveWindow' = 2,
   'wdHelpContents' = 3,
   'wdHelpExamplesAndDemos' = 4,
   'wdHelpIndex' = 5,
   'wdHelpKeyboard' = 6,
   'wdHelpPSSHelp' = 7,
   'wdHelpQuickPreview' = 8,
   'wdHelpSearch' = 9,
   'wdHelpUsingHelp' = 10,
   'wdHelpIchitaro' = 11,
   'wdHelpPE2' = 12,
   'wdHelpHWP' = 13 
 )
storage.mode( WdHelpTypeEnum ) = 'integer'
'WdHelpTypeHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdHelpTypeHIDEnum ) = 'integer'
'WdKeyCategoryEnum' = c(
   'wdKeyCategoryNil' = -1,
   'wdKeyCategoryDisable' = 0,
   'wdKeyCategoryCommand' = 1,
   'wdKeyCategoryMacro' = 2,
   'wdKeyCategoryFont' = 3,
   'wdKeyCategoryAutoText' = 4,
   'wdKeyCategoryStyle' = 5,
   'wdKeyCategorySymbol' = 6,
   'wdKeyCategoryPrefix' = 7 
 )
storage.mode( WdKeyCategoryEnum ) = 'integer'
'WdKeyEnum' = c(
   'wdNoKey' = 255,
   'wdKeyShift' = 256,
   'wdKeyControl' = 512,
   'wdKeyCommand' = 512,
   'wdKeyAlt' = 1024,
   'wdKeyOption' = 1024,
   'wdKeyA' = 65,
   'wdKeyB' = 66,
   'wdKeyC' = 67,
   'wdKeyD' = 68,
   'wdKeyE' = 69,
   'wdKeyF' = 70,
   'wdKeyG' = 71,
   'wdKeyH' = 72,
   'wdKeyI' = 73,
   'wdKeyJ' = 74,
   'wdKeyK' = 75,
   'wdKeyL' = 76,
   'wdKeyM' = 77,
   'wdKeyN' = 78,
   'wdKeyO' = 79,
   'wdKeyP' = 80,
   'wdKeyQ' = 81,
   'wdKeyR' = 82,
   'wdKeyS' = 83,
   'wdKeyT' = 84,
   'wdKeyU' = 85,
   'wdKeyV' = 86,
   'wdKeyW' = 87,
   'wdKeyX' = 88,
   'wdKeyY' = 89,
   'wdKeyZ' = 90,
   'wdKey0' = 48,
   'wdKey1' = 49,
   'wdKey2' = 50,
   'wdKey3' = 51,
   'wdKey4' = 52,
   'wdKey5' = 53,
   'wdKey6' = 54,
   'wdKey7' = 55,
   'wdKey8' = 56,
   'wdKey9' = 57,
   'wdKeyBackspace' = 8,
   'wdKeyTab' = 9,
   'wdKeyNumeric5Special' = 12,
   'wdKeyReturn' = 13,
   'wdKeyPause' = 19,
   'wdKeyEsc' = 27,
   'wdKeySpacebar' = 32,
   'wdKeyPageUp' = 33,
   'wdKeyPageDown' = 34,
   'wdKeyEnd' = 35,
   'wdKeyHome' = 36,
   'wdKeyInsert' = 45,
   'wdKeyDelete' = 46,
   'wdKeyNumeric0' = 96,
   'wdKeyNumeric1' = 97,
   'wdKeyNumeric2' = 98,
   'wdKeyNumeric3' = 99,
   'wdKeyNumeric4' = 100,
   'wdKeyNumeric5' = 101,
   'wdKeyNumeric6' = 102,
   'wdKeyNumeric7' = 103,
   'wdKeyNumeric8' = 104,
   'wdKeyNumeric9' = 105,
   'wdKeyNumericMultiply' = 106,
   'wdKeyNumericAdd' = 107,
   'wdKeyNumericSubtract' = 109,
   'wdKeyNumericDecimal' = 110,
   'wdKeyNumericDivide' = 111,
   'wdKeyF1' = 112,
   'wdKeyF2' = 113,
   'wdKeyF3' = 114,
   'wdKeyF4' = 115,
   'wdKeyF5' = 116,
   'wdKeyF6' = 117,
   'wdKeyF7' = 118,
   'wdKeyF8' = 119,
   'wdKeyF9' = 120,
   'wdKeyF10' = 121,
   'wdKeyF11' = 122,
   'wdKeyF12' = 123,
   'wdKeyF13' = 124,
   'wdKeyF14' = 125,
   'wdKeyF15' = 126,
   'wdKeyF16' = 127,
   'wdKeyScrollLock' = 145,
   'wdKeySemiColon' = 186,
   'wdKeyEquals' = 187,
   'wdKeyComma' = 188,
   'wdKeyHyphen' = 189,
   'wdKeyPeriod' = 190,
   'wdKeySlash' = 191,
   'wdKeyBackSingleQuote' = 192,
   'wdKeyOpenSquareBrace' = 219,
   'wdKeyBackSlash' = 220,
   'wdKeyCloseSquareBrace' = 221,
   'wdKeySingleQuote' = 222 
 )
storage.mode( WdKeyEnum ) = 'integer'
'WdOLETypeEnum' = c(
   'wdOLELink' = 0,
   'wdOLEEmbed' = 1,
   'wdOLEControl' = 2 
 )
storage.mode( WdOLETypeEnum ) = 'integer'
'WdOLEVerbEnum' = c(
   'wdOLEVerbPrimary' = 0,
   'wdOLEVerbShow' = -1,
   'wdOLEVerbOpen' = -2,
   'wdOLEVerbHide' = -3,
   'wdOLEVerbUIActivate' = -4,
   'wdOLEVerbInPlaceActivate' = -5,
   'wdOLEVerbDiscardUndoState' = -6 
 )
storage.mode( WdOLEVerbEnum ) = 'integer'
'WdOLEPlacementEnum' = c(
   'wdInLine' = 0,
   'wdFloatOverText' = 1 
 )
storage.mode( WdOLEPlacementEnum ) = 'integer'
'WdEnvelopeOrientationEnum' = c(
   'wdLeftPortrait' = 0,
   'wdCenterPortrait' = 1,
   'wdRightPortrait' = 2,
   'wdLeftLandscape' = 3,
   'wdCenterLandscape' = 4,
   'wdRightLandscape' = 5,
   'wdLeftClockwise' = 6,
   'wdCenterClockwise' = 7,
   'wdRightClockwise' = 8 
 )
storage.mode( WdEnvelopeOrientationEnum ) = 'integer'
'WdLetterStyleEnum' = c(
   'wdFullBlock' = 0,
   'wdModifiedBlock' = 1,
   'wdSemiBlock' = 2 
 )
storage.mode( WdLetterStyleEnum ) = 'integer'
'WdLetterheadLocationEnum' = c(
   'wdLetterTop' = 0,
   'wdLetterBottom' = 1,
   'wdLetterLeft' = 2,
   'wdLetterRight' = 3 
 )
storage.mode( WdLetterheadLocationEnum ) = 'integer'
'WdSalutationTypeEnum' = c(
   'wdSalutationInformal' = 0,
   'wdSalutationFormal' = 1,
   'wdSalutationBusiness' = 2,
   'wdSalutationOther' = 3 
 )
storage.mode( WdSalutationTypeEnum ) = 'integer'
'WdSalutationGenderEnum' = c(
   'wdGenderFemale' = 0,
   'wdGenderMale' = 1,
   'wdGenderNeutral' = 2,
   'wdGenderUnknown' = 3 
 )
storage.mode( WdSalutationGenderEnum ) = 'integer'
'WdMovementTypeEnum' = c(
   'wdMove' = 0,
   'wdExtend' = 1 
 )
storage.mode( WdMovementTypeEnum ) = 'integer'
'WdConstantsEnum' = c(
   'wdUndefined' = 9999999,
   'wdToggle' = 9999998,
   'wdForward' = 1073741823,
   'wdBackward' = -1073741823,
   'wdAutoPosition' = 0,
   'wdFirst' = 1,
   'wdCreatorCode' = 1297307460 
 )
storage.mode( WdConstantsEnum ) = 'integer'
'WdPasteDataTypeEnum' = c(
   'wdPasteOLEObject' = 0,
   'wdPasteRTF' = 1,
   'wdPasteText' = 2,
   'wdPasteMetafilePicture' = 3,
   'wdPasteBitmap' = 4,
   'wdPasteDeviceIndependentBitmap' = 5,
   'wdPasteHyperlink' = 7,
   'wdPasteShape' = 8,
   'wdPasteEnhancedMetafile' = 9,
   'wdPasteHTML' = 10 
 )
storage.mode( WdPasteDataTypeEnum ) = 'integer'
'WdPrintOutItemEnum' = c(
   'wdPrintDocumentContent' = 0,
   'wdPrintProperties' = 1,
   'wdPrintComments' = 2,
   'wdPrintMarkup' = 2,
   'wdPrintStyles' = 3,
   'wdPrintAutoTextEntries' = 4,
   'wdPrintKeyAssignments' = 5,
   'wdPrintEnvelope' = 6,
   'wdPrintDocumentWithMarkup' = 7 
 )
storage.mode( WdPrintOutItemEnum ) = 'integer'
'WdPrintOutPagesEnum' = c(
   'wdPrintAllPages' = 0,
   'wdPrintOddPagesOnly' = 1,
   'wdPrintEvenPagesOnly' = 2 
 )
storage.mode( WdPrintOutPagesEnum ) = 'integer'
'WdPrintOutRangeEnum' = c(
   'wdPrintAllDocument' = 0,
   'wdPrintSelection' = 1,
   'wdPrintCurrentPage' = 2,
   'wdPrintFromTo' = 3,
   'wdPrintRangeOfPages' = 4 
 )
storage.mode( WdPrintOutRangeEnum ) = 'integer'
'WdDictionaryTypeEnum' = c(
   'wdSpelling' = 0,
   'wdGrammar' = 1,
   'wdThesaurus' = 2,
   'wdHyphenation' = 3,
   'wdSpellingComplete' = 4,
   'wdSpellingCustom' = 5,
   'wdSpellingLegal' = 6,
   'wdSpellingMedical' = 7,
   'wdHangulHanjaConversion' = 8,
   'wdHangulHanjaConversionCustom' = 9 
 )
storage.mode( WdDictionaryTypeEnum ) = 'integer'
'WdDictionaryTypeHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdDictionaryTypeHIDEnum ) = 'integer'
'WdSpellingWordTypeEnum' = c(
   'wdSpellword' = 0,
   'wdWildcard' = 1,
   'wdAnagram' = 2 
 )
storage.mode( WdSpellingWordTypeEnum ) = 'integer'
'WdSpellingErrorTypeEnum' = c(
   'wdSpellingCorrect' = 0,
   'wdSpellingNotInDictionary' = 1,
   'wdSpellingCapitalization' = 2 
 )
storage.mode( WdSpellingErrorTypeEnum ) = 'integer'
'WdProofreadingErrorTypeEnum' = c(
   'wdSpellingError' = 0,
   'wdGrammaticalError' = 1 
 )
storage.mode( WdProofreadingErrorTypeEnum ) = 'integer'
'WdInlineShapeTypeEnum' = c(
   'wdInlineShapeEmbeddedOLEObject' = 1,
   'wdInlineShapeLinkedOLEObject' = 2,
   'wdInlineShapePicture' = 3,
   'wdInlineShapeLinkedPicture' = 4,
   'wdInlineShapeOLEControlObject' = 5,
   'wdInlineShapeHorizontalLine' = 6,
   'wdInlineShapePictureHorizontalLine' = 7,
   'wdInlineShapeLinkedPictureHorizontalLine' = 8,
   'wdInlineShapePictureBullet' = 9,
   'wdInlineShapeScriptAnchor' = 10,
   'wdInlineShapeOWSAnchor' = 11 
 )
storage.mode( WdInlineShapeTypeEnum ) = 'integer'
'WdArrangeStyleEnum' = c(
   'wdTiled' = 0,
   'wdIcons' = 1 
 )
storage.mode( WdArrangeStyleEnum ) = 'integer'
'WdSelectionFlagsEnum' = c(
   'wdSelStartActive' = 1,
   'wdSelAtEOL' = 2,
   'wdSelOvertype' = 4,
   'wdSelActive' = 8,
   'wdSelReplace' = 16 
 )
storage.mode( WdSelectionFlagsEnum ) = 'integer'
'WdAutoVersionsEnum' = c(
   'wdAutoVersionOff' = 0,
   'wdAutoVersionOnClose' = 1 
 )
storage.mode( WdAutoVersionsEnum ) = 'integer'
'WdOrganizerObjectEnum' = c(
   'wdOrganizerObjectStyles' = 0,
   'wdOrganizerObjectAutoText' = 1,
   'wdOrganizerObjectCommandBars' = 2,
   'wdOrganizerObjectProjectItems' = 3 
 )
storage.mode( WdOrganizerObjectEnum ) = 'integer'
'WdFindMatchEnum' = c(
   'wdMatchParagraphMark' = 65551,
   'wdMatchTabCharacter' = 9,
   'wdMatchCommentMark' = 5,
   'wdMatchAnyCharacter' = 65599,
   'wdMatchAnyDigit' = 65567,
   'wdMatchAnyLetter' = 65583,
   'wdMatchCaretCharacter' = 11,
   'wdMatchColumnBreak' = 14,
   'wdMatchEmDash' = 8212,
   'wdMatchEnDash' = 8211,
   'wdMatchEndnoteMark' = 65555,
   'wdMatchField' = 19,
   'wdMatchFootnoteMark' = 65554,
   'wdMatchGraphic' = 1,
   'wdMatchManualLineBreak' = 65551,
   'wdMatchManualPageBreak' = 65564,
   'wdMatchNonbreakingHyphen' = 30,
   'wdMatchNonbreakingSpace' = 160,
   'wdMatchOptionalHyphen' = 31,
   'wdMatchSectionBreak' = 65580,
   'wdMatchWhiteSpace' = 65655 
 )
storage.mode( WdFindMatchEnum ) = 'integer'
'WdFindWrapEnum' = c(
   'wdFindStop' = 0,
   'wdFindContinue' = 1,
   'wdFindAsk' = 2 
 )
storage.mode( WdFindWrapEnum ) = 'integer'
'WdInformationEnum' = c(
   'wdActiveEndAdjustedPageNumber' = 1,
   'wdActiveEndSectionNumber' = 2,
   'wdActiveEndPageNumber' = 3,
   'wdNumberOfPagesInDocument' = 4,
   'wdHorizontalPositionRelativeToPage' = 5,
   'wdVerticalPositionRelativeToPage' = 6,
   'wdHorizontalPositionRelativeToTextBoundary' = 7,
   'wdVerticalPositionRelativeToTextBoundary' = 8,
   'wdFirstCharacterColumnNumber' = 9,
   'wdFirstCharacterLineNumber' = 10,
   'wdFrameIsSelected' = 11,
   'wdWithInTable' = 12,
   'wdStartOfRangeRowNumber' = 13,
   'wdEndOfRangeRowNumber' = 14,
   'wdMaximumNumberOfRows' = 15,
   'wdStartOfRangeColumnNumber' = 16,
   'wdEndOfRangeColumnNumber' = 17,
   'wdMaximumNumberOfColumns' = 18,
   'wdZoomPercentage' = 19,
   'wdSelectionMode' = 20,
   'wdCapsLock' = 21,
   'wdNumLock' = 22,
   'wdOverType' = 23,
   'wdRevisionMarking' = 24,
   'wdInFootnoteEndnotePane' = 25,
   'wdInCommentPane' = 26,
   'wdInHeaderFooter' = 28,
   'wdAtEndOfRowMarker' = 31,
   'wdReferenceOfType' = 32,
   'wdHeaderFooterType' = 33,
   'wdInMasterDocument' = 34,
   'wdInFootnote' = 35,
   'wdInEndnote' = 36,
   'wdInWordMail' = 37,
   'wdInClipboard' = 38 
 )
storage.mode( WdInformationEnum ) = 'integer'
'WdWrapTypeEnum' = c(
   'wdWrapSquare' = 0,
   'wdWrapTight' = 1,
   'wdWrapThrough' = 2,
   'wdWrapNone' = 3,
   'wdWrapTopBottom' = 4,
   'wdWrapInline' = 7 
 )
storage.mode( WdWrapTypeEnum ) = 'integer'
'WdWrapSideTypeEnum' = c(
   'wdWrapBoth' = 0,
   'wdWrapLeft' = 1,
   'wdWrapRight' = 2,
   'wdWrapLargest' = 3 
 )
storage.mode( WdWrapSideTypeEnum ) = 'integer'
'WdOutlineLevelEnum' = c(
   'wdOutlineLevel1' = 1,
   'wdOutlineLevel2' = 2,
   'wdOutlineLevel3' = 3,
   'wdOutlineLevel4' = 4,
   'wdOutlineLevel5' = 5,
   'wdOutlineLevel6' = 6,
   'wdOutlineLevel7' = 7,
   'wdOutlineLevel8' = 8,
   'wdOutlineLevel9' = 9,
   'wdOutlineLevelBodyText' = 10 
 )
storage.mode( WdOutlineLevelEnum ) = 'integer'
'WdTextOrientationEnum' = c(
   'wdTextOrientationHorizontal' = 0,
   'wdTextOrientationUpward' = 2,
   'wdTextOrientationDownward' = 3,
   'wdTextOrientationVerticalFarEast' = 1,
   'wdTextOrientationHorizontalRotatedFarEast' = 4 
 )
storage.mode( WdTextOrientationEnum ) = 'integer'
'WdTextOrientationHIDEnum' = c(
   'emptyenum' = 0 
 )
storage.mode( WdTextOrientationHIDEnum ) = 'integer'
'WdPageBorderArtEnum' = c(
   'wdArtApples' = 1,
   'wdArtMapleMuffins' = 2,
   'wdArtCakeSlice' = 3,
   'wdArtCandyCorn' = 4,
   'wdArtIceCreamCones' = 5,
   'wdArtChampagneBottle' = 6,
   'wdArtPartyGlass' = 7,
   'wdArtChristmasTree' = 8,
   'wdArtTrees' = 9,
   'wdArtPalmsColor' = 10,
   'wdArtBalloons3Colors' = 11,
   'wdArtBalloonsHotAir' = 12,
   'wdArtPartyFavor' = 13,
   'wdArtConfettiStreamers' = 14,
   'wdArtHearts' = 15,
   'wdArtHeartBalloon' = 16,
   'wdArtStars3D' = 17,
   'wdArtStarsShadowed' = 18,
   'wdArtStars' = 19,
   'wdArtSun' = 20,
   'wdArtEarth2' = 21,
   'wdArtEarth1' = 22,
   'wdArtPeopleHats' = 23,
   'wdArtSombrero' = 24,
   'wdArtPencils' = 25,
   'wdArtPackages' = 26,
   'wdArtClocks' = 27,
   'wdArtFirecrackers' = 28,
   'wdArtRings' = 29,
   'wdArtMapPins' = 30,
   'wdArtConfetti' = 31,
   'wdArtCreaturesButterfly' = 32,
   'wdArtCreaturesLadyBug' = 33,
   'wdArtCreaturesFish' = 34,
   'wdArtBirdsFlight' = 35,
   'wdArtScaredCat' = 36,
   'wdArtBats' = 37,
   'wdArtFlowersRoses' = 38,
   'wdArtFlowersRedRose' = 39,
   'wdArtPoinsettias' = 40,
   'wdArtHolly' = 41,
   'wdArtFlowersTiny' = 42,
   'wdArtFlowersPansy' = 43,
   'wdArtFlowersModern2' = 44,
   'wdArtFlowersModern1' = 45,
   'wdArtWhiteFlowers' = 46,
   'wdArtVine' = 47,
   'wdArtFlowersDaisies' = 48,
   'wdArtFlowersBlockPrint' = 49,
   'wdArtDecoArchColor' = 50,
   'wdArtFans' = 51,
   'wdArtFilm' = 52,
   'wdArtLightning1' = 53,
   'wdArtCompass' = 54,
   'wdArtDoubleD' = 55,
   'wdArtClassicalWave' = 56,
   'wdArtShadowedSquares' = 57,
   'wdArtTwistedLines1' = 58,
   'wdArtWaveline' = 59,
   'wdArtQuadrants' = 60,
   'wdArtCheckedBarColor' = 61,
   'wdArtSwirligig' = 62,
   'wdArtPushPinNote1' = 63,
   'wdArtPushPinNote2' = 64,
   'wdArtPumpkin1' = 65,
   'wdArtEggsBlack' = 66,
   'wdArtCup' = 67,
   'wdArtHeartGray' = 68,
   'wdArtGingerbreadMan' = 69,
   'wdArtBabyPacifier' = 70,
   'wdArtBabyRattle' = 71,
   'wdArtCabins' = 72,
   'wdArtHouseFunky' = 73,
   'wdArtStarsBlack' = 74,
   'wdArtSnowflakes' = 75,
   'wdArtSnowflakeFancy' = 76,
   'wdArtSkyrocket' = 77,
   'wdArtSeattle' = 78,
   'wdArtMusicNotes' = 79,
   'wdArtPalmsBlack' = 80,
   'wdArtMapleLeaf' = 81,
   'wdArtPaperClips' = 82,
   'wdArtShorebirdTracks' = 83,
   'wdArtPeople' = 84,
   'wdArtPeopleWaving' = 85,
   'wdArtEclipsingSquares2' = 86,
   'wdArtHypnotic' = 87,
   'wdArtDiamondsGray' = 88,
   'wdArtDecoArch' = 89,
   'wdArtDecoBlocks' = 90,
   'wdArtCirclesLines' = 91,
   'wdArtPapyrus' = 92,
   'wdArtWoodwork' = 93,
   'wdArtWeavingBraid' = 94,
   'wdArtWeavingRibbon' = 95,
   'wdArtWeavingAngles' = 96,
   'wdArtArchedScallops' = 97,
   'wdArtSafari' = 98,
   'wdArtCelticKnotwork' = 99,
   'wdArtCrazyMaze' = 100,
   'wdArtEclipsingSquares1' = 101,
   'wdArtBirds' = 102,
   'wdArtFlowersTeacup' = 103,
   'wdArtNorthwest' = 104,
   'wdArtSouthwest' = 105,
   'wdArtTribal6' = 106,
   'wdArtTribal4' = 107,
   'wdArtTribal3' = 108,
   'wdArtTribal2' = 109,
   'wdArtTribal5' = 110,
   'wdArtXIllusions' = 111,
   'wdArtZanyTriangles' = 112,
   'wdArtPyramids' = 113,
   'wdArtPyramidsAbove' = 114,
   'wdArtConfettiGrays' = 115,
   'wdArtConfettiOutline' = 116,
   'wdArtConfettiWhite' = 117,
   'wdArtMosaic' = 118,
   'wdArtLightning2' = 119,
   'wdArtHeebieJeebies' = 120,
   'wdArtLightBulb' = 121,
   'wdArtGradient' = 122,
   'wdArtTriangleParty' = 123,
   'wdArtTwistedLines2' = 124,
   'wdArtMoons' = 125,
   'wdArtOvals' = 126,
   'wdArtDoubleDiamonds' = 127,
   'wdArtChainLink' = 128,
   'wdArtTriangles' = 129,
   'wdArtTribal1' = 130,
   'wdArtMarqueeToothed' = 131,
   'wdArtSharksTeeth' = 132,
   'wdArtSawtooth' = 133,
   'wdArtSawtoothGray' = 134,
   'wdArtPostageStamp' = 135,
   'wdArtWeavingStrips' = 136,
   'wdArtZigZag' = 137,
   'wdArtCrossStitch' = 138,
   'wdArtGems' = 139,
   'wdArtCirclesRectangles' = 140,
   'wdArtCornerTriangles' = 141,
   'wdArtCreaturesInsects' = 142,
   'wdArtZigZagStitch' = 143,
   'wdArtCheckered' = 144,
   'wdArtCheckedBarBlack' = 145,
   'wdArtMarquee' = 146,
   'wdArtBasicWhiteDots' = 147,
   'wdArtBasicWideMidline' = 148,
   'wdArtBasicWideOutline' = 149,
   'wdArtBasicWideInline' = 150,
   'wdArtBasicThinLines' = 151,
   'wdArtBasicWhiteDashes' = 152,
   'wdArtBasicWhiteSquares' = 153,
   'wdArtBasicBlackSquares' = 154,
   'wdArtBasicBlackDashes' = 155,
   'wdArtBasicBlackDots' = 156,
   'wdArtStarsTop' = 157,
   'wdArtCertificateBanner' = 158,
   'wdArtHandmade1' = 159,
   'wdArtHandmade2' = 160,
   'wdArtTornPaper' = 161,
   'wdArtTornPaperBlack' = 162,
   'wdArtCouponCutoutDashes' = 163,
   'wdArtCouponCutoutDots' = 164 
 )
storage.mode( WdPageBorderArtEnum ) = 'integer'
'WdBorderDistanceFromEnum' = c(
   'wdBorderDistanceFromText' = 0,
   'wdBorderDistanceFromPageEdge' = 1 
 )
storage.mode( WdBorderDistanceFromEnum ) = 'integer'
'WdReplaceEnum' = c(
   'wdReplaceNone' = 0,
   'wdReplaceOne' = 1,
   'wdReplaceAll' = 2 
 )
storage.mode( WdReplaceEnum ) = 'integer'
'WdFontBiasEnum' = c(
   'wdFontBiasDontCare' = 255,
   'wdFontBiasDefault' = 0,
   'wdFontBiasFareast' = 1 
 )
storage.mode( WdFontBiasEnum ) = 'integer'
'WdBrowserLevelEnum' = c(
   'wdBrowserLevelV4' = 0,
   'wdBrowserLevelMicrosoftInternetExplorer5' = 1,
   'wdBrowserLevelMicrosoftInternetExplorer6' = 2 
 )
storage.mode( WdBrowserLevelEnum ) = 'integer'
'WdEnclosureTypeEnum' = c(
   'wdEnclosureCircle' = 0,
   'wdEnclosureSquare' = 1,
   'wdEnclosureTriangle' = 2,
   'wdEnclosureDiamond' = 3 
 )
storage.mode( WdEnclosureTypeEnum ) = 'integer'
'WdEncloseStyleEnum' = c(
   'wdEncloseStyleNone' = 0,
   'wdEncloseStyleSmall' = 1,
   'wdEncloseStyleLarge' = 2 
 )
storage.mode( WdEncloseStyleEnum ) = 'integer'
'WdHighAnsiTextEnum' = c(
   'wdHighAnsiIsFarEast' = 0,
   'wdHighAnsiIsHighAnsi' = 1,
   'wdAutoDetectHighAnsiFarEast' = 2 
 )
storage.mode( WdHighAnsiTextEnum ) = 'integer'
'WdLayoutModeEnum' = c(
   'wdLayoutModeDefault' = 0,
   'wdLayoutModeGrid' = 1,
   'wdLayoutModeLineGrid' = 2,
   'wdLayoutModeGenko' = 3 
 )
storage.mode( WdLayoutModeEnum ) = 'integer'
'WdDocumentMediumEnum' = c(
   'wdEmailMessage' = 0,
   'wdDocument' = 1,
   'wdWebPage' = 2 
 )
storage.mode( WdDocumentMediumEnum ) = 'integer'
'WdMailerPriorityEnum' = c(
   'wdPriorityNormal' = 1,
   'wdPriorityLow' = 2,
   'wdPriorityHigh' = 3 
 )
storage.mode( WdMailerPriorityEnum ) = 'integer'
'WdDocumentViewDirectionEnum' = c(
   'wdDocumentViewRtl' = 0,
   'wdDocumentViewLtr' = 1 
 )
storage.mode( WdDocumentViewDirectionEnum ) = 'integer'
'WdArabicNumeralEnum' = c(
   'wdNumeralArabic' = 0,
   'wdNumeralHindi' = 1,
   'wdNumeralContext' = 2,
   'wdNumeralSystem' = 3 
 )
storage.mode( WdArabicNumeralEnum ) = 'integer'
'WdMonthNamesEnum' = c(
   'wdMonthNamesArabic' = 0,
   'wdMonthNamesEnglish' = 1,
   'wdMonthNamesFrench' = 2 
 )
storage.mode( WdMonthNamesEnum ) = 'integer'
'WdCursorMovementEnum' = c(
   'wdCursorMovementLogical' = 0,
   'wdCursorMovementVisual' = 1 
 )
storage.mode( WdCursorMovementEnum ) = 'integer'
'WdVisualSelectionEnum' = c(
   'wdVisualSelectionBlock' = 0,
   'wdVisualSelectionContinuous' = 1 
 )
storage.mode( WdVisualSelectionEnum ) = 'integer'
'WdTableDirectionEnum' = c(
   'wdTableDirectionRtl' = 0,
   'wdTableDirectionLtr' = 1 
 )
storage.mode( WdTableDirectionEnum ) = 'integer'
'WdFlowDirectionEnum' = c(
   'wdFlowLtr' = 0,
   'wdFlowRtl' = 1 
 )
storage.mode( WdFlowDirectionEnum ) = 'integer'
'WdDiacriticColorEnum' = c(
   'wdDiacriticColorBidi' = 0,
   'wdDiacriticColorLatin' = 1 
 )
storage.mode( WdDiacriticColorEnum ) = 'integer'
'WdGutterStyleEnum' = c(
   'wdGutterPosLeft' = 0,
   'wdGutterPosTop' = 1,
   'wdGutterPosRight' = 2 
 )
storage.mode( WdGutterStyleEnum ) = 'integer'
'WdGutterStyleOldEnum' = c(
   'wdGutterStyleLatin' = -10,
   'wdGutterStyleBidi' = 2 
 )
storage.mode( WdGutterStyleOldEnum ) = 'integer'
'WdSectionDirectionEnum' = c(
   'wdSectionDirectionRtl' = 0,
   'wdSectionDirectionLtr' = 1 
 )
storage.mode( WdSectionDirectionEnum ) = 'integer'
'WdDateLanguageEnum' = c(
   'wdDateLanguageBidi' = 10,
   'wdDateLanguageLatin' = 1033 
 )
storage.mode( WdDateLanguageEnum ) = 'integer'
'WdCalendarTypeBiEnum' = c(
   'wdCalendarTypeBidi' = 99,
   'wdCalendarTypeGregorian' = 100 
 )
storage.mode( WdCalendarTypeBiEnum ) = 'integer'
'WdCalendarTypeEnum' = c(
   'wdCalendarWestern' = 0,
   'wdCalendarArabic' = 1,
   'wdCalendarHebrew' = 2,
   'wdCalendarChina' = 3,
   'wdCalendarJapan' = 4,
   'wdCalendarThai' = 5,
   'wdCalendarKorean' = 6,
   'wdCalendarSakaEra' = 7 
 )
storage.mode( WdCalendarTypeEnum ) = 'integer'
'WdReadingOrderEnum' = c(
   'wdReadingOrderRtl' = 0,
   'wdReadingOrderLtr' = 1 
 )
storage.mode( WdReadingOrderEnum ) = 'integer'
'WdHebSpellStartEnum' = c(
   'wdFullScript' = 0,
   'wdPartialScript' = 1,
   'wdMixedScript' = 2,
   'wdMixedAuthorizedScript' = 3 
 )
storage.mode( WdHebSpellStartEnum ) = 'integer'
'WdAraSpellerEnum' = c(
   'wdNone' = 0,
   'wdInitialAlef' = 1,
   'wdFinalYaa' = 2,
   'wdBoth' = 3 
 )
storage.mode( WdAraSpellerEnum ) = 'integer'
'WdColorEnum' = c(
   'wdColorAutomatic' = -16777216,
   'wdColorBlack' = 0,
   'wdColorBlue' = 16711680,
   'wdColorTurquoise' = 16776960,
   'wdColorBrightGreen' = 65280,
   'wdColorPink' = 16711935,
   'wdColorRed' = 255,
   'wdColorYellow' = 65535,
   'wdColorWhite' = 16777215,
   'wdColorDarkBlue' = 8388608,
   'wdColorTeal' = 8421376,
   'wdColorGreen' = 32768,
   'wdColorViolet' = 8388736,
   'wdColorDarkRed' = 128,
   'wdColorDarkYellow' = 32896,
   'wdColorBrown' = 13209,
   'wdColorOliveGreen' = 13107,
   'wdColorDarkGreen' = 13056,
   'wdColorDarkTeal' = 6697728,
   'wdColorIndigo' = 10040115,
   'wdColorOrange' = 26367,
   'wdColorBlueGray' = 10053222,
   'wdColorLightOrange' = 39423,
   'wdColorLime' = 52377,
   'wdColorSeaGreen' = 6723891,
   'wdColorAqua' = 13421619,
   'wdColorLightBlue' = 16737843,
   'wdColorGold' = 52479,
   'wdColorSkyBlue' = 16763904,
   'wdColorPlum' = 6697881,
   'wdColorRose' = 13408767,
   'wdColorTan' = 10079487,
   'wdColorLightYellow' = 10092543,
   'wdColorLightGreen' = 13434828,
   'wdColorLightTurquoise' = 16777164,
   'wdColorPaleBlue' = 16764057,
   'wdColorLavender' = 16751052,
   'wdColorGray05' = 15987699,
   'wdColorGray10' = 15132390,
   'wdColorGray125' = 14737632,
   'wdColorGray15' = 14277081,
   'wdColorGray20' = 13421772,
   'wdColorGray25' = 12632256,
   'wdColorGray30' = 11776947,
   'wdColorGray35' = 10921638,
   'wdColorGray375' = 10526880,
   'wdColorGray40' = 10066329,
   'wdColorGray45' = 9211020,
   'wdColorGray50' = 8421504,
   'wdColorGray55' = 7566195,
   'wdColorGray60' = 6710886,
   'wdColorGray625' = 6316128,
   'wdColorGray65' = 5855577,
   'wdColorGray70' = 5000268,
   'wdColorGray75' = 4210752,
   'wdColorGray80' = 3355443,
   'wdColorGray85' = 2500134,
   'wdColorGray875' = 2105376,
   'wdColorGray90' = 1644825,
   'wdColorGray95' = 789516 
 )
storage.mode( WdColorEnum ) = 'integer'
'WdShapePositionEnum' = c(
   'wdShapeTop' = -999999,
   'wdShapeLeft' = -999998,
   'wdShapeBottom' = -999997,
   'wdShapeRight' = -999996,
   'wdShapeCenter' = -999995,
   'wdShapeInside' = -999994,
   'wdShapeOutside' = -999993 
 )
storage.mode( WdShapePositionEnum ) = 'integer'
'WdTablePositionEnum' = c(
   'wdTableTop' = -999999,
   'wdTableLeft' = -999998,
   'wdTableBottom' = -999997,
   'wdTableRight' = -999996,
   'wdTableCenter' = -999995,
   'wdTableInside' = -999994,
   'wdTableOutside' = -999993 
 )
storage.mode( WdTablePositionEnum ) = 'integer'
'WdDefaultListBehaviorEnum' = c(
   'wdWord8ListBehavior' = 0,
   'wdWord9ListBehavior' = 1,
   'wdWord10ListBehavior' = 2 
 )
storage.mode( WdDefaultListBehaviorEnum ) = 'integer'
'WdDefaultTableBehaviorEnum' = c(
   'wdWord8TableBehavior' = 0,
   'wdWord9TableBehavior' = 1 
 )
storage.mode( WdDefaultTableBehaviorEnum ) = 'integer'
'WdAutoFitBehaviorEnum' = c(
   'wdAutoFitFixed' = 0,
   'wdAutoFitContent' = 1,
   'wdAutoFitWindow' = 2 
 )
storage.mode( WdAutoFitBehaviorEnum ) = 'integer'
'WdPreferredWidthTypeEnum' = c(
   'wdPreferredWidthAuto' = 1,
   'wdPreferredWidthPercent' = 2,
   'wdPreferredWidthPoints' = 3 
 )
storage.mode( WdPreferredWidthTypeEnum ) = 'integer'
'WdFarEastLineBreakLanguageIDEnum' = c(
   'wdLineBreakJapanese' = 1041,
   'wdLineBreakKorean' = 1042,
   'wdLineBreakSimplifiedChinese' = 2052,
   'wdLineBreakTraditionalChinese' = 1028 
 )
storage.mode( WdFarEastLineBreakLanguageIDEnum ) = 'integer'
'WdViewTypeOldEnum' = c(
   'wdPageView' = 3,
   'wdOnlineView' = 6 
 )
storage.mode( WdViewTypeOldEnum ) = 'integer'
'WdFramesetTypeEnum' = c(
   'wdFramesetTypeFrameset' = 0,
   'wdFramesetTypeFrame' = 1 
 )
storage.mode( WdFramesetTypeEnum ) = 'integer'
'WdFramesetSizeTypeEnum' = c(
   'wdFramesetSizeTypePercent' = 0,
   'wdFramesetSizeTypeFixed' = 1,
   'wdFramesetSizeTypeRelative' = 2 
 )
storage.mode( WdFramesetSizeTypeEnum ) = 'integer'
'WdFramesetNewFrameLocationEnum' = c(
   'wdFramesetNewFrameAbove' = 0,
   'wdFramesetNewFrameBelow' = 1,
   'wdFramesetNewFrameRight' = 2,
   'wdFramesetNewFrameLeft' = 3 
 )
storage.mode( WdFramesetNewFrameLocationEnum ) = 'integer'
'WdScrollbarTypeEnum' = c(
   'wdScrollbarTypeAuto' = 0,
   'wdScrollbarTypeYes' = 1,
   'wdScrollbarTypeNo' = 2 
 )
storage.mode( WdScrollbarTypeEnum ) = 'integer'
'WdTwoLinesInOneTypeEnum' = c(
   'wdTwoLinesInOneNone' = 0,
   'wdTwoLinesInOneNoBrackets' = 1,
   'wdTwoLinesInOneParentheses' = 2,
   'wdTwoLinesInOneSquareBrackets' = 3,
   'wdTwoLinesInOneAngleBrackets' = 4,
   'wdTwoLinesInOneCurlyBrackets' = 5 
 )
storage.mode( WdTwoLinesInOneTypeEnum ) = 'integer'
'WdHorizontalInVerticalTypeEnum' = c(
   'wdHorizontalInVerticalNone' = 0,
   'wdHorizontalInVerticalFitInLine' = 1,
   'wdHorizontalInVerticalResizeLine' = 2 
 )
storage.mode( WdHorizontalInVerticalTypeEnum ) = 'integer'
'WdHorizontalLineAlignmentEnum' = c(
   'wdHorizontalLineAlignLeft' = 0,
   'wdHorizontalLineAlignCenter' = 1,
   'wdHorizontalLineAlignRight' = 2 
 )
storage.mode( WdHorizontalLineAlignmentEnum ) = 'integer'
'WdHorizontalLineWidthTypeEnum' = c(
   'wdHorizontalLinePercentWidth' = -1,
   'wdHorizontalLineFixedWidth' = -2 
 )
storage.mode( WdHorizontalLineWidthTypeEnum ) = 'integer'
'WdPhoneticGuideAlignmentTypeEnum' = c(
   'wdPhoneticGuideAlignmentCenter' = 0,
   'wdPhoneticGuideAlignmentZeroOneZero' = 1,
   'wdPhoneticGuideAlignmentOneTwoOne' = 2,
   'wdPhoneticGuideAlignmentLeft' = 3,
   'wdPhoneticGuideAlignmentRight' = 4,
   'wdPhoneticGuideAlignmentRightVertical' = 5 
 )
storage.mode( WdPhoneticGuideAlignmentTypeEnum ) = 'integer'
'WdNewDocumentTypeEnum' = c(
   'wdNewBlankDocument' = 0,
   'wdNewWebPage' = 1,
   'wdNewEmailMessage' = 2,
   'wdNewFrameset' = 3,
   'wdNewXMLDocument' = 4 
 )
storage.mode( WdNewDocumentTypeEnum ) = 'integer'
'WdKanaEnum' = c(
   'wdKanaKatakana' = 8,
   'wdKanaHiragana' = 9 
 )
storage.mode( WdKanaEnum ) = 'integer'
'WdCharacterWidthEnum' = c(
   'wdWidthHalfWidth' = 6,
   'wdWidthFullWidth' = 7 
 )
storage.mode( WdCharacterWidthEnum ) = 'integer'
'WdNumberStyleWordBasicBiDiEnum' = c(
   'wdListNumberStyleBidi1' = 49,
   'wdListNumberStyleBidi2' = 50,
   'wdCaptionNumberStyleBidiLetter1' = 49,
   'wdCaptionNumberStyleBidiLetter2' = 50,
   'wdNoteNumberStyleBidiLetter1' = 49,
   'wdNoteNumberStyleBidiLetter2' = 50,
   'wdPageNumberStyleBidiLetter1' = 49,
   'wdPageNumberStyleBidiLetter2' = 50 
 )
storage.mode( WdNumberStyleWordBasicBiDiEnum ) = 'integer'
'WdTCSCConverterDirectionEnum' = c(
   'wdTCSCConverterDirectionSCTC' = 0,
   'wdTCSCConverterDirectionTCSC' = 1,
   'wdTCSCConverterDirectionAuto' = 2 
 )
storage.mode( WdTCSCConverterDirectionEnum ) = 'integer'
'WdDisableFeaturesIntroducedAfterEnum' = c(
   'wd70' = 0,
   'wd70FE' = 1,
   'wd80' = 2 
 )
storage.mode( WdDisableFeaturesIntroducedAfterEnum ) = 'integer'
'WdWrapTypeMergedEnum' = c(
   'wdWrapMergeInline' = 0,
   'wdWrapMergeSquare' = 1,
   'wdWrapMergeTight' = 2,
   'wdWrapMergeBehind' = 3,
   'wdWrapMergeFront' = 4,
   'wdWrapMergeThrough' = 5,
   'wdWrapMergeTopBottom' = 6 
 )
storage.mode( WdWrapTypeMergedEnum ) = 'integer'
'WdRecoveryTypeEnum' = c(
   'wdPasteDefault' = 0,
   'wdSingleCellText' = 5,
   'wdSingleCellTable' = 6,
   'wdListContinueNumbering' = 7,
   'wdListRestartNumbering' = 8,
   'wdTableInsertAsRows' = 11,
   'wdTableAppendTable' = 10,
   'wdTableOriginalFormatting' = 12,
   'wdChartPicture' = 13,
   'wdChart' = 14,
   'wdChartLinked' = 15,
   'wdFormatOriginalFormatting' = 16,
   'wdFormatSurroundingFormattingWithEmphasis' = 20,
   'wdFormatPlainText' = 22,
   'wdTableOverwriteCells' = 23,
   'wdListCombineWithExistingList' = 24,
   'wdListDontMerge' = 25 
 )
storage.mode( WdRecoveryTypeEnum ) = 'integer'
'WdLineEndingTypeEnum' = c(
   'wdCRLF' = 0,
   'wdCROnly' = 1,
   'wdLFOnly' = 2,
   'wdLFCR' = 3,
   'wdLSPS' = 4 
 )
storage.mode( WdLineEndingTypeEnum ) = 'integer'
'WdStyleSheetLinkTypeEnum' = c(
   'wdStyleSheetLinkTypeLinked' = 0,
   'wdStyleSheetLinkTypeImported' = 1 
 )
storage.mode( WdStyleSheetLinkTypeEnum ) = 'integer'
'WdStyleSheetPrecedenceEnum' = c(
   'wdStyleSheetPrecedenceHigher' = -1,
   'wdStyleSheetPrecedenceLower' = -2,
   'wdStyleSheetPrecedenceHighest' = 1,
   'wdStyleSheetPrecedenceLowest' = 0 
 )
storage.mode( WdStyleSheetPrecedenceEnum ) = 'integer'
'WdEmailHTMLFidelityEnum' = c(
   'wdEmailHTMLFidelityLow' = 1,
   'wdEmailHTMLFidelityMedium' = 2,
   'wdEmailHTMLFidelityHigh' = 3 
 )
storage.mode( WdEmailHTMLFidelityEnum ) = 'integer'
'WdMailMergeMailFormatEnum' = c(
   'wdMailFormatPlainText' = 0,
   'wdMailFormatHTML' = 1 
 )
storage.mode( WdMailMergeMailFormatEnum ) = 'integer'
'WdMappedDataFieldsEnum' = c(
   'wdUniqueIdentifier' = 1,
   'wdCourtesyTitle' = 2,
   'wdFirstName' = 3,
   'wdMiddleName' = 4,
   'wdLastName' = 5,
   'wdSuffix' = 6,
   'wdNickname' = 7,
   'wdJobTitle' = 8,
   'wdCompany' = 9,
   'wdAddress1' = 10,
   'wdAddress2' = 11,
   'wdCity' = 12,
   'wdState' = 13,
   'wdPostalCode' = 14,
   'wdCountryRegion' = 15,
   'wdBusinessPhone' = 16,
   'wdBusinessFax' = 17,
   'wdHomePhone' = 18,
   'wdHomeFax' = 19,
   'wdEmailAddress' = 20,
   'wdWebPageURL' = 21,
   'wdSpouseCourtesyTitle' = 22,
   'wdSpouseFirstName' = 23,
   'wdSpouseMiddleName' = 24,
   'wdSpouseLastName' = 25,
   'wdSpouseNickname' = 26,
   'wdRubyFirstName' = 27,
   'wdRubyLastName' = 28,
   'wdAddress3' = 29,
   'wdDepartment' = 30 
 )
storage.mode( WdMappedDataFieldsEnum ) = 'integer'
'WdConditionCodeEnum' = c(
   'wdFirstRow' = 0,
   'wdLastRow' = 1,
   'wdOddRowBanding' = 2,
   'wdEvenRowBanding' = 3,
   'wdFirstColumn' = 4,
   'wdLastColumn' = 5,
   'wdOddColumnBanding' = 6,
   'wdEvenColumnBanding' = 7,
   'wdNECell' = 8,
   'wdNWCell' = 9,
   'wdSECell' = 10,
   'wdSWCell' = 11 
 )
storage.mode( WdConditionCodeEnum ) = 'integer'
'WdCompareTargetEnum' = c(
   'wdCompareTargetSelected' = 0,
   'wdCompareTargetCurrent' = 1,
   'wdCompareTargetNew' = 2 
 )
storage.mode( WdCompareTargetEnum ) = 'integer'
'WdMergeTargetEnum' = c(
   'wdMergeTargetSelected' = 0,
   'wdMergeTargetCurrent' = 1,
   'wdMergeTargetNew' = 2 
 )
storage.mode( WdMergeTargetEnum ) = 'integer'
'WdUseFormattingFromEnum' = c(
   'wdFormattingFromCurrent' = 0,
   'wdFormattingFromSelected' = 1,
   'wdFormattingFromPrompt' = 2 
 )
storage.mode( WdUseFormattingFromEnum ) = 'integer'
'WdRevisionsViewEnum' = c(
   'wdRevisionsViewFinal' = 0,
   'wdRevisionsViewOriginal' = 1 
 )
storage.mode( WdRevisionsViewEnum ) = 'integer'
'WdRevisionsModeEnum' = c(
   'wdBalloonRevisions' = 0,
   'wdInLineRevisions' = 1,
   'wdMixedRevisions' = 2 
 )
storage.mode( WdRevisionsModeEnum ) = 'integer'
'WdRevisionsBalloonWidthTypeEnum' = c(
   'wdBalloonWidthPercent' = 0,
   'wdBalloonWidthPoints' = 1 
 )
storage.mode( WdRevisionsBalloonWidthTypeEnum ) = 'integer'
'WdRevisionsBalloonPrintOrientationEnum' = c(
   'wdBalloonPrintOrientationAuto' = 0,
   'wdBalloonPrintOrientationPreserve' = 1,
   'wdBalloonPrintOrientationForceLandscape' = 2 
 )
storage.mode( WdRevisionsBalloonPrintOrientationEnum ) = 'integer'
'WdRevisionsBalloonMarginEnum' = c(
   'wdLeftMargin' = 0,
   'wdRightMargin' = 1 
 )
storage.mode( WdRevisionsBalloonMarginEnum ) = 'integer'
'WdTaskPanesEnum' = c(
   'wdTaskPaneFormatting' = 0,
   'wdTaskPaneRevealFormatting' = 1,
   'wdTaskPaneMailMerge' = 2,
   'wdTaskPaneTranslate' = 3,
   'wdTaskPaneSearch' = 4,
   'wdTaskPaneXMLStructure' = 5,
   'wdTaskPaneDocumentProtection' = 6,
   'wdTaskPaneDocumentActions' = 7,
   'wdTaskPaneSharedWorkspace' = 8,
   'wdTaskPaneHelp' = 9,
   'wdTaskPaneResearch' = 10,
   'wdTaskPaneFaxService' = 11,
   'wdTaskPaneXMLDocument' = 12,
   'wdTaskPaneDocumentUpdates' = 13 
 )
storage.mode( WdTaskPanesEnum ) = 'integer'
'WdShowFilterEnum' = c(
   'wdShowFilterStylesAvailable' = 0,
   'wdShowFilterStylesInUse' = 1,
   'wdShowFilterStylesAll' = 2,
   'wdShowFilterFormattingInUse' = 3,
   'wdShowFilterFormattingAvailable' = 4 
 )
storage.mode( WdShowFilterEnum ) = 'integer'
'WdMergeSubTypeEnum' = c(
   'wdMergeSubTypeOther' = 0,
   'wdMergeSubTypeAccess' = 1,
   'wdMergeSubTypeOAL' = 2,
   'wdMergeSubTypeOLEDBWord' = 3,
   'wdMergeSubTypeWorks' = 4,
   'wdMergeSubTypeOLEDBText' = 5,
   'wdMergeSubTypeOutlook' = 6,
   'wdMergeSubTypeWord' = 7,
   'wdMergeSubTypeWord2000' = 8 
 )
storage.mode( WdMergeSubTypeEnum ) = 'integer'
'WdDocumentDirectionEnum' = c(
   'wdLeftToRight' = 0,
   'wdRightToLeft' = 1 
 )
storage.mode( WdDocumentDirectionEnum ) = 'integer'
'WdLanguageID2000Enum' = c(
   'wdChineseHongKong' = 3076,
   'wdChineseMacao' = 5124,
   'wdEnglishTrinidad' = 11273 
 )
storage.mode( WdLanguageID2000Enum ) = 'integer'
'WdRectangleTypeEnum' = c(
   'wdTextRectangle' = 0,
   'wdShapeRectangle' = 1,
   'wdMarkupRectangle' = 2,
   'wdMarkupRectangleButton' = 3,
   'wdPageBorderRectangle' = 4,
   'wdLineBetweenColumnRectangle' = 5,
   'wdSelection' = 6,
   'wdSystem' = 7 
 )
storage.mode( WdRectangleTypeEnum ) = 'integer'
'WdLineTypeEnum' = c(
   'wdTextLine' = 0,
   'wdTableRow' = 1 
 )
storage.mode( WdLineTypeEnum ) = 'integer'
'WdXMLNodeTypeEnum' = c(
   'wdXMLNodeElement' = 1,
   'wdXMLNodeAttribute' = 2 
 )
storage.mode( WdXMLNodeTypeEnum ) = 'integer'
'WdXMLSelectionChangeReasonEnum' = c(
   'wdXMLSelectionChangeReasonMove' = 0,
   'wdXMLSelectionChangeReasonInsert' = 1,
   'wdXMLSelectionChangeReasonDelete' = 2 
 )
storage.mode( WdXMLSelectionChangeReasonEnum ) = 'integer'
'WdXMLNodeLevelEnum' = c(
   'wdXMLNodeLevelInline' = 0,
   'wdXMLNodeLevelParagraph' = 1,
   'wdXMLNodeLevelRow' = 2,
   'wdXMLNodeLevelCell' = 3 
 )
storage.mode( WdXMLNodeLevelEnum ) = 'integer'
'WdSmartTagControlTypeEnum' = c(
   'wdControlSmartTag' = 1,
   'wdControlLink' = 2,
   'wdControlHelp' = 3,
   'wdControlHelpURL' = 4,
   'wdControlSeparator' = 5,
   'wdControlButton' = 6,
   'wdControlLabel' = 7,
   'wdControlImage' = 8,
   'wdControlCheckbox' = 9,
   'wdControlTextbox' = 10,
   'wdControlListbox' = 11,
   'wdControlCombo' = 12,
   'wdControlActiveX' = 13,
   'wdControlDocumentFragment' = 14,
   'wdControlDocumentFragmentURL' = 15,
   'wdControlRadioGroup' = 16 
 )
storage.mode( WdSmartTagControlTypeEnum ) = 'integer'
'WdEditorTypeEnum' = c(
   'wdEditorEveryone' = -1,
   'wdEditorOwners' = -4,
   'wdEditorEditors' = -5,
   'wdEditorCurrent' = -6 
 )
storage.mode( WdEditorTypeEnum ) = 'integer'
'WdXMLValidationStatusEnum' = c(
   'wdXMLValidationStatusOK' = 0,
   'wdXMLValidationStatusCustom' = -1072898048 
 )
storage.mode( WdXMLValidationStatusEnum ) = 'integer'
'COM._Application.GetProperty'  = list('Application' = function(x) {
		 ans = .COM(x, 'Application', .dispatch = as.integer(2))
		 as(ans, 'Application')
},
'Creator' = function(x) {
		 ans = .COM(x, 'Creator', .dispatch = as.integer(2))
	ans
},
'Parent' = function(x) {
		 ans = .COM(x, 'Parent', .dispatch = as.integer(2))
		 ans
},
'Name' = function(x) {
		 ans = .COM(x, 'Name', .dispatch = as.integer(2))
	ans
},
'Documents' = function(x) {
		 ans = .COM(x, 'Documents', .dispatch = as.integer(2))
		 as(ans, 'Documents')
},
'Windows' = function(x) {
		 ans = .COM(x, 'Windows', .dispatch = as.integer(2))
		 as(ans, 'Windows')
},
'ActiveDocument' = function(x) {
		 ans = .COM(x, 'ActiveDocument', .dispatch = as.integer(2))
		 as(ans, 'Document')
},
'ActiveWindow' = function(x) {
		 ans = .COM(x, 'ActiveWindow', .dispatch = as.integer(2))
		 as(ans, 'Window')
},
'Selection' = function(x) {
		 ans = .COM(x, 'Selection', .dispatch = as.integer(2))
		 as(ans, 'Selection')
},
'WordBasic' = function(x) {
		 ans = .COM(x, 'WordBasic', .dispatch = as.integer(2))
		 ans
},
'RecentFiles' = function(x) {
		 ans = .COM(x, 'RecentFiles', .dispatch = as.integer(2))
		 as(ans, 'RecentFiles')
},
'NormalTemplate' = function(x) {
		 ans = .COM(x, 'NormalTemplate', .dispatch = as.integer(2))
		 as(ans, 'Template')
},
'System' = function(x) {
		 ans = .COM(x, 'System', .dispatch = as.integer(2))
		 as(ans, 'System')
},
'AutoCorrect' = function(x) {
		 ans = .COM(x, 'AutoCorrect', .dispatch = as.integer(2))
		 as(ans, 'AutoCorrect')
},
'FontNames' = function(x) {
		 ans = .COM(x, 'FontNames', .dispatch = as.integer(2))
		 as(ans, 'FontNames')
},
'LandscapeFontNames' = function(x) {
		 ans = .COM(x, 'LandscapeFontNames', .dispatch = as.integer(2))
		 as(ans, 'FontNames')
},
'PortraitFontNames' = function(x) {
		 ans = .COM(x, 'PortraitFontNames', .dispatch = as.integer(2))
		 as(ans, 'FontNames')
},
'Languages' = function(x) {
		 ans = .COM(x, 'Languages', .dispatch = as.integer(2))
		 as(ans, 'Languages')
},
'Assistant' = function(x) {
		 ans = .COM(x, 'Assistant', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'Browser' = function(x) {
		 ans = .COM(x, 'Browser', .dispatch = as.integer(2))
		 as(ans, 'Browser')
},
'FileConverters' = function(x) {
		 ans = .COM(x, 'FileConverters', .dispatch = as.integer(2))
		 as(ans, 'FileConverters')
},
'MailingLabel' = function(x) {
		 ans = .COM(x, 'MailingLabel', .dispatch = as.integer(2))
		 as(ans, 'MailingLabel')
},
'Dialogs' = function(x) {
		 ans = .COM(x, 'Dialogs', .dispatch = as.integer(2))
		 as(ans, 'Dialogs')
},
'CaptionLabels' = function(x) {
		 ans = .COM(x, 'CaptionLabels', .dispatch = as.integer(2))
		 as(ans, 'CaptionLabels')
},
'AutoCaptions' = function(x) {
		 ans = .COM(x, 'AutoCaptions', .dispatch = as.integer(2))
		 as(ans, 'AutoCaptions')
},
'AddIns' = function(x) {
		 ans = .COM(x, 'AddIns', .dispatch = as.integer(2))
		 as(ans, 'AddIns')
},
'Visible' = function(x) {
		 ans = .COM(x, 'Visible', .dispatch = as.integer(2))
	ans
},
'Version' = function(x) {
		 ans = .COM(x, 'Version', .dispatch = as.integer(2))
	ans
},
'ScreenUpdating' = function(x) {
		 ans = .COM(x, 'ScreenUpdating', .dispatch = as.integer(2))
	ans
},
'PrintPreview' = function(x) {
		 ans = .COM(x, 'PrintPreview', .dispatch = as.integer(2))
	ans
},
'Tasks' = function(x) {
		 ans = .COM(x, 'Tasks', .dispatch = as.integer(2))
		 as(ans, 'Tasks')
},
'DisplayStatusBar' = function(x) {
		 ans = .COM(x, 'DisplayStatusBar', .dispatch = as.integer(2))
	ans
},
'SpecialMode' = function(x) {
		 ans = .COM(x, 'SpecialMode', .dispatch = as.integer(2))
	ans
},
'UsableWidth' = function(x) {
		 ans = .COM(x, 'UsableWidth', .dispatch = as.integer(2))
	ans
},
'UsableHeight' = function(x) {
		 ans = .COM(x, 'UsableHeight', .dispatch = as.integer(2))
	ans
},
'MathCoprocessorAvailable' = function(x) {
		 ans = .COM(x, 'MathCoprocessorAvailable', .dispatch = as.integer(2))
	ans
},
'MouseAvailable' = function(x) {
		 ans = .COM(x, 'MouseAvailable', .dispatch = as.integer(2))
	ans
},
'Build' = function(x) {
		 ans = .COM(x, 'Build', .dispatch = as.integer(2))
	ans
},
'CapsLock' = function(x) {
		 ans = .COM(x, 'CapsLock', .dispatch = as.integer(2))
	ans
},
'NumLock' = function(x) {
		 ans = .COM(x, 'NumLock', .dispatch = as.integer(2))
	ans
},
'UserName' = function(x) {
		 ans = .COM(x, 'UserName', .dispatch = as.integer(2))
	ans
},
'UserInitials' = function(x) {
		 ans = .COM(x, 'UserInitials', .dispatch = as.integer(2))
	ans
},
'UserAddress' = function(x) {
		 ans = .COM(x, 'UserAddress', .dispatch = as.integer(2))
	ans
},
'MacroContainer' = function(x) {
		 ans = .COM(x, 'MacroContainer', .dispatch = as.integer(2))
		 ans
},
'DisplayRecentFiles' = function(x) {
		 ans = .COM(x, 'DisplayRecentFiles', .dispatch = as.integer(2))
	ans
},
'CommandBars' = function(x) {
		 ans = .COM(x, 'CommandBars', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'VBE' = function(x) {
		 ans = .COM(x, 'VBE', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'DefaultSaveFormat' = function(x) {
		 ans = .COM(x, 'DefaultSaveFormat', .dispatch = as.integer(2))
	ans
},
'ListGalleries' = function(x) {
		 ans = .COM(x, 'ListGalleries', .dispatch = as.integer(2))
		 as(ans, 'ListGalleries')
},
'ActivePrinter' = function(x) {
		 ans = .COM(x, 'ActivePrinter', .dispatch = as.integer(2))
	ans
},
'Templates' = function(x) {
		 ans = .COM(x, 'Templates', .dispatch = as.integer(2))
		 as(ans, 'Templates')
},
'CustomizationContext' = function(x) {
		 ans = .COM(x, 'CustomizationContext', .dispatch = as.integer(2))
		 ans
},
'KeyBindings' = function(x) {
		 ans = .COM(x, 'KeyBindings', .dispatch = as.integer(2))
		 as(ans, 'KeyBindings')
},
'Caption' = function(x) {
		 ans = .COM(x, 'Caption', .dispatch = as.integer(2))
	ans
},
'Path' = function(x) {
		 ans = .COM(x, 'Path', .dispatch = as.integer(2))
	ans
},
'DisplayScrollBars' = function(x) {
		 ans = .COM(x, 'DisplayScrollBars', .dispatch = as.integer(2))
	ans
},
'StartupPath' = function(x) {
		 ans = .COM(x, 'StartupPath', .dispatch = as.integer(2))
	ans
},
'BackgroundSavingStatus' = function(x) {
		 ans = .COM(x, 'BackgroundSavingStatus', .dispatch = as.integer(2))
	ans
},
'BackgroundPrintingStatus' = function(x) {
		 ans = .COM(x, 'BackgroundPrintingStatus', .dispatch = as.integer(2))
	ans
},
'Left' = function(x) {
		 ans = .COM(x, 'Left', .dispatch = as.integer(2))
	ans
},
'Top' = function(x) {
		 ans = .COM(x, 'Top', .dispatch = as.integer(2))
	ans
},
'Width' = function(x) {
		 ans = .COM(x, 'Width', .dispatch = as.integer(2))
	ans
},
'Height' = function(x) {
		 ans = .COM(x, 'Height', .dispatch = as.integer(2))
	ans
},
'WindowState' = function(x) {
		 ans = .COM(x, 'WindowState', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'DisplayAutoCompleteTips' = function(x) {
		 ans = .COM(x, 'DisplayAutoCompleteTips', .dispatch = as.integer(2))
	ans
},
'Options' = function(x) {
		 ans = .COM(x, 'Options', .dispatch = as.integer(2))
		 as(ans, 'Options')
},
'DisplayAlerts' = function(x) {
		 ans = .COM(x, 'DisplayAlerts', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'CustomDictionaries' = function(x) {
		 ans = .COM(x, 'CustomDictionaries', .dispatch = as.integer(2))
		 as(ans, 'Dictionaries')
},
'PathSeparator' = function(x) {
		 ans = .COM(x, 'PathSeparator', .dispatch = as.integer(2))
	ans
},
'MAPIAvailable' = function(x) {
		 ans = .COM(x, 'MAPIAvailable', .dispatch = as.integer(2))
	ans
},
'DisplayScreenTips' = function(x) {
		 ans = .COM(x, 'DisplayScreenTips', .dispatch = as.integer(2))
	ans
},
'EnableCancelKey' = function(x) {
		 ans = .COM(x, 'EnableCancelKey', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'UserControl' = function(x) {
		 ans = .COM(x, 'UserControl', .dispatch = as.integer(2))
	ans
},
'FileSearch' = function(x) {
		 ans = .COM(x, 'FileSearch', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'MailSystem' = function(x) {
		 ans = .COM(x, 'MailSystem', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'DefaultTableSeparator' = function(x) {
		 ans = .COM(x, 'DefaultTableSeparator', .dispatch = as.integer(2))
	ans
},
'ShowVisualBasicEditor' = function(x) {
		 ans = .COM(x, 'ShowVisualBasicEditor', .dispatch = as.integer(2))
	ans
},
'BrowseExtraFileTypes' = function(x) {
		 ans = .COM(x, 'BrowseExtraFileTypes', .dispatch = as.integer(2))
	ans
},
'HangulHanjaDictionaries' = function(x) {
		 ans = .COM(x, 'HangulHanjaDictionaries', .dispatch = as.integer(2))
		 as(ans, 'HangulHanjaConversionDictionaries')
},
'MailMessage' = function(x) {
		 ans = .COM(x, 'MailMessage', .dispatch = as.integer(2))
		 as(ans, 'MailMessage')
},
'FocusInMailHeader' = function(x) {
		 ans = .COM(x, 'FocusInMailHeader', .dispatch = as.integer(2))
	ans
},
'EmailOptions' = function(x) {
		 ans = .COM(x, 'EmailOptions', .dispatch = as.integer(2))
		 as(ans, 'EmailOptions')
},
'Language' = function(x) {
		 ans = .COM(x, 'Language', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'COMAddIns' = function(x) {
		 ans = .COM(x, 'COMAddIns', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'CheckLanguage' = function(x) {
		 ans = .COM(x, 'CheckLanguage', .dispatch = as.integer(2))
	ans
},
'LanguageSettings' = function(x) {
		 ans = .COM(x, 'LanguageSettings', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'Dummy1' = function(x) {
		 ans = .COM(x, 'Dummy1', .dispatch = as.integer(2))
	ans
},
'AnswerWizard' = function(x) {
		 ans = .COM(x, 'AnswerWizard', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'FeatureInstall' = function(x) {
		 ans = .COM(x, 'FeatureInstall', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'AutomationSecurity' = function(x) {
		 ans = .COM(x, 'AutomationSecurity', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'EmailTemplate' = function(x) {
		 ans = .COM(x, 'EmailTemplate', .dispatch = as.integer(2))
	ans
},
'ShowWindowsInTaskbar' = function(x) {
		 ans = .COM(x, 'ShowWindowsInTaskbar', .dispatch = as.integer(2))
	ans
},
'NewDocument' = function(x) {
		 ans = .COM(x, 'NewDocument', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'ShowStartupDialog' = function(x) {
		 ans = .COM(x, 'ShowStartupDialog', .dispatch = as.integer(2))
	ans
},
'AutoCorrectEmail' = function(x) {
		 ans = .COM(x, 'AutoCorrectEmail', .dispatch = as.integer(2))
		 as(ans, 'AutoCorrect')
},
'TaskPanes' = function(x) {
		 ans = .COM(x, 'TaskPanes', .dispatch = as.integer(2))
		 as(ans, 'TaskPanes')
},
'DefaultLegalBlackline' = function(x) {
		 ans = .COM(x, 'DefaultLegalBlackline', .dispatch = as.integer(2))
	ans
},
'SmartTagRecognizers' = function(x) {
		 ans = .COM(x, 'SmartTagRecognizers', .dispatch = as.integer(2))
		 as(ans, 'SmartTagRecognizers')
},
'SmartTagTypes' = function(x) {
		 ans = .COM(x, 'SmartTagTypes', .dispatch = as.integer(2))
		 as(ans, 'SmartTagTypes')
},
'XMLNamespaces' = function(x) {
		 ans = .COM(x, 'XMLNamespaces', .dispatch = as.integer(2))
		 as(ans, 'XMLNamespaces')
},
'ArbitraryXMLSupportAvailable' = function(x) {
		 ans = .COM(x, 'ArbitraryXMLSupportAvailable', .dispatch = as.integer(2))
	ans
} )
'COM._Application.SetProperty'  = list('Visible' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'Visible', value, .dispatch = as.integer(4))
},
'ScreenUpdating' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ScreenUpdating', value, .dispatch = as.integer(4))
},
'PrintPreview' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'PrintPreview', value, .dispatch = as.integer(4))
},
'DisplayStatusBar' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayStatusBar', value, .dispatch = as.integer(4))
},
'UserName' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'UserName', value, .dispatch = as.integer(4))
},
'UserInitials' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'UserInitials', value, .dispatch = as.integer(4))
},
'UserAddress' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'UserAddress', value, .dispatch = as.integer(4))
},
'DisplayRecentFiles' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayRecentFiles', value, .dispatch = as.integer(4))
},
'DefaultSaveFormat' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'DefaultSaveFormat', value, .dispatch = as.integer(4))
},
'ActivePrinter' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'ActivePrinter', value, .dispatch = as.integer(4))
},
'CustomizationContext' = function(x, value) {
	
	.COM(x, 'CustomizationContext', value, .dispatch = as.integer(4))
},
'Caption' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'Caption', value, .dispatch = as.integer(4))
},
'DisplayScrollBars' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayScrollBars', value, .dispatch = as.integer(4))
},
'StartupPath' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'StartupPath', value, .dispatch = as.integer(4))
},
'Left' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'Left', value, .dispatch = as.integer(4))
},
'Top' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'Top', value, .dispatch = as.integer(4))
},
'Width' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'Width', value, .dispatch = as.integer(4))
},
'Height' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'Height', value, .dispatch = as.integer(4))
},
'WindowState' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'WindowState', value, .dispatch = as.integer(4))
},
'DisplayAutoCompleteTips' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayAutoCompleteTips', value, .dispatch = as.integer(4))
},
'DisplayAlerts' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'DisplayAlerts', value, .dispatch = as.integer(4))
},
'StatusBar' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'StatusBar', value, .dispatch = as.integer(4))
},
'DisplayScreenTips' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DisplayScreenTips', value, .dispatch = as.integer(4))
},
'EnableCancelKey' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'EnableCancelKey', value, .dispatch = as.integer(4))
},
'DefaultTableSeparator' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'DefaultTableSeparator', value, .dispatch = as.integer(4))
},
'ShowVisualBasicEditor' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ShowVisualBasicEditor', value, .dispatch = as.integer(4))
},
'BrowseExtraFileTypes' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'BrowseExtraFileTypes', value, .dispatch = as.integer(4))
},
'CheckLanguage' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'CheckLanguage', value, .dispatch = as.integer(4))
},
'FeatureInstall' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'FeatureInstall', value, .dispatch = as.integer(4))
},
'AutomationSecurity' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'AutomationSecurity', value, .dispatch = as.integer(4))
},
'EmailTemplate' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'EmailTemplate', value, .dispatch = as.integer(4))
},
'ShowWindowsInTaskbar' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ShowWindowsInTaskbar', value, .dispatch = as.integer(4))
},
'ShowStartupDialog' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ShowStartupDialog', value, .dispatch = as.integer(4))
},
'DefaultLegalBlackline' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DefaultLegalBlackline', value, .dispatch = as.integer(4))
} )
'COM._Application.Methods'  = list('Quit' = function( SaveChanges = NA, OriginalFormat = NA, RouteDocument = NA ,  .x){
	
	
	
	ans = .COM(.x, 'Quit', SaveChanges, OriginalFormat, RouteDocument, .dispatch = as.integer(1), .ids =1105, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ScreenRefresh' = function(   .x){
	ans = .COM(.x, 'ScreenRefresh', .dispatch = as.integer(1), .ids =301, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PrintOutOld' = function( Background = NA, Append = NA, Range = NA, OutputFileName = NA, From = NA, To = NA, Item = NA, Copies = NA, Pages = NA, PageType = NA, PrintToFile = NA, Collate = NA, FileName = NA, ActivePrinterMacGX = NA, ManualDuplexPrint = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'PrintOutOld', Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, .dispatch = as.integer(1), .ids =302, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'LookupNameProperties' = function( Name ,  .x){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	Name = as(Name, 'character')
	ans = .COM(.x, 'LookupNameProperties', Name, .dispatch = as.integer(1), .ids =303, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SubstituteFont' = function( UnavailableFont, SubstituteFont ,  .x){
	if( missing( UnavailableFont )||missing( SubstituteFont ) ) {
	stop('You must specify a value for the argument(s)  UnavailableFont, SubstituteFont ')
	}
	UnavailableFont = as(UnavailableFont, 'character')
	SubstituteFont = as(SubstituteFont, 'character')
	ans = .COM(.x, 'SubstituteFont', UnavailableFont, SubstituteFont, .dispatch = as.integer(1), .ids =304, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Repeat' = function( Times = NA ,  .x){
	
	ans = .COM(.x, 'Repeat', Times, .dispatch = as.integer(1), .ids =305, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'DDEExecute' = function( Channel, Command ,  .x){
	if( missing( Channel )||missing( Command ) ) {
	stop('You must specify a value for the argument(s)  Channel, Command ')
	}
	Channel = as(Channel, 'integer')
	Command = as(Command, 'character')
	ans = .COM(.x, 'DDEExecute', Channel, Command, .dispatch = as.integer(1), .ids =310, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DDEInitiate' = function( App, Topic ,  .x){
	if( missing( App )||missing( Topic ) ) {
	stop('You must specify a value for the argument(s)  App, Topic ')
	}
	App = as(App, 'character')
	Topic = as(Topic, 'character')
	ans = .COM(.x, 'DDEInitiate', App, Topic, .dispatch = as.integer(1), .ids =311, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'DDEPoke' = function( Channel, Item, Data ,  .x){
	if( missing( Channel )||missing( Item )||missing( Data ) ) {
	stop('You must specify a value for the argument(s)  Channel, Item, Data ')
	}
	Channel = as(Channel, 'integer')
	Item = as(Item, 'character')
	Data = as(Data, 'character')
	ans = .COM(.x, 'DDEPoke', Channel, Item, Data, .dispatch = as.integer(1), .ids =312, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DDERequest' = function( Channel, Item ,  .x){
	if( missing( Channel )||missing( Item ) ) {
	stop('You must specify a value for the argument(s)  Channel, Item ')
	}
	Channel = as(Channel, 'integer')
	Item = as(Item, 'character')
	ans = .COM(.x, 'DDERequest', Channel, Item, .dispatch = as.integer(1), .ids =313, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'DDETerminate' = function( Channel ,  .x){
	if( missing( Channel ) ) {
	stop('You must specify a value for the argument(s)  Channel ')
	}
	Channel = as(Channel, 'integer')
	ans = .COM(.x, 'DDETerminate', Channel, .dispatch = as.integer(1), .ids =314, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DDETerminateAll' = function(   .x){
	ans = .COM(.x, 'DDETerminateAll', .dispatch = as.integer(1), .ids =315, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'BuildKeyCode' = function( Arg1, Arg2 = NA, Arg3 = NA, Arg4 = NA ,  .x){
	if( missing( Arg1 ) ) {
	stop('You must specify a value for the argument(s)  Arg1 ')
	}
	Arg1 = as(Arg1, 'WdMailSystem')
	
	
	
	ans = .COM(.x, 'BuildKeyCode', Arg1, Arg2, Arg3, Arg4, .dispatch = as.integer(1), .ids =316, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'KeyString' = function( KeyCode, KeyCode2 = NA ,  .x){
	if( missing( KeyCode ) ) {
	stop('You must specify a value for the argument(s)  KeyCode ')
	}
	KeyCode = as(KeyCode, 'integer')
	
	ans = .COM(.x, 'KeyString', KeyCode, KeyCode2, .dispatch = as.integer(1), .ids =317, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'OrganizerCopy' = function( Source, Destination, Name, Object ,  .x){
	if( missing( Source )||missing( Destination )||missing( Name )||missing( Object ) ) {
	stop('You must specify a value for the argument(s)  Source, Destination, Name, Object ')
	}
	Source = as(Source, 'character')
	Destination = as(Destination, 'character')
	Name = as(Name, 'character')
	Object = as(Object, 'WdMailSystem')
	ans = .COM(.x, 'OrganizerCopy', Source, Destination, Name, Object, .dispatch = as.integer(1), .ids =318, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'OrganizerDelete' = function( Source, Name, Object ,  .x){
	if( missing( Source )||missing( Name )||missing( Object ) ) {
	stop('You must specify a value for the argument(s)  Source, Name, Object ')
	}
	Source = as(Source, 'character')
	Name = as(Name, 'character')
	Object = as(Object, 'WdMailSystem')
	ans = .COM(.x, 'OrganizerDelete', Source, Name, Object, .dispatch = as.integer(1), .ids =319, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'OrganizerRename' = function( Source, Name, NewName, Object ,  .x){
	if( missing( Source )||missing( Name )||missing( NewName )||missing( Object ) ) {
	stop('You must specify a value for the argument(s)  Source, Name, NewName, Object ')
	}
	Source = as(Source, 'character')
	Name = as(Name, 'character')
	NewName = as(NewName, 'character')
	Object = as(Object, 'WdMailSystem')
	ans = .COM(.x, 'OrganizerRename', Source, Name, NewName, Object, .dispatch = as.integer(1), .ids =320, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AddAddress' = function( TagID, Value ,  .x){
	if( missing( TagID )||missing( Value ) ) {
	stop('You must specify a value for the argument(s)  TagID, Value ')
	}
	
	
	ans = .COM(.x, 'AddAddress', TagID, Value, .dispatch = as.integer(1), .ids =321, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GetAddress' = function( Name = NA, AddressProperties = NA, UseAutoText = NA, DisplaySelectDialog = NA, SelectDialog = NA, CheckNamesDialog = NA, RecentAddressesChoice = NA, UpdateRecentAddresses = NA ,  .x){
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'GetAddress', Name, AddressProperties, UseAutoText, DisplaySelectDialog, SelectDialog, CheckNamesDialog, RecentAddressesChoice, UpdateRecentAddresses, .dispatch = as.integer(1), .ids =322, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'CheckGrammar' = function( String ,  .x){
	if( missing( String ) ) {
	stop('You must specify a value for the argument(s)  String ')
	}
	String = as(String, 'character')
	ans = .COM(.x, 'CheckGrammar', String, .dispatch = as.integer(1), .ids =323, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'CheckSpelling' = function( Word, CustomDictionary = NA, IgnoreUppercase = NA, MainDictionary = NA, CustomDictionary2 = NA, CustomDictionary3 = NA, CustomDictionary4 = NA, CustomDictionary5 = NA, CustomDictionary6 = NA, CustomDictionary7 = NA, CustomDictionary8 = NA, CustomDictionary9 = NA, CustomDictionary10 = NA ,  .x){
	if( missing( Word ) ) {
	stop('You must specify a value for the argument(s)  Word ')
	}
	Word = as(Word, 'character')
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'CheckSpelling', Word, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10, .dispatch = as.integer(1), .ids =324, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'ResetIgnoreAll' = function(   .x){
	ans = .COM(.x, 'ResetIgnoreAll', .dispatch = as.integer(1), .ids =326, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GetSpellingSuggestions' = function( Word, CustomDictionary = NA, IgnoreUppercase = NA, MainDictionary = NA, SuggestionMode = NA, CustomDictionary2 = NA, CustomDictionary3 = NA, CustomDictionary4 = NA, CustomDictionary5 = NA, CustomDictionary6 = NA, CustomDictionary7 = NA, CustomDictionary8 = NA, CustomDictionary9 = NA, CustomDictionary10 = NA ,  .x){
	if( missing( Word ) ) {
	stop('You must specify a value for the argument(s)  Word ')
	}
	Word = as(Word, 'character')
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'GetSpellingSuggestions', Word, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10, .dispatch = as.integer(1), .ids =327, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'SpellingSuggestions')
},
'GoBack' = function(   .x){
	ans = .COM(.x, 'GoBack', .dispatch = as.integer(1), .ids =328, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Help' = function( HelpType ,  .x){
	if( missing( HelpType ) ) {
	stop('You must specify a value for the argument(s)  HelpType ')
	}
	
	ans = .COM(.x, 'Help', HelpType, .dispatch = as.integer(1), .ids =329, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AutomaticChange' = function(   .x){
	ans = .COM(.x, 'AutomaticChange', .dispatch = as.integer(1), .ids =330, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ShowMe' = function(   .x){
	ans = .COM(.x, 'ShowMe', .dispatch = as.integer(1), .ids =331, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'HelpTool' = function(   .x){
	ans = .COM(.x, 'HelpTool', .dispatch = as.integer(1), .ids =332, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'NewWindow' = function(   .x){
	ans = .COM(.x, 'NewWindow', .dispatch = as.integer(1), .ids =345, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Window')
},
'ListCommands' = function( ListAllCommands ,  .x){
	if( missing( ListAllCommands ) ) {
	stop('You must specify a value for the argument(s)  ListAllCommands ')
	}
	ListAllCommands = as(ListAllCommands, 'logical')
	ans = .COM(.x, 'ListCommands', ListAllCommands, .dispatch = as.integer(1), .ids =346, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ShowClipboard' = function(   .x){
	ans = .COM(.x, 'ShowClipboard', .dispatch = as.integer(1), .ids =349, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'OnTime' = function( When, Name, Tolerance = NA ,  .x){
	if( missing( When )||missing( Name ) ) {
	stop('You must specify a value for the argument(s)  When, Name ')
	}
	
	Name = as(Name, 'character')
	
	ans = .COM(.x, 'OnTime', When, Name, Tolerance, .dispatch = as.integer(1), .ids =350, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'NextLetter' = function(   .x){
	ans = .COM(.x, 'NextLetter', .dispatch = as.integer(1), .ids =351, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'MountVolume' = function( Zone, Server, Volume, User = NA, UserPassword = NA, VolumePassword = NA ,  .x){
	if( missing( Zone )||missing( Server )||missing( Volume ) ) {
	stop('You must specify a value for the argument(s)  Zone, Server, Volume ')
	}
	Zone = as(Zone, 'character')
	Server = as(Server, 'character')
	Volume = as(Volume, 'character')
	
	
	
	ans = .COM(.x, 'MountVolume', Zone, Server, Volume, User, UserPassword, VolumePassword, .dispatch = as.integer(1), .ids =353, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CleanString' = function( String ,  .x){
	if( missing( String ) ) {
	stop('You must specify a value for the argument(s)  String ')
	}
	String = as(String, 'character')
	ans = .COM(.x, 'CleanString', String, .dispatch = as.integer(1), .ids =354, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'SendFax' = function(   .x){
	ans = .COM(.x, 'SendFax', .dispatch = as.integer(1), .ids =356, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ChangeFileOpenDirectory' = function( Path ,  .x){
	if( missing( Path ) ) {
	stop('You must specify a value for the argument(s)  Path ')
	}
	Path = as(Path, 'character')
	ans = .COM(.x, 'ChangeFileOpenDirectory', Path, .dispatch = as.integer(1), .ids =357, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RunOld' = function( MacroName ,  .x){
	if( missing( MacroName ) ) {
	stop('You must specify a value for the argument(s)  MacroName ')
	}
	MacroName = as(MacroName, 'character')
	ans = .COM(.x, 'RunOld', MacroName, .dispatch = as.integer(1), .ids =358, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GoForward' = function(   .x){
	ans = .COM(.x, 'GoForward', .dispatch = as.integer(1), .ids =359, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Move' = function( Left, Top ,  .x){
	if( missing( Left )||missing( Top ) ) {
	stop('You must specify a value for the argument(s)  Left, Top ')
	}
	Left = as(Left, 'integer')
	Top = as(Top, 'integer')
	ans = .COM(.x, 'Move', Left, Top, .dispatch = as.integer(1), .ids =360, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Resize' = function( Width, Height ,  .x){
	if( missing( Width )||missing( Height ) ) {
	stop('You must specify a value for the argument(s)  Width, Height ')
	}
	Width = as(Width, 'integer')
	Height = as(Height, 'integer')
	ans = .COM(.x, 'Resize', Width, Height, .dispatch = as.integer(1), .ids =361, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InchesToPoints' = function( Inches ,  .x){
	if( missing( Inches ) ) {
	stop('You must specify a value for the argument(s)  Inches ')
	}
	
	ans = .COM(.x, 'InchesToPoints', Inches, .dispatch = as.integer(1), .ids =370, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CentimetersToPoints' = function( Centimeters ,  .x){
	if( missing( Centimeters ) ) {
	stop('You must specify a value for the argument(s)  Centimeters ')
	}
	
	ans = .COM(.x, 'CentimetersToPoints', Centimeters, .dispatch = as.integer(1), .ids =371, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'MillimetersToPoints' = function( Millimeters ,  .x){
	if( missing( Millimeters ) ) {
	stop('You must specify a value for the argument(s)  Millimeters ')
	}
	
	ans = .COM(.x, 'MillimetersToPoints', Millimeters, .dispatch = as.integer(1), .ids =372, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PicasToPoints' = function( Picas ,  .x){
	if( missing( Picas ) ) {
	stop('You must specify a value for the argument(s)  Picas ')
	}
	
	ans = .COM(.x, 'PicasToPoints', Picas, .dispatch = as.integer(1), .ids =373, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'LinesToPoints' = function( Lines ,  .x){
	if( missing( Lines ) ) {
	stop('You must specify a value for the argument(s)  Lines ')
	}
	
	ans = .COM(.x, 'LinesToPoints', Lines, .dispatch = as.integer(1), .ids =374, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PointsToInches' = function( Points ,  .x){
	if( missing( Points ) ) {
	stop('You must specify a value for the argument(s)  Points ')
	}
	
	ans = .COM(.x, 'PointsToInches', Points, .dispatch = as.integer(1), .ids =380, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PointsToCentimeters' = function( Points ,  .x){
	if( missing( Points ) ) {
	stop('You must specify a value for the argument(s)  Points ')
	}
	
	ans = .COM(.x, 'PointsToCentimeters', Points, .dispatch = as.integer(1), .ids =381, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PointsToMillimeters' = function( Points ,  .x){
	if( missing( Points ) ) {
	stop('You must specify a value for the argument(s)  Points ')
	}
	
	ans = .COM(.x, 'PointsToMillimeters', Points, .dispatch = as.integer(1), .ids =382, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PointsToPicas' = function( Points ,  .x){
	if( missing( Points ) ) {
	stop('You must specify a value for the argument(s)  Points ')
	}
	
	ans = .COM(.x, 'PointsToPicas', Points, .dispatch = as.integer(1), .ids =383, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PointsToLines' = function( Points ,  .x){
	if( missing( Points ) ) {
	stop('You must specify a value for the argument(s)  Points ')
	}
	
	ans = .COM(.x, 'PointsToLines', Points, .dispatch = as.integer(1), .ids =384, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Activate' = function(   .x){
	ans = .COM(.x, 'Activate', .dispatch = as.integer(1), .ids =385, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PointsToPixels' = function( Points, fVertical = NA ,  .x){
	if( missing( Points ) ) {
	stop('You must specify a value for the argument(s)  Points ')
	}
	
	
	ans = .COM(.x, 'PointsToPixels', Points, fVertical, .dispatch = as.integer(1), .ids =387, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PixelsToPoints' = function( Pixels, fVertical = NA ,  .x){
	if( missing( Pixels ) ) {
	stop('You must specify a value for the argument(s)  Pixels ')
	}
	
	
	ans = .COM(.x, 'PixelsToPoints', Pixels, fVertical, .dispatch = as.integer(1), .ids =388, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'KeyboardLatin' = function(   .x){
	ans = .COM(.x, 'KeyboardLatin', .dispatch = as.integer(1), .ids =400, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'KeyboardBidi' = function(   .x){
	ans = .COM(.x, 'KeyboardBidi', .dispatch = as.integer(1), .ids =401, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ToggleKeyboard' = function(   .x){
	ans = .COM(.x, 'ToggleKeyboard', .dispatch = as.integer(1), .ids =402, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Keyboard' = function( LangId = NA ,  .x){
	if(!missing( LangId )) LangId = as(LangId, 'integer')
	ans = .COM(.x, 'Keyboard', LangId, .dispatch = as.integer(1), .ids =446, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'ProductCode' = function(   .x){
	ans = .COM(.x, 'ProductCode', .dispatch = as.integer(1), .ids =404, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'DefaultWebOptions' = function(   .x){
	ans = .COM(.x, 'DefaultWebOptions', .dispatch = as.integer(1), .ids =405, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'DefaultWebOptions')
},
'DiscussionSupport' = function( Range, cid, piCSE ,  .x){
	if( missing( Range )||missing( cid )||missing( piCSE ) ) {
	stop('You must specify a value for the argument(s)  Range, cid, piCSE ')
	}
	
	
	
	ans = .COM(.x, 'DiscussionSupport', Range, cid, piCSE, .dispatch = as.integer(1), .ids =407, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SetDefaultTheme' = function( Name, DocumentType ,  .x){
	if( missing( Name )||missing( DocumentType ) ) {
	stop('You must specify a value for the argument(s)  Name, DocumentType ')
	}
	Name = as(Name, 'character')
	DocumentType = as(DocumentType, 'WdMailSystem')
	ans = .COM(.x, 'SetDefaultTheme', Name, DocumentType, .dispatch = as.integer(1), .ids =414, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GetDefaultTheme' = function( DocumentType ,  .x){
	if( missing( DocumentType ) ) {
	stop('You must specify a value for the argument(s)  DocumentType ')
	}
	DocumentType = as(DocumentType, 'WdMailSystem')
	ans = .COM(.x, 'GetDefaultTheme', DocumentType, .dispatch = as.integer(1), .ids =416, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'PrintOut2000' = function( Background = NA, Append = NA, Range = NA, OutputFileName = NA, From = NA, To = NA, Item = NA, Copies = NA, Pages = NA, PageType = NA, PrintToFile = NA, Collate = NA, FileName = NA, ActivePrinterMacGX = NA, ManualDuplexPrint = NA, PrintZoomColumn = NA, PrintZoomRow = NA, PrintZoomPaperWidth = NA, PrintZoomPaperHeight = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'PrintOut2000', Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight, .dispatch = as.integer(1), .ids =444, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Run' = function( MacroName, varg1 = NA, varg2 = NA, varg3 = NA, varg4 = NA, varg5 = NA, varg6 = NA, varg7 = NA, varg8 = NA, varg9 = NA, varg10 = NA, varg11 = NA, varg12 = NA, varg13 = NA, varg14 = NA, varg15 = NA, varg16 = NA, varg17 = NA, varg18 = NA, varg19 = NA, varg20 = NA, varg21 = NA, varg22 = NA, varg23 = NA, varg24 = NA, varg25 = NA, varg26 = NA, varg27 = NA, varg28 = NA, varg29 = NA, varg30 = NA ,  .x){
	if( missing( MacroName ) ) {
	stop('You must specify a value for the argument(s)  MacroName ')
	}
	MacroName = as(MacroName, 'character')
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'Run', MacroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29, varg30, .dispatch = as.integer(1), .ids =445, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PrintOut' = function( Background = NA, Append = NA, Range = NA, OutputFileName = NA, From = NA, To = NA, Item = NA, Copies = NA, Pages = NA, PageType = NA, PrintToFile = NA, Collate = NA, FileName = NA, ActivePrinterMacGX = NA, ManualDuplexPrint = NA, PrintZoomColumn = NA, PrintZoomRow = NA, PrintZoomPaperWidth = NA, PrintZoomPaperHeight = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'PrintOut', Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight, .dispatch = as.integer(1), .ids =448, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Dummy2' = function(   .x){
	ans = .COM(.x, 'Dummy2', .dispatch = as.integer(1), .ids =458, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'PutFocusInMailHeader' = function(   .x){
	ans = .COM(.x, 'PutFocusInMailHeader', .dispatch = as.integer(1), .ids =464, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
} )
'COM._Document.GetProperty'  = list('Name' = function(x) {
		 ans = .COM(x, 'Name', .dispatch = as.integer(2))
	ans
},
'Application' = function(x) {
		 ans = .COM(x, 'Application', .dispatch = as.integer(2))
		 as(ans, 'Application')
},
'Creator' = function(x) {
		 ans = .COM(x, 'Creator', .dispatch = as.integer(2))
	ans
},
'Parent' = function(x) {
		 ans = .COM(x, 'Parent', .dispatch = as.integer(2))
		 ans
},
'BuiltInDocumentProperties' = function(x) {
		 ans = .COM(x, 'BuiltInDocumentProperties', .dispatch = as.integer(2))
		 ans
},
'CustomDocumentProperties' = function(x) {
		 ans = .COM(x, 'CustomDocumentProperties', .dispatch = as.integer(2))
		 ans
},
'Path' = function(x) {
		 ans = .COM(x, 'Path', .dispatch = as.integer(2))
	ans
},
'Bookmarks' = function(x) {
		 ans = .COM(x, 'Bookmarks', .dispatch = as.integer(2))
		 as(ans, 'Bookmarks')
},
'Tables' = function(x) {
		 ans = .COM(x, 'Tables', .dispatch = as.integer(2))
		 as(ans, 'Tables')
},
'Footnotes' = function(x) {
		 ans = .COM(x, 'Footnotes', .dispatch = as.integer(2))
		 as(ans, 'Footnotes')
},
'Endnotes' = function(x) {
		 ans = .COM(x, 'Endnotes', .dispatch = as.integer(2))
		 as(ans, 'Endnotes')
},
'Comments' = function(x) {
		 ans = .COM(x, 'Comments', .dispatch = as.integer(2))
		 as(ans, 'Comments')
},
'Type' = function(x) {
		 ans = .COM(x, 'Type', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'AutoHyphenation' = function(x) {
		 ans = .COM(x, 'AutoHyphenation', .dispatch = as.integer(2))
	ans
},
'HyphenateCaps' = function(x) {
		 ans = .COM(x, 'HyphenateCaps', .dispatch = as.integer(2))
	ans
},
'HyphenationZone' = function(x) {
		 ans = .COM(x, 'HyphenationZone', .dispatch = as.integer(2))
	ans
},
'ConsecutiveHyphensLimit' = function(x) {
		 ans = .COM(x, 'ConsecutiveHyphensLimit', .dispatch = as.integer(2))
	ans
},
'Sections' = function(x) {
		 ans = .COM(x, 'Sections', .dispatch = as.integer(2))
		 as(ans, 'Sections')
},
'Paragraphs' = function(x) {
		 ans = .COM(x, 'Paragraphs', .dispatch = as.integer(2))
		 as(ans, 'Paragraphs')
},
'Words' = function(x) {
		 ans = .COM(x, 'Words', .dispatch = as.integer(2))
		 as(ans, 'Words')
},
'Sentences' = function(x) {
		 ans = .COM(x, 'Sentences', .dispatch = as.integer(2))
		 as(ans, 'Sentences')
},
'Characters' = function(x) {
		 ans = .COM(x, 'Characters', .dispatch = as.integer(2))
		 as(ans, 'Characters')
},
'Fields' = function(x) {
		 ans = .COM(x, 'Fields', .dispatch = as.integer(2))
		 as(ans, 'Fields')
},
'FormFields' = function(x) {
		 ans = .COM(x, 'FormFields', .dispatch = as.integer(2))
		 as(ans, 'FormFields')
},
'Styles' = function(x) {
		 ans = .COM(x, 'Styles', .dispatch = as.integer(2))
		 as(ans, 'Styles')
},
'Frames' = function(x) {
		 ans = .COM(x, 'Frames', .dispatch = as.integer(2))
		 as(ans, 'Frames')
},
'TablesOfFigures' = function(x) {
		 ans = .COM(x, 'TablesOfFigures', .dispatch = as.integer(2))
		 as(ans, 'TablesOfFigures')
},
'Variables' = function(x) {
		 ans = .COM(x, 'Variables', .dispatch = as.integer(2))
		 as(ans, 'Variables')
},
'MailMerge' = function(x) {
		 ans = .COM(x, 'MailMerge', .dispatch = as.integer(2))
		 as(ans, 'MailMerge')
},
'Envelope' = function(x) {
		 ans = .COM(x, 'Envelope', .dispatch = as.integer(2))
		 as(ans, 'Envelope')
},
'FullName' = function(x) {
		 ans = .COM(x, 'FullName', .dispatch = as.integer(2))
	ans
},
'Revisions' = function(x) {
		 ans = .COM(x, 'Revisions', .dispatch = as.integer(2))
		 as(ans, 'Revisions')
},
'TablesOfContents' = function(x) {
		 ans = .COM(x, 'TablesOfContents', .dispatch = as.integer(2))
		 as(ans, 'TablesOfContents')
},
'TablesOfAuthorities' = function(x) {
		 ans = .COM(x, 'TablesOfAuthorities', .dispatch = as.integer(2))
		 as(ans, 'TablesOfAuthorities')
},
'PageSetup' = function(x) {
		 ans = .COM(x, 'PageSetup', .dispatch = as.integer(2))
		 as(ans, 'PageSetup')
},
'Windows' = function(x) {
		 ans = .COM(x, 'Windows', .dispatch = as.integer(2))
		 as(ans, 'Windows')
},
'HasRoutingSlip' = function(x) {
		 ans = .COM(x, 'HasRoutingSlip', .dispatch = as.integer(2))
	ans
},
'RoutingSlip' = function(x) {
		 ans = .COM(x, 'RoutingSlip', .dispatch = as.integer(2))
		 as(ans, 'RoutingSlip')
},
'Routed' = function(x) {
		 ans = .COM(x, 'Routed', .dispatch = as.integer(2))
	ans
},
'TablesOfAuthoritiesCategories' = function(x) {
		 ans = .COM(x, 'TablesOfAuthoritiesCategories', .dispatch = as.integer(2))
		 as(ans, 'TablesOfAuthoritiesCategories')
},
'Indexes' = function(x) {
		 ans = .COM(x, 'Indexes', .dispatch = as.integer(2))
		 as(ans, 'Indexes')
},
'Saved' = function(x) {
		 ans = .COM(x, 'Saved', .dispatch = as.integer(2))
	ans
},
'Content' = function(x) {
		 ans = .COM(x, 'Content', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'ActiveWindow' = function(x) {
		 ans = .COM(x, 'ActiveWindow', .dispatch = as.integer(2))
		 as(ans, 'Window')
},
'Kind' = function(x) {
		 ans = .COM(x, 'Kind', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'ReadOnly' = function(x) {
		 ans = .COM(x, 'ReadOnly', .dispatch = as.integer(2))
	ans
},
'Subdocuments' = function(x) {
		 ans = .COM(x, 'Subdocuments', .dispatch = as.integer(2))
		 as(ans, 'Subdocuments')
},
'IsMasterDocument' = function(x) {
		 ans = .COM(x, 'IsMasterDocument', .dispatch = as.integer(2))
	ans
},
'DefaultTabStop' = function(x) {
		 ans = .COM(x, 'DefaultTabStop', .dispatch = as.integer(2))
		 ans
},
'EmbedTrueTypeFonts' = function(x) {
		 ans = .COM(x, 'EmbedTrueTypeFonts', .dispatch = as.integer(2))
	ans
},
'SaveFormsData' = function(x) {
		 ans = .COM(x, 'SaveFormsData', .dispatch = as.integer(2))
	ans
},
'ReadOnlyRecommended' = function(x) {
		 ans = .COM(x, 'ReadOnlyRecommended', .dispatch = as.integer(2))
	ans
},
'SaveSubsetFonts' = function(x) {
		 ans = .COM(x, 'SaveSubsetFonts', .dispatch = as.integer(2))
	ans
},
'StoryRanges' = function(x) {
		 ans = .COM(x, 'StoryRanges', .dispatch = as.integer(2))
		 as(ans, 'StoryRanges')
},
'CommandBars' = function(x) {
		 ans = .COM(x, 'CommandBars', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'IsSubdocument' = function(x) {
		 ans = .COM(x, 'IsSubdocument', .dispatch = as.integer(2))
	ans
},
'SaveFormat' = function(x) {
		 ans = .COM(x, 'SaveFormat', .dispatch = as.integer(2))
	ans
},
'ProtectionType' = function(x) {
		 ans = .COM(x, 'ProtectionType', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'Hyperlinks' = function(x) {
		 ans = .COM(x, 'Hyperlinks', .dispatch = as.integer(2))
		 as(ans, 'Hyperlinks')
},
'Shapes' = function(x) {
		 ans = .COM(x, 'Shapes', .dispatch = as.integer(2))
		 as(ans, 'Shapes')
},
'ListTemplates' = function(x) {
		 ans = .COM(x, 'ListTemplates', .dispatch = as.integer(2))
		 as(ans, 'ListTemplates')
},
'Lists' = function(x) {
		 ans = .COM(x, 'Lists', .dispatch = as.integer(2))
		 as(ans, 'Lists')
},
'UpdateStylesOnOpen' = function(x) {
		 ans = .COM(x, 'UpdateStylesOnOpen', .dispatch = as.integer(2))
	ans
},
'AttachedTemplate' = function(x) {
		 ans = .COM(x, 'AttachedTemplate', .dispatch = as.integer(2))
		 ans
},
'InlineShapes' = function(x) {
		 ans = .COM(x, 'InlineShapes', .dispatch = as.integer(2))
		 as(ans, 'InlineShapes')
},
'Background' = function(x) {
		 ans = .COM(x, 'Background', .dispatch = as.integer(2))
		 as(ans, 'Shape')
},
'GrammarChecked' = function(x) {
		 ans = .COM(x, 'GrammarChecked', .dispatch = as.integer(2))
	ans
},
'SpellingChecked' = function(x) {
		 ans = .COM(x, 'SpellingChecked', .dispatch = as.integer(2))
	ans
},
'ShowGrammaticalErrors' = function(x) {
		 ans = .COM(x, 'ShowGrammaticalErrors', .dispatch = as.integer(2))
	ans
},
'ShowSpellingErrors' = function(x) {
		 ans = .COM(x, 'ShowSpellingErrors', .dispatch = as.integer(2))
	ans
},
'Versions' = function(x) {
		 ans = .COM(x, 'Versions', .dispatch = as.integer(2))
		 as(ans, 'Versions')
},
'ShowSummary' = function(x) {
		 ans = .COM(x, 'ShowSummary', .dispatch = as.integer(2))
	ans
},
'SummaryViewMode' = function(x) {
		 ans = .COM(x, 'SummaryViewMode', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'SummaryLength' = function(x) {
		 ans = .COM(x, 'SummaryLength', .dispatch = as.integer(2))
	ans
},
'PrintFractionalWidths' = function(x) {
		 ans = .COM(x, 'PrintFractionalWidths', .dispatch = as.integer(2))
	ans
},
'PrintPostScriptOverText' = function(x) {
		 ans = .COM(x, 'PrintPostScriptOverText', .dispatch = as.integer(2))
	ans
},
'Container' = function(x) {
		 ans = .COM(x, 'Container', .dispatch = as.integer(2))
		 ans
},
'PrintFormsData' = function(x) {
		 ans = .COM(x, 'PrintFormsData', .dispatch = as.integer(2))
	ans
},
'ListParagraphs' = function(x) {
		 ans = .COM(x, 'ListParagraphs', .dispatch = as.integer(2))
		 as(ans, 'ListParagraphs')
},
'HasPassword' = function(x) {
		 ans = .COM(x, 'HasPassword', .dispatch = as.integer(2))
	ans
},
'WriteReserved' = function(x) {
		 ans = .COM(x, 'WriteReserved', .dispatch = as.integer(2))
	ans
},
'UserControl' = function(x) {
		 ans = .COM(x, 'UserControl', .dispatch = as.integer(2))
	ans
},
'HasMailer' = function(x) {
		 ans = .COM(x, 'HasMailer', .dispatch = as.integer(2))
	ans
},
'Mailer' = function(x) {
		 ans = .COM(x, 'Mailer', .dispatch = as.integer(2))
		 as(ans, 'Mailer')
},
'ReadabilityStatistics' = function(x) {
		 ans = .COM(x, 'ReadabilityStatistics', .dispatch = as.integer(2))
		 as(ans, 'ReadabilityStatistics')
},
'GrammaticalErrors' = function(x) {
		 ans = .COM(x, 'GrammaticalErrors', .dispatch = as.integer(2))
		 as(ans, 'ProofreadingErrors')
},
'SpellingErrors' = function(x) {
		 ans = .COM(x, 'SpellingErrors', .dispatch = as.integer(2))
		 as(ans, 'ProofreadingErrors')
},
'VBProject' = function(x) {
		 ans = .COM(x, 'VBProject', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'FormsDesign' = function(x) {
		 ans = .COM(x, 'FormsDesign', .dispatch = as.integer(2))
	ans
},
'_CodeName' = function(x) {
		 ans = .COM(x, '_CodeName', .dispatch = as.integer(2))
	ans
},
'CodeName' = function(x) {
		 ans = .COM(x, 'CodeName', .dispatch = as.integer(2))
	ans
},
'SnapToGrid' = function(x) {
		 ans = .COM(x, 'SnapToGrid', .dispatch = as.integer(2))
	ans
},
'SnapToShapes' = function(x) {
		 ans = .COM(x, 'SnapToShapes', .dispatch = as.integer(2))
	ans
},
'GridDistanceHorizontal' = function(x) {
		 ans = .COM(x, 'GridDistanceHorizontal', .dispatch = as.integer(2))
		 ans
},
'GridDistanceVertical' = function(x) {
		 ans = .COM(x, 'GridDistanceVertical', .dispatch = as.integer(2))
		 ans
},
'GridOriginHorizontal' = function(x) {
		 ans = .COM(x, 'GridOriginHorizontal', .dispatch = as.integer(2))
		 ans
},
'GridOriginVertical' = function(x) {
		 ans = .COM(x, 'GridOriginVertical', .dispatch = as.integer(2))
		 ans
},
'GridSpaceBetweenHorizontalLines' = function(x) {
		 ans = .COM(x, 'GridSpaceBetweenHorizontalLines', .dispatch = as.integer(2))
	ans
},
'GridSpaceBetweenVerticalLines' = function(x) {
		 ans = .COM(x, 'GridSpaceBetweenVerticalLines', .dispatch = as.integer(2))
	ans
},
'GridOriginFromMargin' = function(x) {
		 ans = .COM(x, 'GridOriginFromMargin', .dispatch = as.integer(2))
	ans
},
'KerningByAlgorithm' = function(x) {
		 ans = .COM(x, 'KerningByAlgorithm', .dispatch = as.integer(2))
	ans
},
'JustificationMode' = function(x) {
		 ans = .COM(x, 'JustificationMode', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'FarEastLineBreakLevel' = function(x) {
		 ans = .COM(x, 'FarEastLineBreakLevel', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'NoLineBreakBefore' = function(x) {
		 ans = .COM(x, 'NoLineBreakBefore', .dispatch = as.integer(2))
	ans
},
'NoLineBreakAfter' = function(x) {
		 ans = .COM(x, 'NoLineBreakAfter', .dispatch = as.integer(2))
	ans
},
'TrackRevisions' = function(x) {
		 ans = .COM(x, 'TrackRevisions', .dispatch = as.integer(2))
	ans
},
'PrintRevisions' = function(x) {
		 ans = .COM(x, 'PrintRevisions', .dispatch = as.integer(2))
	ans
},
'ShowRevisions' = function(x) {
		 ans = .COM(x, 'ShowRevisions', .dispatch = as.integer(2))
	ans
},
'ActiveTheme' = function(x) {
		 ans = .COM(x, 'ActiveTheme', .dispatch = as.integer(2))
	ans
},
'ActiveThemeDisplayName' = function(x) {
		 ans = .COM(x, 'ActiveThemeDisplayName', .dispatch = as.integer(2))
	ans
},
'Email' = function(x) {
		 ans = .COM(x, 'Email', .dispatch = as.integer(2))
		 as(ans, 'Email')
},
'Scripts' = function(x) {
		 ans = .COM(x, 'Scripts', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'LanguageDetected' = function(x) {
		 ans = .COM(x, 'LanguageDetected', .dispatch = as.integer(2))
	ans
},
'FarEastLineBreakLanguage' = function(x) {
		 ans = .COM(x, 'FarEastLineBreakLanguage', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'Frameset' = function(x) {
		 ans = .COM(x, 'Frameset', .dispatch = as.integer(2))
		 as(ans, 'Frameset')
},
'ClickAndTypeParagraphStyle' = function(x) {
		 ans = .COM(x, 'ClickAndTypeParagraphStyle', .dispatch = as.integer(2))
		 ans
},
'HTMLProject' = function(x) {
		 ans = .COM(x, 'HTMLProject', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'WebOptions' = function(x) {
		 ans = .COM(x, 'WebOptions', .dispatch = as.integer(2))
		 as(ans, 'WebOptions')
},
'OpenEncoding' = function(x) {
		 ans = .COM(x, 'OpenEncoding', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'SaveEncoding' = function(x) {
		 ans = .COM(x, 'SaveEncoding', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'OptimizeForWord97' = function(x) {
		 ans = .COM(x, 'OptimizeForWord97', .dispatch = as.integer(2))
	ans
},
'VBASigned' = function(x) {
		 ans = .COM(x, 'VBASigned', .dispatch = as.integer(2))
	ans
},
'MailEnvelope' = function(x) {
		 ans = .COM(x, 'MailEnvelope', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'DisableFeatures' = function(x) {
		 ans = .COM(x, 'DisableFeatures', .dispatch = as.integer(2))
	ans
},
'DoNotEmbedSystemFonts' = function(x) {
		 ans = .COM(x, 'DoNotEmbedSystemFonts', .dispatch = as.integer(2))
	ans
},
'Signatures' = function(x) {
		 ans = .COM(x, 'Signatures', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'DefaultTargetFrame' = function(x) {
		 ans = .COM(x, 'DefaultTargetFrame', .dispatch = as.integer(2))
	ans
},
'HTMLDivisions' = function(x) {
		 ans = .COM(x, 'HTMLDivisions', .dispatch = as.integer(2))
		 as(ans, 'HTMLDivisions')
},
'DisableFeaturesIntroducedAfter' = function(x) {
		 ans = .COM(x, 'DisableFeaturesIntroducedAfter', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'RemovePersonalInformation' = function(x) {
		 ans = .COM(x, 'RemovePersonalInformation', .dispatch = as.integer(2))
	ans
},
'SmartTags' = function(x) {
		 ans = .COM(x, 'SmartTags', .dispatch = as.integer(2))
		 as(ans, 'SmartTags')
},
'EmbedSmartTags' = function(x) {
		 ans = .COM(x, 'EmbedSmartTags', .dispatch = as.integer(2))
	ans
},
'SmartTagsAsXMLProps' = function(x) {
		 ans = .COM(x, 'SmartTagsAsXMLProps', .dispatch = as.integer(2))
	ans
},
'TextEncoding' = function(x) {
		 ans = .COM(x, 'TextEncoding', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'TextLineEnding' = function(x) {
		 ans = .COM(x, 'TextLineEnding', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'StyleSheets' = function(x) {
		 ans = .COM(x, 'StyleSheets', .dispatch = as.integer(2))
		 as(ans, 'StyleSheets')
},
'DefaultTableStyle' = function(x) {
		 ans = .COM(x, 'DefaultTableStyle', .dispatch = as.integer(2))
		 ans
},
'PasswordEncryptionProvider' = function(x) {
		 ans = .COM(x, 'PasswordEncryptionProvider', .dispatch = as.integer(2))
	ans
},
'PasswordEncryptionAlgorithm' = function(x) {
		 ans = .COM(x, 'PasswordEncryptionAlgorithm', .dispatch = as.integer(2))
	ans
},
'PasswordEncryptionKeyLength' = function(x) {
		 ans = .COM(x, 'PasswordEncryptionKeyLength', .dispatch = as.integer(2))
	ans
},
'PasswordEncryptionFileProperties' = function(x) {
		 ans = .COM(x, 'PasswordEncryptionFileProperties', .dispatch = as.integer(2))
	ans
},
'EmbedLinguisticData' = function(x) {
		 ans = .COM(x, 'EmbedLinguisticData', .dispatch = as.integer(2))
	ans
},
'FormattingShowFont' = function(x) {
		 ans = .COM(x, 'FormattingShowFont', .dispatch = as.integer(2))
	ans
},
'FormattingShowClear' = function(x) {
		 ans = .COM(x, 'FormattingShowClear', .dispatch = as.integer(2))
	ans
},
'FormattingShowParagraph' = function(x) {
		 ans = .COM(x, 'FormattingShowParagraph', .dispatch = as.integer(2))
	ans
},
'FormattingShowNumbering' = function(x) {
		 ans = .COM(x, 'FormattingShowNumbering', .dispatch = as.integer(2))
	ans
},
'FormattingShowFilter' = function(x) {
		 ans = .COM(x, 'FormattingShowFilter', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'Permission' = function(x) {
		 ans = .COM(x, 'Permission', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'XMLNodes' = function(x) {
		 ans = .COM(x, 'XMLNodes', .dispatch = as.integer(2))
		 as(ans, 'XMLNodes')
},
'XMLSchemaReferences' = function(x) {
		 ans = .COM(x, 'XMLSchemaReferences', .dispatch = as.integer(2))
		 as(ans, 'XMLSchemaReferences')
},
'SmartDocument' = function(x) {
		 ans = .COM(x, 'SmartDocument', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'SharedWorkspace' = function(x) {
		 ans = .COM(x, 'SharedWorkspace', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'Sync' = function(x) {
		 ans = .COM(x, 'Sync', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'EnforceStyle' = function(x) {
		 ans = .COM(x, 'EnforceStyle', .dispatch = as.integer(2))
	ans
},
'AutoFormatOverride' = function(x) {
		 ans = .COM(x, 'AutoFormatOverride', .dispatch = as.integer(2))
	ans
},
'XMLSaveDataOnly' = function(x) {
		 ans = .COM(x, 'XMLSaveDataOnly', .dispatch = as.integer(2))
	ans
},
'XMLHideNamespaces' = function(x) {
		 ans = .COM(x, 'XMLHideNamespaces', .dispatch = as.integer(2))
	ans
},
'XMLShowAdvancedErrors' = function(x) {
		 ans = .COM(x, 'XMLShowAdvancedErrors', .dispatch = as.integer(2))
	ans
},
'XMLUseXSLTWhenSaving' = function(x) {
		 ans = .COM(x, 'XMLUseXSLTWhenSaving', .dispatch = as.integer(2))
	ans
},
'XMLSaveThroughXSLT' = function(x) {
		 ans = .COM(x, 'XMLSaveThroughXSLT', .dispatch = as.integer(2))
	ans
},
'DocumentLibraryVersions' = function(x) {
		 ans = .COM(x, 'DocumentLibraryVersions', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'ReadingModeLayoutFrozen' = function(x) {
		 ans = .COM(x, 'ReadingModeLayoutFrozen', .dispatch = as.integer(2))
	ans
},
'RemoveDateAndTime' = function(x) {
		 ans = .COM(x, 'RemoveDateAndTime', .dispatch = as.integer(2))
	ans
},
'ChildNodeSuggestions' = function(x) {
		 ans = .COM(x, 'ChildNodeSuggestions', .dispatch = as.integer(2))
		 as(ans, 'XMLChildNodeSuggestions')
},
'XMLSchemaViolations' = function(x) {
		 ans = .COM(x, 'XMLSchemaViolations', .dispatch = as.integer(2))
		 as(ans, 'XMLNodes')
},
'ReadingLayoutSizeX' = function(x) {
		 ans = .COM(x, 'ReadingLayoutSizeX', .dispatch = as.integer(2))
	ans
},
'ReadingLayoutSizeY' = function(x) {
		 ans = .COM(x, 'ReadingLayoutSizeY', .dispatch = as.integer(2))
	ans
} )
'COM._Document.SetProperty'  = list('AutoHyphenation' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'AutoHyphenation', value, .dispatch = as.integer(4))
},
'HyphenateCaps' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'HyphenateCaps', value, .dispatch = as.integer(4))
},
'HyphenationZone' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'HyphenationZone', value, .dispatch = as.integer(4))
},
'ConsecutiveHyphensLimit' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'ConsecutiveHyphensLimit', value, .dispatch = as.integer(4))
},
'PageSetup' = function(x, value) {
	value = as(value, 'PageSetup')
	.COM(x, 'PageSetup', value, .dispatch = as.integer(4))
},
'HasRoutingSlip' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'HasRoutingSlip', value, .dispatch = as.integer(4))
},
'Saved' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'Saved', value, .dispatch = as.integer(4))
},
'Kind' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'Kind', value, .dispatch = as.integer(4))
},
'DefaultTabStop' = function(x, value) {
	
	.COM(x, 'DefaultTabStop', value, .dispatch = as.integer(4))
},
'EmbedTrueTypeFonts' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'EmbedTrueTypeFonts', value, .dispatch = as.integer(4))
},
'SaveFormsData' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'SaveFormsData', value, .dispatch = as.integer(4))
},
'ReadOnlyRecommended' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ReadOnlyRecommended', value, .dispatch = as.integer(4))
},
'SaveSubsetFonts' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'SaveSubsetFonts', value, .dispatch = as.integer(4))
},
'Compatibility' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'Compatibility', value, .dispatch = as.integer(4))
},
'UpdateStylesOnOpen' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'UpdateStylesOnOpen', value, .dispatch = as.integer(4))
},
'AttachedTemplate' = function(x, value) {
	
	.COM(x, 'AttachedTemplate', value, .dispatch = as.integer(4))
},
'Background' = function(x, value) {
	value = as(value, 'Shape')
	.COM(x, 'Background', value, .dispatch = as.integer(4))
},
'GrammarChecked' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'GrammarChecked', value, .dispatch = as.integer(4))
},
'SpellingChecked' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'SpellingChecked', value, .dispatch = as.integer(4))
},
'ShowGrammaticalErrors' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ShowGrammaticalErrors', value, .dispatch = as.integer(4))
},
'ShowSpellingErrors' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ShowSpellingErrors', value, .dispatch = as.integer(4))
},
'ShowSummary' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ShowSummary', value, .dispatch = as.integer(4))
},
'SummaryViewMode' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'SummaryViewMode', value, .dispatch = as.integer(4))
},
'SummaryLength' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'SummaryLength', value, .dispatch = as.integer(4))
},
'PrintFractionalWidths' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'PrintFractionalWidths', value, .dispatch = as.integer(4))
},
'PrintPostScriptOverText' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'PrintPostScriptOverText', value, .dispatch = as.integer(4))
},
'PrintFormsData' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'PrintFormsData', value, .dispatch = as.integer(4))
},
'Password' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'Password', value, .dispatch = as.integer(4))
},
'WritePassword' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'WritePassword', value, .dispatch = as.integer(4))
},
'ActiveWritingStyle' = function(x, value) {
	
	.COM(x, 'ActiveWritingStyle', value, .dispatch = as.integer(4))
},
'UserControl' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'UserControl', value, .dispatch = as.integer(4))
},
'HasMailer' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'HasMailer', value, .dispatch = as.integer(4))
},
'_CodeName' = function(x, value) {
	value = as(value, 'character')
	.COM(x, '_CodeName', value, .dispatch = as.integer(4))
},
'SnapToGrid' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'SnapToGrid', value, .dispatch = as.integer(4))
},
'SnapToShapes' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'SnapToShapes', value, .dispatch = as.integer(4))
},
'GridDistanceHorizontal' = function(x, value) {
	
	.COM(x, 'GridDistanceHorizontal', value, .dispatch = as.integer(4))
},
'GridDistanceVertical' = function(x, value) {
	
	.COM(x, 'GridDistanceVertical', value, .dispatch = as.integer(4))
},
'GridOriginHorizontal' = function(x, value) {
	
	.COM(x, 'GridOriginHorizontal', value, .dispatch = as.integer(4))
},
'GridOriginVertical' = function(x, value) {
	
	.COM(x, 'GridOriginVertical', value, .dispatch = as.integer(4))
},
'GridSpaceBetweenHorizontalLines' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'GridSpaceBetweenHorizontalLines', value, .dispatch = as.integer(4))
},
'GridSpaceBetweenVerticalLines' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'GridSpaceBetweenVerticalLines', value, .dispatch = as.integer(4))
},
'GridOriginFromMargin' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'GridOriginFromMargin', value, .dispatch = as.integer(4))
},
'KerningByAlgorithm' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'KerningByAlgorithm', value, .dispatch = as.integer(4))
},
'JustificationMode' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'JustificationMode', value, .dispatch = as.integer(4))
},
'FarEastLineBreakLevel' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'FarEastLineBreakLevel', value, .dispatch = as.integer(4))
},
'NoLineBreakBefore' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'NoLineBreakBefore', value, .dispatch = as.integer(4))
},
'NoLineBreakAfter' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'NoLineBreakAfter', value, .dispatch = as.integer(4))
},
'TrackRevisions' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'TrackRevisions', value, .dispatch = as.integer(4))
},
'PrintRevisions' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'PrintRevisions', value, .dispatch = as.integer(4))
},
'ShowRevisions' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ShowRevisions', value, .dispatch = as.integer(4))
},
'LanguageDetected' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'LanguageDetected', value, .dispatch = as.integer(4))
},
'FarEastLineBreakLanguage' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'FarEastLineBreakLanguage', value, .dispatch = as.integer(4))
},
'ClickAndTypeParagraphStyle' = function(x, value) {
	
	.COM(x, 'ClickAndTypeParagraphStyle', value, .dispatch = as.integer(4))
},
'SaveEncoding' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'SaveEncoding', value, .dispatch = as.integer(4))
},
'OptimizeForWord97' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'OptimizeForWord97', value, .dispatch = as.integer(4))
},
'DisableFeatures' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DisableFeatures', value, .dispatch = as.integer(4))
},
'DoNotEmbedSystemFonts' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DoNotEmbedSystemFonts', value, .dispatch = as.integer(4))
},
'DefaultTargetFrame' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'DefaultTargetFrame', value, .dispatch = as.integer(4))
},
'DisableFeaturesIntroducedAfter' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'DisableFeaturesIntroducedAfter', value, .dispatch = as.integer(4))
},
'RemovePersonalInformation' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'RemovePersonalInformation', value, .dispatch = as.integer(4))
},
'EmbedSmartTags' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'EmbedSmartTags', value, .dispatch = as.integer(4))
},
'SmartTagsAsXMLProps' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'SmartTagsAsXMLProps', value, .dispatch = as.integer(4))
},
'TextEncoding' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'TextEncoding', value, .dispatch = as.integer(4))
},
'TextLineEnding' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'TextLineEnding', value, .dispatch = as.integer(4))
},
'EmbedLinguisticData' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'EmbedLinguisticData', value, .dispatch = as.integer(4))
},
'FormattingShowFont' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'FormattingShowFont', value, .dispatch = as.integer(4))
},
'FormattingShowClear' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'FormattingShowClear', value, .dispatch = as.integer(4))
},
'FormattingShowParagraph' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'FormattingShowParagraph', value, .dispatch = as.integer(4))
},
'FormattingShowNumbering' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'FormattingShowNumbering', value, .dispatch = as.integer(4))
},
'FormattingShowFilter' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'FormattingShowFilter', value, .dispatch = as.integer(4))
},
'EnforceStyle' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'EnforceStyle', value, .dispatch = as.integer(4))
},
'AutoFormatOverride' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'AutoFormatOverride', value, .dispatch = as.integer(4))
},
'XMLSaveDataOnly' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'XMLSaveDataOnly', value, .dispatch = as.integer(4))
},
'XMLHideNamespaces' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'XMLHideNamespaces', value, .dispatch = as.integer(4))
},
'XMLShowAdvancedErrors' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'XMLShowAdvancedErrors', value, .dispatch = as.integer(4))
},
'XMLUseXSLTWhenSaving' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'XMLUseXSLTWhenSaving', value, .dispatch = as.integer(4))
},
'XMLSaveThroughXSLT' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'XMLSaveThroughXSLT', value, .dispatch = as.integer(4))
},
'ReadingModeLayoutFrozen' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ReadingModeLayoutFrozen', value, .dispatch = as.integer(4))
},
'RemoveDateAndTime' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'RemoveDateAndTime', value, .dispatch = as.integer(4))
},
'ReadingLayoutSizeX' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'ReadingLayoutSizeX', value, .dispatch = as.integer(4))
},
'ReadingLayoutSizeY' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'ReadingLayoutSizeY', value, .dispatch = as.integer(4))
} )
'COM._Document.Methods'  = list('Close' = function( SaveChanges = NA, OriginalFormat = NA, RouteDocument = NA ,  .x){
	
	
	
	ans = .COM(.x, 'Close', SaveChanges, OriginalFormat, RouteDocument, .dispatch = as.integer(1), .ids =1105, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SaveAs2000' = function( FileName = NA, FileFormat = NA, LockComments = NA, Password = NA, AddToRecentFiles = NA, WritePassword = NA, ReadOnlyRecommended = NA, EmbedTrueTypeFonts = NA, SaveNativePictureFormat = NA, SaveFormsData = NA, SaveAsAOCELetter = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'SaveAs2000', FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, .dispatch = as.integer(1), .ids =102, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Repaginate' = function(   .x){
	ans = .COM(.x, 'Repaginate', .dispatch = as.integer(1), .ids =103, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'FitToPages' = function(   .x){
	ans = .COM(.x, 'FitToPages', .dispatch = as.integer(1), .ids =104, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ManualHyphenation' = function(   .x){
	ans = .COM(.x, 'ManualHyphenation', .dispatch = as.integer(1), .ids =105, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Select' = function(   .x){
	ans = .COM(.x, 'Select', .dispatch = as.integer(1), .ids =65535, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DataForm' = function(   .x){
	ans = .COM(.x, 'DataForm', .dispatch = as.integer(1), .ids =106, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Route' = function(   .x){
	ans = .COM(.x, 'Route', .dispatch = as.integer(1), .ids =107, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Save' = function(   .x){
	ans = .COM(.x, 'Save', .dispatch = as.integer(1), .ids =108, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PrintOutOld' = function( Background = NA, Append = NA, Range = NA, OutputFileName = NA, From = NA, To = NA, Item = NA, Copies = NA, Pages = NA, PageType = NA, PrintToFile = NA, Collate = NA, ActivePrinterMacGX = NA, ManualDuplexPrint = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'PrintOutOld', Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, .dispatch = as.integer(1), .ids =109, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SendMail' = function(   .x){
	ans = .COM(.x, 'SendMail', .dispatch = as.integer(1), .ids =110, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Range' = function( Start = NA, End = NA ,  .x){
	
	
	ans = .COM(.x, 'Range', Start, End, .dispatch = as.integer(1), .ids =2000, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'RunAutoMacro' = function( Which ,  .x){
	if( missing( Which ) ) {
	stop('You must specify a value for the argument(s)  Which ')
	}
	Which = as(Which, 'WdMailSystem')
	ans = .COM(.x, 'RunAutoMacro', Which, .dispatch = as.integer(1), .ids =112, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Activate' = function(   .x){
	ans = .COM(.x, 'Activate', .dispatch = as.integer(1), .ids =113, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PrintPreview' = function(   .x){
	ans = .COM(.x, 'PrintPreview', .dispatch = as.integer(1), .ids =114, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GoTo' = function( What = NA, Which = NA, Count = NA, Name = NA ,  .x){
	
	
	
	
	ans = .COM(.x, 'GoTo', What, Which, Count, Name, .dispatch = as.integer(1), .ids =115, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'Undo' = function( Times = NA ,  .x){
	
	ans = .COM(.x, 'Undo', Times, .dispatch = as.integer(1), .ids =116, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Redo' = function( Times = NA ,  .x){
	
	ans = .COM(.x, 'Redo', Times, .dispatch = as.integer(1), .ids =117, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'ComputeStatistics' = function( Statistic, IncludeFootnotesAndEndnotes = NA ,  .x){
	if( missing( Statistic ) ) {
	stop('You must specify a value for the argument(s)  Statistic ')
	}
	Statistic = as(Statistic, 'WdMailSystem')
	
	ans = .COM(.x, 'ComputeStatistics', Statistic, IncludeFootnotesAndEndnotes, .dispatch = as.integer(1), .ids =118, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MakeCompatibilityDefault' = function(   .x){
	ans = .COM(.x, 'MakeCompatibilityDefault', .dispatch = as.integer(1), .ids =119, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Protect2002' = function( Type, NoReset = NA, Password = NA ,  .x){
	if( missing( Type ) ) {
	stop('You must specify a value for the argument(s)  Type ')
	}
	Type = as(Type, 'WdMailSystem')
	
	
	ans = .COM(.x, 'Protect2002', Type, NoReset, Password, .dispatch = as.integer(1), .ids =120, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Unprotect' = function( Password = NA ,  .x){
	
	ans = .COM(.x, 'Unprotect', Password, .dispatch = as.integer(1), .ids =121, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'EditionOptions' = function( Type, Option, Name, Format = NA ,  .x){
	if( missing( Type )||missing( Option )||missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Type, Option, Name ')
	}
	Type = as(Type, 'WdMailSystem')
	Option = as(Option, 'WdMailSystem')
	Name = as(Name, 'character')
	
	ans = .COM(.x, 'EditionOptions', Type, Option, Name, Format, .dispatch = as.integer(1), .ids =122, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RunLetterWizard' = function( LetterContent = NA, WizardMode = NA ,  .x){
	
	
	ans = .COM(.x, 'RunLetterWizard', LetterContent, WizardMode, .dispatch = as.integer(1), .ids =123, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GetLetterContent' = function(   .x){
	ans = .COM(.x, 'GetLetterContent', .dispatch = as.integer(1), .ids =124, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'LetterContent')
},
'SetLetterContent' = function( LetterContent ,  .x){
	if( missing( LetterContent ) ) {
	stop('You must specify a value for the argument(s)  LetterContent ')
	}
	
	ans = .COM(.x, 'SetLetterContent', LetterContent, .dispatch = as.integer(1), .ids =125, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CopyStylesFromTemplate' = function( Template ,  .x){
	if( missing( Template ) ) {
	stop('You must specify a value for the argument(s)  Template ')
	}
	Template = as(Template, 'character')
	ans = .COM(.x, 'CopyStylesFromTemplate', Template, .dispatch = as.integer(1), .ids =126, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'UpdateStyles' = function(   .x){
	ans = .COM(.x, 'UpdateStyles', .dispatch = as.integer(1), .ids =127, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CheckGrammar' = function(   .x){
	ans = .COM(.x, 'CheckGrammar', .dispatch = as.integer(1), .ids =131, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CheckSpelling' = function( CustomDictionary = NA, IgnoreUppercase = NA, AlwaysSuggest = NA, CustomDictionary2 = NA, CustomDictionary3 = NA, CustomDictionary4 = NA, CustomDictionary5 = NA, CustomDictionary6 = NA, CustomDictionary7 = NA, CustomDictionary8 = NA, CustomDictionary9 = NA, CustomDictionary10 = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'CheckSpelling', CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10, .dispatch = as.integer(1), .ids =132, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'FollowHyperlink' = function( Address = NA, SubAddress = NA, NewWindow = NA, AddHistory = NA, ExtraInfo = NA, Method = NA, HeaderInfo = NA ,  .x){
	
	
	
	
	
	
	
	ans = .COM(.x, 'FollowHyperlink', Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo, .dispatch = as.integer(1), .ids =135, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AddToFavorites' = function(   .x){
	ans = .COM(.x, 'AddToFavorites', .dispatch = as.integer(1), .ids =136, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Reload' = function(   .x){
	ans = .COM(.x, 'Reload', .dispatch = as.integer(1), .ids =137, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AutoSummarize' = function( Length = NA, Mode = NA, UpdateProperties = NA ,  .x){
	
	
	
	ans = .COM(.x, 'AutoSummarize', Length, Mode, UpdateProperties, .dispatch = as.integer(1), .ids =138, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'RemoveNumbers' = function( NumberType = NA ,  .x){
	
	ans = .COM(.x, 'RemoveNumbers', NumberType, .dispatch = as.integer(1), .ids =140, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ConvertNumbersToText' = function( NumberType = NA ,  .x){
	
	ans = .COM(.x, 'ConvertNumbersToText', NumberType, .dispatch = as.integer(1), .ids =141, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CountNumberedItems' = function( NumberType = NA, Level = NA ,  .x){
	
	
	ans = .COM(.x, 'CountNumberedItems', NumberType, Level, .dispatch = as.integer(1), .ids =142, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Post' = function(   .x){
	ans = .COM(.x, 'Post', .dispatch = as.integer(1), .ids =143, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ToggleFormsDesign' = function(   .x){
	ans = .COM(.x, 'ToggleFormsDesign', .dispatch = as.integer(1), .ids =144, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Compare2000' = function( Name ,  .x){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	Name = as(Name, 'character')
	ans = .COM(.x, 'Compare2000', Name, .dispatch = as.integer(1), .ids =145, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'UpdateSummaryProperties' = function(   .x){
	ans = .COM(.x, 'UpdateSummaryProperties', .dispatch = as.integer(1), .ids =146, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GetCrossReferenceItems' = function( ReferenceType ,  .x){
	if( missing( ReferenceType ) ) {
	stop('You must specify a value for the argument(s)  ReferenceType ')
	}
	
	ans = .COM(.x, 'GetCrossReferenceItems', ReferenceType, .dispatch = as.integer(1), .ids =147, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AutoFormat' = function(   .x){
	ans = .COM(.x, 'AutoFormat', .dispatch = as.integer(1), .ids =148, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ViewCode' = function(   .x){
	ans = .COM(.x, 'ViewCode', .dispatch = as.integer(1), .ids =149, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ViewPropertyBrowser' = function(   .x){
	ans = .COM(.x, 'ViewPropertyBrowser', .dispatch = as.integer(1), .ids =150, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ForwardMailer' = function(   .x){
	ans = .COM(.x, 'ForwardMailer', .dispatch = as.integer(1), .ids =250, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Reply' = function(   .x){
	ans = .COM(.x, 'Reply', .dispatch = as.integer(1), .ids =251, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ReplyAll' = function(   .x){
	ans = .COM(.x, 'ReplyAll', .dispatch = as.integer(1), .ids =252, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SendMailer' = function( FileFormat = NA, Priority = NA ,  .x){
	
	
	ans = .COM(.x, 'SendMailer', FileFormat, Priority, .dispatch = as.integer(1), .ids =253, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'UndoClear' = function(   .x){
	ans = .COM(.x, 'UndoClear', .dispatch = as.integer(1), .ids =254, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PresentIt' = function(   .x){
	ans = .COM(.x, 'PresentIt', .dispatch = as.integer(1), .ids =255, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SendFax' = function( Address, Subject = NA ,  .x){
	if( missing( Address ) ) {
	stop('You must specify a value for the argument(s)  Address ')
	}
	Address = as(Address, 'character')
	
	ans = .COM(.x, 'SendFax', Address, Subject, .dispatch = as.integer(1), .ids =256, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Merge2000' = function( FileName ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	FileName = as(FileName, 'character')
	ans = .COM(.x, 'Merge2000', FileName, .dispatch = as.integer(1), .ids =257, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ClosePrintPreview' = function(   .x){
	ans = .COM(.x, 'ClosePrintPreview', .dispatch = as.integer(1), .ids =258, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CheckConsistency' = function(   .x){
	ans = .COM(.x, 'CheckConsistency', .dispatch = as.integer(1), .ids =259, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CreateLetterContent' = function( DateFormat, IncludeHeaderFooter, PageDesign, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientReference, MailingInstructions, AttentionLine, Subject, CCList, ReturnAddress, SenderName, Closing, SenderCompany, SenderJobTitle, SenderInitials, EnclosureNumber, InfoBlock = NA, RecipientCode = NA, RecipientGender = NA, ReturnAddressShortForm = NA, SenderCity = NA, SenderCode = NA, SenderGender = NA, SenderReference = NA ,  .x){
	if( missing( DateFormat )||missing( IncludeHeaderFooter )||missing( PageDesign )||missing( LetterStyle )||missing( Letterhead )||missing( LetterheadLocation )||missing( LetterheadSize )||missing( RecipientName )||missing( RecipientAddress )||missing( Salutation )||missing( SalutationType )||missing( RecipientReference )||missing( MailingInstructions )||missing( AttentionLine )||missing( Subject )||missing( CCList )||missing( ReturnAddress )||missing( SenderName )||missing( Closing )||missing( SenderCompany )||missing( SenderJobTitle )||missing( SenderInitials )||missing( EnclosureNumber ) ) {
	stop('You must specify a value for the argument(s)  DateFormat, IncludeHeaderFooter, PageDesign, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientReference, MailingInstructions, AttentionLine, Subject, CCList, ReturnAddress, SenderName, Closing, SenderCompany, SenderJobTitle, SenderInitials, EnclosureNumber ')
	}
	DateFormat = as(DateFormat, 'character')
	IncludeHeaderFooter = as(IncludeHeaderFooter, 'logical')
	PageDesign = as(PageDesign, 'character')
	LetterStyle = as(LetterStyle, 'WdMailSystem')
	Letterhead = as(Letterhead, 'logical')
	LetterheadLocation = as(LetterheadLocation, 'WdMailSystem')
	
	RecipientName = as(RecipientName, 'character')
	RecipientAddress = as(RecipientAddress, 'character')
	Salutation = as(Salutation, 'character')
	SalutationType = as(SalutationType, 'WdMailSystem')
	RecipientReference = as(RecipientReference, 'character')
	MailingInstructions = as(MailingInstructions, 'character')
	AttentionLine = as(AttentionLine, 'character')
	Subject = as(Subject, 'character')
	CCList = as(CCList, 'character')
	ReturnAddress = as(ReturnAddress, 'character')
	SenderName = as(SenderName, 'character')
	Closing = as(Closing, 'character')
	SenderCompany = as(SenderCompany, 'character')
	SenderJobTitle = as(SenderJobTitle, 'character')
	SenderInitials = as(SenderInitials, 'character')
	EnclosureNumber = as(EnclosureNumber, 'integer')
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'CreateLetterContent', DateFormat, IncludeHeaderFooter, PageDesign, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientReference, MailingInstructions, AttentionLine, Subject, CCList, ReturnAddress, SenderName, Closing, SenderCompany, SenderJobTitle, SenderInitials, EnclosureNumber, InfoBlock, RecipientCode, RecipientGender, ReturnAddressShortForm, SenderCity, SenderCode, SenderGender, SenderReference, .dispatch = as.integer(1), .ids =260, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'LetterContent')
},
'AcceptAllRevisions' = function(   .x){
	ans = .COM(.x, 'AcceptAllRevisions', .dispatch = as.integer(1), .ids =317, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RejectAllRevisions' = function(   .x){
	ans = .COM(.x, 'RejectAllRevisions', .dispatch = as.integer(1), .ids =318, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DetectLanguage' = function(   .x){
	ans = .COM(.x, 'DetectLanguage', .dispatch = as.integer(1), .ids =151, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ApplyTheme' = function( Name ,  .x){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	Name = as(Name, 'character')
	ans = .COM(.x, 'ApplyTheme', Name, .dispatch = as.integer(1), .ids =322, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RemoveTheme' = function(   .x){
	ans = .COM(.x, 'RemoveTheme', .dispatch = as.integer(1), .ids =323, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'WebPagePreview' = function(   .x){
	ans = .COM(.x, 'WebPagePreview', .dispatch = as.integer(1), .ids =325, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ReloadAs' = function( Encoding ,  .x){
	if( missing( Encoding ) ) {
	stop('You must specify a value for the argument(s)  Encoding ')
	}
	Encoding = as(Encoding, 'WdMailSystem')
	ans = .COM(.x, 'ReloadAs', Encoding, .dispatch = as.integer(1), .ids =331, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PrintOut2000' = function( Background = NA, Append = NA, Range = NA, OutputFileName = NA, From = NA, To = NA, Item = NA, Copies = NA, Pages = NA, PageType = NA, PrintToFile = NA, Collate = NA, ActivePrinterMacGX = NA, ManualDuplexPrint = NA, PrintZoomColumn = NA, PrintZoomRow = NA, PrintZoomPaperWidth = NA, PrintZoomPaperHeight = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'PrintOut2000', Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight, .dispatch = as.integer(1), .ids =444, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'sblt' = function( s ,  .x){
	if( missing( s ) ) {
	stop('You must specify a value for the argument(s)  s ')
	}
	s = as(s, 'character')
	ans = .COM(.x, 'sblt', s, .dispatch = as.integer(1), .ids =445, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ConvertVietDoc' = function( CodePageOrigin ,  .x){
	if( missing( CodePageOrigin ) ) {
	stop('You must specify a value for the argument(s)  CodePageOrigin ')
	}
	CodePageOrigin = as(CodePageOrigin, 'integer')
	ans = .COM(.x, 'ConvertVietDoc', CodePageOrigin, .dispatch = as.integer(1), .ids =447, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PrintOut' = function( Background = NA, Append = NA, Range = NA, OutputFileName = NA, From = NA, To = NA, Item = NA, Copies = NA, Pages = NA, PageType = NA, PrintToFile = NA, Collate = NA, ActivePrinterMacGX = NA, ManualDuplexPrint = NA, PrintZoomColumn = NA, PrintZoomRow = NA, PrintZoomPaperWidth = NA, PrintZoomPaperHeight = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'PrintOut', Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight, .dispatch = as.integer(1), .ids =446, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Compare2002' = function( Name, AuthorName = NA, CompareTarget = NA, DetectFormatChanges = NA, IgnoreAllComparisonWarnings = NA, AddToRecentFiles = NA ,  .x){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	Name = as(Name, 'character')
	
	
	
	
	
	ans = .COM(.x, 'Compare2002', Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, .dispatch = as.integer(1), .ids =345, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CheckIn' = function( SaveChanges = NA, Comments = NA, MakePublic = NA ,  .x){
	if(!missing( SaveChanges )) SaveChanges = as(SaveChanges, 'logical')
	
	if(!missing( MakePublic )) MakePublic = as(MakePublic, 'logical')
	ans = .COM(.x, 'CheckIn', SaveChanges, Comments, MakePublic, .dispatch = as.integer(1), .ids =349, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CanCheckin' = function(   .x){
	ans = .COM(.x, 'CanCheckin', .dispatch = as.integer(1), .ids =351, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Merge' = function( FileName, MergeTarget = NA, DetectFormatChanges = NA, UseFormattingFrom = NA, AddToRecentFiles = NA ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	FileName = as(FileName, 'character')
	
	
	
	
	ans = .COM(.x, 'Merge', FileName, MergeTarget, DetectFormatChanges, UseFormattingFrom, AddToRecentFiles, .dispatch = as.integer(1), .ids =362, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SendForReview' = function( Recipients = NA, Subject = NA, ShowMessage = NA, IncludeAttachment = NA ,  .x){
	
	
	
	
	ans = .COM(.x, 'SendForReview', Recipients, Subject, ShowMessage, IncludeAttachment, .dispatch = as.integer(1), .ids =353, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ReplyWithChanges' = function( ShowMessage = NA ,  .x){
	
	ans = .COM(.x, 'ReplyWithChanges', ShowMessage, .dispatch = as.integer(1), .ids =354, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'EndReview' = function(   .x){
	ans = .COM(.x, 'EndReview', .dispatch = as.integer(1), .ids =356, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SetPasswordEncryptionOptions' = function( PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties = NA ,  .x){
	if( missing( PasswordEncryptionProvider )||missing( PasswordEncryptionAlgorithm )||missing( PasswordEncryptionKeyLength ) ) {
	stop('You must specify a value for the argument(s)  PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength ')
	}
	PasswordEncryptionProvider = as(PasswordEncryptionProvider, 'character')
	PasswordEncryptionAlgorithm = as(PasswordEncryptionAlgorithm, 'character')
	PasswordEncryptionKeyLength = as(PasswordEncryptionKeyLength, 'integer')
	
	ans = .COM(.x, 'SetPasswordEncryptionOptions', PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties, .dispatch = as.integer(1), .ids =361, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RecheckSmartTags' = function(   .x){
	ans = .COM(.x, 'RecheckSmartTags', .dispatch = as.integer(1), .ids =363, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RemoveSmartTags' = function(   .x){
	ans = .COM(.x, 'RemoveSmartTags', .dispatch = as.integer(1), .ids =364, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SetDefaultTableStyle' = function( Style, SetInTemplate ,  .x){
	if( missing( Style )||missing( SetInTemplate ) ) {
	stop('You must specify a value for the argument(s)  Style, SetInTemplate ')
	}
	
	SetInTemplate = as(SetInTemplate, 'logical')
	ans = .COM(.x, 'SetDefaultTableStyle', Style, SetInTemplate, .dispatch = as.integer(1), .ids =366, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DeleteAllComments' = function(   .x){
	ans = .COM(.x, 'DeleteAllComments', .dispatch = as.integer(1), .ids =371, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AcceptAllRevisionsShown' = function(   .x){
	ans = .COM(.x, 'AcceptAllRevisionsShown', .dispatch = as.integer(1), .ids =372, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RejectAllRevisionsShown' = function(   .x){
	ans = .COM(.x, 'RejectAllRevisionsShown', .dispatch = as.integer(1), .ids =373, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DeleteAllCommentsShown' = function(   .x){
	ans = .COM(.x, 'DeleteAllCommentsShown', .dispatch = as.integer(1), .ids =374, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ResetFormFields' = function(   .x){
	ans = .COM(.x, 'ResetFormFields', .dispatch = as.integer(1), .ids =375, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SaveAs' = function( FileName = NA, FileFormat = NA, LockComments = NA, Password = NA, AddToRecentFiles = NA, WritePassword = NA, ReadOnlyRecommended = NA, EmbedTrueTypeFonts = NA, SaveNativePictureFormat = NA, SaveFormsData = NA, SaveAsAOCELetter = NA, Encoding = NA, InsertLineBreaks = NA, AllowSubstitutions = NA, LineEnding = NA, AddBiDiMarks = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'SaveAs', FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks, .dispatch = as.integer(1), .ids =376, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CheckNewSmartTags' = function(   .x){
	ans = .COM(.x, 'CheckNewSmartTags', .dispatch = as.integer(1), .ids =378, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SendFaxOverInternet' = function( Recipients = NA, Subject = NA, ShowMessage = NA ,  .x){
	
	
	
	ans = .COM(.x, 'SendFaxOverInternet', Recipients, Subject, ShowMessage, .dispatch = as.integer(1), .ids =464, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'TransformDocument' = function( Path, DataOnly = NA ,  .x){
	if( missing( Path ) ) {
	stop('You must specify a value for the argument(s)  Path ')
	}
	Path = as(Path, 'character')
	if(!missing( DataOnly )) DataOnly = as(DataOnly, 'logical')
	ans = .COM(.x, 'TransformDocument', Path, DataOnly, .dispatch = as.integer(1), .ids =500, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Protect' = function( Type, NoReset = NA, Password = NA, UseIRM = NA, EnforceStyleLock = NA ,  .x){
	if( missing( Type ) ) {
	stop('You must specify a value for the argument(s)  Type ')
	}
	Type = as(Type, 'WdMailSystem')
	
	
	
	
	ans = .COM(.x, 'Protect', Type, NoReset, Password, UseIRM, EnforceStyleLock, .dispatch = as.integer(1), .ids =467, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SelectAllEditableRanges' = function( EditorID = NA ,  .x){
	
	ans = .COM(.x, 'SelectAllEditableRanges', EditorID, .dispatch = as.integer(1), .ids =468, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DeleteAllEditableRanges' = function( EditorID = NA ,  .x){
	
	ans = .COM(.x, 'DeleteAllEditableRanges', EditorID, .dispatch = as.integer(1), .ids =469, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DeleteAllInkAnnotations' = function(   .x){
	ans = .COM(.x, 'DeleteAllInkAnnotations', .dispatch = as.integer(1), .ids =479, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AddDocumentWorkspaceHeader' = function( RichFormat, Url, Title, Description, ID ,  .x){
	if( missing( RichFormat )||missing( Url )||missing( Title )||missing( Description )||missing( ID ) ) {
	stop('You must specify a value for the argument(s)  RichFormat, Url, Title, Description, ID ')
	}
	RichFormat = as(RichFormat, 'logical')
	Url = as(Url, 'character')
	Title = as(Title, 'character')
	Description = as(Description, 'character')
	ID = as(ID, 'character')
	ans = .COM(.x, 'AddDocumentWorkspaceHeader', RichFormat, Url, Title, Description, ID, .dispatch = as.integer(1), .ids =482, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RemoveDocumentWorkspaceHeader' = function( ID ,  .x){
	if( missing( ID ) ) {
	stop('You must specify a value for the argument(s)  ID ')
	}
	ID = as(ID, 'character')
	ans = .COM(.x, 'RemoveDocumentWorkspaceHeader', ID, .dispatch = as.integer(1), .ids =483, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Compare' = function( Name, AuthorName = NA, CompareTarget = NA, DetectFormatChanges = NA, IgnoreAllComparisonWarnings = NA, AddToRecentFiles = NA, RemovePersonalInformation = NA, RemoveDateAndTime = NA ,  .x){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	Name = as(Name, 'character')
	
	
	
	
	
	
	
	ans = .COM(.x, 'Compare', Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation, RemoveDateAndTime, .dispatch = as.integer(1), .ids =485, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'RemoveLockedStyles' = function(   .x){
	ans = .COM(.x, 'RemoveLockedStyles', .dispatch = as.integer(1), .ids =487, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SelectSingleNode' = function( XPath, PrefixMapping = NA, FastSearchSkippingTextNodes = NA ,  .x){
	if( missing( XPath ) ) {
	stop('You must specify a value for the argument(s)  XPath ')
	}
	XPath = as(XPath, 'character')
	if(!missing( PrefixMapping )) PrefixMapping = as(PrefixMapping, 'character')
	if(!missing( FastSearchSkippingTextNodes )) FastSearchSkippingTextNodes = as(FastSearchSkippingTextNodes, 'logical')
	ans = .COM(.x, 'SelectSingleNode', XPath, PrefixMapping, FastSearchSkippingTextNodes, .dispatch = as.integer(1), .ids =488, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'XMLNode')
},
'SelectNodes' = function( XPath, PrefixMapping = NA, FastSearchSkippingTextNodes = NA ,  .x){
	if( missing( XPath ) ) {
	stop('You must specify a value for the argument(s)  XPath ')
	}
	XPath = as(XPath, 'character')
	if(!missing( PrefixMapping )) PrefixMapping = as(PrefixMapping, 'character')
	if(!missing( FastSearchSkippingTextNodes )) FastSearchSkippingTextNodes = as(FastSearchSkippingTextNodes, 'logical')
	ans = .COM(.x, 'SelectNodes', XPath, PrefixMapping, FastSearchSkippingTextNodes, .dispatch = as.integer(1), .ids =489, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'XMLNodes')
} )
'COM.Documents.GetProperty'  = list('_NewEnum' = function(x) {
		 ans = .COM(x, '_NewEnum', .dispatch = as.integer(2))
		 ans
},
'Count' = function(x) {
		 ans = .COM(x, 'Count', .dispatch = as.integer(2))
	ans
},
'Application' = function(x) {
		 ans = .COM(x, 'Application', .dispatch = as.integer(2))
		 as(ans, 'Application')
},
'Creator' = function(x) {
		 ans = .COM(x, 'Creator', .dispatch = as.integer(2))
	ans
},
'Parent' = function(x) {
		 ans = .COM(x, 'Parent', .dispatch = as.integer(2))
		 ans
} )
'COM.Documents.SetProperty'  = list( )
'COM.Documents.Methods'  = list('Item' = function( Index ,  .x){
	if( missing( Index ) ) {
	stop('You must specify a value for the argument(s)  Index ')
	}
	
	ans = .COM(.x, 'Item', Index, .dispatch = as.integer(1), .ids =0, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Document')
},
'Close' = function( SaveChanges = NA, OriginalFormat = NA, RouteDocument = NA ,  .x){
	
	
	
	ans = .COM(.x, 'Close', SaveChanges, OriginalFormat, RouteDocument, .dispatch = as.integer(1), .ids =1105, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AddOld' = function( Template = NA, NewTemplate = NA ,  .x){
	
	
	ans = .COM(.x, 'AddOld', Template, NewTemplate, .dispatch = as.integer(1), .ids =11, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Document')
},
'OpenOld' = function( FileName, ConfirmConversions = NA, ReadOnly = NA, AddToRecentFiles = NA, PasswordDocument = NA, PasswordTemplate = NA, Revert = NA, WritePasswordDocument = NA, WritePasswordTemplate = NA, Format = NA ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'OpenOld', FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, .dispatch = as.integer(1), .ids =12, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Document')
},
'Save' = function( NoPrompt = NA, OriginalFormat = NA ,  .x){
	
	
	ans = .COM(.x, 'Save', NoPrompt, OriginalFormat, .dispatch = as.integer(1), .ids =13, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Add' = function( Template = NA, NewTemplate = NA, DocumentType = NA, Visible = NA ,  .x){
	
	
	
	
	ans = .COM(.x, 'Add', Template, NewTemplate, DocumentType, Visible, .dispatch = as.integer(1), .ids =14, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Document')
},
'Open2000' = function( FileName, ConfirmConversions = NA, ReadOnly = NA, AddToRecentFiles = NA, PasswordDocument = NA, PasswordTemplate = NA, Revert = NA, WritePasswordDocument = NA, WritePasswordTemplate = NA, Format = NA, Encoding = NA, Visible = NA ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'Open2000', FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, .dispatch = as.integer(1), .ids =15, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Document')
},
'CheckOut' = function( FileName ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	FileName = as(FileName, 'character')
	ans = .COM(.x, 'CheckOut', FileName, .dispatch = as.integer(1), .ids =16, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CanCheckOut' = function( FileName ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	FileName = as(FileName, 'character')
	ans = .COM(.x, 'CanCheckOut', FileName, .dispatch = as.integer(1), .ids =17, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Open2002' = function( FileName, ConfirmConversions = NA, ReadOnly = NA, AddToRecentFiles = NA, PasswordDocument = NA, PasswordTemplate = NA, Revert = NA, WritePasswordDocument = NA, WritePasswordTemplate = NA, Format = NA, Encoding = NA, Visible = NA, OpenAndRepair = NA, DocumentDirection = NA, NoEncodingDialog = NA ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'Open2002', FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, .dispatch = as.integer(1), .ids =18, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Document')
},
'Open' = function( FileName, ConfirmConversions = NA, ReadOnly = NA, AddToRecentFiles = NA, PasswordDocument = NA, PasswordTemplate = NA, Revert = NA, WritePasswordDocument = NA, WritePasswordTemplate = NA, Format = NA, Encoding = NA, Visible = NA, OpenAndRepair = NA, DocumentDirection = NA, NoEncodingDialog = NA, XMLTransform = NA ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'Open', FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform, .dispatch = as.integer(1), .ids =19, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Document')
} )
'COM.Range.GetProperty'  = list('Text' = function(x) {
		 ans = .COM(x, 'Text', .dispatch = as.integer(2))
	ans
},
'FormattedText' = function(x) {
		 ans = .COM(x, 'FormattedText', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'Start' = function(x) {
		 ans = .COM(x, 'Start', .dispatch = as.integer(2))
	ans
},
'End' = function(x) {
		 ans = .COM(x, 'End', .dispatch = as.integer(2))
	ans
},
'Font' = function(x) {
		 ans = .COM(x, 'Font', .dispatch = as.integer(2))
		 as(ans, 'Font')
},
'Duplicate' = function(x) {
		 ans = .COM(x, 'Duplicate', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'StoryType' = function(x) {
		 ans = .COM(x, 'StoryType', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'Tables' = function(x) {
		 ans = .COM(x, 'Tables', .dispatch = as.integer(2))
		 as(ans, 'Tables')
},
'Words' = function(x) {
		 ans = .COM(x, 'Words', .dispatch = as.integer(2))
		 as(ans, 'Words')
},
'Sentences' = function(x) {
		 ans = .COM(x, 'Sentences', .dispatch = as.integer(2))
		 as(ans, 'Sentences')
},
'Characters' = function(x) {
		 ans = .COM(x, 'Characters', .dispatch = as.integer(2))
		 as(ans, 'Characters')
},
'Footnotes' = function(x) {
		 ans = .COM(x, 'Footnotes', .dispatch = as.integer(2))
		 as(ans, 'Footnotes')
},
'Endnotes' = function(x) {
		 ans = .COM(x, 'Endnotes', .dispatch = as.integer(2))
		 as(ans, 'Endnotes')
},
'Comments' = function(x) {
		 ans = .COM(x, 'Comments', .dispatch = as.integer(2))
		 as(ans, 'Comments')
},
'Cells' = function(x) {
		 ans = .COM(x, 'Cells', .dispatch = as.integer(2))
		 as(ans, 'Cells')
},
'Sections' = function(x) {
		 ans = .COM(x, 'Sections', .dispatch = as.integer(2))
		 as(ans, 'Sections')
},
'Paragraphs' = function(x) {
		 ans = .COM(x, 'Paragraphs', .dispatch = as.integer(2))
		 as(ans, 'Paragraphs')
},
'Borders' = function(x) {
		 ans = .COM(x, 'Borders', .dispatch = as.integer(2))
		 as(ans, 'Borders')
},
'Shading' = function(x) {
		 ans = .COM(x, 'Shading', .dispatch = as.integer(2))
		 as(ans, 'Shading')
},
'TextRetrievalMode' = function(x) {
		 ans = .COM(x, 'TextRetrievalMode', .dispatch = as.integer(2))
		 as(ans, 'TextRetrievalMode')
},
'Fields' = function(x) {
		 ans = .COM(x, 'Fields', .dispatch = as.integer(2))
		 as(ans, 'Fields')
},
'FormFields' = function(x) {
		 ans = .COM(x, 'FormFields', .dispatch = as.integer(2))
		 as(ans, 'FormFields')
},
'Frames' = function(x) {
		 ans = .COM(x, 'Frames', .dispatch = as.integer(2))
		 as(ans, 'Frames')
},
'ParagraphFormat' = function(x) {
		 ans = .COM(x, 'ParagraphFormat', .dispatch = as.integer(2))
		 as(ans, 'ParagraphFormat')
},
'ListFormat' = function(x) {
		 ans = .COM(x, 'ListFormat', .dispatch = as.integer(2))
		 as(ans, 'ListFormat')
},
'Bookmarks' = function(x) {
		 ans = .COM(x, 'Bookmarks', .dispatch = as.integer(2))
		 as(ans, 'Bookmarks')
},
'Application' = function(x) {
		 ans = .COM(x, 'Application', .dispatch = as.integer(2))
		 as(ans, 'Application')
},
'Creator' = function(x) {
		 ans = .COM(x, 'Creator', .dispatch = as.integer(2))
	ans
},
'Parent' = function(x) {
		 ans = .COM(x, 'Parent', .dispatch = as.integer(2))
		 ans
},
'Bold' = function(x) {
		 ans = .COM(x, 'Bold', .dispatch = as.integer(2))
	ans
},
'Italic' = function(x) {
		 ans = .COM(x, 'Italic', .dispatch = as.integer(2))
	ans
},
'Underline' = function(x) {
		 ans = .COM(x, 'Underline', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'EmphasisMark' = function(x) {
		 ans = .COM(x, 'EmphasisMark', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'DisableCharacterSpaceGrid' = function(x) {
		 ans = .COM(x, 'DisableCharacterSpaceGrid', .dispatch = as.integer(2))
	ans
},
'Revisions' = function(x) {
		 ans = .COM(x, 'Revisions', .dispatch = as.integer(2))
		 as(ans, 'Revisions')
},
'Style' = function(x) {
		 ans = .COM(x, 'Style', .dispatch = as.integer(2))
		 ans
},
'StoryLength' = function(x) {
		 ans = .COM(x, 'StoryLength', .dispatch = as.integer(2))
	ans
},
'LanguageID' = function(x) {
		 ans = .COM(x, 'LanguageID', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'SynonymInfo' = function(x) {
		 ans = .COM(x, 'SynonymInfo', .dispatch = as.integer(2))
		 as(ans, 'SynonymInfo')
},
'Hyperlinks' = function(x) {
		 ans = .COM(x, 'Hyperlinks', .dispatch = as.integer(2))
		 as(ans, 'Hyperlinks')
},
'ListParagraphs' = function(x) {
		 ans = .COM(x, 'ListParagraphs', .dispatch = as.integer(2))
		 as(ans, 'ListParagraphs')
},
'Subdocuments' = function(x) {
		 ans = .COM(x, 'Subdocuments', .dispatch = as.integer(2))
		 as(ans, 'Subdocuments')
},
'GrammarChecked' = function(x) {
		 ans = .COM(x, 'GrammarChecked', .dispatch = as.integer(2))
	ans
},
'SpellingChecked' = function(x) {
		 ans = .COM(x, 'SpellingChecked', .dispatch = as.integer(2))
	ans
},
'HighlightColorIndex' = function(x) {
		 ans = .COM(x, 'HighlightColorIndex', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'Columns' = function(x) {
		 ans = .COM(x, 'Columns', .dispatch = as.integer(2))
		 as(ans, 'Columns')
},
'Rows' = function(x) {
		 ans = .COM(x, 'Rows', .dispatch = as.integer(2))
		 as(ans, 'Rows')
},
'CanEdit' = function(x) {
		 ans = .COM(x, 'CanEdit', .dispatch = as.integer(2))
	ans
},
'CanPaste' = function(x) {
		 ans = .COM(x, 'CanPaste', .dispatch = as.integer(2))
	ans
},
'IsEndOfRowMark' = function(x) {
		 ans = .COM(x, 'IsEndOfRowMark', .dispatch = as.integer(2))
	ans
},
'BookmarkID' = function(x) {
		 ans = .COM(x, 'BookmarkID', .dispatch = as.integer(2))
	ans
},
'PreviousBookmarkID' = function(x) {
		 ans = .COM(x, 'PreviousBookmarkID', .dispatch = as.integer(2))
	ans
},
'Find' = function(x) {
		 ans = .COM(x, 'Find', .dispatch = as.integer(2))
		 as(ans, 'Find')
},
'PageSetup' = function(x) {
		 ans = .COM(x, 'PageSetup', .dispatch = as.integer(2))
		 as(ans, 'PageSetup')
},
'ShapeRange' = function(x) {
		 ans = .COM(x, 'ShapeRange', .dispatch = as.integer(2))
		 as(ans, 'ShapeRange')
},
'Case' = function(x) {
		 ans = .COM(x, 'Case', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'ReadabilityStatistics' = function(x) {
		 ans = .COM(x, 'ReadabilityStatistics', .dispatch = as.integer(2))
		 as(ans, 'ReadabilityStatistics')
},
'GrammaticalErrors' = function(x) {
		 ans = .COM(x, 'GrammaticalErrors', .dispatch = as.integer(2))
		 as(ans, 'ProofreadingErrors')
},
'SpellingErrors' = function(x) {
		 ans = .COM(x, 'SpellingErrors', .dispatch = as.integer(2))
		 as(ans, 'ProofreadingErrors')
},
'Orientation' = function(x) {
		 ans = .COM(x, 'Orientation', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'InlineShapes' = function(x) {
		 ans = .COM(x, 'InlineShapes', .dispatch = as.integer(2))
		 as(ans, 'InlineShapes')
},
'NextStoryRange' = function(x) {
		 ans = .COM(x, 'NextStoryRange', .dispatch = as.integer(2))
		 as(ans, 'Range')
},
'LanguageIDFarEast' = function(x) {
		 ans = .COM(x, 'LanguageIDFarEast', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'LanguageIDOther' = function(x) {
		 ans = .COM(x, 'LanguageIDOther', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'LanguageDetected' = function(x) {
		 ans = .COM(x, 'LanguageDetected', .dispatch = as.integer(2))
	ans
},
'FitTextWidth' = function(x) {
		 ans = .COM(x, 'FitTextWidth', .dispatch = as.integer(2))
		 ans
},
'HorizontalInVertical' = function(x) {
		 ans = .COM(x, 'HorizontalInVertical', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'TwoLinesInOne' = function(x) {
		 ans = .COM(x, 'TwoLinesInOne', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'CombineCharacters' = function(x) {
		 ans = .COM(x, 'CombineCharacters', .dispatch = as.integer(2))
	ans
},
'NoProofing' = function(x) {
		 ans = .COM(x, 'NoProofing', .dispatch = as.integer(2))
	ans
},
'TopLevelTables' = function(x) {
		 ans = .COM(x, 'TopLevelTables', .dispatch = as.integer(2))
		 as(ans, 'Tables')
},
'Scripts' = function(x) {
		 ans = .COM(x, 'Scripts', .dispatch = as.integer(2))
		 as(ans, 'COMIDispatch')
},
'CharacterWidth' = function(x) {
		 ans = .COM(x, 'CharacterWidth', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'Kana' = function(x) {
		 ans = .COM(x, 'Kana', .dispatch = as.integer(2))
		 as(ans, 'WdMailSystem')
},
'BoldBi' = function(x) {
		 ans = .COM(x, 'BoldBi', .dispatch = as.integer(2))
	ans
},
'ItalicBi' = function(x) {
		 ans = .COM(x, 'ItalicBi', .dispatch = as.integer(2))
	ans
},
'ID' = function(x) {
		 ans = .COM(x, 'ID', .dispatch = as.integer(2))
	ans
},
'HTMLDivisions' = function(x) {
		 ans = .COM(x, 'HTMLDivisions', .dispatch = as.integer(2))
		 as(ans, 'HTMLDivisions')
},
'SmartTags' = function(x) {
		 ans = .COM(x, 'SmartTags', .dispatch = as.integer(2))
		 as(ans, 'SmartTags')
},
'ShowAll' = function(x) {
		 ans = .COM(x, 'ShowAll', .dispatch = as.integer(2))
	ans
},
'Document' = function(x) {
		 ans = .COM(x, 'Document', .dispatch = as.integer(2))
		 as(ans, 'Document')
},
'FootnoteOptions' = function(x) {
		 ans = .COM(x, 'FootnoteOptions', .dispatch = as.integer(2))
		 as(ans, 'FootnoteOptions')
},
'EndnoteOptions' = function(x) {
		 ans = .COM(x, 'EndnoteOptions', .dispatch = as.integer(2))
		 as(ans, 'EndnoteOptions')
},
'XMLNodes' = function(x) {
		 ans = .COM(x, 'XMLNodes', .dispatch = as.integer(2))
		 as(ans, 'XMLNodes')
},
'XMLParentNode' = function(x) {
		 ans = .COM(x, 'XMLParentNode', .dispatch = as.integer(2))
		 as(ans, 'XMLNode')
},
'Editors' = function(x) {
		 ans = .COM(x, 'Editors', .dispatch = as.integer(2))
		 as(ans, 'Editors')
},
'EnhMetaFileBits' = function(x) {
		 ans = .COM(x, 'EnhMetaFileBits', .dispatch = as.integer(2))
		 ans
} )
'COM.Range.SetProperty'  = list('Text' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'Text', value, .dispatch = as.integer(4))
},
'FormattedText' = function(x, value) {
	value = as(value, 'Range')
	.COM(x, 'FormattedText', value, .dispatch = as.integer(4))
},
'Start' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'Start', value, .dispatch = as.integer(4))
},
'End' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'End', value, .dispatch = as.integer(4))
},
'Font' = function(x, value) {
	value = as(value, 'Font')
	.COM(x, 'Font', value, .dispatch = as.integer(4))
},
'Borders' = function(x, value) {
	value = as(value, 'Borders')
	.COM(x, 'Borders', value, .dispatch = as.integer(4))
},
'TextRetrievalMode' = function(x, value) {
	value = as(value, 'TextRetrievalMode')
	.COM(x, 'TextRetrievalMode', value, .dispatch = as.integer(4))
},
'ParagraphFormat' = function(x, value) {
	value = as(value, 'ParagraphFormat')
	.COM(x, 'ParagraphFormat', value, .dispatch = as.integer(4))
},
'Bold' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'Bold', value, .dispatch = as.integer(4))
},
'Italic' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'Italic', value, .dispatch = as.integer(4))
},
'Underline' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'Underline', value, .dispatch = as.integer(4))
},
'EmphasisMark' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'EmphasisMark', value, .dispatch = as.integer(4))
},
'DisableCharacterSpaceGrid' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'DisableCharacterSpaceGrid', value, .dispatch = as.integer(4))
},
'Style' = function(x, value) {
	
	.COM(x, 'Style', value, .dispatch = as.integer(4))
},
'LanguageID' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'LanguageID', value, .dispatch = as.integer(4))
},
'GrammarChecked' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'GrammarChecked', value, .dispatch = as.integer(4))
},
'SpellingChecked' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'SpellingChecked', value, .dispatch = as.integer(4))
},
'HighlightColorIndex' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'HighlightColorIndex', value, .dispatch = as.integer(4))
},
'PageSetup' = function(x, value) {
	value = as(value, 'PageSetup')
	.COM(x, 'PageSetup', value, .dispatch = as.integer(4))
},
'Case' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'Case', value, .dispatch = as.integer(4))
},
'Orientation' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'Orientation', value, .dispatch = as.integer(4))
},
'LanguageIDFarEast' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'LanguageIDFarEast', value, .dispatch = as.integer(4))
},
'LanguageIDOther' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'LanguageIDOther', value, .dispatch = as.integer(4))
},
'LanguageDetected' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'LanguageDetected', value, .dispatch = as.integer(4))
},
'FitTextWidth' = function(x, value) {
	
	.COM(x, 'FitTextWidth', value, .dispatch = as.integer(4))
},
'HorizontalInVertical' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'HorizontalInVertical', value, .dispatch = as.integer(4))
},
'TwoLinesInOne' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'TwoLinesInOne', value, .dispatch = as.integer(4))
},
'CombineCharacters' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'CombineCharacters', value, .dispatch = as.integer(4))
},
'NoProofing' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'NoProofing', value, .dispatch = as.integer(4))
},
'CharacterWidth' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'CharacterWidth', value, .dispatch = as.integer(4))
},
'Kana' = function(x, value) {
	value = as(value, 'WdMailSystem')
	.COM(x, 'Kana', value, .dispatch = as.integer(4))
},
'BoldBi' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'BoldBi', value, .dispatch = as.integer(4))
},
'ItalicBi' = function(x, value) {
	value = as(value, 'integer')
	.COM(x, 'ItalicBi', value, .dispatch = as.integer(4))
},
'ID' = function(x, value) {
	value = as(value, 'character')
	.COM(x, 'ID', value, .dispatch = as.integer(4))
},
'ShowAll' = function(x, value) {
	value = as(value, 'logical')
	.COM(x, 'ShowAll', value, .dispatch = as.integer(4))
} )
'COM.Range.Methods'  = list('Select' = function(   .x){
	ans = .COM(.x, 'Select', .dispatch = as.integer(1), .ids =65535, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SetRange' = function( Start, End ,  .x){
	if( missing( Start )||missing( End ) ) {
	stop('You must specify a value for the argument(s)  Start, End ')
	}
	Start = as(Start, 'integer')
	End = as(End, 'integer')
	ans = .COM(.x, 'SetRange', Start, End, .dispatch = as.integer(1), .ids =100, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Collapse' = function( Direction = NA ,  .x){
	
	ans = .COM(.x, 'Collapse', Direction, .dispatch = as.integer(1), .ids =101, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertBefore' = function( Text ,  .x){
	if( missing( Text ) ) {
	stop('You must specify a value for the argument(s)  Text ')
	}
	Text = as(Text, 'character')
	ans = .COM(.x, 'InsertBefore', Text, .dispatch = as.integer(1), .ids =102, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertAfter' = function( Text ,  .x){
	if( missing( Text ) ) {
	stop('You must specify a value for the argument(s)  Text ')
	}
	Text = as(Text, 'character')
	ans = .COM(.x, 'InsertAfter', Text, .dispatch = as.integer(1), .ids =104, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Next' = function( Unit = NA, Count = NA ,  .x){
	
	
	ans = .COM(.x, 'Next', Unit, Count, .dispatch = as.integer(1), .ids =105, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'Previous' = function( Unit = NA, Count = NA ,  .x){
	
	
	ans = .COM(.x, 'Previous', Unit, Count, .dispatch = as.integer(1), .ids =106, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'StartOf' = function( Unit = NA, Extend = NA ,  .x){
	
	
	ans = .COM(.x, 'StartOf', Unit, Extend, .dispatch = as.integer(1), .ids =107, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'EndOf' = function( Unit = NA, Extend = NA ,  .x){
	
	
	ans = .COM(.x, 'EndOf', Unit, Extend, .dispatch = as.integer(1), .ids =108, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Move' = function( Unit = NA, Count = NA ,  .x){
	
	
	ans = .COM(.x, 'Move', Unit, Count, .dispatch = as.integer(1), .ids =109, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MoveStart' = function( Unit = NA, Count = NA ,  .x){
	
	
	ans = .COM(.x, 'MoveStart', Unit, Count, .dispatch = as.integer(1), .ids =110, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MoveEnd' = function( Unit = NA, Count = NA ,  .x){
	
	
	ans = .COM(.x, 'MoveEnd', Unit, Count, .dispatch = as.integer(1), .ids =111, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MoveWhile' = function( Cset, Count = NA ,  .x){
	if( missing( Cset ) ) {
	stop('You must specify a value for the argument(s)  Cset ')
	}
	
	
	ans = .COM(.x, 'MoveWhile', Cset, Count, .dispatch = as.integer(1), .ids =112, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MoveStartWhile' = function( Cset, Count = NA ,  .x){
	if( missing( Cset ) ) {
	stop('You must specify a value for the argument(s)  Cset ')
	}
	
	
	ans = .COM(.x, 'MoveStartWhile', Cset, Count, .dispatch = as.integer(1), .ids =113, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MoveEndWhile' = function( Cset, Count = NA ,  .x){
	if( missing( Cset ) ) {
	stop('You must specify a value for the argument(s)  Cset ')
	}
	
	
	ans = .COM(.x, 'MoveEndWhile', Cset, Count, .dispatch = as.integer(1), .ids =114, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MoveUntil' = function( Cset, Count = NA ,  .x){
	if( missing( Cset ) ) {
	stop('You must specify a value for the argument(s)  Cset ')
	}
	
	
	ans = .COM(.x, 'MoveUntil', Cset, Count, .dispatch = as.integer(1), .ids =115, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MoveStartUntil' = function( Cset, Count = NA ,  .x){
	if( missing( Cset ) ) {
	stop('You must specify a value for the argument(s)  Cset ')
	}
	
	
	ans = .COM(.x, 'MoveStartUntil', Cset, Count, .dispatch = as.integer(1), .ids =116, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'MoveEndUntil' = function( Cset, Count = NA ,  .x){
	if( missing( Cset ) ) {
	stop('You must specify a value for the argument(s)  Cset ')
	}
	
	
	ans = .COM(.x, 'MoveEndUntil', Cset, Count, .dispatch = as.integer(1), .ids =117, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Cut' = function(   .x){
	ans = .COM(.x, 'Cut', .dispatch = as.integer(1), .ids =119, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Copy' = function(   .x){
	ans = .COM(.x, 'Copy', .dispatch = as.integer(1), .ids =120, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Paste' = function(   .x){
	ans = .COM(.x, 'Paste', .dispatch = as.integer(1), .ids =121, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertBreak' = function( Type = NA ,  .x){
	
	ans = .COM(.x, 'InsertBreak', Type, .dispatch = as.integer(1), .ids =122, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertFile' = function( FileName, Range = NA, ConfirmConversions = NA, Link = NA, Attachment = NA ,  .x){
	if( missing( FileName ) ) {
	stop('You must specify a value for the argument(s)  FileName ')
	}
	FileName = as(FileName, 'character')
	
	
	
	
	ans = .COM(.x, 'InsertFile', FileName, Range, ConfirmConversions, Link, Attachment, .dispatch = as.integer(1), .ids =123, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InStory' = function( Range ,  .x){
	if( missing( Range ) ) {
	stop('You must specify a value for the argument(s)  Range ')
	}
	Range = as(Range, 'Range')
	ans = .COM(.x, 'InStory', Range, .dispatch = as.integer(1), .ids =125, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'InRange' = function( Range ,  .x){
	if( missing( Range ) ) {
	stop('You must specify a value for the argument(s)  Range ')
	}
	Range = as(Range, 'Range')
	ans = .COM(.x, 'InRange', Range, .dispatch = as.integer(1), .ids =126, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Delete' = function( Unit = NA, Count = NA ,  .x){
	
	
	ans = .COM(.x, 'Delete', Unit, Count, .dispatch = as.integer(1), .ids =127, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'WholeStory' = function(   .x){
	ans = .COM(.x, 'WholeStory', .dispatch = as.integer(1), .ids =128, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Expand' = function( Unit = NA ,  .x){
	
	ans = .COM(.x, 'Expand', Unit, .dispatch = as.integer(1), .ids =129, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'InsertParagraph' = function(   .x){
	ans = .COM(.x, 'InsertParagraph', .dispatch = as.integer(1), .ids =160, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertParagraphAfter' = function(   .x){
	ans = .COM(.x, 'InsertParagraphAfter', .dispatch = as.integer(1), .ids =161, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ConvertToTableOld' = function( Separator = NA, NumRows = NA, NumColumns = NA, InitialColumnWidth = NA, Format = NA, ApplyBorders = NA, ApplyShading = NA, ApplyFont = NA, ApplyColor = NA, ApplyHeadingRows = NA, ApplyLastRow = NA, ApplyFirstColumn = NA, ApplyLastColumn = NA, AutoFit = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'ConvertToTableOld', Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, .dispatch = as.integer(1), .ids =162, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Table')
},
'InsertDateTimeOld' = function( DateTimeFormat = NA, InsertAsField = NA, InsertAsFullWidth = NA ,  .x){
	
	
	
	ans = .COM(.x, 'InsertDateTimeOld', DateTimeFormat, InsertAsField, InsertAsFullWidth, .dispatch = as.integer(1), .ids =163, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertSymbol' = function( CharacterNumber, Font = NA, Unicode = NA, Bias = NA ,  .x){
	if( missing( CharacterNumber ) ) {
	stop('You must specify a value for the argument(s)  CharacterNumber ')
	}
	CharacterNumber = as(CharacterNumber, 'integer')
	
	
	
	ans = .COM(.x, 'InsertSymbol', CharacterNumber, Font, Unicode, Bias, .dispatch = as.integer(1), .ids =164, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertCrossReference_2002' = function( ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink = NA, IncludePosition = NA ,  .x){
	if( missing( ReferenceType )||missing( ReferenceKind )||missing( ReferenceItem ) ) {
	stop('You must specify a value for the argument(s)  ReferenceType, ReferenceKind, ReferenceItem ')
	}
	
	ReferenceKind = as(ReferenceKind, 'WdMailSystem')
	
	
	
	ans = .COM(.x, 'InsertCrossReference_2002', ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, .dispatch = as.integer(1), .ids =165, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertCaptionXP' = function( Label, Title = NA, TitleAutoText = NA, Position = NA ,  .x){
	if( missing( Label ) ) {
	stop('You must specify a value for the argument(s)  Label ')
	}
	
	
	
	
	ans = .COM(.x, 'InsertCaptionXP', Label, Title, TitleAutoText, Position, .dispatch = as.integer(1), .ids =166, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CopyAsPicture' = function(   .x){
	ans = .COM(.x, 'CopyAsPicture', .dispatch = as.integer(1), .ids =167, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SortOld' = function( ExcludeHeader = NA, FieldNumber = NA, SortFieldType = NA, SortOrder = NA, FieldNumber2 = NA, SortFieldType2 = NA, SortOrder2 = NA, FieldNumber3 = NA, SortFieldType3 = NA, SortOrder3 = NA, SortColumn = NA, Separator = NA, CaseSensitive = NA, LanguageID = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'SortOld', ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, LanguageID, .dispatch = as.integer(1), .ids =168, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SortAscending' = function(   .x){
	ans = .COM(.x, 'SortAscending', .dispatch = as.integer(1), .ids =169, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SortDescending' = function(   .x){
	ans = .COM(.x, 'SortDescending', .dispatch = as.integer(1), .ids =170, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'IsEqual' = function( Range ,  .x){
	if( missing( Range ) ) {
	stop('You must specify a value for the argument(s)  Range ')
	}
	Range = as(Range, 'Range')
	ans = .COM(.x, 'IsEqual', Range, .dispatch = as.integer(1), .ids =171, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Calculate' = function(   .x){
	ans = .COM(.x, 'Calculate', .dispatch = as.integer(1), .ids =172, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GoTo' = function( What = NA, Which = NA, Count = NA, Name = NA ,  .x){
	
	
	
	
	ans = .COM(.x, 'GoTo', What, Which, Count, Name, .dispatch = as.integer(1), .ids =173, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'GoToNext' = function( What ,  .x){
	if( missing( What ) ) {
	stop('You must specify a value for the argument(s)  What ')
	}
	What = as(What, 'WdMailSystem')
	ans = .COM(.x, 'GoToNext', What, .dispatch = as.integer(1), .ids =174, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'GoToPrevious' = function( What ,  .x){
	if( missing( What ) ) {
	stop('You must specify a value for the argument(s)  What ')
	}
	What = as(What, 'WdMailSystem')
	ans = .COM(.x, 'GoToPrevious', What, .dispatch = as.integer(1), .ids =175, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'PasteSpecial' = function( IconIndex = NA, Link = NA, Placement = NA, DisplayAsIcon = NA, DataType = NA, IconFileName = NA, IconLabel = NA ,  .x){
	
	
	
	
	
	
	
	ans = .COM(.x, 'PasteSpecial', IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel, .dispatch = as.integer(1), .ids =176, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'LookupNameProperties' = function(   .x){
	ans = .COM(.x, 'LookupNameProperties', .dispatch = as.integer(1), .ids =177, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ComputeStatistics' = function( Statistic ,  .x){
	if( missing( Statistic ) ) {
	stop('You must specify a value for the argument(s)  Statistic ')
	}
	Statistic = as(Statistic, 'WdMailSystem')
	ans = .COM(.x, 'ComputeStatistics', Statistic, .dispatch = as.integer(1), .ids =178, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
	ans
},
'Relocate' = function( Direction ,  .x){
	if( missing( Direction ) ) {
	stop('You must specify a value for the argument(s)  Direction ')
	}
	Direction = as(Direction, 'integer')
	ans = .COM(.x, 'Relocate', Direction, .dispatch = as.integer(1), .ids =179, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CheckSynonyms' = function(   .x){
	ans = .COM(.x, 'CheckSynonyms', .dispatch = as.integer(1), .ids =180, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'SubscribeTo' = function( Edition, Format = NA ,  .x){
	if( missing( Edition ) ) {
	stop('You must specify a value for the argument(s)  Edition ')
	}
	Edition = as(Edition, 'character')
	
	ans = .COM(.x, 'SubscribeTo', Edition, Format, .dispatch = as.integer(1), .ids =181, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CreatePublisher' = function( Edition = NA, ContainsPICT = NA, ContainsRTF = NA, ContainsText = NA ,  .x){
	
	
	
	
	ans = .COM(.x, 'CreatePublisher', Edition, ContainsPICT, ContainsRTF, ContainsText, .dispatch = as.integer(1), .ids =182, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertAutoText' = function(   .x){
	ans = .COM(.x, 'InsertAutoText', .dispatch = as.integer(1), .ids =183, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertDatabase' = function( Format = NA, Style = NA, LinkToSource = NA, Connection = NA, SQLStatement = NA, SQLStatement1 = NA, PasswordDocument = NA, PasswordTemplate = NA, WritePasswordDocument = NA, WritePasswordTemplate = NA, DataSource = NA, From = NA, To = NA, IncludeFields = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'InsertDatabase', Format, Style, LinkToSource, Connection, SQLStatement, SQLStatement1, PasswordDocument, PasswordTemplate, WritePasswordDocument, WritePasswordTemplate, DataSource, From, To, IncludeFields, .dispatch = as.integer(1), .ids =194, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'AutoFormat' = function(   .x){
	ans = .COM(.x, 'AutoFormat', .dispatch = as.integer(1), .ids =195, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CheckGrammar' = function(   .x){
	ans = .COM(.x, 'CheckGrammar', .dispatch = as.integer(1), .ids =204, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'CheckSpelling' = function( CustomDictionary = NA, IgnoreUppercase = NA, AlwaysSuggest = NA, CustomDictionary2 = NA, CustomDictionary3 = NA, CustomDictionary4 = NA, CustomDictionary5 = NA, CustomDictionary6 = NA, CustomDictionary7 = NA, CustomDictionary8 = NA, CustomDictionary9 = NA, CustomDictionary10 = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'CheckSpelling', CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10, .dispatch = as.integer(1), .ids =205, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GetSpellingSuggestions' = function( CustomDictionary = NA, IgnoreUppercase = NA, MainDictionary = NA, SuggestionMode = NA, CustomDictionary2 = NA, CustomDictionary3 = NA, CustomDictionary4 = NA, CustomDictionary5 = NA, CustomDictionary6 = NA, CustomDictionary7 = NA, CustomDictionary8 = NA, CustomDictionary9 = NA, CustomDictionary10 = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'GetSpellingSuggestions', CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10, .dispatch = as.integer(1), .ids =209, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'SpellingSuggestions')
},
'InsertParagraphBefore' = function(   .x){
	ans = .COM(.x, 'InsertParagraphBefore', .dispatch = as.integer(1), .ids =212, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'NextSubdocument' = function(   .x){
	ans = .COM(.x, 'NextSubdocument', .dispatch = as.integer(1), .ids =219, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PreviousSubdocument' = function(   .x){
	ans = .COM(.x, 'PreviousSubdocument', .dispatch = as.integer(1), .ids =220, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ConvertHangulAndHanja' = function( ConversionsMode = NA, FastConversion = NA, CheckHangulEnding = NA, EnableRecentOrdering = NA, CustomDictionary = NA ,  .x){
	
	
	
	
	
	ans = .COM(.x, 'ConvertHangulAndHanja', ConversionsMode, FastConversion, CheckHangulEnding, EnableRecentOrdering, CustomDictionary, .dispatch = as.integer(1), .ids =221, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PasteAsNestedTable' = function(   .x){
	ans = .COM(.x, 'PasteAsNestedTable', .dispatch = as.integer(1), .ids =222, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ModifyEnclosure' = function( Style, Symbol = NA, EnclosedText = NA ,  .x){
	if( missing( Style ) ) {
	stop('You must specify a value for the argument(s)  Style ')
	}
	
	
	
	ans = .COM(.x, 'ModifyEnclosure', Style, Symbol, EnclosedText, .dispatch = as.integer(1), .ids =223, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PhoneticGuide' = function( Text, Alignment = NA, Raise = NA, FontSize = NA, FontName = NA ,  .x){
	if( missing( Text ) ) {
	stop('You must specify a value for the argument(s)  Text ')
	}
	Text = as(Text, 'character')
	if(!missing( Alignment )) Alignment = as(Alignment, 'WdMailSystem')
	if(!missing( Raise )) Raise = as(Raise, 'integer')
	if(!missing( FontSize )) FontSize = as(FontSize, 'integer')
	if(!missing( FontName )) FontName = as(FontName, 'character')
	ans = .COM(.x, 'PhoneticGuide', Text, Alignment, Raise, FontSize, FontName, .dispatch = as.integer(1), .ids =224, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertDateTime' = function( DateTimeFormat = NA, InsertAsField = NA, InsertAsFullWidth = NA, DateLanguage = NA, CalendarType = NA ,  .x){
	
	
	
	
	
	ans = .COM(.x, 'InsertDateTime', DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType, .dispatch = as.integer(1), .ids =444, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'Sort' = function( ExcludeHeader = NA, FieldNumber = NA, SortFieldType = NA, SortOrder = NA, FieldNumber2 = NA, SortFieldType2 = NA, SortOrder2 = NA, FieldNumber3 = NA, SortFieldType3 = NA, SortOrder3 = NA, SortColumn = NA, Separator = NA, CaseSensitive = NA, BidiSort = NA, IgnoreThe = NA, IgnoreKashida = NA, IgnoreDiacritics = NA, IgnoreHe = NA, LanguageID = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'Sort', ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID, .dispatch = as.integer(1), .ids =484, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'DetectLanguage' = function(   .x){
	ans = .COM(.x, 'DetectLanguage', .dispatch = as.integer(1), .ids =203, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'ConvertToTable' = function( Separator = NA, NumRows = NA, NumColumns = NA, InitialColumnWidth = NA, Format = NA, ApplyBorders = NA, ApplyShading = NA, ApplyFont = NA, ApplyColor = NA, ApplyHeadingRows = NA, ApplyLastRow = NA, ApplyFirstColumn = NA, ApplyLastColumn = NA, AutoFit = NA, AutoFitBehavior = NA, DefaultTableBehavior = NA ,  .x){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(.x, 'ConvertToTable', Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior, .dispatch = as.integer(1), .ids =498, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Table')
},
'TCSCConverter' = function( WdTCSCConverterDirection = NA, CommonTerms = NA, UseVariants = NA ,  .x){
	if(!missing( WdTCSCConverterDirection )) WdTCSCConverterDirection = as(WdTCSCConverterDirection, 'WdMailSystem')
	if(!missing( CommonTerms )) CommonTerms = as(CommonTerms, 'logical')
	if(!missing( UseVariants )) UseVariants = as(UseVariants, 'logical')
	ans = .COM(.x, 'TCSCConverter', WdTCSCConverterDirection, CommonTerms, UseVariants, .dispatch = as.integer(1), .ids =499, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PasteAndFormat' = function( Type ,  .x){
	if( missing( Type ) ) {
	stop('You must specify a value for the argument(s)  Type ')
	}
	Type = as(Type, 'WdMailSystem')
	ans = .COM(.x, 'PasteAndFormat', Type, .dispatch = as.integer(1), .ids =412, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PasteExcelTable' = function( LinkedToExcel, WordFormatting, RTF ,  .x){
	if( missing( LinkedToExcel )||missing( WordFormatting )||missing( RTF ) ) {
	stop('You must specify a value for the argument(s)  LinkedToExcel, WordFormatting, RTF ')
	}
	LinkedToExcel = as(LinkedToExcel, 'logical')
	WordFormatting = as(WordFormatting, 'logical')
	RTF = as(RTF, 'logical')
	ans = .COM(.x, 'PasteExcelTable', LinkedToExcel, WordFormatting, RTF, .dispatch = as.integer(1), .ids =413, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'PasteAppendTable' = function(   .x){
	ans = .COM(.x, 'PasteAppendTable', .dispatch = as.integer(1), .ids =414, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'GoToEditableRange' = function( EditorID = NA ,  .x){
	
	ans = .COM(.x, 'GoToEditableRange', EditorID, .dispatch = as.integer(1), .ids =415, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 as(ans, 'Range')
},
'InsertXML' = function( XML, Transform = NA ,  .x){
	if( missing( XML ) ) {
	stop('You must specify a value for the argument(s)  XML ')
	}
	XML = as(XML, 'character')
	
	ans = .COM(.x, 'InsertXML', XML, Transform, .dispatch = as.integer(1), .ids =416, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertCaption' = function( Label, Title = NA, TitleAutoText = NA, Position = NA, ExcludeLabel = NA ,  .x){
	if( missing( Label ) ) {
	stop('You must specify a value for the argument(s)  Label ')
	}
	
	
	
	
	
	ans = .COM(.x, 'InsertCaption', Label, Title, TitleAutoText, Position, ExcludeLabel, .dispatch = as.integer(1), .ids =417, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
},
'InsertCrossReference' = function( ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink = NA, IncludePosition = NA, SeparateNumbers = NA, SeparatorString = NA ,  .x){
	if( missing( ReferenceType )||missing( ReferenceKind )||missing( ReferenceItem ) ) {
	stop('You must specify a value for the argument(s)  ReferenceType, ReferenceKind, ReferenceItem ')
	}
	
	ReferenceKind = as(ReferenceKind, 'WdMailSystem')
	
	
	
	
	
	ans = .COM(.x, 'InsertCrossReference', ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString, .dispatch = as.integer(1), .ids =418, .suppliedArgs = match(names(formals()), names(sys.call())[-1]))
		 ans
} )
