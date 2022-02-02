﻿INSERT INTO tnx.base_structure VALUES (@TowerRec,
										@TowerDatabase,
										@TowerName,
										@TowerHeight,
										@TowerFaceWidth,
										@TowerNumSections,
										@TowerSectionLength,
										@TowerDiagonalSpacing,
										@TowerDiagonalSpacingEx,
										@TowerBraceType,
										@TowerFaceBevel,
										@TowerTopGirtOffset,
										@TowerBotGirtOffset,
										@TowerHasKBraceEndPanels,
										@TowerHasHorizontals,
										@TowerLegType,
										@TowerLegSize,
										@TowerLegGrade,
										@TowerLegMatlGrade,
										@TowerDiagonalGrade,
										@TowerDiagonalMatlGrade,
										@TowerInnerBracingGrade,
										@TowerInnerBracingMatlGrade,
										@TowerTopGirtGrade,
										@TowerTopGirtMatlGrade,
										@TowerBotGirtGrade,
										@TowerBotGirtMatlGrade,
										@TowerInnerGirtGrade,
										@TowerInnerGirtMatlGrade,
										@TowerLongHorizontalGrade,
										@TowerLongHorizontalMatlGrade,
										@TowerShortHorizontalGrade,
										@TowerShortHorizontalMatlGrade,
										@TowerDiagonalType,
										@TowerDiagonalSize,
										@TowerInnerBracingType,
										@TowerInnerBracingSize,
										@TowerTopGirtType,
										@TowerTopGirtSize,
										@TowerBotGirtType,
										@TowerBotGirtSize,
										@TowerNumInnerGirts,
										@TowerInnerGirtType,
										@TowerInnerGirtSize,
										@TowerLongHorizontalType,
										@TowerLongHorizontalSize,
										@TowerShortHorizontalType,
										@TowerShortHorizontalSize,
										@TowerRedundantGrade,
										@TowerRedundantMatlGrade,
										@TowerRedundantType,
										@TowerRedundantDiagType,
										@TowerRedundantSubDiagonalType,
										@TowerRedundantSubHorizontalType,
										@TowerRedundantVerticalType,
										@TowerRedundantHipType,
										@TowerRedundantHipDiagonalType,
										@TowerRedundantHorizontalSize,
										@TowerRedundantHorizontalSize2,
										@TowerRedundantHorizontalSize3,
										@TowerRedundantHorizontalSize4,
										@TowerRedundantDiagonalSize,
										@TowerRedundantDiagonalSize2,
										@TowerRedundantDiagonalSize3,
										@TowerRedundantDiagonalSize4,
										@TowerRedundantSubHorizontalSize,
										@TowerRedundantSubDiagonalSize,
										@TowerSubDiagLocation,
										@TowerRedundantVerticalSize,
										@TowerRedundantHipSize,
										@TowerRedundantHipSize2,
										@TowerRedundantHipSize3,
										@TowerRedundantHipSize4,
										@TowerRedundantHipDiagonalSize,
										@TowerRedundantHipDiagonalSize2,
										@TowerRedundantHipDiagonalSize3,
										@TowerRedundantHipDiagonalSize4,
										@TowerSWMult,
										@TowerWPMult,
										@TowerAutoCalcKSingleAngle,
										@TowerAutoCalcKSolidRound,
										@TowerAfGusset,
										@TowerTfGusset,
										@TowerGussetBoltEdgeDistance,
										@TowerGussetGrade,
										@TowerGussetMatlGrade,
										@TowerAfMult,
										@TowerArMult,
										@TowerFlatIPAPole,
										@TowerRoundIPAPole,
										@TowerFlatIPALeg,
										@TowerRoundIPALeg,
										@TowerFlatIPAHorizontal,
										@TowerRoundIPAHorizontal,
										@TowerFlatIPADiagonal,
										@TowerRoundIPADiagonal,
										@TowerCSA_S37_SpeedUpFactor,
										@TowerKLegs,
										@TowerKXBracedDiags,
										@TowerKKBracedDiags,
										@TowerKZBracedDiags,
										@TowerKHorzs,
										@TowerKSecHorzs,
										@TowerKGirts,
										@TowerKInners,
										@TowerKXBracedDiagsY,
										@TowerKKBracedDiagsY,
										@TowerKZBracedDiagsY,
										@TowerKHorzsY,
										@TowerKSecHorzsY,
										@TowerKGirtsY,
										@TowerKInnersY,
										@TowerKRedHorz,
										@TowerKRedDiag,
										@TowerKRedSubDiag,
										@TowerKRedSubHorz,
										@TowerKRedVert,
										@TowerKRedHip,
										@TowerKRedHipDiag,
										@TowerKTLX,
										@TowerKTLZ,
										@TowerKTLLeg,
										@TowerInnerKTLX,
										@TowerInnerKTLZ,
										@TowerInnerKTLLeg,
										@TowerStitchBoltLocationHoriz,
										@TowerStitchBoltLocationDiag,
										@TowerStitchBoltLocationRed,
										@TowerStitchSpacing,
										@TowerStitchSpacingDiag,
										@TowerStitchSpacingHorz,
										@TowerStitchSpacingRed,
										@TowerLegNetWidthDeduct,
										@TowerLegUFactor,
										@TowerDiagonalNetWidthDeduct,
										@TowerTopGirtNetWidthDeduct,
										@TowerBotGirtNetWidthDeduct,
										@TowerInnerGirtNetWidthDeduct,
										@TowerHorizontalNetWidthDeduct,
										@TowerShortHorizontalNetWidthDeduct,
										@TowerDiagonalUFactor,
										@TowerTopGirtUFactor,
										@TowerBotGirtUFactor,
										@TowerInnerGirtUFactor,
										@TowerHorizontalUFactor,
										@TowerShortHorizontalUFactor,
										@TowerLegConnType,
										@TowerLegNumBolts,
										@TowerDiagonalNumBolts,
										@TowerTopGirtNumBolts,
										@TowerBotGirtNumBolts,
										@TowerInnerGirtNumBolts,
										@TowerHorizontalNumBolts,
										@TowerShortHorizontalNumBolts,
										@TowerLegBoltGrade,
										@TowerLegBoltSize,
										@TowerDiagonalBoltGrade,
										@TowerDiagonalBoltSize,
										@TowerTopGirtBoltGrade,
										@TowerTopGirtBoltSize,
										@TowerBotGirtBoltGrade,
										@TowerBotGirtBoltSize,
										@TowerInnerGirtBoltGrade,
										@TowerInnerGirtBoltSize,
										@TowerHorizontalBoltGrade,
										@TowerHorizontalBoltSize,
										@TowerShortHorizontalBoltGrade,
										@TowerShortHorizontalBoltSize,
										@TowerLegBoltEdgeDistance,
										@TowerDiagonalBoltEdgeDistance,
										@TowerTopGirtBoltEdgeDistance,
										@TowerBotGirtBoltEdgeDistance,
										@TowerInnerGirtBoltEdgeDistance,
										@TowerHorizontalBoltEdgeDistance,
										@TowerShortHorizontalBoltEdgeDistance,
										@TowerDiagonalGageG1Distance,
										@TowerTopGirtGageG1Distance,
										@TowerBotGirtGageG1Distance,
										@TowerInnerGirtGageG1Distance,
										@TowerHorizontalGageG1Distance,
										@TowerShortHorizontalGageG1Distance,
										@TowerRedundantHorizontalBoltGrade,
										@TowerRedundantHorizontalBoltSize,
										@TowerRedundantHorizontalNumBolts,
										@TowerRedundantHorizontalBoltEdgeDistance,
										@TowerRedundantHorizontalGageG1Distance,
										@TowerRedundantHorizontalNetWidthDeduct,
										@TowerRedundantHorizontalUFactor,
										@TowerRedundantDiagonalBoltGrade,
										@TowerRedundantDiagonalBoltSize,
										@TowerRedundantDiagonalNumBolts,
										@TowerRedundantDiagonalBoltEdgeDistance,
										@TowerRedundantDiagonalGageG1Distance,
										@TowerRedundantDiagonalNetWidthDeduct,
										@TowerRedundantDiagonalUFactor,
										@TowerRedundantSubDiagonalBoltGrade,
										@TowerRedundantSubDiagonalBoltSize,
										@TowerRedundantSubDiagonalNumBolts,
										@TowerRedundantSubDiagonalBoltEdgeDistance,
										@TowerRedundantSubDiagonalGageG1Distance,
										@TowerRedundantSubDiagonalNetWidthDeduct,
										@TowerRedundantSubDiagonalUFactor,
										@TowerRedundantSubHorizontalBoltGrade,
										@TowerRedundantSubHorizontalBoltSize,
										@TowerRedundantSubHorizontalNumBolts,
										@TowerRedundantSubHorizontalBoltEdgeDistance,
										@TowerRedundantSubHorizontalGageG1Distance,
										@TowerRedundantSubHorizontalNetWidthDeduct,
										@TowerRedundantSubHorizontalUFactor,
										@TowerRedundantVerticalBoltGrade,
										@TowerRedundantVerticalBoltSize,
										@TowerRedundantVerticalNumBolts,
										@TowerRedundantVerticalBoltEdgeDistance,
										@TowerRedundantVerticalGageG1Distance,
										@TowerRedundantVerticalNetWidthDeduct,
										@TowerRedundantVerticalUFactor,
										@TowerRedundantHipBoltGrade,
										@TowerRedundantHipBoltSize,
										@TowerRedundantHipNumBolts,
										@TowerRedundantHipBoltEdgeDistance,
										@TowerRedundantHipGageG1Distance,
										@TowerRedundantHipNetWidthDeduct,
										@TowerRedundantHipUFactor,
										@TowerRedundantHipDiagonalBoltGrade,
										@TowerRedundantHipDiagonalBoltSize,
										@TowerRedundantHipDiagonalNumBolts,
										@TowerRedundantHipDiagonalBoltEdgeDistance,
										@TowerRedundantHipDiagonalGageG1Distance,
										@TowerRedundantHipDiagonalNetWidthDeduct,
										@TowerRedundantHipDiagonalUFactor,
										@TowerDiagonalOutOfPlaneRestraint,
										@TowerTopGirtOutOfPlaneRestraint,
										@TowerBottomGirtOutOfPlaneRestraint,
										@TowerMidGirtOutOfPlaneRestraint,
										@TowerHorizontalOutOfPlaneRestraint,
										@TowerSecondaryHorizontalOutOfPlaneRestraint,
										@TowerUniqueFlag,
										@TowerDiagOffsetNEY,
										@TowerDiagOffsetNEX,
										@TowerDiagOffsetPEY,
										@TowerDiagOffsetPEX,
										@TowerKbraceOffsetNEY,
										@TowerKbraceOffsetNEX,
										@TowerKbraceOffsetPEY,
										@TowerKbraceOffsetPEX)
SELECT SCOPE_IDENTITY()
		--SELECT @baseSectionID = baseSectionID FROM @baseSection

--INSERT INTO tnx.base_structure_xref VALUES(@baseSectionID, @tnxID)s