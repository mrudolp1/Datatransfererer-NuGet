﻿--Model Declarations
DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
DECLARE @STR_TYPE VARCHAR(50)
--Foundation Type Declarations
DECLARE @Foundation TABLE(FndID INT)
DECLARE @FndID INT
DECLARE @FndType VARCHAR(255)
--Pile Declarations
DECLARE @PID INT
DECLARE @Pile TABLE(PID INT)
Declare @IsCONFIG VARCHAR(50)

	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'

	--Foundation ID will need passed in. Either a number or the text NULL without quotes
	SET @FndType = '[FOUNDATION TYPE]'
	SET @PID = '[Pile ID]'
	Set @IsCONFIG = '[CONFIGURATION]'

	BEGIN
		IF EXISTS(SELECT * FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True') 
			INSERT INTO @Model (ModelID) SELECT ID FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True'
		ELSE
			INSERT INTO structure_model (bus_unit,structure_id,existing_geometry) OUTPUT INSERTED.ID INTO @Model VALUES (@BU,@STR_ID,'True')
	END --Select existing model ID or insert new

	SELECT @ModelID=ModelID FROM @Model
		
	BEGIN
		IF @PID IS NULL 
			BEGIN
				INSERT INTO foundation_details (model_id,foundation_type) OUTPUT INSERTED.ID INTO @Foundation VALUES(@ModelID,@FndType)
				SELECT @FndID=FndID FROM @Foundation
				--INSERT INTO pile_details VALUES ([INSERT ALL PILE DETAILS] DNU3)
			END
		ELSE
			BEGIN
				SELECT @FndID=foundation_id FROM pile_details WHERE ID=@PID
				--(SELECT * FROM TEMPORARY DNU3)
			END
	END --If foundation ID is NULL, insert a foundation based on the type provided and output the new foundation ID

	BEGIN
		IF @PID IS NULL
			BEGIN
				INSERT INTO pile_details OUTPUT INSERTED.ID INTO @Pile VALUES ([INSERT ALL PILE DETAILS])
				SELECT @PID=PID FROM @Pile
				--SELECT @PID=PID,@IsCONFIG=IsCONFIG FROM @Pile

				--INSERT Soil Layers 
				INSERT INTO pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])

				--INSERT Pile Location Information if required (lines 60 and 61 are formatted to be easily replaced when ID already exists)
				BEGIN IF @IsCONFIG = 'Asymmetric'
						INSERT INTO pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End

			End
		Else
			(SELECT * FROM TEMPORARY)
	End