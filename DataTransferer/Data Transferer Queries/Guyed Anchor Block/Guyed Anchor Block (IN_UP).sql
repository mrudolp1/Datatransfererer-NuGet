DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
DECLARE @STR_TYPE VARCHAR(50)
DECLARE @Foundation TABLE(FndID INT)
DECLARE @FndID INT
DECLARE @FndType VARCHAR(255)
DECLARE @GabID INT
DECLARE @GuyedAnchorBlock TABLE(GabID INT) 

--WORK IN PROGRESS

	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'
	--Foundation ID will need passed in. Either a number or the text NULL without quotes
	SET @FndType = '[FOUNDATION TYPE]'
	SET @GabID = '[GUYED ANCHOR BLOCK ID]'
	--If ID is NULL, insert an object based on the information provided and output the new object ID	

	BEGIN
		IF EXISTS(SELECT * FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True') 
			INSERT INTO @Model (ModelID) SELECT ID FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True'
		ELSE
			INSERT INTO structure_model (bus_unit,structure_id,existing_geometry) OUTPUT INSERTED.ID INTO @Model VALUES (@BU,@STR_ID,'True')
	END --Select existing model ID or insert new

	SELECT @ModelID=ModelID FROM @Model
	   		
	BEGIN
		IF @GabID IS NULL 
			BEGIN
				INSERT INTO foundation_details (model_id,foundation_type) OUTPUT INSERTED.ID INTO @Foundation VALUES(@ModelID,@FndType)
				SELECT @FndID=FndID FROM @Foundation
			END
		ELSE
			BEGIN
				SELECT @FndID=foundation_id FROM anchor_block_details WHERE ID=@GabID
			END
	END --If foundation ID is NULL, insert a foundation based on the type provided and output the new foundation ID

	BEGIN
		IF @GabID IS NULL
			BEGIN
				INSERT INTO anchor_block_details OUTPUT INSERTED.ID INTO @GuyedAnchorBlock VALUES ([INSERT ALL GUYED ANCHOR BLOCK DETAILS])
				SELECT @GabID=GabID FROM @GuyedAnchorBlock	

				--INSERT Soil Layers 
				INSERT INTO anchor_block_soil_layer VALUES ([INSERT ALL SOIL LAYERS])

				--INSERT Profiles 
				INSERT INTO anchor_block_profile VALUES ([INSERT ALL GUYED ANCHOR BLOCK PROFILES])

			END
		ELSE
			(SELECT * FROM TEMPORARY)
	END
