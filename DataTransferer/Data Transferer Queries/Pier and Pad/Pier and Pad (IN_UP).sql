--Model Declarations
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
--Pier & Pad Declarations
DECLARE @PnPID INT
DECLARE @PierPad TABLE(PnPID INT)

	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'

	--Foundation ID will need passed in. Either a number or the text NULL without quotes
	SET @FndType = '[FOUNDATION TYPE]'
	SET @PnPID = '[PIER AND PAD ID]'

	BEGIN
		IF EXISTS(SELECT * FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True') 
			INSERT INTO @Model (ModelID) SELECT ID FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True'
		ELSE
			INSERT INTO structure_model (bus_unit,structure_id,existing_geometry) OUTPUT INSERTED.ID INTO @Model VALUES (@BU,@STR_ID,'True')
	END --Select existing model ID or insert new

	SELECT @ModelID=ModelID FROM @Model
		
	BEGIN
		IF @PnPID IS NULL 
			BEGIN
				INSERT INTO foundation_details (model_id,foundation_type) OUTPUT INSERTED.ID INTO @Foundation VALUES(@ModelID,@FndType)
				SELECT @FndID=FndID FROM @Foundation
				INSERT INTO pier_pad_details VALUES ([INSERT ALL PIER AND PAD DETAILS])
			END
		ELSE
			BEGIN
				(SELECT * FROM TEMPORARY)
			END
	END --If foundation ID is NULL, insert a foundation based on the type provided and output the new foundation ID