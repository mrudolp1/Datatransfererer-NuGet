DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
DECLARE @STR_TYPE VARCHAR(50)
DECLARE @Foundation TABLE(FndID INT)
DECLARE @FndID INT
DECLARE @FndType VARCHAR(255)
DECLARE @DpID INT
DECLARE @DrilledPier TABLE(DpID INT, IsEmbed BIT, IsBelled BIT)
DECLARE @IsEmbed BIT
DECLARE @IsBelled BIT
DECLARE @EmbeddedPole TABLE(EmbedID INT)
DECLARE @EmbedID INT
DECLARE @DrilledPierSection Table(SecID INT)
DEClARE @SecID INT	

	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'
	--Foundation ID will need passed in. Either a number or the text NULL without quotes
	SET @FndType = '[FOUNDATION TYPE]'
	SET @DpID = '[DRILLED PIER ID]'
	--If Drilled Pier ID is NULL, insert a drilled pier based on the information provided and output the new drilled pier ID for Sections, Rebar, & Soil Layers	
	SET @IsEmbed = '[EMBED BOOLEAN]'
	SET @IsBelled = '[BELL BOOLEAN]'

	BEGIN
		IF EXISTS(SELECT * FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True') 
			INSERT INTO @Model (ModelID) SELECT ID FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True'
		ELSE
			INSERT INTO structure_model (bus_unit,structure_id,existing_geometry) OUTPUT INSERTED.ID INTO @Model VALUES (@BU,@STR_ID,'True')
	END --Select existing model ID or insert new

	SELECT @ModelID=ModelID FROM @Model
		
	BEGIN
		IF @DpID IS NULL 
			BEGIN
				INSERT INTO foundation_details (model_id,foundation_type) OUTPUT INSERTED.ID INTO @Foundation VALUES(@ModelID,@FndType)
				SELECT @FndID=FndID FROM @Foundation
			END
		ELSE
			BEGIN
				SELECT @FndID=foundation_id FROM drilled_pier_details WHERE ID=@DpID
			END
	END --If foundation ID is NULL, insert a foundation based on the type provided and output the new foundation ID

	BEGIN
		IF @DpID IS NULL
			BEGIN
				INSERT INTO drilled_pier_details OUTPUT INSERTED.ID,INSERTED.embedded_pole,INSERTED.belled_pier INTO @DrilledPier VALUES ([INSERT ALL PIER DETAILS])
				SELECT @DpID=DpID,@IsEmbed=IsEmbed,@IsBelled=IsBelled FROM @DrilledPier	
				
					
				BEGIN
					IF @IsBelled = 'True'
						INSERT INTO belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])
				END --INSERT Belled Pier information if required

					
				BEGIN
					IF @IsEmbed = 'True'
						INSERT INTO embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])								
						SELECT @EmbedID=EmbedID FROM @EmbeddedPole

						--INSERT Embedded Pole Sections
						INSERT INTO embedded_pole_section VALUES ([INSERT ALL EMBEDDED SECTIONS])
				END --INSERT Embedded Pole information if required

				--INSERT Soil Layers 
				INSERT INTO drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])
					
				--INSERT Drilled Pier Sections & Rebar
				--*[DRILLED PIER SECTIONS]*--
			END
		ELSE
			(SELECT * FROM TEMPORARY)
	END