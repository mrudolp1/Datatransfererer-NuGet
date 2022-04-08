--Model Declarations
DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
DECLARE @ModelNeeded BIT --NEW
--DECLARE @STR_TYPE VARCHAR(50)
--Foundation Group Declarations
DECLARE @Fndgrp TABLE(FndgrpID INT)
DECLARE @FndgrpID INT
--Foundation Type Declarations
DECLARE @Foundation TABLE(FndID INT)
DECLARE @FndID INT
DECLARE @FndType VARCHAR(255)
DECLARE @FndGroupNeeded BIT 
--Drilled Pier Declarations
DECLARE @DpID INT
--DECLARE @DrilledPier TABLE(DpID INT, IsEmbed BIT, IsBelled BIT)
DECLARE @DrilledPier TABLE(DpID INT)
DECLARE @IsEmbed BIT
DECLARE @IsBelled BIT
DECLARE @EmbeddedPole TABLE(EmbedID INT)
DECLARE @EmbedID INT
DECLARE @DrilledPierSection Table(SecID INT)
DEClARE @SecID INT
DECLARE @DPNeeded BIT

	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'
	SET @ModelNeeded = '[Model ID Needed]'

	--Foundation ID will need passed in. Either a number or the text NULL without quotes
	SET @FndType = '[FOUNDATION TYPE]'
	SET @DpID = '[DRILLED PIER ID]'
	--If Drilled Pier ID is NULL, insert a drilled pier based on the information provided and output the new drilled pier ID for Sections, Rebar, & Soil Layers	
	SET @IsEmbed = '[EMBED BOOLEAN]'
	SET @IsBelled = '[BELL BOOLEAN]'

	Set @FndGroupNeeded = '[Fnd GRP ID Needed]'
	Set @DPNeeded = '[DRILLED PIER ID Needed]'

	--Determine model_id (Table Impacts: gen.structure_model_xref & gen.structure_model)
IF EXISTS(SELECT * FROM gen.structure_model_xref WHERE bus_unit=@BU AND structure_id=@STR_ID) 
	BEGIN
		--If exists, select model_id from structure_model_xref
		INSERT INTO @Model (ModelID) SELECT model_id FROM gen.structure_model_xref WHERE bus_unit=@BU AND structure_id=@STR_ID AND isActive='True' --ORDER BY model_id
		SELECT @ModelID=ModelID FROM @Model
		--If changes occurred with any field in structure_model table, create new model ID for reference
		IF @ModelNeeded = 1 --TRUE (Reference ismodelneeded)
		BEGIN
			--Update status to FALSE for existing model_id
			UPDATE gen.structure_model_xref Set isActive='False' WHERE model_id=@ModelID
			
			--Create new Model ID by copying previous data and pasting new row into Structure_model
			--INSERT INTO gen.structure_model (connection_group_id,foundation_group_id,guy_config_id,lattice_structure_id,pole_structure_id,critera_id) OUTPUT Inserted.id INTO @Model SELECT connection_group_id,foundation_group_id,guy_config_id,lattice_structure_id,pole_structure_id,critera_id FROM gen.structure_model WHERE id=@ModelID
			
				--Delete temp table if already exists
				IF OBJECT_ID(N'tempdb..#TempTable') IS NOT NULL
				Begin
					DROP TABLE #TempTable
				End

				SELECT * INTO #TempTable FROM gen.structure_model WHERE ID=@ModelID
				ALTER TABLE #TempTable DROP COLUMN ID
				INSERT INTO gen.structure_model OUTPUT INSERTED.ID INTO @Model SELECT * FROM #TempTable
				DROP TABLE #TempTable

			SELECT @ModelID=ModelID FROM @Model
			--Create new row in structure_model_xref, associating BU to newly created Model ID
			INSERT INTO gen.structure_model_xref (model_id,bus_unit,structure_id,isActive) VALUES (@ModelID,@BU,@STR_ID,'True')
		END
	END
ELSE
	BEGIN
		-- Create new Model ID by adding row to Structure_model 
		INSERT INTO gen.structure_model OUTPUT Inserted.ID INTO @Model DEFAULT VALUES
		SELECT @ModelID=ModelID FROM @Model
		--Create new row in structure_model_xref, associating BU to newly created Model ID
		INSERT INTO gen.structure_model_xref (model_id,bus_unit,structure_id,isActive) VALUES (@ModelID,@BU,@STR_ID,'True')
	END--Select existing model ID or insert new


--MRP - 10-11-21 EDITED this to separate foundation group from foundation ID. There should only be one group added per drilled pier tool
---------------------------------------------------------------------------------------------------------------------------------------------------
--Determine foundation_group_id (Table Impacts: gen.structure_model & fnd.foundation_group & fnd.foundation_details)
IF @FndGroupNeeded = 1 --TRUE (Reference isfndGroupNeeded)
	BEGIN
		---Before creating new foundation ID, Need to select foundation detail per most recent foundation group and insert new row in foundation details
		IF @DpID IS NULL
			BEGIN
				-- Create new Foundation ID by adding row to foundation_details
				INSERT INTO fnd.foundation_details (foundation_type) OUTPUT Inserted.id INTO @Foundation VALUES (@FndType)
				SELECT @FndID=FndID FROM @Foundation
			END
		ELSE
			BEGIN
				--Create new Foundation ID by copying previous data and pasting new row into foundation_details
				SELECT @FndgrpID=foundation_group_id FROM gen.structure_model WHERE ID=@ModelID
				INSERT INTO fnd.foundation_details (foundation_group_id,foundation_type,guy_group_id,details_id) OUTPUT Inserted.id INTO @Foundation SELECT foundation_group_id,foundation_type,guy_group_id,details_id FROM fnd.foundation_details WHERE foundation_group_id=@FndgrpID AND foundation_type=@FndType AND details_id=@DpID
				SELECT @FndID=FndID FROM @Foundation
			END

		--Create new foundation group ID by adding row to foundation_group
		INSERT INTO fnd.foundation_group OUTPUT Inserted.ID INTO @Fndgrp DEFAULT VALUES
		SELECT @FndgrpID=FndgrpID FROM @Fndgrp
		UPDATE gen.structure_model Set foundation_group_id=@FndgrpID WHERE ID=@ModelID
		UPDATE fnd.foundation_details Set foundation_group_id=@FndgrpID WHERE ID=@FndID

		--Determine Foundation_ID
		IF @DPNeeded = 1 --TRUE  
			BEGIN
				--INSERT Details
				INSERT INTO fnd.drilled_pier_details (local_drilled_pier_id,local_drilled_pier_profile,foundation_depth,extension_above_grade,groundwater_depth,assume_min_steel,check_shear_along_depth,utilize_shear_friction_methodology,embedded_pole,belled_pier,soil_layer_quantity,concrete_compressive_strength,tie_yield_strength,longitudinal_rebar_yield_strength,rebar_effective_depths,rebar_cage_2_fy_override,rebar_cage_3_fy_override,shear_override_crit_depth,shear_crit_depth_override_comp,shear_crit_depth_override_uplift,bearing_toggle_type,tool_version,modified) OUTPUT INSERTED.ID INTO @DrilledPier VALUES ([INSERT ALL DRILLED PIER DETAILS])
				SELECT @DpID=DpID FROM @DrilledPier

				--INSERT Soil Layers 
				--INSERT INTO fnd.drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])

				--INSERT Profile
				INSERT INTO fnd.drilled_pier_profile VALUES ([INSERT ALL DRILLED PIER PROFILES])

				BEGIN --Belled Pier
					IF @IsBelled = 'True'
						INSERT INTO fnd.belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])
				END --INSERT Belled Pier information if required

				BEGIN --Embedded Pole
					IF @IsEmbed = 'True'
						INSERT INTO fnd.embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])								
						SELECT @EmbedID=EmbedID FROM @EmbeddedPole
				END --INSERT Embedded Pole information if required
					
				--INSERT Drilled Pier Sections & Rebar
				--*[DRILLED PIER SECTIONS]*--

				UPDATE fnd.foundation_details Set details_id=@DpID WHERE ID=@FndID

			END

	END--Select existing foundation group or insert new

IF @FndGroupNeeded = 0 --FALSE --if Foundation Group is not Needed
	BEGIN
		--Determine Foundation_ID
		IF @DPNeeded = 1 --TRUE  
			BEGIN
				--INSERT Details
				INSERT INTO fnd.drilled_pier_details (local_drilled_pier_id,local_drilled_pier_profile,foundation_depth,extension_above_grade,groundwater_depth,assume_min_steel,check_shear_along_depth,utilize_shear_friction_methodology,embedded_pole,belled_pier,soil_layer_quantity,concrete_compressive_strength,tie_yield_strength,longitudinal_rebar_yield_strength,rebar_effective_depths,rebar_cage_2_fy_override,rebar_cage_3_fy_override,shear_override_crit_depth,shear_crit_depth_override_comp,shear_crit_depth_override_uplift,bearing_toggle_type,tool_version,modified) OUTPUT INSERTED.ID INTO @DrilledPier VALUES ([INSERT ALL DRILLED PIER DETAILS])
				SELECT @DpID=DpID FROM @DrilledPier

				--INSERT Profile
				INSERT INTO fnd.drilled_pier_profile VALUES ([INSERT ALL DRILLED PIER PROFILES])
			
				--INSERT Soil Layers 
				INSERT INTO fnd.drilled_pier_soil_layer VALUES ([INSERT ALL SOIL LAYERS])

				BEGIN --Belled Pier
					IF @IsBelled = 'True'
						INSERT INTO fnd.belled_pier_details VALUES ([INSERT ALL BELLED PIER DETAILS])
				END --INSERT Belled Pier information if required

				BEGIN --Embedded Pole
					IF @IsEmbed = 'True'
						INSERT INTO fnd.embedded_pole_details OUTPUT INSERTED.ID INTO @EmbeddedPole VALUES ([INSERT ALL EMBEDDED POLE DETAILS])								
						SELECT @EmbedID=EmbedID FROM @EmbeddedPole
				END --INSERT Embedded Pole information if required
					
				--INSERT Drilled Pier Sections & Rebar -> SubQuery [Drilled Piers Sections (IN_UP).sql]
				--*[DRILLED PIER SECTIONS]*--

				
				--SELECT EXISTING FOUNDATION GROUP ID
				SELECT @FndgrpID=foundation_group_id FROM gen.structure_model WHERE ID=@ModelID
			
				--INSERT Foundation ID into Foundation Details
				INSERT INTO fnd.foundation_details (foundation_group_id,foundation_type,details_id) OUTPUT INSERTED.ID INTO @Foundation VALUES (@FndgrpID,@FndType,@DpID)
				SELECT @FndID=FndID FROM @Foundation

			END
	END