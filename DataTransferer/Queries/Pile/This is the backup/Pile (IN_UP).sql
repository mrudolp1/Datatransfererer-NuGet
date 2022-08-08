--Model Declarations
DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
DECLARE @ModelNeeded BIT --NEW
--DECLARE @STR_TYPE VARCHAR(50)
--Foundation Group Declarations (NEW)
DECLARE @Fndgrp TABLE(FndgrpID INT)
DECLARE @FndgrpID INT
--Foundation Type Declarations
DECLARE @Foundation TABLE(FndID INT)
DECLARE @FndID INT
DECLARE @FndType VARCHAR(255)
DECLARE @FndGroupNeeded BIT --NEW
--Pile Declarations
DECLARE @PID INT
DECLARE @Pile TABLE(PID INT)
Declare @IsCONFIG VARCHAR(50)
Declare @PileCap BIT --NEW
DECLARE @PileNeeded BIT --NEW

	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'
	Set @ModelNeeded = '[Model ID Needed]'

	--Foundation ID will need passed in. Either a number or the text NULL without quotes
	SET @FndType = '[FOUNDATION TYPE]'
	SET @PID = '[Pile ID]'
	Set @IsCONFIG = '[CONFIGURATION]'
	Set @FndGroupNeeded = '[Fnd GRP ID Needed]'
	Set @PileNeeded = '[Pile ID Needed]'

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


--Determine foundation_group_id (Table Impacts: gen.structure_model & fnd.foundation_group & fnd.foundation_details)
IF @FndGroupNeeded = 1 --TRUE (Reference isfndGroupNeeded)
BEGIN
	---Before creating new foundation ID, Need to select foundation detail per most recent foundation group and insert new row in foundation details
	IF @PID IS NULL
	Begin
		-- Create new Foundation ID by adding row to foundation_details
		INSERT INTO fnd.foundation_details (foundation_type) OUTPUT Inserted.id INTO @Foundation VALUES (@FndType)
		SELECT @FndID=FndID FROM @Foundation
	End
	ELSE
		BEGIN
			--Check and see if PID was accidentally copied from a different file. If doesn't match for BU, create new Foundation ID in foundation_details
			SELECT @FndgrpID=foundation_group_id FROM gen.structure_model WHERE ID=@ModelID
			IF EXISTS(SELECT * FROM fnd.foundation_details WHERE foundation_group_id=@FndgrpID AND foundation_type=@FndType AND details_id=@PID)
				BEGIN
					--Create new Foundation ID by copying previous data and pasting new row into foundation_details
					SELECT @FndgrpID=foundation_group_id FROM gen.structure_model WHERE ID=@ModelID
					INSERT INTO fnd.foundation_details (foundation_group_id,foundation_type,guy_group_id,details_id) OUTPUT Inserted.id INTO @Foundation SELECT foundation_group_id,foundation_type,guy_group_id,details_id FROM fnd.foundation_details WHERE foundation_group_id=@FndgrpID AND foundation_type=@FndType AND details_id=@PID
					SELECT @FndID=FndID FROM @Foundation
				END
			ELSE
				BEGIN
					-- Create new Foundation ID by adding row to foundation_details
					INSERT INTO fnd.foundation_details (foundation_type) OUTPUT Inserted.id INTO @Foundation VALUES (@FndType)
					SELECT @FndID=FndID FROM @Foundation
				END
		END

	--Create new foundation group ID by adding row to foundation_group
	INSERT INTO fnd.foundation_group OUTPUT Inserted.ID INTO @Fndgrp DEFAULT VALUES
	SELECT @FndgrpID=FndgrpID FROM @Fndgrp
	UPDATE gen.structure_model Set foundation_group_id=@FndgrpID WHERE ID=@ModelID
	UPDATE fnd.foundation_details Set foundation_group_id=@FndgrpID WHERE ID=@FndID

	--Determine Foundation_ID
	IF @PileNeeded = 1 --TRUE  
	BEGIN
		--INSERT Details
		INSERT INTO fnd.pile_details (load_eccentricity,bolt_circle_bearing_plate_width,pile_shape,pile_material,pile_length,pile_diameter_width,pile_pipe_thickness,pile_soil_capacity_given,steel_yield_strength,pile_type_option,rebar_quantity,pile_group_config,foundation_depth,pad_thickness,pad_width_dir1,pad_width_dir2,pad_rebar_size_bottom,pad_rebar_size_top,pad_rebar_quantity_bottom_dir1,pad_rebar_quantity_top_dir1,pad_rebar_quantity_bottom_dir2,pad_rebar_quantity_top_dir2,pier_shape,pier_diameter,extension_above_grade,pier_rebar_size,pier_rebar_quantity,pier_tie_size,rebar_grade,concrete_compressive_strength,groundwater_depth,total_soil_unit_weight,cohesion,friction_angle,neglect_depth,spt_blow_count,pile_negative_friction_force,pile_ultimate_compression,pile_ultimate_tension,top_and_bottom_rebar_different,ultimate_gross_end_bearing,skin_friction_given,pile_quantity_circular,group_diameter_circular,pile_column_quantity,pile_row_quantity,pile_columns_spacing,pile_row_spacing,group_efficiency_factor_given,group_efficiency_factor,cap_type,pile_quantity_asymmetric,pile_spacing_min_asymmetric,quantity_piles_surrounding,pile_cap_reference,tool_version,Soil_110,Structural_105) OUTPUT INSERTED.ID INTO @Pile VALUES ([INSERT ALL PILE DETAILS])
		SELECT @PID=PID FROM @Pile

		--INSERT Soil Layers 
		--IF @PileCap = 0 --FALSE
		--BEGIN
			INSERT INTO fnd.pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])
		--END

		--INSERT Pile Location Information if required (lines 98 and 99 are formatted to be easily replaced when ID already exists)
		BEGIN IF @IsCONFIG = 'Asymmetric'
		INSERT INTO fnd.pile_location VALUES ([INSERT ALL PILE LOCATIONS]) End

		UPDATE fnd.foundation_details Set details_id=@PID WHERE ID=@FndID

	END

END--Select existing foundation group or insert new


--Don't need query listed below. 
--IF @FndGroupNeeded = 0 --FALSE (think only need to report what happens if true)
--BEGIN
	--INSERT INTO @Fndgrp (FndgrpID) SELECT foundation_group_id FROM gen.structure_model WHERE ID=@ModelID
	--SELECT @FndgrpID=FndgrpID FROM @Fndgrp
	--code right above performs the same function. 
	--SELECT @FndgrpID=foundation_group_id FROM gen.structure_model WHERE ID=@ModelID
	--IF @FndgrpID Is Null --applies when BU's are entered into database for the first time. Might be able to remove if @FNDChangeFlag can detect Null values. 
	----Create new foundation group ID by adding row to foundation_group
	--BEGIN
	--	INSERT INTO fnd.foundation_group OUTPUT Inserted.ID INTO @Fndgrp DEFAULT VALUES
	--	SELECT @FndgrpID=FndgrpID FROM @Fndgrp
	--	UPDATE gen.structure_model Set foundation_group_id=@FndgrpID WHERE ID=@ModelID
	--END
	----ELSE
	----BEGIN
	--	--INSERT INTO @Fndgrp (FndgrpID) SELECT foundation_group_id FROM gen.structure_model WHERE ID=@ModelID --No longer required based on code from Ian (line 61)
	--	--SELECT @FndgrpID=FndgrpID FROM @Fndgrp
	----END
--END