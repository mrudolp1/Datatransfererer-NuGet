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
--Pier and Pad Declarations
DECLARE @PPID INT
DECLARE @PP TABLE(PPID INT)
Declare @IsCONFIG VARCHAR(50)
Declare @PPCap BIT --NEW
DECLARE @PPNeeded BIT --NEW

	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'
	Set @ModelNeeded = '[Model ID Needed]'

	--Foundation ID will need passed in. Either a number or the text NULL without quotes
	SET @FndType = '[FOUNDATION TYPE]'
	SET @PPID = '[PIER AND PAD ID]'
	Set @IsCONFIG = '[CONFIGURATION]'
	Set @FndGroupNeeded = '[Fnd GRP ID Needed]'
	Set @PPNeeded = '[PIER AND PAD ID Needed]'

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
				INSERT INTO gen.structure_model (connection_group_id,foundation_group_id,guy_config_id,lattice_structure_id,pole_structure_id,critera_id) OUTPUT Inserted.id INTO @Model SELECT connection_group_id,foundation_group_id,guy_config_id,lattice_structure_id,pole_structure_id,critera_id FROM gen.structure_model WHERE id=@ModelID
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
		IF @PPID IS NULL
			BEGIN
				-- Create new Foundation ID by adding row to foundation_details
				INSERT INTO fnd.foundation_details (foundation_type) OUTPUT Inserted.id INTO @Foundation VALUES (@FndType)
				SELECT @FndID=FndID FROM @Foundation
			END
		--END
		ELSE
			BEGIN
				--Create new Foundation ID by copying previous data and pasting new row into foundation_details
				SELECT @FndgrpID=foundation_group_id FROM gen.structure_model WHERE ID=@ModelID
				INSERT INTO fnd.foundation_details (foundation_group_id,foundation_type,guy_group_id,details_id) OUTPUT Inserted.id INTO @Foundation SELECT foundation_group_id,foundation_type,guy_group_id,details_id FROM fnd.foundation_details WHERE foundation_group_id=@FndgrpID AND foundation_type=@FndType AND details_id=@PPID
				SELECT @FndID=FndID FROM @Foundation
		END

	--Create new foundation group ID by adding row to foundation_group
	INSERT INTO fnd.foundation_group OUTPUT Inserted.ID INTO @Fndgrp DEFAULT VALUES
	SELECT @FndgrpID=FndgrpID FROM @Fndgrp
	UPDATE gen.structure_model Set foundation_group_id=@FndgrpID WHERE ID=@ModelID
	UPDATE fnd.foundation_details Set foundation_group_id=@FndgrpID WHERE ID=@FndID

	--Determine Foundation_ID
	IF @PPNeeded = 1 --TRUE  
		BEGIN
			--INSERT Details
			--Determine Foundation_ID
			--IF @PPNeeded = 1 --TRUE  
				--BEGIN
					--INSERT Details
					INSERT INTO fnd.pier_pad_details (pier_shape
						,pier_diameter
						,extension_above_grade
						,pier_rebar_size
						,pier_tie_size
						,pier_tie_quantity
						,pier_reinforcement_type
						,pier_clear_cover
						,foundation_depth
						,pad_width_1
						,pad_width_2
						,pad_thickness
						,pad_rebar_size_top_dir1
						,pad_rebar_size_bottom_dir1
						,pad_rebar_size_top_dir2
						,pad_rebar_size_bottom_dir2
						,pad_rebar_quantity_top_dir1
						,pad_rebar_quantity_bottom_dir1
						,pad_rebar_quantity_top_dir2
						,pad_rebar_quantity_bottom_dir2
						,pad_clear_cover
						,rebar_grade
						,concrete_compressive_strength
						,dry_concrete_density
						,total_soil_unit_weight
						,bearing_type
						,nominal_bearing_capacity
						,cohesion
						,friction_angle
						,spt_blow_count
						,base_friction_factor
						,neglect_depth
						,bearing_distribution_type
						,groundwater_depth
						,top_and_bottom_rebar_different
						,block_foundation
						,rectangular_foundation
						,base_plate_distance_above_foundation
						,bolt_circle_bearing_plate_width
						,pier_rebar_quantity
						,basic_soil_check
						,structural_check
						,tool_version,
						modified) OUTPUT INSERTED.ID INTO @PP VALUES ('[INSERT ALL PIER AND PAD DETAILS]')
					SELECT @PPID=PPID FROM @PP

					UPDATE fnd.foundation_details Set details_id=@PPID WHERE ID=@FndID

				END

END--Select existing foundation group or insert new