--Model Declarations
DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
DECLARE @ModelNeeded BIT
--
DECLARE @STR_TYPE VARCHAR(50)
--CCIpole Declarations
DECLARE @PoleStructure TABLE(PoleID INT, CriteriaID INT)
DECLARE @PoleID INT
DECLARE @PoleNeeded BIT
--Other CCIpole Declarations
DECLARE @Criteria TABLE (CriteriaID INT)
DECLARE @CriteriaID INT
DECLARE @PoleSection TABLE (PoleSectionID INT)
DECLARE @PoleSectionID INT
DECLARE @PoleReinfSection TABLE (PoleReinfSectionID INT)
DECLARE @PoleReinfSectionID INT
DECLARE @ReinfGroups TABLE (ReinfGroupID INT)
DECLARE @ReinfGroupID INT
DECLARE @ReinfDetails TABLE (ReinfDetailID INT)
DECLARE @ReinfDetailID INT
DECLARE @IntGroups TABLE (IntGroupID INT)
DECLARE @IntGroupID INT
DECLARE @IntDetails TABLE (IntDetailID INT)
DECLARE @IntDetailID INT
DECLARE @PoleReinfResults TABLE (PoleReinfResultID INT)
DECLARE @PoleReinfResultID INT
DECLARE @PropReinf TABLE (ReinfID INT)
DECLARE @ReinfID INT
DECLARE @PropBolt TABLE (BoltID INT)
DECLARE @BoltID INT
DECLARE @PropMatl TABLE (MatlID INT)
DECLARE @MatlID INT


	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'
	Set @ModelNeeded = '[Model ID Needed]'

	--Foundation ID will need passed in. Either a number or the text NULL without quotes
	SET @STR_TYPE = '[STRUCTURE TYPE]'
	SET @PoleID = '[CCIPOLE ID]'
	Set @PoleNeeded = '[Pole ID Needed]'


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



--Determine pole_structure_id (Table Impacts: gen.structure_model & pole.pole_section & pole.pole_reinf_section & pole.pole_reinf_group & pole.pole_interference_group & pole.pole_reinf_results & pole.memb_prop_flat_plate & pole.bolt_prop_flat_plate & pole.matl_prop_flat_plate)
IF @PoleNeeded = 1 --TRUE (Reference isPoleNeeded)
BEGIN
	---Before creating new foundation ID, Need to select foundation detail per most recent foundation group and insert new row in foundation details
	IF @PoleID IS NULL
		--BEGIN
		--	-- Create new CCIpole ID by adding row to foundation_details
		--	INSERT INTO pole.pole_structure (foundation_type) OUTPUT Inserted.id INTO @Foundation VALUES (@FndType) --Dont know this is necessary for Pole DB
		--	SELECT @FndID=FndID FROM @Foundation
		--END
	ELSE
		BEGIN
			--Create new CCIpole ID by copying previous data and pasting new row into pole_structure
			SELECT @PoleID=pole_structure_id FROM gen.structure_model WHERE ID=@ModelID
			INSERT INTO pole.pole_structure (pole_structure_id,criteria_id) OUTPUT Inserted.id INTO @PoleStructure SELECT foundation_group_id,foundation_type,guy_group_id,details_id FROM fnd.foundation_details WHERE foundation_group_id=@FndgrpID AND foundation_type=@FndType AND details_id=@DpID
			SELECT @FndID=FndID FROM @Foundation
		END

	--Create new foundation group ID by adding row to foundation_group
	INSERT INTO fnd.foundation_group OUTPUT Inserted.ID INTO @Fndgrp DEFAULT VALUES
	SELECT @FndgrpID=FndgrpID FROM @Fndgrp
	UPDATE gen.structure_model Set foundation_group_id=@FndgrpID WHERE ID=@ModelID
	UPDATE fnd.foundation_details Set foundation_group_id=@FndgrpID WHERE ID=@FndID

	INSERT INTO fnd.drilled_pier_details (local_drilled_pier_id,local_drilled_pier_profile,foundation_depth,extension_above_grade,groundwater_depth,assume_min_steel,check_shear_along_depth,utilize_shear_friction_methodology,embedded_pole,belled_pier,soil_layer_quantity,concrete_compressive_strength,tie_yield_strength,longitudinal_rebar_yield_strength,rebar_effective_depths,rebar_cage_2_fy_override,rebar_cage_3_fy_override,shear_override_crit_depth,shear_crit_depth_override_comp,shear_crit_depth_override_uplift,bearing_toggle_type,tool_version,modified) OUTPUT INSERTED.ID INTO @DrilledPier VALUES ([INSERT ALL DRILLED PIER DETAILS])
			
	SELECT @PoleID=PoleID FROM @PoleStructure

	--Criteria 
	INSERT INTO pole.pole_analysis_criteria VALUES ([INSERT CRITERIA])

	--Unreinf Geometry
	INSERT INTO pole.pole_section VALUES ([INSERT ALL POLE SECTIONS])

	--Reinf Geometry
	INSERT INTO pole.pole_reinf_section VALUES ([INSERT ALL REINF POLE SECTIONS])

	--Reinf Groups
	INSERT INTO pole.pole_reinf_group VALUES ([INSERT ALL REINF GROUPS])

	--Reinf Details
	INSERT INTO pole.pole_reinf_details VALUES ([INSERT ALL REINF DETAILS])

	--Interference Groups
	INSERT INTO pole.pole_interference_group VALUES ([INSERT ALL INTERFERENCE GROUPS])

	--Interference Details
	INSERT INTO pole.pole_interference_details VALUES ([INSERT ALL INTERFERENCE DETAILS])

	--Reinf Results
	INSERT INTO pole.pole_reinf_results VALUES ([INSERT ALL REINF RESULTS])
	
	--Custom Reinf Properties
	INSERT INTO pole.memb_prop_flat_plate VALUES ([INSERT ALL REINF PROPERTIES])
	
	--Custom Bolt Properties
	INSERT INTO pole.bolt_prop_flat_plate VALUES ([INSERT ALL BOLT PROPERTIES])

	--Custom Matl Properties
	INSERT INTO pole.matl_prop_flat_plate VALUES ([INSERT ALL MATL PROPERTIES])
					

	UPDATE pole.pole_structure Set pole_structure_id=@PoleID WHERE ID=@FndID


END--Select existing CCIpole or insert new

IF @PoleNeeded = 0 --FALSE

	--Determine Foundation_ID
	IF @DPNeeded = 1 --TRUE  
		BEGIN
			--INSERT Details
			INSERT INTO fnd.drilled_pier_details (local_drilled_pier_id,local_drilled_pier_profile,foundation_depth,extension_above_grade,groundwater_depth,assume_min_steel,check_shear_along_depth,utilize_shear_friction_methodology,embedded_pole,belled_pier,soil_layer_quantity,concrete_compressive_strength,tie_yield_strength,longitudinal_rebar_yield_strength,rebar_effective_depths,rebar_cage_2_fy_override,rebar_cage_3_fy_override,shear_override_crit_depth,shear_crit_depth_override_comp,shear_crit_depth_override_uplift,bearing_toggle_type,tool_version,modified) OUTPUT INSERTED.ID INTO @DrilledPier VALUES ([INSERT ALL DRILLED PIER DETAILS])
			SELECT @DpID=DpID FROM @DrilledPier

			--Criteria 
			INSERT INTO pole.pole_analysis_criteria VALUES ([INSERT CRITERIA])

			--Unreinf Geometry
			INSERT INTO pole.pole_section VALUES ([INSERT ALL POLE SECTIONS])

			--Reinf Geometry
			INSERT INTO pole.pole_reinf_section VALUES ([INSERT ALL REINF POLE SECTIONS])

			--Reinf Groups
			INSERT INTO pole.pole_reinf_group VALUES ([INSERT ALL REINF GROUPS])

			--Reinf Details
			INSERT INTO pole.pole_reinf_details VALUES ([INSERT ALL REINF DETAILS])

			--Interference Groups
			INSERT INTO pole.pole_interference_group VALUES ([INSERT ALL INTERFERENCE GROUPS])

			--Interference Details
			INSERT INTO pole.pole_interference_details VALUES ([INSERT ALL INTERFERENCE DETAILS])

			--Reinf Results
			INSERT INTO pole.pole_reinf_results VALUES ([INSERT ALL REINF RESULTS])
	
			--Custom Reinf Properties
			INSERT INTO pole.memb_prop_flat_plate VALUES ([INSERT ALL REINF PROPERTIES])
	
			--Custom Bolt Properties
			INSERT INTO pole.bolt_prop_flat_plate VALUES ([INSERT ALL BOLT PROPERTIES])

			--Custom Matl Properties
			INSERT INTO pole.matl_prop_flat_plate VALUES ([INSERT ALL MATL PROPERTIES])

			--SELECT EXISTING FOUNDATION GROUP ID
			SELECT @FndgrpID=foundation_group_id FROM gen.structure_model WHERE ID=@ModelID
			
			--INSERT Foundation ID into Foundation Details
			INSERT INTO fnd.foundation_details (foundation_group_id,foundation_type,details_id) OUTPUT INSERTED.ID INTO @Foundation VALUES (@FndgrpID,@FndType,@DpID)
			SELECT @FndID=FndID FROM @Foundation

		END

