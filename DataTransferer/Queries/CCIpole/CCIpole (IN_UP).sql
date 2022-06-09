--Model Declarations
DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
DECLARE @ModelNeeded BIT
--CCIpole Declarations
DECLARE @PoleStructure TABLE(PoleID INT)
DECLARE @PoleID INT
DECLARE @PoleNeeded BIT
DECLARE @Criteria TABLE (CriteriaID INT)
DECLARE @CriteriaID INT
--Other CCIpole Declarations
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
	DECLARE @BotBoltID INT
	DECLARE @TopBoltID INT
DECLARE @PropMatl TABLE (MatlID INT)
DECLARE @MatlID INT


	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'
	Set @ModelNeeded = '[Model ID Needed]'

	--ID will need passed in. Either a number or the text NULL without quotes
	SET @PoleID = '[CCIpole ID]'
	SET @CriteriaID = '[Pole Criteria ID]'
	Set @PoleNeeded = '[CCIpole ID Needed]'


--Determine model_id (Table Impacts: gen.structure_model_xref & gen.structure_model)
IF EXISTS(SELECT * FROM gen.structure_model_xref WHERE bus_unit=@BU AND structure_id=@STR_ID) 
	BEGIN
		--If BU/StructureID exists within EDS, select the model_id from structure_model_xref
		INSERT INTO @Model (ModelID) SELECT model_id FROM gen.structure_model_xref WHERE bus_unit=@BU AND structure_id=@STR_ID AND isActive='True' --ORDER BY model_id
		SELECT @ModelID=ModelID FROM @Model
		--If changes occurred with any field in structure_model table, create new model ID for reference
		IF @ModelNeeded = 1 --TRUE (Reference ismodelneeded)
			BEGIN
				--Update status to FALSE for existing model_id
				UPDATE gen.structure_model_xref Set isActive='False' WHERE model_id=@ModelID
				----Create new Model ID by copying previous data and pasting new row into Structure_model -> Replaced with TempTable code below
				--INSERT INTO gen.structure_model (connection_group_id,foundation_group_id,guy_config_id,lattice_structure_id,pole_structure_id,critera_id, tnx_id) OUTPUT Inserted.id INTO @Model SELECT connection_group_id,foundation_group_id,guy_config_id,lattice_structure_id,pole_structure_id,critera_id, tnx_id FROM gen.structure_model WHERE id=@ModelID
	
				--Delete temp table if already exists
				IF OBJECT_ID(N'tempdb..#TempTable') IS NOT NULL
					DROP TABLE #TempTable
					
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
		-- No previous BU/StructureID in EDS > Create new Model ID by adding row to Structure_model 
		INSERT INTO gen.structure_model OUTPUT Inserted.ID INTO @Model DEFAULT VALUES
		SELECT @ModelID=ModelID FROM @Model
		--Create new row in structure_model_xref, associating BU to newly created Model ID
		INSERT INTO gen.structure_model_xref (model_id,bus_unit,structure_id,isActive) VALUES (@ModelID,@BU,@STR_ID,'True')
	END--Select existing model ID or insert new

---------------------------------------------------------------------------------------------------------------------------------------------------

--Determine pole_structure_id (Table Impacts: gen.structure_model & pole.pole_structure & pole.pole_section & pole.pole_reinf_section & pole.pole_reinf_group & pole.pole_interference_group & pole.pole_reinf_results & pole.memb_prop_flat_plate & pole.bolt_prop_flat_plate & pole.matl_prop_flat_plate)
IF @PoleNeeded = 1 --TRUE (Reference isPoleNeeded)

	BEGIN 
		--Create new CCIpole ID in pole_structure
		INSERT INTO pole.pole_structure OUTPUT INSERTED.ID INTO @PoleStructure DEFAULT VALUES
		SELECT @PoleID = PoleID FROM @PoleStructure
		UPDATE gen.structure_model SET pole_structure_id = @PoleID WHERE ID = @ModelID

		--Create new Pole Criteria ID in pole_analysis_criteria
		--CriteriaDNU IF @CriteriaID IS NULL
		--CriteriaDNU 	BEGIN
		--CriteriaDNU 		IF EXISTS(SELECT * FROM pole.pole_analysis_criteria WHERE upper_structure_type = '[upper_structure_type]' AND analysis_deg = '[analysis_deg]' AND geom_increment_length = '[geom_increment_length]' AND vnum = '[vnum]' AND check_connections = '[check_connections]' AND hole_deformation = '[hole_deformation]' AND ineff_mod_check = '[ineff_mod_check]' AND modified = '[modified]')
		--CriteriaDNU 			SELECT @CriteriaID = ID FROM pole.pole_analysis_criteria WHERE upper_structure_type = '[upper_structure_type]' AND analysis_deg = '[analysis_deg]' AND geom_increment_length = '[geom_increment_length]' AND vnum = '[vnum]' AND check_connections = '[check_connections]' AND hole_deformation = '[hole_deformation]' AND ineff_mod_check = '[ineff_mod_check]' AND modified = '[modified]'
		--CriteriaDNU 		ELSE
		--CriteriaDNU 			BEGIN
		--CriteriaDNU 				INSERT INTO pole.pole_analysis_criteria OUTPUT INSERTED.ID INTO @Criteria VALUES ('[INSERT POLE CRITERIA]')
		--CriteriaDNU 				SELECT @CriteriaID = CriteriaID FROM @Criteria
		--CriteriaDNU 			END
		--CriteriaDNU 	END

		UPDATE pole.pole_structure SET criteria_id = @CriteriaID WHERE ID = @PoleID

		--'[SUBQUERY]'

	END	

				
	----Custom Matl Properties
	--INSERT INTO pole.matl_prop_flat_plate VALUES ('[INSERT ALL MATL PROPERTIES]')
	----Custom Bolt Properties
	--INSERT INTO pole.bolt_prop_flat_plate VALUES ('[INSERT ALL BOLT PROPERTIES]')
	----Custom Reinf Properties
	--INSERT INTO pole.memb_prop_flat_plate VALUES ('[INSERT ALL REINF PROPERTIES]')

	----Reinf Results
	--INSERT INTO pole.pole_reinf_results VALUES ('[INSERT ALL REINF RESULTS]')