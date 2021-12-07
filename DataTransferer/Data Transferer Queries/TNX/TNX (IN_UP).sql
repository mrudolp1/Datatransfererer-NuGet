--gen.structure_model
----tnx.tnx_structure
------tnx.tnx_individual_inputs
------tnx.base_structure
------tnx.base_structure_xref

--Model Declarations
DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
--DECLARE @ModelNeeded BIT
--TNX Declarations
DECLARE @tnxID INT
DECLARE @TNX TABLE(tnxID INT)
DECLARE @isTNXStructureNeeded BIT
DECLARE @tnxInputID INT
DECLARE @tnxInput TABLE(tnxInputID INT)
--TNX Record Declarations
DECLARE @baseSectionID INT
DECLARE @baseSection TABLE(baseSectionID INT)
DECLARE @upperSectionID INT
DECLARE @upperSection TABLE(upperSectionID INT)
DECLARE @guyID INT
DECLARE @guy TABLE(guyID INT)

	--Minimum information needed to insert a new model into structure_model
	SET @BU = [BU NUMBER]
	SET @STR_ID = [STRUCTURE ID]
	--SET @ModelNeeded = '[Model ID Needed]'
BEGIN
	SET @tnxID = [TNX ID] --This will either exist in .NET or be set to NULL .REPLACE("'[TNX ID]'","NULL")
	SET @tnxInputID = [TNX INDIVIDUAL INPUT ID]
	If @tnxID is Null --There are changes in the tnx model, create new record
		--
		--Create new model, update old model to inactive if it exist
		--This needs to be moved to the Structure Model Class or we will be creating a new model for every upload query
		--
		IF EXISTS(SELECT * FROM gen.structure_model_xref WHERE bus_unit=@BU AND structure_id=@STR_ID) 
			BEGIN
				--If exists, select model_id from structure_model_xref
				INSERT INTO @Model (ModelID) SELECT model_id FROM gen.structure_model_xref WHERE bus_unit=@BU AND structure_id=@STR_ID AND isActive='True' --ORDER BY model_id
				SELECT @ModelID=ModelID FROM @Model
				--Update status to FALSE for existing model_id
				UPDATE gen.structure_model_xref Set isActive='False' WHERE model_id=@ModelID
				--Create new Model ID by copying previous data and pasting new row into Structure_model
				--INSERT INTO gen.structure_model (connection_group_id,foundation_group_id,guy_config_id,lattice_structure_id,pole_structure_id,critera_id,tnx_id) OUTPUT Inserted.id INTO @Model SELECT connection_group_id,foundation_group_id,guy_config_id,lattice_structure_id,pole_structure_id,critera_id,tnx_id FROM gen.structure_model WHERE id=@ModelID
				--SELECT @ModelID=ModelID FROM @Model
				--# Denotes a temporary table
				SELECT * INTO #TempTable FROM gen.structure_model WHERE ID=@ModelID
				ALTER TABLE #TempTable DROP COLUMN ID
				INSERT INTO gen.structure_model OUTPUT INSERTED.ID INTO @Model SELECT * FROM #TempTable
				DROP TABLE #TempTable
				SELECT @ModelID=ModelID FROM @Model
				--Create new row in structure_model_xref, associating BU to newly created Model ID
				INSERT INTO gen.structure_model_xref (model_id,bus_unit,structure_id,isActive) VALUES (@ModelID,@BU,@STR_ID,'True')

			END
		ELSE
			BEGIN
				-- Create new Model ID by adding row to Structure_model 
				INSERT INTO gen.structure_model OUTPUT Inserted.ID INTO @Model DEFAULT VALUES
				SELECT @ModelID=ModelID FROM @Model
				--Create new row in structure_model_xref, associating BU to newly created Model ID
				INSERT INTO gen.structure_model_xref (model_id,bus_unit,structure_id,isActive) VALUES (@ModelID,@BU,@STR_ID,'True')
			END--Select existing model ID or insert new
		--
		--Create new tnx record
		--
		BEGIN
			IF @tnxInputID IS NULL --individual inputs have changed, create a new record
				Begin
					INSERT INTO tnx.tnx_individual_inputs ([TNX INDIVIDUAL INPUT COLUMNS]) OUTPUT INSERTED.ID INTO @tnxInput VALUES ([TNX INDIVIDUAL INPUT VALUES])
					SELECT @tnxInputID = tnxInputID FROM @tnxInput
				End

			INSERT INTO tnx.tnx_structure OUTPUT INSERTED.ID INTO @TNX VALUES (@tnxInputID,[TNX FILE PATH])
			SELECT @tnxID=tnxID FROM @TNX
			UPDATE gen.structure_model SET tnx_id = @tnxID WHERE ID = @ModelID

			[BASE STRUCTURE]

			[UPPER STRUCTURE]

			[GUY LEVELS]
		END
END

--BEGIN
--	--Individual Inputs
--	SET @tnxInputID = '[TNX Individual Input ID]'
--	IF @tnxInputID IS NULL 
--		BEGIN
--			DELETE FROM @tnxInput 
--			INSERT INTO tnx.tnx_individual_inputs ('[TNX Individual Input Columns]') OUTPUT INSERTED.ID INTO @tnxInput VALUES ('[TNX Individual Input Values]')
--			SELECT @tnxInputID = tnxInputID FROM @tnxInput
--		END
--	INSERT INTO tnx.tnx_structure VALUES(@baseStructureID, @tnxID)
--END

--Repeat the following for each section of the tnx database
--Base Structure Sub Query START
--BEGIN
--	SET @baseStructureID = 1 --This will either exist in .NET or be set to NULL .REPLACE("'[BASE ID]'","NULL")
--	IF @baseStructureID IS NULL 
--		BEGIN
--			DELETE FROM @basestructure 
--			INSERT INTO tnx.base_structure (TowerRec) OUTPUT INSERTED.ID INTO @basestructure VALUES('[ALL BASE STRUCTURE VALUES]')
--			SELECT @baseStructureID = baseid FROM @basestructure 
--		END

--	INSERT INTO tnx.base_structure_xref VALUES(@baseStructureID, @tnxID)
--END
--Base Structure Sub Query END



--INSERT INTO tnx.tnx_structure DEFAULT VALUES
--SELECT * FROM tnx.base_structure 
--SELECT * FROM tnx.base_structure_xref 
--SELECT * FROM tnx.tnx_structure 

--POTENTIALLY USEFUL CODE. DON'T DELETE YET
--BEGIN
	--SELECT * INTO TempTable FROM tnx.base_structure WHERE ID=@baseStructureID
	--ALTER TABLE TempTable DROP COLUMN ID
	--INSERT INTO tnx.base_structure OUTPUT INSERTED.ID INTO @basestructure SELECT * FROM TempTable
	--DROP TABLE TempTable
--END