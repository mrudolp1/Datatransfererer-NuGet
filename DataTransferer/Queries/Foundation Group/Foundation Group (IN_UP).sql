--Model Declarations
DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
--Foundatation Declarations
DECLARE @FndGrp TABLE(FndgrpID INT)
DECLARE @FndGrpID INT
DECLARE @Fnd TABLE(FndID INT)
DECLARE @FndID INT
DECLARE @FndType VARCHAR(255)
DECLARE @GuyGrpID INT
DECLARE @FndList Table(FndGrpID INT, FndIndex INT,FndType VARCHAR(255),FndID INt)
DECLARE @FndListID INT

--Minimum information needed to insert a new model into structure_model
	SET @BU = [BU NUMBER]
	SET @STR_ID = [STRUCTURE ID]

Begin
	SET @FndGrpID = [FOUNDATION GROUP ID] --This will either exist in .NET or be set to NULL .REPLACE("'[TNX ID]'","NULL")
	If @FndGrpID is Null
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
		Begin
			INSERT INTO fnd.foundation_group OUTPUT Inserted.ID INTO @FndGrp DEFAULT VALUES
			SELECT @FndGrpID=FndGrpID FROM @FndGrp
			UPDATE gen.structure_model SET foundation_group_id = @FndGrpID WHERE ID = @ModelID

			[FOUNDATIONS]

			SELECT * FROM @FndList
		END
END
