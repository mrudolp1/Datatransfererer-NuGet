--structure_model -> P: model_id, foundation_group_id
--structure model xref -> model_id, bus_unit, structure_id
--foundation_group_id -> P: foundation_id, details_id, foundation_type, foundation_group_id, guy_group_id
--anchor_block_details -> P: anchor_id
--anchor_block_soil_layer -> P: layer_id, F: anchor_id
--guy profile. Similar logic to soil layer

--create model ID first
--associate BU and Structure to it in structure xref
--create foundation group id
--assign foundation group (how is this determined, folder in which all tools are saved?). For this query, just create new entries to insert for all cases. Must create foundation group first
--create foundations ids (P) for each foundaiton in the group (F)
--create anchor block details ID (P) group (F)
--retroactively add details_id to the foundation details ID (must have details ID created first)
--create soil layer ID (P). Accociate with anchor id (F)
--guy profile. Similar logic to soil layer


DECLARE @ModelID INT
DECLARE @Model TABLE(ModelID INT)

DECLARE @BU VARCHAR(10)

DECLARE @StrID VARCHAR(10)
DECLARE @StrType VARCHAR(50)

DECLARE @FndGrpID INT
DECLARE @FndGrp TABLE(FndGrpID INT)

DECLARE @FndID INT
DECLARE @Fnd TABLE(FndID INT)

DECLARE @FndType VARCHAR(255)

DECLARE @GabID INT
DECLARE @Gab TABLE(GabID INT)


SET @BU = '[BU Number]'
SET @StrID = '[STRUCTURE ID]'
SET @FndType = '[FOUNDATION TYPE]'
--SET @GabID = '[GUYED ANCHOR BLOCK ID]'

--CREATE ANCHOR DETAIL
--CREATE SOIL LAYERS ASSOCIATED WITH ANCHOR DETAIL
--CREATE PROFILES ASSOCIATED WITH ANCHOR DETAIL

------(DO THESE ONLY ONCE PER TOOL, NOT PER ANCHOR!!!!) ADD CODE IF FIRST THEN, ELSE------
--CREATE FOUNDATION GROUP 
--CREATE/UPDATE STRUCTURE MODEL
--CREATE/UPDATE STRUCTURE MODEL XREF

--CREATE FOUNDATION DETAILS, LINKING TO FOUNDATION GROUP AND ANCHOR DETAILS


--START ANCHOR DETAILS
BEGIN

	INSERT INTO fnd.anchor_block_details OUTPUT INSERTED.ID INTO @Gab VALUES ([INSERT ALL GUYED ANCHOR BLOCK DETAILS])
	SELECT @GabID=GabID FROM @Gab

	--SOIL LAYERS
	INSERT INTO fnd.anchor_block_soil_layer VALUES([INSERT ALL SOIL LAYERS])

	--PROFILES
	INSERT INTO fnd.anchor_block_profile VALUES([INSERT ALL PROFILES])

END
--END ANCHOR DETAILS


--START SITE-SPECIFIC
BEGIN
	
	--FOUNDATION GROUP (may need to edit to DEFAULT VALUES if inserting into table with only a primary key)
	INSERT INTO fnd.foundation_group OUTPUT INSERTED.ID INTO @FndGrp 
	SELECT @FndGrpID=FndGrpID FROM @FndGrp

	--STRUCUTRE MODEL (must incorporate with other code. Should not create a new structure model simply because of foundation information, must include all tables)
	INSERT INTO gen.structure_model OUTPUT INSERTED.ID INTO @Model
	SELECT @ModelID=ModelID FROM @Model

	--STRUCUTRE MODEL CROSS-REFERENCE (must incorporate with other code. Should not create a new BU simply because of foundation information, must include all tables)
	INSERT INTO gen.structure_model_xref VALUES(@ModelID,@BU,@StrID)

END
--END SITE-SPECIFIC

--START FOUNDATION DETAILS
BEGIN

	INSERT INTO fnd.foundation_details VALUES(@FndGrpID,@FndType,@GabID)

END
--END FOUNDATION DETAILS