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
DECLARE @memberID INT
DECLARE @member TABLE(memberID INT)
DECLARE @materialID INT
DECLARE @material TABLE(materialID INT)

	--Minimum information needed to insert a new model into structure_model
	SET @BU = [BU NUMBER]
	SET @STR_ID = [STRUCTURE ID]
	--SET @ModelNeeded = '[Model ID Needed]'
BEGIN
	SET @tnxID = [TNX ID] --This will either exist in .NET or be set to NULL .REPLACE("'[TNX ID]'","NULL")
	If @tnxID is Null --There are changes in the tnx model, create new record
	--SET @tnxInputID = [TNX INDIVIDUAL INPUT ID]

	Begin
		INSERT INTO tnx.tnx ([TNX INDIVIDUAL INPUT COLUMNS], bus_unit, structure_id) OUTPUT INSERTED.ID INTO @TNX VALUES ([TNX INDIVIDUAL INPUT VALUES], @BU, @STR_ID)
		SELECT @tnxID = tnxID FROM @TNX
	End

	[BASE STRUCTURE]

	--[UPPER STRUCTURE]

	--[GUY LEVELS]

	--[MEMBERS]

	--[MATERIALS]

END
