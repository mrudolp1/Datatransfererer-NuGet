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

--TNX Record Declarations
DECLARE @baseSectionID INT
DECLARE @baseSection TABLE(baseSectionID INT)

	--SET @ModelNeeded = '[Model ID Needed]'
BEGIN
	[BASE STRUCTURE]
END
