DECLARE @ModelID INT
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)

SET @BU = '[BU NUMBER]'
SET @STR_ID = '[STRUCTURE_ID]'
--selecting most recent model_id associated to BU. (Eventually will need to update to include status (proposed/existing)
SELECT @ModelID=model_id FROM gen.structure_model_xref WHERE bus_unit=@BU AND structure_id=@STR_ID and isActive='True'
--ORDER BY
--model_id --Noticed inserted rows don't always get added in ascending order which is why Order By is included. 
