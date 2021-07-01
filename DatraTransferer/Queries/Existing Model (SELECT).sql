DECLARE @ModelID INT
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)

SET @BU = '[BU NUMBER]'
SET @STR_ID = '[STRUCTURE_ID]'
SELECT @ModelID=ID FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True'

