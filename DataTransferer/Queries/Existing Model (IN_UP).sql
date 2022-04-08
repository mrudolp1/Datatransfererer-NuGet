DECLARE @Model TABLE(ModelID INT)
DECLARE @ModelID INT
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)
DECLARE @STR_TYPE VARCHAR(50)

	--Minimum information needed to insert a new model into structure_model
	SET @BU = '[BU NUMBER]'
	SET @STR_ID = '[STRUCTURE ID]'

	--Select existing model ID or insert new
	BEGIN
		IF EXISTS(SELECT * FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True') 
			INSERT INTO @Model (ModelID) SELECT ID FROM structure_model WHERE bus_unit=@BU AND structure_id=@STR_ID AND existing_geometry='True'
		ELSE
			INSERT INTO structure_model (bus_unit,structure_id,existing_geometry) OUTPUT INSERTED.ID INTO @Model VALUES (@BU,@STR_ID,'True')
	END

	SELECT @ModelID=ModelID FROM @Model
