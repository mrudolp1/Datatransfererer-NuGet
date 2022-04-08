--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)

	SET @BU = [BU NUMBER]
	SET @STR_ID = [STRUCTURE ID]

Begin
	SELECT
		ID
	FROM
		fnd.pier_pad pp
	WHERE
		pp.bus_unit=@BU
		AND pp.structure_id=@STR_ID
END