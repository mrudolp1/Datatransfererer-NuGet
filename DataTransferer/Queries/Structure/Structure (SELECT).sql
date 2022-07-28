--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @strID VARCHAR(10)
DECLARE @SoilProfileIDs TABLE(soil_profile_id INT)

SET @BU = [BU]
SET @strID = [STRID]

Begin
	--TNX
	Select * From tnx.tnx WHERE bus_unit=@BU AND structure_id=@strID

	--Base Structure
	Select bs.* 
	From 
		tnx.base_structure_sections bs
		,tnx.tnx tnx
	WHERE
		tnx.bus_unit = @BU
		AND tnx.structure_id = @strID
		AND tnx.ID = bs.tnx_ID

	--Upper Structure
	Select us.* 
	From 
		tnx.upper_structure_sections us
		,tnx.tnx tnx
	WHERE
		tnx.bus_unit = @BU
		AND tnx.structure_id = @strID
		AND tnx.ID = us.tnx_ID

	--Guys
	Select guys.* 
	From 
		tnx.guys guys
		,tnx.tnx tnx
	WHERE
		tnx.bus_unit = @BU
		AND tnx.structure_id = @strID
		AND tnx.ID = guys.tnx_ID

	--Members
	Select mb.* 
	From 
		tnx.members mb
		,tnx.members_xref xmb
		,tnx.tnx tnx
	WHERE
		tnx.bus_unit = @BU
		AND tnx.structure_id = @strID
		AND tnx.ID = xmb.tnx_ID
		AND xmb.member_id = mb.ID

	--Materials
	Select mt.* 
	From 
		tnx.materials mt
		,tnx.materials_xref xmt
		,tnx.tnx tnx
	WHERE
		tnx.bus_unit = @BU
		AND tnx.structure_id = @strID
		AND tnx.ID = xmt.tnx_ID
		AND xmt.material_id = mt.ID

	--Pier and Pad
	Select * From fnd.pier_pad WHERE bus_unit=@BU AND structure_id=@strID

	--Unit Base
	Select * From fnd.unit_base WHERE bus_unit=@BU AND structure_id=@strID

	--Pile
	Select * From fnd.pile WHERE bus_unit=@BU AND structure_id=@strID

	--Drilled Pier
	Select * From fnd.drilled_pier WHERE bus_unit=@BU AND structure_id=@strID

	--Anchor Block
	Select * From fnd.anchor_block WHERE bus_unit=@BU AND structure_id=@strID

	--Soil
	INSERT INTO @SoilProfileIDs 
	SELECT soil_profile_id sp_id FROM fnd.drilled_pier WHERE bus_unit = @BU AND structure_id = @strID
	UNION ALL
	SELECT soil_profile_id sp_id FROM fnd.pile WHERE bus_unit = @BU AND structure_id = @strID
	UNION ALL
	SELECT soil_profile_id sp_id FROM fnd.anchor_block WHERE bus_unit = @BU AND structure_id = @strID

	--Soil Profiles
	Select sp.* 
	From 
		fnd.soil_profile sp
		,@SoilProfileIDs fnd
	WHERE
		fnd.soil_profile_id = sp.id

	--Soil Layers
	Select sl.* 
	From 
		fnd.soil_layer sl
		,@SoilProfileIDs fnd
	WHERE
		fnd.soil_profile_id = sl.soil_profile_id

	--CCIPlate
	Select * From conn.connections WHERE bus_unit=@BU AND structure_id=@strID

	--CCIPole
	Select * From pole.pole WHERE bus_unit=@BU AND structure_id=@strID

	--Site code criteria
	SELECT * FROM gen.site_code_criteria WHERE bus_unit = @BU

END