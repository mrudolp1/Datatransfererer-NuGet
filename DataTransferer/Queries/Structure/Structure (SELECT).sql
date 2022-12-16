﻿--Structure Info Declarations
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

	--Pile Location
	Select pl.* 
	From 
		fnd.pile_location pl
		,fnd.pile fnd
	WHERE
		fnd.bus_unit = @BU
		AND fnd.structure_id = @strID
		AND fnd.ID = pl.pile_id

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

	--Connections
	Select pl.* 
	From 
		conn.plates pl
		,conn.connections con
	WHERE
		con.bus_unit = @BU
		AND con.structure_id = @strID
		AND con.ID = pl.connection_id

	--Plate Details
	Select pd.* 
	From 
		conn.plate_details pd
		,conn.plates pl
		,conn.connections con
	WHERE
		con.bus_unit = @BU
		AND con.structure_id = @strID
		AND con.ID = pl.connection_id
		AND pl.ID = pd.plate_id

	--Bolt Groups
	Select bg.* 
	From 
		conn.bolts bg
		,conn.plates pl
		,conn.connections con
	WHERE
		con.bus_unit = @BU
		AND con.structure_id = @strID
		AND con.ID = pl.connection_id
		AND pl.ID = bg.plate_id

	--Bolt Details
	Select bd.* 
	From 
		conn.bolt_details bd
		,conn.bolts bg
		,conn.plates pl
		,conn.connections con
	WHERE
		con.bus_unit = @BU
		AND con.structure_id = @strID
		AND con.ID = pl.connection_id
		AND pl.ID = bg.plate_id
		AND bg.ID = bd.bolt_id

	--CCIplate Materials
	Select * From gen.connection_material_properties

	--Stiffener Groups
	Select sg.* 
	From 
		conn.stiffeners sg
		,conn.plates pl
		,conn.connections con
		,conn.plate_details pd
	WHERE
		con.bus_unit = @BU
		AND con.structure_id = @strID
		AND con.ID = pl.connection_id
		AND pl.ID = pd.plate_id
		AND pd.ID = sg.plate_details_id

	--Stiffener Details
	Select sd.* 
	From 
		conn.stiffener_details sd
		,conn.stiffeners sg
		,conn.plates pl
		,conn.connections con
		,conn.plate_details pd
	WHERE
		con.bus_unit = @BU
		AND con.structure_id = @strID
		AND con.ID = pl.connection_id
		AND pl.ID=pd.plate_id
		AND pd.ID = sg.plate_details_id
		AND sg.ID=sd.stiffener_id

	--Bridge Stiffener Details
	Select bsd.* 
	From 
		conn.bridge_stiffeners bsd
		,conn.plates pl
		,conn.connections con
	WHERE
		con.bus_unit = @BU
		AND con.structure_id = @strID
		AND con.ID = pl.connection_id
		AND pl.ID=bsd.plate_id

	--CCIPole
	Select * From pole.pole WHERE bus_unit=@BU AND structure_id=@strID

	--Site code criteria
	SELECT * FROM gen.site_code_criteria WHERE bus_unit = @BU

END