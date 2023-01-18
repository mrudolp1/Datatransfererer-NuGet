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

	--Pile Location
	Select pl.* 
	From 
		fnd.pile_location pl
		,fnd.pile fnd
	WHERE
		fnd.bus_unit = @BU
		AND fnd.structure_id = @strID
		AND fnd.ID = pl.pile_id

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

	--File Upload
	SELECT * FROM gen.file_upload fu, (SELECT MAX(work_order_seq_num) work_order_seq_num FROM gen.work_orders WHERE bus_unit = @BU AND structure_id = @strID) wo WHERE fu.work_order_seq_num = wo.work_order_seq_num

	--Drilled Pier
	Select dp.* From fnd.drilled_pier dp WHERE (dp.bus_unit=@BU AND dp.structure_id=@strID)

	--Drilled Pier Profile
	SELECT dpp.* FROM fnd.drilled_pier_profile dpp, fnd.drilled_pier dp WHERE (dp.bus_unit=@BU AND dp.structure_id=@strID) AND dpp.id = dp.drilled_pier_profile_id
	
	--Drilled Pier Section
	SELECT dps.* FROM fnd.drilled_pier_section dps, fnd.drilled_pier_profile dpp, fnd.drilled_pier dp WHERE (dp.bus_unit=@BU AND dp.structure_id=@strID) AND dps.drilled_pier_profile_id = dpp.id AND dpp.id = dp.drilled_pier_profile_id

	--Drilled Pier Rebar
	SELECT dpr.* FROM fnd.drilled_pier_rebar dpr, fnd.drilled_pier_section dps ,fnd.drilled_pier_profile dpp, fnd.drilled_pier dp WHERE (dp.bus_unit=@BU AND dp.structure_id=@strID) AND dps.id = dpr.section_id AND dps.drilled_pier_profile_id = dpp.id AND dpp.id = dp.drilled_pier_profile_id
	
	--Belled Pier
	SELECT bp.* FROM fnd.belled_pier bp, fnd.drilled_pier_profile dpp, fnd.drilled_pier dp WHERE (dp.bus_unit=@BU AND dp.structure_id=@strID) AND bp.drilled_pier_profile_id = dpp.id AND dpp.id = dp.drilled_pier_profile_id
	
	--Embedded Pole
	SELECT ep.* FROM fnd.embedded_pole ep, fnd.drilled_pier_profile dpp, fnd.drilled_pier dp WHERE (dp.bus_unit=@BU AND dp.structure_id=@strID) AND ep.drilled_pier_profile_id = dpp.id AND dpp.id = dp.drilled_pier_profile_id

	--Drilled Pier Tool
	Select * From fnd.drilled_pier_tool WHERE bus_unit=@BU AND structure_id=@strID
END