--Structure Info Declarations
--This line is a test. Did you pass?
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
	--Select * From fnd.anchor_block WHERE bus_unit=@BU AND structure_id=@strID

	--Soil
	INSERT INTO @SoilProfileIDs 
	SELECT soil_profile_id sp_id FROM fnd.drilled_pier WHERE bus_unit = @BU AND structure_id = @strID
	UNION ALL
	SELECT soil_profile_id sp_id FROM fnd.pile WHERE bus_unit = @BU AND structure_id = @strID
	--UNION ALL
	--SELECT soil_profile_id sp_id FROM fnd.anchor_block WHERE bus_unit = @BU AND structure_id = @strID

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
	SELECT * FROM pole.pole WHERE bus_unit=@BU AND structure_id=@strID

	--CCIpole Unreinf Sections
	SELECT us.* 
	FROM 
		pole.sections us
		,pole.pole pole
	WHERE
		pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = us.pole_id

	--CCIpole Reinf Sections
	SELECT rs.* 
	FROM 
		pole.reinforced_sections rs
		,pole.pole pole
	WHERE
		pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = rs.pole_id

	--CCIpole Reinf Groups
	SELECT rg.* 
	FROM 
		pole.reinforcements rg
		,pole.pole pole
	WHERE
		pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = rg.pole_id

	--CCIpole Reinf Details
	SELECT rd.* 
	FROM 
		pole.reinforcement_details rd
		,pole.reinforcements rg
		,pole.pole pole
	WHERE
		pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = rg.pole_id
		AND rg.ID = rd.group_id

	--CCIpole Interference Groups
	SELECT ig.* 
	FROM 
		pole.interferences ig
		,pole.pole pole
	WHERE
		pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = ig.pole_id

	--CCIpole Interference Details
	SELECT id.* 
	FROM 
		pole.interference_details id
		,pole.interferences ig
		,pole.pole pole
	WHERE
		pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = ig.pole_id
		AND ig.ID = id.group_id

	--Create TempTables of Matls referenced in analysis
	SELECT DISTINCT reinfs.* INTO #TempReinfs
	FROM 
		gen.pole_reinforcements reinfs
		,pole.reinforcements rg
		,pole.pole pole
	WHERE pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = rg.pole_id
		AND reinfs.ID = rg.reinf_id

	SELECT DISTINCT matls.*  INTO #TempSection
	FROM 
		gen.pole_matls matls
		,pole.sections us
		,pole.pole pole
	WHERE
		pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = us.pole_id
		AND us.matl_id = matls.ID

	SELECT DISTINCT matls.*  INTO #TempRSection
	FROM 
		gen.pole_matls matls
		,pole.reinforced_sections rs
		,pole.pole pole
	WHERE
		pole.bus_unit = @BU
		AND pole.structure_id = @strID
		AND pole.ID = rs.pole_id
		AND rs.matl_id = matls.ID
	
	--CCIPole Matls
	SELECT DISTINCT matls.*
	FROM 
		gen.pole_matls matls
		,#TempSection sec
		,#TempRSection rsec
		,#TempReinfs reinfs
	WHERE matls.ID = sec.ID
		OR matls.ID = rsec.ID
		Or matls.ID = reinfs.matl_id

	--CCIPole Bolts
	SELECT DISTINCT bolts.* 
	FROM 
		gen.pole_bolts bolts
		,#TempReinfs reinfs
	WHERE bolts.ID = reinfs.bolt_id_top
		OR bolts.ID = reinfs.bolt_id_bot

	--CCIPole Reinfs
	SELECT * FROM #TempReinfs

	Drop Table #TempReinfs
	Drop Table #TempSection
	Drop Table #TempRSection


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

	--Leg Reinforcement
	Select * From tnx.memb_leg_reinforcement WHERE bus_unit=@BU AND structure_id=@strID

	--Leg Reinforcement Details
	Select lrdet.* 
	From 
		tnx.memb_leg_reinforcement lr
		,tnx.memb_leg_reinforcement_details lrdet
	WHERE
		lr.bus_unit = @BU
		AND lr.structure_id = @strID
		AND lr.ID = lrdet.leg_reinforcement_id
END