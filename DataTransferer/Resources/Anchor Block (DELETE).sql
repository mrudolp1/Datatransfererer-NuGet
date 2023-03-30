
BEGIN

	--Soil Layers
	DELETE FROM fnd.soil_layer WHERE soil_profile_id = [SOIL PROFILE ID]
	
	--Soil Profiles
	DELETE FROM fnd.soil_profile WHERE ID = [SOIL PROFILE ID]

	--Anchor Profiles
	DELETE FROM fnd.anchor_block_profile WHERE ID = [ANCHOR PROFILE ID]

	--Anchor Results
	DELETE FROM fnd.anchor_block_results WHERE anchor_block_id = [ANCHOR ID]

	--Anchor Block
	DELETE FROM fnd.anchor_block WHERE ID = [ANCHOR ID]

END