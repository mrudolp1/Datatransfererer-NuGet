--Pier and Pad Sub Insert
BEGIN
	SET @FndID = [FOUNDATION ID]
	SET @FndType = [FOUNDATION TYPE]
	SET @GuyGrpID = [GUY GROUP ID]
	IF @FndID IS NULL 
		BEGIN
			DELETE FROM @Fnd
			INSERT INTO fnd.pier_pad_details (pier_shape
						,pier_diameter
						,extension_above_grade
						,pier_rebar_size
						,pier_tie_size
						,pier_tie_quantity
						,pier_reinforcement_type
						,pier_clear_cover
						,foundation_depth
						,pad_width_1
						,pad_width_2
						,pad_thickness
						,pad_rebar_size_top_dir1
						,pad_rebar_size_bottom_dir1
						,pad_rebar_size_top_dir2
						,pad_rebar_size_bottom_dir2
						,pad_rebar_quantity_top_dir1
						,pad_rebar_quantity_bottom_dir1
						,pad_rebar_quantity_top_dir2
						,pad_rebar_quantity_bottom_dir2
						,pad_clear_cover
						,rebar_grade
						,concrete_compressive_strength
						,dry_concrete_density
						,total_soil_unit_weight
						,bearing_type
						,nominal_bearing_capacity
						,cohesion
						,friction_angle
						,spt_blow_count
						,base_friction_factor
						,neglect_depth
						,bearing_distribution_type
						,groundwater_depth
						,top_and_bottom_rebar_different
						,block_foundation
						,rectangular_foundation
						,base_plate_distance_above_foundation
						,bolt_circle_bearing_plate_width
						,pier_rebar_quantity
						,basic_soil_check
						,structural_check
						,tool_version,
						modified) OUTPUT INSERTED.ID INTO @Fnd VALUES ([FOUNDATION VALUES])
			SELECT @FndID = FndID FROM @Fnd
		END
	INSERT INTO fnd.foundation_details (foundation_group_id, foundation_type, guy_group_id, details_id) VALUES (@FndGrpID, @FndType, @GuyGrpID, @FndID)
	INSERT INTO @Fndlist (FndGrpID, FndIndex, FndType, FndID) VALUES (@FndGrpID, [i], @FndType, @FndID)
END