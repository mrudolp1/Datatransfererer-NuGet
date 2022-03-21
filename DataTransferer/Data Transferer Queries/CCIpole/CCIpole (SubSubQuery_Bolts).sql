BEGIN --Bolt DB SubSubQuery BEGIN

	SET @BoltID = '[BOLT ID]'

	IF @BoltID IS NULL
		BEGIN
			IF EXISTS(SELECT * FROM pole.bolt_prop_flat_plate WHERE local_id = '[local_id]' AND name = '[name]' AND description = '[description]' AND diam = '[diam]' AND area = '[area]' AND fu_bolt = '[fu_bolt]' AND sleeve_diam_out = '[sleeve_diam_out]' AND sleeve_diam_in = '[sleeve_diam_in]' AND fu_sleeve = '[fu_sleeve]' AND bolt_n_sleeve_shear_revF = '[bolt_n_sleeve_shear_revF]' AND bolt_x_sleeve_shear_revF = '[bolt_x_sleeve_shear_revF]' AND bolt_n_sleeve_shear_revG = '[bolt_n_sleeve_shear_revG]' AND bolt_x_sleeve_shear_revG = '[bolt_x_sleeve_shear_revG]' AND bolt_n_sleeve_shear_revH = '[bolt_n_sleeve_shear_revH]' AND bolt_x_sleeve_shear_revH = '[bolt_x_sleeve_shear_revH]' AND rb_applied_revH = '[rb_applied_revH]') 
				SELECT @BoltID = ID FROM pole.bolt_prop_flat_plate WHERE local_id = '[local_id]' AND name = '[name]' AND description = '[description]' AND diam = '[diam]' AND area = '[area]' AND fu_bolt = '[fu_bolt]' AND sleeve_diam_out = '[sleeve_diam_out]' AND sleeve_diam_in = '[sleeve_diam_in]' AND fu_sleeve = '[fu_sleeve]' AND bolt_n_sleeve_shear_revF = '[bolt_n_sleeve_shear_revF]' AND bolt_x_sleeve_shear_revF = '[bolt_x_sleeve_shear_revF]' AND bolt_n_sleeve_shear_revG = '[bolt_n_sleeve_shear_revG]' AND bolt_x_sleeve_shear_revG = '[bolt_x_sleeve_shear_revG]' AND bolt_n_sleeve_shear_revH = '[bolt_n_sleeve_shear_revH]' AND bolt_x_sleeve_shear_revH = '[bolt_x_sleeve_shear_revH]' AND rb_applied_revH = '[rb_applied_revH]'
			ELSE
				BEGIN
					INSERT INTO pole.bolt_prop_flat_plate OUTPUT INSERTED.ID INTO @PropBolt VALUES ('[INSERT BOLT PROP]')
					SELECT @BoltID = BoltID FROM @PropBolt
				END		
		END


END --Bolt DB SubSubQuer END

--'[BOLT SUB-SUBQUERY]'