
BEGIN --Unreinf Section SubQuery BEGIN

	--Material DB ID	
	SET @SubLevel4ID = [MATL ID]

	--MatlDNU IF @SubLevel4ID IS NULL
	--MatlDNU 	BEGIN
	--MatlDNU 		IF EXISTS(SELECT * FROM pole.pole_matls WHERE local_matl_id = '[local_matl_id]' AND name = '[name]' AND fy = '[fy]' AND fu = '[fu]' AND ind_default = '[ind_default]') 
	--MatlDNU 			SELECT @SubLevel4ID = ID FROM pole.pole_matls WHERE local_matl_id = '[local_matl_id]' AND name = '[name]' AND fy = '[fy]' AND fu = '[fu]' AND ind_default = '[ind_default]'
	--MatlDNU 		ELSE
	--MatlDNU 			BEGIN
	--MatlDNU 				INSERT INTO pole.pole_matls OUTPUT INSERTED.ID INTO @SubLevel4 VALUES ([MATL PROP VALUES])
	--MatlDNU 				SELECT @SubLevel4ID = ID FROM @SubLevel4
	--MatlDNU 			END
	--MatlDNU 	END

	--Section ID	
	INSERT INTO pole.sections ([UNREINF SECTION FIELDS]) 
		OUTPUT INSERTED.ID INTO @SubLevel1
		VALUES([UNREINF SECTION VALUES])
		SELECT @SubLevel1ID=ID FROM @SubLevel1


END --Unreinf Section SubQuery END

--[UNREINF SECTION INSERT]