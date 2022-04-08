--Pier and Pad Partial Upload
---Use specific foundation partial queries instead of this generic one. - DHS
BEGIN
	SET @FndID = [FOUNDATION ID]
	SET @FndType = [FOUNDATION TYPE]
	SET @GuyGrpID = [GUY GROUP ID]
	IF @FndID IS NULL 
		BEGIN
			DELETE FROM @Fnd
			INSERT INTO [FOUNDATION TABLE] ([FOUNDATION COLUMNS]) OUTPUT INSERTED.ID INTO @Fnd VALUES ([FOUNDATION VALUES])
			SELECT @FndID = FndID FROM @Fnd
		END
	INSERT INTO fnd.foundation_details (foundation_group_id, foundation_type, guy_group_id, details_id) VALUES (@FndGrpID, @FndType, @GuyGrpID, @FndID)
	INSERT INTO @Fndlist VALUES (@FndGrpID, [i], @FndType, @FndID)
END