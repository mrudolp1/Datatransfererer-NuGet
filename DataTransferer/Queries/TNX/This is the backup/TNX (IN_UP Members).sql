BEGIN
	SET @memberID = [MEMBER ID]
	IF @memberID IS NULL 
		BEGIN
			DELETE FROM @member 
			INSERT INTO tnx.members OUTPUT INSERTED.ID INTO @member VALUES([ALL MEMBER VALUES])
			SELECT @memberID = memberID FROM @member
		END
	INSERT INTO tnx.members_xref VALUES(@memberID, @tnxID)
END