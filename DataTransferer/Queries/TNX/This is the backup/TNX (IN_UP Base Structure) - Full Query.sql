INSERT INTO tnx.base_structure VALUES (@baseSectionValues)
SELECT SCOPE_IDENTITY()
		--SELECT @baseSectionID = baseSectionID FROM @baseSection

--INSERT INTO tnx.base_structure_xref VALUES(@baseSectionID, @tnxID)s