﻿
BEGIN
	INSERT INTO fnd.pier_pad ([FOUNDATION FIELDS]) 
	OUTPUT INSERTED.ID INTO @Prev
	VALUES([FOUNDATION VALUES])

	Set @IncResults = [INCLUDE RESULTS]
	IF @IncResults = 1
		BEGIN
			[RESULTS]
		END
	DELETE FROM @Prev
END