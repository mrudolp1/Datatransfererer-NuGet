--Structure Info Declarations
DECLARE @BU VARCHAR(10)
DECLARE @STR_ID VARCHAR(10)

DECLARE @FndGrp TABLE(FndgrpID INT)
DECLARE @FndGrpID INT
DECLARE @Fnd TABLE(FndID INT)
DECLARE @FndID INT
DECLARE @FndType VARCHAR(255)
DECLARE @GuyGrpID INT
DECLARE @FndList Table(FndGrpID INT, FndIndex INT,FndType VARCHAR(255),FndID INt)
DECLARE @FndListID INT

Begin
	SET @FndGrpID = [FOUNDATION GROUP ID] --This will either exist in .NET or be set to NULL .REPLACE("'[TNX ID]'","NULL")
	If @FndGrpID is Null
		Begin
			INSERT INTO fnd.foundation_group OUTPUT Inserted.ID INTO @FndGrp DEFAULT VALUES
			SELECT @FndGrpID=FndGrpID FROM @FndGrp

			[FOUNDATIONS]

			SELECT * FROM @FndList
		END
END
