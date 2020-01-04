IF EXISTS (
  SELECT * 
    FROM SYS.all_objects O
	JOIN sys.schemas S ON S.schema_id = O.schema_id
   WHERE S.name = N'DBO'
     AND O.name = N'sp_ImportFromExcel' 
	 AND O.type = 'P'
)
   DROP PROCEDURE DBO.sp_ImportFromExcel
GO

CREATE PROCEDURE DBO.sp_ImportFromExcel
	@Excel VARCHAR(MAX),
	@DesiredTableName VARCHAR(MAX) = ImportedFromExcel
AS

DECLARE @Error VARCHAR(MAX) = N'
The Process Failed!
Table Already Exists. Either Select a New Name for your table or Drop ' + @DesiredTableName + '. ---At your own Risk!---'

IF EXISTS (SELECT *
		  FROM INFORMATION_SCHEMA.TABLES T
		  WHERE T.TABLE_NAME = @DesiredTableName)
    BEGIN
	   RAISERROR(@Error, 1, 1, 1)
    END
ELSE BEGIN
	DECLARE @Hdrs NVARCHAR(MAX);
	DECLARE @CreateTable NVARCHAR(MAX);
	DECLARE @Data NVARCHAR(MAX);
	DECLARE @ModifiedExcel NVARCHAR(MAX);
	DECLARE @InsertStatement NVARCHAR(MAX);
	DECLARE @SelectStatement NVARCHAR(MAX);

	SET @ModifiedExcel = CHAR(39) +
					REPLACE(REPLACE(@Excel, 
								 CHAR(9), 
								 CHAR(39) + CHAR(9) + CHAR(39)
						   ),
						   CHAR(13) + CHAR(10),
						   CHAR(39) + CHAR(13) + CHAR(10) + CHAR(39)
					)


	SET @Hdrs = LEFT(@Excel,
				 CHARINDEX(CHAR(13),
						   @Excel) - 1)

	SET @CreateTable = 
	'CREATE TABLE ' + @DesiredTableName  + '
	([' + REPLACE(@Hdrs,
				CHAR(9), 
				'] VARCHAR(MAX),' + CHAR(13) + '[') + '] VARCHAR(MAX))'

	SET @Data = REPLACE('(' + 
						REPLACE(
						   REPLACE(
							  REPLACE(
								 @ModifiedExcel,
								 CHAR(9),
								 ','),
							CHAR(10),
							'),('),
						CHAR(13),''),
						'),(',
						'),' + CHAR(13) + '(')

	SET @Data = LEFT(@Data,
				 LEN(@Data) - 3)

	SET @InsertStatement = 
		'INSERT ' + @DesiredTableName +' 
	SELECT * 
	FROM (
		VALUES ' + 
		   LEFT(
			  RIGHT(@Data,
					LEN(@Data) - CHARINDEX(CHAR(13), 
										@Data)),
			  LEN(RIGHT(@Data,
						LEN(@Data) - CHARINDEX(CHAR(13),
										   @Data))) - 1) +
		') V(['+
		   REPLACE(@Hdrs,
				 CHAR(9), 
				 '],[') +
		'])'

	EXEC sp_sqlexec @CreateTable
	EXEC sp_sqlexec @InsertStatement
	SET @SelectStatement = '
		SELECT *
		FROM ' + @DesiredTableName

	EXEC sp_sqlexec @SelectStatement
	 
END