ALTER PROCEDURE [USP_QUERY_TO_INSERT_INTO]
(
	@InQuery NVARCHAR(MAX)
   ,@DestinyTable VARCHAR(300)='[DestinyTable_Here]'
   ,@ExportPath VARCHAR(300)=NULL
)
AS
DECLARE @NewQuery NVARCHAR(MAX)
	   ,@ColumnTable VARCHAR (3000)=''
	   ,@ValuesColumn VARCHAR(5000)=''
	   ,@ExportQuery NVARCHAR(3000)=''
BEGIN TRY
	
	 IF OBJECT_ID('tempdb..##QUERY_TABLE') IS NOT NULL DROP TABLE ##QUERY_TABLE
	 SELECT @NewQuery= STUFF(@InQuery,CHARINDEX('FROM',@InQuery),0,'INTO ##QUERY_TABLE ')
	 EXEC sp_executesql @NewQuery

	 SET @NewQuery=''

	 SELECT @ColumnTable=COALESCE(@ColumnTable+'+'',''+','')+''''+QUOTENAME(TCOL.COLUMN_NAME)+''''
	 FROM tempdb.INFORMATION_SCHEMA.COLUMNS AS TCOL
	 WHERE TCOL.TABLE_NAME='##QUERY_TABLE'

	 SELECT @ColumnTable=SUBSTRING(@ColumnTable,6,LEN(@ColumnTable))

	 SELECT @ValuesColumn=COALESCE(@ValuesColumn+'+'',''+','')+
						 CASE WHEN TCOL.DATA_TYPE IN('bigint','bit','int','numeric','decimal','float','money','smallint','smallmoney','tinyint','real')
						 THEN 'ISNULL(CAST('+QUOTENAME(TCOL.COLUMN_NAME)+' AS VARCHAR(MAX)),''NULL'''+')'
						 WHEN TCOL.DATA_TYPE IN('date','datetime')
						 THEN 'ISNULL(''''''''+'+'CONVERT(VARCHAR,'+QUOTENAME(TCOL.COLUMN_NAME)+',112)'+'+'''''''',''NULL'''+')'
						 WHEN TCOL.DATA_TYPE IN('binary','image','varbinary')
						 THEN 'ISNULL(CONVERT(VARCHAR(MAX),CAST('+QUOTENAME(TCOL.COLUMN_NAME)+' AS VARBINARY(MAX)),1)+'')'',''NULL'''+')'
						 ELSE 'ISNULL(''''''''+'+QUOTENAME(TCOL.COLUMN_NAME)+'+'''''''',''NULL'''+')'
						 END 
	 FROM tempdb.INFORMATION_SCHEMA.COLUMNS AS TCOL
	 WHERE TCOL.TABLE_NAME='##QUERY_TABLE'

	 SELECT @ValuesColumn=SUBSTRING(@ValuesColumn,6,LEN(@ValuesColumn))

	 SELECT @NewQuery=( N'SELECT CASE WHEN RINS.RowNo=1 THEN ''VALUES''+RINS.ResultInsert WHEN RINS.RowNo>1 THEN '',''+RINS.ResultInsert ELSE RINS.ResultInsert END [ResultInsert] FROM ('+
						N'SELECT 0[RowNo],''INSERT INTO ''+'+''''+@DestinyTable+'(''+'+''+@ColumnTable+'+'')''[ResultInsert]'+
						N' UNION '+
					    N'SELECT ROW_NUMBER()OVER(ORDER BY'+SUBSTRING(@ColumnTable,2,CHARINDEX('''+''',@ColumnTable)-2)+')[RowNo]
						,''(''+'+ISNULL(@ValuesColumn,'NULL')+'+'')''[ResultInsert]'+'FROM ##QUERY_TABLE'+
						N') AS RINS'+
						N' ORDER BY RINS.RowNo,RINS.ResultInsert'
					  )
					
	 EXEC sp_executesql @NewQuery
	 PRINT @NewQuery

	 --IF @ExportPath IS NOT NULL
	 --BEGIN
		--EXEC xp_CmdShell @NewQuery
		--https://www.sqlteam.com/articles/exporting-data-programatically-with-bcp-and-xp_cmdshell
		--https://www.sqlservercentral.com/Forums/846362/BCP-using-xpcmdshell
	 --END

END TRY
BEGIN CATCH
	PRINT CAST(ERROR_NUMBER()AS VARCHAR(10))+'--'+ERROR_MESSAGE()
	PRINT @InQuery
	PRINT @NewQuery
END CATCH