USE DBAdmin;
GO


SET ANSI_NULLS ON;
GO

SET QUOTED_IDENTIFIER ON;
GO



CREATE OR ALTER PROCEDURE dbo.SQLTableToHTMLTable
    @TableName VARCHAR(MAX) = NULL,
    @IsTempTable BIT = NULL,
    @ColumnsToExclude VARCHAR(MAX) = NULL, --Comma seperated list of column names not to include in HTML.
    @ColumnWidth INT = 0, --0 = never wrap, otherwise number of pixels.
    @Font VARCHAR(MAX) = 'font face="times new roman,times,serif" size="2"',
    @TableBorder INT = 1,
    @CellPadding INT = 1,
    @TrimTextTo BIGINT = 2000, --Trim text for each cell to this length if it is greater.
    @HTML VARCHAR(MAX) OUTPUT,
    @WhereClause VARCHAR(MAX) = NULL, --Exclude data from table using this, will be ANDed to actual dynamic sql
    @MaxRows INT = 1000, --Maximum number of rows to put in the table
    @Debug INT = 0 --Debug mode
AS
/****************************************************************************************************************************
*	 DBAdmin.[dbo].[SQLTableToHTMLTable]
*	Creator:        Aaron Hayes
*	Date:           12/7/2011
*	Project:        NA
* 	
*	Description:   Creates an HTML string containing a table from a temp table.  Primary use is for notification emails.
*	Notes:         
*                   
*	Return Values:   
*	Usage:	         
                


*	Modifications:   
*   Developer Name     	Date     	Brief Description
*   ------------------  ---------- 	--------------------------------------------------------------------------------------------
*	
********************************************************************************************************************************/

DECLARE @NL VARCHAR(MAX);
SET @NL = CHAR(13) + CHAR(10);
DECLARE @Columns TABLE (Offset INT IDENTITY(1, 1),
                        Name VARCHAR(MAX));
DECLARE @CurCol INT,
        @MaxCol INT;
DECLARE @CurRow INT,
        @MaxRow INT;
DECLARE @iTableName VARCHAR(MAX);
SET @iTableName = '##iTable_SQLTABLETOHTMLTABLE_' + CAST(@@SPID AS VARCHAR(10));
DECLARE @TableWidth VARCHAR(MAX);
DECLARE @NumberOfColumns INT;
DECLARE @DBandSchema VARCHAR(MAX);
DECLARE @DB VARCHAR(MAX);

IF @IsTempTable IS NULL
BEGIN
    IF @TableName LIKE '#%'
        SET @IsTempTable = 1;
    ELSE
        SET @IsTempTable = 0;
END;

IF @IsTempTable = 1
BEGIN
    SET @DBandSchema = 'tempdb.dbo';
    SET @DB = 'tempdb';
END;
ELSE
BEGIN
    SET @DBandSchema = (SELECT LEFT(@TableName, PATINDEX('%.%', @TableName)
                                                + PATINDEX(
                                                      '%.%',
                                                      RIGHT(@TableName, LEN(@TableName) - PATINDEX('%.%', @TableName)))
                                                - 1));
    SET @DB = LEFT(@DBandSchema, PATINDEX('%.%', @DBandSchema) - 1);
    IF @DBandSchema IS NULL
    OR @DBandSchema NOT LIKE '%_%.%_%'
    OR @DB IS NULL
        RAISERROR(
            '[ DBAdmin.dbo.SQLTableToHTMLTable] Failed to identify database and schema names.  Check table name, used a fully qualified name for non-temp tables.',
            12,
            1);
    IF @Debug > 0
    BEGIN
        PRINT @DB;
        PRINT @DBandSchema;
    END;
END;

IF  @IsTempTable = 1
AND @TableName LIKE '#%'
    SET @TableName = 'tempdb.dbo.' + @TableName;

BEGIN TRY
    INSERT INTO @Columns
    EXEC ('select Name from ' + @DB + '.sys.columns where object_id = object_id(''' + @TableName + ''')');

    SET @MaxCol = (SELECT MAX(Offset)FROM @Columns);

    IF @MaxCol IS NULL
    OR @MaxCol < 1
        RAISERROR(
            '[SQLTableToHTMLTable] No columns found.  Check table name, used a fully qualified name for non-temp tables.',
            12,
            1);

    SET @NumberOfColumns = (SELECT COUNT(*)FROM @Columns);

    IF @ColumnsToExclude IS NOT NULL
    BEGIN
        DELETE @Columns
         WHERE Name IN (   SELECT T.Value
                             FROM (   SELECT ROW_NUMBER() OVER (ORDER BY id ASC) AS Id,
                                             Value
                                        FROM (   SELECT Num AS id,
                                                        LTRIM(
                                                            RTRIM(
                                                                SUBSTRING(
                                                                    @ColumnsToExclude,
                                                                    Num,
                                                                    CHARINDEX(',', @ColumnsToExclude + ',', Num) - Num))) AS Value
                                                   FROM dbo.Numbers
                                                  WHERE Num                                        <= DATALENGTH(@ColumnsToExclude) + 1
                                                    AND SUBSTRING(',' + @ColumnsToExclude, Num, 1) = ',') R ) AS T );
    END;

    IF @ColumnWidth = 0
        SET @TableWidth = '';
    ELSE
        SET @TableWidth = 'style = "width: ' + CAST(@ColumnWidth * @NumberOfColumns AS VARCHAR(MAX)) + ';"';

    SET @HTML
        = '<TABLE ' + @TableWidth + ' BORDER="' + CAST(@TableBorder AS VARCHAR(10)) + '" CELLPADDING="'
          + CAST(@CellPadding AS VARCHAR(10)) + '" CELLSPACING = "1">' + @NL;

    DECLARE @CreateIntermediateTable VARCHAR(MAX);
    SET @CreateIntermediateTable = 'create table ' + @iTableName + ' (RKey int identity(1,1) primary key clustered ';
    DECLARE @PopulateIntermediateTable VARCHAR(MAX);
    SET @PopulateIntermediateTable
        = 'insert into ' + @iTableName + ' ' + @NL + 'select top ' + CAST(@MaxRows AS VARCHAR(20)) + ' ';

    SET @CurCol = (SELECT MIN(Offset)FROM @Columns);

    WHILE @CurCol <= @MaxCol
    BEGIN
        SET @CreateIntermediateTable
            = @CreateIntermediateTable + ', Col' + CAST(@CurCol AS VARCHAR(10)) + ' varchar(max)';

        IF @CurCol <> (SELECT MIN(Offset)FROM @Columns)
            SET @PopulateIntermediateTable = @PopulateIntermediateTable + ', ';
        SELECT @PopulateIntermediateTable
            = @PopulateIntermediateTable + Name + ' as Col' + CAST(@CurCol AS VARCHAR(10))
          FROM @Columns
         WHERE Offset = @CurCol;

        SET @CurCol = (SELECT MIN(Offset)FROM @Columns WHERE Offset > @CurCol);
    END;

    SET @CreateIntermediateTable = @CreateIntermediateTable + ' )';
    SET @PopulateIntermediateTable = @PopulateIntermediateTable + ' from ' + @TableName + '';

    IF @WhereClause IS NOT NULL
        SET @PopulateIntermediateTable = @PopulateIntermediateTable + ' where ' + @WhereClause;

    IF @Debug > 0
        PRINT @CreateIntermediateTable;

    EXEC (@CreateIntermediateTable);

    IF @Debug > 0
        PRINT @PopulateIntermediateTable;

    EXEC (@PopulateIntermediateTable);
    SET @MaxRow = @@ROWCOUNT;

    IF @Debug > 0
    BEGIN
        PRINT 'MaxRow: ';
        PRINT @MaxRow;
    END;

    DECLARE @Content VARCHAR(MAX);
    CREATE TABLE #CellText (Value VARCHAR(MAX));
    DECLARE @GetCellText VARCHAR(MAX);

    DECLARE @ShadedRow INT;
    SET @ShadedRow = 0;

    SET @CurRow = 0;


    WHILE @CurRow <= @MaxRow
    BEGIN
        --Row start
        SET @HTML = @HTML + '  <tr>' + @NL;
        SET @CurCol = (SELECT MIN(Offset)FROM @Columns);

        IF @ShadedRow = 0
       AND @CurRow <> 0
            SET @ShadedRow = 1;
        ELSE
            SET @ShadedRow = 0;

        WHILE @CurCol <= @MaxCol
        BEGIN
            IF @CurRow = 0 --Header Row
            BEGIN
                SET @HTML = @HTML + '<td style="';

                IF @ColumnWidth = 0
                    SET @HTML = @HTML + 'white-space:nowrap;';
                ELSE
                    SET @HTML = @HTML + 'width: ' + CAST(@ColumnWidth AS VARCHAR(10)) + 'px;';

                SET @HTML = @HTML + 'background-color: rgb(210, 210, 210);';

                SET @HTML = @HTML + '">';

                SET @HTML = @HTML + '<' + @Font + '>';

                IF @ColumnWidth = 0
                    SET @HTML = @HTML + '<NOBR>';

                SET @HTML = @HTML + (SELECT Name FROM @Columns WHERE Offset = @CurCol);

                IF @ColumnWidth = 0
                    SET @HTML = @HTML + '</NOBR>';

                SET @HTML = @HTML + '</font>';

                SET @HTML = @HTML + '</TD>';
            END;
            ELSE
            BEGIN
                SET @HTML = @HTML + '		<td style="';

                IF @ColumnWidth = 0
                    SET @HTML = @HTML + 'white-space:nowrap;';
                ELSE
                    SET @HTML = @HTML + 'width: ' + CAST(@ColumnWidth AS VARCHAR(10)) + 'px;';

                IF @ShadedRow = 1
                    SET @HTML = @HTML + 'background-color: rgb(255, 245, 240);';

                SET @HTML = @HTML + '">';

                SET @HTML = @HTML + '<' + @Font + '>';

                IF @ColumnWidth = 0
                    SET @HTML = @HTML + '<NOBR>';

                DELETE #CellText;
                SET @GetCellText
                    = 'select SUBSTRING(cast(Col' + CAST(@CurCol AS VARCHAR(10)) + ' as varchar(max)),1,'
                      + CAST(@TrimTextTo AS VARCHAR(MAX)) + ') from ' + @iTableName + ' where RKey = '
                      + CAST(@CurRow AS VARCHAR(10));
                IF @Debug > 1
                    PRINT @GetCellText;
                INSERT INTO #CellText
                EXEC (@GetCellText);

                SET @HTML = @HTML + (SELECT ISNULL(Value, 'NULL')FROM #CellText);

                IF @ColumnWidth = 0
                    SET @HTML = @HTML + '</NOBR>';

                SET @HTML = @HTML + '</font>';

                SET @HTML = @HTML + '</TD>' + @NL;
            END;

            SET @CurCol = (SELECT MIN(Offset)FROM @Columns WHERE Offset > @CurCol);
        END;
        SET @CurRow = @CurRow + 1;

        --Row end
        SET @HTML = @HTML + '  </tr>' + @NL;
    END;

    SET @HTML = @HTML + '</TABLE>' + @NL;

    EXEC ('drop table ' + @iTableName);
END TRY
BEGIN CATCH
    --Make sure the global temp table is dropped
    EXEC ('if exists (select * from tempdb.sys.tables where name like ''%' + @iTableName + '%'') drop table ' + @iTableName);
    DECLARE @ErrorMsg VARCHAR(MAX);
    SET @ErrorMsg
        = DB_NAME(DB_ID()) + '.' + OBJECT_SCHEMA_NAME(@@procid, DB_ID()) + '.' + OBJECT_NAME(@@procid, DB_ID()) + ' - '
          + ERROR_MESSAGE();
    RAISERROR(@ErrorMsg, 11, 1);
    RETURN;
END CATCH;



RETURN;


GO


