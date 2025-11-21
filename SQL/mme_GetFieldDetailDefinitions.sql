IF OBJECT_ID('dbo.mme_GetFieldDetailDefinitions', 'P') IS NOT NULL
    DROP PROCEDURE dbo.mme_GetFieldDetailDefinitions;
GO

CREATE PROCEDURE dbo.mme_GetFieldDetailDefinitions
(
    @WFCON_Guid uniqueidentifier
)
AS
BEGIN
    SET NOCOUNT ON;

    IF OBJECT_ID('tempdb..#src') IS NOT NULL
        DROP TABLE #src;

    SELECT
        DCN_Prompt     AS ColumnName,
        FDD_Name       AS DatabaseName,
        FDD_Guid       AS Guid,
        ObjectName     AS DataType,
        EnglishName    AS ColumnType
    INTO #src
    FROM WFDetailConfigs
    JOIN WFConfigurations 
        ON WFCON_ID   = DCN_WFCONID
       AND WFCON_Guid = @WFCON_Guid
    JOIN WFFieldDetailDefinitions ON DCN_FDDID = FDD_ID
    JOIN DicFieldDetailTypes      ON DCN_FieldDetailTypeID = TypeID;

    IF NOT EXISTS (SELECT 1 FROM #src)
    BEGIN
        SELECT CAST(NULL AS nvarchar(50)) AS [ ]
        WHERE 1 = 0;
        RETURN;
    END;

    DECLARE @cols nvarchar(max),
            @sql  nvarchar(max);

    SELECT @cols = STRING_AGG(QUOTENAME(ColumnName), ',')
    FROM (SELECT DISTINCT ColumnName FROM #src) AS c;

    SET @sql = N'
    SELECT CASE RowType
               WHEN ''ColumnType'' THEN ''ID''
               ELSE ''''
           END AS [ ],' + @cols + '
    FROM (
        SELECT ColumnName, ''DatabaseName'' AS RowType, CAST(DatabaseName AS nvarchar(400)) AS Value FROM #src
        UNION ALL SELECT ColumnName, ''Guid'',       CAST(Guid       AS nvarchar(400)) FROM #src
        UNION ALL SELECT ColumnName, ''ColumnType'', CAST(ColumnType AS nvarchar(400)) FROM #src
    ) d
    PIVOT (
        MAX(Value)
        FOR ColumnName IN (' + @cols + ')
    ) p
    ORDER BY CASE RowType
                 WHEN ''DatabaseName'' THEN 1
                 WHEN ''Guid''         THEN 2
                 WHEN ''ColumnType''   THEN 3
             END;';

    EXEC sp_executesql @sql;
END;
GO