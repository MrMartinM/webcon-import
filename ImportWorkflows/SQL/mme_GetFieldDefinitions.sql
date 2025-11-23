IF OBJECT_ID('dbo.mme_GetFieldDefinitions', 'P') IS NOT NULL
    DROP PROCEDURE dbo.mme_GetFieldDefinitions;
GO

CREATE PROCEDURE dbo.mme_GetFieldDefinitions
(
    @WF_Guid uniqueidentifier
)
AS
BEGIN
    SET NOCOUNT ON;

    IF OBJECT_ID('tempdb..#src') IS NOT NULL
        DROP TABLE #src;

    SELECT
        WFCON_Prompt AS ColumnName,
        FDEF_Name    AS DatabaseName,
        WFCON_Guid   AS Guid,
        ObjectName   AS DataType,
        EnglishName  AS ColumnType
    INTO #src
    FROM WorkFlows
    JOIN WFConfigurations   ON WF_WFDEFID   = WFCON_DEFID
    JOIN WFFieldDefinitions ON WFCON_FDEFID = FDEF_ID
    JOIN DicWFFieldTypes    ON TypeID       = FDEF_WFFieldTypeID
    WHERE WF_Guid = @WF_Guid
      AND FDEF_Name NOT LIKE 'SEL_%'
      AND FDEF_Name <> 'WFD_SubElems'

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