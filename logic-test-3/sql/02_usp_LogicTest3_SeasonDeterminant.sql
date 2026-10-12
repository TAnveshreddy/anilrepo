/* =============================================================================
   dbo.usp_LogicTest3_SeasonDeterminant
   -----------------------------------------------------------------------------
   SQL implementation of the "Logic Test 3" sheet (Season / Determinant Logic
   Test - Common Property version) of PQSpecshop_Validation_V4.xlsx, including
   all helper columns of "Logic Calc 3".

   Inputs  : @Source, @Product, @Destination, @ETA   (Logic Test 3 B4:B7)
   Returns : result set 1 - header  (Spec sheet, ETA key, Data check, Terminal DB check)
             result set 2 - results (one row per IQDM Property Tag / Description)

   Step numbers in the comments refer to the "Logic BPMN 3" step register.
   Text comparisons rely on a case-insensitive collation, like Excel.
   Requires SQL Server 2017+.
   ============================================================================= */
CREATE OR ALTER PROCEDURE dbo.usp_LogicTest3_SeasonDeterminant
    @Source       NVARCHAR(255),
    @Product      NVARCHAR(255),
    @Destination  NVARCHAR(255),
    @ETA          DATE
AS
BEGIN
    SET NOCOUNT ON;

    /* ---------------------------------------------------------------------
       Steps 1-2: inputs and ETA key (month*100 + day, year ignored)
       --------------------------------------------------------------------- */
    SET @Source      = TRIM(ISNULL(@Source, N''));
    SET @Product     = TRIM(ISNULL(@Product, N''));
    SET @Destination = TRIM(ISNULL(@Destination, N''));

    IF @ETA IS NULL OR @ETA <= '2000-01-01'
        THROW 50001, N'ETA must be a date after 1-Jan-2000.', 1;

    DECLARE @ETAKey INT = MONTH(@ETA) * 100 + DAY(@ETA);

    /* ---------------------------------------------------------------------
       Steps 3-6: SpecificationSheet from Terminals Specsheet Master Data
       (fill Source/Product down, match Source + Product + Destination,
       first matching row wins)
       --------------------------------------------------------------------- */
    DECLARE @Spec NVARCHAR(400) = N'',
            @DestMatchCount INT = 0;

    IF @Source <> N'' AND @Product <> N'' AND @Destination <> N''
    BEGIN
        ;WITH Master AS (
            SELECT  m.RowId,
                    Src  = ISNULL(src.Val, N''),
                    Prd  = ISNULL(prd.Val, N''),
                    Spec = TRIM(ISNULL(m.SpecificationSheet, N''))
            FROM    dbo.TerminalsSpecsheetMaster m
            OUTER APPLY (SELECT TOP (1) Val = TRIM(m2.Source)
                         FROM dbo.TerminalsSpecsheetMaster m2
                         WHERE m2.RowId <= m.RowId AND TRIM(ISNULL(m2.Source, N'')) <> N''
                         ORDER BY m2.RowId DESC) src
            OUTER APPLY (SELECT TOP (1) Val = TRIM(m2.Product)
                         FROM dbo.TerminalsSpecsheetMaster m2
                         WHERE m2.RowId <= m.RowId AND TRIM(ISNULL(m2.Product, N'')) <> N''
                         ORDER BY m2.RowId DESC) prd
        )
        SELECT  ms.RowId, ms.Spec
        INTO    #DestMatch
        FROM    Master ms
        JOIN    dbo.TerminalsSpecsheetMaster m ON m.RowId = ms.RowId
        WHERE   ms.Spec <> N''
          AND   ms.Src = @Source
          AND   ms.Prd = @Product
          AND   EXISTS (SELECT 1
                        FROM (VALUES (m.Dest01),(m.Dest02),(m.Dest03),(m.Dest04),
                                     (m.Dest05),(m.Dest06),(m.Dest07),(m.Dest08),
                                     (m.Dest09),(m.Dest10),(m.Dest11),(m.Dest12),
                                     (m.Dest13),(m.Dest14),(m.Dest15),(m.Dest16)) d(Dest)
                        WHERE TRIM(d.Dest) = @Destination);

        SET @DestMatchCount = (SELECT COUNT(*) FROM #DestMatch);
        SET @Spec = ISNULL((SELECT TOP (1) Spec FROM #DestMatch ORDER BY RowId),
                           N'Not found in master data');
    END;

    /* ---------------------------------------------------------------------
       Steps 8 / 20 helper: quality -> Common Property Name
       Property Mapping col C is carried down over blank cells; a row with a
       blank Quality Name breaks the carry (same as Logic Calc 3 col AM).
       First mapping row for "SourceDB|Quality" wins.
       --------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#QualityMap') IS NOT NULL DROP TABLE #QualityMap;

    ;WITH PM AS (
        SELECT  p.RowId,
                Quality = TRIM(ISNULL(p.QualityName, N'')),
                SrcDb   = ISNULL(p.SourceDB, N''),
                CommonProperty =
                    CASE WHEN TRIM(ISNULL(p.QualityName, N'')) = N'' THEN N''
                         ELSE ISNULL(( SELECT TOP (1)
                                              CASE WHEN TRIM(ISNULL(k.QualityName, N'')) = N'' THEN N''
                                                   ELSE TRIM(k.CommonPropertyName) END
                                       FROM   dbo.PropertyMapping k
                                       WHERE  k.RowId <= p.RowId
                                         AND (TRIM(ISNULL(k.CommonPropertyName, N'')) <> N''
                                              OR TRIM(ISNULL(k.QualityName, N'')) = N'')
                                       ORDER BY k.RowId DESC), N'')
                    END
        FROM    dbo.PropertyMapping p
    ),
    Ranked AS (
        SELECT  LookupKey = SrcDb + N'|' + Quality,
                CommonProperty,
                rn = ROW_NUMBER() OVER (PARTITION BY SrcDb + N'|' + Quality ORDER BY RowId)
        FROM    PM
        WHERE   Quality <> N''
    )
    SELECT  LookupKey, CommonProperty
    INTO    #QualityMap
    FROM    Ranked
    WHERE   rn = 1;

    /* ---------------------------------------------------------------------
       Steps 7-17: Refinery DB rows for Spec + Source + Product
       (Logic Calc 3 cols A:M; only InputMatch = 1 rows are kept)
       --------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#Ref') IS NOT NULL DROP TABLE #Ref;

    SELECT  r.RowId,
            Test       = CASE WHEN qm.LookupKey IS NOT NULL THEN qm.CommonProperty
                              ELSE TRIM(ISNULL(r.Tests, N'')) END,
            HasSeason  = CASE WHEN r.HasSeason = 1 THEN 1 ELSE 0 END,
            Season     = CASE WHEN r.HasSeason = 1 AND UPPER(TRIM(ISNULL(r.Season, N''))) <> N'NULL'
                              THEN TRIM(ISNULL(r.Season, N'')) ELSE N'' END,
            k.StartKey,
            k.EndKey,
            MinV       = r.MinimumNumericDeterminant,
            MaxV       = r.MaximumNumericDeterminant,
            SeasonStatus = CAST(NULL AS NVARCHAR(20)),
            Applicable   = CAST(0 AS INT)
    INTO    #Ref
    FROM    dbo.RefineryDB r
    LEFT JOIN #QualityMap qm
           ON qm.LookupKey = N'Refinery DB|' + TRIM(ISNULL(r.Tests, N''))
    CROSS APPLY (SELECT ss = TRIM(ISNULL(r.SeasonStart, N'')), se = TRIM(ISNULL(r.SeasonEnd, N''))) t
    CROSS APPLY (SELECT
            StartKey = CASE WHEN r.HasSeason = 1 AND CHARINDEX(N'/', t.ss) > 1
                            THEN TRY_CONVERT(INT, LEFT(t.ss, CHARINDEX(N'/', t.ss) - 1)) * 100
                               + TRY_CONVERT(INT, NULLIF(SUBSTRING(t.ss, CHARINDEX(N'/', t.ss) + 1, 2), N'')) END,
            EndKey   = CASE WHEN r.HasSeason = 1 AND CHARINDEX(N'/', t.se) > 1
                            THEN TRY_CONVERT(INT, LEFT(t.se, CHARINDEX(N'/', t.se) - 1)) * 100
                               + TRY_CONVERT(INT, NULLIF(SUBSTRING(t.se, CHARINDEX(N'/', t.se) + 1, 2), N'')) END
    ) k
    WHERE   TRIM(ISNULL(r.SpecificationSheet, N'')) = @Spec
      AND   TRIM(ISNULL(r.Source, N''))             = @Source
      AND   TRIM(ISNULL(r.Product, N''))            = @Product;

    -- Steps 10-17: season status and applicability (window may wrap the year end)
    UPDATE #Ref
    SET    SeasonStatus =
               CASE WHEN HasSeason <> 1 THEN N'No season'
                    WHEN StartKey IS NULL OR EndKey IS NULL THEN N'Dates missing'
                    WHEN (StartKey <= EndKey AND @ETAKey BETWEEN StartKey AND EndKey)
                      OR (StartKey >  EndKey AND (@ETAKey >= StartKey OR @ETAKey <= EndKey))
                         THEN N'In season'
                    ELSE N'Out of season' END;

    UPDATE #Ref
    SET    Applicable = CASE WHEN SeasonStatus IN (N'No season', N'In season') THEN 1 ELSE 0 END;

    /* ---------------------------------------------------------------------
       Steps 18-29: Terminal DB rows for Destination + resolved product
       (Logic Calc 3 cols Z:AH, AK)
       --------------------------------------------------------------------- */
    DECLARE @TdbProduct NVARCHAR(255) =
        ISNULL((SELECT TOP (1) TRIM(o.TerminalDBProduct)
                FROM dbo.ProductMappingOverride o
                WHERE TRIM(o.SpecSheetProduct) = @Product), @Product);

    IF OBJECT_ID('tempdb..#Tdb') IS NOT NULL DROP TABLE #Tdb;

    SELECT  r.RowId,
            Test       = CASE WHEN qm.LookupKey IS NOT NULL THEN qm.CommonProperty
                              ELSE TRIM(ISNULL(r.Tests, N'')) END,
            HasSeason  = CASE WHEN r.HasSeason = 1 THEN 1 ELSE 0 END,
            Season     = CASE WHEN r.HasSeason = 1 AND UPPER(TRIM(ISNULL(r.Season, N''))) <> N'NULL'
                              THEN TRIM(ISNULL(r.Season, N'')) ELSE N'' END,
            k.StartKey,
            k.EndKey,
            MinV       = r.MinimumNumericDeterminant,
            MaxV       = r.MaximumNumericDeterminant,
            SeasonStatus = CAST(NULL AS NVARCHAR(20)),
            Applicable   = CAST(0 AS INT)
    INTO    #Tdb
    FROM    dbo.TerminalDB r
    LEFT JOIN #QualityMap qm
           ON qm.LookupKey = N'Terminal DB|' + TRIM(ISNULL(r.Tests, N''))
    CROSS APPLY (SELECT ss = TRIM(ISNULL(r.SeasonStart, N'')), se = TRIM(ISNULL(r.SeasonEnd, N''))) t
    CROSS APPLY (SELECT
            StartKey = CASE WHEN r.HasSeason = 1 AND CHARINDEX(N'/', t.ss) > 1
                            THEN TRY_CONVERT(INT, LEFT(t.ss, CHARINDEX(N'/', t.ss) - 1)) * 100
                               + TRY_CONVERT(INT, NULLIF(SUBSTRING(t.ss, CHARINDEX(N'/', t.ss) + 1, 2), N'')) END,
            EndKey   = CASE WHEN r.HasSeason = 1 AND CHARINDEX(N'/', t.se) > 1
                            THEN TRY_CONVERT(INT, LEFT(t.se, CHARINDEX(N'/', t.se) - 1)) * 100
                               + TRY_CONVERT(INT, NULLIF(SUBSTRING(t.se, CHARINDEX(N'/', t.se) + 1, 2), N'')) END
    ) k
    WHERE   TRIM(ISNULL(r.Source, N''))  = @Destination
      AND   TRIM(ISNULL(r.Product, N'')) = @TdbProduct;

    UPDATE #Tdb
    SET    SeasonStatus =
               CASE WHEN HasSeason <> 1 THEN N'No season'
                    WHEN StartKey IS NULL OR EndKey IS NULL THEN N'Dates missing'
                    WHEN (StartKey <= EndKey AND @ETAKey BETWEEN StartKey AND EndKey)
                      OR (StartKey >  EndKey AND (@ETAKey >= StartKey OR @ETAKey <= EndKey))
                         THEN N'In season'
                    ELSE N'Out of season' END;

    UPDATE #Tdb
    SET    Applicable = CASE WHEN SeasonStatus IN (N'No season', N'In season') THEN 1 ELSE 0 END;

    /* ---------------------------------------------------------------------
       Steps 30-32: property list
       Refinery properties (first-occurrence order), then Terminal-only ones;
       stable sort: non-seasonal first, then seasonal.
       --------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#Props') IS NOT NULL DROP TABLE #Props;

    ;WITH RefList AS (
        SELECT Test, FirstRow = MIN(RowId), Seasonal = MAX(HasSeason)
        FROM   #Ref
        WHERE  Test <> N''
        GROUP BY Test
    ),
    TdbOnly AS (
        SELECT Test, FirstRow = MIN(RowId), Seasonal = MAX(HasSeason)
        FROM   #Tdb t
        WHERE  Test <> N''
          AND  NOT EXISTS (SELECT 1 FROM RefList rl WHERE rl.Test = t.Test)
        GROUP BY Test
    ),
    AllProps AS (
        SELECT Test, Seasonal, SrcOrder = 1, FirstRow FROM RefList
        UNION ALL
        SELECT Test, Seasonal, SrcOrder = 2, FirstRow FROM TdbOnly
    )
    SELECT  PropSeq = ROW_NUMBER() OVER (ORDER BY Seasonal, SrcOrder, FirstRow),
            CommonProperty = Test,
            Seasonal,
            InRefinery = CASE WHEN SrcOrder = 1 THEN 1 ELSE 0 END
    INTO    #Props
    FROM    AllProps;

    /* ---------------------------------------------------------------------
       Steps 33-35: expand each property to its IQDM Tag / Description pairs
       (Property Mapping col C carried down; properties with no pair dropped)
       --------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#Pairs') IS NOT NULL DROP TABLE #Pairs;

    SELECT  p.RowId,
            Tag  = TRIM(ISNULL(p.IQDMPropertyTag, N'')),
            Descr = TRIM(ISNULL(p.IQDMPropertyDescription, N'')),
            CommonProperty = ISNULL(cp.Val, N'')
    INTO    #Pairs
    FROM    dbo.PropertyMapping p
    OUTER APPLY (SELECT TOP (1) Val = TRIM(k.CommonPropertyName)
                 FROM dbo.PropertyMapping k
                 WHERE k.RowId <= p.RowId AND TRIM(ISNULL(k.CommonPropertyName, N'')) <> N''
                 ORDER BY k.RowId DESC) cp
    WHERE   TRIM(ISNULL(p.IQDMPropertyTag, N'')) <> N''
       OR   TRIM(ISNULL(p.IQDMPropertyDescription, N'')) <> N'';

    /* ---------------------------------------------------------------------
       Steps 36-37: limits and flags per property
       Min = MAX of Mins, Max = MIN of Maxes over applicable rows
       --------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#Calc') IS NOT NULL DROP TABLE #Calc;

    SELECT  pr.PropSeq,
            pr.CommonProperty,
            pr.InRefinery,
            RefMatchCnt        = ref.MatchCnt,
            RefApplicCnt       = ref.ApplicCnt,
            RefApplicSeasonCnt = ref.ApplicSeasonCnt,
            RefDatesMissing    = ref.DatesMissingCnt,
            RefMin = ref.MinV,
            RefMax = ref.MaxV,
            TdbMin = tdb.MinV,
            TdbMax = tdb.MaxV,
            TdbApplicCnt        = tdb.ApplicCnt,
            TdbApplicSeasonName = tdb.ApplicSeasonNameCnt,
            TdbDatesMissing     = tdb.DatesMissingCnt,
            MinSeason = CASE WHEN ref.MinV IS NULL THEN N''
                             ELSE ISNULL(NULLIF((SELECT TOP (1) r.Season FROM #Ref r
                                                 WHERE r.Test = pr.CommonProperty AND r.Applicable = 1
                                                   AND r.MinV = ref.MinV ORDER BY r.RowId), N''), N'N/A') END,
            MaxSeason = CASE WHEN ref.MaxV IS NULL THEN N''
                             ELSE ISNULL(NULLIF((SELECT TOP (1) r.Season FROM #Ref r
                                                 WHERE r.Test = pr.CommonProperty AND r.Applicable = 1
                                                   AND r.MaxV = ref.MaxV ORDER BY r.RowId), N''), N'N/A') END,
            RefSeasonList = ISNULL((SELECT STRING_AGG(s.Season, N', ') WITHIN GROUP (ORDER BY s.FirstRow)
                                    FROM (SELECT r.Season, FirstRow = MIN(r.RowId) FROM #Ref r
                                          WHERE r.Test = pr.CommonProperty AND r.Applicable = 1
                                            AND r.HasSeason = 1 AND r.Season <> N''
                                          GROUP BY r.Season) s), N''),
            TdbSeasonList = ISNULL((SELECT STRING_AGG(s.Season, N', ') WITHIN GROUP (ORDER BY s.FirstRow)
                                    FROM (SELECT t.Season, FirstRow = MIN(t.RowId) FROM #Tdb t
                                          WHERE t.Test = pr.CommonProperty AND t.Applicable = 1
                                            AND t.Season <> N''
                                          GROUP BY t.Season) s), N'')
    INTO    #Calc
    FROM    #Props pr
    OUTER APPLY (SELECT MatchCnt        = COUNT(*),
                        ApplicCnt       = SUM(r.Applicable),
                        ApplicSeasonCnt = SUM(CASE WHEN r.Applicable = 1 AND r.HasSeason = 1 THEN 1 ELSE 0 END),
                        DatesMissingCnt = SUM(CASE WHEN r.SeasonStatus = N'Dates missing' THEN 1 ELSE 0 END),
                        MinV = MAX(CASE WHEN r.Applicable = 1 THEN r.MinV END),
                        MaxV = MIN(CASE WHEN r.Applicable = 1 THEN r.MaxV END)
                 FROM   #Ref r WHERE r.Test = pr.CommonProperty) ref
    OUTER APPLY (SELECT ApplicCnt           = ISNULL(SUM(t.Applicable), 0),
                        ApplicSeasonNameCnt = ISNULL(SUM(CASE WHEN t.Applicable = 1 AND t.Season <> N'' THEN 1 ELSE 0 END), 0),
                        DatesMissingCnt     = ISNULL(SUM(CASE WHEN t.SeasonStatus = N'Dates missing' THEN 1 ELSE 0 END), 0),
                        MinV = MAX(CASE WHEN t.Applicable = 1 THEN t.MinV END),
                        MaxV = MIN(CASE WHEN t.Applicable = 1 THEN t.MaxV END)
                 FROM   #Tdb t WHERE t.Test = pr.CommonProperty) tdb;

    /* ---------------------------------------------------------------------
       Result set 1: header (Logic Test 3 B8:B11)
       --------------------------------------------------------------------- */
    DECLARE @RefMatched INT = (SELECT COUNT(*) FROM #Ref),
            @RefApplic  INT = (SELECT COUNT(*) FROM #Ref WHERE Applicable = 1),
            @TdbMatched INT = (SELECT COUNT(*) FROM #Tdb),
            @TdbApplic  INT = (SELECT COUNT(*) FROM #Tdb WHERE Applicable = 1);

    SELECT  SpecificationSheet = @Spec,
            ETAKey             = @ETAKey,
            TerminalDBProduct  = @TdbProduct,
            DataCheck = CASE
                WHEN @DestMatchCount = 0
                    THEN N'No specification sheet in Terminals Specsheet Master Data for this Source / Product / Destination - re-pick Product and Destination'
                WHEN @RefMatched = 0
                    THEN N'Spec sheet ''' + @Spec + N''' found in master data but no matching rows in Refinery DB - check spec name'
                ELSE CONCAT(@RefMatched, N' data rows matched; ', @RefApplic, N' applicable on ETA',
                            CASE WHEN @DestMatchCount > 1
                                 THEN CONCAT(N'; ', @DestMatchCount, N' spec sheets share this destination - first one used')
                                 ELSE N'' END)
                END,
            TerminalDBCheck = CONCAT(N'Terminal DB: ', @TdbMatched, N' rows match Destination + Product; ',
                                     @TdbApplic, N' applicable on ETA');

    /* ---------------------------------------------------------------------
       Result set 2: results (Logic Test 3 A13:N...)  - steps 38-43
       --------------------------------------------------------------------- */
    SELECT  Season = CASE
                WHEN c.RefMatchCnt = 0 THEN                                   -- step 39: Terminal-only
                    CASE WHEN c.TdbApplicCnt = 0
                              THEN CASE WHEN c.TdbDatesMissing > 0 THEN N'Dates missing' ELSE N'Out of season' END
                         WHEN c.TdbApplicSeasonName = 0 THEN N'N/A'
                         ELSE c.TdbSeasonList END
                WHEN c.RefApplicCnt = 0 THEN                                   -- step 41
                    CASE WHEN c.RefDatesMissing > 0 THEN N'Dates missing' ELSE N'Out of season' END
                WHEN c.RefMax IS NOT NULL THEN c.MaxSeason                     -- step 42
                WHEN c.RefMin IS NOT NULL THEN c.MinSeason
                WHEN c.RefApplicSeasonCnt = 0 THEN N'N/A'
                ELSE c.RefSeasonList END,
            [IQDM Property Tag]                     = pa.Tag,
            [IQDM Property Description]             = pa.Descr,
            MinimumNumericDeterminant               = c.RefMin,
            MaximumNumericDeterminant               = c.RefMax,
            [Terminal DB MinimumNumericDeterminant] = c.TdbMin,
            [Terminal DB MaximumNumericDeterminant] = c.TdbMax,
            [Overlap?]      = CASE WHEN c.RefMatchCnt = 0 THEN N''
                                   WHEN c.RefApplicSeasonCnt > 1 THEN N'Yes' ELSE N'No' END,
            [Rows applied]  = CASE WHEN c.RefMatchCnt = 0 THEN NULL ELSE c.RefApplicCnt END,
            [Rule applied]  = CASE
                WHEN c.RefMatchCnt = 0
                    THEN N'Terminal DB only - property not in the Refinery DB spec list; limits taken from Terminal DB (columns E:F)'
                WHEN c.RefApplicCnt = 0
                    THEN CASE WHEN c.RefDatesMissing > 0
                              THEN N'Season dates not in data - cannot be evaluated for ETA'
                              ELSE N'No season window covers the ETA - no limits apply' END
                WHEN c.RefApplicSeasonCnt > 1
                    THEN N'Overlap: Min = MAX of Mins, Max = MIN of Maxes; Season = row giving the final value'
                       + CASE WHEN c.RefMin IS NOT NULL AND c.RefMax IS NOT NULL AND c.MinSeason <> c.MaxSeason
                              THEN N' (Min from ' + c.MinSeason + N', Max from ' + c.MaxSeason + N')' ELSE N'' END
                WHEN c.RefApplicCnt > 1
                    THEN N'Several qualities / rows pooled - Min = MAX of Mins, Max = MIN of Maxes'
                WHEN c.RefApplicSeasonCnt = 1
                    THEN N'Single season window - limits used as is'
                ELSE N'No season - limits used as is' END,
            [Season dates missing?] = CASE WHEN c.RefMatchCnt = 0 THEN N''
                                           WHEN c.RefDatesMissing > 0 THEN N'Yes' ELSE N'No' END,
            [Min taken from season] = c.MinSeason,
            [Max taken from season] = c.MaxSeason,
            [Property source]       = CASE WHEN c.RefMatchCnt > 0 THEN N'Refinery DB' ELSE N'Terminal DB only' END,
            [Common Property Name]  = c.CommonProperty
    FROM    #Calc c
    JOIN    #Pairs pa ON pa.CommonProperty = c.CommonProperty
    ORDER BY c.PropSeq, pa.RowId;
END;
GO

/* Example (inputs used in the workbook):
EXEC dbo.usp_LogicTest3_SeasonDeterminant
     @Source      = N'Canton, OH : MPC Refinery',
     @Product     = N'84 Blend Grade Conventional',
     @Destination = N'East Sparta, OH : MPL Tank Farm',
     @ETA         = '2026-09-29';
*/
