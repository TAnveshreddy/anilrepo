/* =============================================================================
   Logic Test 3 - staging tables
   -----------------------------------------------------------------------------
   One table per source sheet of PQSpecshop_Validation_V4.xlsx.
   RowId must preserve the original sheet row order: the logic uses
   "first row in data" for tie-breaks and for the order of the results.
   Target: SQL Server 2017+ (TRIM, STRING_AGG, CREATE OR ALTER).
   ============================================================================= */

-- Sheet "Refinery DB" (only the columns the logic reads are required)
IF OBJECT_ID('dbo.RefineryDB') IS NULL
CREATE TABLE dbo.RefineryDB (
    RowId                       INT            NOT NULL PRIMARY KEY, -- sheet row number
    SpecificationSheet          NVARCHAR(400)  NULL,  -- col A
    Source                      NVARCHAR(255)  NULL,  -- col B (refinery)
    Product                     NVARCHAR(255)  NULL,  -- col F
    HasSeason                   INT            NULL,  -- col I (1 = seasonal)
    Season                      NVARCHAR(100)  NULL,  -- col J
    Tests                       NVARCHAR(400)  NULL,  -- col M (quality name)
    SeasonStart                 NVARCHAR(20)   NULL,  -- col O, text 'm/d'
    SeasonEnd                   NVARCHAR(20)   NULL,  -- col P, text 'm/d'
    MinimumNumericDeterminant   FLOAT          NULL,  -- col R (NULL when not numeric)
    MaximumNumericDeterminant   FLOAT          NULL   -- col S (NULL when not numeric)
);

-- Sheet "Terminal DB" (Source = the destination terminal)
IF OBJECT_ID('dbo.TerminalDB') IS NULL
CREATE TABLE dbo.TerminalDB (
    RowId                       INT            NOT NULL PRIMARY KEY,
    SpecificationSheet          NVARCHAR(400)  NULL,  -- col A
    Source                      NVARCHAR(255)  NULL,  -- col B (terminal = destination)
    Product                     NVARCHAR(255)  NULL,  -- col F
    HasSeason                   INT            NULL,  -- col I
    Season                      NVARCHAR(100)  NULL,  -- col J
    Tests                       NVARCHAR(400)  NULL,  -- col L
    SeasonStart                 NVARCHAR(20)   NULL,  -- col N
    SeasonEnd                   NVARCHAR(20)   NULL,  -- col O
    MinimumNumericDeterminant   FLOAT          NULL,  -- col Q
    MaximumNumericDeterminant   FLOAT          NULL   -- col R
);

-- Sheet "Terminals Specsheet Master Data" (rows 7+, cols C:U, wide layout kept)
-- Source / Product may be blank on continuation rows; the procedure fills them down.
IF OBJECT_ID('dbo.TerminalsSpecsheetMaster') IS NULL
CREATE TABLE dbo.TerminalsSpecsheetMaster (
    RowId               INT           NOT NULL PRIMARY KEY,
    Source              NVARCHAR(255) NULL,   -- col C
    Product             NVARCHAR(255) NULL,   -- col D
    SpecificationSheet  NVARCHAR(400) NULL,   -- col E
    Dest01 NVARCHAR(255) NULL, Dest02 NVARCHAR(255) NULL, Dest03 NVARCHAR(255) NULL, Dest04 NVARCHAR(255) NULL,
    Dest05 NVARCHAR(255) NULL, Dest06 NVARCHAR(255) NULL, Dest07 NVARCHAR(255) NULL, Dest08 NVARCHAR(255) NULL,
    Dest09 NVARCHAR(255) NULL, Dest10 NVARCHAR(255) NULL, Dest11 NVARCHAR(255) NULL, Dest12 NVARCHAR(255) NULL,
    Dest13 NVARCHAR(255) NULL, Dest14 NVARCHAR(255) NULL, Dest15 NVARCHAR(255) NULL, Dest16 NVARCHAR(255) NULL  -- cols F:U
);

-- Sheet "Property Mapping" (rows 7+, cols A:E)
-- CommonPropertyName is only filled on the first row of each group; the procedure carries it down.
IF OBJECT_ID('dbo.PropertyMapping') IS NULL
CREATE TABLE dbo.PropertyMapping (
    RowId                    INT           NOT NULL PRIMARY KEY,
    IQDMPropertyTag          NVARCHAR(50)  NULL,  -- col A
    IQDMPropertyDescription  NVARCHAR(255) NULL,  -- col B
    CommonPropertyName       NVARCHAR(255) NULL,  -- col C
    QualityName              NVARCHAR(400) NULL,  -- col D (Refinery DB col M / Terminal DB col L)
    SourceDB                 NVARCHAR(50)  NULL   -- col E ('Refinery DB' / 'Terminal DB')
);

-- Sheet "Product Mapping" cols J:K (override list)
IF OBJECT_ID('dbo.ProductMappingOverride') IS NULL
CREATE TABLE dbo.ProductMappingOverride (
    SpecSheetProduct   NVARCHAR(255) NOT NULL PRIMARY KEY,  -- col J, e.g. 'No. 2 MV 15 / Heating Oil'
    TerminalDBProduct  NVARCHAR(255) NOT NULL               -- col K, e.g. 'No. 2 MV 15'
);

-- Recommended indexes for the lookups the procedure does
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_RefineryDB_Spec_Source_Product')
    CREATE INDEX IX_RefineryDB_Spec_Source_Product ON dbo.RefineryDB (Source, Product) INCLUDE (SpecificationSheet);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_TerminalDB_Source_Product')
    CREATE INDEX IX_TerminalDB_Source_Product ON dbo.TerminalDB (Source, Product);
