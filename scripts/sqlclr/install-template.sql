/*
    DBA-ONLY, PERSISTENT SERVER/DATABASE CHANGE.

    Install into a dedicated utility database named OneCAdmin, never into the
    1C application database. Review and execute manually as a sysadmin/DBA.

    SQL Server 2014 requires the server option "clr enabled". This script does
    not enable it automatically.
*/
SET NOCOUNT ON;

IF DB_NAME() <> N'OneCAdmin'
    THROW 50101, 'Connect to the dedicated OneCAdmin utility database.', 1;

IF CAST(SERVERPROPERTY('ProductMajorVersion') AS int) >= 14
    PRINT 'SQL Server 2017+ may also require a signed assembly or a trusted assembly entry.';

IF EXISTS (SELECT 1 FROM sys.assemblies WHERE name = N'OneCSqlMetadata')
    THROW 50102, 'Assembly OneCSqlMetadata already exists; review upgrade manually.', 1;

CREATE ASSEMBLY [OneCSqlMetadata]
AUTHORIZATION [dbo]
FROM __ASSEMBLY_HEX__
WITH PERMISSION_SET = SAFE;
GO
CREATE FUNCTION [dbo].[OneC_InflateUtf8](@input varbinary(max))
RETURNS nvarchar(max)
AS EXTERNAL NAME [OneCSqlMetadata].[OneC.SqlClr.OneCMetadata].[InflateUtf8];
GO

CREATE FUNCTION [dbo].[OneC_DecodeUtf8](@input varbinary(max))
RETURNS nvarchar(max)
AS EXTERNAL NAME [OneCSqlMetadata].[OneC.SqlClr.OneCMetadata].[DecodeUtf8];
GO

CREATE FUNCTION [dbo].[OneC_ParseDbNames](@input varbinary(max))
RETURNS TABLE
(
    MetadataId uniqueidentifier,
    Token nvarchar(64),
    Code int
)
AS EXTERNAL NAME [OneCSqlMetadata].[OneC.SqlClr.OneCMetadata].[ParseDbNames];
GO

CREATE FUNCTION [dbo].[OneC_ParseConfigNames](@input varbinary(max))
RETURNS TABLE
(
    MetadataId uniqueidentifier,
    LogicalName nvarchar(256)
)
AS EXTERNAL NAME [OneCSqlMetadata].[OneC.SqlClr.OneCMetadata].[ParseConfigNames];
GO

CREATE FUNCTION [dbo].[OneC_ParseFixedReferenceEdges]
(
    @schema varbinary(max),
    @targetTable nvarchar(128)
)
RETURNS TABLE
(
    SourceTable nvarchar(128),
    SourceField nvarchar(128),
    TargetTable nvarchar(128)
)
AS EXTERNAL NAME [OneCSqlMetadata].[OneC.SqlClr.OneCMetadata].[ParseFixedReferenceEdges];
GO
