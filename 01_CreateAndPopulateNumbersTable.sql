Use DBAdmin
GO

--DROP TABLE IF EXISTS dbo.Number;
--GO


DECLARE @UpperBound INT = 1000000;

;WITH cteN (Number)
   AS (SELECT      ROW_NUMBER() OVER (ORDER BY s1.object_id) - 1
         FROM      sys.all_columns AS s1
        CROSS JOIN sys.all_columns AS s2)
SELECT Number
  INTO dbo.Numbers
  FROM cteN
 WHERE Number <= @UpperBound;

CREATE UNIQUE CLUSTERED INDEX CIX_Number
ON dbo.Numbers (Number)
WITH (FILLFACTOR = 100, -- in the event server default has been changed
      DATA_COMPRESSION = ROW -- if Enterprise & table large enough to matter
);
