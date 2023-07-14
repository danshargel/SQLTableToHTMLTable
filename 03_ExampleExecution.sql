USE DBAdmin;
GO

IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'temp')
    EXEC ('create schema temp');

DROP TABLE IF EXISTS temp.tbTestMail;
SELECT TOP 50 *
  INTO temp.tbTestMail
  FROM master.sys.objects
 ORDER BY name;


SET NOCOUNT ON;

DECLARE @HTMLMail VARCHAR(MAX),
        @Message  VARCHAR(MAX);

EXEC DBAdmin.dbo.SQLTableToHTMLTable @TableName = 'DBAdmin.temp.tbTestMail',
                                     @ColumnsToExclude = 'RowID',
                                     @HTML = @HTMLMail OUTPUT;

SELECT @Message = '<h3>Hey look at this, dummy.</h3><br><br>' + @HTMLMail;

EXEC msdb.dbo.sp_send_dbmail @profile_name = DEFAULT,
                             @body_format = 'HTML',
                             @recipients = 'xyz@xyz.com',
                             @subject = 'test email',
                             @body = @Message;

DROP TABLE IF EXISTS temp.tbTestMail;