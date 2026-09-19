/*
    DBA_SQL_DriveSpaceReport
    Creates ONE enabled SQL Server Agent job and ONE private-use daily schedule.
    Run this entire file in SSMS on the Windows SQL instance hosting the job.
    SQL Server 2016+ is required for the current-offset safety check.

    PREREQUISITES
    - The SQL Agent host uses Windows "Pacific Standard Time", with automatic
      daylight-saving adjustment enabled. Verify with Get-TimeZone on that host.
    - SQL Server Agent is running.
    - The existing, working report script and servers.txt are on this host.
    - Run this deployment as an authorized sysadmin. It grants NO permissions.
    - The job's execution account has the required file, remote CIM/WMI,
      SQL login, and Database Mail profile permissions.

    BEHAVIOR
    - 07:00 server-local time every day, including weekends.
    - Interprets "7 AM PST" as 7 AM Pacific local time (PST/PDT), NOT fixed UTC-8.
    - Full report every morning; -OnlySendOnAlert is intentionally omitted.
    - Keeps the existing script, its mail settings, and its volume exclusions.
    - Stops if the job or schedule already exists. Does not overwrite/drop jobs.
    - No automatic retry, no immediate start, no service restart, no xp_cmdshell.
    - Job creation does NOT verify report execution or mailbox delivery.
    - The offset guard checks NOW only; it does not prove the time-zone ID or
      future DST behavior. The Windows time-zone prerequisite is mandatory.

    References:
    https://learn.microsoft.com/en-us/sql/relational-databases/system-stored-procedures/sp-add-job-transact-sql
    https://learn.microsoft.com/en-us/sql/relational-databases/system-stored-procedures/sp-add-jobstep-transact-sql
    https://learn.microsoft.com/en-us/sql/relational-databases/system-stored-procedures/sp-add-jobschedule-transact-sql
    https://learn.microsoft.com/en-us/sql/t-sql/queries/at-time-zone-transact-sql
*/
USE msdb;
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;

-- Optional: use an EXISTING, approved CmdExec proxy instead of the Agent account.
DECLARE @ProxyName sysname = NULL; -- Example: N'DBA_CmdExec_Proxy'
-- Job ownership is separate from the Windows identity that executes CmdExec.
-- Replace with an approved long-lived owner login when required by policy.
DECLARE @JobOwner sysname = SUSER_SNAME();

DECLARE @JobName sysname = N'DBA_SQL_DriveSpaceReport';
DECLARE @ScheduleName sysname = N'DBA_SQL_DriveSpaceReport_Daily_0700_Pacific';
DECLARE @JobId uniqueidentifier;
DECLARE @ReturnCode int;
DECLARE @StartDate int = CONVERT(int, CONVERT(char(8), GETDATE(), 112));
DECLARE @ServerNow datetimeoffset = SYSDATETIMEOFFSET();
DECLARE @PacificNow datetimeoffset;
SET @PacificNow = @ServerNow AT TIME ZONE 'Pacific Standard Time';

-- No ExecutionPolicy bypass. Use your organization's approved script policy.
DECLARE @Command nvarchar(max) =
    N'"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"'
  + N' -NoProfile -NonInteractive'
  + N' -File "D:\DriveSpaceReport\Send-SQLDriveCapacityReport.ps1"'
  + N' -ServerListPath "D:\DriveSpaceReport\servers.txt"'
  + N' -OutputFolder "D:\DriveSpaceReport\Reports"';

IF ISNULL(IS_SRVROLEMEMBER(N'sysadmin'), 0) <> 1
    THROW 50001, 'Run this deployment using an authorized sysadmin login. No permissions are granted by this script.', 1;

-- Never silently schedule 07:00 UTC/Eastern/etc. as if it were Pacific time.
IF DATEPART(TZOFFSET, @ServerNow) <> DATEPART(TZOFFSET, @PacificNow)
    THROW 50002, 'Server current UTC offset differs from Pacific time. Stop: this 07:00 local-time schedule is not suitable for this server. Do not change a production server time zone merely for this job.', 1;

IF EXISTS (SELECT 1 FROM dbo.sysjobs WHERE name = @JobName)
    THROW 50003, 'DBA_SQL_DriveSpaceReport already exists. Existing job was not modified. Review it in SQL Server Agent.', 1;

IF EXISTS (SELECT 1 FROM dbo.sysschedules WHERE name = @ScheduleName)
    THROW 50004, 'The proposed schedule name already exists. Nothing was changed; review the existing schedule before proceeding.', 1;

IF @ProxyName IS NOT NULL
   AND NOT EXISTS
   (
       SELECT 1
       FROM dbo.sysproxies
       WHERE name = @ProxyName AND enabled = 1
   )
    THROW 50005, 'The requested proxy does not exist or is disabled. Specify an approved, enabled CmdExec proxy or leave NULL.', 1;

BEGIN TRY
    BEGIN TRANSACTION;

    EXEC @ReturnCode = dbo.sp_add_job
        @job_name = @JobName,
        @enabled = 1,
        @description = N'Daily drive-capacity report at 07:00 Pacific local time. Requires a Pacific-time SQL Agent host. Runs existing PowerShell script and submits HTML through Database Mail. No cleanup or restarts.',
        @owner_login_name = @JobOwner,
        @start_step_id = 1,
        @notify_level_eventlog = 2,
        @notify_level_email = 0,
        @delete_level = 0,
        @job_id = @JobId OUTPUT;
    IF @ReturnCode <> 0
        THROW 50010, 'Could not create the SQL Agent job.', 1;

    EXEC @ReturnCode = dbo.sp_add_jobstep
        @job_id = @JobId,
        @step_id = 1,
        @step_name = N'Collect drive usage and submit Database Mail report',
        @subsystem = N'CmdExec',
        @command = @Command,
        @proxy_name = @ProxyName,
        @cmdexec_success_code = 0,
        @on_success_action = 1,
        @on_fail_action = 2,
        @retry_attempts = 0,
        @retry_interval = 0,
        @output_file_name = N'D:\DriveSpaceReport\DBA_SQL_DriveSpaceReport_Agent.log',
        @flags = 0;
    IF @ReturnCode <> 0
        THROW 50011, 'Could not create the CmdExec job step. Check proxy/subsystem permissions when a proxy is specified.', 1;

    EXEC @ReturnCode = dbo.sp_add_jobschedule
        @job_id = @JobId,
        @name = @ScheduleName,
        @enabled = 1,
        @freq_type = 4,            -- Daily
        @freq_interval = 1,        -- Every day, including weekends
        @freq_subday_type = 1,     -- Once, at the specified time
        @freq_subday_interval = 0,
        @active_start_date = @StartDate,
        @active_end_date = 99991231,
        @active_start_time = 070000,
        @active_end_time = 235959;
    IF @ReturnCode <> 0
        THROW 50012, 'Could not create and attach the daily 07:00 schedule.', 1;

    EXEC @ReturnCode = dbo.sp_add_jobserver
        @job_id = @JobId,
        @server_name = N'(LOCAL)';
    IF @ReturnCode <> 0
        THROW 50013, 'Could not assign the job to this SQL Server Agent instance.', 1;

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;
    THROW;
END CATCH;

-- Verification: reads the actual stored job and schedule definition.
SELECT
    j.name AS JobName,
    j.enabled AS JobEnabled,
    SUSER_SNAME(j.owner_sid) AS JobOwner,
    st.subsystem AS StepType,
    COALESCE(p.name, N'(SQL Server Agent service account)') AS RunAs,
    s.name AS ScheduleName,
    s.enabled AS ScheduleEnabled,
    s.freq_type AS FrequencyType_4_MeansDaily,
    s.freq_interval AS EveryNDays,
    s.active_start_date AS StartDate_YYYYMMDD,
    STUFF(STUFF(RIGHT('000000' + CONVERT(varchar(6), s.active_start_time), 6), 3, 0, ':'), 6, 0, ':') AS RunTime_ServerLocal,
    st.command AS StepCommand,
    st.output_file_name AS AgentOutputFile
FROM dbo.sysjobs AS j
JOIN dbo.sysjobsteps AS st ON st.job_id = j.job_id AND st.step_id = 1
JOIN dbo.sysjobschedules AS js ON js.job_id = j.job_id
JOIN dbo.sysschedules AS s ON s.schedule_id = js.schedule_id
LEFT JOIN dbo.sysproxies AS p ON p.proxy_id = st.proxy_id
WHERE j.job_id = @JobId;

PRINT 'Job created and enabled. Review the displayed definition and perform ONE manual test under SQL Agent.';
PRINT 'The creation script does not start the job immediately. The schedule is active from today.';
PRINT 'Daily time is 07:00 SERVER LOCAL TIME. Pacific zone with automatic DST adjustment is required.';
GO
