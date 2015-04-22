
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 04/15/2015 11:58:26
-- Generated from EDMX file: d:\data\visual studio 2013\Projects\McoEasyTool\McoEasyTool\DataModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [McoEasyTool_v4_DB];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[FK_AdReportFaultyServer_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[FaultyServerReports] DROP CONSTRAINT [FK_AdReportFaultyServer_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_FaultyServerFaultyServer_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[FaultyServerReports] DROP CONSTRAINT [FK_FaultyServerFaultyServer_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_PoolBackupServer]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Servers_BackupServer] DROP CONSTRAINT [FK_PoolBackupServer];
GO
IF OBJECT_ID(N'[dbo].[FK_BackupReportBackupServer_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[BackupServerReports] DROP CONSTRAINT [FK_BackupReportBackupServer_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_AppReportApplication_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[ApplicationAppReports] DROP CONSTRAINT [FK_AppReportApplication_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_AppScheduleScheduled_Application]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[ScheduledApplications] DROP CONSTRAINT [FK_AppScheduleScheduled_Application];
GO
IF OBJECT_ID(N'[dbo].[FK_ApplicationApplication_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[ApplicationAppReports] DROP CONSTRAINT [FK_ApplicationApplication_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_ApplicationAppHtmlElement]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AppHtmlElements] DROP CONSTRAINT [FK_ApplicationAppHtmlElement];
GO
IF OBJECT_ID(N'[dbo].[FK_ApplicationScheduled_Application]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[ScheduledApplications] DROP CONSTRAINT [FK_ApplicationScheduled_Application];
GO
IF OBJECT_ID(N'[dbo].[FK_Application_AppServer]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Servers_AppServer] DROP CONSTRAINT [FK_Application_AppServer];
GO
IF OBJECT_ID(N'[dbo].[FK_AppServer_AppServerReport]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AppServerAppReports] DROP CONSTRAINT [FK_AppServer_AppServerReport];
GO
IF OBJECT_ID(N'[dbo].[FK_Application_ReportAppServer_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AppServerAppReports] DROP CONSTRAINT [FK_Application_ReportAppServer_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_ReportEmail]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Reports] DROP CONSTRAINT [FK_ReportEmail];
GO
IF OBJECT_ID(N'[dbo].[FK_ScheduleReport]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Reports] DROP CONSTRAINT [FK_ScheduleReport];
GO
IF OBJECT_ID(N'[dbo].[FK_SpaceServerSpaceServer_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[SpaceServer_Reports] DROP CONSTRAINT [FK_SpaceServerSpaceServer_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_SpaceReportSpaceServer_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[SpaceServer_Reports] DROP CONSTRAINT [FK_SpaceReportSpaceServer_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_AdReport_inherits_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Reports_AdReport] DROP CONSTRAINT [FK_AdReport_inherits_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_FaultyServer_inherits_Server]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Servers_FaultyServer] DROP CONSTRAINT [FK_FaultyServer_inherits_Server];
GO
IF OBJECT_ID(N'[dbo].[FK_BackupServer_inherits_Server]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Servers_BackupServer] DROP CONSTRAINT [FK_BackupServer_inherits_Server];
GO
IF OBJECT_ID(N'[dbo].[FK_BackupReport_inherits_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Reports_BackupReport] DROP CONSTRAINT [FK_BackupReport_inherits_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_AppReport_inherits_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Reports_AppReport] DROP CONSTRAINT [FK_AppReport_inherits_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_AppSchedule_inherits_Schedule]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Schedules_AppSchedule] DROP CONSTRAINT [FK_AppSchedule_inherits_Schedule];
GO
IF OBJECT_ID(N'[dbo].[FK_AppServer_inherits_Server]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Servers_AppServer] DROP CONSTRAINT [FK_AppServer_inherits_Server];
GO
IF OBJECT_ID(N'[dbo].[FK_SpaceServer_inherits_Server]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Servers_SpaceServer] DROP CONSTRAINT [FK_SpaceServer_inherits_Server];
GO
IF OBJECT_ID(N'[dbo].[FK_SpaceReport_inherits_Report]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Reports_SpaceReport] DROP CONSTRAINT [FK_SpaceReport_inherits_Report];
GO
IF OBJECT_ID(N'[dbo].[FK_AdSchedule_inherits_Schedule]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Schedules_AdSchedule] DROP CONSTRAINT [FK_AdSchedule_inherits_Schedule];
GO
IF OBJECT_ID(N'[dbo].[FK_BackupSchedule_inherits_Schedule]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Schedules_BackupSchedule] DROP CONSTRAINT [FK_BackupSchedule_inherits_Schedule];
GO
IF OBJECT_ID(N'[dbo].[FK_SpaceSchedule_inherits_Schedule]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Schedules_SpaceSchedule] DROP CONSTRAINT [FK_SpaceSchedule_inherits_Schedule];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[FaultyServerReports]', 'U') IS NOT NULL
    DROP TABLE [dbo].[FaultyServerReports];
GO
IF OBJECT_ID(N'[dbo].[Recipients]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Recipients];
GO
IF OBJECT_ID(N'[dbo].[ReftechServers]', 'U') IS NOT NULL
    DROP TABLE [dbo].[ReftechServers];
GO
IF OBJECT_ID(N'[dbo].[Pools]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Pools];
GO
IF OBJECT_ID(N'[dbo].[BackupServerReports]', 'U') IS NOT NULL
    DROP TABLE [dbo].[BackupServerReports];
GO
IF OBJECT_ID(N'[dbo].[AD_Settings]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AD_Settings];
GO
IF OBJECT_ID(N'[dbo].[Applications]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Applications];
GO
IF OBJECT_ID(N'[dbo].[ApplicationAppReports]', 'U') IS NOT NULL
    DROP TABLE [dbo].[ApplicationAppReports];
GO
IF OBJECT_ID(N'[dbo].[AppServerAppReports]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AppServerAppReports];
GO
IF OBJECT_ID(N'[dbo].[AppHtmlElements]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AppHtmlElements];
GO
IF OBJECT_ID(N'[dbo].[ScheduledApplications]', 'U') IS NOT NULL
    DROP TABLE [dbo].[ScheduledApplications];
GO
IF OBJECT_ID(N'[dbo].[Emails]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Emails];
GO
IF OBJECT_ID(N'[dbo].[Reports]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Reports];
GO
IF OBJECT_ID(N'[dbo].[Schedules]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Schedules];
GO
IF OBJECT_ID(N'[dbo].[Servers]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Servers];
GO
IF OBJECT_ID(N'[dbo].[AppDomains]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AppDomains];
GO
IF OBJECT_ID(N'[dbo].[SpaceServer_Reports]', 'U') IS NOT NULL
    DROP TABLE [dbo].[SpaceServer_Reports];
GO
IF OBJECT_ID(N'[dbo].[Accounts]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Accounts];
GO
IF OBJECT_ID(N'[dbo].[Reports_AdReport]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Reports_AdReport];
GO
IF OBJECT_ID(N'[dbo].[Servers_FaultyServer]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Servers_FaultyServer];
GO
IF OBJECT_ID(N'[dbo].[Servers_BackupServer]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Servers_BackupServer];
GO
IF OBJECT_ID(N'[dbo].[Reports_BackupReport]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Reports_BackupReport];
GO
IF OBJECT_ID(N'[dbo].[Reports_AppReport]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Reports_AppReport];
GO
IF OBJECT_ID(N'[dbo].[Schedules_AppSchedule]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Schedules_AppSchedule];
GO
IF OBJECT_ID(N'[dbo].[Servers_AppServer]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Servers_AppServer];
GO
IF OBJECT_ID(N'[dbo].[Servers_SpaceServer]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Servers_SpaceServer];
GO
IF OBJECT_ID(N'[dbo].[Reports_SpaceReport]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Reports_SpaceReport];
GO
IF OBJECT_ID(N'[dbo].[Schedules_AdSchedule]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Schedules_AdSchedule];
GO
IF OBJECT_ID(N'[dbo].[Schedules_BackupSchedule]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Schedules_BackupSchedule];
GO
IF OBJECT_ID(N'[dbo].[Schedules_SpaceSchedule]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Schedules_SpaceSchedule];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'FaultyServerReports'
CREATE TABLE [dbo].[FaultyServerReports] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Ping] nvarchar(max)  NULL,
    [Details] nvarchar(max)  NULL,
    [AbsenceDuration] nvarchar(max)  NULL,
    [Ticket] nvarchar(max)  NULL,
    [AdReportId] int  NOT NULL,
    [FaultyServerId] int  NOT NULL,
    [AdReport_Id] int  NOT NULL,
    [FaultyServer_Id] int  NOT NULL
);
GO

-- Creating table 'Recipients'
CREATE TABLE [dbo].[Recipients] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [RelativeAddress] nvarchar(max)  NULL,
    [AbsoluteAddress] nvarchar(max)  NULL,
    [Name] nvarchar(max)  NULL,
    [Module] nvarchar(max)  NOT NULL,
    [Included] bit  NULL
);
GO

-- Creating table 'ReftechServers'
CREATE TABLE [dbo].[ReftechServers] (
    [IdServeur] int  NOT NULL,
    [NomMachineServeur] varchar(30)  NOT NULL,
    [NomLogiqueServeur] varchar(100)  NULL,
    [ServiceMajeur] char(15)  NULL,
    [Environnement] char(15)  NOT NULL,
    [Perimetre] char(15)  NOT NULL,
    [CodeDomaine] char(3)  NOT NULL,
    [IsDedie] char(1)  NOT NULL,
    [IsHauteDispo] char(1)  NULL,
    [IsMaquette] char(1)  NOT NULL,
    [NumSerie] varchar(50)  NULL,
    [EtatServeur] char(1)  NOT NULL,
    [IdSite] varchar(130)  NULL,
    [LocalisationPhysique] varchar(30)  NULL,
    [IdUserExploitLocal] int  NULL,
    [RemarquesServeur] varchar(400)  NULL,
    [DateInsertServeur] datetime  NOT NULL,
    [IP] varchar(30)  NULL,
    [DateUpdateServeur] datetime  NULL,
    [IsPCI] char(1)  NOT NULL,
    [IdTypeSupervision] int  NULL,
    [NiveauSupervision] int  NULL,
    [IdExploitantServeur] int  NULL,
    [IdKM] varchar(20)  NULL,
    [IdServeurParent] int  NULL,
    [IdTypeServeur] int  NULL
);
GO

-- Creating table 'Pools'
CREATE TABLE [dbo].[Pools] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [BackupDay] int  NOT NULL,
    [CheckDay] int  NOT NULL,
    [CellColor] nvarchar(max)  NULL,
    [CheckFolder] nvarchar(max)  NULL,
    [BackupManager] nvarchar(max)  NULL,
    [ExecutionAccount] nvarchar(max)  NULL,
    [CheckAccount] nvarchar(max)  NULL
);
GO

-- Creating table 'BackupServerReports'
CREATE TABLE [dbo].[BackupServerReports] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [State] nvarchar(max)  NOT NULL,
    [Ping] nvarchar(max)  NOT NULL,
    [Services] nvarchar(max)  NOT NULL,
    [Relaunched] nvarchar(max)  NOT NULL,
    [Details] nvarchar(max)  NOT NULL,
    [BackupReportId] int  NOT NULL,
    [BackupServerId] int  NOT NULL,
    [BackupReport_Id] int  NOT NULL
);
GO

-- Creating table 'AD_Settings'
CREATE TABLE [dbo].[AD_Settings] (
    [DurationFilter] bit  NOT NULL,
    [StateFilter] bit  NOT NULL,
    [PingFilter] bit  NOT NULL,
    [ErrorFilter] bit  NOT NULL,
    [Duration] int  NOT NULL,
    [State] nvarchar(max)  NULL,
    [Errors] nvarchar(max)  NULL,
    [Id] int  NOT NULL,
    [MemoryErrors] nvarchar(max)  NULL
);
GO

-- Creating table 'Applications'
CREATE TABLE [dbo].[Applications] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Url] nvarchar(max)  NULL,
    [Domain] int  NULL,
    [Navigator] nvarchar(max)  NOT NULL
);
GO

-- Creating table 'ApplicationAppReports'
CREATE TABLE [dbo].[ApplicationAppReports] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [State] nvarchar(max)  NOT NULL,
    [Details] nvarchar(max)  NULL,
    [Linkable] nvarchar(max)  NULL,
    [Authentified] nvarchar(max)  NULL,
    [AppReportId] int  NOT NULL,
    [ApplicationId] int  NOT NULL,
    [AppReport_Id] int  NOT NULL
);
GO

-- Creating table 'AppServerAppReports'
CREATE TABLE [dbo].[AppServerAppReports] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Ping] nvarchar(max)  NOT NULL,
    [State] nvarchar(max)  NOT NULL,
    [Details] nvarchar(max)  NOT NULL,
    [Application_ReportId] int  NOT NULL,
    [AppServerId] int  NOT NULL,
    [AppServer_Id] int  NOT NULL
);
GO

-- Creating table 'AppHtmlElements'
CREATE TABLE [dbo].[AppHtmlElements] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [TagName] nvarchar(max)  NULL,
    [AttrId] nvarchar(max)  NULL,
    [AttrName] nvarchar(max)  NULL,
    [AttrClass] nvarchar(max)  NULL,
    [Value] nvarchar(max)  NULL,
    [Type] nvarchar(max)  NOT NULL,
    [ApplicationId] int  NOT NULL,
    [AttrXpath] nvarchar(max)  NULL
);
GO

-- Creating table 'ScheduledApplications'
CREATE TABLE [dbo].[ScheduledApplications] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [AppScheduleId] int  NOT NULL,
    [ApplicationId] int  NOT NULL,
    [AppSchedule_Id] int  NOT NULL
);
GO

-- Creating table 'Emails'
CREATE TABLE [dbo].[Emails] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Recipients] nvarchar(max)  NULL,
    [Sender] nvarchar(max)  NULL,
    [Subject] nvarchar(max)  NULL,
    [Body] nvarchar(max)  NULL,
    [Module] nvarchar(max)  NOT NULL,
    [Sent] bit  NULL
);
GO

-- Creating table 'Reports'
CREATE TABLE [dbo].[Reports] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [DateTime] datetime  NOT NULL,
    [Duration] time  NULL,
    [TotalChecked] int  NULL,
    [TotalErrors] int  NULL,
    [ResultPath] nvarchar(max)  NULL,
    [Module] nvarchar(max)  NOT NULL,
    [ScheduleId] int  NULL,
    [Author] nvarchar(max)  NULL,
    [Email_Id] int  NOT NULL
);
GO

-- Creating table 'Schedules'
CREATE TABLE [dbo].[Schedules] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [CreationTime] datetime  NOT NULL,
    [Multiplicity] nvarchar(max)  NOT NULL,
    [TaskName] nvarchar(max)  NULL,
    [Generator] nvarchar(max)  NULL,
    [Executed] int  NULL,
    [NextExecution] datetime  NULL,
    [State] nvarchar(max)  NOT NULL,
    [Module] nvarchar(max)  NOT NULL
);
GO

-- Creating table 'Servers'
CREATE TABLE [dbo].[Servers] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [IpAddress] nvarchar(max)  NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Status] nvarchar(max)  NULL,
    [Location] nvarchar(max)  NULL,
    [Version] nvarchar(max)  NULL,
    [ActiveDirecotryDomain] nvarchar(max)  NULL
);
GO

-- Creating table 'AppDomains'
CREATE TABLE [dbo].[AppDomains] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Applications] int  NULL
);
GO

-- Creating table 'SpaceServer_Reports'
CREATE TABLE [dbo].[SpaceServer_Reports] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [State] nvarchar(max)  NOT NULL,
    [Ping] nvarchar(max)  NULL,
    [Details] nvarchar(max)  NULL,
    [SpaceReportId] int  NOT NULL,
    [SpaceServerId] int  NOT NULL,
    [SpaceServer_Id] int  NOT NULL,
    [SpaceReport_Id] int  NOT NULL
);
GO

-- Creating table 'Accounts'
CREATE TABLE [dbo].[Accounts] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Domain] nvarchar(max)  NOT NULL,
    [Username] nvarchar(max)  NOT NULL,
    [Password] nvarchar(max)  NOT NULL,
    [DisplayName] nvarchar(max)  NOT NULL,
    [IsSystem] bit  NOT NULL
);
GO

-- Creating table 'Reports_AdReport'
CREATE TABLE [dbo].[Reports_AdReport] (
    [FatalErrors] int  NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Servers_FaultyServer'
CREATE TABLE [dbo].[Servers_FaultyServer] (
    [IdSite] nvarchar(max)  NULL,
    [Site] nvarchar(max)  NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Servers_BackupServer'
CREATE TABLE [dbo].[Servers_BackupServer] (
    [Disks] nvarchar(max)  NULL,
    [PoolId] int  NOT NULL,
    [Included] bit  NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Reports_BackupReport'
CREATE TABLE [dbo].[Reports_BackupReport] (
    [WeekNumber] int  NOT NULL,
    [LastUpdate] datetime  NOT NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Reports_AppReport'
CREATE TABLE [dbo].[Reports_AppReport] (
    [Id] int  NOT NULL
);
GO

-- Creating table 'Schedules_AppSchedule'
CREATE TABLE [dbo].[Schedules_AppSchedule] (
    [AutoRelaunch] bit  NOT NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Servers_AppServer'
CREATE TABLE [dbo].[Servers_AppServer] (
    [StartOrder] nvarchar(max)  NOT NULL,
    [StopOrder] nvarchar(max)  NULL,
    [ApplicationId] int  NOT NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Servers_SpaceServer'
CREATE TABLE [dbo].[Servers_SpaceServer] (
    [Included] bit  NULL,
    [Disks] nvarchar(max)  NOT NULL,
    [IsShare] bit  NOT NULL,
    [CellColor] nvarchar(max)  NULL,
    [ExecutionAccount] nvarchar(max)  NULL,
    [CheckAccount] nvarchar(max)  NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Reports_SpaceReport'
CREATE TABLE [dbo].[Reports_SpaceReport] (
    [Id] int  NOT NULL
);
GO

-- Creating table 'Schedules_AdSchedule'
CREATE TABLE [dbo].[Schedules_AdSchedule] (
    [AutoCorrect] bit  NOT NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Schedules_BackupSchedule'
CREATE TABLE [dbo].[Schedules_BackupSchedule] (
    [AutoRelaunch] bit  NOT NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Schedules_SpaceSchedule'
CREATE TABLE [dbo].[Schedules_SpaceSchedule] (
    [Id] int  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'FaultyServerReports'
ALTER TABLE [dbo].[FaultyServerReports]
ADD CONSTRAINT [PK_FaultyServerReports]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Recipients'
ALTER TABLE [dbo].[Recipients]
ADD CONSTRAINT [PK_Recipients]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [IdServeur], [NomMachineServeur], [Environnement], [Perimetre], [CodeDomaine], [IsDedie], [IsMaquette], [EtatServeur], [DateInsertServeur], [IsPCI] in table 'ReftechServers'
ALTER TABLE [dbo].[ReftechServers]
ADD CONSTRAINT [PK_ReftechServers]
    PRIMARY KEY CLUSTERED ([IdServeur], [NomMachineServeur], [Environnement], [Perimetre], [CodeDomaine], [IsDedie], [IsMaquette], [EtatServeur], [DateInsertServeur], [IsPCI] ASC);
GO

-- Creating primary key on [Id] in table 'Pools'
ALTER TABLE [dbo].[Pools]
ADD CONSTRAINT [PK_Pools]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'BackupServerReports'
ALTER TABLE [dbo].[BackupServerReports]
ADD CONSTRAINT [PK_BackupServerReports]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'AD_Settings'
ALTER TABLE [dbo].[AD_Settings]
ADD CONSTRAINT [PK_AD_Settings]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Applications'
ALTER TABLE [dbo].[Applications]
ADD CONSTRAINT [PK_Applications]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'ApplicationAppReports'
ALTER TABLE [dbo].[ApplicationAppReports]
ADD CONSTRAINT [PK_ApplicationAppReports]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'AppServerAppReports'
ALTER TABLE [dbo].[AppServerAppReports]
ADD CONSTRAINT [PK_AppServerAppReports]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'AppHtmlElements'
ALTER TABLE [dbo].[AppHtmlElements]
ADD CONSTRAINT [PK_AppHtmlElements]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'ScheduledApplications'
ALTER TABLE [dbo].[ScheduledApplications]
ADD CONSTRAINT [PK_ScheduledApplications]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Emails'
ALTER TABLE [dbo].[Emails]
ADD CONSTRAINT [PK_Emails]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Reports'
ALTER TABLE [dbo].[Reports]
ADD CONSTRAINT [PK_Reports]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Schedules'
ALTER TABLE [dbo].[Schedules]
ADD CONSTRAINT [PK_Schedules]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Servers'
ALTER TABLE [dbo].[Servers]
ADD CONSTRAINT [PK_Servers]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'AppDomains'
ALTER TABLE [dbo].[AppDomains]
ADD CONSTRAINT [PK_AppDomains]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'SpaceServer_Reports'
ALTER TABLE [dbo].[SpaceServer_Reports]
ADD CONSTRAINT [PK_SpaceServer_Reports]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Accounts'
ALTER TABLE [dbo].[Accounts]
ADD CONSTRAINT [PK_Accounts]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Reports_AdReport'
ALTER TABLE [dbo].[Reports_AdReport]
ADD CONSTRAINT [PK_Reports_AdReport]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Servers_FaultyServer'
ALTER TABLE [dbo].[Servers_FaultyServer]
ADD CONSTRAINT [PK_Servers_FaultyServer]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Servers_BackupServer'
ALTER TABLE [dbo].[Servers_BackupServer]
ADD CONSTRAINT [PK_Servers_BackupServer]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Reports_BackupReport'
ALTER TABLE [dbo].[Reports_BackupReport]
ADD CONSTRAINT [PK_Reports_BackupReport]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Reports_AppReport'
ALTER TABLE [dbo].[Reports_AppReport]
ADD CONSTRAINT [PK_Reports_AppReport]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Schedules_AppSchedule'
ALTER TABLE [dbo].[Schedules_AppSchedule]
ADD CONSTRAINT [PK_Schedules_AppSchedule]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Servers_AppServer'
ALTER TABLE [dbo].[Servers_AppServer]
ADD CONSTRAINT [PK_Servers_AppServer]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Servers_SpaceServer'
ALTER TABLE [dbo].[Servers_SpaceServer]
ADD CONSTRAINT [PK_Servers_SpaceServer]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Reports_SpaceReport'
ALTER TABLE [dbo].[Reports_SpaceReport]
ADD CONSTRAINT [PK_Reports_SpaceReport]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Schedules_AdSchedule'
ALTER TABLE [dbo].[Schedules_AdSchedule]
ADD CONSTRAINT [PK_Schedules_AdSchedule]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Schedules_BackupSchedule'
ALTER TABLE [dbo].[Schedules_BackupSchedule]
ADD CONSTRAINT [PK_Schedules_BackupSchedule]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Schedules_SpaceSchedule'
ALTER TABLE [dbo].[Schedules_SpaceSchedule]
ADD CONSTRAINT [PK_Schedules_SpaceSchedule]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [AdReport_Id] in table 'FaultyServerReports'
ALTER TABLE [dbo].[FaultyServerReports]
ADD CONSTRAINT [FK_AdReportFaultyServer_Report]
    FOREIGN KEY ([AdReport_Id])
    REFERENCES [dbo].[Reports_AdReport]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_AdReportFaultyServer_Report'
CREATE INDEX [IX_FK_AdReportFaultyServer_Report]
ON [dbo].[FaultyServerReports]
    ([AdReport_Id]);
GO

-- Creating foreign key on [FaultyServer_Id] in table 'FaultyServerReports'
ALTER TABLE [dbo].[FaultyServerReports]
ADD CONSTRAINT [FK_FaultyServerFaultyServer_Report]
    FOREIGN KEY ([FaultyServer_Id])
    REFERENCES [dbo].[Servers_FaultyServer]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_FaultyServerFaultyServer_Report'
CREATE INDEX [IX_FK_FaultyServerFaultyServer_Report]
ON [dbo].[FaultyServerReports]
    ([FaultyServer_Id]);
GO

-- Creating foreign key on [PoolId] in table 'Servers_BackupServer'
ALTER TABLE [dbo].[Servers_BackupServer]
ADD CONSTRAINT [FK_PoolBackupServer]
    FOREIGN KEY ([PoolId])
    REFERENCES [dbo].[Pools]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_PoolBackupServer'
CREATE INDEX [IX_FK_PoolBackupServer]
ON [dbo].[Servers_BackupServer]
    ([PoolId]);
GO

-- Creating foreign key on [BackupReport_Id] in table 'BackupServerReports'
ALTER TABLE [dbo].[BackupServerReports]
ADD CONSTRAINT [FK_BackupReportBackupServer_Report]
    FOREIGN KEY ([BackupReport_Id])
    REFERENCES [dbo].[Reports_BackupReport]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_BackupReportBackupServer_Report'
CREATE INDEX [IX_FK_BackupReportBackupServer_Report]
ON [dbo].[BackupServerReports]
    ([BackupReport_Id]);
GO

-- Creating foreign key on [AppReport_Id] in table 'ApplicationAppReports'
ALTER TABLE [dbo].[ApplicationAppReports]
ADD CONSTRAINT [FK_AppReportApplication_Report]
    FOREIGN KEY ([AppReport_Id])
    REFERENCES [dbo].[Reports_AppReport]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_AppReportApplication_Report'
CREATE INDEX [IX_FK_AppReportApplication_Report]
ON [dbo].[ApplicationAppReports]
    ([AppReport_Id]);
GO

-- Creating foreign key on [AppSchedule_Id] in table 'ScheduledApplications'
ALTER TABLE [dbo].[ScheduledApplications]
ADD CONSTRAINT [FK_AppScheduleScheduled_Application]
    FOREIGN KEY ([AppSchedule_Id])
    REFERENCES [dbo].[Schedules_AppSchedule]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_AppScheduleScheduled_Application'
CREATE INDEX [IX_FK_AppScheduleScheduled_Application]
ON [dbo].[ScheduledApplications]
    ([AppSchedule_Id]);
GO

-- Creating foreign key on [ApplicationId] in table 'ApplicationAppReports'
ALTER TABLE [dbo].[ApplicationAppReports]
ADD CONSTRAINT [FK_ApplicationApplication_Report]
    FOREIGN KEY ([ApplicationId])
    REFERENCES [dbo].[Applications]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ApplicationApplication_Report'
CREATE INDEX [IX_FK_ApplicationApplication_Report]
ON [dbo].[ApplicationAppReports]
    ([ApplicationId]);
GO

-- Creating foreign key on [ApplicationId] in table 'AppHtmlElements'
ALTER TABLE [dbo].[AppHtmlElements]
ADD CONSTRAINT [FK_ApplicationAppHtmlElement]
    FOREIGN KEY ([ApplicationId])
    REFERENCES [dbo].[Applications]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ApplicationAppHtmlElement'
CREATE INDEX [IX_FK_ApplicationAppHtmlElement]
ON [dbo].[AppHtmlElements]
    ([ApplicationId]);
GO

-- Creating foreign key on [ApplicationId] in table 'ScheduledApplications'
ALTER TABLE [dbo].[ScheduledApplications]
ADD CONSTRAINT [FK_ApplicationScheduled_Application]
    FOREIGN KEY ([ApplicationId])
    REFERENCES [dbo].[Applications]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ApplicationScheduled_Application'
CREATE INDEX [IX_FK_ApplicationScheduled_Application]
ON [dbo].[ScheduledApplications]
    ([ApplicationId]);
GO

-- Creating foreign key on [ApplicationId] in table 'Servers_AppServer'
ALTER TABLE [dbo].[Servers_AppServer]
ADD CONSTRAINT [FK_Application_AppServer]
    FOREIGN KEY ([ApplicationId])
    REFERENCES [dbo].[Applications]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_Application_AppServer'
CREATE INDEX [IX_FK_Application_AppServer]
ON [dbo].[Servers_AppServer]
    ([ApplicationId]);
GO

-- Creating foreign key on [AppServer_Id] in table 'AppServerAppReports'
ALTER TABLE [dbo].[AppServerAppReports]
ADD CONSTRAINT [FK_AppServer_AppServerReport]
    FOREIGN KEY ([AppServer_Id])
    REFERENCES [dbo].[Servers_AppServer]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_AppServer_AppServerReport'
CREATE INDEX [IX_FK_AppServer_AppServerReport]
ON [dbo].[AppServerAppReports]
    ([AppServer_Id]);
GO

-- Creating foreign key on [Application_ReportId] in table 'AppServerAppReports'
ALTER TABLE [dbo].[AppServerAppReports]
ADD CONSTRAINT [FK_Application_ReportAppServer_Report]
    FOREIGN KEY ([Application_ReportId])
    REFERENCES [dbo].[ApplicationAppReports]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_Application_ReportAppServer_Report'
CREATE INDEX [IX_FK_Application_ReportAppServer_Report]
ON [dbo].[AppServerAppReports]
    ([Application_ReportId]);
GO

-- Creating foreign key on [Email_Id] in table 'Reports'
ALTER TABLE [dbo].[Reports]
ADD CONSTRAINT [FK_ReportEmail]
    FOREIGN KEY ([Email_Id])
    REFERENCES [dbo].[Emails]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ReportEmail'
CREATE INDEX [IX_FK_ReportEmail]
ON [dbo].[Reports]
    ([Email_Id]);
GO

-- Creating foreign key on [ScheduleId] in table 'Reports'
ALTER TABLE [dbo].[Reports]
ADD CONSTRAINT [FK_ScheduleReport]
    FOREIGN KEY ([ScheduleId])
    REFERENCES [dbo].[Schedules]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ScheduleReport'
CREATE INDEX [IX_FK_ScheduleReport]
ON [dbo].[Reports]
    ([ScheduleId]);
GO

-- Creating foreign key on [SpaceServer_Id] in table 'SpaceServer_Reports'
ALTER TABLE [dbo].[SpaceServer_Reports]
ADD CONSTRAINT [FK_SpaceServerSpaceServer_Report]
    FOREIGN KEY ([SpaceServer_Id])
    REFERENCES [dbo].[Servers_SpaceServer]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_SpaceServerSpaceServer_Report'
CREATE INDEX [IX_FK_SpaceServerSpaceServer_Report]
ON [dbo].[SpaceServer_Reports]
    ([SpaceServer_Id]);
GO

-- Creating foreign key on [SpaceReport_Id] in table 'SpaceServer_Reports'
ALTER TABLE [dbo].[SpaceServer_Reports]
ADD CONSTRAINT [FK_SpaceReportSpaceServer_Report]
    FOREIGN KEY ([SpaceReport_Id])
    REFERENCES [dbo].[Reports_SpaceReport]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_SpaceReportSpaceServer_Report'
CREATE INDEX [IX_FK_SpaceReportSpaceServer_Report]
ON [dbo].[SpaceServer_Reports]
    ([SpaceReport_Id]);
GO

-- Creating foreign key on [BackupServerId] in table 'BackupServerReports'
ALTER TABLE [dbo].[BackupServerReports]
ADD CONSTRAINT [FK_BackupServerBackupServer_Report]
    FOREIGN KEY ([BackupServerId])
    REFERENCES [dbo].[Servers_BackupServer]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_BackupServerBackupServer_Report'
CREATE INDEX [IX_FK_BackupServerBackupServer_Report]
ON [dbo].[BackupServerReports]
    ([BackupServerId]);
GO

-- Creating foreign key on [Id] in table 'Reports_AdReport'
ALTER TABLE [dbo].[Reports_AdReport]
ADD CONSTRAINT [FK_AdReport_inherits_Report]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Reports]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Servers_FaultyServer'
ALTER TABLE [dbo].[Servers_FaultyServer]
ADD CONSTRAINT [FK_FaultyServer_inherits_Server]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Servers]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Servers_BackupServer'
ALTER TABLE [dbo].[Servers_BackupServer]
ADD CONSTRAINT [FK_BackupServer_inherits_Server]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Servers]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Reports_BackupReport'
ALTER TABLE [dbo].[Reports_BackupReport]
ADD CONSTRAINT [FK_BackupReport_inherits_Report]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Reports]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Reports_AppReport'
ALTER TABLE [dbo].[Reports_AppReport]
ADD CONSTRAINT [FK_AppReport_inherits_Report]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Reports]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Schedules_AppSchedule'
ALTER TABLE [dbo].[Schedules_AppSchedule]
ADD CONSTRAINT [FK_AppSchedule_inherits_Schedule]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Schedules]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Servers_AppServer'
ALTER TABLE [dbo].[Servers_AppServer]
ADD CONSTRAINT [FK_AppServer_inherits_Server]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Servers]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Servers_SpaceServer'
ALTER TABLE [dbo].[Servers_SpaceServer]
ADD CONSTRAINT [FK_SpaceServer_inherits_Server]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Servers]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Reports_SpaceReport'
ALTER TABLE [dbo].[Reports_SpaceReport]
ADD CONSTRAINT [FK_SpaceReport_inherits_Report]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Reports]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Schedules_AdSchedule'
ALTER TABLE [dbo].[Schedules_AdSchedule]
ADD CONSTRAINT [FK_AdSchedule_inherits_Schedule]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Schedules]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Schedules_BackupSchedule'
ALTER TABLE [dbo].[Schedules_BackupSchedule]
ADD CONSTRAINT [FK_BackupSchedule_inherits_Schedule]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Schedules]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Schedules_SpaceSchedule'
ALTER TABLE [dbo].[Schedules_SpaceSchedule]
ADD CONSTRAINT [FK_SpaceSchedule_inherits_Schedule]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Schedules]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------