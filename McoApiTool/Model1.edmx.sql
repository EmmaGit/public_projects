
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 04/20/2015 18:26:15
-- Generated from EDMX file: d:\data\visual studio 2013\Projects\McoApiTool\McoApiTool\Model1.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [RoomNetDB];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------


-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------


-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'Users'
CREATE TABLE [dbo].[Users] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Username] nvarchar(max)  NOT NULL,
    [Email] nvarchar(max)  NOT NULL,
    [FirstConnection] nvarchar(max)  NOT NULL,
    [LastConnection] nvarchar(max)  NOT NULL,
    [Picture] nvarchar(max)  NOT NULL,
    [Hostname] nvarchar(max)  NOT NULL,
    [RoleId] int  NOT NULL,
    [ExchangeId] int  NOT NULL,
    [Setting_Id] int  NOT NULL
);
GO

-- Creating table 'Roles'
CREATE TABLE [dbo].[Roles] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL
);
GO

-- Creating table 'Resources'
CREATE TABLE [dbo].[Resources] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Extension] nvarchar(max)  NOT NULL,
    [Type] nvarchar(max)  NOT NULL,
    [Filename] nvarchar(max)  NOT NULL,
    [Path] nvarchar(max)  NOT NULL,
    [Size] nvarchar(max)  NOT NULL,
    [CreationTime] datetime  NOT NULL,
    [UpdateTime] datetime  NOT NULL,
    [MediaryId] int  NOT NULL
);
GO

-- Creating table 'Feeds'
CREATE TABLE [dbo].[Feeds] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Title] nvarchar(max)  NOT NULL,
    [Content] nvarchar(max)  NOT NULL,
    [CreationTime] datetime  NOT NULL,
    [UpdateTime] datetime  NOT NULL,
    [UserId] int  NOT NULL
);
GO

-- Creating table 'Likes'
CREATE TABLE [dbo].[Likes] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [UserId] int  NOT NULL,
    [FeedId] int  NOT NULL
);
GO

-- Creating table 'Settings'
CREATE TABLE [dbo].[Settings] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Color] nvarchar(max)  NOT NULL,
    [Language] nvarchar(max)  NOT NULL
);
GO

-- Creating table 'Mediaries'
CREATE TABLE [dbo].[Mediaries] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Settings] nvarchar(max)  NOT NULL,
    [CreationTime] datetime  NOT NULL,
    [UpdateTime] datetime  NOT NULL,
    [UserId] int  NOT NULL
);
GO

-- Creating table 'Exchanges'
CREATE TABLE [dbo].[Exchanges] (
    [Id] int IDENTITY(1,1) NOT NULL
);
GO

-- Creating table 'Feeds_Message'
CREATE TABLE [dbo].[Feeds_Message] (
    [ExchangeId] int  NOT NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Feeds_Post'
CREATE TABLE [dbo].[Feeds_Post] (
    [Availability] int IDENTITY(1,1) NOT NULL,
    [Id] int  NOT NULL
);
GO

-- Creating table 'Feeds_Comment'
CREATE TABLE [dbo].[Feeds_Comment] (
    [PostId] int  NOT NULL,
    [Id] int  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'Users'
ALTER TABLE [dbo].[Users]
ADD CONSTRAINT [PK_Users]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Roles'
ALTER TABLE [dbo].[Roles]
ADD CONSTRAINT [PK_Roles]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Resources'
ALTER TABLE [dbo].[Resources]
ADD CONSTRAINT [PK_Resources]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Feeds'
ALTER TABLE [dbo].[Feeds]
ADD CONSTRAINT [PK_Feeds]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Likes'
ALTER TABLE [dbo].[Likes]
ADD CONSTRAINT [PK_Likes]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Settings'
ALTER TABLE [dbo].[Settings]
ADD CONSTRAINT [PK_Settings]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Mediaries'
ALTER TABLE [dbo].[Mediaries]
ADD CONSTRAINT [PK_Mediaries]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Exchanges'
ALTER TABLE [dbo].[Exchanges]
ADD CONSTRAINT [PK_Exchanges]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Feeds_Message'
ALTER TABLE [dbo].[Feeds_Message]
ADD CONSTRAINT [PK_Feeds_Message]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Feeds_Post'
ALTER TABLE [dbo].[Feeds_Post]
ADD CONSTRAINT [PK_Feeds_Post]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Feeds_Comment'
ALTER TABLE [dbo].[Feeds_Comment]
ADD CONSTRAINT [PK_Feeds_Comment]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [RoleId] in table 'Users'
ALTER TABLE [dbo].[Users]
ADD CONSTRAINT [FK_RoleUser]
    FOREIGN KEY ([RoleId])
    REFERENCES [dbo].[Roles]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_RoleUser'
CREATE INDEX [IX_FK_RoleUser]
ON [dbo].[Users]
    ([RoleId]);
GO

-- Creating foreign key on [Setting_Id] in table 'Users'
ALTER TABLE [dbo].[Users]
ADD CONSTRAINT [FK_UserSetting]
    FOREIGN KEY ([Setting_Id])
    REFERENCES [dbo].[Settings]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_UserSetting'
CREATE INDEX [IX_FK_UserSetting]
ON [dbo].[Users]
    ([Setting_Id]);
GO

-- Creating foreign key on [UserId] in table 'Mediaries'
ALTER TABLE [dbo].[Mediaries]
ADD CONSTRAINT [FK_UserMediary]
    FOREIGN KEY ([UserId])
    REFERENCES [dbo].[Users]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_UserMediary'
CREATE INDEX [IX_FK_UserMediary]
ON [dbo].[Mediaries]
    ([UserId]);
GO

-- Creating foreign key on [MediaryId] in table 'Resources'
ALTER TABLE [dbo].[Resources]
ADD CONSTRAINT [FK_MediaryResource]
    FOREIGN KEY ([MediaryId])
    REFERENCES [dbo].[Mediaries]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_MediaryResource'
CREATE INDEX [IX_FK_MediaryResource]
ON [dbo].[Resources]
    ([MediaryId]);
GO

-- Creating foreign key on [UserId] in table 'Feeds'
ALTER TABLE [dbo].[Feeds]
ADD CONSTRAINT [FK_UserFeed]
    FOREIGN KEY ([UserId])
    REFERENCES [dbo].[Users]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_UserFeed'
CREATE INDEX [IX_FK_UserFeed]
ON [dbo].[Feeds]
    ([UserId]);
GO

-- Creating foreign key on [UserId] in table 'Likes'
ALTER TABLE [dbo].[Likes]
ADD CONSTRAINT [FK_UserLike]
    FOREIGN KEY ([UserId])
    REFERENCES [dbo].[Users]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_UserLike'
CREATE INDEX [IX_FK_UserLike]
ON [dbo].[Likes]
    ([UserId]);
GO

-- Creating foreign key on [FeedId] in table 'Likes'
ALTER TABLE [dbo].[Likes]
ADD CONSTRAINT [FK_FeedLike]
    FOREIGN KEY ([FeedId])
    REFERENCES [dbo].[Feeds]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_FeedLike'
CREATE INDEX [IX_FK_FeedLike]
ON [dbo].[Likes]
    ([FeedId]);
GO

-- Creating foreign key on [ExchangeId] in table 'Feeds_Message'
ALTER TABLE [dbo].[Feeds_Message]
ADD CONSTRAINT [FK_ExchangeMessage]
    FOREIGN KEY ([ExchangeId])
    REFERENCES [dbo].[Exchanges]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ExchangeMessage'
CREATE INDEX [IX_FK_ExchangeMessage]
ON [dbo].[Feeds_Message]
    ([ExchangeId]);
GO

-- Creating foreign key on [ExchangeId] in table 'Users'
ALTER TABLE [dbo].[Users]
ADD CONSTRAINT [FK_ExchangeUser]
    FOREIGN KEY ([ExchangeId])
    REFERENCES [dbo].[Exchanges]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ExchangeUser'
CREATE INDEX [IX_FK_ExchangeUser]
ON [dbo].[Users]
    ([ExchangeId]);
GO

-- Creating foreign key on [PostId] in table 'Feeds_Comment'
ALTER TABLE [dbo].[Feeds_Comment]
ADD CONSTRAINT [FK_PostComment]
    FOREIGN KEY ([PostId])
    REFERENCES [dbo].[Feeds_Post]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_PostComment'
CREATE INDEX [IX_FK_PostComment]
ON [dbo].[Feeds_Comment]
    ([PostId]);
GO

-- Creating foreign key on [Id] in table 'Feeds_Message'
ALTER TABLE [dbo].[Feeds_Message]
ADD CONSTRAINT [FK_Message_inherits_Feed]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Feeds]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Feeds_Post'
ALTER TABLE [dbo].[Feeds_Post]
ADD CONSTRAINT [FK_Post_inherits_Feed]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Feeds]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Id] in table 'Feeds_Comment'
ALTER TABLE [dbo].[Feeds_Comment]
ADD CONSTRAINT [FK_Comment_inherits_Feed]
    FOREIGN KEY ([Id])
    REFERENCES [dbo].[Feeds]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------