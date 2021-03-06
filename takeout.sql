USE [master]
GO
/****** Object:  Database [takeout]    Script Date: 11/28/2016 22:32:31 ******/
CREATE DATABASE [takeout] ON  PRIMARY 
( NAME = N'takeout', FILENAME = N'C:\数据库实验\takeout\takeout.mdf' , SIZE = 3072KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'takeout_log', FILENAME = N'C:\数据库实验\takeout\takeout_log.ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO
ALTER DATABASE [takeout] SET COMPATIBILITY_LEVEL = 100
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [takeout].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [takeout] SET ANSI_NULL_DEFAULT OFF
GO
ALTER DATABASE [takeout] SET ANSI_NULLS OFF
GO
ALTER DATABASE [takeout] SET ANSI_PADDING OFF
GO
ALTER DATABASE [takeout] SET ANSI_WARNINGS OFF
GO
ALTER DATABASE [takeout] SET ARITHABORT OFF
GO
ALTER DATABASE [takeout] SET AUTO_CLOSE OFF
GO
ALTER DATABASE [takeout] SET AUTO_CREATE_STATISTICS ON
GO
ALTER DATABASE [takeout] SET AUTO_SHRINK OFF
GO
ALTER DATABASE [takeout] SET AUTO_UPDATE_STATISTICS ON
GO
ALTER DATABASE [takeout] SET CURSOR_CLOSE_ON_COMMIT OFF
GO
ALTER DATABASE [takeout] SET CURSOR_DEFAULT  GLOBAL
GO
ALTER DATABASE [takeout] SET CONCAT_NULL_YIELDS_NULL OFF
GO
ALTER DATABASE [takeout] SET NUMERIC_ROUNDABORT OFF
GO
ALTER DATABASE [takeout] SET QUOTED_IDENTIFIER OFF
GO
ALTER DATABASE [takeout] SET RECURSIVE_TRIGGERS OFF
GO
ALTER DATABASE [takeout] SET  DISABLE_BROKER
GO
ALTER DATABASE [takeout] SET AUTO_UPDATE_STATISTICS_ASYNC OFF
GO
ALTER DATABASE [takeout] SET DATE_CORRELATION_OPTIMIZATION OFF
GO
ALTER DATABASE [takeout] SET TRUSTWORTHY OFF
GO
ALTER DATABASE [takeout] SET ALLOW_SNAPSHOT_ISOLATION OFF
GO
ALTER DATABASE [takeout] SET PARAMETERIZATION SIMPLE
GO
ALTER DATABASE [takeout] SET READ_COMMITTED_SNAPSHOT OFF
GO
ALTER DATABASE [takeout] SET HONOR_BROKER_PRIORITY OFF
GO
ALTER DATABASE [takeout] SET  READ_WRITE
GO
ALTER DATABASE [takeout] SET RECOVERY FULL
GO
ALTER DATABASE [takeout] SET  MULTI_USER
GO
ALTER DATABASE [takeout] SET PAGE_VERIFY CHECKSUM
GO
ALTER DATABASE [takeout] SET DB_CHAINING OFF
GO
EXEC sys.sp_db_vardecimal_storage_format N'takeout', N'ON'
GO
USE [takeout]
GO
/****** Object:  Table [dbo].[user]    Script Date: 11/28/2016 22:32:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[user](
	[id] [int] NOT NULL,
	[username] [nvarchar](20) NOT NULL,
	[password] [char](32) NOT NULL,
	[type] [char](1) NOT NULL,
	[balance] [money] NULL,
 CONSTRAINT [PK_user] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
INSERT [dbo].[user] ([id], [username], [password], [type], [balance]) VALUES (1, N'admin', N'1a1dc91c907325c69271ddf0c944bc72', N'A', 0.0000)
INSERT [dbo].[user] ([id], [username], [password], [type], [balance]) VALUES (4, N'岳麓餐馆', N'1a1dc91c907325c69271ddf0c944bc72', N'S', 0.0000)
INSERT [dbo].[user] ([id], [username], [password], [type], [balance]) VALUES (5, N'15367893593', N'1a1dc91c907325c69271ddf0c944bc72', N'U', 76.0000)
/****** Object:  Table [dbo].[menu]    Script Date: 11/28/2016 22:32:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[menu](
	[id] [int] NOT NULL,
	[seller_name] [nvarchar](50) NOT NULL,
	[name] [nvarchar](50) NOT NULL,
	[price] [money] NOT NULL,
	[type] [nvarchar](50) NOT NULL,
	[list] [nvarchar](50) NOT NULL,
	[score] [float] NOT NULL,
	[count] [int] NOT NULL,
	[total] [int] NOT NULL,
 CONSTRAINT [PK_menu] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
INSERT [dbo].[menu] ([id], [seller_name], [name], [price], [type], [list], [score], [count], [total]) VALUES (1, N'岳麓餐馆', N'辣椒炒肉', 12.0000, N'单点', N'', 4.5, 2, 2)
INSERT [dbo].[menu] ([id], [seller_name], [name], [price], [type], [list], [score], [count], [total]) VALUES (2, N'岳麓餐馆', N'土豆炒肉', 11.0000, N'单点', N'', 0, 0, 0)
INSERT [dbo].[menu] ([id], [seller_name], [name], [price], [type], [list], [score], [count], [total]) VALUES (3, N'岳麓餐馆', N'鸡蛋汤', 3.0000, N'单点', N'', 0, 0, 0)
INSERT [dbo].[menu] ([id], [seller_name], [name], [price], [type], [list], [score], [count], [total]) VALUES (4, N'岳麓餐馆', N'家常', 14.0000, N'中餐', N'辣椒炒肉,鸡蛋汤', 0, 0, 0)
/****** Object:  Table [dbo].[history]    Script Date: 11/28/2016 22:32:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[history](
	[id] [int] NOT NULL,
	[telephone] [nvarchar](20) NOT NULL,
	[menu_id] [int] NOT NULL,
	[method] [nvarchar](50) NOT NULL,
	[score] [float] NOT NULL,
	[datetime] [datetime] NOT NULL,
 CONSTRAINT [PK_history] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
INSERT [dbo].[history] ([id], [telephone], [menu_id], [method], [score], [datetime]) VALUES (1, N'15367893593', 1, N'盒饭', 5, CAST(0x0000A6CC016C56B4 AS DateTime))
INSERT [dbo].[history] ([id], [telephone], [menu_id], [method], [score], [datetime]) VALUES (2, N'15367893593', 1, N'盒饭', 4, CAST(0x0000A6CC016E93C0 AS DateTime))
/****** Object:  Default [DF_menu_score]    Script Date: 11/28/2016 22:32:31 ******/
ALTER TABLE [dbo].[menu] ADD  CONSTRAINT [DF_menu_score]  DEFAULT ((0)) FOR [score]
GO
/****** Object:  Default [DF_menu_count]    Script Date: 11/28/2016 22:32:31 ******/
ALTER TABLE [dbo].[menu] ADD  CONSTRAINT [DF_menu_count]  DEFAULT ((0)) FOR [count]
GO
/****** Object:  Default [DF_menu_total]    Script Date: 11/28/2016 22:32:31 ******/
ALTER TABLE [dbo].[menu] ADD  CONSTRAINT [DF_menu_total]  DEFAULT ((0)) FOR [total]
GO
