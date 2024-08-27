USE [master]
GO
/****** Object:  Database [SB_IM]    Script Date: 8/27/2024 11:06:35 AM ******/
CREATE DATABASE [SB_IM]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'SB_IM', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\DATA\SB_IM.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'SB_IM_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\DATA\SB_IM_log.ldf' , SIZE = 8192KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT
GO
ALTER DATABASE [SB_IM] SET COMPATIBILITY_LEVEL = 150
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [SB_IM].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [SB_IM] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [SB_IM] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [SB_IM] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [SB_IM] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [SB_IM] SET ARITHABORT OFF 
GO
ALTER DATABASE [SB_IM] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [SB_IM] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [SB_IM] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [SB_IM] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [SB_IM] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [SB_IM] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [SB_IM] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [SB_IM] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [SB_IM] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [SB_IM] SET  DISABLE_BROKER 
GO
ALTER DATABASE [SB_IM] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [SB_IM] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [SB_IM] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [SB_IM] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [SB_IM] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [SB_IM] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [SB_IM] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [SB_IM] SET RECOVERY FULL 
GO
ALTER DATABASE [SB_IM] SET  MULTI_USER 
GO
ALTER DATABASE [SB_IM] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [SB_IM] SET DB_CHAINING OFF 
GO
ALTER DATABASE [SB_IM] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [SB_IM] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [SB_IM] SET DELAYED_DURABILITY = DISABLED 
GO
ALTER DATABASE [SB_IM] SET ACCELERATED_DATABASE_RECOVERY = OFF  
GO
EXEC sys.sp_db_vardecimal_storage_format N'SB_IM', N'ON'
GO
ALTER DATABASE [SB_IM] SET QUERY_STORE = OFF
GO
USE [SB_IM]
GO
/****** Object:  Table [dbo].[Sys_SsoRcvDetail_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoRcvDetail_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[RcvHOID] [int] NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[BRGNAME] [varchar](100) NOT NULL,
	[BRGUNIT] [varchar](50) NOT NULL,
	[RcvDQty] [int] NOT NULL,
	[RcvDNote] [varchar](450) NULL,
	[RcvDNoteUserOID] [int] NULL,
	[RcvDNoteDatetime] [datetime] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoRcvScan_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoRcvScan_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[RcvHOID] [int] NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[RcvScanNote] [varchar](100) NOT NULL,
	[RcvScanQty] [int] NOT NULL,
	[RcvScanUserOID] [int] NOT NULL,
	[RcvScanDatetime] [datetime] NOT NULL,
	[RcvScanDeleted] [bit] NOT NULL,
	[RcvScanDeletedNote] [varchar](100) NULL,
	[RcvScanDeletedUserOID] [int] NULL,
	[RcvScanDeletedDatetime] [datetime] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoUser_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoUser_MA](
	[OID] [int] NOT NULL,
	[UserID] [varchar](45) NOT NULL,
	[UserNip] [varchar](10) NOT NULL,
	[UserName] [varchar](85) NOT NULL,
	[UserAdmin] [bit] NOT NULL,
	[UserGroupOID] [int] NOT NULL,
	[UserPassword] [varchar](245) NOT NULL,
	[UserLocationOID] [int] NOT NULL,
	[UserCompanyCode] [varchar](10) NULL,
	[Status] [varchar](10) NULL,
	[CreationDatetime] [datetime] NULL,
	[CreationUserOID] [int] NULL,
	[ModificationDatetime] [datetime] NULL,
	[ModificationUserOID] [int] NULL,
 CONSTRAINT [PK_Sys_SsoUser_MA] PRIMARY KEY CLUSTERED 
(
	[OID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  UserDefinedFunction [dbo].[fnTbl_SsoRcvScan]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create function [dbo].[fnTbl_SsoRcvScan](@vriRcvHOID int)
returns table as
return(
Select isnull(sso.OID,0) OID,
       isnull(sso.RcvHOID,ssc.RcvHOID)RcvHOID,
	   isnull(sso.BRGCODE,ssc.BRGCODE)BRGCODE,
       isnull(sso.RcvDQty,0)RcvDQty,
	   isnull(ssc.vSumRcvScanQty,0)vSumRcvScanQty,
	   (isnull(sso.RcvDQty,0) - isnull(ssc.vSumRcvScanQty,0)) vRcvScanVarian,
	   sso.RcvDNote,
	   replace(sso.RcvDNote,char(10),'<br />')vRcvDNote,
	   sso.RcvDNoteUserOID,sso.RcvDNoteDatetime,
	   ssu.UserName vRcvDNoteBy,
	   (convert(varchar(11),sso.RcvDNoteDatetime,106) +' '+ convert(varchar(5),sso.RcvDNoteDatetime,108)) vRcvDNoteDatetime
       From (Select * From Sys_SsoRcvDetail_TR with(nolock) Where RcvHOID=@vriRcvHOID)sso
	        left outer join Sys_SsoUser_MA ssu with(nolock) on ssu.OID=sso.RcvDNoteUserOID
	        full outer join
			     (Select b.RcvHOID,b.BRGCODE,sum(b.RcvScanQty)vSumRcvScanQty
				         From Sys_SsoRcvScan_TR b with(nolock)
						Where b.RcvHOID=@vriRcvHOID
						Group by b.RcvHOID,b.BRGCODE) ssc on ssc.RcvHOID=sso.RcvHOID and ssc.BRGCODE=sso.BRGCODE
)
GO
/****** Object:  Table [dbo].[Sys_SsoLocation_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoLocation_MA](
	[OID] [int] NOT NULL,
	[LocationCode] [varchar](2) NOT NULL,
	[LocationName] [varchar](45) NOT NULL,
	[LocationCity] [varchar](45) NULL,
 CONSTRAINT [PK_Sys_SsoLocation_MA] PRIMARY KEY CLUSTERED 
(
	[OID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoRcvHeader_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoRcvHeader_TR](
	[OID] [int] NOT NULL,
	[RcvNo] [varchar](45) NOT NULL,
	[RcvDate] [date] NOT NULL,
	[RcvCompanyCode] [varchar](15) NOT NULL,
	[RcvGdgCode] [varchar](15) NOT NULL,
	[RcvLocationOID] [int] NOT NULL,
	[RcvNote] [varchar](450) NOT NULL,
	[RcvCancelNote] [varchar](450) NULL,
	[RcvCloseNote] [varchar](450) NULL,
	[TransCode] [varchar](4) NOT NULL,
	[TransStatus] [int] NOT NULL,
	[CreationDatetime] [datetime] NOT NULL,
	[CreationUserOID] [int] NOT NULL,
	[ModificationDatetime] [datetime] NULL,
	[ModificationUserOID] [int] NULL,
	[ScanOpenDatetime] [datetime] NULL,
	[ScanOpenUserOID] [int] NULL,
	[ScanClosedDatetime] [datetime] NULL,
	[ScanClosedUserOID] [int] NULL,
	[ClosedDatetime] [datetime] NULL,
	[ClosedUserOID] [int] NULL,
	[CancelledDatetime] [datetime] NULL,
	[CancelledUserOID] [int] NULL,
 CONSTRAINT [PK_Sys_SsoRcvHeader_TR] PRIMARY KEY CLUSTERED 
(
	[OID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [Un_Sys_SsoRcvHeader_TR] UNIQUE NONCLUSTERED 
(
	[RcvNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoTransStatus_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoTransStatus_MA](
	[TransCode] [varchar](4) NOT NULL,
	[TransStatus] [smallint] NOT NULL,
	[TransStatusDescr] [varchar](50) NOT NULL,
 CONSTRAINT [Un_Sys_SsoTransStatus_MA] UNIQUE NONCLUSTERED 
(
	[TransCode] ASC,
	[TransStatus] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  UserDefinedFunction [dbo].[fnTbl_SsoTallyRcv]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Function [dbo].[fnTbl_SsoTallyRcv](@vriHOID int,@vriUser varchar(100))returns table
as
return(
Select sh.OID,sh.RcvNo,sh.RcvDate,sh.RcvNote,
       sh.RcvLocationOID,sh.RcvCompanyCode,sh.RcvGdgCode,
	   lm.LocationName,
	   sh.RcvCloseNote,sh.RcvCancelNote,
	   sh.TransCode,sh.TransStatus,
	   st.TransStatusDescr,
	   sh.CreationUserOID,sh.CreationDatetime,
	   sh.ScanOpenUserOID,sh.ScanOpenDatetime,
	   sh.ScanClosedUserOID,sh.ScanClosedDatetime,
	   sh.ClosedUserOID,sh.ClosedDatetime,
	   sh.CancelledUserOID,sh.CancelledDatetime,
       sd.OID vDOID,
       sd.BRGCODE,sd.RcvDQty,sd.vSumRcvScanQty,sd.vRcvScanVarian,
       sd.RcvDNote,sd.vRcvDNoteBy,sd.vRcvDNoteDatetime,
	   Convert(varchar(11),getdate(),106)+' '+Convert(varchar(5),getdate(),108) vPrintDate,
	   @vriUser vPrintUser
  From Sys_SsoRcvHeader_TR sh with(nolock)
       inner join Sys_SsoLocation_MA lm with(nolock) on lm.OID=sh.RcvLocationOID
       inner join Sys_SsoTransStatus_MA st with(nolock) on st.TransCode=sh.TransCode and st.TransStatus=sh.TransStatus
       inner join fnTbl_SsoRcvScan(@vriHOID) sd on sd.RcvHOID=sh.OID
)

GO
/****** Object:  Table [dbo].[Sys_SsoSOScan_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoSOScan_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[SOHOID] [int] NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[SOScanNote] [varchar](100) NOT NULL,
	[SOScanQty] [int] NOT NULL,
	[SOScanUserOID] [int] NOT NULL,
	[SOScanDatetime] [datetime] NOT NULL,
	[SOScanDeleted] [bit] NOT NULL,
	[SOScanDeletedNote] [varchar](100) NULL,
	[SOScanDeletedUserOID] [int] NULL,
	[SOScanDeletedDatetime] [datetime] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoSOStock_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoSOStock_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[SOHOID] [int] NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[BRGNAME] [varchar](100) NOT NULL,
	[BRGUNIT] [varchar](50) NOT NULL,
	[SOStockQty] [int] NOT NULL,
	[SOStockNote] [varchar](450) NULL,
	[SOStockNoteUserOID] [int] NULL,
	[SOStockNoteDatetime] [datetime] NULL
) ON [PRIMARY]
GO
/****** Object:  UserDefinedFunction [dbo].[fnTbl_SsoSOStockScan]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE function [dbo].[fnTbl_SsoSOStockScan](@vriSOHOID int)
returns table as
return(
Select isnull(sso.OID,0) OID,
       isnull(sso.SOHOID,ssc.SOHOID)SOHOID,
	   isnull(sso.BRGCODE,ssc.BRGCODE)BRGCODE,
       isnull(sso.SOStockQty,0)SOStockQty,
	   isnull(ssc.vSumSOScanQty,0)vSumSOScanQty,
	   (isnull(sso.SOStockQty,0) - isnull(ssc.vSumSOScanQty,0)) vSOStockScanVarian,
	   sso.SOStockNote,
	   replace(sso.SOStockNote,char(10),'<br />')vSOStockNote,
	   sso.SOStockNoteUserOID,sso.SOStockNoteDatetime,
	   ssu.UserName vSOStockNoteBy,
	   (convert(varchar(11),sso.SOStockNoteDatetime,106) +' '+ convert(varchar(5),sso.SOStockNoteDatetime,108)) vSOStockNoteDatetime
       From (Select * From Sys_SsoSOStock_TR with(nolock) Where SOHOID=@vriSOHOID)sso
	        left outer join Sys_SsoUser_MA ssu with(nolock) on ssu.OID=sso.SOStockNoteUserOID
	        full outer join
			     (Select b.SOHOID,b.BRGCODE,sum(b.SOScanQty)vSumSOScanQty
				         From Sys_SsoSOScan_TR b with(nolock)
						Where b.SOHOID=@vriSOHOID
						Group by b.SOHOID,b.BRGCODE) ssc on ssc.SOHOID=sso.SOHOID and ssc.BRGCODE=sso.BRGCODE

--Select sso.OID,sso.SOHOID,sso.BRGCODE,sso.BRGNAME,sso.BRGUNIT,
--       sso.SOStockQty,ssc.vSumSOScanQty,
--	   (sso.SOStockQty - ssc.vSumSOScanQty) vSOStockScanVarian,
--	   sso.SOStockNote,sso.SOStockNoteUserOID,sso.SOStockNoteDatetime,
--	   ssu.UserName vSOStockNoteBy,
--	   (convert(varchar(11),sso.SOStockNoteDatetime,106) +' '+ convert(varchar(5),sso.SOStockNoteDatetime,108)) vSOStockNoteDatetime
--       From Sys_SsoSOStock_TR sso with(nolock)
--	        left outer join Sys_SsoUser_MA ssu with(nolock) on ssu.OID=sso.SOStockNoteUserOID
--	        left outer join
--			     (Select b.SOHOID,b.BRGCODE,sum(b.SOScanQty)vSumSOScanQty
--				         From Sys_SsoSOScan_TR b with(nolock)
--						Where b.SOHOID=@vriSOHOID
--						Group by b.SOHOID,b.BRGCODE) ssc on ssc.SOHOID=sso.SOHOID and ssc.BRGCODE=sso.BRGCODE
--	 Where sso.SOHOID=@vriSOHOID
)
GO
/****** Object:  Table [dbo].[Sys_SsoSOHeader_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoSOHeader_TR](
	[OID] [int] NOT NULL,
	[SONo] [varchar](45) NOT NULL,
	[SOCutOff] [datetime] NOT NULL,
	[SOCompanyCode] [varchar](15) NOT NULL,
	[SOGdgCode] [varchar](15) NOT NULL,
	[SOLocationOID] [int] NOT NULL,
	[SONote] [varchar](450) NOT NULL,
	[SOCancelNote] [varchar](450) NULL,
	[SOCloseNote] [varchar](450) NULL,
	[SOXlsFileName] [varchar](250) NOT NULL,
	[SOXlsSheetName] [varchar](50) NOT NULL,
	[TransCode] [varchar](4) NOT NULL,
	[TransStatus] [int] NOT NULL,
	[CreationDatetime] [datetime] NOT NULL,
	[CreationUserOID] [int] NOT NULL,
	[ModificationDatetime] [datetime] NULL,
	[ModificationUserOID] [int] NULL,
	[ScanOpenDatetime] [datetime] NULL,
	[ScanOpenUserOID] [int] NULL,
	[ScanClosedDatetime] [datetime] NULL,
	[ScanClosedUserOID] [int] NULL,
	[ClosedDatetime] [datetime] NULL,
	[ClosedUserOID] [int] NULL,
	[CancelledDatetime] [datetime] NULL,
	[CancelledUserOID] [int] NULL,
 CONSTRAINT [PK_Sys_SsoSOHeader_TR] PRIMARY KEY CLUSTERED 
(
	[OID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [Un_Sys_SsoSOHeader_TR] UNIQUE NONCLUSTERED 
(
	[SONo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  UserDefinedFunction [dbo].[fnTbl_SsoTally]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Function [dbo].[fnTbl_SsoTally](@vriHOID int,@vriUser varchar(100))returns table
as
return(
Select sh.OID,sh.SONo,sh.SOCutOff,sh.SONote,
       sh.SOLocationOID,sh.SOCompanyCode,sh.SOGdgCode,
	   lm.LocationName,
	   sh.SOCloseNote,sh.SOCancelNote,
	   sh.TransCode,sh.TransStatus,
	   st.TransStatusDescr,
	   sh.CreationUserOID,sh.CreationDatetime,
	   sh.ScanOpenUserOID,sh.ScanOpenDatetime,
	   sh.ScanClosedUserOID,sh.ScanClosedDatetime,
	   sh.ClosedUserOID,sh.ClosedDatetime,
	   sh.CancelledUserOID,sh.CancelledDatetime,
       sd.OID vDOID,
       sd.BRGCODE,sd.SOStockQty,sd.vSumSOScanQty,sd.vSOStockScanVarian,
       sd.SOStockNote,sd.vSOStockNoteBy,sd.vSOStockNoteDatetime,
	   Convert(varchar(11),getdate(),106)+' '+Convert(varchar(5),getdate(),108) vPrintDate,
	   @vriUser vPrintUser
  From Sys_SsoSOHeader_TR sh with(nolock)
       inner join Sys_SsoLocation_MA lm with(nolock) on lm.OID=sh.SOLocationOID
       inner join Sys_SsoTransStatus_MA st with(nolock) on st.TransCode=sh.TransCode and st.TransStatus=sh.TransStatus
       inner join fnTbl_SsoSOStockScan(@vriHOID) sd on sd.SOHOID=sh.OID
)

GO
/****** Object:  Table [dbo].[Sys_SsoPickDetail_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoPickDetail_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[PickHOID] [int] NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[BRGNAME] [varchar](100) NOT NULL,
	[BRGUNIT] [varchar](50) NOT NULL,
	[PickDQty] [int] NOT NULL,
	[PickDQtyBonus] [int] NOT NULL,
	[PickDNote] [varchar](450) NULL,
	[PickDNoteUserOID] [int] NULL,
	[PickDNoteDatetime] [datetime] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoPickScan_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoPickScan_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[PickHOID] [int] NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[PickScanSerialNo] [varchar](45) NULL,
	[PickScanNote] [varchar](100) NOT NULL,
	[PickScanQty] [int] NOT NULL,
	[PickScanUserOID] [int] NOT NULL,
	[PickScanDatetime] [datetime] NOT NULL,
	[PickScanDeleted] [bit] NOT NULL,
	[PickScanDeletedNote] [varchar](100) NULL,
	[PickScanDeletedUserOID] [int] NULL,
	[PickScanDeletedDatetime] [datetime] NULL
) ON [PRIMARY]
GO
/****** Object:  UserDefinedFunction [dbo].[fnTbl_SsoPickScan]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE function [dbo].[fnTbl_SsoPickScan](@vriPickHOID int)
returns table as
return(
Select isnull(sso.OID,0) OID,
       isnull(sso.PickHOID,ssc.PickHOID)SOHOID,
	   isnull(sso.BRGCODE,ssc.BRGCODE)BRGCODE,
       isnull(sso.PickDQty,0)+isnull(sso.PickDQtyBonus,0)vPickDQtyTotal,
	   isnull(ssc.vSumPickScanQty,0)vSumPickScanQty,
	   (isnull(sso.PickDQty,0) - isnull(ssc.vSumPickScanQty,0)) vPickScanVarian,
	   sso.PickDNote,
	   replace(sso.PickDNote,char(10),'<br />')vPickDNote,
	   sso.PickDNoteUserOID,sso.PickDNoteDatetime,
	   ssu.UserName vPickDNoteBy,
	   (convert(varchar(11),sso.PickDNoteDatetime,106) +' '+ convert(varchar(5),sso.PickDNoteDatetime,108)) vPickDNoteDatetime
       From (Select * From Sys_SsoPickDetail_TR with(nolock) Where PickHOID=@vriPickHOID)sso
	        left outer join Sys_SsoUser_MA ssu with(nolock) on ssu.OID=sso.PickDNoteUserOID
	        full outer join
			     (Select b.PickHOID,b.BRGCODE,sum(b.PickScanQty)vSumPickScanQty
				         From Sys_SsoPickScan_TR b with(nolock)
						Where b.PickHOID=@vriPickHOID and abs(b.PickScanDeleted)=0
						Group by b.PickHOID,b.BRGCODE) ssc on ssc.PickHOID=sso.PickHOID and ssc.BRGCODE=sso.BRGCODE
)
GO
/****** Object:  UserDefinedFunction [dbo].[fnTbl_SsoRcvDetailScan]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE function [dbo].[fnTbl_SsoRcvDetailScan](@vriRcvHOID int)
returns table as
return(
Select isnull(sso.OID,0) OID,
       isnull(sso.RcvHOID,ssc.RcvHOID)RcvHOID,
	   isnull(sso.BRGCODE,ssc.BRGCODE)BRGCODE,
       isnull(sso.RcvDQty,0)RcvDQty,
	   isnull(ssc.vSumRcvScanQty,0)vSumRcvScanQty,
	   (isnull(sso.RcvDQty,0) - isnull(ssc.vSumRcvScanQty,0)) vSumRcvScanVarian,
	   sso.RcvDNote,
	   replace(sso.RcvDNote,char(10),'<br />')vRcvDNote,
	   sso.RcvDNoteUserOID,sso.RcvDNoteDatetime,
	   ssu.UserName vRcvDNoteBy,
	   (convert(varchar(11),sso.RcvDNoteDatetime,106) +' '+ convert(varchar(5),sso.RcvDNoteDatetime,108)) vRcvDNoteDatetime
       From (Select * From Sys_SsoRcvDetail_TR with(nolock) Where RcvHOID=@vriRcvHOID)sso
	        left outer join Sys_SsoUser_MA ssu with(nolock) on ssu.OID=sso.RcvDNoteUserOID
	        full outer join
			     (Select b.RcvHOID,b.BRGCODE,sum(b.RcvScanQty)vSumRcvScanQty
				         From Sys_SsoRcvScan_TR b with(nolock)
						Where b.RcvHOID=@vriRcvHOID
						Group by b.RcvHOID,b.BRGCODE) ssc on ssc.RcvHOID=sso.RcvHOID and ssc.BRGCODE=sso.BRGCODE
)
GO
/****** Object:  UserDefinedFunction [dbo].[fnTbl_SsoUserSso]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE function [dbo].[fnTbl_SsoUserSso](@vriUserID varchar(100))returns table as
Return
(
Select OID,UserID,UserNip,UserName,UserAdmin,UserGroupOID,UserCompanyCode,UserLocationOID,UserPassword,Status,CreationDatetime,CreationUserOID,ModificationDatetime,ModificationUserOID
       From Sys_SsoUser_MA Where UserID=@vriUserID
)


GO
/****** Object:  UserDefinedFunction [dbo].[fnTbl_SsoUserSsoByUID]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create function [dbo].[fnTbl_SsoUserSsoByUID](@vriUserID varchar(100),@vriUserPwd varchar(255))returns table as
Return
(
Select OID,UserID,UserNip,UserName,UserAdmin,UserGroupOID,UserCompanyCode,UserLocationOID,UserPassword,Status,CreationDatetime,CreationUserOID,ModificationDatetime,ModificationUserOID
       From Sys_SsoUser_MA
	   Where UserID=@vriUserID and UserPassword=@vriUserPwd
)


GO
/****** Object:  Table [dbo].[Sys_SsoPickHeader_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoPickHeader_TR](
	[OID] [int] NOT NULL,
	[PickNo] [varchar](45) NOT NULL,
	[SchDTypeOID] [int] NOT NULL,
	[PickRefOID] [int] NOT NULL,
	[PickRefNo] [varchar](45) NOT NULL,
	[PickRefDate] [date] NOT NULL,
	[PickDate] [date] NOT NULL,
	[PickCompanyCode] [varchar](15) NOT NULL,
	[PickLocationOID] [int] NOT NULL,
	[PickGdgCode] [varchar](15) NOT NULL,
	[PickGdgCodeName] [varchar](145) NOT NULL,
	[PickGdgCodeTujuan] [varchar](15) NULL,
	[PickGdgCodeTujuanName] [varchar](145) NULL,
	[PickCustCode] [varchar](50) NULL,
	[PickCustName] [varchar](150) NULL,
	[PickNote] [varchar](450) NOT NULL,
	[PickCancelNote] [varchar](450) NULL,
	[PickCloseNote] [varchar](450) NULL,
	[TransCode] [varchar](4) NOT NULL,
	[TransStatus] [int] NOT NULL,
	[CreationDatetime] [datetime] NOT NULL,
	[CreationUserOID] [int] NOT NULL,
	[ModificationDatetime] [datetime] NULL,
	[ModificationUserOID] [int] NULL,
	[ScanOpenDatetime] [datetime] NULL,
	[ScanOpenUserOID] [int] NULL,
	[ScanClosedDatetime] [datetime] NULL,
	[ScanClosedUserOID] [int] NULL,
	[ClosedDatetime] [datetime] NULL,
	[ClosedUserOID] [int] NULL,
	[CancelledDatetime] [datetime] NULL,
	[CancelledUserOID] [int] NULL,
 CONSTRAINT [PK_Sys_SsoPickHeader_TR] PRIMARY KEY CLUSTERED 
(
	[OID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoPickStatus_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoPickStatus_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[PickHOID] [int] NOT NULL,
	[TransCode] [varchar](4) NOT NULL,
	[TransStatus] [int] NOT NULL,
	[TransStatusUserOID] [int] NOT NULL,
	[TransStatusDatetime] [datetime] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoPrintQRBarang_Temp]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoPrintQRBarang_Temp](
	[OID] [int] NULL,
	[UserOID] [int] NULL,
	[BcProductGenCode011] [varchar](45) NULL,
	[BcProductGenCodeImg011] [image] NULL,
	[BcProductGenCode012] [varchar](45) NULL,
	[BcProductGenCodeImg012] [image] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoPrintQRBarang_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoPrintQRBarang_TR](
	[OID] [int] NOT NULL,
	[CompanyCode] [varchar](5) NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[BRGNAME] [varchar](100) NOT NULL,
	[BRGUNIT] [varchar](50) NOT NULL,
	[PrintCount] [int] NOT NULL,
	[PrintNote] [varchar](450) NOT NULL,
	[PrintDatetime] [datetime] NOT NULL,
	[PrintUserOID] [int] NOT NULL,
 CONSTRAINT [Un_Sys_SsoPrintQRBarang_TR] UNIQUE NONCLUSTERED 
(
	[OID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoPrintSNRequest_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoPrintSNRequest_TR](
	[OID] [int] NOT NULL,
	[CompanyCode] [varchar](5) NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[BRGNAME] [varchar](100) NOT NULL,
	[BRGUNIT] [varchar](50) NOT NULL,
	[ReqSN] [varchar](50) NOT NULL,
	[ReqSNNote] [varchar](450) NOT NULL,
	[ReqSNDatetime] [datetime] NOT NULL,
	[ReqSNUserOID] [int] NOT NULL,
	[PrintSNDatetime] [datetime] NULL,
	[PrintSNUserOID] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoRcvStatus_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoRcvStatus_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[RcvHOID] [int] NOT NULL,
	[TransCode] [varchar](4) NOT NULL,
	[TransStatus] [int] NOT NULL,
	[TransStatusUserOID] [int] NOT NULL,
	[TransStatusDatetime] [datetime] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoSOStatus_TR]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoSOStatus_TR](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[SOHOID] [int] NOT NULL,
	[TransCode] [varchar](4) NOT NULL,
	[TransStatus] [int] NOT NULL,
	[TransStatusUserOID] [int] NOT NULL,
	[TransStatusDatetime] [datetime] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoSOStock_HS]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoSOStock_HS](
	[OID] [int] IDENTITY(1,1) NOT NULL,
	[SOSOID] [int] NOT NULL,
	[SOHOID] [int] NOT NULL,
	[BRGCODE] [varchar](20) NOT NULL,
	[BRGNAME] [varchar](100) NOT NULL,
	[BRGUNIT] [varchar](50) NOT NULL,
	[SOStockQty] [int] NOT NULL,
	[SOStockNote] [varchar](450) NULL,
	[SOStockNoteUserOID] [int] NULL,
	[SOStockNoteDatetime] [datetime] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoTrAccess_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoTrAccess_MA](
	[TrAccessCode] [varchar](4) NOT NULL,
	[TrAccessName] [varchar](50) NOT NULL,
 CONSTRAINT [PK_Sys_SsoTrAccess_MA] PRIMARY KEY CLUSTERED 
(
	[TrAccessCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoTransAccess_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoTransAccess_MA](
	[TransCode] [varchar](4) NULL,
	[TrAccessCode] [varchar](4) NULL,
 CONSTRAINT [Un_Sys_SsoTransAccess_MA] UNIQUE NONCLUSTERED 
(
	[TransCode] ASC,
	[TrAccessCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoTransName_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoTransName_MA](
	[TransCode] [varchar](4) NOT NULL,
	[TransName] [varchar](50) NOT NULL,
	[IsTransMenu] [bit] NOT NULL,
 CONSTRAINT [PK_Sys_SsoTransName_MA] PRIMARY KEY CLUSTERED 
(
	[TransCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoUserCompany_HS]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoUserCompany_HS](
	[UserOID] [int] NOT NULL,
	[CompanyCode] [varchar](10) NOT NULL,
	[HistorySeq] [int] NOT NULL,
	[HistoryDatetime] [datetime] NOT NULL,
	[HistoryUserOID] [int] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoUserCompany_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoUserCompany_MA](
	[UserOID] [int] NOT NULL,
	[CompanyCode] [varchar](10) NOT NULL,
 CONSTRAINT [Un_Sys_SsoUserCompany_MA] UNIQUE NONCLUSTERED 
(
	[UserOID] ASC,
	[CompanyCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoUserGroup_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoUserGroup_MA](
	[OID] [int] NOT NULL,
	[SsoUserGroupName] [varchar](45) NOT NULL,
	[SsoUserGroupDescr] [varchar](450) NOT NULL,
	[Status] [varchar](10) NOT NULL,
	[CreationDatetime] [datetime] NOT NULL,
	[CreationUserOID] [int] NOT NULL,
	[ModificationDatetime] [datetime] NULL,
	[ModificationUserOID] [int] NULL,
 CONSTRAINT [PK_Sys_SsoUserGroup_MA] PRIMARY KEY CLUSTERED 
(
	[OID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoUserGroupAccess_MA]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoUserGroupAccess_MA](
	[UserGroupOID] [int] NOT NULL,
	[TransCode] [varchar](4) NOT NULL,
	[TrAccessCode] [varchar](4) NOT NULL,
 CONSTRAINT [Un_Sys_SsoUserGroupAccess_MA] UNIQUE NONCLUSTERED 
(
	[UserGroupOID] ASC,
	[TransCode] ASC,
	[TrAccessCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sys_SsoUserUserPwd_HS]    Script Date: 8/27/2024 11:06:35 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sys_SsoUserUserPwd_HS](
	[UserOID] [int] NOT NULL,
	[ChangePwdDatetime] [datetime] NOT NULL,
	[ChangePwdUserOID] [int] NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_RcvNo]  DEFAULT (NULL) FOR [PickNo]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_PickNo1]  DEFAULT (NULL) FOR [PickRefNo]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_PickDate1]  DEFAULT (NULL) FOR [PickRefDate]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_RcvDate]  DEFAULT (NULL) FOR [PickDate]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_TransStatus]  DEFAULT ((0)) FOR [TransStatus]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_CreationDatetime]  DEFAULT (NULL) FOR [CreationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_CreationUserOID]  DEFAULT (NULL) FOR [CreationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_ModificationDatetime]  DEFAULT (NULL) FOR [ModificationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_ModificationUserOID]  DEFAULT (NULL) FOR [ModificationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_ScanOpenUserOID]  DEFAULT (NULL) FOR [ScanOpenUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_ScanClosedUserOID]  DEFAULT (NULL) FOR [ScanClosedUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_ClosedUserOID]  DEFAULT (NULL) FOR [ClosedUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoPickHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoPickHeader_TR_CancelledUserOID]  DEFAULT (NULL) FOR [CancelledUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoPickScan_TR] ADD  CONSTRAINT [DF_Sys_SsoPickScan_TR_IsDeleted]  DEFAULT ((0)) FOR [PickScanDeleted]
GO
ALTER TABLE [dbo].[Sys_SsoPickStatus_TR] ADD  CONSTRAINT [DF_Sys_SsoPickStatus_TR_TransStatus]  DEFAULT ((0)) FOR [TransStatus]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_SONo]  DEFAULT (NULL) FOR [RcvNo]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_SOCutOff]  DEFAULT (NULL) FOR [RcvDate]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_TransStatus]  DEFAULT ((0)) FOR [TransStatus]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_CreationDatetime]  DEFAULT (NULL) FOR [CreationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_CreationUserOID]  DEFAULT (NULL) FOR [CreationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_ModificationDatetime]  DEFAULT (NULL) FOR [ModificationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_ModificationUserOID]  DEFAULT (NULL) FOR [ModificationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_ScanOpenUserOID]  DEFAULT (NULL) FOR [ScanOpenUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_ScanClosedUserOID]  DEFAULT (NULL) FOR [ScanClosedUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_ClosedUserOID]  DEFAULT (NULL) FOR [ClosedUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoRcvHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvHeader_TR_CancelledUserOID]  DEFAULT (NULL) FOR [CancelledUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoRcvScan_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvScan_TR_PickScanDeleted]  DEFAULT ((0)) FOR [RcvScanDeleted]
GO
ALTER TABLE [dbo].[Sys_SsoRcvStatus_TR] ADD  CONSTRAINT [DF_Sys_SsoRcvStatus_TR_TransStatus]  DEFAULT ((0)) FOR [TransStatus]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_DcmPLNo]  DEFAULT (NULL) FOR [SONo]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_DcmPLDate]  DEFAULT (NULL) FOR [SOCutOff]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_TransStatus]  DEFAULT ((0)) FOR [TransStatus]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_CreationDatetime]  DEFAULT (NULL) FOR [CreationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_CreationUserOID]  DEFAULT (NULL) FOR [CreationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_CreationDatetime1]  DEFAULT (NULL) FOR [ModificationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_CreationUserOID1]  DEFAULT (NULL) FOR [ModificationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_ScanEndUserOID1]  DEFAULT (NULL) FOR [ScanOpenUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_PreparedUserOID1]  DEFAULT (NULL) FOR [ScanClosedUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_ScanEndUserOID1_1]  DEFAULT (NULL) FOR [ClosedUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoSOHeader_TR] ADD  CONSTRAINT [DF_Sys_SsoSOHeader_TR_ClosedUserOID1]  DEFAULT (NULL) FOR [CancelledUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoSOScan_TR] ADD  CONSTRAINT [DF_Sys_SsoSOScan_TR_RcvScanDeleted]  DEFAULT ((0)) FOR [SOScanDeleted]
GO
ALTER TABLE [dbo].[Sys_SsoSOStatus_TR] ADD  CONSTRAINT [DF_Sys_SsoSOStatus_TR_TransStatus]  DEFAULT ((0)) FOR [TransStatus]
GO
ALTER TABLE [dbo].[Sys_SsoTransName_MA] ADD  CONSTRAINT [DF_Sys_SsoTransName_MA_IsTransMenu]  DEFAULT ((1)) FOR [IsTransMenu]
GO
ALTER TABLE [dbo].[Sys_SsoTransStatus_MA] ADD  CONSTRAINT [DF_Sys_SsoTransStatus_MA_TransStatus]  DEFAULT ((0)) FOR [TransStatus]
GO
ALTER TABLE [dbo].[Sys_SsoTransStatus_MA] ADD  CONSTRAINT [DF_Sys_SsoTransStatus_MA_TransStatusDescr]  DEFAULT ((0)) FOR [TransStatusDescr]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_UserID]  DEFAULT (NULL) FOR [UserID]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_UserNip]  DEFAULT ('') FOR [UserNip]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_UserName]  DEFAULT (NULL) FOR [UserName]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_UserAdmin]  DEFAULT (NULL) FOR [UserAdmin]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_UserGroup]  DEFAULT ((1)) FOR [UserGroupOID]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_UserPassword]  DEFAULT (NULL) FOR [UserPassword]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_Status]  DEFAULT (NULL) FOR [Status]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_CreationDatetime]  DEFAULT (NULL) FOR [CreationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_CreationUserOID]  DEFAULT (NULL) FOR [CreationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_ModificationDatetime]  DEFAULT (NULL) FOR [ModificationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] ADD  CONSTRAINT [DF_Sys_SsoUser_MA_ModificationUserOID]  DEFAULT (NULL) FOR [ModificationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoUserCompany_HS] ADD  CONSTRAINT [DF_Sys_SsoUserCompany_HS_CreationDatetime]  DEFAULT (NULL) FOR [HistoryDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoUserCompany_HS] ADD  CONSTRAINT [DF_Sys_SsoUserCompany_HS_CreationUserOID]  DEFAULT (NULL) FOR [HistoryUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoUserGroup_MA] ADD  CONSTRAINT [DF_Sys_SsoUserGroup_MA_DcmUserGroup]  DEFAULT ((0)) FOR [OID]
GO
ALTER TABLE [dbo].[Sys_SsoUserGroup_MA] ADD  CONSTRAINT [DF_Sys_SsoUserGroup_MA_DcmUserGroupName]  DEFAULT ('') FOR [SsoUserGroupName]
GO
ALTER TABLE [dbo].[Sys_SsoUserGroup_MA] ADD  CONSTRAINT [DF_Sys_SsoUserGroup_MA_Status]  DEFAULT (NULL) FOR [Status]
GO
ALTER TABLE [dbo].[Sys_SsoUserGroup_MA] ADD  CONSTRAINT [DF_Sys_SsoUserGroup_MA_CreationDatetime]  DEFAULT (NULL) FOR [CreationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoUserGroup_MA] ADD  CONSTRAINT [DF_Sys_SsoUserGroup_MA_CreationUserOID]  DEFAULT (NULL) FOR [CreationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoUserGroup_MA] ADD  CONSTRAINT [DF_Sys_SsoUserGroup_MA_ModificationDatetime]  DEFAULT (NULL) FOR [ModificationDatetime]
GO
ALTER TABLE [dbo].[Sys_SsoUserGroup_MA] ADD  CONSTRAINT [DF_Sys_SsoUserGroup_MA_ModificationUserOID]  DEFAULT (NULL) FOR [ModificationUserOID]
GO
ALTER TABLE [dbo].[Sys_SsoPickDetail_TR]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoPickDetail_TR_Sys_SsoPickHeader_TR] FOREIGN KEY([PickHOID])
REFERENCES [dbo].[Sys_SsoPickHeader_TR] ([OID])
GO
ALTER TABLE [dbo].[Sys_SsoPickDetail_TR] CHECK CONSTRAINT [FK_Sys_SsoPickDetail_TR_Sys_SsoPickHeader_TR]
GO
ALTER TABLE [dbo].[Sys_SsoPickScan_TR]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoPickScan_TR_Sys_SsoPickHeader_TR] FOREIGN KEY([PickHOID])
REFERENCES [dbo].[Sys_SsoPickHeader_TR] ([OID])
GO
ALTER TABLE [dbo].[Sys_SsoPickScan_TR] CHECK CONSTRAINT [FK_Sys_SsoPickScan_TR_Sys_SsoPickHeader_TR]
GO
ALTER TABLE [dbo].[Sys_SsoPickStatus_TR]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoPickStatus_TR_Sys_SsoPickHeader_TR] FOREIGN KEY([PickHOID])
REFERENCES [dbo].[Sys_SsoPickHeader_TR] ([OID])
GO
ALTER TABLE [dbo].[Sys_SsoPickStatus_TR] CHECK CONSTRAINT [FK_Sys_SsoPickStatus_TR_Sys_SsoPickHeader_TR]
GO
ALTER TABLE [dbo].[Sys_SsoRcvStatus_TR]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoRcvStatus_TR_Sys_SsoRcvHeader_TR] FOREIGN KEY([RcvHOID])
REFERENCES [dbo].[Sys_SsoRcvHeader_TR] ([OID])
GO
ALTER TABLE [dbo].[Sys_SsoRcvStatus_TR] CHECK CONSTRAINT [FK_Sys_SsoRcvStatus_TR_Sys_SsoRcvHeader_TR]
GO
ALTER TABLE [dbo].[Sys_SsoSOScan_TR]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoSOScan_TR_Sys_SsoSOHeader_TR] FOREIGN KEY([SOHOID])
REFERENCES [dbo].[Sys_SsoSOHeader_TR] ([OID])
GO
ALTER TABLE [dbo].[Sys_SsoSOScan_TR] CHECK CONSTRAINT [FK_Sys_SsoSOScan_TR_Sys_SsoSOHeader_TR]
GO
ALTER TABLE [dbo].[Sys_SsoTransAccess_MA]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoTransAccess_MA_Sys_SsoTrAccess_MA] FOREIGN KEY([TrAccessCode])
REFERENCES [dbo].[Sys_SsoTrAccess_MA] ([TrAccessCode])
GO
ALTER TABLE [dbo].[Sys_SsoTransAccess_MA] CHECK CONSTRAINT [FK_Sys_SsoTransAccess_MA_Sys_SsoTrAccess_MA]
GO
ALTER TABLE [dbo].[Sys_SsoTransAccess_MA]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoTransAccess_MA_Sys_SsoTransName_MA] FOREIGN KEY([TransCode])
REFERENCES [dbo].[Sys_SsoTransName_MA] ([TransCode])
GO
ALTER TABLE [dbo].[Sys_SsoTransAccess_MA] CHECK CONSTRAINT [FK_Sys_SsoTransAccess_MA_Sys_SsoTransName_MA]
GO
ALTER TABLE [dbo].[Sys_SsoTransStatus_MA]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoTransStatus_MA_Sys_SsoTransName_MA] FOREIGN KEY([TransCode])
REFERENCES [dbo].[Sys_SsoTransName_MA] ([TransCode])
GO
ALTER TABLE [dbo].[Sys_SsoTransStatus_MA] CHECK CONSTRAINT [FK_Sys_SsoTransStatus_MA_Sys_SsoTransName_MA]
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoUser_MA_Sys_SsoUserGroup_MA] FOREIGN KEY([UserGroupOID])
REFERENCES [dbo].[Sys_SsoUserGroup_MA] ([OID])
GO
ALTER TABLE [dbo].[Sys_SsoUser_MA] CHECK CONSTRAINT [FK_Sys_SsoUser_MA_Sys_SsoUserGroup_MA]
GO
ALTER TABLE [dbo].[Sys_SsoUserCompany_HS]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoUserCompany_HS_Sys_SsoUser_MA] FOREIGN KEY([UserOID])
REFERENCES [dbo].[Sys_SsoUser_MA] ([OID])
GO
ALTER TABLE [dbo].[Sys_SsoUserCompany_HS] CHECK CONSTRAINT [FK_Sys_SsoUserCompany_HS_Sys_SsoUser_MA]
GO
ALTER TABLE [dbo].[Sys_SsoUserCompany_MA]  WITH CHECK ADD  CONSTRAINT [FK_Sys_SsoUserCompany_MA_Sys_SsoUser_MA] FOREIGN KEY([UserOID])
REFERENCES [dbo].[Sys_SsoUser_MA] ([OID])
GO
ALTER TABLE [dbo].[Sys_SsoUserCompany_MA] CHECK CONSTRAINT [FK_Sys_SsoUserCompany_MA_Sys_SsoUser_MA]
GO
USE [master]
GO
ALTER DATABASE [SB_IM] SET  READ_WRITE 
GO
