USE [BrickLinkCache]
GO

/****** Object:  Table [dbo].[Sets]    Script Date: 8/27/2023 1:02:05 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Sets](
	[ID] [nvarchar](max) NOT NULL,
	[name] [nvarchar](max) NULL,
	[type] [nvarchar](max) NULL,
	[categoryID] [nvarchar](max) NULL,
	[imageURL] [nvarchar](max) NULL,
	[thumbnail_url] [nvarchar](max) NULL,
	[year_released] [nvarchar](max) NULL,
	[avg_price] [nvarchar](max) NULL,
	[date_updated] [datetime] NULL,
	[partnum] [int] NULL,
	[minifignum] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


