USE [BrickLinkCache]
GO

/****** Object:  Table [dbo].[Sets]    Script Date: 11/12/2023 7:08:43 PM ******/
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
	[minifignum] [int] NULL,
	[UPC] [nvarchar](max) NULL,
	[description] [nvarchar](max) NULL,
	[original_price] [nvarchar](max) NULL,
	[minifigset] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


