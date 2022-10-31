USE [BrickLinkCache]
GO

/****** Object:  Table [dbo].[Sets]    Script Date: 10/31/2022 8:13:45 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Sets](
	[ID] [nchar](10) NOT NULL,
	[name] [nvarchar](max) NULL,
	[type] [nvarchar](50) NULL,
	[categoryID] [nchar](10) NULL,
	[imageURL] [nvarchar](max) NULL,
	[thumbnail_url] [nvarchar](max) NULL,
	[year_released] [nchar](10) NULL,
	[avg_price] [nchar](10) NULL,
	[date_updated] [datetime] NULL,
PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


