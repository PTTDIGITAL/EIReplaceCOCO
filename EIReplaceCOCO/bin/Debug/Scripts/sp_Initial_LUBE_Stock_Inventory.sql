USE [POSDB]
GO

/****** Object:  StoredProcedure [dbo].[sp_Initial_LUBE_Stock_Inventory]    Script Date: 1/9/2560 12:09:08 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO







---GO

---GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[sp_Initial_LUBE_Stock_Inventory]

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
	SET QUOTED_IDENTIFIER ON
	SET noexec off;


		begin transaction;
		--GO

		--Cancel Intial History
		UPDATE Document

		SET DocumentStatus = 99
		, CancelDate = GETDATE()

		WHERE CreateBy = 'ADMIN INITAIL STOCK'

---------------------------------------------------------------------------------------------------------------------------------------

		--Document ID
		declare @DocID  int;
		SET @DocID = (SELECT ISNULL((SELECT ISNULL(MAX(DocumentID), 0) + 1  FROM Document), 1) ) ;

		--Shop ID
		declare @ShopD  int;
		SET @ShopD = (SELECT TOP 1 (ShopID) FROM shop_data WHERE Deleted = 0) ;

		--DocumentNumber
		declare @DocumentNumber  int;
		SET @DocumentNumber = (SELECT ISNULL((SELECT ISNULL(MAX(DocumentNumber), 0) + 1  FROM Document WHERE DocumentTypeID = 102 AND DocumentYear = YEAR(GETDATE()) AND DocumentMonth = MONTH(GETDATE())), 1)) ;

---------------------------------------------------------------------------------------------------------------------------------------

		if @@error != 0 set noexec on;
		
		-- Insert Intial Stock Header
		INSERT INTO Document([DocumentID], [ShopID], [DocumentTypeID], [DocumentYear], [DocumentMonth], [DocumentNumber]
		         , [DocumentStatus], [DocumentDate], [Remark], [ProductLevelID], [ToInvID],[InputBy]
				 , [UpdateBy]  , [ApproveBy] ,[InsertDate],[UpdateDate] ,[ApproveDate],  [CreateBy],[ShiftDay], [ShiftNo] )

		SELECT  @DocID
		      , @ShopD
			  , 102
			  , YEAR(GETDATE())
			  , MONTH(GETDATE())
			  , @DocumentNumber
			  , 2
			  , CONVERT(date, GETDATE())  
			  , 'เริ่มต้นสต๊อก ' +  CONVERT(varchar, CONVERT(date, GETDATE()) )
			  , @ShopD
			  , @ShopD
			  , 1
			  , 1
			  , 1
			  , GETDATE()
			  , GETDATE()
			  , GETDATE()
			  , 'ADMIN INITAIL STOCK'
			  , 1
			  , 1

---------------------------------------------------------------------------------------------------------------------------------------
		-- Insert Intial Stock Detail
		if @@error != 0 set noexec on;

		INSERT INTO [dbo].[DocDetail]
           ([DocDetailID]
           ,[DocumentID]
           ,[ShopID]
           ,[ProductID]
           ,[ProductCode]
           ,[ProductName]
           ,[SupplierMaterialCode]
           ,[SupplierMaterialName]
           ,[UnitID]
           ,[ProductUnit]
           ,[UnitName]
           ,[UnitSmallAmount]
           ,[ProductAmount]
		   ,[ProductTaxType])

		   SELECT RANK() OVER (ORDER BY M.MAT_ID) 
		   , @DocID
		   , @ShopD
		   , P.ProductID
		   , M.MAT_ID
		   , M.MAT_NAME
		   , P.ProductID
		   , M.MAT_ID
		   , P.ProductId
		   , P.ProductId
		   , P.ProductUnitName
		   , M.STOCK
		   , M.STOCK
		   , 1

		   FROM TBMATERIAL M 
		   INNER JOIN Products P ON M.MAT_ID = P.ProductCode AND P.Deleted = 0

		   WHERE M.STOCK >= 0
		   AND (M.BLOCK <> 'X' or M.BLOCK is null)
		   AND M.MAT_ID2 IS NULL

---------------------------------------------------------------------------------------------------------------------------------------

		declare @finished bit;
		set @finished = 1;

		SET noexec off;

		IF @finished = 1
		BEGIN
			PRINT 'Committing changes'
			COMMIT TRANSACTION
	
		END
		ELSE
		BEGIN
			PRINT 'Errors occured. Rolling back changes'
			ROLLBACK TRANSACTION
		END

END






GO


