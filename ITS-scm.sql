/*
query: ITS
modified: 27.08.2020
*/

Set NOCOUNT ON;

DECLARE @db nvarchar(100), @sdate smalldatetime, @edate smalldatetime, 
        @ExchDate smalldatetime, @ExchRate float, @Country nvarchar(10), 
	    @SysCurrncy nvarchar(5), @MainCurncy nvarchar(5), @DirectRate nvarchar(3)

DECLARE @sql_stmt NVARCHAR(3000), @Parmdef nvarchar (1000)

set @db = '$(DB)'
set @edate = '$(eDate)'
set @sdate = DATEADD(MONTH, -6, @edate+1)
set @ExchDate = getdate()

if @ExchDate > @edate
  set @ExchDate = @edate

CREATE TABLE #itemrecs
(
 ItemCode nvarchar(50),
 cb_qty float,
 cb_cnsg float,
 cb_val money,
 issueqty numeric(16,3),
 psiteqty numeric(16,3),
 icnsgqty numeric(16,3),
 dcnsgqty numeric(16,3)
)

create table #ITS
(
Country nvarchar(5),
ItemCode nvarchar(50),
ItemName nvarchar(100),
SAP_MatReference nvarchar(50),
ClosingQty float,
CQty_Cnsg float,
UnitMsr nvarchar(10),
ClosingVal money,
ExchangeRate float,
ClosingValEUR money,
ConsQtyLast6Mths float,
psiteqty float,
ConsCnsgQtyLast6Mths float,
dcnsgqty float,
CardCode nvarchar(20), 
CardName nvarchar(250),
MHMSalNo nvarchar(50), 
PkgCode nvarchar(50), 
ItemType nvarchar(50), 
TreeType nvarchar(5),
Substitute nvarchar(50),
ProdHierarchy nvarchar(50),
SPPROPERTY nvarchar(50),
SPPName nvarchar(200),
ProductGrpCode nvarchar(20)
)

select @Country = 'UK ', @SysCurrncy = 'EUR', @ExchRate = 1, @DirectRate = 'N'
SET @parmdef = N'@Country nvarchar(10) OUTPUT, @SysCurrncy nvarchar(5) OUTPUT, @MainCurncy nvarchar(5) OUTPUT, @DirectRate nvarchar(3) OUTPUT'
SET @sql_stmt = 'select top 1 @Country = Country, @SysCurrncy = SysCurrncy, 
                              @MainCurncy = MainCurncy, @DirectRate = DirectRate 
                   from '+@db+'.dbo.OADM'
EXEC sp_executeSQL @sql_stmt, @Parmdef, @Country OUTPUT, @SysCurrncy OUTPUT, @MainCurncy OUTPUT, @DirectRate OUTPUT

if @SysCurrncy <> 'EUR' --and @SysCurrncy = @MainCurncy
  set @SysCurrncy = 'EUR' --to take of non-standard sites
  
if @SysCurrncy <> @MainCurncy
begin
  SET @parmdef = N'@ExchRate numeric(16,8) OUTPUT, @SysCurrncy nvarchar(5), @ExchDate smalldatetime'
  SET @sql_stmt = 'select top 1 @ExchRate = Rate
                   from '+@db+'.dbo.ORTT 
				   where Currency = @SysCurrncy
				     and RateDate <= @ExchDate
				   order by RateDate DESC'
  EXEC sp_executeSQL @sql_stmt, @Parmdef, @ExchRate OUTPUT, @SysCurrncy, @ExchDate
end
  
if @DirectRate = 'Y'
  set @ExchRate = 1 / @ExchRate
  
truncate table #itemrecs

SET @parmdef = N'@sdate smalldatetime, @edate smalldatetime'
if @Country = 'GB'
SET @sql_stmt = 'insert #itemrecs
                 SELECT T0.ItemCode, sum(T0.InQty-T0.OutQty), 
				        sum(case when isnull(T1.[U_WHtype],''S'') in (''C'',''N'') then T0.InQty-T0.OutQty else 0 end), 
				        sum(t0.TransValue),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') not in (''C'',''N''))
						          and (T0.[TransType] in (13,1499,6099,165,1669) 
								   OR (T0.[TransType] = 5999 and T0.[ApplObj]<>202)
				   				   OR (T0.[TransType] = 15 and isnull(DN.[U_OrderTyp],'''') not in (''1005'',''1006''))
								   OR (T0.[TransType] = 1699 and isnull(RD.[U_OrderTyp],'''') not in (''1007'')))
						           THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') not in (''C'',''N''))
						          and (T0.[TransType] = 60 and T0.[ApplObj] = 202) 
								  THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and isnull(T1.[U_WHtype],''S'') in (''C'',''N'')
						          and (T0.[TransType] in (13,1499,15,6099,165,1669) 
								   OR (T0.[TransType] = 5999 and T0.[ApplObj]<>202)
								   OR (T0.[TransType] = 1699 and isnull(RD.[U_OrderTyp],'''') not in (''1007'')))
						           THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') not in (''C'',''N''))
						          and (T0.[TransType] in (13,15) and isnull(DN.[U_OrderTyp],'''') in (''1005'',''1006'')) 
								  THEN T0.OutQty-T0.InQty ELSE 0 END)
                   FROM '+@db+'.dbo.OINM T0
				     left outer join '+@db+'.dbo.OWHS T1 on T1.[WhsCode] = T0.[Warehouse]
					 left outer join '+@db+'.dbo.ODLN DN on T0.[TransType] = 15 and DN.[DocEntry] = T0.[CreatedBy]
					 left outer join '+@db+'.dbo.ORDN RD on T0.[TransType] = 16 and RD.[DocEntry] = T0.[CreatedBy]
                   where T0.DocDate <= @edate
				   group by T0.ItemCode'
else if @Country = 'FR'
SET @sql_stmt = 'insert #itemrecs
                 SELECT T0.ItemCode, sum(T0.InQty-T0.OutQty), 
				        sum(case when isnull(T1.[U_WHtype],''S'') in (''C'',''N'') then T0.InQty-T0.OutQty else 0 end), 
				        sum(t0.TransValue), 
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') not in (''C'',''N''))
						          and (T0.[TransType] in (13,1499,6099,165,1669) 
								   OR (T0.[TransType] = 5999 and T0.[ApplObj]<>202)
				   				   OR (T0.[TransType] = 15 and isnull(DN.[U_OrderTyp],'''') not in (''XX'',''XX-Retour''))
								   OR (T0.[TransType] = 1699 and isnull(RD.[U_OrderTyp],'''') not in (''XX'',''XX-Retour'')))
						           THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') not in (''C'',''N''))
						          and (T0.[TransType] = 60 and T0.[ApplObj] = 202) 
								  THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and isnull(T1.[U_WHtype],''S'') in (''C'',''N'')
						          and (T0.[TransType] in (13,1499,15,6099,165,1669) 
								   OR (T0.[TransType] = 5999 and T0.[ApplObj]<>202)
								   OR (T0.[TransType] = 1699 and isnull(RD.[U_OrderTyp],'''') not in (''XX'',''XX-Retour'')))
						           THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') not in (''C'',''N''))
						          and (T0.[TransType] in (13,15) and isnull(DN.[U_OrderTyp],'''') in (''XX'')) 
								  THEN T0.OutQty-T0.InQty ELSE 0 END)
                   FROM '+@db+'.dbo.OINM T0
				     left outer join '+@db+'.dbo.OWHS T1 on T1.[WhsCode] = T0.[Warehouse]
					 left outer join '+@db+'.dbo.ODLN DN on T0.[TransType] = 15 and DN.[DocEntry] = T0.[CreatedBy]
					 left outer join '+@db+'.dbo.ORDN RD on T0.[TransType] = 16 and RD.[DocEntry] = T0.[CreatedBy]
                   where T0.DocDate <= @edate
				   group by T0.ItemCode'
else
SET @sql_stmt = 'insert #itemrecs
                 SELECT T0.ItemCode, sum(T0.InQty-T0.OutQty), 
				        sum(case when isnull(T1.[U_WHtype],''S'') in (''C'',''N'') then T0.InQty-T0.OutQty else 0 end), 
				        sum(t0.TransValue), 
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') not in (''C'',''N''))
						          and (T0.[TransType] in (13,1499,15,1699,6099,165,1669) 
								   OR (T0.[TransType] = 5999 and T0.[ApplObj]<>202)) 
								  THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') not in (''C'',''N''))
						          and (T0.[TransType] = 60 and T0.[ApplObj] = 202) 
								  THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						          and (isnull(T1.[U_WHtype],''S'') in (''C'',''N''))
						          and (T0.[TransType] in (13,1499,15,1699,6099,165,1669) 
								   OR (T0.[TransType] = 5999 and T0.[ApplObj]<>202))
						         THEN T0.OutQty-T0.InQty ELSE 0 END),
                        sum(CASE WHEN T0.DocDate between @sdate and @edate 
						           and (isnull(T1.[U_WHtype],''S'') in (''C'',''N'')) and T0.[TransType] = 67
						         THEN T0.InQty-T0.OutQty ELSE 0 END)
                   FROM '+@db+'.dbo.OINM T0
				     left outer join '+@db+'.dbo.OWHS T1 on T1.[WhsCode] = T0.[Warehouse]
                   where T0.DocDate <= @edate
				   group by T0.ItemCode'
EXEC sp_executeSQL @sql_stmt, @Parmdef, @sdate, @edate

SET @parmdef = N'@Country nvarchar(10), @ExchRate numeric(16,8)'
SET @sql_stmt = 'Insert #ITS
                 select @Country, T0.ItemCode, T1.ItemName, 
                        case when isnull(T1.U_M_SAP_MatReference,'''') = '''' then ''''
	                         when isnull(t1.U_M_Gebinde,'''') = '''' then isnull(T1.U_M_SAP_MatReference,'''')
	                         else isnull(T1.U_M_SAP_MatReference,'''') + ''-'' + t1.U_M_Gebinde
					    end, T0.cb_qty, T0.cb_cnsg, T1.InvntryUom, T0.cb_val, 
						convert(varchar, @ExchRate), T0.cb_val*@ExchRate, 
						T0.issueqty,  T0.psiteqty,  T0.icnsgqty, case when T0.dcnsgqty < 0 then 0 else T0.dcnsgqty end, 
	                    T1.CardCode, T2.CardName, T1.U_M_MHMSalNo, T1.U_M_Gebinde, T1.U_M_Item_Type,
	                   (case when isnull(B1.[TreeType], '' '') = ''P'' then ''Yes'' else ''No'' end),
					   S0.[Substitute], T1.[U_M_ProdHierarchy], T1.U_SPPROPERTY, S1.[U_LName], 
					   isnull(T1.[U_M_ProductGrpCode],S1.[U_M_ProductGrpCode])
                   from #itemrecs T0
                     inner join '+@db+'.dbo.OITM T1 on T1.itemcode = T0.ItemCode collate SQL_Latin1_General_CP850_CI_AS
	                 left outer join '+@db+'.dbo.OITT B1 on B1.[Code] = T1.ItemCode
                     left outer join '+@db+'.dbo.OCRD T2 on T2.CardCode = T1.CardCode
					 left outer join (select t.ItemCode, t.CardCode, max(t.Substitute) as Substitute from '+@db+'.dbo.OSCN t group by t.ItemCode, t.CardCode) S0 on S0.[ItemCode] = T1.ItemCode and S0.[CardCode] = T1.CardCode
					 left outer join '+@db+'.dbo.[@SPGROUP] S1 on S1.[Code] = T1.U_SPPROPERTY
                   where T0.cb_qty<>0 or T0.issueqty<>0 or T0.psiteqty<>0 or T0.icnsgqty<>0 or T0.dcnsgqty<>0'
EXEC sp_executeSQL @sql_stmt, @Parmdef, @Country, @ExchRate

select isnull(T0.Country,'') [Country], 
       isnull(T0.ItemCode,'') [ItemCode], 
       --isnull(T0.ItemName,'') [ItemName],
       replace(replace(replace(isnull(T0.ItemName,''),char(9),' '),char(10),' '),char(13),' ') [ItemName],
       isnull(T0.SAP_MatReference,'') [SAP_MatReference], 
	   isnull(T0.MHMSalNo,'') [SAP_EU_Code],
	   isnull(str(T0.[ClosingQty],16,3),'') [quantityTotal], 
	   isnull(str(T0.[ClosingQty]-T0.[CQty_Cnsg],16,3),'') [quantityAtSite], 
	   isnull(str(T0.[CQty_Cnsg],16,3),'') [quantityCons], 
	   isnull(T0.UnitMsr,'') [UnitOfMeasure], 
	   str(T0.[ClosingVal],16,2) [LocalVAL], 
	   str(T0.[ClosingValEUR],16,2) [ValEUR], 
	   str(T0.[ClosingVal]-(case when T0.[ClosingQty]>0 then T0.[CQty_Cnsg]*T0.[ClosingVal]/T0.[ClosingQty] else 0 end),16,2) [LocalValSite], 
	   str(T0.[ClosingValEUR]-(case when T0.[ClosingQty]>0 then T0.[CQty_Cnsg]*T0.[ClosingValEUR]/T0.[ClosingQty] else 0 end),16,2) [EURValSite], 
	   --str(case when T0.[ClosingQty]>0 then T0.[ClosingValEUR]/T0.[ClosingQty] else 0 end,16,2) [UnitpriceEUR], 
	   str(case when T0.[ClosingQty]>0 then T0.[CQty_Cnsg]*T0.[ClosingVal]/T0.[ClosingQty] else 0 end,16,2) [LocalValCons], 
	   str(case when T0.[ClosingQty]>0 then T0.[CQty_Cnsg]*T0.[ClosingValEUR]/T0.[ClosingQty] else 0 end,16,2) [EURValCons], 
	   str(T0.[ExchangeRate],16,8) [ExchangeRate],
	   str(case when T0.[ClosingQty]>0 then T0.[ClosingVal]/T0.[ClosingQty] else 0 end,16,2) [Unitprice localVAL], 
	   str(T0.[ConsQtyLast6Mths]+T0.[psiteqty]+T0.[ConsCnsgQtyLast6Mths],16,3) [ConsTotal],
	   str(T0.[ConsQtyLast6Mths],16,3) [ConsLocalSiteSales],
	   str(T0.[psiteqty],16,3) [ConsLocalSiteProd],
	   str(T0.[ConsCnsgQtyLast6Mths],16,3) [ConsToCustomer],
	   --str(T0.[dcnsgQty],16,3) [ConsDeliveryToCnsgStock],
	   T0.CardCode [Card_Code], 
	   --T0.CardName [Card_Name], 
       replace(replace(replace(isnull(T0.CardName,''),char(9),' '),char(10),' '),char(13),' ') [Card_Name],
	   T0.MHMSalNo [U_M_MHMSalNo],
       T0.PkgCode [U_M_Gebinde],
	   T0.ItemType [U_M_Item_Type], 
	   T0.Substitute [Substitute],
	   T0.ProdHierarchy [U_M_ProdHierarchy],
	   T0.SPPROPERTY [U_SPPROPERTY],
	   T0.SPPName [U_LName],
	   T0.ProductGrpCode [ProductGrpCode],
	   T0.[TreeType] [TreeType]
  from #ITS T0
  order by T0.[Country], T0.[ItemCode]

drop table #itemrecs, #ITS