<%
'####################################################
' Description : 할인관리
' History : 2010.09.28 한용민 생성
'####################################################

Class CSale
	public FSCode
	public FECode
	public FTotCnt
	public FCPage
	public FPSize	
	public FSearchTxt
	public FSearchType
	public FBrand		
	public FDateType   
	public FSDate		
	public FEDate		
	public FSStatus
	public FSName 		
	public FSRate 		
	public FSMargin 		
	public FEGroupCode	 	
	public FSRegdate 
	public FSUsing 	
	public FSAdminid 
	public FOpenDate
	public FSMarginValue
	public FCloseDate

	'== 할인관리 리스트 가져오기 '//academy/sale/salelist.asp
	public Function fnGetSaleList
		Dim strSqlCnt, strSql, strSearch,iDelCnt
		
		strSearch = ""
		IF FECode <> "" THEN
			strSearch = " and evt_code ="&FECode
		END IF	
		
		IF FSearchTxt <> "" THEN
			IF FSearchType = 1 THEN 
				strSearch = strSearch & " and sale_code = "&FSearchTxt
			ELSEIF FSearchType= 2 THEN
				strSearch = strSearch & " and evt_code = "&FSearchTxt	
			ELSEIF FSearchType=3 THEN
				strSearch = strSearch & " and sale_name like '%"& FSearchTxt &"%' "
			END IF	
		END IF					
		
				
		IF FSDate <> "" AND FEDate <> "" THEN
			if CStr(FDateType) = "S" THEN
				strSearch  = strSearch & " and  datediff(day, '"&FSDate&"', sale_startdate) >= 0 and  datediff(day,'"&FEDate&"', sale_startdate) <=0  "
			elseif CStr(FDateType) = "E" THEN
				strSearch  = strSearch & " and  datediff(day,'"&FSDate&"',sale_enddate) >= 0 and  datediff(day,'"&FEDate&"',sale_enddate) <=0  "
			end if
		END IF
		
		IF FSStatus <> "" THEN			
			strSearch = strSearch & " and sale_status = "&FSStatus
		END IF	
	
		strSqlCnt = " SELECT COUNT(sale_code) FROM [db_academy].[dbo].[tbl_sale] WHERE sale_using =1 "	&strSearch		
		
		rsacademyget.Open strSqlCnt,dbacademyget 
		IF not rsacademyget.EOF THEN
			FTotCnt = rsacademyget(0)
		End IF
		rsacademyget.Close		
		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = " SELECT TOP "&FPSize&"  [sale_code], [sale_name], [sale_rate], [sale_margin], [evt_code], [evtgroup_code]"&_
					", [sale_startdate], [sale_enddate], [sale_status], [availPayType], [regdate], [sale_using], [adminid] "&_		
					", (select count(itemid) from [db_academy].[dbo].tbl_saleItem  where sale_code = A.sale_code ) as saleitem_cnt "&_	
					", sale_marginvalue, opendate, closedate "&_		
					" FROM [db_academy].[dbo].[tbl_sale] as A "&_
					" WHERE sale_using =1 AND sale_code <= ( SELECT Min(sale_code) FROM ( SELECT TOP "&iDelCnt&" sale_code "&_
					" FROM [db_academy].[dbo].[tbl_sale] WHERE sale_using =1 "&strSearch&" ORDER BY sale_code DESC ) as T ) "&strSearch&" ORDER BY sale_code DESC "																													
			rsacademyget.Open strSql,dbacademyget,3 			
			IF not rsacademyget.EOF THEN
				fnGetSaleList = rsacademyget.getRows()
			End IF
			rsacademyget.Close
			
		END IF					
	End Function

	'== 할인관리 내용가져오기 '/academy/sale/saleReg.asp
	public Function fnGetSaleConts
		Dim strSql
		
		strSql = " SELECT [sale_code], [sale_name], [sale_rate], [sale_margin], [evt_code], [evtgroup_code] "&_
				", [sale_startdate], [sale_enddate], [sale_status], [availPayType], [regdate], [sale_using] "&_
				" , [adminid], [opendate], [closedate],sale_marginvalue "&_		
				" FROM [db_academy].[dbo].[tbl_sale] "&_ 
				" WHERE sale_code = "&FSCode
		
		'response.write strSql &"<br>"		
		rsacademyget.Open strSql,dbacademyget 
			
			IF not rsacademyget.EOF THEN
				FSCode 		= rsacademyget("sale_code")
				FSName 		= rsacademyget("sale_name")
				FSRate 		= rsacademyget("sale_rate")
				FSMargin 	= rsacademyget("sale_margin")
				FECode 		= rsacademyget("evt_code")
				FEGroupCode = rsacademyget("evtgroup_code")
				FSDate 		= rsacademyget("sale_startdate")
				FEDate		= rsacademyget("sale_enddate")
				FSStatus 	= rsacademyget("sale_status")
				FSRegdate 	= rsacademyget("regdate")
				FSUsing 	= rsacademyget("sale_using")
				FSAdminid 	= rsacademyget("adminid")
				FOpenDate	= rsacademyget("opendate")
				FCloseDate	= rsacademyget("closedate")
				FSMarginValue	= rsacademyget("sale_marginvalue")				
			END IF
		rsacademyget.close
			
	End Function	
END Class	

'할인상품 
Class CSaleItem
	public FSCode	
	public FTotCnt
	public FCPage
	public FPSize
	public FItemid
	public FBrand
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small		
	public FChkPreSale
	
	'--특정 할인마스터에 해당하는 상품 리스트 가져오기 '/academy/sale/saleItemReg.asp
	public Function fnGetSaleItemList
		Dim strSqlCnt, strSql
		
		strSqlCnt = " SELECT COUNT(i.itemid) FROM [db_academy].[dbo].[tbl_saleItem] A, [db_academy].dbo.tbl_diy_item i"&_
				"	WHERE A.itemid = i.itemid AND A.sale_code = "&FSCode
		
		'response.write strSql &"<br>"
		rsacademyget.Open strSqlCnt,dbacademyget 
		IF not rsacademyget.EOF THEN
			FTotCnt = rsacademyget(0)
		End IF
		rsacademyget.Close
		
		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = " SELECT TOP "&FPSize&" A.sale_code, A.itemid, A.saleprice, A.salesupplycash, A.saleItem_status, A.limitno, A.orglimityn "&_
					"	  ,i.makerid, i.itemname, i.smallimage ,i.saleyn,i.sellcash, i.buycash,i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash"&_ 
					"	 , i.mwdiv,i.limityn,i.limitno, i.limitsold,i.isusing "&_  
					"	FROM [db_academy].[dbo].[tbl_saleItem] A, [db_academy].dbo.tbl_diy_item i "&_
					"	WHERE A.itemid = i.itemid AND A.sale_code = "&FSCode&"   AND A.saleitem_idx <="&_
					"			(select min(saleitem_idx) from (select top "&iDelCnt&" saleitem_idx "&_
					"			 from [db_academy].[dbo].[tbl_saleItem] b, [db_academy].dbo.tbl_diy_item d "&_
					"			 where b.itemid = d.itemid and b.sale_code ="&FSCode&_
					"			 order by b.saleitem_idx desc) as T ) order by A.saleitem_idx desc "					
			
			'response.write strSql &"<br>"
			rsacademyget.Open strSql,dbacademyget 
			IF not rsacademyget.EOF THEN
				fnGetSaleItemList = rsacademyget.getRows()
			End IF
			rsacademyget.Close
		END IF				
				
	End Function				

End Class
%>	