<%
'####################################################
' Description : 할인관리
' History : 2008.04.01
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
	public FSType

	'== 할인관리 리스트 가져오기
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

	    IF FSType <> "" THEN
	        strSearch = strSearch & " and sale_type = "&FSType
	    END IF

		strSqlCnt = " SELECT COUNT(sale_code) FROM [db_event].[dbo].[tbl_sale] with (nolock) WHERE sale_using =1 "	&strSearch

		'response.write strSqlCnt & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSqlCnt, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = " SELECT TOP "&FPSize&"  [sale_code], [sale_name], [sale_rate], [sale_margin], [evt_code], [evtgroup_code]"&_
					", [sale_startdate], [sale_enddate], [sale_status], [availPayType], [regdate], [sale_using], [adminid] "&_
					", (select count(itemid) from [db_event].[dbo].tbl_saleItem  where sale_code = A.sale_code ) as saleitem_cnt "&_
					", sale_marginvalue, opendate, closedate, sale_type  "&_
					" FROM [db_event].[dbo].[tbl_sale] as A with (nolock)"&_
					" WHERE sale_using =1 AND sale_code <= ( SELECT Min(sale_code) FROM ( SELECT TOP "&iDelCnt&" sale_code "&_
					" FROM [db_event].[dbo].[tbl_sale] with (nolock) WHERE sale_using =1 "&strSearch&" ORDER BY sale_code DESC ) as T ) "&strSearch&" ORDER BY sale_code DESC "

			'response.write strSql & "<Br>"
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				fnGetSaleList = rsget.getRows()
			End IF
			rsget.Close

		END IF
	End Function

	'== 할인관리 내용가져오기
	public Function fnGetSaleConts
		Dim strSql
		strSql = " SELECT [sale_code], [sale_name], [sale_rate], [sale_margin], [evt_code], [evtgroup_code] "&_
				", convert(varchar(19),sale_startdate,121) as [sale_startdate], convert(varchar(19),sale_enddate,121) as [sale_enddate], [sale_status], [availPayType], [regdate], [sale_using] "&_
				" , [adminid], convert(varchar(19),[opendate],121) as  opendate, [closedate],sale_marginvalue , sale_type "&_
				" FROM [db_event].[dbo].[tbl_sale] "&_
				" WHERE sale_code = "&FSCode
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				FSCode 		= rsget("sale_code")
				FSName 		= rsget("sale_name")
				FSRate 		= rsget("sale_rate")
				FSMargin 	= rsget("sale_margin")
				FECode 		= rsget("evt_code")
				FEGroupCode = rsget("evtgroup_code")
				FSDate 		= rsget("sale_startdate")
				FEDate		= rsget("sale_enddate")
				FSStatus 	= rsget("sale_status")
				FSRegdate 	= rsget("regdate")
				FSUsing 	= rsget("sale_using")
				FSAdminid 	= rsget("adminid")
				FOpenDate	= rsget("opendate")
				FCloseDate	= rsget("closedate")
				FSMarginValue	= rsget("sale_marginvalue")
				FSType      = rsget("sale_type")
			END IF
		rsget.close
	End Function
END Class

'할인상품
Class CSaleItem
	public FSCode
	public FTotCnt
	public FCPage
	public FPSize
	public FSPageNo
	public FEPageNo
	public FItemid
	public FBrand
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
    public FRectDispCate

	public FRectItemSaleStatus
    public FRectSaleStatus
	public FRectMakerid
	public FRectsailyn
	public FRectinvalidmargin
	public FRectItemidArr


	'--특정 할인마스터에 해당하는 상품 리스트 가져오기
	public Function fnGetSaleItemList
		Dim strSqlCnt, strSql, strSqlAdd
		dim iDelCnt

		strSqlAdd = ""
		if FRectMakerid <> "" then
		strSqlAdd = strSqlAdd & " and i.makerid = '"&FRectMakerid&"'"
		end if

		if FRectSailYn<>"" then
			strSqlAdd = strSqlAdd & " and i.sailyn='"&FRectSailYn&"'"
		end if

        if FRectInvalidMargin="Y" then
            strSqlAdd = strSqlAdd & " and   (( ( B.saleprice-B.salesupplycash)/B.saleprice*100)  < 0  "
            strSqlAdd = strSqlAdd & "       or"
            strSqlAdd = strSqlAdd & "  (B.salesupplycash>i.orgsuplycash))" ''2018/07/20
        end if

		if FRectItemidArr <> "" then
		strSqlAdd = strSqlAdd & " and B.itemid in ("&FRectItemidArr&")"
		end if

		 '//6-오픈, 7-오픈요청, 8-종료,9-종료요청
        if (FRectSaleStatus <> "") then
        	strSqlAdd = strSqlAdd & " and A.sale_status =  "&FRectSaleStatus
        end if

        if (FRectItemSaleStatus <> "") then
        	strSqlAdd = strSqlAdd & " and B.saleitem_status =  "&FRectItemSaleStatus
        end if

		strSqlCnt = " SELECT COUNT(i.itemid) " &VbCrlf
		strSqlCnt = strSqlCnt & " FROM  [db_item].[dbo].[tbl_item] as i" &VbCrlf
		strSqlCnt = strSqlCnt & "		inner join [db_event].[dbo].[tbl_saleItem] as B on B.itemid = i.itemid " &VbCrlf
		strSqlCnt = strSqlCnt & "		inner join [db_event].[dbo].[tbl_sale] as A on A.sale_code = B.sale_code " &VbCrlf
		strSqlCnt = strSqlCnt & "		 	WHERE   A.sale_code = "&FSCode & strSqlAdd
		''rw strSqlCnt

		rsget.Open strSqlCnt,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
				dim iSPageNo, iEPageNo
				iSPageNo = (FPSize*(FCPage-1)) + 1
				iEPageNo = FPSize*FCPage

			strSql = " SELECT   sale_code,  itemid,  saleprice,  salesupplycash,  saleItem_status,  slimitno,  orglimityn " &VbCrlf
			strSql =	strSql & "	  , makerid,  itemname,  smallimage , sailyn, sellcash,  buycash, orgprice,  orgsuplycash,  sailprice,  sailsuplycash" &VbCrlf
			strSql =	strSql & "	 ,  mwdiv, limityn, limitno,  limitsold, isusing ,  orgsailprice,  orgsailsuplycash,  orgsailyn, optioncnt"   &VbCrlf
			strSql =	strSql & "	FROM " 	&VbCrlf
			strSql =  strSql & " 		( SELECT   ROW_NUMBER() OVER (ORDER BY  saleitem_idx desc ) as RowNum , A.sale_code, B.itemid, B.saleprice, B.salesupplycash, B.saleItem_status, B.limitno as slimitno, B.orglimityn " &VbCrlf
			strSql =  strSql &		"	  ,i.makerid, i.itemname, i.smallimage ,i.sailyn,i.sellcash, i.buycash,i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash" &VbCrlf
			strSql =  strSql &		"	 , i.mwdiv,i.limityn,i.limitno, i.limitsold,i.isusing , B.orgsailprice, B.orgsailsuplycash, B.orgsailyn , i.optioncnt " &VbCrlf
			strSql =  strSql & "		 FROM  [db_item].[dbo].[tbl_item] as i" &VbCrlf
			strSql =  strSql & "		inner join [db_event].[dbo].[tbl_saleItem] as B on B.itemid = i.itemid " &VbCrlf
			strSql =  strSql & "		inner join [db_event].[dbo].[tbl_sale] as A on A.sale_code = B.sale_code " &VbCrlf
			strSql =  strSql & "		 	WHERE   A.sale_code = "&FSCode & strSqlAdd   &VbCrlf
			strSql =  strSql &		") AS TB "&VbCrlf
			strSql =  strSql &		" WHERE TB.RowNum Between "&iSPageNo&" AND "  &iEPageNo & " "&VbCrlf
			strSql =  strSql &		" order by  TB.RowNum  "
			''rw strSql

			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetSaleItemList = rsget.getRows()
			End IF
			rsget.Close
		END IF

	End Function


	'// 진행중인 쿠폰 목록 출력
	Function fnGetCouponListBySaleInfo()
		Dim strSql
		strSql = "Select cm.itemcouponidx, cm.itemcoupontype, cm.itemcouponvalue, cm.itemcouponname, cd.itemid, cd.couponbuyprice, cm.itemcouponstartdate, cm.itemcouponexpiredate "
		strSql = strSql & "from db_item.dbo.tbl_item_coupon_master as cm "
		strSql = strSql & "	join db_item.dbo.tbl_item_coupon_detail as cd "
		strSql = strSql & "		on cm.itemcouponidx=cd.itemcouponidx "
		strSql = strSql & "	join db_event.dbo.tbl_saleItem as sd "
		strSql = strSql & "		on cd.itemid=sd.itemid "
		strSql = strSql & "	join db_event.dbo.tbl_sale as sm "
		strSql = strSql & "		on sm.sale_code=sd.sale_code "
		strSql = strSql & "			and sm.sale_using=1 "
		strSql = strSql & "where cm.itemcouponstartdate<=sm.sale_enddate "
		strSql = strSql & "	and cm.itemcouponexpiredate>=sm.sale_startdate "
		strSql = strSql & "	and cm.itemcoupontype in ('1','2') "
		strSql = strSql & "	and cm.openstate<9 "
		strSql = strSql & "	and sm.sale_code=" & FSCode
		strSql = strSql & "order by cd.itemid, cm.itemcouponidx "

		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetCouponListBySaleInfo = rsget.getRows()
		End IF
		rsget.Close
	end Function


	'-- 할인 대기,오픈중인 상품리스트
	public Function fnGetSaleOnItemList
		Dim strSqlCnt, addSql, strSql


		if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if

        if (FBrand <> "") then
            addSql = addSql & " and i.makerid='" + FBrand + "'"
        end if

        if (FItemid <> "") then
            addSql = addSql & " and i.itemid in (" + FItemid + ")"
        end if
        '//6-오픈, 7-오픈요청, 8-종료,9-종료요청
        if (FRectSaleStatus <> "") then
        	addSql = addSql & " and A.sale_status =  "&FRectSaleStatus
        end if

        if (FRectItemSaleStatus <> "") then
        	addSql = addSql & " and B.saleitem_status =  "&FRectItemSaleStatus
        end if

        if FRectDispCate<>"" then
                addSql = addSql + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"  ''조건추가 2016/04/14
		    	addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		strSqlCnt = " SELECT COUNT(i.itemid) FROM  [db_item].[dbo].[tbl_item] i,[db_event].[dbo].[tbl_sale] A, [db_event].[dbo].[tbl_saleItem] B "&_
				"	WHERE i.itemid= B.itemid and A.sale_code = B.sale_code  AND A.sale_using = 1  "&addSql
		''rsget.Open strSqlCnt,dbget
		rsget.CursorLocation = adUseClient
        rsget.Open strSqlCnt,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close
		'AND datediff(d,A.sale_enddate,getdate())<=0 삭제 2014-09
		IF FTotCnt >0 THEN
		    FSPageNo = (FPSize*(FCPage-1)) + 1
			FEPageNo = FPSize*FCPage

			strSql = "SELECT  sale_code,  itemid,  saleprice,  salesupplycash,  saleItem_status,  salelimitno,  orglimityn, makerid,  itemname,  smallimage "& vbcrlf
			strSql = strSql& " , sailyn, sellcash,  buycash, orgprice,  orgsuplycash,  sailprice, sailsuplycash,  mwdiv, limityn, limitno "& vbcrlf
			strSql = strSql& " , limitsold, isusing,  evt_code,  evtgroup_code ,  sale_startdate,  sale_enddate,  sale_status,  orgsailYN,  orgSailPrice, orgSailSuplyCash "& vbcrlf
			strSql = strSql&" FROM ( " &vbcrlf
			strSql = strSql& "      SELECT ROW_NUMBER() OVER (ORDER BY  B.saleitem_idx desc ) as RowNum "&vbcrlf
			strSql = strSql& "      ,A.sale_code, B.itemid, B.saleprice, B.salesupplycash, B.saleItem_status, B.limitno as salelimitno, B.orglimityn "& vbcrlf
			strSql = strSql& "	    ,i.makerid, i.itemname, i.smallimage ,i.sailyn,i.sellcash, i.buycash,i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash"&vbcrlf
			strSql = strSql& "	    , i.mwdiv,i.limityn,i.limitno, i.limitsold,i.isusing, A.evt_code, A.evtgroup_code , A.sale_startdate, A.sale_enddate, A.sale_status "&vbcrlf
			strSql = strSql& "	    , B.orgsailYN, B.orgSailPrice, B.orgSailSuplyCash "&vbcrlf
			strSql = strSql& " FROM [db_item].[dbo].[tbl_item] as i "&vbcrlf
			strSql = strSql&"      inner join [db_event].[dbo].[tbl_saleItem] as B on i.itemid= B.itemid "&vbcrlf
			strSql = strSql&"      inner join [db_event].[dbo].[tbl_sale] as A on A.sale_code = B.sale_code"&vbcrlf
			strSql = strSql& 		" WHERE   A.sale_using = 1  "& addSql &vbcrlf
			strSql = strSql&") as TB " &vbcrlf
			strSql = strSql&	" WHERE TB.RowNum Between "&FSPageNo&" AND "  &FEPageNo  &vbcrlf
			strSql = strSql&	" order by TB.RowNum "

'			strSql = ""
'            strSql = strSql& " SELECT top "&FEPageNo&" "&vbcrlf
'			strSql = strSql& "      A.sale_code, B.itemid, B.saleprice, B.salesupplycash, B.saleItem_status, B.limitno as salelimitno, B.orglimityn "& vbcrlf
'			strSql = strSql& "	    ,i.makerid, i.itemname, i.smallimage ,i.sailyn,i.sellcash, i.buycash,i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash"&vbcrlf
'			strSql = strSql& "	    , i.mwdiv,i.limityn,i.limitno, i.limitsold,i.isusing, A.evt_code, A.evtgroup_code , A.sale_startdate, A.sale_enddate, A.sale_status "&vbcrlf
'			strSql = strSql& "	    , B.orgsailYN, B.orgSailPrice, B.orgSailSuplyCash "&vbcrlf
'			strSql = strSql& " FROM [db_item].[dbo].[tbl_item] as i "&vbcrlf
'			strSql = strSql&"      inner join [db_event].[dbo].[tbl_saleItem] as B on i.itemid= B.itemid "&vbcrlf
'			strSql = strSql&"      inner join [db_event].[dbo].[tbl_sale] as A on A.sale_code = B.sale_code"&vbcrlf
'			strSql = strSql& 		" WHERE   A.sale_using = 1  "& addSql &vbcrlf
'			strSql = strSql&	" ORDER BY  A.sale_code desc, B.saleitem_idx desc "  ''A.sale_code desc 이걸 붙여야 빨라짐..
	'	response.write strSql
			''rsget.Open strSql,dbget
'			rsget.pageSize = FPSize
'			rsget.CursorLocation = adUseClient
            rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
			 '   rsget.absolutepage = FCPage
				fnGetSaleOnItemList = rsget.getRows()
			End IF
			rsget.Close
		END IF

	End Function
End Class
%>
