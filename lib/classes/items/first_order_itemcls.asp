<%
Class CDealItem
	Public FTotCnt
	public FRectMasterIDX
	Public FPSize
	Public FCPage

	Public FRectMakerid
	Public FRectItemid
	Public FRectItemName
	Public FRectSellYN
	Public FRectDispCate
	Public FESGroup
	Public FESSort

	public Function fnGetDealEventItemNew

		Dim strSql, strSqlCnt,iDelCnt
		Dim strSort,strGroup, striSort,addSql
		dim addSort

		addSort = ""

        if (FRectMakerid <> "") then
            addSql = addSql & " and B.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            addSql = addSql & " and B.itemid in (" + FRectItemid + ")"
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and B.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectSellYN <> "") then
            addSql = addSql & " and B.sellyn='" + FRectSellYN + "'"
        end if

		if FRectDispCate<>"" then
			if LEN(FRectDispCate)>3 then
					addSql = addSql + " and B.dispcate1='"&LEFT(FRectDispCate,3)&"'"
			end if
			addSql = addSql + " and B.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		IF FESSort = "1" THEN
			strSort = "A.itemid DESC "
		ELSEIF FESSort = "2" THEN
			strSort = "B.sellyn DESC, lsold, B.sellcash ASC"
		ELSEIF FESSort = "3" THEN
			strSort = "B.sellyn DESC, lsold, A.sort, A.itemid DESC"
		ELSEIF FESSort = "4" THEN
			strSort = "B.sellyn DESC, lsold, E.recentsellcount DESC, E.sellcount DESC, B.itemid DESC"
		ELSEIF FESSort = "5" THEN
			strSort = "B.sellyn DESC, lsold, B.sellcash DESC"
		ELSEIF FESSort = "6" THEN
			strSort = "sailpercent DESC, B.sellyn DESC, lsold, B.sellcash DESC"
		ELSEIF FESSort = "7" THEN
			strSort = "B.sellyn DESC, lsold, E.favcount DESC"
		ELSEIF FESSort = "8" THEN
			strSort = "E.favcount DESC"
		ELSEIF FESSort = "9" THEN
			strSort = "group_code DESC, A.sort, A.itemid DESC"
		ELSEIF FESSort = "10" THEN
			strSort = "makerid, A.sort, A.itemid DESC"
		ELSE
			strSort = "A.sort, A.itemid DESC"
		END IF

		strSqlCnt = strSqlCnt & " SELECT COUNT(A.itemid) as Totcnt" & vbcrlf
		strSqlCnt = strSqlCnt & " FROM [db_event].[dbo].[tbl_event_first_order_item] AS A with (nolock)" & vbcrlf
		strSqlCnt = strSqlCnt & " join [db_item].[dbo].[tbl_item] B with (nolock) on A.itemid=B.itemid" & vbcrlf
		strSqlCnt =	strSqlCnt & " WHERE 1=1 " & addSql
		rsget.Open strSqlCnt,dbget,1
		If Not rsget.EOF Then
			FTotCnt = rsget(0)
		End If
		rsget.Close   
		If FTotCnt >0 Then
			iDelCnt =  (FCPage - 1) * FPSize
			strSql = " SELECT  TOP " & FPSize * FCPage & " A.itemid, B.itemname, B.buycash, B.sellcash, B.orgprice, B.orgsuplycash, B.sailprice, B.sailsuplycash, B.sailyn, B.itemcouponyn" & vbcrlf
			strSql = strSql & " , B.sellyn, B.limityn, B.limitno, B.limitsold, B.smallimage, B.itemcoupontype, B.itemcouponvalue, B.listimage, A.sort" & vbcrlf
			strSql = strSql & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=B.curritemcouponidx and itemid=B.itemid) end as couponbuyprice" & vbcrlf
			strSql = strSql & " , case B.sailyn when 'Y' then ((B.orgprice-B.sailprice)/B.orgprice)*100 else 0 end as sailpercent" & vbcrlf
			strSql = strSql & " , case B.limityn when 'Y' then case when ((B.limitno-B.limitsold)<=0) then '2' else '1' end Else '1' end as lsold, B.makerid, A.isusing" & vbcrlf
			strSql = strSql & " FROM [db_event].[dbo].[tbl_event_first_order_item] AS A with (nolock) " & vbcrlf
			strSql = strSql & " LEFT JOIN [db_item].[dbo].[tbl_item] B with (nolock) on B.itemid=A.itemid" & vbcrlf
			strSql = strSql & " LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid " & vbcrlf
			strSql = strSql & " WHERE 1=1 " & FRectMasterIDX & strGroup & addSql & vbcrlf
			strSql = strSql & " ORDER BY " & strSort
			'  response.write strSql&"<BR>"
			'  response.end
			
			rsget.pagesize = FPSize
			rsget.Open strSql,dbget,1
			
			rsget.absolutepage = FCPage
			IF not rsget.EOF THEN
				fnGetDealEventItemNew = rsget.getRows()
			End IF
			rsget.Close
		END IF
	End Function

    public Function IsSoldOut(FSellYn,FLimitYn,FLimitNo,FLimitSold)
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa(FLimitNo,FLimitSold)<1))
	end function

    public function GetLimitEa(FLimitNo,FLimitSold)
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	Private Sub Class_Initialize()
		FCPage=1
		FPSize=100
	End Sub

	Private Sub Class_Terminate()
    End Sub
end Class

Class CDealSelectItem
    public Fidx
	public Fitemid
    public Fitemname
    public Fviewidx
	public FSellYn
	public FLimitYn
	public FLimitNo
	public FLimitSold

    public Function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

	public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class
%>