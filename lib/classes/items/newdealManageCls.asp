<%
'####################################################
' Page : /lib/classes/event/dealManageCls.asp
' Description :  딜 이벤트 메인 클래스
' History : 2017.08.22 정태훈 생성
'####################################################

Class ClsDeal
	public FCPage
	public FPSize
	public FSfDate
	public FSsDate
	public FSeDate
	public FSdispCate
	public FSeTxt
	public FTotCnt
	public FMaxNum
	public FSearchDiv
	public FSearchDateDiv
	public FSIsUsing
	public FSViewDiv
	public FMasterIDX
	public FRectMasterIDX

	'## fnGetDealList : 생성 메일 목록  ##
	public Function fnGetDealList
	Dim strSql, strSqlCnt, strSearch, strSubSort, strSort
	
	strSearch = ""
 
	'//검색조건    
	If FSsDate <> ""  or FSeDate <> "" Then
		If FSearchDateDiv="R" Then
			strSearch  = strSearch & " and  d.regdate>='"&FSsDate&"' and  d.regdate<='"&FSeDate&"'"
		Else
			If FSearchDateDiv="S" Then
				strSearch  = strSearch & " and  datediff(day, '"&FSsDate&"', d.startdate) >= 0 and  datediff(day,'"&FSeDate&"', d.startdate) <=0  "
			Else
				strSearch  = strSearch & " and  datediff(day, '"&FSsDate&"', d.enddate) >= 0 and  datediff(day,'"&FSeDate&"', d.enddate) <=0  "
			End If
		End If
	End If

	If FSViewDiv<>"" Then
		strSearch  = strSearch & " and  d.viewdiv='"&FSViewDiv&"'"	
	End If

	If FSIsUsing<>"" Then
		strSearch  = strSearch & " and  i.isusing='"&FSIsUsing&"'"	
	End If

	If FSdispCate<>"" Then
		strSearch  = strSearch & " and  left(c.catecode,6)='"&left(FSdispCate,6)&"'"	
	End If

	If FSeTxt <> "" Then
		If FSearchDiv="itemid" Then
			strSearch  = strSearch & " and  d.dealitemid='"&FSeTxt&"'"	
		ElseIf FSearchDiv="itemname" Then
			strSearch  = strSearch & " and  i.itemname like '%"&FSeTxt&"%'"	
		Else
			strSearch  = strSearch & " and  p.username like '%"&FSeTxt&"%'"	
		End If
	End If

	strSqlCnt = " SELECT COUNT(d.idx) FROM [db_event].[dbo].[tbl_deal_event] d" 
	strSqlCnt = strSqlCnt + " left join [db_item].[dbo].[tbl_item] i on i.itemid=d.dealitemid" + vbcrlf
	strSqlCnt = strSqlCnt + " left join [db_item].[dbo].[tbl_display_cate_item] c on c.itemid=i.itemid and c.isDefault='y'" + vbcrlf
	strSqlCnt = strSqlCnt + " LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as p on p.userid=d.regname" + vbcrlf
	strSqlCnt =	strSqlCnt &	" WHERE d.dealitemid<>0"&strSearch
'response.Write strSqlCnt
	rsget.Open strSqlCnt,dbget
	IF not rsget.EOF THEN
		FTotCnt = rsget(0)
	End IF
	rsget.Close
 

	IF FTotCnt >0 THEN
		'이벤트 기간 종료시 상태 종료로 , 이벤트 오픈요청상태에서 기간이 진행중일때 상태 오픈으로 view 처리
		dim iSPageNo, iEPageNo
		iSPageNo = (FPSize*(FCPage-1)) + 1
		iEPageNo = FPSize*FCPage	
					
		strSql = " SELECT TOP "&FPSize*FCPage&" d.idx, d.dealitemid, d.viewdiv, d.startdate, d.enddate, i.itemname, d.mastersellcash, d.masterdiscountrate, p.username, d.regdate, d.pricesdash, d.sailsdash, cn.cateFullName, i.isusing" + vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_deal_event] d" + vbcrlf
		strSql = strSql + " left join [db_item].[dbo].[tbl_item] i on i.itemid=d.dealitemid" + vbcrlf
		strSql = strSql + " left join [db_item].[dbo].[tbl_display_cate_item] c on c.itemid=i.itemid and c.isDefault='y'" + vbcrlf
		strSql = strSql + " left join [ANALDB].[db_analyze_data_raw].[dbo].[tbl_display_cate] cn on cn.catecode=c.catecode" + vbcrlf
		strSql = strSql + " LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as p on p.userid=d.regname" + vbcrlf
		strSql = strSql + " WHERE  d.dealitemid<>0" + strSearch + vbcrlf
		strSql = strSql + " order by d.idx desc"
		  
		'response.Write strSql
		'Response.end
		rsget.pagesize = FPSize
		rsget.Open strSql,dbget,1
		rsget.absolutepage = FCPage
		IF not rsget.EOF THEN
			fnGetDealList = rsget.getRows()
		End IF
		rsget.Close
	End IF


	End Function

	public Function fnGetMAXDealMasterNum
		Dim strSql
		strSql = " insert into [db_event].[dbo].[tbl_deal_event](viewdiv, regname) values(0,'admin')" 
		dbget.execute strSql

		strSql = "select SCOPE_IDENTITY()"
		rsget.Open strSql, dbget, 0
		FMasterIDX = rsget(0)
		rsget.Close
	End Function

	public Function fnGetMAXDealSalePer
		Dim strSql
		strSql = " SELECT  TOP 1 B.orgprice, B.sailprice, B.sailyn, case sailyn when 'Y' then ((orgprice-sailprice)/ orgprice)*100 else 0 end  as sailpercent, B.itemid "&_
				"	FROM [db_event].[dbo].[tbl_deal_event_item] AS A " &_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid "&_
				"	WHERE A.dealcode = " &FRectMasterIDX & " and  A.isusing='Y'"&_ 
				"   order by sailpercent desc"
		'  response.write strSql&"<BR>"
		rsget.Open strSql,dbget,1
		IF not rsget.EOF THEN
			fnGetMAXDealSalePer = rsget.getRows()
		End IF
		rsget.Close
	End Function

	public Function fnGetMAXDealCouponSalePer
		Dim strSql
		strSql = " SELECT  TOP 1 B.itemcouponvalue, B.itemid"&_
				"	FROM [db_event].[dbo].[tbl_deal_event_item] AS A " &_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid "&_
				"	WHERE A.dealcode = " &FRectMasterIDX & " and  A.isusing='Y' and B.itemcoupontype=1"&_ 
				"   order by itemcouponvalue desc"
		'  response.write strSql&"<BR>"
		rsget.Open strSql,dbget,1
		IF not rsget.EOF THEN
			fnGetMAXDealCouponSalePer = rsget.getRows()
		End IF
		rsget.Close
	End Function

	public Function fnGetDealItemMinPrice
		Dim strSql
		strSql = " SELECT  TOP 1 B.sellcash, B.itemid "&_
				"	FROM [db_event].[dbo].[tbl_deal_event_item] AS A " &_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid "&_
				"	WHERE A.dealcode = " &FRectMasterIDX & " and  A.isusing='Y'"&_ 
				"   order by sellcash asc"
		'  response.write strSql&"<BR>"
		rsget.Open strSql,dbget,1
		IF not rsget.EOF THEN
			fnGetDealItemMinPrice = rsget.getRows()
		End IF
		rsget.Close
	End Function

End Class

Class CDealView
	public Fidx
	public Fdealitemid
	public Fmasteritemcode
	public Fviewdiv
	public Fstartdate
	public Fenddate
	public Fitemname
	public Fmastersellcash
	public Fmasterdiscountrate
	public Fregname
	public Fregdate
	public Fisusing
	public Fsellyn
	public Fpricesdash
	public Fsailsdash
	public Fwork_notice

	public FRectMasterIDX
	public FResultCount

	Public Function GetDealView()
        dim sqlStr, addSql, i

        '// 본문 내용 접수
        sqlStr = "select top 1"
        sqlStr = sqlStr & " d.idx, d.dealitemid, d.masteritemcode, d.viewdiv, d.startdate, d.enddate, i.itemname, d.mastersellcash, d.masterdiscountrate, d.regname, d.regdate, i.isusing, d.pricesdash, d.sailsdash, d.work_notice, i.sellyn"
        sqlStr = sqlStr & " from [db_event].[dbo].[tbl_deal_event] d, [db_item].[dbo].[tbl_item] i"
        sqlStr = sqlStr & " WHERE  d.dealitemid=i.itemid and d.idx=" & CStr(FRectMasterIDX)
        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
        if Not(rsget.EOF or rsget.BOF) then
			Fidx	= rsget("idx")
			Fdealitemid	= rsget("dealitemid")
			Fmasteritemcode	= rsget("masteritemcode")
			Fviewdiv	= rsget("viewdiv")
			Fstartdate	= rsget("startdate")
			Fenddate	= rsget("enddate")
			Fitemname	= rsget("itemname")
			Fmastersellcash	= rsget("mastersellcash")
			Fmasterdiscountrate	= rsget("masterdiscountrate")
			Fregname	= rsget("regname")
			Fregdate	= rsget("regdate")
			Fisusing	= rsget("isusing")
			Fpricesdash	= rsget("pricesdash")
			Fsailsdash	= rsget("sailsdash")
			Fwork_notice	= rsget("work_notice")
			Fsellyn		= rsget("sellyn")
        end if
        rsget.Close
    end function

	Private Sub Class_Initialize()
		FResultCount = 0
	End Sub

	Private Sub Class_Terminate()
    End Sub
end Class

Class CDealItem
	Public FTotCnt
	public FRectMasterIDX
	Public FPSize
	Public FCPage

	'## fnGetDealEventItem :딜 이벤트 상품 가져오기 ##
	public Function fnGetDealEventItem

		Dim strSql, strSqlCnt,iDelCnt
		Dim strSort,strGroup, striSort,addSql
		dim addSort

		addSort = ""

		strSqlCnt = strSqlCnt &" SELECT COUNT(d.itemid) as Totcnt "&vbCrlf
		strSqlCnt = strSqlCnt &" FROM [db_event].[dbo].[tbl_deal_event_item] d, [db_item].[dbo].[tbl_item] i "&vbCrlf
		strSqlCnt =	strSqlCnt &"	WHERE  d.itemid=i.itemid  and d.dealcode = "&FRectMasterIDX
		rsget.Open strSqlCnt,dbget,1
		If Not rsget.EOF Then
			FTotCnt = rsget(0)
		End If
		rsget.Close   
		If FTotCnt >0 Then
			iDelCnt =  (FCPage - 1) * FPSize
			strSql = " SELECT  TOP "&FPSize*FCPage&" d.viewidx, d.itemid, i.itemname, i.buycash,  i.sellcash "&_
					" ,i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash, i.sailyn, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue "&_
					" , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice"&_
					" FROM  [db_event].[dbo].[tbl_deal_event_item] d "&_
					" left join [db_item].[dbo].[tbl_item] i on i.itemid=d.itemid" &_
					" WHERE d.dealcode = " & FRectMasterIDX &_
					" ORDER BY   d.viewidx asc"
			'  response.write strSql&"<BR>"
			
			rsget.pagesize = FPSize
			rsget.Open strSql,dbget,1
			
			rsget.absolutepage = FCPage
			IF not rsget.EOF THEN
				fnGetDealEventItem = rsget.getRows()
			End IF
			rsget.Close
		END IF
	End Function

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

Class CDealSelect
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectDealCode

	public function GetDealSelectItemList()
        dim sqlStr, addSql, i

        '// 본문 내용 접수
        sqlStr = "select top 200"
        sqlStr = sqlStr & " d.idx, d.itemid, d.itemname, d.viewidx, i.sellyn, i.limityn, i.limitno, i.limitsold"
        sqlStr = sqlStr & " from [db_event].[dbo].[tbl_deal_event_item] d"
		sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item] i on i.itemid=d.itemid"
        sqlStr = sqlStr & " where dealcode=" & CStr(FRectDealCode)
		sqlStr = sqlStr & " Order by viewidx asc"
        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)
        i=0
        if Not(rsget.EOF or rsget.BOF) then
            do until rsget.EOF
                set FItemList(i) = new CDealSelectItem
				FItemList(i).Fidx	= rsget("idx")
                FItemList(i).Fitemid	= rsget("itemid")
                FItemList(i).Fitemname	= rsget("itemname")
                FItemList(i).Fviewidx	= rsget("viewidx")
				FItemList(i).FSellYn	= rsget("sellyn")
				FItemList(i).FLimitYn	= rsget("limityn")
				FItemList(i).FLimitNo	= rsget("limitno")
				FItemList(i).FLimitSold	= rsget("limitsold")
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
    End Sub
end Class
%>