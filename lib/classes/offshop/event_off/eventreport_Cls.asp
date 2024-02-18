<%
'###########################################################
' Description :  오프라인이벤트 통계 클래스
' History : 2010.03.25 한용민 생성
'###########################################################

class Cevtreport_item
	public fshopname
	public fshopid
	public fevt_code
	public fevt_name
	public fevt_startdate
	public fevt_enddate
	public fsum_cnt
	public fsellsum
	public fshopregdate
	public Fitemid
	public fitemname
	public fitemoptionname
	public fitemoption
	public fmakerid
	public fitemgubun
	public fissale
	public fisgift
	public fisrack
	public fisprize
	public fisracknum
	public fevt_kind
	public fitem_count
	public fcate_nm1
	public fcate_nm2
	public ftotselljumuncnt
	public fimgbasic
	public fpartmdname

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class Cevtreport_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public maxc
	PUBLIC FRectOldJumun
	public frectevt_kind
	public frectevt_startdate
	public frectevt_enddate
	public frectevt_code
	public FRectReportType
	public FRectItemID
	public FRectShopID
	public frectevt_name
	public frectisgift
	public frectisrack
	public frectisprize
	public frectissale
	public FRectStartDay
	public FRectEndDay
	public frectdatefg
	public frectevt_cateL
	public frectevt_cateM
	public FRectInc3pl

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'/이벤트 통계 에서 가져옴
	'///admin/offshop/event_off/event_report.asp
    Public Sub getevent_sum()
        dim sqlStr,i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if

		'/이벤트기간 기준
		if frectdatefg = "event" then
			if frectevt_startdate <> "" and frectevt_enddate <> "" then
				sqlsearch = sqlsearch & " and ("
				sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&frectevt_startdate&"' and E.evt_startdate < '"&frectevt_enddate&"')"
				'sqlsearch = sqlsearch & " 	or ( dateadd(day,1,E.evt_enddate) >= '"&frectevt_startdate&"' and dateadd(day,1,E.evt_enddate) < '"&frectevt_enddate&"')"

				'sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&frectevt_startdate&"' and convert(varchar(10),E.evt_enddate,121) >= '"&frectevt_startdate&"') "
				'sqlsearch = sqlsearch & " 	or (E.evt_startdate > '"&frectevt_enddate&"' and convert(varchar(10),E.evt_enddate,121) > '"&frectevt_enddate&"') "
				sqlsearch = sqlsearch & " )"
			end if

		'/매출기간기준
		elseif frectdatefg = "jumun" then
			if frectevt_startdate <> "" and frectevt_enddate <> "" then
				sqlsearch = sqlsearch & " and ("
				sqlsearch = sqlsearch & " 	s.yyyymmdd >= '"&frectevt_startdate&"' and s.yyyymmdd < '"&frectevt_enddate&"'"
				sqlsearch = sqlsearch & " )"
			end if
		else
			if frectevt_startdate <> "" and frectevt_enddate <> "" then
				sqlsearch = sqlsearch & " and ("
				sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&frectevt_startdate&"' and E.evt_startdate < '"&frectevt_enddate&"')"
				'sqlsearch = sqlsearch & " 	or ( dateadd(day,1,E.evt_enddate) >= '"&frectevt_startdate&"' and dateadd(day,1,E.evt_enddate) < '"&frectevt_enddate&"')"

				'sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&frectevt_startdate&"' and convert(varchar(10),E.evt_enddate,121) >= '"&frectevt_startdate&"') "
				'sqlsearch = sqlsearch & " 	or (E.evt_startdate > '"&frectevt_enddate&"' and convert(varchar(10),E.evt_enddate,121) > '"&frectevt_enddate&"') "
				sqlsearch = sqlsearch & " )"
			end if
		end if
		if frectevt_kind<>"1" then
			sqlsearch = sqlsearch & " and E.evt_kind='" & frectevt_kind & "' "
		end if
		if frectevt_code <> "" then
			sqlsearch = sqlsearch & " and e.evt_code ='" & frectevt_code & "'"
		end if
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and es.AssignShopid ='" & FRectShopID & "'"
		end if
		if frectevt_name <> "" then
			sqlsearch = sqlsearch & " and e.evt_name like '%" & frectevt_name & "%'"
		end if
		IF frectissale <> "" THEN sqlsearch  = sqlsearch & " and b.issale = '"&frectissale&"'"
		IF frectisgift <> "" THEN sqlsearch  = sqlsearch & " and b.isgift = '"&frectisgift&"'"
		IF frectisrack <> "" THEN sqlsearch  = sqlsearch & " and b.israck = '"&frectisrack&"'"
		IF frectisprize <> "" THEN sqlsearch  = sqlsearch & " and b.isprize = '"&frectisprize&"'"

		if frectevt_cateL <> "" then
			sqlsearch = sqlsearch & " and c1.code_large = '" & frectevt_cateL & "' "
		end if
		if frectevt_cateM <> "" then
			sqlsearch = sqlsearch & " and c2.code_large = '" & frectevt_cateL & "' and c2.code_mid = '" & frectevt_cateM & "' "
		end if

		sqlStr = "SELECT top 500"
		sqlStr = sqlStr & " E.evt_code,E.evt_name,E.evt_startdate,E.evt_enddate ,es.AssignShopid , e.evt_kind,u.shopname"
		sqlStr = sqlStr & " , b.issale,b.isgift ,b.israck ,b.isprize, b.isracknum ,b.brand"
		sqlStr = sqlStr & " ,isnull(sum(s.totSellCnt),0) as sum_cnt"
		sqlStr = sqlStr & " ,isnull(sum(s.totselljumuncnt),0) as totselljumuncnt"
		sqlStr = sqlStr & " ,isnull(sum(s.totSellSum),0) as sellsum"
		sqlStr = sqlStr & " ,c1.code_nm as cate_nm1, isNull(B.img_basic, '') AS img_basic "

		if frectevt_cateM <> "" then
		sqlStr = sqlStr & " ,c2.code_nm as cate_nm2 "
		else
		sqlStr = sqlStr & " ,'' as cate_nm2 "
		end if

		sqlStr = sqlStr & " ,(select username from db_partner.dbo.tbl_user_tenbyten where userid <> '' and userid = isNull(B.partMDid,'')) AS partMDname "
		sqlStr = sqlStr & " FROM db_shop.[dbo].tbl_event_off E"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_event_off_AssignedShop es"
		sqlStr = sqlStr & " 	ON E.evt_code = es.evt_code"
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_event_off_sell_summary_daily S"
		sqlStr = sqlStr & " 	on E.evt_code=S.evt_code"
		sqlStr = sqlStr & " 	and es.AssignShopid = s.shopid"
		sqlStr = sqlStr & " LEFT JOIN db_shop.dbo.tbl_event_off_display as B"
		sqlStr = sqlStr & " 	ON e.evt_code = B.evt_code"
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr & " 	on es.AssignShopid = u.userid"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_Cate_large c1"
		sqlStr = sqlStr & "		on b.evt_category = c1.code_large"
		sqlStr = sqlStr & "		and c1.display_yn = 'Y'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on es.AssignShopid=p.id "

		if frectevt_cateM <> "" then
			sqlStr = sqlStr & " left join db_item.dbo.tbl_Cate_mid c2"
			sqlStr = sqlStr & "		on b.evt_category = c1.code_large"
			sqlStr = sqlStr & "		and c1.code_large = c2.code_large"
			sqlStr = sqlStr & "		and b.evt_catemid = c2.code_mid"
			sqlStr = sqlStr & "		and c2.display_yn = 'Y'"
		end if

		sqlStr = sqlStr & " where E.evt_using='Y'"
		sqlStr = sqlStr & " and s.yyyymmdd >= E.evt_startdate and s.yyyymmdd < dateadd(day,1,E.evt_enddate)"
		sqlStr = sqlStr & " and E.evt_state>=5 " & sqlsearch
		sqlStr = sqlStr & " GROUP BY E.evt_code,E.evt_name,E.evt_startdate,E.evt_enddate, e.evt_kind, u.shopname, es.AssignShopid"
		sqlStr = sqlStr & " 	, b.issale, b.isgift ,b.israck ,b.isprize, b.isracknum, b.brand, B.img_basic, B.partMDid"
		sqlStr = sqlStr & " 	,c1.code_nm"

		if frectevt_cateM <> "" then
		sqlStr = sqlStr & ", c2.code_nm "
		end if

		sqlStr = sqlStr & " order by E.evt_code desc"

        ''response.write sqlStr &"<Br>"
        rsget.open sqlStr,dbget

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new Cevtreport_item

			FItemList(i).fcate_nm1 = db2html(rsget("cate_nm1"))
			FItemList(i).fcate_nm2 = db2html(rsget("cate_nm2"))
			'FItemList(i).fitem_count = rsget("item_count")
			FItemList(i).ftotselljumuncnt = rsget("totselljumuncnt")
			FItemList(i).fevt_kind = rsget("evt_kind")
			FItemList(i).fisracknum = rsget("isracknum")
			FItemList(i).fissale = rsget("issale")
			FItemList(i).fisgift = rsget("isgift")
			FItemList(i).fisrack = rsget("israck")
			FItemList(i).fisprize = rsget("isprize")
			FItemList(i).fevt_code = rsget("evt_code")
			FItemList(i).fshopname = rsget("shopname")
			FItemList(i).fshopid = rsget("AssignShopid")
			FItemList(i).fevt_name = db2html(rsget("evt_name"))
			FItemList(i).fevt_startdate = left(rsget("evt_startdate"),10)
			FItemList(i).fevt_enddate = left(rsget("evt_enddate"),10)
			FItemList(i).fsum_cnt = rsget("sum_cnt")
			FItemList(i).fsellsum = rsget("sellsum")
			FItemList(i).fmakerid = rsget("brand")
			FItemList(i).fimgbasic = rsget("img_basic")
			FItemList(i).fpartmdname = rsget("partMDname")

		rsget.MoveNext
		i = i + 1
		loop

		rsget.close
    End Sub

	'///admin/offshop/event_off/event_report_detail.asp
    Public Sub geteventdate_sum()
		dim sqlStr,i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectevt_code <> "" then
			sqlsearch = sqlsearch & " and s.evt_code ='" & frectevt_code & "'"
		end if
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and es.AssignShopid ='" & FRectShopID & "'"
		end if
		'/이벤트기간 기준
		if frectdatefg = "event" then
			if FRectStartDay <> "" and FRectEndDay <> "" then
				sqlsearch = sqlsearch & " and ("
				'sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&FRectStartDay&"' and convert(varchar(10),E.evt_enddate,121) >= '"&FRectStartDay&"') "
				'sqlsearch = sqlsearch & " 	or (E.evt_startdate > '"&FRectEndDay&"' and convert(varchar(10),E.evt_enddate,121) > '"&FRectEndDay&"') "

				sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&FRectStartDay&"' and E.evt_startdate < '"&FRectEndDay&"')"
				'sqlsearch = sqlsearch & " 	or ( dateadd(day,1,E.evt_enddate) >= '"&FRectStartDay&"' and dateadd(day,1,E.evt_enddate) < '"&FRectEndDay&"')"
				sqlsearch = sqlsearch & " )"
			end if

		'/매출기간기준
		elseif frectdatefg = "jumun" then
			if FRectStartDay <> "" and FRectEndDay <> "" then
				sqlsearch = sqlsearch & " and ("
				sqlsearch = sqlsearch & " 	s.yyyymmdd >= '"&FRectStartDay&"' and s.yyyymmdd < '"&FRectEndDay&"'"
				sqlsearch = sqlsearch & " )"
			end if
		else
			if FRectStartDay <> "" and FRectEndDay <> "" then
				sqlsearch = sqlsearch & " and ("
				'sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&FRectStartDay&"' and convert(varchar(10),E.evt_enddate,121) >= '"&FRectStartDay&"') "
				'sqlsearch = sqlsearch & " 	or (E.evt_startdate > '"&FRectEndDay&"' and convert(varchar(10),E.evt_enddate,121) > '"&FRectEndDay&"') "

				sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&FRectStartDay&"' and E.evt_startdate < '"&FRectEndDay&"')"
				'sqlsearch = sqlsearch & " 	or ( dateadd(day,1,E.evt_enddate) >= '"&FRectStartDay&"' and dateadd(day,1,E.evt_enddate) < '"&FRectEndDay&"')"
				sqlsearch = sqlsearch & " )"
			end if
		end if

		'데이터 리스트
		sqlStr = "select top 1000"
		sqlStr = sqlStr & " isnull(s.totSellCnt,0) as sum_cnt"
		sqlStr = sqlStr & " ,isnull(s.totselljumuncnt,0) as totselljumuncnt"
		sqlStr = sqlStr & " , isnull(s.totSellSum,0) as sellsum"
		sqlStr = sqlStr & " ,s.yyyymmdd as shopregdate, e.evt_name, e.evt_code, es.AssignShopid"
		sqlStr = sqlStr & " FROM db_shop.[dbo].tbl_event_off E"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_event_off_AssignedShop es"
		sqlStr = sqlStr & " 	ON E.evt_code = es.evt_code"
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_event_off_sell_summary_daily S"
		sqlStr = sqlStr & " 	on E.evt_code=S.evt_code"
		sqlStr = sqlStr & " 	and es.AssignShopid = s.shopid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on es.AssignShopid=p.id "
		sqlStr = sqlStr & " where E.evt_using='Y'"
		sqlStr = sqlStr & " and s.yyyymmdd >= E.evt_startdate and s.yyyymmdd < dateadd(day,1,E.evt_enddate)"
		sqlStr = sqlStr & " and E.evt_state>=5 " & sqlsearch
		sqlStr = sqlStr & " order by shopregdate desc ,es.AssignShopid asc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FItemList(i) = new Cevtreport_item

				FItemList(i).ftotselljumuncnt = rsget("totselljumuncnt")
				FItemList(i).fshopid = rsget("AssignShopid")
				FItemList(i).fshopregdate = rsget("shopregdate")
				FItemList(i).fsum_cnt = rsget("sum_cnt")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fevt_name = rsget("evt_name")
				FItemList(i).fevt_code = rsget("evt_code")

				if Not IsNull(FItemList(i).fsellsum) then
					maxc = MaxVal(maxc,FItemList(i).fsellsum)
				end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/상품별 이벤트 상품 판매 통계
	Public Sub geteventitem_sum
		dim sqlStr , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectevt_code <> "" then
			sqlsearch = sqlsearch & " and e.evt_code ='" & frectevt_code & "'"
		end if
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and es.AssignShopid ='" & FRectShopID & "'"
		end if
		'/이벤트기간 기준
		if frectdatefg = "event" then
			if FRectStartDay <> "" and FRectEndDay <> "" then
				sqlsearch = sqlsearch & " and ("
				'sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&FRectStartDay&"' and convert(varchar(10),E.evt_enddate,121) >= '"&FRectStartDay&"') "
				'sqlsearch = sqlsearch & " 	or (E.evt_startdate > '"&FRectEndDay&"' and convert(varchar(10),E.evt_enddate,121) > '"&FRectEndDay&"') "

				sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&FRectStartDay&"' and E.evt_startdate < '"&FRectEndDay&"')"
				'sqlsearch = sqlsearch & " 	or ( dateadd(day,1,E.evt_enddate) >= '"&FRectStartDay&"' and dateadd(day,1,E.evt_enddate) < '"&FRectEndDay&"')"
				sqlsearch = sqlsearch & " )"
			end if

		'/매출기간기준
		elseif frectdatefg = "jumun" then
			if FRectStartDay <> "" and FRectEndDay <> "" then
				sqlsearch = sqlsearch & " and ("
				sqlsearch = sqlsearch & " 	om.IXyyyymmdd >= '"&FRectStartDay&"' and om.IXyyyymmdd < '"&FRectEndDay&"'"
				sqlsearch = sqlsearch & " )"
			end if
		else
			if FRectStartDay <> "" and FRectEndDay <> "" then
				sqlsearch = sqlsearch & " and ("
				'sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&FRectStartDay&"' and convert(varchar(10),E.evt_enddate,121) >= '"&FRectStartDay&"') "
				'sqlsearch = sqlsearch & " 	or (E.evt_startdate > '"&FRectEndDay&"' and convert(varchar(10),E.evt_enddate,121) > '"&FRectEndDay&"') "

				sqlsearch = sqlsearch & " 	(E.evt_startdate >= '"&FRectStartDay&"' and E.evt_startdate < '"&FRectEndDay&"')"
				'sqlsearch = sqlsearch & " 	or ( dateadd(day,1,E.evt_enddate) >= '"&FRectStartDay&"' and dateadd(day,1,E.evt_enddate) < '"&FRectEndDay&"')"
				sqlsearch = sqlsearch & " )"
			end if
		end if

		sqlStr = "SELECT top 1000"
		sqlStr = sqlStr & " OD.itemgubun ,OD.itemid ,OD.itemoption,OD.itemname,od.itemoptionname, od.makerid"
		sqlStr = sqlStr & " ,isnull(sum(OD.itemNo),0) as sum_cnt"
		sqlStr = sqlStr & " ,isnull(sum(OD.realsellprice*OD.itemNo),0) as sellsum"
		sqlStr = sqlStr & " ,e.evt_code,E.evt_name"
		sqlStr = sqlStr & " ,es.AssignShopid"
		sqlStr = sqlStr & " FROM db_shop.[dbo].tbl_event_off E "
		sqlStr = sqlStr & " JOIN db_shop.dbo.tbl_eventitem_off ET"
		sqlStr = sqlStr & " 	ON E.evt_code = ET.evt_code "
		sqlStr = sqlStr & " join db_shop.dbo.tbl_event_off_AssignedShop es"
		sqlStr = sqlStr & " 	ON E.evt_code = es.evt_code"
		sqlStr = sqlStr & " JOIN db_shop.dbo.tbl_shopjumun_detail OD "
		sqlStr = sqlStr & " 	ON ET.itemgubun = OD.itemgubun"
		sqlStr = sqlStr & " 	and ET.itemid = OD.itemid"
		sqlStr = sqlStr & " 	and ET.itemoption = OD.itemoption"
		sqlStr = sqlStr & " JOIN db_shop.dbo.tbl_shopjumun_master OM"
		sqlStr = sqlStr & " 	On OM.idx=OD.masteridx"
		sqlStr = sqlStr & " 	and es.AssignShopid = OM.shopid"
		sqlStr = sqlStr & " 	and OM.idx<>0"
		sqlStr = sqlStr & " 	and OM.cancelyn='N'"
		sqlStr = sqlStr & " 	and OD.cancelyn='N'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on es.AssignShopid=p.id "
		sqlStr = sqlStr & " WHERE E.evt_using='Y'"
		sqlStr = sqlStr & " and om.IXyyyymmdd >= E.evt_startdate and om.IXyyyymmdd < dateadd(day,1,E.evt_enddate)"
		sqlStr = sqlStr & " and E.evt_state>=5 " & sqlsearch
		sqlStr = sqlStr & " group by"
		sqlStr = sqlStr & " OD.itemgubun ,OD.itemid ,OD.itemoption ,od.makerid,OD.itemname,od.itemoptionname"
		sqlStr = sqlStr & " ,e.evt_code, E.evt_name"
		sqlStr = sqlStr & " ,es.AssignShopid"
		sqlStr = sqlStr & " order by sellsum desc , OD.itemid desc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if not rsget.eof then
			do until rsget.eof
				set FItemList(i) = new Cevtreport_item

				FItemList(i).fshopid 	= rsget("AssignShopid")
				FItemList(i).Fitemid 	= rsget("itemid")
				FItemList(i).fitemgubun 	= rsget("itemgubun")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fsum_cnt 	= rsget("sum_cnt")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fevt_name = db2html(rsget("evt_name"))
				FItemList(i).fitemname = db2html(rsget("itemname"))
				FItemList(i).fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fmakerid = rsget("makerid")

				if Not IsNull(FItemList(i).fsellsum) then
					maxc = MaxVal(maxc,FItemList(i).fsellsum)
				end if

			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close
	end Sub

end class
%>
