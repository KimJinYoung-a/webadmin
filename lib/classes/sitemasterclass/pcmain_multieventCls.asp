<%
'###############################################
' PageName :multievent
' Discription : PC메인 1~8 이벤트
' History : 2018-03-13 이종화 생성
'###############################################

Class CMainbannerItem
	public Fidx
	public Feventid
	public Fmaincopy
	public Fsubcopy
	public Fstartdate
	public Fenddate
	public Fevtstdate
	public Fevteddate
	public Fadminid
	public Flastadminid
	public Fisusing
	public Fordertext
	public Fsortnum
	public Fevtmolistbanner
	public Fsale_per
	public Fcoupon_per
	Public Fregdate
	Public Flastupdate
	Public Flinkurl
	public Ftag_only
	public FdispOption
	public FcontentImg
	public FcontentType
	public FItemId
	'추가
	public FEventInfo
	public FEventInfoOption
	public FESale
	public FEGift
	public FECoupon
	public FECommnet
	public FSisOnlyTen
	public FEOneplusOne
	public FEFreedelivery
	public FENew
	public FESalePer
	public FECsalePer	

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CMainbanner
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	Public FRectvaliddate
	public FRectSelDateTime
	public FRectDispOption

    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	Public FRectsedatechk
	Public FRecttype

	'//admin/appmanage/today/enjoyevent/enjoy_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 t.* , d.evt_mo_listbanner "
		sqlStr = sqlStr & " , d.issale, d.isgift, d.iscoupon, d.isOnlyTen, d.isoneplusone, d.isfreedelivery, d.isbookingsell, d.iscomment, d.ISNEW, d.SALEPER, d.SALECPER "		
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_pcmain_enjoyevent as t "
		sqlStr = sqlStr & " left outer join db_event.dbo.tbl_event_display as d "
        sqlStr = sqlStr & " on t.eventid = d.evt_code"
        sqlStr = sqlStr & " where idx=" & CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainbannerItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
			FOneItem.Feventid			= rsget("eventid")
			FOneItem.Fmaincopy			= rsget("maincopy")
			FOneItem.Fsubcopy			= rsget("subcopy")
			FOneItem.Fstartdate			= rsget("startdate")
			FOneItem.Fenddate			= rsget("enddate")
			FOneItem.Fevtstdate			= rsget("evtstdate")
			FOneItem.Fevteddate			= rsget("evteddate")
			FOneItem.Fadminid			= rsget("adminid")
			FOneItem.Flastadminid		= rsget("lastadminid")
			FOneItem.Fisusing			= rsget("isusing")
			FOneItem.Fordertext			= rsget("ordertext")
			FOneItem.Fsortnum			= rsget("sortnum")
			FOneItem.Fevtmolistbanner	= rsget("evt_mo_listbanner")
			FOneItem.Fsale_per			= rsget("sale_per")
			FOneItem.Fcoupon_per		= rsget("coupon_per")
			FOneItem.Flinkurl			= rsget("linkurl")
			FOneItem.Ftag_only			= rsget("tag_only")
			FOneItem.FdispOption		= rsget("dispOption")
			FOneItem.FcontentType		= rsget("contentType")
			FOneItem.FcontentImg 		= rsget("contentImg")	
			FOneItem.FItemId	 		= rsget("itemId")	

			'추가
			FOneItem.FEventInfoOption	 = rsget("event_info_option")
			FOneItem.FEventInfo			 = rsget("event_info")
			FOneItem.FESale				 = rsget("issale")
			FOneItem.FEGift				 = rsget("isgift")
			FOneItem.FECoupon			 = rsget("iscoupon")
			FOneItem.FECommnet			 = rsget("iscomment")	
			FOneItem.FSisOnlyTen		 = rsget("isOnlyTen")
			FOneItem.FEOneplusOne		 = rsget("isoneplusone") 
			FOneItem.FEFreedelivery		 = rsget("isfreedelivery")
			FOneItem.FENew 				 = rsget("isnew")
			FOneItem.FECsalePer 		 = rsget("SALECPER")
			FOneItem.FESalePer   	 	 = rsget("SALEPER")			
        end If

        rsget.Close
    end Sub

	'//admin/appmanage/today/enjoyevent/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_pcmain_enjoyevent "
		sqlStr = sqlStr & " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr & " and isusing='" & CStr(Fisusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and startdate = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then
			sqlStr = sqlStr & " and '" & Fsdt & "' between convert(varchar(10),startdate,120) and convert(varchar(10),enddate,120) "
		End If

		if Fsdt<> "" and FRectSelDateTime <> "00" then 
            sqlStr = sqlStr + " and datepart(hh , startdate) >=" &FRectSelDateTime
        end if 

		If FRectvaliddate = "on" Then
			sqlStr = sqlStr & " and enddate > getdate() "
		End If

		if FRectDispOption <> "" then 
            sqlStr = sqlStr + " and dispOption =" &FRectDispOption
        end if 

		'response.write sqlStr &"<br><br/>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "SELECT TOP " & CStr(FPageSize * FCurrPage) & " "
		 sqlStr = sqlStr & "t.idx  , d.evt_mo_listbanner , t.maincopy , t.subcopy , t.startdate , t.enddate ,t.regdate , t.adminid , t.lastadminid , t.lastupdate , t.sortnum , t.isusing, t.dispOption, t.contentType, t.contentImg  "
        sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_pcmain_enjoyevent AS t WITH(NOLOCK) "
        sqlStr = sqlStr & " LEFT OUTER JOIN db_event.dbo.tbl_event_display AS d WITH(NOLOCK)"
        sqlStr = sqlStr & " ON t.eventid = d.evt_code"
        sqlStr = sqlStr & " WHERE 1=1"

		'Response.write sqlStr
        if Fisusing<>"" then
            sqlStr = sqlStr & " and isusing='" & CStr(Fisusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and startdate = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then
			sqlStr = sqlStr & " and '" & Fsdt & "' between convert(varchar(10),startdate,120) and convert(varchar(10),enddate,120) "
		End If

		if Fsdt<> "" and FRectSelDateTime <> "00" then 
            sqlStr = sqlStr + " and datepart(hh , startdate) >=" &FRectSelDateTime
        end if 

		If FRecttype <> "" Then
			sqlStr = sqlStr & " and addtype = '"& FRecttype &"'"
		End If

		If FRectvaliddate = "on" Then
			sqlStr = sqlStr & " and t.enddate > getdate() "
		End If

		if FRectDispOption <> "" then 
            sqlStr = sqlStr + " and dispOption =" &FRectDispOption
        end if 

		sqlStr = sqlStr & " order by t.sortnum asc , t.idx desc "

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMainbannerItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fevtmolistbanner	= rsget("evt_mo_listbanner")
				FItemList(i).Fmaincopy			= rsget("maincopy")
				FItemList(i).Fsubcopy			= rsget("subcopy")
				FItemList(i).Fstartdate			= rsget("startdate")
				FItemList(i).Fenddate			= rsget("enddate")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid		= rsget("lastadminid")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fsortnum			= rsget("sortnum")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).FdispOption		= rsget("dispOption")
				FItemList(i).FcontentType		= rsget("contentType")				
				FItemList(i).FcontentImg		= rsget("contentImg")				

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub


    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

	End Sub

	Private Sub Class_Terminate()

    End Sub

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

'// STAFF 이름 접수
public Function getStaffUserName(uid)
	if uid="" or isNull(uid) then
		exit Function
	end if

	Dim strSql
	strSql = "Select top 1 username From db_partner.dbo.tbl_user_tenbyten Where userid='" & uid & "'"
	rsget.Open strSql, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		getStaffUserName = rsget("username")
	end if
	rsget.Close
End Function
%>
