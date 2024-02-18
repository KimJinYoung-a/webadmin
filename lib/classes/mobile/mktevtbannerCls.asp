<%
'###############################################
' PageName : mktbanner
' Discription : 마케팅 배너 - 모바일 메인 
' History : 2015-01-07 이종화 생성
'###############################################

Class CEvtMktbannerItem
	Public Fidx
	Public Fgubun
	Public Fmktimg
	Public Fm_eventid
	Public Fa_eventid
	Public Fstartdate
	Public Fenddate
	Public Fisusing
	Public Fregdate
	Public Fadminid
	Public Flastadminid
	Public Fsortnum
	Public Flastupdate
	Public Faltname
	Public Ftopfixed
	Public Fevtgubun	''2016-04-28 유태욱 추가(기획전1, 마케팅2 이벤트 구분)
	
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CEvtMktbanner
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	Public FRectvaliddate
	Public FRectgubun
       
    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	
	'//admin/mobile/mktissuebanner/mktban_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 t.* "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_event_mktbanner as t "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CEvtMktbannerItem
        
        if Not rsget.Eof then
			FOneItem.Fidx			= rsget("idx")
			FOneItem.Fgubun			= rsget("gubun")
			FOneItem.Fmktimg		= staticImgUrl & "/mobile/enjoyevent" & rsget("mktimg")
			FOneItem.Fm_eventid		= rsget("m_eventid")
			FOneItem.Fa_eventid		= rsget("a_eventid")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fsortnum		= rsget("sortnum")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.Faltname		= rsget("altname")
			FOneItem.Ftopfixed		= rsget("topfixed")
			FOneItem.Fevtgubun		= rsget("evtgubun")	''2016-04-28 유태욱 추가(기획전1, 마케팅2 이벤트 구분)

        end If
        
        rsget.Close
    end Sub
	
	'//admin/mobile/mktissuebanner/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_event_mktbanner "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate>='" & Fsdt & " 00:00:00' and StartDate<='" & Fsdt & " 23:59:59' "

		If FRectvaliddate = "on" Then 
			sqlStr = sqlStr + " and enddate > getdate() "
		End If 

		If FRectgubun <> "" Then
			sqlStr = sqlStr + " and gubun = " + FRectgubun + ""
		End If 

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr + " t.idx , t.gubun , t.m_eventid , t.a_eventid  "
		sqlStr = sqlStr + ", t.regdate ,  t.startdate , t.enddate "
		sqlStr = sqlStr + ", t.adminid , t.lastadminid , t.lastupdate , t.sortnum , t.isusing , t.mktimg , t.altname , t.topfixed, t.evtgubun"
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_event_mktbanner as t "
        sqlStr = sqlStr + " where 1=1"

		'Response.write sqlStr

        if Fisusing<>"" then
            sqlStr = sqlStr + " and t.isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and t.StartDate>='" & Fsdt & " 00:00:00' and t.StartDate<='" & Fsdt & " 23:59:59' "
        
		If FRectvaliddate = "on" Then 
		sqlStr = sqlStr + " and t.enddate > getdate() "
		End If 

		If FRectgubun <> "" Then
			sqlStr = sqlStr + " and t.gubun = " + FRectgubun + ""
		End If 

		sqlStr = sqlStr + " order by t.topfixed desc , t.startdate asc , t.sortnum asc " 

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
			set FItemList(i) = new CEvtMktbannerItem
			
			FItemList(i).Fidx			= rsget("idx")
			FItemList(i).Fgubun			= rsget("gubun")
			FItemList(i).Fmktimg		= staticImgUrl & "/mobile/enjoyevent" & rsget("mktimg")
			FItemList(i).Fm_eventid		= rsget("m_eventid")
			FItemList(i).Fa_eventid		= rsget("a_eventid")
			FItemList(i).Fstartdate		= rsget("startdate")
			FItemList(i).Fenddate		= rsget("enddate")
			FItemList(i).Fisusing		= rsget("isusing")
			FItemList(i).Fregdate		= rsget("regdate")
			FItemList(i).Fadminid		= rsget("adminid")
			FItemList(i).Flastadminid	= rsget("lastadminid")
			FItemList(i).Fsortnum		= rsget("sortnum")
			FItemList(i).Flastupdate	= rsget("lastupdate")
			FItemList(i).Faltname		= rsget("altname")
			FItemList(i).Ftopfixed		= rsget("topfixed")
			FItemList(i).Fevtgubun		= rsget("evtgubun")	''2016-04-28 유태욱 추가(기획전1, 마케팅2 이벤트 구분)

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

'// 구분 이름 접수
Public Function getGubun(v1)
	Dim gubunname1
	Dim gubunname2

	select case v1
		case "1"
			gubunname1 = "Mobile & Apps "
		case "2"
			gubunname1 = "Mobile"
		case "3"
			gubunname1 = "Apps"
		case else
			gubunname1 = ""
	end Select

	Response.write gubunname1 
End function
%>