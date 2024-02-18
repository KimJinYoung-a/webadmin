<%
'###############################################
' PageName :todaydeal
' Discription : 투데이딜 
' History : 2014.07.01 이종화 생성
'###############################################

Class CMainbannerItem
	Public Fidx
	Public FSmallimg		
	Public Fitemurl
	Public Fitemurlmo
	Public Fdealtitle		
	Public Fstartdate		
	Public Fenddate		
	Public Fadminid		
	Public Flastadminid	
	Public Fisusing		
	Public Fregdate		
	Public Flastupdate	
	Public Fxmlregdate	
	Public Fsortnum		
	Public Fitemid		
	Public Fgubun1		
	Public Fgubun2		
	Public Flimityn
	Public Flimitno
	Public Fitemname
	
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
	Public FRectgubun
       
    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	
	'//admin/appmanage/today/todaydeal/deal_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 t.* , d.limityn , (d.limitno-d.limitsold) as limitno , d.smallimage , d.limityn  "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_todaydeal as t "
		sqlStr = sqlStr + " inner join db_item.dbo.tbl_item as d "
        sqlStr = sqlStr + " on  t.itemid = d.itemid"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
			FOneItem.Fidx			= rsget("idx")
			FOneItem.FSmallimg		= webImgUrl & "/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage")
			FOneItem.Fitemurl		= rsget("itemurl")
			FOneItem.Fdealtitle		= rsget("dealtitle")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.Fxmlregdate	= rsget("xmlregdate")
			FOneItem.Fsortnum		= rsget("sortnum")
			FOneItem.Fgubun1		= rsget("gubun1")
			FOneItem.Fgubun2		= rsget("gubun2")
			FOneItem.Flimityn		= rsget("limityn")
			FOneItem.Flimitno		= rsget("limitno")
			FOneItem.Fitemid		= rsget("itemid")
			FOneItem.Fitemname		= rsget("itemname")
			FOneItem.Fitemurlmo		= rsget("itemurlmo")
        end If
        
        rsget.Close
    end Sub
	
	'//admin/appmanage/today/todaydeal/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_main_todaydeal "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate>='" & Fsdt & " 00:00:00' and StartDate<='" & Fsdt & " 23:59:59' "

		If FRectvaliddate = "on" Then 
			sqlStr = sqlStr + " and enddate > getdate() "
		End If 

		If FRectgubun <> "" Then
			sqlStr = sqlStr + " and gubun1 = " + FRectgubun + ""
		End If 

		'response.write sqlStr &"<br>"
		'response.end
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr + " t.idx , t.xmlregdate , t.gubun1 , t.gubun2 , t.itemid "
		sqlStr = sqlStr + ", d.limityn , (d.limitno-d.limitsold) as limitno , d.smallimage "
		sqlStr = sqlStr + ", t.dealtitle , t.regdate ,  t.startdate , t.enddate "
		sqlStr = sqlStr + ", t.adminid , t.lastadminid , t.lastupdate , t.sortnum , t.isusing , t.itemurl , t.itemname , t.itemurlmo "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_todaydeal as t "
        sqlStr = sqlStr + " inner join db_item.dbo.tbl_item as d "
        sqlStr = sqlStr + " on t.itemid = d.itemid"
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
			sqlStr = sqlStr + " and t.gubun1 = " + FRectgubun + ""
		End If 

		'sqlStr = sqlStr + " order by t.startdate asc , t.sortnum asc " 
		sqlStr = sqlStr + " order by t.idx desc " 

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
			
			FItemList(i).Fidx			= rsget("idx")
			FItemList(i).FSmallimg		= webImgUrl & "/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage")
			FItemList(i).Fitemurl		= rsget("itemurl")
			FItemList(i).Fdealtitle		= rsget("dealtitle")
			FItemList(i).Fstartdate		= rsget("startdate")
			FItemList(i).Fenddate		= rsget("enddate")
			FItemList(i).Fadminid		= rsget("adminid")
			FItemList(i).Flastadminid	= rsget("lastadminid")
			FItemList(i).Fisusing		= rsget("isusing")
			FItemList(i).Fregdate		= rsget("regdate")
			FItemList(i).Flastupdate	= rsget("lastupdate")
			FItemList(i).Fxmlregdate	= rsget("xmlregdate")
			FItemList(i).Fsortnum		= rsget("sortnum")
			FItemList(i).Fitemid		= rsget("itemid")
			FItemList(i).Fgubun1		= rsget("gubun1")
			FItemList(i).Fgubun2		= rsget("gubun2")
			FItemList(i).Flimityn		= rsget("limityn")
			FItemList(i).Flimitno		= rsget("limitno")
			FItemList(i).Fitemname		= rsget("itemname")
			FItemList(i).Fitemurlmo		= rsget("itemurlmo")

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
Public Function getGubun(v1,v2)
	Dim gubunname1
	Dim gubunname2

	select case v1
		case "1"
			gubunname1 = "TIME SALE"
		case "2"
			gubunname1 = "WISH NO.1"
		case "3"
			gubunname1 = "ISSUE ITEM<br/>"
		case else
			gubunname1 = ""
	end Select

	If v2 <> "" And v1= "3" Then 
		select case v2
			case "1"
				gubunname2 = "<span style='color:red'>한정 재입고</span>"
			case "2"
				gubunname2 = "<span style='color:red'>HOT ITEM</span>"
			case "3"
				gubunname2 = "<span style='color:red'>SPECIAL EDITION</span>"
			case "4"
				gubunname2 = "<span style='color:red'>10x10 ONLY</span>"
			case else
				gubunname2 = ""
		end Select
	End If 

	Response.write gubunname1 & gubunname2
End function
%>