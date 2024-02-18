<%
'###########################################################
' Description : 텐바이텐 대량구매 사이트 게시판 클래스
' Hieditor : 2013.05.13 한용민 생성
'###########################################################

Class board_item
	Public Fbrd_subject
	Public Fbrd_hit
	Public Fbrd_regdate
	Public Fbrd_fixed
	Public Fbrd_sn
	Public Fbrd_username
	Public Fbrd_team
	Public Fbrd_type
	public fbrd_isusing
	public fuserid
	public Fbrd_content
	public flastuserid
	public fbrd_lastupdate
End Class

Class board
	Public FItemList()
	Public FcmtList()
	Public Fbrd_sn
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount
	public FOneItem
	
	public Frectdetail_search
	public Frectsearchstr
	public Frectsearch_type
	public Frectbrd_sn
	public frectbrd_isusing
	
	'//admin/wholesale/notice/board_edit.asp
	Public Sub fnBoardmodify
        dim strSql, sqlsearch
		
		if Frectbrd_sn = "" then exit Sub
		
		strSql = " select top 1"
		strSql = strSql & " b.brd_sn, B.userid, B.brd_subject, B.brd_content, B.brd_hit, B.brd_regdate"
		strSql = strSql & " , B.brd_fixed, B.brd_type, B.brd_isusing, b.lastuserid, b.brd_lastupdate"
		strSql = strSql & " from db_board.dbo.tbl_board_wholesale as B"
		strSql = strSql & " where b.brd_gubun = 1 and B.brd_sn = '"& Frectbrd_sn &"'"

        'response.write strSql & "<br>"
        rsget.Open strSql, dbget, 1
        ftotalcount = rsget.RecordCount
        fresultcount = rsget.RecordCount
        
        set FOneItem = new board_item
        
        if Not rsget.Eof then
			
			FOneItem.flastuserid 			= rsget("lastuserid")
			FOneItem.fbrd_lastupdate 			= rsget("brd_lastupdate")
			FOneItem.Fbrd_sn 			= rsget("brd_sn")
    		FOneItem.fuserid 			= rsget("userid")
    		FOneItem.Fbrd_subject	= db2html(rsget("brd_subject"))
    		FOneItem.Fbrd_content 	= db2html(rsget("brd_content"))
    		FOneItem.Fbrd_hit 		= rsget("brd_hit")
    		FOneItem.Fbrd_regdate 	= rsget("brd_regdate")
    		FOneItem.Fbrd_fixed 		= rsget("brd_fixed")
    		FOneItem.Fbrd_isusing	= rsget("brd_isusing")
    		FOneItem.Fbrd_type		= rsget("brd_type")
			           
        end if
        rsget.Close
    end Sub
	
	'/admin/wholesale/notice/board_list.asp
	public Sub fnBoardlist()
		Dim strSql, i, sqlsearch

		If Frectdetail_search = "subject" Then
			sqlsearch = sqlsearch & " and B.brd_subject like '%"&Frectsearchstr&"%' "
			
		ElseIf Frectdetail_search = "content" Then
			sqlsearch = sqlsearch & " and B.brd_content like '%"&Frectsearchstr&"%' "
			
		ElseIf Frectdetail_search = "writer" Then
			sqlsearch = sqlsearch & " and b.userid like '%"&Frectsearchstr&"%' "
		End IF
		
		If Frectsearch_type <> "" Then
			sqlsearch = sqlsearch & " and B.brd_type = '" & Frectsearch_type & "' "
		End IF
		
		if frectbrd_isusing <> "" then
			sqlsearch = sqlsearch & " and B.brd_isusing = '" & frectbrd_isusing & "' "			
		end if
		
		strSql = "select count(*) as cnt "
		strSql = strSql & " from db_board.dbo.tbl_board_wholesale as B"
		strSql = strSql & " where b.brd_gubun = 1 " & sqlsearch

		'response.write strSql & "<br>"
        rsget.Open strSql,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit Sub
			
        strSql = "select top "& Cstr(FPageSize * FCurrPage)
		strSql = strSql & " B.brd_sn, B.brd_subject, B.brd_hit, B.brd_regdate, B.brd_fixed,b.brd_type , b.brd_isusing"
		strSql = strSql & " ,B.userid"
		strSql = strSql & " from db_board.dbo.tbl_board_wholesale as B"
		strSql = strSql & " where b.brd_gubun = 1 " & sqlsearch
		strSql = strSql & " order by B.brd_fixed asc, B.brd_sn desc"
		
		'response.write strSql & "<br>"
        rsget.pagesize = FPageSize
        rsget.Open strSql,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
				Set FItemList(i) = new board_item
				
				FItemList(i).fuserid 		= rsget("userid")
				FItemList(i).fbrd_type 		= rsget("brd_type")
				FItemList(i).fbrd_isusing 		= rsget("brd_isusing")
				FItemList(i).Fbrd_sn 		= rsget("brd_sn")
				FItemList(i).Fbrd_subject	= db2html(rsget("brd_subject"))
				FItemList(i).Fbrd_hit 		= rsget("brd_hit")
				FItemList(i).Fbrd_regdate 	= rsget("brd_regdate")
				FItemList(i).Fbrd_fixed 		= rsget("brd_fixed")
                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function
	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function
	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function
End Class

Function fnBrdType(view, isAll, boxname,value, onchange)
	Dim vBody
	
	If view = "w" Then	'### write
		vBody = "<select name=" & boxname & " class=""a"" " & onchange & ">"
		If isAll = "Y" Then
			vBody = vBody & "	<option value="""">CHOICE</option>"
		End If
		vBody = vBody & "	<option value=""1"" " & CHKIIF(value="1","selected","") & ">NOTICE</option>"
		vBody = vBody & "</select>"
		
	ElseIf view = "v" Then	'### view
		Select Case value
			Case "1" : vBody = "NOTICE"
			Case Else : vBody = ""
		End Select
	End IF
	fnBrdType = vBody
End Function
%>	