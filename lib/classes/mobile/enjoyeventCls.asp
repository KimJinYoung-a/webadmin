<%
'###############################################
' PageName :enjoyevent
' Discription : 사이트 메인 공지 배너 관리
' History : 2014.06.09 이종화 생성
'###############################################

Class CMainbannerItem
	public fidx
	Public Fimg1
	Public Fimg2
	Public Fimg3
	Public Fimg4
	Public Fimg1alt
	Public Fimg2alt
	Public Fimg3alt
	Public Fimg4alt
	Public Fimg1url
	Public Fimg2url
	Public Fimg3url
	Public Fimg4url
	Public Fimg1text
	Public Fimg2text
	Public Fimg3text
	Public Fimg4text
	Public Fimg1sale
	Public Fimg2sale
	Public Fimg3sale
	Public Fimg4sale

	Public Fimg1sc
	Public Fimg2sc
	Public Fimg3sc
	Public Fimg4sc

	Public Fimg1stdate
	Public Fimg2stdate
	Public Fimg3stdate
	Public Fimg4stdate

	Public Fimg1eddate
	Public Fimg2eddate
	Public Fimg3eddate
	Public Fimg4eddate

	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fordertext

	Public Fxmlregdate
	
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
       
    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	
	'//admin/appmanage/today/enjoyevent/enjoy_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_enjoyevent "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
    		FOneItem.fidx				= rsget("idx")
			FOneItem.Fimg1				= staticImgUrl & "/mobile/enjoyevent" & rsget("img1")
			FOneItem.Fimg2				= staticImgUrl & "/mobile/enjoyevent" & rsget("img2")
			FOneItem.Fimg3				= staticImgUrl & "/mobile/enjoyevent" & rsget("img3")
			FOneItem.Fimg4				= staticImgUrl & "/mobile/enjoyevent" & rsget("img4")

			FOneItem.Fimg1alt			= rsget("img1alt")
			FOneItem.Fimg2alt			= rsget("img2alt")
			FOneItem.Fimg3alt			= rsget("img3alt")
			FOneItem.Fimg4alt			= rsget("img4alt")
			
			FOneItem.Fimg1url			= rsget("img1url")
			FOneItem.Fimg2url			= rsget("img2url")
			FOneItem.Fimg3url			= rsget("img3url")
			FOneItem.Fimg4url			= rsget("img4url")

			FOneItem.Fimg1text			= rsget("img1text")
			FOneItem.Fimg2text			= rsget("img2text")
			FOneItem.Fimg3text			= rsget("img3text")
			FOneItem.Fimg4text			= rsget("img4text")

			FOneItem.Fimg1sale			= rsget("img1sale")
			FOneItem.Fimg2sale			= rsget("img2sale")
			FOneItem.Fimg3sale			= rsget("img3sale")
			FOneItem.Fimg4sale			= rsget("img4sale")

			FOneItem.Fimg1stdate		= rsget("img1stdate")
			FOneItem.Fimg2stdate		= rsget("img2stdate")
			FOneItem.Fimg3stdate		= rsget("img3stdate")
			FOneItem.Fimg4stdate		= rsget("img4stdate")

			FOneItem.Fimg1eddate		= rsget("img1eddate")
			FOneItem.Fimg2eddate		= rsget("img2eddate")
			FOneItem.Fimg3eddate		= rsget("img3eddate")
			FOneItem.Fimg4eddate		= rsget("img4eddate")

			FOneItem.Fimg1sc			= rsget("img1sc")
			FOneItem.Fimg2sc			= rsget("img2sc")
			FOneItem.Fimg3sc			= rsget("img3sc")
			FOneItem.Fimg4sc			= rsget("img4sc")

			FOneItem.Fstartdate			= rsget("startdate")
			FOneItem.Fenddate			= rsget("enddate")
			FOneItem.Fadminid			= rsget("adminid")
			FOneItem.Flastadminid		= rsget("lastadminid")
			FOneItem.Fisusing			= rsget("isusing")
			FOneItem.Fordertext			= rsget("ordertext")
        end If
        
        rsget.Close
    end Sub
	
	'//admin/appmanage/today/enjoyevent/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_main_enjoyevent "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate>='" & Fsdt & " 00:00:00' and StartDate<='" & Fsdt & " 23:59:59' "

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " t.idx , t.img1 , t.img2 , t.img3 , t.img4 , t.img1alt , t.img2alt , t.img3alt , t.img4alt , t.img1url , t.img2url , t.img3url , t.img4url , t.img1text , t.img2text , t.img3text , t.img4text , t.img1sale , t.img2sale , t.img3sale , t.img4sale , t.startdate , t.enddate , t.adminid , t.lastadminid , t.isusing ,t.regdate , t.lastupdate , t.xmlregdate "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_enjoyevent as t "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate>='" & Fsdt & " 00:00:00' and StartDate<='" & Fsdt & " 23:59:59' "
        
		sqlStr = sqlStr + " order by  t.idx desc" 

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
				
				FItemList(i).fidx				= rsget("idx")
				FItemList(i).Fimg1				= staticImgUrl & "/mobile/enjoyevent" & rsget("img1")
				FItemList(i).Fimg2				= staticImgUrl & "/mobile/enjoyevent" & rsget("img2")
				FItemList(i).Fimg3				= staticImgUrl & "/mobile/enjoyevent" & rsget("img3")
				FItemList(i).Fimg4				= staticImgUrl & "/mobile/enjoyevent" & rsget("img4")
				
				FItemList(i).Fimg1alt			= rsget("img1alt")
				FItemList(i).Fimg2alt			= rsget("img2alt")
				FItemList(i).Fimg3alt			= rsget("img3alt")
				FItemList(i).Fimg4alt			= rsget("img4alt")

				FItemList(i).Fimg1url			= rsget("img1url")
				FItemList(i).Fimg2url			= rsget("img2url")
				FItemList(i).Fimg3url			= rsget("img3url")
				FItemList(i).Fimg4url			= rsget("img4url")

				FItemList(i).Fimg1text			= rsget("img1text")
				FItemList(i).Fimg2text			= rsget("img2text")
				FItemList(i).Fimg3text			= rsget("img3text")
				FItemList(i).Fimg4text			= rsget("img4text")

				FItemList(i).Fimg1sale			= rsget("img1sale")
				FItemList(i).Fimg2sale			= rsget("img2sale")
				FItemList(i).Fimg3sale			= rsget("img3sale")
				FItemList(i).Fimg4sale			= rsget("img4sale")

				FItemList(i).Fstartdate			= rsget("startdate")
				FItemList(i).Fenddate			= rsget("enddate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid		= rsget("lastadminid")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fxmlregdate		= rsget("xmlregdate")

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