<%
'###########################################################
' Description :  기프트
' History : 2014.03.19 한용민 생성
' History : 2014.10.31 유태욱 mtitle 추가
'###########################################################

class Cgiftday_item
	public Fkeywordtype
	public Fkeywordidx
	public Fkeywordname
	public Fsortno
	public Fisusing
	public Fregdate
	public fmasteridx
	public ftitle
	public fmtitle
	public fstartdate
	public fenddate
	public flisttopimg_w
	public flisttopimg_m
	public fregtopimg_w
	public fregtopimg_m
	public fmainimg_W
	public fdetailidx
	public fgiftgubun
	public fuserid
	public fcontents
	public fviewcount
	public fcommentcount
	public fkeywordcount
	public fdevice
	public fimagesmall
	public FImageBasic
	public fage
	public FsubIdx		
	Public Flistidx
	Public Fitemid
	Public Fsortnum
	Public FitemName
	Public Fitemcnt
	Public FsmallImage
	public fdetailcount
	public fjoincount
	public fwinnercount
	public fuserlevel

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Cgiftday_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount

	public FRectkeywordtype
	public FRectkeywordidx
	public Frecttitle
	public Frectmtitle
	public Frectisusing
	public Frectmasteridx
	public Frectdetailidx
	Public FRectlistidx
	Public Fisusing
	public Frectuserid
	public frectorder

	'//admin/sitemaster/gift/day/giftday_edit.asp
	Public Sub getgiftday_master_one
		Dim sqlStr, i, sqlsearch
		
		if Frectmasteridx="" then exit Sub

		if Frectmasteridx<>"" then
			sqlsearch = sqlsearch & " and m.masteridx = "&Frectmasteridx&""
		end if
		
		sqlStr = "SELECT TOP 1"
		sqlStr = sqlStr & " m.masteridx, m.title, m.mtitle, m.startdate, m.enddate, m.listtopimg_w, m.listtopimg_m, m.regtopimg_w, m.regtopimg_m, m.mainimg_W"
		sqlStr = sqlStr & " , m.regdate, m.isusing"
		sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftday_master m with (nolock)"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		
		ftotalcount = rsget.recordcount
        SET FOneItem = new Cgiftday_item
	        If Not rsget.Eof then

				FOneItem.fmasteridx = rsget("masteridx")
				FOneItem.ftitle = db2html(rsget("title"))
				FOneItem.fmtitle = db2html(rsget("mtitle"))
				FOneItem.fstartdate = rsget("startdate")
				FOneItem.fenddate = rsget("enddate")
				FOneItem.flisttopimg_w = rsget("listtopimg_w")
				FOneItem.flisttopimg_m = rsget("listtopimg_m")				
				FOneItem.fregtopimg_w = rsget("regtopimg_w")
				FOneItem.fregtopimg_m = rsget("regtopimg_m")
				FOneItem.fmainimg_W = rsget("mainimg_W")
				FOneItem.fregdate = rsget("regdate")
				FOneItem.fisusing = rsget("isusing")
				
        	End If
        rsget.Close
	End Sub


	'//admin/sitemaster/gift/day/giftdaywinner.asp
	public sub getgiftday_winner()
		dim sqlStr,i, sqladd

		If Frectmasteridx <> "" Then
			sqladd = sqladd & " and d.masteridx = "&Frectmasteridx&" " 
		End If
		If Frectdetailidx <> "" Then
			sqladd = sqladd & " and d.detailidx = "&Frectdetailidx&" " 
		End If		
		If Frectuserid <> "" Then
			sqladd = sqladd & " and d.userid = '"&Frectuserid&"' " 
		End If
		If Frectisusing <> "" Then
			sqladd = sqladd & " and d.isusing = '"&Frectisusing&"' " 
		End If
		
		'총 갯수 구하기
		sqlStr = "SELECT count(*) as cnt"
		sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftday_detail d"
		sqlStr = sqlStr & " join db_board.dbo.tbl_giftday_detail_item di"
		sqlStr = sqlStr & " 	on d.detailidx=di.detailidx"
		sqlStr = sqlStr & " 	and di.isusing='Y'"
		sqlStr = sqlStr & " join db_item.dbo.tbl_item i"
		sqlStr = sqlStr & " 	on di.itemid=i.itemid"
		sqlStr = sqlStr & " join [db_user].dbo.tbl_user_n n"
		sqlStr = sqlStr & " 	on d.userid=n.userid"
		sqlStr = sqlStr & " JOIN [db_user].dbo.tbl_logindata L"
		sqlStr = sqlStr & " 	on n.userid=L.userid"
		sqlStr = sqlStr & " where 1=1 " & sqladd

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " d.detailidx, d.masteridx, d.giftgubun, d.userid, d.contents, d.viewcount, d.commentcount"
		sqlStr = sqlStr & " , d.keywordcount, d.device, d.regdate, d.isusing"
		sqlStr = sqlStr & " , di.itemid, i.smallimage, i.basicimage, i.itemname, (datediff(year,n.birthday,getdate())+1) age"
		sqlStr = sqlStr & " , L.userlevel"		
		sqlStr = sqlStr & " ,(select count(*) from db_board.dbo.tbl_giftday_detail_comment where d.userid=userid and isusing='Y') as commentcount"
		sqlStr = sqlStr & " ,(select count(*) from db_board.dbo.tbl_giftday_detail where d.userid=userid and isusing='Y') as joincount"
		sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftday_detail d"
		sqlStr = sqlStr & " join db_board.dbo.tbl_giftday_detail_item di"
		sqlStr = sqlStr & " 	on d.detailidx=di.detailidx"
		sqlStr = sqlStr & " 	and di.isusing='Y'"
		sqlStr = sqlStr & " join db_item.dbo.tbl_item i"
		sqlStr = sqlStr & " 	on di.itemid=i.itemid"
		sqlStr = sqlStr & " join [db_user].dbo.tbl_user_n n"
		sqlStr = sqlStr & " 	on d.userid=n.userid"
		sqlStr = sqlStr & " JOIN [db_user].dbo.tbl_logindata L"
		sqlStr = sqlStr & " 	on n.userid=L.userid"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd

		if frectorder="comment" then
			sqlStr = sqlStr & " ORDER BY commentcount DESC, joincount desc"
		elseif frectorder="join" then
			sqlStr = sqlStr & " ORDER BY joincount DESC, commentcount desc"
		else
			sqlStr = sqlStr & " ORDER BY d.detailidx desc"
		end if

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new Cgiftday_item
				
	            FItemList(i).fdetailidx = rsget("detailidx")
	            FItemList(i).fmasteridx = rsget("masteridx")
	            FItemList(i).fgiftgubun = rsget("giftgubun")
	            FItemList(i).fuserid = rsget("userid")
	            FItemList(i).fcontents = db2html(rsget("contents"))
	            FItemList(i).fviewcount = rsget("viewcount")
	            FItemList(i).fcommentcount = rsget("commentcount")
	            FItemList(i).fkeywordcount = rsget("keywordcount")
	            FItemList(i).fdevice = rsget("device")
	            FItemList(i).fregdate = rsget("regdate")
	            FItemList(i).fisusing = rsget("isusing")
	            FItemList(i).fitemid		= rsget("itemid")
	            FItemList(i).fitemname = rsget("itemname")
	            FItemList(i).fimagesmall = "http://webimage.10x10.co.kr/image/small/"&GetImageSubFolderByItemid(rsget("itemid"))&"/"&rsget("smallimage")
				FItemList(i).FImageBasic = "http://webimage.10x10.co.kr/image/basic/"&GetImageSubFolderByItemid(rsget("itemid"))&"/"&rsget("basicimage")
				FItemList(i).fage = rsget("age")
				FItemList(i).fuserlevel = rsget("userlevel")
				FItemList(i).fcommentcount = rsget("commentcount")
				FItemList(i).fjoincount = rsget("joincount")
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/admin/sitemaster/gift/day/giftday.asp
	public sub getgiftday_master()
		dim sqlStr,i, sqladd

		If Frectmasteridx <> "" Then
			sqladd = sqladd & " and m.masteridx = "&Frectmasteridx&" " 
		End If
		If Frecttitle <> "" Then
			sqladd = sqladd & " and m.title = '"&Frecttitle&"' " 
		End If
		If Frectmtitle <> "" Then
			sqladd = sqladd & " and m.mtitle = '"&Frectmtitle&"' " 
		End If
		If Frectisusing <> "" Then
			sqladd = sqladd & " and m.isusing = '"&Frectisusing&"' " 
		End If
		
		'총 갯수 구하기
		sqlStr = "SELECT count(*) as cnt"
		sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftday_master m with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqladd

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.masteridx, m.title, m.mtitle, m.startdate, m.enddate, m.listtopimg_w, m.listtopimg_m, m.regtopimg_w, m.regtopimg_m, m.regdate, m.isusing"
		sqlStr = sqlStr & " ,(select count(*) from db_board.dbo.tbl_giftday_detail with (nolock) where m.masteridx = masteridx and isusing='Y') as detailcount"
		sqlStr = sqlStr & " , isnull((select count(itemid) from db_board.dbo.tbl_giftday_master_item as S with (nolock) where S.masteridx = m.masteridx and S.isusing = 'Y'),0) as itemcnt "
		sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftday_master m with (nolock)"		
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY m.masteridx DESC"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new Cgiftday_item

					FItemList(i).fmasteridx = rsget("masteridx")
					FItemList(i).ftitle = db2html(rsget("title"))
					FItemList(i).fmtitle = db2html(rsget("mtitle"))
					FItemList(i).fstartdate = rsget("startdate")
					FItemList(i).fenddate = rsget("enddate")
					FItemList(i).flisttopimg_w = rsget("listtopimg_w")
					FItemList(i).flisttopimg_m = rsget("listtopimg_m")
					FItemList(i).fregtopimg_w = rsget("regtopimg_w")
					FItemList(i).fregtopimg_m = rsget("regtopimg_m")
					FItemList(i).fregdate = rsget("regdate")
					FItemList(i).fisusing = rsget("isusing")
					FItemList(i).fdetailcount = rsget("detailcount")
	                FItemList(i).Fitemcnt	= rsget("itemcnt")
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/sitemaster/gift/day/keyword/giftday_keyword.asp
	public Function getgiftdaykeywordList
		Dim strSql, i, subSql

		if keywordtype="" then exit Function

		strSql = "SELECT top " & cstr(FPageSize*FCurrpage)
		strSql = strSql & " keywordtype, keywordidx, keywordname, sortno, isusing, regdate"
		strSql = strSql & " FROM db_board.dbo.tbl_gift_keyword with (nolock)"		
		strSql = strSql & " WHERE keywordtype='"& keywordtype &"' " & subSql
		strSql = strSql & " ORDER BY sortno ASC, keywordidx desc"
		
		'response.write strSql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

		if  not rsget.EOF  then
			getgiftdaykeywordList = rsget.getRows()
		end if
		rsget.close
		
	End Function
	
	'//admin/sitemaster/gift/day/keyword/giftday_keyword.asp
	public Function getgiftdaykeywordDetail
		Dim strSql, i, subSql
		
		if FRectkeywordidx="" or keywordtype="" then exit Function

		strSql = "SELECT top 1"
		strSql = strSql & " keywordtype, keywordidx, keywordname, sortno, isusing, regdate"
		strSql = strSql & " FROM db_board.dbo.tbl_gift_keyword with (nolock)"
		strSql = strSql & " WHERE keywordtype='"& keywordtype &"' and keywordidx = '" & FRectkeywordidx & "'"
		
		'response.write strSql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

		if  not rsget.EOF  then
			set FOneItem = new Cgiftday_item

			FOneItem.Fkeywordtype = rsget("keywordtype")
			FOneItem.Fkeywordidx = rsget("keywordidx")
			FOneItem.Fkeywordname = db2html(rsget("keywordname"))
			FOneItem.Fsortno = rsget("sortno")
			FOneItem.Fisusing = rsget("isusing")
			FOneItem.Fregdate = rsget("regdate")

		end if
		rsget.close
		
	End Function

	'//admin/sitemaster/gift/day/item_isert.asp
    public Sub GetContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " select count(masteridx) as cnt from db_board.dbo.tbl_giftday_master_item "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr & " and  masteridx='" & FRectlistidx & "'"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " s.subidx , s.masteridx , s.itemid , s.isusing as itemusing , s.sortnum , i.itemname, i.smallImage "
        sqlStr = sqlStr & "From db_board.dbo.tbl_giftday_master_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & "Where masteridx='" & FRectlistidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If
        
		sqlStr = sqlStr + " order by sortnum asc" 

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
				set FItemList(i) = new Cgiftday_item
				
				FItemList(i).FsubIdx			= rsget("subidx")
	            FItemList(i).Flistidx			= rsget("masteridx")
	            FItemList(i).Fitemid			= rsget("itemid")
	            FItemList(i).Fsortnum			= rsget("sortnum")
	            FItemList(i).FIsUsing			= rsget("itemusing")
	            FItemList(i).FitemName			= rsget("itemname")
	            FItemList(i).FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'//subitem
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "Select top 1 s.*, i.itemname, i.smallImage "
        sqlStr = sqlStr & "From db_board.dbo.tbl_giftday_master_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.Itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        SqlStr = SqlStr + " where subidx=" + CStr(FRectSubIdx)

		'rw SqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new Cgiftday_item
        if Not rsget.Eof then
            FOneItem.FsubIdx			= rsget("subidx")
            FOneItem.Flistidx			= rsget("masteridx")
            FOneItem.FItemid			= rsget("Itemid")
            FOneItem.Fsortnum			= rsget("sortnum")
            FOneItem.Fisusing			= rsget("isusing")
            FOneItem.FitemName			= rsget("itemname")
            FOneItem.FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FItemid) & "/" & rsget("smallImage"),"")
        end if
        rsget.close
	End Sub
    	    
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
%>