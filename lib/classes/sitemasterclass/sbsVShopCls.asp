<%
'###############################################
' PageName : sbsvshopCls
' Discription : SBS V-SHOP 
' History : 2018-04-26 이종화 생성
'###############################################

Class sbsvshopItem
	public Fidx			'//drama idx
	Public Fposterimage
	Public Fdramatitle
	public Fisusing 

	public Flistidx		'//drama list idx
	public Fdramaidx
	public Ftitle
	public Fcontents
	public Fmainimage
	public Fvideourl
	public Fsubimage1
	public Fsubimage2
	public Fsubimage3
	public Fsubimage4
	public Fsubimage5
	Public Fkakaoshareimage

	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid

	Public Fregdate
	Public Fusername
	Public Flastupdate

	Public Fordertext

	Public FsubIdx
	Public Fsortnum
	Public FitemName
	Public FsmallImage
	Public FItemid
	Public Flinkurl


'20180731 최종원 추가
	Public FbannerIsUsing
	Public FeventCode
	Public FBannerImg
	Public FMainCopy
	Public FSubCopy
	Public FSalePer
	Public FOpenDate
	Public FCloseFate
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class sbsvshop
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
       
    public FRectIdx
    public FRectisusing
	Public Fsdt
	Public Fedt
	Public FRectSubIdx
	Public FRectlistidx

	'// 드라마 popup
    public Sub fnDramaGet()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_sbsvshop_drama "
        sqlStr = sqlStr & " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new sbsvshopItem
        
        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
    		FOneItem.Fposterimage		= rsget("posterimage")
    		FOneItem.Fdramatitle		= rsget("dramatitle")
    		FOneItem.Fisusing			= rsget("isusing")
        end If
        
        rsget.Close
    end Sub

	'// 드라마 popup list
	public Sub fnDramaListGet()
        dim sqlStr, i

		sqlStr = "select count(idx) as cnt from db_sitemaster.dbo.tbl_sbsvshop_drama "

		'rw sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr & " * "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_sbsvshop_drama "
		sqlStr = sqlStr & " order by idx desc" 

		'rw sqlStr &"<br>"
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
				set FItemList(i) = new sbsvshopItem
				
	    		FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fposterimage		= rsget("posterimage")
				FItemList(i).Fdramatitle		= rsget("dramatitle")
				FItemList(i).Fisusing			= rsget("isusing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub
	

	'// 드라마 컨텐츠
    public Sub fnDramaContentsGet()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_sbsvshop_list "
        sqlStr = sqlStr & " where listidx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new sbsvshopItem
        
        if Not rsget.Eof then
			FOneItem.Flistidx		=	rsget("listidx")    
			FOneItem.Fdramaidx		=	rsget("dramaidx")   
			FOneItem.Ftitle			=	db2html(rsget("title"))
			FOneItem.Fcontents		=	db2html(rsget("contents"))
			FOneItem.Fmainimage		=	chkiif(rsget("mainimage") <> "", staticImgUrl &"/mobile/drama/"& rsget("mainimage") , "")
			FOneItem.Fvideourl		=	rsget("videourl")   
			FOneItem.Fsubimage1		=	chkiif(rsget("subimage1") <> "", staticImgUrl &"/mobile/drama/"& rsget("subimage1") , "")
			FOneItem.Fsubimage2		=	chkiif(rsget("subimage2") <> "", staticImgUrl &"/mobile/drama/"& rsget("subimage2") , "")
			FOneItem.Fsubimage3		=	chkiif(rsget("subimage3") <> "", staticImgUrl &"/mobile/drama/"& rsget("subimage3") , "")
			FOneItem.Fsubimage4		=	chkiif(rsget("subimage4") <> "", staticImgUrl &"/mobile/drama/"& rsget("subimage4") , "")
			FOneItem.Fsubimage5		=	chkiif(rsget("subimage5") <> "", staticImgUrl &"/mobile/drama/"& rsget("subimage5") , "")
			FOneItem.Fisusing		=	rsget("isusing")    
			FOneItem.Fordertext		=	db2html(rsget("ordertext"))
			FOneItem.Fstartdate		=	rsget("startdate")  
			FOneItem.Fenddate		=	rsget("enddate")    
			FOneItem.Fkakaoshareimage=	chkiif(rsget("kakaoshareimage") <> "", staticImgUrl &"/mobile/drama/"& rsget("kakaoshareimage") , "")   
'20180731 최종원 추가
			FOneItem.FbannerIsUsing		=	db2html(rsget("banner_isusing"))
			FOneItem.FeventCode			=	rsget("evt_code")    
			FOneItem.FBannerImg			=	rsget("bannerimg")    
			FOneItem.FMainCopy			=	db2html(rsget("maincopy"))
			FOneItem.FSubCopy			=	db2html(rsget("subcopy"))
			FOneItem.FSalePer			=	rsget("saleper")    
			FOneItem.FOpenDate			=	rsget("opendate")    
			FOneItem.FCloseFate			=	rsget("closedate")    
        end If
        
        rsget.Close
    end Sub

	'// 드라마 컨텐츠 리스트
    public Sub fnDramaContentsListGet()
        dim sqlStr, i

		sqlStr = " select count(listidx) as cnt from db_sitemaster.dbo.tbl_sbsvshop_list "
		sqlStr = sqlStr & " where 1=1 "
        
        if FRectisusing<>"" then
            sqlStr = sqlStr & " and isusing ='" + CStr(FRectisusing) + "'"
        end If

		If FRectidx <> "" Then
			sqlStr = sqlStr & " and dramaidx ='" + CStr(FRectidx) + "'"
		End If 

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >='" & Fsdt & " 00:00:00' and  EndDate <='" & Fsdt & " 23:59:59' "

		'rw sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr & " * , D.dramatitle "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_sbsvshop_list "
		sqlStr = sqlStr & "	CROSS APPLY ( "
		sqlStr = sqlStr & "					SELECT dramatitle FROM db_sitemaster.dbo.tbl_sbsvshop_drama where idx = dramaidx "
		sqlStr = sqlStr & "	) D"
        sqlStr = sqlStr & " where 1=1"

        if FRectisusing <> "" then
            sqlStr = sqlStr & " and isusing ='" + CStr(FRectisusing) + "'"
        end If

		If FRectidx <> "" Then
			sqlStr = sqlStr & " and dramaidx ='" + CStr(FRectidx) + "'"
		End If 

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >='" & Fsdt & " 00:00:00' and  EndDate <='" & Fsdt & " 23:59:59' "

		sqlStr = sqlStr & " order by listidx desc" 

'		rw sqlStr &"<br>"
'		Response.end
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
				set FItemList(i) = new sbsvshopItem
				
				FItemList(i).Flistidx		=	rsget("listidx")    
				FItemList(i).Fdramaidx		=	rsget("dramaidx")   
				FItemList(i).Ftitle			=	rsget("title")      
				FItemList(i).Fcontents		=	rsget("contents")   
				FItemList(i).Fmainimage		=	rsget("mainimage")  
				FItemList(i).Fisusing		=	rsget("isusing")    
				FItemList(i).Fstartdate		=	rsget("startdate")  
				FItemList(i).Fenddate		=	rsget("enddate")    
				FItemList(i).Fregdate		=	rsget("regdate")    
				FItemList(i).Flastupdate	=	rsget("lastupdate") 
				FItemList(i).Fadminid		=	rsget("adminid")    
				FItemList(i).Flastadminid	=	rsget("lastadminid")
				FItemList(i).Fdramatitle	=	rsget("dramatitle")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	
    public Sub fnDramaContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " select count(listidx) as cnt from db_sitemaster.dbo.tbl_sbsvshop_items "
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & " and  listidx='" & FRectlistidx & "'"
        
        if FRectisusing<>"" then
            sqlStr = sqlStr & " and isusing='" + CStr(FRectisusing) + "'"
        end If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " s.subidx , s.listidx , s.itemid , s.isusing as itemusing , s.sortnum, isnull(s.itemname,i.itemname) as itemname , i.smallImage "
        sqlStr = sqlStr & "From db_sitemaster.dbo.tbl_sbsvshop_items as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & "Where listidx='" & FRectlistidx & "'"

        if FRectisusing<>"" then
            sqlStr = sqlStr & " and isusing='" + CStr(FRectisusing) + "'"
        end If

		sqlStr = sqlStr & " order by sortnum asc" 

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
				set FItemList(i) = new sbsvshopItem
				
				FItemList(i).FsubIdx					= rsget("subidx")
	            FItemList(i).Flistidx					= rsget("listidx")
	            FItemList(i).Fitemid					= rsget("itemid")
	            FItemList(i).Fsortnum					= rsget("sortnum")
	            FItemList(i).FIsUsing					= rsget("itemusing")
	            FItemList(i).FitemName					= rsget("itemname")
	            FItemList(i).FsmallImage				= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")

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

'// 드라마 리스트 select
Public Function getdramaname(selectBoxName,selectedId,v)
	Dim strSql , returnSelect , tmp_str
	strSql = "SELECT idx , dramatitle FROM db_sitemaster.dbo.tbl_sbsvshop_drama WHERE isusing = 1"
	rsget.Open strSql, dbget, 1

	returnSelect = "<select class='select' name='"& selectBoxName &"'>"
	If v = "on" Then
		returnSelect = returnSelect & "<option value='' "& chkiif(selectedId="", " selected" ,"") &">선택하세요</option>"
	End If 

	if Not(rsget.EOF or rsget.BOF) then
		
		rsget.Movefirst
		do until rsget.EOF
		   if Lcase(selectedId) = Lcase(rsget("idx")) then
			   tmp_str = " selected"
		   end if
		   returnSelect = returnSelect &"<option value='"&rsget("idx")&"' "&tmp_str&">"& rsget("dramatitle") &"</option>"
   		   tmp_str = ""
		   rsget.MoveNext
		loop

	end If

	returnSelect = returnSelect & "</select>"
	rsget.Close

	Response.write returnSelect
End Function
%>