<%
'###########################################################
' Description : 아이띵소 사이트 관리
' Hieditor : 2013.05.15 한용민 생성
'###########################################################

Class csitemanager_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fcode
	public fcodetype
	public fcodename
	public fimagetype
	public fimagewidth
	public fimageheight
	public fisusing
	public fimagecount
	public fregdate
	public flastdate
	public fregadminid
	public flastupdateadminid
	public idx
	public imagepath
	public imagepath2
	public imagepath3
	public linkpath
	public image_order
	public fidx
	public fimagepath
	public fimagepath2
	public fimagepath3
	public flinkpath
	public fimage_order
end class

class csitemanager_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frectcode
	public frectisusing
	public frectidx
	
	'//admin/ithinkso/sitemanager/sitemanager_contents_ithinkso.asp
    public Sub fsitemanager_one()		
        dim SqlStr, sqlsearch
        
        if frectidx = "" then exit Sub
		
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and s.idx = "&frectidx&""
		end if

        SqlStr = "select top 1"
		sqlStr = sqlStr & " s.idx, s.code, s.imagepath, s.imagepath2, s.imagepath3, s.linkpath, s.isusing, s.image_order"
		sqlStr = sqlStr & " , s.regdate, s.lastdate, s.regadminid, s.lastupdateadminid"		
		sqlStr = sqlStr & " , c.imagetype, c.codename, c.imagewidth, c.imageheight, c.imagecount"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_sitemanager_ithinkso s"
		sqlStr = sqlStr & " join db_contents.dbo.tbl_sitemanager_code_ithinkso c"
		sqlStr = sqlStr & " 	on s.code = c.code"
		sqlStr = sqlStr & " 	and c.isusing='Y'"
		sqlStr = sqlStr & " 	and c.codetype=1"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
         
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new csitemanager_oneitem
        if Not rsget.Eof then

    		FOneItem.fimagepath3 = rsget("imagepath3")
    		FOneItem.fimagepath2 = rsget("imagepath2")
			FOneItem.fcode = rsget("code")
			FOneItem.fcodename = db2html(rsget("codename"))
			FOneItem.fimagetype = rsget("imagetype")
			FOneItem.fimagewidth = rsget("imagewidth")
			FOneItem.fimageheight = rsget("imageheight")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fidx = rsget("idx")
			FOneItem.fimagepath = db2html(rsget("imagepath"))
			FOneItem.flinkpath = db2html(rsget("linkpath"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fimagecount = rsget("imagecount") 
			FOneItem.fimage_order = rsget("image_order") 
 			FOneItem.flastdate = rsget("lastdate")
			FOneItem.fregadminid = rsget("regadminid")
			FOneItem.flastupdateadminid = rsget("lastupdateadminid")
			
        end if
        rsget.close
    end Sub
    
	'//admin/ithinkso/sitemanager/sitemanager_list_ithinkso.asp
	public sub fsitemanager_list()
		dim sqlStr,i, sqlsearch
		
		if frectcode <> "" then
			sqlsearch = sqlsearch & " and s.code = "&frectcode&""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and s.isusing = '"&frectisusing&"'"
		end if
		
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_sitemanager_ithinkso s"
		sqlStr = sqlStr & " join db_contents.dbo.tbl_sitemanager_code_ithinkso c"
		sqlStr = sqlStr & " 	on s.code = c.code"
		sqlStr = sqlStr & " 	and c.isusing='Y'"
		sqlStr = sqlStr & " 	and c.codetype=1"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub
			
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.idx, s.code, s.imagepath, s.imagepath2, s.imagepath3, s.linkpath, s.isusing, s.image_order"
		sqlStr = sqlStr & " , s.regdate, s.lastdate, s.regadminid, s.lastupdateadminid"		
		sqlStr = sqlStr & " , c.imagetype, c.codename, c.imagewidth, c.imageheight, c.imagecount"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_sitemanager_ithinkso s"
		sqlStr = sqlStr & " join db_contents.dbo.tbl_sitemanager_code_ithinkso c"
		sqlStr = sqlStr & " 	on s.code = c.code"
		sqlStr = sqlStr & " 	and c.isusing='Y'"
		sqlStr = sqlStr & " 	and c.codetype=1"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by s.idx Desc"

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
				set FItemList(i) = new csitemanager_oneitem
				
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fcode = rsget("code")
				FItemList(i).fimagepath = rsget("imagepath")
				FItemList(i).fimagepath2 = rsget("imagepath2")
				FItemList(i).fimagepath3 = rsget("imagepath3")
				FItemList(i).flinkpath = rsget("linkpath")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fimage_order = rsget("image_order")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastdate = rsget("lastdate")
				FItemList(i).fregadminid = rsget("regadminid")
				FItemList(i).flastupdateadminid = rsget("lastupdateadminid")
				FItemList(i).fcodename = db2html(rsget("codename"))
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fimagecount = rsget("imagecount")
															
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/ithinkso/sitemanager/sitemanager_list_ithinkso.asp
    public Sub fsitemanager_code_one()		
        dim SqlStr, sqlsearch
        
        if frectcode = "" then exit Sub
		
		if frectcode <> "" then
			sqlsearch = sqlsearch & " and c.code = "&frectcode&""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and c.isusing = '"&frectisusing&"'"
		end if
		
        SqlStr = "select top 1"
		sqlStr = sqlStr & " c.code,codetype, c.codename, c.imagetype, c.imagewidth, c.imageheight, c.isusing, c.imagecount"
		sqlStr = sqlStr & " , c.regdate, c.lastdate, c.regadminid, c.lastupdateadminid"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_sitemanager_code_ithinkso c"
		sqlStr = sqlStr & " where c.codetype=1 " & sqlsearch
        
        'response.write sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new csitemanager_oneitem
        if Not rsget.Eof then

            FOneItem.fcode = rsget("code")
            FOneItem.fcodetype = rsget("codetype")
            FOneItem.fcodename = db2html(rsget("codename"))
            FOneItem.fimagetype = rsget("imagetype")
            FOneItem.fimagewidth = rsget("imagewidth")
            FOneItem.fimageheight = rsget("imageheight")
            FOneItem.fisusing = rsget("isusing")
            FOneItem.fimagecount = rsget("imagecount")
            FOneItem.fregdate = rsget("regdate")
            FOneItem.flastdate = rsget("lastdate")
            FOneItem.fregadminid = rsget("regadminid")
            FOneItem.flastupdateadminid = rsget("lastupdateadminid")
                       
        end if
        rsget.close
    end Sub

	'//admin/ithinkso/sitemanager/sitemanager_code_ithinkso.asp
    public Sub fsitemanager_code_list()
		dim sqlStr, i, sqlsearch

		if frectcode <> "" then
			sqlsearch = sqlsearch & " and c.code = "&frectcode&""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and c.isusing = '"&frectisusing&"'"
		end if
		
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_sitemanager_code_ithinkso c"
		sqlStr = sqlStr & " where c.codetype=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit Sub
		
        SqlStr = "select top 1"
		sqlStr = sqlStr & " c.code,codetype, c.codename, c.imagetype, c.imagewidth, c.imageheight, c.isusing, c.imagecount"
		sqlStr = sqlStr & " , c.regdate, c.lastdate, c.regadminid, c.lastupdateadminid"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_sitemanager_code_ithinkso c"
		sqlStr = sqlStr & " where c.codetype=1 " & sqlsearch
		sqlStr = sqlStr & " order by c.lastdate desc"
		
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
				set FItemList(i) = new csitemanager_oneitem

	            FItemList(i).fcode = rsget("code")
	            FItemList(i).fcodetype = rsget("codetype")
	            FItemList(i).fcodename = db2html(rsget("codename"))
	            FItemList(i).fimagetype = rsget("imagetype")
	            FItemList(i).fimagewidth = rsget("imagewidth")
	            FItemList(i).fimageheight = rsget("imageheight")
	            FItemList(i).fisusing = rsget("isusing")
	            FItemList(i).fimagecount = rsget("imagecount")
	            FItemList(i).fregdate = rsget("regdate")
	            FItemList(i).flastdate = rsget("lastdate")
	            FItemList(i).fregadminid = rsget("regadminid")
	            FItemList(i).flastupdateadminid = rsget("lastupdateadminid")
														
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

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

'//적용구분 
function DrawsitemanagerCode(selectBoxName, selectedId, changeFlag)
	dim tmp_str,query1
	
	query1 = "select code,codename"
	query1 = query1 & " from db_contents.dbo.tbl_sitemanager_code_ithinkso where isusing = 'Y'"
	
	'response.write query1 & "<Br>"
	rsget.Open query1,dbget,1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if selectedId="" then response.write " selected"%> >CHOICE</option>
<%		
		if not rsget.EOF  then
		   do until rsget.EOF
		       if Lcase(selectedId) = Lcase(rsget("code")) then
		           tmp_str = " selected"
		       end if
		       response.write("<option value='"&rsget("code")&"' "&tmp_str&">" + db2html(rsget("codename")) + "</option>")
		       tmp_str = ""
		       rsget.MoveNext
		   loop
		end if
	rsget.close
	
	response.write("</select>")
end function
%>