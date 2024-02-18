<%
'###########################################################
' Description : 酒捞厄家 疯合 包府
' Hieditor : 2013.05.15 茄侩刮 积己
'###########################################################

Class clookbook_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public Fidx
	public Flookbookgubun
	public Ftitle
	public Fcontents
	public Fimagemain
	public Fimagemain_over
	public Flinkpath
	public Fisusing
	public Fregdate
	public Flastdate
	public Fregadminid
	public Flastupdateadminid
end class

class clookbook_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frectlookbookgubun
	public frectisusing
	public FRectIdx
	
	'//admin/ithinkso/lookbook/concept/concept_list_ithinkso.asp
	public sub flookbook_concept_list()
		dim sqlStr,i, sqlsearch

		if frectlookbookgubun <> "" then
			sqlsearch = sqlsearch & " and lookbookgubun = "&frectlookbookgubun&""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"&frectisusing&"'"
		end if
		
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_lookbook_ithinkso"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub
			
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, lookbookgubun, title, contents, imagemain, imagemain_over, linkpath"
		sqlStr = sqlStr & " ,isusing, regdate, lastdate, regadminid, lastupdateadminid"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_lookbook_ithinkso"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"

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
				set FItemList(i) = new clookbook_oneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).flookbookgubun = rsget("lookbookgubun")
				FItemList(i).ftitle = db2html(rsget("title"))
				FItemList(i).fcontents = db2html(rsget("contents"))
				FItemList(i).fimagemain = rsget("imagemain")
				FItemList(i).fimagemain_over = rsget("imagemain_over")
				FItemList(i).flinkpath = rsget("linkpath")
				FItemList(i).fisusing = rsget("isusing")
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

	'//admin/ithinkso/lookbook/concept/concept_contents_ithinkso.asp
    public Sub flookbook_concept_one()		
        dim SqlStr, sqlsearch
        
        if frectidx = "" then exit Sub
		
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&""
		end if

        SqlStr = "select top 1"
		sqlStr = sqlStr & " idx, lookbookgubun, title, contents, imagemain, imagemain_over, linkpath"
		sqlStr = sqlStr & " ,isusing, regdate, lastdate, regadminid, lastupdateadminid"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_lookbook_ithinkso"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
         
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new clookbook_oneitem
        if Not rsget.Eof then

			FOneItem.fidx = rsget("idx")
			FOneItem.flookbookgubun = rsget("lookbookgubun")
			FOneItem.ftitle = db2html(rsget("title"))
			FOneItem.fcontents = db2html(rsget("contents"))
			FOneItem.fimagemain = rsget("imagemain")
			FOneItem.fimagemain_over = rsget("imagemain_over")
			FOneItem.flinkpath = rsget("linkpath")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.flastdate = rsget("lastdate")
			FOneItem.fregadminid = rsget("regadminid")
			FOneItem.flastupdateadminid = rsget("lastupdateadminid")
			
        end if
        rsget.close
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