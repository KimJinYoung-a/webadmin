<%
'###########################################################
' Description : 아이띵소 제휴 관리
' Hieditor : 2013.05.14 한용민 생성
'###########################################################

Class Ccontact_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fcontact_gubun
	public fusername
	public femail
	public fhp
	public fcountryname
	public ftitle
	public fcontents
	public fuploadfileurl
	public fisusing
	public fregdate
	
end class

class Ccontact_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public frectusername
	public frectisusing
	public frecttitle
	public frectidx

	'//admin/ithinkso/contact/contact_edit.asp
	Public Sub fcontact_one
        dim sqlStr, sqlsearch
		
		if frectidx = "" then exit Sub

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "& frectidx &""
		end if
		
		sqlStr = " select top 1"
		sqlStr = sqlStr & " idx, contact_gubun, username, email, hp, countryname, title, contents, uploadfileurl"
		sqlStr = sqlStr & " , isusing, regdate"
		sqlStr = sqlStr & " from db_board.dbo.tbl_contact_ithinkso"
		sqlStr = sqlStr & " where contact_gubun=1 " & sqlsearch

        'response.write sqlStr & "<br>"
        rsget.Open sqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        fresultcount = rsget.RecordCount
        
        set FOneItem = new Ccontact_oneitem
        
        if Not rsget.Eof then

			FOneItem.fidx = rsget("idx")
			FOneItem.fcontact_gubun = rsget("contact_gubun")
			FOneItem.fusername = db2html(rsget("username"))
			FOneItem.femail = db2html(rsget("email"))
			FOneItem.fhp = db2html(rsget("hp"))
			FOneItem.fcountryname = db2html(rsget("countryname"))
			FOneItem.ftitle = db2html(rsget("title"))
			FOneItem.fcontents = db2html(rsget("contents"))
			FOneItem.fuploadfileurl = db2html(rsget("uploadfileurl"))
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fregdate = rsget("regdate")
			           
        end if
        rsget.Close
    end Sub
	
	'//admin/ithinkso/contact/contact_list.asp
	public sub fcontact_list()
		dim sqlStr,i, sqlsearch
		
		if frectusername <> "" then
			sqlsearch = sqlsearch & " and username = '"& frectusername &"'"
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"& frectisusing &"'"
		end if
		if frecttitle <> "" then
			sqlsearch = sqlsearch & " and title like '%"& frecttitle &"%'"
		end if
		
		sqlStr = "select"
		sqlStr = sqlStr & " count(*) as cnt"
		sqlStr = sqlStr & " from db_board.dbo.tbl_contact_ithinkso"
		sqlStr = sqlStr & " where contact_gubun=1 " & sqlsearch
		
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, contact_gubun, username, email, hp, countryname, title, contents, uploadfileurl"
		sqlStr = sqlStr & " , isusing, regdate"
		sqlStr = sqlStr & " from db_board.dbo.tbl_contact_ithinkso"
		sqlStr = sqlStr & " where contact_gubun=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"

		'response.write sqlStr & "<br>"
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
				set FItemList(i) = new Ccontact_oneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fcontact_gubun = rsget("contact_gubun")
				FItemList(i).fusername = db2html(rsget("username"))
				FItemList(i).femail = db2html(rsget("email"))
				FItemList(i).fhp = db2html(rsget("hp"))
				FItemList(i).fcountryname = db2html(rsget("countryname"))
				FItemList(i).ftitle = db2html(rsget("title"))
				FItemList(i).fcontents = db2html(rsget("contents"))
				FItemList(i).fuploadfileurl = db2html(rsget("uploadfileurl"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = rsget("regdate")
																
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
%>