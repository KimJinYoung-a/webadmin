<%
'###############################################
' Discription : 모바일 keywordbanner 클래스
' History : 2013.12.16 한용민
'###############################################

Class ckeywordbanner_oneitem
	public fidx
	public fkeywordtype
	public fkeyword
	public fimagepath
	public flinkpath
	public fisusing
	public forderno
	public fregdate
	public flastdate
	public fregadminid
	public flastadminid
	public fkeywordtypename
	public fimgalt
	Public fstartdate
	Public fenddate

	Public Fxmlregdate
End Class

Class ckeywordbanner
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount
	
	Public frectisusing
	public frectidx
	Public frectdate
	
	'//admin/mobile/keywordbanner/keywordbanner_edit.asp
	Public Sub getkeywordbanner_one
		Dim sqlStr, i, sqlsearch
		
		if frectisusing<>"" then
			sqlsearch = sqlsearch & " and isusing='"& frectisusing &"'"
		end if
		if FRectIdx<>"" then
			sqlsearch = sqlsearch & " and idx="& FRectIdx &""
		end if
		
		sqlStr = "SELECT TOP 1"
		sqlStr = sqlStr & " idx, keywordtype, keyword, imagepath, linkpath, isusing, orderno, regdate"
		sqlStr = sqlStr & " , lastdate, regadminid, lastadminid, imgalt"
		sqlStr = sqlStr & " ,(case when keywordtype=1 then '이미지타입' else '키워드타입' end) as keywordtypename"		
		sqlStr = sqlStr & " ,startdate , enddate"		
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_mobile_main_keywordbanner"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
		ftotalcount = rsget.recordcount
		
        SET FOneItem = new ckeywordbanner_oneitem
	        If Not rsget.Eof then

				FOneItem.fidx			= rsget("idx")
				FOneItem.fkeywordtype			= rsget("keywordtype")
				FOneItem.fkeyword					= db2html(rsget("keyword"))
				FOneItem.fimagepath				= rsget("imagepath")
				FOneItem.flinkpath					= db2html(rsget("linkpath"))
				FOneItem.fisusing						= rsget("isusing")
				FOneItem.forderno					= rsget("orderno")
				FOneItem.fregdate					= rsget("regdate")
				FOneItem.flastdate					= rsget("lastdate")
				FOneItem.fregadminid				= rsget("regadminid")
				FOneItem.flastadminid				= rsget("lastadminid")
				FOneItem.fkeywordtypename	= rsget("keywordtypename")
				FOneItem.fimgalt						= rsget("imgalt")
				FOneItem.fstartdate					= rsget("startdate")
				FOneItem.Fenddate						= rsget("enddate")
					
        	End If
        rsget.Close
	End Sub
	
	'/mobile/keywordbanner/keywordbanner_list.asp
	Public Sub getkeywordbanner_list
		Dim sqlStr, i, sqlsearch , addSql
		
		if frectisusing<>"" then
			sqlsearch = sqlsearch & " and isusing='"& frectisusing &"'"
		end If
		
		If frectdate <> "" Then
			sqlsearch = sqlsearch & " and ('"& frectdate &"' between convert(varchar(10),startdate,120) and convert(varchar(10),enddate,120)) "
		End If 

		If frectdate <> "" Then
			addSql = addSql & " orderno ASC , regdate DESC"
		Else 
			addSql = addSql & " regdate DESC"
		End If 
		
		sqlStr = "SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_mobile_main_keywordbanner"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, keywordtype, keyword, imagepath, linkpath, isusing, orderno, regdate"
		sqlStr = sqlStr & " , lastdate, regadminid, lastadminid"
		sqlStr = sqlStr & " ,(case when keywordtype=1 then '이미지타입' else '키워드타입' end) as keywordtypename"
		sqlStr = sqlStr & " ,startdate , enddate ,xmlregdate "
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_mobile_main_keywordbanner"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		sqlStr = sqlStr & " ORDER BY " & addSql
		rsget.pagesize = FPageSize
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new ckeywordbanner_oneitem

					FItemList(i).fidx			= rsget("idx")
					FItemList(i).fkeywordtype			= rsget("keywordtype")
					FItemList(i).fkeyword				= db2html(rsget("keyword"))
					FItemList(i).fimagepath				= rsget("imagepath")
					FItemList(i).flinkpath					= db2html(rsget("linkpath"))
					FItemList(i).fisusing					= rsget("isusing")
					FItemList(i).forderno					= rsget("orderno")
					FItemList(i).fregdate					= rsget("regdate")
					FItemList(i).flastdate					= rsget("lastdate")
					FItemList(i).fregadminid			= rsget("regadminid")
					FItemList(i).flastadminid			= rsget("lastadminid")
					FItemList(i).fkeywordtypename	= rsget("keywordtypename")
					FItemList(i).fstartdate				= rsget("startdate")
					FItemList(i).fenddate					= rsget("enddate")
					FItemList(i).Fxmlregdate			= rsget("xmlregdate")

				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
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

'//키워드배너타입 select박스
Function drawSelectBoxkeywordtype(selectBoxName,selectedId,chplg)
%>
   <select name="<%=selectBoxName%>" <%= chplg %>>
	   <option value="" <% if selectedId="" then response.write "selected" %>>CHOICE</option>
	   <option value="1" <% if selectedId="1" then response.write "selected" %>>이미지타입</option>
	   <option value="2" <% if selectedId="2" then response.write "selected" %>>키워드타입</option>
   </select>
<%
End Function
%>