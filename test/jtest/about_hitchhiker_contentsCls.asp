<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###########################################################
' Description : 히치하이커 컨텐츠 클래스
' Hieditor : 2014.07.17 유태욱 생성
'###########################################################
%>
<%
Sub DrawSelectBoxHitchhikerGubun(boxname,iselid,etcVal)
%>
<select name='<%=boxname%>' class='select' <%=etcVal%>>
	<option value="">선택하세요</option>
	<option value="1" <% if iselid = "1" then response.write " selected" %>>PC</option>
	<option value="2" <% if iselid = "2" then response.write " selected" %>>MOBILE</option>
	<option value="3" <% if iselid = "3" then response.write " selected" %>>MOVIE</option>
</select>
<%
end Sub

function getHitchhikerGubun(v)
	if v = 1 then
		getHitchhikerGubun = "PC"
	elseif v = 2 then
		getHitchhikerGubun = "MOBILE"
	else
		getHitchhikerGubun = "MOVIE"
	end if
end function

class CHitchhikerItem
	public Fidx
	public FSdate
	public FEdate
	public Fgubun
	public FIsusing
	public Fcon_title
	public FRegdate
	public FSortnum
	public Fdeviceidx
	public Fcon_detail
	public Fcontentsidx
	public FContentslink
	public FDevicename
	public FContentsSize
	public Fcon_movieurl
	public Fcon_viewthumbimg
end class

class CAbouthitchhiker
	public FItemList()
	public FDevice
	public Foneitem
	public FPageSize
	public FCurrPage
	public FTotalPage
	public Frectgubun
	public Frectdevice
	public FPageCount
	public FTotalCount
	public FScrollCount
	public FrectIsusing
	public FResultCount
	public Frectcon_title
	public Frectcontentsidx

	public Sub fnGetHitchhiker_oneitem()
	    dim sqlStr, sqlsearch

	if Frectcontentsidx <> "" Then
		sqlsearch = sqlsearch & " AND contentsidx ='"& Frectcontentsidx &"'"
	end if

	    sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " contentsidx,gubun,con_viewthumbimg,con_title,con_sdate,con_edate"
		sqlStr = sqlStr & " ,con_movieurl,con_regdate,isusing, con_detail"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_contents_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by contentsidx Desc"

	    'response.write sqlStr&"<br>"
	    rsget.Open SqlStr, dbget, 1
	    FResultCount = rsget.RecordCount

	    set FOneItem = new CHitchhikerItem

	    if Not rsget.Eof then

	    '// db2html 넣을것
		Foneitem.Fgubun = rsget("gubun")
		Foneitem.FIsusing = rsget("isusing")
		Foneitem.FSdate = rsget("con_sdate")
		Foneitem.FEdate = rsget("con_edate")
		Foneitem.FRegdate = rsget("con_regdate")
		Foneitem.Fcontentsidx = rsget("contentsidx")
		Foneitem.Fcon_title = db2html(rsget("con_title"))
		Foneitem.Fcon_detail = db2html(rsget("con_detail"))
		Foneitem.Fcon_movieurl = db2html(rsget("con_movieurl"))
		Foneitem.Fcon_viewthumbimg = rsget("con_viewthumbimg")

	    end if
	    rsget.Close
	end Sub

	public Sub getHitchhiker_oneitem()
		Dim sData, rst, i, objJson, iBody, istrParam
		istrParam = "?ContentIdx="&Frectcontentsidx
		SET objJson = CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objJson.OPEN "GET", "http://localhost:58658/api/Hitchhiker/View" & istrParam, true
			objJSON.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
			objJson.Send()

			If objJson.ReadyState <> 4 Then
				objJson.waitForResponse 150
			End If

			If objJson.Status = "200" Then
				iBody = BinaryToText(objJson.ResponseBody, "utf-8")
				Set rst = JSON.parse(iBody)
					set FOneItem = new CHitchhikerItem
						FResultCount = 1
						Foneitem.Fgubun = rst.gubun
						Foneitem.FIsusing = rst.isUsing
						Foneitem.FSdate = rst.conSDate
						Foneitem.FEdate = rst.conEDate
						Foneitem.FRegdate = rst.conRegDate
						Foneitem.Fcontentsidx = rst.contentIdx
						Foneitem.Fcon_title = db2html(rst.conTitle)
						Foneitem.Fcon_detail = db2html(rst.gubun)
						Foneitem.Fcon_movieurl = db2html(rst.conMovieURL)
						Foneitem.Fcon_viewthumbimg = rst.conviewThumbImg
				Set rst = nothing
			End If
		SET objJson = nothing
	End Sub

	Public Sub getHitchhikerList
		Dim sData, rst, i, objJson, iBody, istrParam, lst
		istrParam = "?Gubun="&Frectgubun&"&ConTitle="&Frectcon_title&"&IsUsing="&FrectIsusing&"&SPageNo="&FCurrPage&"&PageCount="&FPageSize&""
		SET objJson = CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objJson.OPEN "GET", "http://localhost:58658/api/Hitchhiker/List" & istrParam, true
			objJSON.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
			objJson.Send()

			If objJson.ReadyState <> 4 Then
				objJson.waitForResponse 150
			End If

			If objJson.Status = "200" Then
				iBody = BinaryToText(objJson.ResponseBody, "utf-8")
				Set rst = JSON.parse(iBody)
					SET lst = rst.List
						FTotalCount = rst.totalCount
						If (FCurrPage * FPageSize < FTotalCount) Then
							FResultCount = FPageSize
						Else
							FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
						End If

						FTotalPage = (FTotalCount \ FPageSize)
						If (FTotalPage<>FTotalCount / FPageSize) Then FTotalPage = FTotalPage +1
						Redim preserve FItemList(FResultCount)
						FPageCount = FCurrPage - 1
						If FResultCount > 0 Then
							For i = 0 to FResultCount - 1
								Set FItemList(i) = new CHitchhikerItem
									FItemList(i).Fgubun = lst.get(i).gubun
									FItemList(i).FIsusing = lst.get(i).isUsing
									FItemList(i).FSdate = lst.get(i).conSDate
									FItemList(i).FEdate = lst.get(i).conEDate
									FItemList(i).FRegdate = lst.get(i).conRegDate
									FItemList(i).Fcontentsidx = lst.get(i).contentsIdx
									FItemList(i).Fcon_title = lst.get(i).conTitle
									FItemList(i).Fcon_viewthumbimg = lst.get(i).conViewThumbImg
							Next
						End If
					Set lst = nothing
				Set rst = nothing
			End If
		SET objJson = nothing
	End Sub

	public sub fnGetHitchhikerList
		dim sqlStr,i, sqlsearch

		if Frectgubun <> "" Then
			sqlsearch = sqlsearch & " AND gubun ='"& Frectgubun &"'"
		end if

		if Frectcon_title <> "" Then
			sqlsearch = sqlsearch & " AND con_title like'%"& Frectcon_title &"%'"
		end if

		if FrectIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FrectIsusing &"'"
		end if

		'글의 총 갯수 구하기
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_contents_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " contentsidx,gubun,con_viewthumbimg,con_title,con_sdate,con_edate"
		sqlStr = sqlStr & " ,con_movieurl,con_regdate,isusing"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_contents_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by contentsidx Desc"

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
				set FItemList(i) = new CHitchhikerItem

					'//db2html 넣을것
					FItemList(i).Fgubun = rsget("gubun")
					FItemList(i).FIsusing = rsget("isusing")
					FItemList(i).FSdate = rsget("con_sdate")
					FItemList(i).FEdate = rsget("con_edate")
					FItemList(i).FRegdate = rsget("con_regdate")
					FItemList(i).Fcontentsidx = rsget("contentsidx")
					FItemList(i).Fcon_title = db2html(rsget("con_title"))
					FItemList(i).Fcon_viewthumbimg = rsget("con_viewthumbimg")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub fnGetDeviceList
		dim sqlStr,i, sqlsearch, hicprogbn

			if Frectgubun <> "" Then
				sqlsearch = sqlsearch & " AND gubun ='"& Frectgubun &"'"
			end if

			if Frectisusing <> "" Then
				sqlsearch = sqlsearch & " AND isusing ='"& Frectisusing &"'"
			end if

			sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
			sqlStr = sqlStr & " deviceidx,gubun,device_name,contents_size,isusing,sortnum,regdate"
			sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_device_size"
			sqlStr = sqlStr & " where 1=1 " & sqlsearch
			sqlStr = sqlStr & " order by isusing Desc, sortnum Asc"

			'response.write sqlStr &"<br>"
			rsget.pagesize = FPageSize
			rsget.Open sqlStr,dbget,1

			FTotalCount = rsget.recordcount
			FResultCount =  rsget.recordcount

			redim preserve FItemList(FResultCount)

			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.EOF
					set FItemList(i) = new CHitchhikerItem

					'//db2html 넣을것
					FItemList(i).Fgubun = rsget("gubun")
					FItemList(i).FIsusing = rsget("isusing")
					FItemList(i).FRegdate = rsget("regdate")
					FItemList(i).FDeviceidx = rsget("deviceidx")
					FItemList(i).FSortnum = db2html(rsget("sortnum"))
					FItemList(i).FDevicename = db2html(rsget("device_name"))
					FItemList(i).FContentsSize = db2html(rsget("contents_size"))

					rsget.movenext
					i=i+1
				loop
			end if
			rsget.Close
	end sub

	public sub fnGetContents_link
		dim sqlStr,i, sqlsearch, hicprogbn

			if Frectgubun <> "" Then
				sqlsearch = sqlsearch & " AND gubun ='"& Frectgubun &"'"
			end if

			if Frectisusing <> "" Then
				sqlsearch = sqlsearch & " AND D.isusing ='"& Frectisusing &"'"
			end if

			sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
			sqlStr = sqlStr & " d.deviceidx, D.device_name, D.contents_size, D.gubun, D.isusing, D.sortnum, D.regdate, K.contentslink "
			sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_device_size as D"
			sqlStr = sqlStr & "	left join db_sitemaster.dbo.tbl_hitchhiker_contents_link as K"
			sqlStr = sqlStr & "	on D.deviceidx = K.deviceidx"
				if Frectcontentsidx <> "" Then
					sqlStr = sqlStr & " AND K.contentsidx ='"& Frectcontentsidx &"'"
				end if
			sqlStr = sqlStr & " where 1=1 " & sqlsearch
			sqlStr = sqlStr & " order by D.sortnum asc"

			'response.write sqlStr &"<br>"
			rsget.pagesize = FPageSize
			rsget.Open sqlStr,dbget,1

			FTotalCount = rsget.recordcount
			FResultCount =  rsget.recordcount

			redim preserve FItemList(FResultCount)

			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.EOF
					set FItemList(i) = new CHitchhikerItem

					'//db2html 넣을것
					FItemList(i).Fgubun = rsget("gubun")
					FItemList(i).FIsusing = rsget("isusing")
					FItemList(i).FRegdate = rsget("regdate")
					FItemList(i).FDeviceidx = rsget("deviceidx")
					FItemList(i).FSortnum = db2html(rsget("sortnum"))
					FItemList(i).FContentslink = db2html(rsget("contentslink"))
					FItemList(i).FDevicename = db2html(rsget("device_name"))
					FItemList(i).FContentsSize = db2html(rsget("contents_size"))

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
end class
%>








