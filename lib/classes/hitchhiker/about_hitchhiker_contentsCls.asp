<%
'###########################################################
' Description : 히치하이커 컨텐츠 클래스
' Hieditor : 2014.07.17 유태욱 생성
'			 	  2022.07.07 한용민 수정(isms취약점보안조치)
'###########################################################
%>
<%
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
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_contents_list with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by contentsidx Desc"
	
	    'response.write sqlStr&"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	    FResultCount = rsget.RecordCount
	    
	    set FOneItem = new CHitchhikerItem
	    
	    if Not rsget.Eof then

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
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_contents_list with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " contentsidx,gubun,con_viewthumbimg,con_title,con_sdate,con_edate"
		sqlStr = sqlStr & " ,con_movieurl,con_regdate,isusing"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_contents_list with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by contentsidx Desc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize		
		'response.write sqlStr &"<br>"
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

Sub DrawSelectBoxHitchhikerGubun(boxname,iselid,etcVal)
%>
<select name='<%=boxname%>' class='select' <%=etcVal%>>
	<option value="">선택하세요</option>
	<option value="1" <% if iselid = "1" then response.write " selected" %>>PC</option>
	<option value="2" <% if iselid = "2" then response.write " selected" %>>MOBILE</option>
	<option value="3" <% if iselid = "3" then response.write " selected" %>>MOVIE</option>
	<option value="4" <% if iselid = "4" then response.write " selected" %>>MOBILE배경</option>
</select>
<%
end Sub

function getHitchhikerGubun(v)
	if v = 1 then
		getHitchhikerGubun = "PC"
	elseif v = 2 then
		getHitchhikerGubun = "MOBILE"
	elseif v = 3 then
		getHitchhikerGubun = "MOVIE"
	else
		getHitchhikerGubun = "MOBILE배경"
	end if
end function
%>






	

		