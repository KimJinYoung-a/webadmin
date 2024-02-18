<%
'###########################################################
' Description : cs템플릿
' Hieditor : 이상구 생성
'			 2019.09.17 한용민 수정
'###########################################################

'// /cscenter/board/cs_template_select_process.asp
Sub SelectBoxCSTemplateGubun(byval mastergubun, selectedgubun)
   Call SelectBoxCSTemplateGubunNew(mastergubun, "gubun", selectedgubun)
End Sub

Sub SelectBoxCSTemplateGubunNew(byval mastergubun, gubunname, selectedgubun)
   dim tmp_str,query1
   %><select class="select" name="<%= gubunname %>" onchange="TnCSTemplateGubunChanged(this.options[this.selectedIndex].value);">
     <option value="">선택</option><%
   query1 = " select G.gubun,G.gubunname from db_cs.dbo.tbl_cs_template as G"
   query1 = query1 & " where G.mastergubun='" + Cstr(mastergubun) + "' and G.isusing='Y' "
   query1 = query1 & " order by G.disporder Asc"

	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedgubun) = Lcase(rsget("gubun")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("gubun")&"' "&tmp_str&">"&db2html(rsget("gubunname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

function GetMasterGubunName(mastergubun)
	Select Case mastergubun
		Case "31"
			GetMasterGubunName = "업체게시판"
		Case Else
			GetMasterGubunName = "CS접수"
	End Select
end function

Class CCSTemplateItem

	public Fidx
	Public Fmastergubun
	public Fgubun
	Public Fgubunname
	public Fcontents
	public Fdisporder
	public Fisusing
	public Fregdate
	public Flastupdate

	function GetTitle()
		dim v

		v = Split(Fcontents, "__|__")
		GetTitle = v(0)
	end function

	function GetContents()
		dim v

		GetContents = ""
		v = Split(Fcontents, "__|__")

		if (UBound(v) > 0) then
			GetContents = v(1)
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CCSTemplate
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FOneItem

	Public FRectMasterGubun
	public FRectGubun
	public FRectIdx

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub GetCSTemplateone()
		dim SqlStr

		sqlStr = "select" & vbcrlf
		sqlStr = sqlStr & " idx, mastergubun, gubun, gubunname, contents, disporder, isusing, regdate, lastupdate" & vbcrlf
		sqlStr = sqlStr & " from db_cs.dbo.tbl_cs_template with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where 1=1" & vbcrlf

		if FRectMasterGubun <> "" then
			sqlStr = sqlStr & " and mastergubun = '" & FRectMasterGubun & "'"
		end if
		if FRectGubun <> "" then
			sqlStr = sqlStr & " and gubun='" & FRectGubun & "'"
		end If
		if FRectIdx<>"" then
			sqlStr = sqlStr & " and idx='" & FRectIdx & "'"
		end If

		sqlStr = sqlStr & " order by mastergubun asc, disporder asc, idx asc" & vbcrlf
		
		'Response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount
		
		if not rsget.EOF then
			set FOneItem = new CCSTemplateItem	
			
            FOneItem.fidx = rsget("idx")
			FOneItem.fmastergubun = rsget("mastergubun")
			FOneItem.fgubun = rsget("gubun")
			FOneItem.fgubunname = rsget("gubunname")
			FOneItem.fcontents = db2html(rsget("contents"))
			FOneItem.fdisporder = rsget("disporder")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.flastupdate = rsget("lastupdate")
            					
		end if
		rsget.close
	end sub	

	public Function GetCSTemplateList()
		dim sqlStr, addSql, i

		'// ===================================================================
		addSql = addSql + " from db_cs.dbo.tbl_cs_template with (nolock)"
		addSql = addSql + " where "
		addSql = addSql + " 	1 = 1 "

		if FRectMasterGubun <> "" then
			addSql = addSql + " and mastergubun = '" + FRectMasterGubun + "'"
		end if

		if FRectGubun <> "" then
			addSql = addSql + " and gubun='" + FRectGubun + "'"
		end If

		if FRectIdx<>"" then
			addSql = addSql + " and idx='" + FRectIdx + "'"
		end If


		'// ===================================================================
		sqlStr = "select count(idx) as cnt "
		sqlStr = sqlStr + addSql

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close


		'// ===================================================================
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " * " + vbCrLf
		sqlStr = sqlStr + addSql
		sqlStr = sqlStr + " order by mastergubun, disporder, idx "
		'response.write sqlStr

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSTemplateItem

				FItemList(i).Fidx     		= rsget("idx")
				FItemList(i).Fmastergubun   = rsget("mastergubun")
				FItemList(i).Fgubun     	= rsget("gubun")
				FItemList(i).Fgubunname     = db2html(rsget("gubunname"))
				FItemList(i).Fcontents     	= db2html(rsget("contents"))
				FItemList(i).Fdisporder     = rsget("disporder")
				FItemList(i).Fisusing     	= rsget("isusing")
				FItemList(i).Fregdate     	= rsget("regdate")
				FItemList(i).Flastupdate    = rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Function

	public function GetImageFolerName(byval i)
		GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public Function GetMDSRecommendList()
		dim sqlStr,i
		sqlStr = "select count(idx) as cnt from db_cs.dbo.tbl_cs_template with (nolock)"
		sqlStr = sqlStr + " where idx <> 0"
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectGubun + "'"
		end if
		if FRectidx<>"" then
			sqlStr = sqlStr + " and idx='" + FRectidx + "'"
		end if


		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbCrLf
		sqlStr = sqlStr + " idx, gubun, contents, regdate, isusing" + vbCrLf
		sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_template with (nolock)" + vbCrLf
		sqlStr = sqlStr + " where idx <> 0" + vbCrLf
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun = '" + FRectGubun + "'" + vbCrLf
		end if
		if FRectidx<>"" then
			sqlStr = sqlStr + " and idx='" + FRectidx + "'"
		end if

		sqlStr = sqlStr + " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSTemplateItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fgubun      = rsget("gubun")
				FItemList(i).Fcontents       = db2html(rsget("contents"))
				FItemList(i).Fregdate =  rsget("regdate")
				FItemList(i).Fisusing      = rsget("isusing")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function GetQnaComplimentGubun()
		dim sqlStr,i
		sqlStr = "select count(gubun) as cnt from db_cs.dbo.tbl_cs_template with (nolock)"
		sqlStr = sqlStr + " where gubun <> ''"
		if FRectidx<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectidx + "'"
		end If
		if FRectMasterGubun<>"" then
			sqlStr = sqlStr + " and mastergubun='" + FRectMasterGubun + "'"
		end if

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbCrLf
		sqlStr = sqlStr + " gubun, gubunname" + vbCrLf
		sqlStr = sqlStr + " from db_cs.dbo.tbl_sms_gubun with (nolock)" + vbCrLf
		sqlStr = sqlStr + " where gubun <> 0" + vbCrLf
		if FRectidx<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectidx + "'"
		end If
		if FRectMasterGubun<>"" then
			sqlStr = sqlStr + " and mastergubun='" + FRectMasterGubun + "'"
		end if
		sqlStr = sqlStr + " order by gubun desc"

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSTemplateItem

				FItemList(i).Fgubun     = rsget("gubun")
				FItemList(i).Fgubunname      = rsget("gubunname")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

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
