<%

'// /cscenter/board/cs_template_select_process.asp
Sub SelectBoxCSTemplateGubun(byval mastergubun, selectedgubun)
   Call SelectBoxCSTemplateGubunNew(mastergubun, "gubun", selectedgubun)
End Sub

Sub SelectBoxCSTemplateGubunNew(byval mastergubun, gubunname, selectedgubun)
	dim tmp_str,query1
	response.write "<select class='select' name='" & gubunname & "' onchange='TnCSTemplateGubunChanged(this.options[this.selectedIndex].value);'>"
	response.write "<option value=''>º±≈√</option>"

	query1 = " select G.gubun,G.gubunname from [db_academy].[dbo].[tbl_ACA_cs_template] as G"
	query1 = query1 & " where G.mastergubun='" + Cstr(mastergubun) + "' and G.isusing='Y' "
	query1 = query1 & " order by G.disporder Asc"
	rsACADEMYget.Open query1,dbACADEMYget,1

	if  not rsACADEMYget.EOF  then
		rsACADEMYget.Movefirst

		do until rsACADEMYget.EOF
			if Lcase(selectedgubun) = Lcase(rsACADEMYget("gubun")) then
				tmp_str = " selected"
			end if
			response.write("<option value='"&rsACADEMYget("gubun")&"' "&tmp_str&">"&db2html(rsACADEMYget("gubunname"))&"</option>")
			tmp_str = ""
			rsACADEMYget.MoveNext
		loop
	end if
	rsACADEMYget.close
	response.write("</select>")
End Sub

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

	public Function GetCSTemplateList()
		dim sqlStr, addSql, i

		'// ===================================================================
		addSql = addSql + " from [db_academy].[dbo].[tbl_ACA_cs_template] "
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

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close


		'// ===================================================================
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " * " + vbCrLf
		sqlStr = sqlStr + addSql
		sqlStr = sqlStr + " order by mastergubun, disporder, idx "
		'response.write sqlStr

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CCSTemplateItem

				FItemList(i).Fidx     		= rsACADEMYget("idx")
				FItemList(i).Fmastergubun   = rsACADEMYget("mastergubun")
				FItemList(i).Fgubun     	= rsACADEMYget("gubun")
				FItemList(i).Fgubunname     = db2html(rsACADEMYget("gubunname"))
				FItemList(i).Fcontents     	= db2html(rsACADEMYget("contents"))
				FItemList(i).Fdisporder     = rsACADEMYget("disporder")
				FItemList(i).Fisusing     	= rsACADEMYget("isusing")
				FItemList(i).Fregdate     	= rsACADEMYget("regdate")
				FItemList(i).Flastupdate    = rsACADEMYget("lastupdate")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end Function

	public function GetImageFolerName(byval i)
		GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public Function GetMDSRecommendList()
		dim sqlStr,i
		sqlStr = "select count(idx) as cnt from [db_academy].[dbo].[tbl_ACA_cs_template]"
		sqlStr = sqlStr + " where idx <> 0"
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectGubun + "'"
		end if
		if FRectidx<>"" then
			sqlStr = sqlStr + " and idx='" + FRectidx + "'"
		end if


		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbCrLf
		sqlStr = sqlStr + " idx, gubun, contents, regdate, isusing" + vbCrLf
		sqlStr = sqlStr + " from [db_academy].[dbo].[tbl_ACA_cs_template]" + vbCrLf
		sqlStr = sqlStr + " where idx <> 0" + vbCrLf
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun = '" + FRectGubun + "'" + vbCrLf
		end if
		if FRectidx<>"" then
			sqlStr = sqlStr + " and idx='" + FRectidx + "'"
		end if

		sqlStr = sqlStr + " order by idx desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CCSTemplateItem

				FItemList(i).Fidx     = rsACADEMYget("idx")
				FItemList(i).Fgubun      = rsACADEMYget("gubun")
				FItemList(i).Fcontents       = db2html(rsACADEMYget("contents"))
				FItemList(i).Fregdate =  rsACADEMYget("regdate")
				FItemList(i).Fisusing      = rsACADEMYget("isusing")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end function

	public Function GetQnaComplimentGubun()
		dim sqlStr,i
		sqlStr = "select count(gubun) as cnt from [db_academy].[dbo].[tbl_ACA_cs_template]"
		sqlStr = sqlStr + " where gubun <> ''"
		if FRectidx<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectidx + "'"
		end If
		if FRectMasterGubun<>"" then
			sqlStr = sqlStr + " and mastergubun='" + FRectMasterGubun + "'"
		end if

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbCrLf
		sqlStr = sqlStr + " gubun, gubunname" + vbCrLf
		sqlStr = sqlStr + " from db_cs.dbo.tbl_sms_gubun" + vbCrLf
		sqlStr = sqlStr + " where gubun <> 0" + vbCrLf
		if FRectidx<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectidx + "'"
		end If
		if FRectMasterGubun<>"" then
			sqlStr = sqlStr + " and mastergubun='" + FRectMasterGubun + "'"
		end if
		sqlStr = sqlStr + " order by gubun desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CCSTemplateItem

				FItemList(i).Fgubun     = rsACADEMYget("gubun")
				FItemList(i).Fgubunname      = rsACADEMYget("gubunname")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
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
