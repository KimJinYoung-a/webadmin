<%
'###########################################################
' Description : 고객센터 게시판,1대1문의 답변 관리 클래스
' History : 이상구 생성
'			2021.09.10 한용민 수정
'###########################################################

Sub drawSelectBoxReplyMaster(selectBoxName, selectedId, gubunCode, masterUseYN)
	dim tmp_str,sqlStr, tmpTitle
%><select class="select" name="<%=selectBoxName%>">
	<option value=""></option>
<%
	sqlStr = " select m.idx, m.title, m.useYN from db_cs.dbo.tbl_reply_master m "
	sqlStr = sqlStr + " where 1 = 1 "
	if (gubunCode <> "") then
		sqlStr = sqlStr + " and m.gubunCode = '" + CStr(gubunCode) + "' "
	end if
	if (masterUseYN <> "") then
		sqlStr = sqlStr + " and m.useYN = '" + CStr(masterUseYN) + "' "
	end if
	sqlStr = sqlStr + " order by m.dispOrderNo, m.idx desc "
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

   if  not rsget.EOF  then
       do until rsget.EOF
		   tmp_str = ""
           if CStr(selectedId) = CStr(rsget("idx")) then
               tmp_str = " selected"
           end if
		   tmpTitle = CStr(db2html(rsget("title")))
		   if (rsget("title") = "N") then
			   tmpTitle = tmpTitle + "(사용안함)"
		   end if
           response.write("<option value='"&rsget("idx")&"' "&tmp_str&">" + tmpTitle + "</option>")

           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Class CReplyMasterItem
    public Fidx
	public FgubunCode
	public Ftitle
	public FdispOrderNo
	public FuseYN
	public Freguserid
	public Fmodiuserid
	public Fregdate
	public Flastupdate
	public fsitename

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CReplyDetailItem
    public Fidx
	public Fmasteridx
	public Ftitle
	public Fsubtitle
	public Fcontents
	public FdispOrderNo
	public FuseYN
	public Freguserid
	public Fmodiuserid
	public Fregdate
	public Flastupdate
	public fsitename
	public FgubunCode
	public FmasterUseYN

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CReply
    public FItemList()
    public FOneItem

    public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount

	public FRectGubunCode			'// 0001 : 1:1상담
	public FRectMasterUseYN
	public FRectDetailUseYN
	public FRectMasterIDX
	public FRectDetailIDX
	public FRectsitename

	' /cscenter/board/cs_replymaster_list.asp
	public Sub GetReplyMasterList
		dim sqlStr, addSqlStr, i

		addSqlStr = ""

		if (FRectGubunCode <> "") then
			addSqlStr = addSqlStr + " and m.gubunCode = '" + CStr(FRectGubunCode) + "' "
		end if
		if (FRectMasterUseYN <> "") then
			addSqlStr = addSqlStr + " and m.useYN = '" + CStr(FRectMasterUseYN) + "' "
		end if
		if FRectsitename <> "" then
			addSqlStr = addSqlStr & " and m.sitename = '" & FRectsitename & "' "
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr & " from db_cs.dbo.tbl_reply_master m with (nolock)"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " m.idx, m.gubunCode, m.title, m.dispOrderNo, m.useYN, m.reguserid, m.modiuserid, m.regdate, m.lastupdate, m.sitename"
		sqlStr = sqlStr & " from db_cs.dbo.tbl_reply_master m with (nolock)"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
		sqlStr = sqlStr + " order by m.dispOrderNo, m.idx desc"

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
				set FItemList(i) = new CReplyMasterItem

				FItemList(i).Fidx = rsget("idx")
				FItemList(i).FgubunCode = rsget("gubunCode")
				FItemList(i).Ftitle = db2html(rsget("title"))
				FItemList(i).FdispOrderNo = rsget("dispOrderNo")
				FItemList(i).FuseYN = rsget("useYN")
				FItemList(i).Freguserid = rsget("reguserid")
				FItemList(i).Fmodiuserid = rsget("modiuserid")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Flastupdate = rsget("lastupdate")
				FItemList(i).fsitename = db2html(rsget("sitename"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
    end Sub

	' /cscenter/board/cs_replymaster_view.asp
    public Sub GetReplyMasterOne()
        dim sqlStr

		if FRectMasterIDX="" or isnull(FRectMasterIDX) then exit Sub

		sqlStr = " select top 1"
		sqlStr = sqlStr & " m.idx, m.gubunCode, m.title, m.dispOrderNo, m.useYN, m.reguserid, m.modiuserid, m.regdate, m.lastupdate, m.sitename"
		sqlStr = sqlStr & " from db_cs.dbo.tbl_reply_master m with (nolock)"
		sqlStr = sqlStr & " where m.idx = " & FRectMasterIDX

        'response.write sqlStr&"<br>"
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount

        set FOneItem = new CReplyMasterItem

        if Not rsget.Eof then
			FOneItem.Fidx = rsget("idx")
			FOneItem.FgubunCode = rsget("gubunCode")
			FOneItem.Ftitle = db2html(rsget("title"))
			FOneItem.FdispOrderNo = rsget("dispOrderNo")
			FOneItem.FuseYN = rsget("useYN")
			FOneItem.Freguserid = rsget("reguserid")
			FOneItem.Fmodiuserid = rsget("modiuserid")
			FOneItem.Fregdate = rsget("regdate")
			FOneItem.Flastupdate = rsget("lastupdate")
			FOneItem.fsitename = db2html(rsget("sitename"))
        end if
        rsget.Close
    end Sub

    public Sub GetReplyMasterEmptyOne()
        dim sqlStr

        set FOneItem = new CReplyMasterItem
    end Sub

	' /cscenter/board/cs_reply_xml_response.asp
	public Sub GetReplysitenameList
		dim sqlStr, addSqlStr, i

		addSqlStr = ""

		if (FRectGubunCode <> "") then
			addSqlStr = addSqlStr + " and m.gubunCode = '" + CStr(FRectGubunCode) + "' "
		end if
		if (FRectMasterUseYN <> "") then
			addSqlStr = addSqlStr + " and m.useYN = '" + CStr(FRectMasterUseYN) + "' "
		end if
		if FRectsitename <> "" then
			addSqlStr = addSqlStr & " and m.sitename = '" & FRectsitename & "' "
		end if

		sqlStr = " select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " m.sitename"
		sqlStr = sqlStr & " from db_cs.dbo.tbl_reply_master m with (nolock)"
		sqlStr = sqlStr & " where "
		sqlStr = sqlStr & " 	1 = 1 "
		sqlStr = sqlStr & addSqlStr
		sqlStr = sqlStr & " group by m.sitename"
		sqlStr = sqlStr & " order by m.sitename asc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		' if (FCurrPage * FPageSize < FTotalCount) then
		' 	FResultCount = FPageSize
		' else
		' 	FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		' end if
		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CReplyMasterItem

				FItemList(i).fsitename = db2html(rsget("sitename"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
    end Sub

	' /cscenter/board/cs_replydetail_list.asp
	public Sub GetReplyDetailList
		dim sqlStr, addSqlStr, i

		addSqlStr = ""

		if (FRectGubunCode <> "") then
			addSqlStr = addSqlStr + " and m.gubunCode = '" + CStr(FRectGubunCode) + "' "
		end if

		if (FRectMasterIDX <> "") then
			addSqlStr = addSqlStr + " and m.idx = " + CStr(FRectMasterIDX) + " "
		end if

		if (FRectMasterUseYN <> "") then
			addSqlStr = addSqlStr + " and m.useYN = '" + CStr(FRectMasterUseYN) + "' "
		end if

		if (FRectDetailUseYN <> "") then
			addSqlStr = addSqlStr + " and d.useYN = '" + CStr(FRectDetailUseYN) + "' "
		end if
		if FRectsitename <> "" then
			addSqlStr = addSqlStr & " and m.sitename = '" & FRectsitename & "' "
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr & " from db_cs.dbo.tbl_reply_master m with (nolock)"
		sqlStr = sqlStr & " join db_cs.dbo.tbl_reply_detail d with (nolock)"
		sqlStr = sqlStr & " 	on m.idx = d.masteridx "
		sqlStr = sqlStr & " where 1 = 1 " & addSqlStr
		sqlStr = sqlStr 

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " m.title, m.gubunCode, m.useYN as masterUseYN, m.sitename"
		sqlStr = sqlStr & " , d.idx, d.masteridx, d.subtitle, d.contents, d.dispOrderNo, d.useYN, d.reguserid, d.modiuserid, d.regdate, d.lastupdate "
		sqlStr = sqlStr & " from db_cs.dbo.tbl_reply_master m with (nolock)"
		sqlStr = sqlStr & " join db_cs.dbo.tbl_reply_detail d with (nolock)"
		sqlStr = sqlStr & " 	on m.idx = d.masteridx "
		sqlStr = sqlStr & " where 1 = 1 " & addSqlStr
		sqlStr = sqlStr + " order by m.dispOrderNo, d.dispOrderNo, m.idx desc, d.idx desc "

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
				set FItemList(i) = new CReplyDetailItem

				FItemList(i).Fidx = rsget("idx")
				FItemList(i).Fmasteridx = rsget("masteridx")

				FItemList(i).Ftitle = db2html(rsget("title"))
				FItemList(i).Fsubtitle = db2html(rsget("subtitle"))
				FItemList(i).Fcontents = db2html(rsget("contents"))

				FItemList(i).FdispOrderNo = rsget("dispOrderNo")
				FItemList(i).FuseYN = rsget("useYN")
				FItemList(i).Freguserid = rsget("reguserid")
				FItemList(i).Fmodiuserid = rsget("modiuserid")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Flastupdate = rsget("lastupdate")

				FItemList(i).FgubunCode = rsget("gubunCode")
				FItemList(i).FmasterUseYN = rsget("masterUseYN")
				FItemList(i).fsitename = db2html(rsget("sitename"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
    end Sub

	public Sub GetReplyDetailOne
		dim sqlStr, addSqlStr, i

		addSqlStr = ""

		if (FRectDetailIDX <> "") then
			addSqlStr = addSqlStr + " and d.idx = " + CStr(FRectDetailIDX) + " "
		end if

		sqlStr = " select top 1 m.title, m.gubunCode, m.useYN as masterUseYN, d.idx, d.masteridx, d.subtitle, d.contents, d.dispOrderNo, d.useYN, d.reguserid, d.modiuserid, d.regdate, d.lastupdate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_cs.dbo.tbl_reply_master m "
		sqlStr = sqlStr + " 	join db_cs.dbo.tbl_reply_detail d "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		m.idx = d.masteridx "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
		sqlStr = sqlStr + " order by m.dispOrderNo, d.dispOrderNo, m.idx desc, d.idx desc "
		''response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount

        set FOneItem = new CReplyDetailItem

        if Not rsget.Eof then
			FOneItem.Fidx = rsget("idx")

			FOneItem.Fidx = rsget("idx")
			FOneItem.Fmasteridx = rsget("masteridx")

			FOneItem.Ftitle = db2html(rsget("title"))
			FOneItem.Fsubtitle = db2html(rsget("subtitle"))
			FOneItem.Fcontents = db2html(rsget("contents"))

			FOneItem.FdispOrderNo = rsget("dispOrderNo")
			FOneItem.FuseYN = rsget("useYN")
			FOneItem.Freguserid = rsget("reguserid")
			FOneItem.Fmodiuserid = rsget("modiuserid")
			FOneItem.Fregdate = rsget("regdate")
			FOneItem.Flastupdate = rsget("lastupdate")

			FOneItem.FgubunCode = rsget("gubunCode")
			FOneItem.FmasterUseYN = rsget("masterUseYN")
		end if
		rsget.Close
    end Sub

    public Sub GetReplyDetailEmptyOne()
        dim sqlStr

        set FOneItem = new CReplyDetailItem
    end Sub

    Private Sub Class_Initialize()
		ReDim FItemList(0)

		FCurrPage		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0

		FRectGubunCode = "0001"				'// 0001 : 1:1상담
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
End Class

' 자사몰 구분		' 2021.09.10 한용민 생성
function replysitename(vsitename)
    dim ret

    if vsitename="10x10" then
		ret="자사몰"
	elseif vsitename="etcmall" then
		ret="제휴몰"
	else
		ret=""
	end if
	replysitename=ret
end function

' 자사몰 구분		' 2021.09.10 한용민 생성
function Drawreplysitename(selBoxName,selVal,chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
        <option value='10x10' <% if cstr(selVal)=cstr("10x10") then response.write " selected" %> >자사몰</option>
		<option value='etcmall' <% if cstr(selVal)=cstr("etcmall") then response.write " selected" %> >제휴몰</option>
	</select>
<%
end Function

%>
