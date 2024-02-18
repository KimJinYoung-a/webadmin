<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 히치하이커 컨텐츠
' Hieditor : 2014.07.17 유태욱 생성
'			 2022.07.07 한용민 수정(isms취약점보안조치)
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhiker_contentsCls.asp"-->

<%
dim contentslinkarr, deviceidxarr
dim deviceidx, contentsidx, mode, sortnum, isusing, contentslink
dim gubun, con_viewthumbimg, con_title, con_sdate, con_edate, con_movieurl, con_regdate, con_detail

	deviceidx = Request("deviceidx")
	contentslink = request("contentslink")
	mode = requestCheckvar(Request("mode"),16)
	isusing = requestCheckvar(Request("isusing"),1)
	gubun = requestCheckvar(Request("hicprogbn"),1)
	sortnum = requestCheckvar(Request("sortnum"),10)
	contentsidx = requestCheckvar(Request("contentsidx"),10)

	con_title = requestCheckvar(Request("con_title"),60)
	con_sdate = requestCheckvar(Request("con_sdate"),10)
	con_edate = requestCheckvar(Request("con_edate"),10)
	con_detail = requestCheckvar(Request("con_detail"),150)
	con_regdate = requestCheckvar(Request("con_regdate"),10)
	con_movieurl = requestCheckvar(Request("con_movieurl"),500)
	con_viewthumbimg = requestCheckvar(Request("con_viewthumbimg"),150)

	dim sqlstr, getdate, i
	if mode = "EDIT" then
		if con_title <> "" and not(isnull(con_title)) then
			con_title = ReplaceBracket(con_title)
		end If
		if con_detail <> "" and not(isnull(con_detail)) then
			con_detail = ReplaceBracket(con_detail)
		end If
		if con_movieurl <> "" and not(isnull(con_movieurl)) then
			con_movieurl = ReplaceBracket(con_movieurl)
		end If

		sqlstr = " update db_sitemaster.dbo.tbl_hitchhiker_contents_list set " '수정모드일때 db업데이트
		sqlstr = sqlstr & " gubun = '"& gubun &"' "
		sqlstr = sqlstr & " ,isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,con_sdate = '"& con_sdate &"' "
		sqlstr = sqlstr & " ,con_edate = '"& con_edate &"' "
		sqlstr = sqlstr & " ,con_title = '"& html2db(con_title) &"' "
		sqlstr = sqlstr & " ,con_detail = '"& html2db(con_detail) &"' "
		sqlstr = sqlstr & " ,con_movieurl = '"& html2db(con_movieurl) &"' "
		sqlstr = sqlstr & " ,con_viewthumbimg = '"& con_viewthumbimg &"' "
		sqlstr = sqlstr & " where contentsidx = "& contentsidx &" "
		'response.write sqlstr
		dbget.execute sqlstr

		deviceidxarr = split(deviceidx,",")
		contentslinkarr = split(contentslink,",")
		for i = 0 to ubound(contentslinkarr)

			sqlstr = " if exists(select top 1 * from db_sitemaster.dbo.tbl_hitchhiker_contents_link where contentsidx = '"& contentsidx &"' and deviceidx = '"& deviceidxarr(i)&"') "
			sqlstr = sqlstr & " update db_sitemaster.dbo.tbl_hitchhiker_contents_link set "
'			sqlstr = sqlstr & " contentslink = '"& html2db(Trim(contentslinkarr(i))) &"' "
			sqlstr = sqlstr & " contentslink = '"& Trim(contentslinkarr(i)) &"' "
			sqlstr = sqlstr & " where contentsidx = '"& contentsidx &"' and deviceidx = '"& deviceidxarr(i)&"' "
			
			sqlstr = sqlstr & " else "
			sqlstr = sqlstr & " insert into db_sitemaster.dbo.tbl_hitchhiker_contents_link (contentsidx, deviceidx, contentslink)"
'			sqlstr = sqlstr & " values (" & contentsidx & " , " & deviceidxarr(i) & " , '" & html2db(Trim(contentslinkarr(i))) & "')"
			sqlstr = sqlstr & " values (" & contentsidx & " , " & deviceidxarr(i) & " , '" & Trim(contentslinkarr(i)) & "')"
			'response.write sqlstr & "<br>"
		dbget.execute sqlstr
		next

	elseif mode = "NEW" then
		if con_title <> "" and not(isnull(con_title)) then
			con_title = ReplaceBracket(con_title)
		end If
		if con_detail <> "" and not(isnull(con_detail)) then
			con_detail = ReplaceBracket(con_detail)
		end If
		if con_movieurl <> "" and not(isnull(con_movieurl)) then
			con_movieurl = ReplaceBracket(con_movieurl)
		end If

		sqlstr = "insert into db_sitemaster.dbo.tbl_hitchhiker_contents_list (gubun, con_viewthumbimg, con_title, con_sdate, con_edate, con_movieurl, isusing, con_detail, con_regdate)"
		sqlstr = sqlstr & " values (" & gubun & ",'" & con_viewthumbimg & "' , '" & html2db(con_title) & "' , '" & con_sdate & "', '" & con_edate & "' , '" & html2db(con_movieurl) & "', '" & isusing & "' , '" & html2db(con_detail) &"' , getdate())"
		'response.write sqlstr
		dbget.execute sqlstr
		sqlstr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_hitchhiker_contents_list') as contentsidx "
		rsget.Open SqlStr, dbget, 1
		
		if Not rsget.Eof then
			contentsidx = rsget("contentsidx")
		end if
		rsget.Close

		contentslinkarr = split(contentslink,",")
		deviceidxarr = split(deviceidx,",")
		for i = 0 to ubound(contentslinkarr)
			sqlstr = " if exists(select top 1 * from db_sitemaster.dbo.tbl_hitchhiker_contents_link where contentsidx = '"& contentsidx &"' and deviceidx = '"& deviceidxarr(i)&"') "
			sqlstr = sqlstr & " update db_sitemaster.dbo.tbl_hitchhiker_contents_link set "
'			sqlstr = sqlstr & " contentslink = '"& html2db(Trim(contentslinkarr(i))) &"' "
			sqlstr = sqlstr & " contentslink = '"& Trim(contentslinkarr(i)) &"' "
			sqlstr = sqlstr & " where contentsidx = '"& contentsidx &"' and deviceidx = '"& deviceidxarr(i)&"' "
			
			'//html2db 넣을것
			sqlstr = sqlstr & " else "
			sqlstr = sqlstr & " insert into db_sitemaster.dbo.tbl_hitchhiker_contents_link (contentsidx, deviceidx, contentslink)"
'			sqlstr = sqlstr & " values (" & contentsidx & " , " & deviceidxarr(i) & " , '" & html2db(Trim(contentslinkarr(i))) & "')"
			sqlstr = sqlstr & " values (" & contentsidx & " , " & deviceidxarr(i) & " , '" & Trim(contentslinkarr(i)) & "')"
			'response.write sqlstr & "<br>"
		dbget.execute sqlstr
		next
	end if
%>

<script type='text/javascript'>
	alert("저장되었습니다.");
	opener.location.reload();
	self.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->