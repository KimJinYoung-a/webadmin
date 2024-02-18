<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN(페이지관리_프리뷰-process)
'	History		: 2014.07.09 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_previewCls.asp"-->

<%
dim idx, mode, isusing, sortnum, title, cash, mileage, preview_detail, preview_thumbimg
dim startdate, enddate
	idx = requestcheckvar(request("idx"),10)
	cash = requestcheckvar(request("cash"),10)
	mode = requestcheckvar(request("mode"),4)
	title = requestcheckvar(request("title"),150)
	mileage = requestcheckvar(request("mileage"),10)
	isusing = requestcheckvar(request("isusing"),10)
	sortnum = requestcheckvar(request("sortnum"),3)
	preview_detail = requestcheckvar(request("preview_detail"),150)
	preview_thumbimg = requestcheckvar(request("preview_thumbimg"),150)
	
	startdate	= Request("StartDate")& " " &Request("sTm")
	enddate	= Request("EndDate")& " " &Request("eTm")

	dim sqlstr, getdate
	if mode = "EDIT" then
		sqlstr = " update db_sitemaster.dbo.tbl_hitchhiker_preview_list set " '수정모드
		sqlstr = sqlstr & " isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,edate = '"& enddate &"' "
		sqlstr = sqlstr & " ,sdate = '"& startdate &"' "
		sqlstr = sqlstr & " ,sortnum = '"& sortnum &"' "
		sqlstr = sqlstr & " ,cash = '"& cash &"' "
		sqlstr = sqlstr & " ,mileage = '"& mileage &"' "
		sqlstr = sqlstr & " ,title = '"& html2db(title) &"' "
		sqlstr = sqlstr & " ,preview_detail = '"& html2db(preview_detail) &"' "
		sqlstr = sqlstr & " ,preview_thumbimg = '"& html2db(preview_thumbimg) &"' "
		sqlstr = sqlstr & " where idx = "& idx &" "

		'response.write sqlstr
		dbget.execute sqlstr
	
		response.write "<script language='javascript'>"
		response.write "	alert('저장되었습니다');"
		response.write "	location.replace('/admin/hitchhiker/preview/index.asp?menupos="&menupos&"');"
		response.write "</script>"	

	elseif mode = "NEW" then
		
		sqlstr = "insert into db_sitemaster.dbo.tbl_hitchhiker_preview_list (title, sdate, edate, isusing, preview_detail , preview_thumbimg , sortnum, cash, mileage, regdate)" ' 신규입력모드
		sqlstr = sqlstr & " values ('" & title & "'  , '" & startdate & "' , '" & enddate & "' , '" & isusing & "' , '" & html2db(preview_detail) & "' , '" & html2db(preview_thumbimg) & "'"
		sqlstr = sqlstr & " ,'" & sortnum & "', '" & cash & "' , '" & mileage & "' , getdate())"

		'response.write sqlstr
		dbget.execute sqlstr
			
		sqlStr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_hitchhiker_preview_list') as idx"
		rsget.Open sqlStr, dbget, 1

		If Not rsget.Eof then
			idx = rsget("idx")
		End If
		rsget.close

		Response.Write "<script language='javascript'>"
		Response.Write "	alert('저장되었습니다');"
		Response.Write "	location.replace('/admin/hitchhiker/preview/hitchhiker_preview_write.asp?idx="&idx&"&menupos="&menupos&"');"
		Response.Write "</script>"		
	end if
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->