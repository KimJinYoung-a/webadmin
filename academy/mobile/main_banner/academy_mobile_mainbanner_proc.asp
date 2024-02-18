<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 핑거스 모바일 상단 메인 배너 처리
'	History		: 2016.07.29 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/academy/mobile/main_banner/academy_mobile_mainbannerCls.asp"-->

<%
dim idx, mode, isusing , gubun , sortnum , linkurl_etc , layerpopurl , con_viewthumbimg, linknum, art_text
dim startdate, enddate, loginuserid

	loginuserid		=	session("ssBctId")
	idx				=	requestcheckvar(request("idx"),10)
	mode				=	requestcheckvar(request("mode"),10)
	isusing			=	requestcheckvar(request("isusing"),2)
	sortnum			=	requestcheckvar(request("sortnum"),3)
	linknum			=	requestcheckvar(request("linknum"),10)
	art_text			=	requestcheckvar(request("art_text"),256)
	gubun				=	requestcheckvar(request("SearchGubun"),2)
	linkurl_etc		=	requestcheckvar(request("linkurl_etc"),256)
	con_viewthumbimg	=	requestCheckvar(Request("con_viewthumbimg"),150)
  	if art_text <> "" then
		if checkNotValidHTML(art_text) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if linkurl_etc <> "" then
		if checkNotValidHTML(linkurl_etc) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if con_viewthumbimg <> "" then
		if checkNotValidHTML(con_viewthumbimg) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	startdate	= requestCheckvar(Request("StartDate"),10)& " " &requestCheckvar(Request("sTm"),10)
	enddate	= requestCheckvar(Request("EndDate"),10)& " " &requestCheckvar(Request("eTm"),10)

	dim sqlstr, getdate
	if mode = "EDIT" then

		sqlstr = " update db_academy.dbo.tbl_academy_mobile_mainbanner_list set " '수정모드일때 db업데이트
		sqlstr = sqlstr & " isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,edate = '"& enddate &"' "
		sqlstr = sqlstr & " ,sdate = '"& startdate &"' "
		sqlstr = sqlstr & " ,gubun = '"& gubun &"' "
		sqlstr = sqlstr & " ,sortnum = '"& sortnum &"' "
		sqlstr = sqlstr & " ,linknum = '"& html2db(linknum) &"' "
		sqlstr = sqlstr & " ,art_text = '"& html2db(art_text) &"' "
		sqlstr = sqlstr & " ,linkurl_etc = '"& html2db(linkurl_etc) &"' "
		sqlstr = sqlstr & " ,con_viewthumbimg = '"& html2db(con_viewthumbimg) &"' "
		sqlstr = sqlstr & " where idx = "& idx &" "

		'response.write sqlstr
		dbACADEMYget.execute sqlstr

	elseif mode = "NEW" then
		sqlstr = "insert into db_academy.dbo.tbl_academy_mobile_mainbanner_list (gubun, sdate, edate, isusing , linkurl_etc, sortnum, con_viewthumbimg, regdate, linknum, art_text, makerid)" ' 신규입력모드
		sqlstr = sqlstr & " values ('" & gubun & "'  , '" & startdate & "' , '" & enddate & "' , '" & isusing & "' , '" & html2db(linkurl_etc) & "'"
		sqlstr = sqlstr & " ,'" &sortnum & "', '" & con_viewthumbimg & "' , getdate(), '" & html2db(linknum) & "', '" & html2db(art_text) & "', '" & loginuserid & "') "

'response.write sqlstr
'response.end
		dbACADEMYget.execute sqlstr
	end if
%>
<script language = "javascript">
	alert("저장되었습니다.");
	opener.location.reload();
	self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->