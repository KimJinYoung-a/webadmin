<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN(페이지관리->이슈영역)
'	History		: 2014.07.09 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_mainbannerCls.asp"-->

<%
dim idx, mode, isusing , gubun , sortnum , linkurl , layerpopurl , con_viewthumbimg
dim startdate, enddate
	idx = request("idx")
	mode = request("mode")
	isusing = request("isusing")
	gubun = request("SearchGubun")
	sortnum =  requestcheckvar(request("sortnum"),3)
	linkurl =  requestcheckvar(request("linkurl"),256)
	con_viewthumbimg = requestCheckvar(Request("con_viewthumbimg"),150)
	
	startdate	= Request("StartDate")& " " &Request("sTm")
	enddate	= Request("EndDate")& " " &Request("eTm")

	dim sqlstr, getdate
	if mode = "EDIT" then
		
		sqlstr = " update db_sitemaster.dbo.tbl_hitchhiker_mainbanner_list set " '수정모드일때 db업데이트
		sqlstr = sqlstr & " isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,edate = '"& enddate &"' "
		sqlstr = sqlstr & " ,sdate = '"& startdate &"' "
		sqlstr = sqlstr & " ,gubun = '"& gubun &"' "
		sqlstr = sqlstr & " ,sortnum = '"& sortnum &"' "
		sqlstr = sqlstr & " ,linkurl = '"& html2db(linkurl) &"' "
		sqlstr = sqlstr & " ,con_viewthumbimg = '"& html2db(con_viewthumbimg) &"' "
		sqlstr = sqlstr & " where idx = "& idx &" "

		'response.write sqlstr
		dbget.execute sqlstr

	elseif mode = "NEW" then
		
		'//html2db  넣을것
		sqlstr = "insert into db_sitemaster.dbo.tbl_hitchhiker_mainbanner_list (gubun, sdate, edate, isusing , linkurl, sortnum, con_viewthumbimg, regdate)" ' 신규입력모드
		sqlstr = sqlstr & " values ('" & gubun & "'  , '" & startdate & "' , '" & enddate & "' , '" & isusing & "' , '" & html2db(linkurl) & "'"
		sqlstr = sqlstr & " ,'" &sortnum & "', '" & con_viewthumbimg & "' , getdate())"

		'response.write sqlstr
		dbget.execute sqlstr
	end if
%>
<script language = "javascript">
	alert("저장되었습니다."); //저장되었습니다 라는 메시지띄움
	opener.location.reload(); //이창을 띄운 부모창을 리로드함
	self.close();			  //이창을 닫음
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->