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
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhikerCls.asp"-->

<%
dim idx, mode, sortnum, isusing, gubun, vol1, vol2, issueimg
dim evt_title, startdate, enddate, imghtmltext
	idx = request("idx")
	vol1 = request("vol1")
	vol2 = request("vol2")
	mode = request("mode")
	isusing = request("isusing")
	issueimg = request("issueimg")
	gubun = request("SearchGubun")
	sortnum =  requestcheckvar(request("sortnum"),5)
	evt_title = requestcheckvar(request("evt_title"),150)
	imghtmltext = request("imghtmltext")

	startdate	= Request("StartDate")& " " &Request("sTm")
	enddate	= Request("EndDate")& " " &Request("eTm")

	dim sqlstr, getdate
	if mode = "EDIT" then
		
		sqlstr = " update db_sitemaster.dbo.tbl_hitchhiker_list set " '수정모드일때 db업데이트
		sqlstr = sqlstr & " isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,edate = '"& enddate &"' "
		sqlstr = sqlstr & " ,sdate = '"& startdate &"' "
		sqlstr = sqlstr & " ,vol1 = '"& vol1 &"' "
		sqlstr = sqlstr & " ,vol2 = '"& vol2 &"' "
		sqlstr = sqlstr & " ,sortnum = '"& sortnum &"' "
		sqlstr = sqlstr & " ,issueimg = '"& issueimg &"' "
		sqlstr = sqlstr & " ,hic_title = '"& html2db(evt_title) &"' "
		sqlstr = sqlstr & " ,imghtmltext = '"& html2db(imghtmltext) &"' "
		sqlstr = sqlstr & " where idx = "& idx &" "

		'response.write sqlstr
		dbget.execute sqlstr

	elseif mode = "NEW" then
		
		'//html2db  넣을것
		sqlstr = "insert into db_sitemaster.dbo.tbl_hitchhiker_list (gubun, hic_title, sdate, edate, isusing, imghtmltext ,sortnum, vol1, vol2, issueimg, regdate)" ' 신규입력모드
		sqlstr = sqlstr & " values ('" & gubun & "' , '" & html2db(evt_title) & "' , '" & startdate & "' , '" & enddate & "' , '" & isusing & "' , '" & html2db(imghtmltext) & "'"
		sqlstr = sqlstr & " ,'" &sortnum & "', '" &vol1 & "', '" &vol2 & "', '" &issueimg & "', getdate())"

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