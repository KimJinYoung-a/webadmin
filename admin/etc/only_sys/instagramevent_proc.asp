<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 인스타그램 이벤트용 수동 등록 처리페이지
' Hieditor : 2016.06.23 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/only_sys/instagrameventCls.asp"-->

<%
dim mode, isusing, evt_code, instaidx
dim userid, imgurl, linkurl

	evt_code = requestCheckvar(request("evt_code"),10)
	mode = requestCheckvar(Request("mode"),5)
	instaidx = requestCheckvar(Request("contentsidx"),10)
	isusing = requestCheckvar(Request("isusing"),1)
	userid = requestCheckvar(Request("userid"),20)
	imgurl = requestCheckvar(Request("imgurl"),500)
	linkurl = requestCheckvar(Request("linkurl"),250)
	
	dim sqlstr
	if mode = "EDIT" then
		sqlstr = " update [db_temp].[dbo].[tbl_event_instagram] " &_
			" set evt_code=" & evt_code & " " &_
			" , imgurl = '" & imgurl & "' " &_
			" , userid = '" & userid & "' " &_
			" , linkurl = '" & linkurl & "' " &_
			" , isusing = '" & isusing & "' " &_
			" where idx=" & instaidx & " "
'response.write sqlstr
		dbget.execute sqlstr

	elseif mode = "NEW" then
		sqlstr = "insert into [db_temp].[dbo].[tbl_event_instagram] (evt_code, imgurl, userid, linkurl, isusing)"
		sqlstr = sqlstr & " values (" & evt_code & ",'" & imgurl & "' , '" & userid & "' , '" & linkurl & "', '" & isusing & "')"
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