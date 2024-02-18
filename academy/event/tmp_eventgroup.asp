<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 그룹 등록  '도메인 문제로 우회 시킴
' History : 2010.09.28 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim strUrl
	strUrl = request("strUrl")
%>

<script language="javascript">

	alert("등록되었습니다.");
	//opener.location.href='<%=strUrl%>';
	opener.location.reload(); 		
	window.close();

</script>

