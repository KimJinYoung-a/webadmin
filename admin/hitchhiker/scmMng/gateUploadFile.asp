<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이미지 등록처리
' History : 2011.03.16 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	Dim  strImgUrl 
 
	strImgUrl	= requestCheckVar(Request("sImgUrl"),100)  
	 
%> 
<script type="text/javascript"> 
	alert("파일이 업로드 되었습니다.");
	opener.document.getElementById("sfimg").value = "<%=strImgUrl%>";		
	opener.document.all.dvFUrl.innerHTML = "<%=strImgUrl%>";		
	window.close(); 
</script>
 