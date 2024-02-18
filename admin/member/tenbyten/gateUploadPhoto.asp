<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 사원 이미지 등록
' History : 2011.03.16 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	dim userimage ,userimageUrl
	userimage = ReplaceRequestSpecialChar(request("sfimg"))  
	userimageUrl = ReplaceRequestSpecialChar(request("sfimgUrl"))  
%>
 <div id="dAddFile">
	<img src="<%=userimageUrl%>" width="120" height="132" style="cursor:pointer" onClick="window.open('http://www.10x10.co.kr/common/showimage.asp?img=<%=userimageUrl%>', 'imageView', 'width=10,height=10,status=no,resizable=yes,scrollbars=yes');">
	<div style="text-align:right;">
		<a href="javascript:jsFileDel('<%=userimageUrl%>')" style="font-size:10px;color:blue;">[x]</a> 
	</div>
	<input type="hidden" name="sfImg" value="<%=userimageUrl%>"> 
</div>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="javascript">
<!--
$(document).ready(function(){ 
	 var sValue = $("#dAddFile").html(); 
	 $(opener.document).find("#dFile").html(sValue);   
	 self.close();
});
//-->
</script>
 
 
  