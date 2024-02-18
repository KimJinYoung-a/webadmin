<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결재라인 등록
' History : 2011.03.16 정윤정  생성
'###########################################################
%>

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<%
dim sFilePath, sFileName, sType, pvWidth
  sFilePath = requestCheckVar(Request("SFP"),128) 
  sFileName = requestCheckVar(Request("SFN"),60) 	 
  sType = requestCheckVar(Request("sType"),10) 
  pvWidth = requestCheckVar(Request("pvWidth"),10) 
  if pvWidth="" then pvWidth=105
   %>
<script language="javascript">
<!--
$(document).ready(function(){
	$(opener.document).find("#hid<%=sType%>").val("<%=sFilePath%>");
    self.close();
});
//-->
</script>