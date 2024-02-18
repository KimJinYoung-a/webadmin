<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결재라인 등록
' History : 2011.03.16 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />

<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<%
dim sFilePath, sFileName, sType, pvWidth
  sFilePath = requestCheckVar(Request("SFP"),128) 
  sFileName = requestCheckVar(Request("SFN"),60) 	 
  sType = requestCheckVar(Request("sType"),10) 
  pvWidth = requestCheckVar(Request("pvWidth"),10) 
  if pvWidth="" then pvWidth=105
   %>
<div id="pvImg"> 
 <button type="button" onclick="jsDelimg('<%=sType%>');">X</button>
 <img src="<%=sFilePath%>" alt="" <%if pvWidth <> "" then%>style="width:<%=pvWidth%>px;"<%end if%> />
</div>
<script language="javascript">
<!--
$(document).ready(function(){ 
	var sValue = $("#pvImg").html();    		 
	 $(opener.document).find("#<%=sType%>Img").empty(); 
	 $(opener.document).find("#hid<%=sType%>").val("");

	 $(opener.document).find("#<%=sType%>Img").append(sValue);   
	 $(opener.document).find("#hid<%=sType%>").val("<%=sFilePath%>");
 self.close();
});
//-->
</script>
 


 
 