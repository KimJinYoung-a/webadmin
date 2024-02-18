<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim vQuery, vGroupID, vGubun, vTIdx, vCompNOchgOX
	vGroupID 		= Request("groupid")
	vGubun 			= Request("gb")
	vTIdx 			= Request("tidx")
	vCompNOchgOX 	= Request("compnochgox")
	
	
	If vTIdx = "" Then
		vQuery = "SELECT TOP 1 (SELECT username FROM [db_partner].[dbo].[tbl_user_tenbyten] WHERE userid = A.reguserid) FROM [db_partner].[dbo].[tbl_partner_temp_info] AS A WHERE (groupid = '" & vGroupID & "' OR groupid_old = '" & vGroupID & "') and status IN(1,2) "
		rsget.Open vQuery,dbget
		IF Not rsget.EOF THEN
			Response.Write "<script>alert('" & rsget(0) & " 님이 동업체의 신청한 내용건이 있습니다.\n그 건이 완료된 후 신청할 수 있습니다.');window.close();</script>"
			rsget.close()
			dbget.close()
			Response.End
		Else
			rsget.close()
		END IF
	END IF
%>
<!--DOCTYPE HTML-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body onload="javascript:window.resizeTo(<%=CHKIIF(InStr(UCase(cstr(request.ServerVariables("HTTP_USER_AGENT"))),"MSIE"),"1400,1000","1500,1000")%>);">
<table width="100%" height="100%" cellpadding="0" cellspacing="0" border="0" class="a">
<tr>
	<td width="49%" height="100%"><iframe src="/admin/member/partner/upcheinfo_edit_child1.asp?groupid=<%=vGroupID%>&gb=<%=vGubun%>&tidx=<%=vTIdx%>" name="child1" width="100%" height="100%" frameborder="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="auto"></iframe></td>
	<td width="2%">&nbsp;</td>
	<% If vCompNOchgOX = "o" Then %>
	<td width="49%" height="100%"><iframe src="/admin/member/partner/upcheinfo_edit_child_compnosearch.asp?groupid=<%=vGroupID%>&gb=<%=vGubun%>&tidx=<%=vTIdx%>" name="child2" width="100%" height="100%" frameborder="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="auto"></iframe></td>
	<% Else %>
	<td width="49%" height="100%"><iframe src="/admin/member/partner/upcheinfo_edit_child2.asp?groupid=<%=vGroupID%>&gb=<%=vGubun%>&tidx=<%=vTIdx%>" name="child2" width="100%" height="100%" frameborder="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="auto"></iframe></td>
	<% End If %>
</tr>
</table>
</body>
</html>