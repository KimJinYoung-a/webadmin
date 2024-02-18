<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<%
Dim reportidx
reportidx = requestCheckvar(Request("iridx"),10)
 
	'// 품의서에 해당하는 결제요청  리스트
	dim clseapp, arrList,intLoop
	set clseapp = new CEApproval
	clseapp.Freportidx 		= reportidx 
	arrList = clseapp.fnGetPayRequestreportList 
	set clseapp = nothing 
%>
 <%IF isArray(arrList) THEN %>
<table border=0 cellpadding=0 cellspacing=5 bgcolor="#EFEFEF" class="a"> 
<tr>
	<td>[ 품의서 : <%=reportidx%> ]
	 <table border=0 cellpadding =5 cellspacing=1 class="a"  bgcolor=#BABABA>
		<Tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td>결제요청서</td><td>결제요청금액</td><td>결제요청상태</td> 
		</tr>
	<%
	 	For intLoop = 0 To UBound(arrList,2)
	 	%>
	 	<tr bgcolor="#FFFFFF" align="center"> 
	 		<td><a href="javascript:jsGoMenuSetIdx('M02<%=arrList(1,intLoop)%>','<%=reportidx%>','<%=arrList(0,intLoop)%>');"><%=arrList(0,intLoop)%></a></td>
	 			<td align="right"><a href="javascript:jsGoMenuSetIdx('M02<%=arrList(1,intLoop)%>','<%=reportidx%>','<%=arrList(0,intLoop)%>');"><%=formatnumber(arrList(2,intLoop),0)%></a></td>
	 		<td><a href="javascript:jsGoMenuSetIdx('M02<%=arrList(1,intLoop)%>','<%=reportidx%>','<%=arrList(0,intLoop)%>');"><%=fnGetPayRequestState(arrList(1,intLoop))%></a></td>
	 	</tr>
	<%	Next
 %> 
	 </table>
	</td> 
</tr> 
<tr>
 		<td align="right"><a href="javascript:jsReSetHtm(<%=reportidx%>);">[close]</a></td>
 	</tr> 
</table>
<%	END IF%>
<!-- #include virtual="/lib/db/dbclose.asp" -->