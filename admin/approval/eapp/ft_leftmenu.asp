<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 재무회계전자결재 왼쪽메뉴
' History : 2011.03.14 정윤정  생성
'###########################################################  
%> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->  
<!-- #include virtual="/lib/db/dbopen.asp" -->  
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<%
Dim clsLeapp
Dim iPayRequeststate000,iPayRequeststate001,iPayRequeststate110,iPayRequeststate111,iPayRequeststate710,iPayRequeststate711,iPayRequeststate970,iPayRequeststate971,iPayRequeststate550,iPayRequeststate551
Dim iRectMenu 
iRectMenu = requestCheckvar(Request("iRM"),10)
IF iRectMenu ="" THEN iRectMenu = "F100"
set clsLeapp = new CPayRequest
clsLeapp.FadminId = session("ssBctId")
clsLeapp.fnGetLeftMenu

iPayRequeststate000	= clsLeapp.FPayRequeststate000
iPayRequeststate001 = clsLeapp.FPayRequeststate001
iPayRequeststate110 = clsLeapp.FPayRequeststate110
iPayRequeststate111 = clsLeapp.FPayRequeststate111
iPayRequeststate710 = clsLeapp.FPayRequeststate710
iPayRequeststate711 = clsLeapp.FPayRequeststate711
iPayRequeststate970 = clsLeapp.FPayRequeststate970
iPayRequeststate971 = clsLeapp.FPayRequeststate971
iPayRequeststate550 = clsLeapp.FPayRequeststate550
iPayRequeststate551 = clsLeapp.FPayRequeststate551

set clsLeapp = nothing
%> 
<!-- #include virtual="/lib/db/dbclose.asp" -->
<html>
<head>
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->
<script language="javascript">
	function jsGoMenu(iRMenu){
		parent.location.href = "/admin/approval/eapp/ft_index.asp?iRM="+iRMenu; 
	}
</script>
</head>
<body leftmargin ="0" topmargin="0"	>
<table width="100%" height="100%" align="center" cellpadding="3" cellspacing="0" class="a"   border="0">    
<tr height="150">
	<td  valign="top">결제요청서<Br>
		<table width="100%"  align="center" cellpadding="0" cellspacing="1" class="a" border="0">
		<tr>
			<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F100');"><%IF iRectMenu="F100" THEN%><font color="#4E9FC6"><b><%END IF%>결제요청전승인 (<font color="blue"><%=iPayRequeststate000%></font>/<%=iPayRequeststate001%>)</a></td>
		</tr> 
		<tr>
			<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F110');"><%IF iRectMenu="F110" THEN%><font color="#4E9FC6"><b><%END IF%>결제요청 (<font color="blue"><%=iPayRequeststate110%></font>/<%=iPayRequeststate111%>)</a></td>
		</tr>	
		<tr>	
			<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F711');"><%IF iRectMenu="F711" THEN%><font color="#4E9FC6"><b><%END IF%>결제확인(결제예정) (<font color="blue"><%=iPayRequeststate710%></font>/<%=iPayRequeststate711%>)</a></td>
		</tr>	
		<tr>	
			<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F971');"><%IF iRectMenu="F971" THEN%><font color="#4E9FC6"><b><%END IF%>결제완료 (<font color="blue"><%=iPayRequeststate970%></font>/<%=iPayRequeststate971%>)</a></td>
		</tr>	
		<tr>	
			<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F551');"><%IF iRectMenu="F551" THEN%><font color="#4E9FC6"><b><%ELSE%><font color="gray"><%END IF%>결제반려 (<font color="blue"><%=iPayRequeststate550%></font>/<%=iPayRequeststate551%>)</font></a></td>
		</tr>	 
		</table>
	</td>  
</tr>  
</table>
</body>
</html>
 
	
	
	
