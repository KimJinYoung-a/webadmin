<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<% 
dim detailidx 
dim i
dim idx,masteridx,startday ,endday,totalday ,tatedivcd,regdate, vHalfGubun
Dim empno,userid,divcd,totstartday,totendday,totalvacationday,usedvacationday ,requestedday,divcdStr
   
detailidx = Request("detailidx")
 
dim oVacation
Set oVacation = new CTenByTenVacation 
oVacation.FRectdetailIdx = detailidx 
oVacation.GetDetailOne 

idx				  	= oVacation.Fidx				      
masteridx           = oVacation.Fmasteridx              
startday            = FormatDate(oVacation.Fstartday,"0000-00-00")
endday              = FormatDate(oVacation.Fendday,"0000-00-00")                 
totalday            = oVacation.Ftotalday               
tatedivcd           = oVacation.statedivcd              
regdate             = oVacation.Fregdate                
empno               = oVacation.Fempno                  
userid              = oVacation.Fuserid                 
divcd               = oVacation.Fdivcd     
divcdStr			= oVacation.FdivcdStr             
totstartday         = oVacation.Ftotstartday            
totendday           = oVacation.Ftotendday              
totalvacationday    = oVacation.Ftotalvacationday       
usedvacationday     = oVacation.Fusedvacationday        
requestedday        = oVacation.Frequestedday
vHalfGubun	        = oVacation.Fhalfgubun
Set oVacation = nothing
%>
<html>
<head>
<title>����(�ް�) ��û����</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script> 
<script language="javascript">
function changeHalfgubun(gb)
{
	if(gb == "am" || gb == "pm")
	{
		document.frm1.halfgubun.value = gb;
		if(document.frm1.halfgubun.value == "")
		{
			alert("Error\n�ý������� ���ǹٶ�.");
			return false;
		}
		document.frm1.submit();
	}
}
</script>
</head>
<body leftmargin="5" topmargin="5">

<table width="470" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
	<tr height="25">
		<td valign="bottom" colspan=2  bgcolor="F4F4F4">
			<font color="red"><strong>����(�ް�) ��û����</strong></font>
		</td>
	</tr>
	<tr height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">idx</td>
		<td bgcolor="#FFFFFF" width="300">
			<%= idx %>
		</td>
	</tr>
	<tr height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">���� ���̵�</td>
		<td bgcolor="#FFFFFF">
			<%= userid %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF">
			<%=divcdStr%>
		</td>
	</tr>
	<tr height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">����ϼ�/���δ��/���ϼ� </td>
    	<td bgcolor="#FFFFFF">
    		<%=usedvacationday %> / <%=requestedday%> / <%=  totalvacationday%>
    	</td>
    </tr>
	<tr height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">��û�Ⱓ</td>
    	<td bgcolor="#FFFFFF">
    		<%=startday%>
    		-
    		<%=endday%> 
    	</td>
    </tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��û�ϼ�</td>
		<td bgcolor="#FFFFFF">
			<%=totalday%>
			<%
				If totalday = "0.5" Then
					If vHalfGubun = "am" Then
						Response.Write "&nbsp;[��������]"
						If userid = session("ssBctId") Then
							Response.Write "&nbsp;&nbsp;&nbsp;<input type=""button"" class=""button"" value=""���ķ� ����"" onClick=""changeHalfgubun('pm')"">"
						End If
					ElseIf vHalfGubun = "pm" Then
						Response.Write "&nbsp;[���Ĺ���]"
						If userid = session("ssBctId") Then
							Response.Write "&nbsp;&nbsp;&nbsp;<input type=""button"" class=""button"" value=""�������� ����"" onClick=""changeHalfgubun('am')"">"
						End If
					End IF
				End If
			%>
		</td>
	</tr> 
</table><br>
<center><input type="button" class="button" value="�ݱ�" onclick="self.close();"></center>
<form name="frm1" action="pop_vacation_detail_view_proc.asp" method="post" target="iframe1">
<input type="hidden" name="detailidx" value="<%=idx%>">
<input type="hidden" name="halfgubun" value="">
</form>
<iframe name="iframe1" src="about:blank" width="0" height="0" border="0"></iframe>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->