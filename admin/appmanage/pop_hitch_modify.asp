<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/appmanage/hitchhiker.asp" -->
<%
Dim idx, hitch, mode, strSql, vol
idx = request("idx")
vol = request("vol")
mode = request("mode")

Set hitch = new Hitchhiker
	hitch.Sidx = idx

If mode = "U" Then
	Dim opendate, openstate
	opendate 	= request("opendate")
	openstate 	= request("openstate")

	hitch.SMode = mode
	hitch.Sopendate = opendate
	hitch.Sopenstate = openstate
	hitch.Svol = vol
	hitch.HitchProcess
Set hitch = nothing
	response.write "<script language = 'javascript'>alert('���� �Ǿ����ϴ�');location.replace('/admin/appmanage/pop_hitch_modify.asp?idx="&idx&"');opener.location.reload();window.close();</script>"
	response.end
Else
	hitch.HitchModify
End If

Dim TodayDate
TodayDate = FormatDate(now(), "0000-00-00")
%>
<script language = "javascript">
function HitchModify(){
	var frm = document.frmcontents;
	if(frm.opendate.value==''){
		alert('�������� �����ϼ���');
		frm.opendate.focus();
		return;
	}
	if(frm.opendate.value < "<%=TodayDate%>"){
		alert('���� ��¥���� �����Դϴ�.\n�ٽ� �������ּ���');
		frm.opendate.focus();
		return;
	}
	if(frm.openstate.value==''){
		alert('���¸� �����ϼ���');
		frm.openstate.focus();
		return;
	}
	if(confirm("���� �Ͻðڽ��ϱ�?")){
		document.getElementById("mode").value = "U";
		frm.submit();
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="pop_hitch_modify.asp" onsubmit="return false;">
<input type="hidden" name="mode" id="mode">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="vol" value="<%=vol%>">
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">Idx :</td>
    <td><%=idx%></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">������ :</td>
    <td>
		<input type="text" name="opendate" size="10" maxlength=10 readonly value="<%=Left(hitch.Sopendate,10)%>"> 00:00:00
		<a href="javascript:calendarOpen(frmcontents.opendate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">���� :</td>
    <td>
    	<select name="openstate" class="select">
    		<option value=""> -- ���� --
    		<option value="0" <%=chkiif(hitch.Sopenstate="0","selected","")%> >������
			<option value="3" <%=chkiif(hitch.Sopenstate="3","selected","")%>>devOpen
			<option value="7" <%=chkiif(hitch.Sopenstate="7","selected","")%>>����
			<option value="9" <%=chkiif(hitch.Sopenstate="9","selected","")%>>����
    	</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center" colspan=2>
    	<input type="button" value="����" onClick="HitchModify();" class="button">
    </td>
</tr>
</form>
</table>
<% Set hitch = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
