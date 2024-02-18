<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �¶��� ��������-�Ǹ�ó��
' History : 2012.10.09 ���ر� ����
'			2013.01.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim i, cStatistic, vSiteName, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vIsBanPum, v6Ago
vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
vEYear		= NullFillWith(request("eyear"),Year(now))
vEMonth		= NullFillWith(request("emonth"),Month(now))

Dim strSql, arrRows
strSql = "exec [db_statistics_order].[dbo].[usp_TEN_meachul_kjy] '"& vSYear & "-" & TwoNumber(vSMonth) & "-01" &"', '"& vEYear & "-" & TwoNumber(vEMonth) & "-01" &"'"
rsSTSget.CursorLocation = adUseClient
rsSTSget.CursorType = adOpenStatic
rsSTSget.LockType = adLockOptimistic
rsSTSget.Open strSql, dbSTSget
If Not(rsSTSget.EOF or rsSTSget.BOF) Then
	arrRows = rsSTSget.getRows
End If
rsSTSget.close

rw strSql
%>
<script language="javascript">
function searchSubmit(){
	frm.submit();
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				<%
					'### ��
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To 2001 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""smonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'#############################

					'### ��
					Response.Write "<select name=""eyear"" class=""select"">"
					For i=Year(now) To 2001 Step -1 ''Year(v6MonthDate)
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""emonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"
				%>
				&nbsp;&nbsp;
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
* cjmall : 2021-03-11(��ǰ�з�����) / 2021-11-03 (�з�,ī�װ���Ī���� ��������)</br>
* 11���� : 2021-09-27���� �۾�����</br>
* lfmall : 2021-10-19(���οϷ�) / 2021-11-02(��Ͻ����ٷ� ���������� ����)</br>
* ������ũ : 2021-12-15 ��ǰ ��� ���� �Ϸ�</br>
* ������� : 2022-03-31 ��Ͻ����ٷ� ���������� ����</br>
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">���޸�</td>
	<td align="center">����</td>
    <td align="center">�ֹ���</td>
    <td align="center">�����Ѿ�</td>
    <td align="center">���ʽ���������</td>
    <td align="center">�����</td>
</tr>
<%
If isArray(arrRows) Then
	For i=0 To Ubound(arrRows, 2)
%>
<tr <%= Chkiif(arrRows(6, i)="1","bgcolor=SKYBLUE","bgcolor=#FFFFFF") %>>
	<td align="center"><%= arrRows(0, i) %></td>
	<td align="center"><%= arrRows(1, i) %></td>
	<td align="center"><%= FormatNumber(arrRows(2, i), 0) %></td>
	<td align="center"><%= FormatNumber(arrRows(3, i), 0) %></td>
	<td align="center"><%= FormatNumber(arrRows(4, i), 0) %></td>
	<td align="center"><%= FormatNumber(arrRows(5, i), 0) %></td>
</tr>
<%
	Next
End If
%>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
