<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/analysiscls.asp"-->

<%
'response.write "������ ���� ���"
'dbget.close()	:	response.End

dim yyyy1,mm1
yyyy1 = request("yyyy1")
mm1 = request("mm1")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim oanal
set oanal = new CAnalysis
oanal.FRectYYYYMM = yyyy1 + "-" + mm1
oanal.getOnLineMonthGainSum

dim i, shopmmttl, shopsuppttl
shopmmttl = 0
shopsuppttl = 0

for i=0 to oanal.FResultCount-1
    shopmmttl = shopmmttl + oanal.FItemList(i).FTotsum
    shopsuppttl = shopsuppttl + oanal.FItemList(i).FSuplysum
next
%>

<table width="900" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
		�˻������:<% DrawYMBox yyyy1,mm1 %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<h3>������</h3>

<span class=a>** Admin ���⳻��</span>
<table width="900" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td>�������</td>
	<td>��Ÿ�������</td>
	<td>��ۺ�</td>
	<td>��۰Ǽ�</td>
	<td>�Ѹ����</td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td><%= FormatNumber(oanal.FOneItem.FWebTotalSel,0) %></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
</table>
<br>
<span class=a>** onLine ���곻��</span>
<table width="900" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td>��ü���</td>
	<td>����</td>
	<td>��Ź</td>
	<td>��Ÿ���</td>
	<td>�Ѹ��Ծ�</td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td><%= FormatNumber(oanal.FOneItem.FUbTotal,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FMeTotal,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FWiTotal,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FEtTotal,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.getTotalMeaip,0) %></td>
</tr>
</table>
<br><br>
<table width="900" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td width=180>����</td>
	<td width=180>����</td>
	<td>����</td>
	<td>����(������)</td>
	<td>���<br>(�������ݿ� ���Ծ�)</td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td rowspan=5 align=center>�¶���</td>
	<td align=left>��ü���</td>
	<td rowspan=3 bgcolor="#337799"><%= FormatNumber(oanal.FOneItem.FWebTotalSel,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FUbTotal,0) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td align=left>��Ź</td>
	<td><%= FormatNumber(oanal.FOneItem.FwiTotal,0) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td align=left>����</td>
	<td><%= FormatNumber(oanal.FOneItem.FMeTotal,0) %></td>
	<td><%= FormatNumber(shopmmttl,0) %></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td align=left>��Ÿ</td>
	<td></td>
	<td><%= FormatNumber(oanal.FOneItem.FEtTotal,0) %></td>
	<td></td>
</tr>
<tr bgcolor="#DDDDDD" align=right>
	<td align=left>�Ұ�</td>
	<td><%= FormatNumber(oanal.FOneItem.FWebTotalSel+shopsuppttl,0) %></td>
	<td><%= FormatNumber(oanal.FOneItem.FMeTotal + oanal.FOneItem.FwiTotal + oanal.FOneItem.FUbTotal + oanal.FOneItem.FEtTotal,0) %></td>
	<td></td>
</tr>

<% for i=0 to oanal.FResultCount-1 %>
<tr bgcolor="#FFFFFF" align=right>
    <% if i=0 then %>
    <td rowspan=<%= oanal.FResultCount + 1 %> align=center>��������<br>���<br>(���Ի�ǰ)</td>
    <% end if %>
	<td align=left>(<%= oanal.FItemList(i).FShopid %>)</td>
	<td><%= FormatNumber(oanal.FItemList(i).FSuplysum,0) %></td>
	<td></td>
	<td><%= FormatNumber(oanal.FItemList(i).FTotsum,0) %></td>
</tr>
<% next %>

<tr bgcolor="#DDDDDD" align=right>
	<td align=left>�Ұ�</td>
	<td></td>
	<td></td>
	<td><%= FormatNumber(shopmmttl,0) %></td>
</tr>

<tr bgcolor="#FFFFFF" align=right>
	<td rowspan=6 align=center>��Ÿ���</td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=right>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>

<tr bgcolor="#FFFFFF" align=right>
	<td align=center>�Ѱ�</td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
</table>
<br>
<table width="800" border="0" cellpadding="2" cellspacing="1" class="a" >
<tr>
	<td width=100 bgcolor="#337799"></td>
	<td>ǥ�� �Ǿ� �ִ� ������ ���λ� �����ִ� �����Դϴ�.</td>
</tr>
<tr>
	<td colspan="2">������ �����Ϸ� ���� / ������ ��ۿϷ� �����Դϴ�.</td>
</tr>
</table>
<%
set oanal = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
