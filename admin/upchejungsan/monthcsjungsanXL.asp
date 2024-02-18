<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
Dim yyyy1, mm1, jgubun
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)
jgubun = "CC"

dim oCSetcjungsan
set oCSetcjungsan = new CUpcheJungsanTax
	oCSetcjungsan.FPageSize = 5000
	oCSetcjungsan.FCurrPage = 1
	oCSetcjungsan.FRectYYYYMM = yyyy1 & "-" & mm1
	oCSetcjungsan.FRectJGubun = jgubun
	oCSetcjungsan.getMonthCsjungsanList
dim i
%>

<!-- �������Ϸ� ���� ��� �κ� -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
Response.ContentType = "application/unknown"
Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=CS��Ÿ����_"&yyyy1&"-"&mm1&".xls"
%>
<style type="text/css">
/* ���� �ٿ�ε�� ����� ���ڷ� ǥ�õ� ��� ���� */
.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="100">�귣��ID</td>
    <td width="100">�����ڵ�</td>
    <td width="100">���޸��ֹ���ȣ</td>
    <td width="100">�Ǹ�ä��</td>
    <td width="100">������</td>
    <td width="100">������</td>
    <td width="100">��ǰ�ڵ�</td>
    <td width="100">�ɼ��ڵ�</td>
    <td width="100">��ǰ��</td>
    <td width="180">�ɼǸ�</td>
    <td width="100">����</td>
    <td width="100">�����Ѿ�</td>
    <td width="100">�⺻�Ǹż�����</td>
    <td width="100">�������ξ�(�ٹ����ٺδ�)</td>
    <td width="200">�����ֹ���(���»�����)</td>
    <td width="100">������</td>
  	<td width="200">�������������</td>
  	<td width="100">�����</td>
  	<td width="200">�����հ�(����*�����)</td>
  	<td width="100">�ְ�������</td>
</tr>
<% For i=0 to oCSetcjungsan.FresultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FDesignerid%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FMastercode%></td>
    <td class="txt"><% if oCSetcjungsan.FItemList(i).FSitename<>"10x10" then %><%=oCSetcjungsan.FItemList(i).Fauthcode%><% end if %></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FSitename%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FBuyname%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FReqname%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemid%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemoption%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemname%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemoptionname%></td>
    <td class="txt"><%=oCSetcjungsan.FItemList(i).FItemno%></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FSellcash, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FCouponPlusCommi, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FCoupoonDiscount, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FReducedprice, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FCommission, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FPgcommission, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FSuplycash, 0) %></td>
    <td align="right"><%= FormatNumber(oCSetcjungsan.FItemList(i).FSumsuplycash,0) %></td>
    <td class="txt"><%= oCSetcjungsan.FItemList(i).FPaymethod %></td>
</tr>
<% Next %>
</table>
</body>
</html>
<% Set oCSetcjungsan = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
