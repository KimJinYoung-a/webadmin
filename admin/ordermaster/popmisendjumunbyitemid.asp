<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%
dim itemid,obalju

itemid = request("itemid")

set obalju = New COldMiSend
obalju.FRectItemid = itemid
obalju.GetMiSendOrderByitemid

dim i
%>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	��ǰ�ڵ� : <input type="text" name="orderserial" value="<%= obalju.FRectItemid %>" size="12">
	        </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80" >�ֹ���ȣ</td>
		<td width="80">������ /<br>������</td>
	    <td width="80">�ֹ��� /<br>������</td>
	   	<td width="60">����Ʈ��</td>
		<td width="80">���̵�</td>
		<td width="60">�����ݾ�</td>
		<td width="70">�ֹ�����/<br>����No</td>
		<td width="50">�ֹ�����</td>
		<td width="50">��������</td>
		<td width="70">����<br>����</td>
		<td>��û����</td>
		<td width="70">ó��<br>���</td>
		<td width="70">ó��<br>����</td>
	</tr>
	<% for i=0 to obalju.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= obalju.FItemList(i).Forderserial %></td>
		<td><%= obalju.FItemList(i).FBuyName %> <br><%= obalju.FItemList(i).FReqName %></td>
	    <td><%= Left(obalju.FItemList(i).FRegDate,10) %> <br><%= Left(obalju.FItemList(i).FIpkumDate,10) %></td>
	   	<td><%= obalju.FItemList(i).FSiteName %></td>
		<td><%= obalju.FItemList(i).FUserId %></td>
		<td><%= FormatNumber(obalju.FItemList(i).FSubTotalPrice,0) %></td>
		<td><font color="<%= obalju.FItemList(i).IpkumDivColor %>"><%= obalju.FItemList(i).IpkumDivName %></font><br><%= obalju.FItemList(i).FDeliveryNo %></td>
		<td><%= obalju.FItemList(i).Fitemno %></td>
		<td><font color="red"><b><%= obalju.FItemList(i).FItemLackNo %></b></font></td>
		<td><%= obalju.FItemList(i).getMiSendCodeName %></td>
		<td><%= obalju.FItemList(i).FrequestString  %></td>
		<td><%= obalju.FItemList(i).FfinishString  %></td>
		<td><%= obalju.FItemList(i).GetStateString  %></td>
	</tr>
	<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set obalju = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->