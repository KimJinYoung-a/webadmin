<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ȸ�� ���� �����丮
' Hieditor : 2011.02.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #Include virtual = "/lib/classes/totalpoint/totalpointCls.asp" -->

<%
Dim ohistory, i, page , orderno ,totrealprice
dim vCardNo, vUserName, vUserID, posuid, pssnkey, dummikey, shopid
	orderno = requestCheckVar(Request("orderno"),20)
	vCardNo			= requestCheckVar(Request("cardno"),20)
	vUserName		= requestCheckVar(Request("username"),20)
	vUserID			= requestCheckVar(Request("userid"),32)
	posuid			= Request("posuid")
	pssnkey			= Request("pssnkey")
	dummikey		= Request("dummikey")
	shopid = request("shopid")
	menupos = request("menupos")

set ohistory = new TotalPoint
	ohistory.frectorderno = orderno
	ohistory.fsell_history_detail()
%>

<script language="javascript">

function refer(){
	frm.action='/admin/totalpoint/customer_sell_history.asp';
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="posuid" value="<%=posuid%>">
<input type="hidden" name="pssnkey" value="<%=pssnkey%>">
<input type="hidden" name="dummikey" value="<%=dummikey%>">
<input type="hidden" name="cardno" value="<%=vCardNo%>">
<input type="hidden" name="username" value="<%=vUserName%>">
<input type="hidden" name="userid" value="<%=vUserID%>">
<input type="hidden" name="shopid" value="<%=shopid%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�ֹ���ȣ: <input type="text" class="text" name="orderno" value="<%=orderno%>">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% '<input type="button" class="button" value="�������" onClick="refer();"> %>
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ohistory.FTotalCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ohistory.FTotalCount %></b> ���� 1000�� ���� �˻� �˴ϴ�
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�ֹ���ȣ</td>
	<td>�����̸�</td>
	<td>����ID</td>
	<td>��ǰ��ȣ</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td>�Ǹűݾ�</td>
	<td>�ǰ�����</td>
	<td>�Ǹż���</td>
	<td>�հ�</td>
	<td>���</td>
</tr>
<%

for i=0 to ohistory.FTotalCount-1

if ohistory.FItemList(i).fcancelyn = "N" then
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td>
		<%= ohistory.FItemList(i).forderno %>
		
		<% if ohistory.FItemList(i).fcancelyn = "Y" then %>
			<br>(���)
		<% end if %>
	</td>
	<td><%= ohistory.FItemList(i).fshopname %></td>
	<td><%= ohistory.FItemList(i).fshopid %></td>
	<td><%= ohistory.FItemList(i).fitemgubun %>-<%= CHKIIF(ohistory.FItemList(i).fitemid>=1000000,Format00(8,ohistory.FItemList(i).fitemid),Format00(6,ohistory.FItemList(i).fitemid)) %>-<%= ohistory.FItemList(i).fitemoption %></td>
	<td><%= ohistory.FItemList(i).fitemname %><br></td>
	<td><%= ohistory.FItemList(i).fitemoptionname %></td>
	<td><%= FormatNumber(ohistory.FItemList(i).fsellprice,0) %></td>
	<td><%= FormatNumber(ohistory.FItemList(i).frealsellprice,0) %></td>
	<td><%= ohistory.FItemList(i).fitemno %></td>
	<td><%= FormatNumber(ohistory.FItemList(i).frealsellprice*ohistory.FItemList(i).fitemno,0) %></td>
	<td></td>
</tr>
<%
totrealprice = totrealprice + (ohistory.FItemList(i).frealsellprice*ohistory.FItemList(i).fitemno)
next
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=6>�հ�</td>
	<td><%= FormatNumber(totrealprice,0) %></td>
	<td colspan=10></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set ohistory = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->