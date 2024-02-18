<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/resendOrderCls.asp"-->

<%

dim oResend,  isCancel, itemid, itemoption, vSiteName, designer
itemid  = RequestCheckVar(request("itemid"),10)
itemoption  = RequestCheckVar(request("itemoption"),10)
isCancel = request("isCancel")
vSiteName		= requestCheckVar(request("sitename"),10)
designer		= requestCheckVar(request("designer"),32)

if isCancel="" then isCancel="A"

set oResend = New CReSend
oResend.FPageSize = 500
oResend.FRectIsCancel = isCancel
oResend.FRectSiteName = vSiteName

oResend.FRectMakerid = designer
oResend.FRectItemID = itemid
oResend.FRectItemOption = itemoption
oResend.GetResendOrderList

dim i, tmp

%>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� : <% drawSelectBoxDesigner "designer", designer %>
			&nbsp;
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="8" maxlength="10">
            &nbsp;
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="8" maxlength="10">
            &nbsp;
			Site :
			<select name="sitename" class="select">
				<option value="">-��ü-</option>
				<option value="10x10" <%=CHKIIF(vSiteName="10x10","selected","")%>>�ٹ�����</option>
				<option value="NOTTEN" <%=CHKIIF(vSiteName="NOTTEN","selected","")%>>���޻�</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="radio" name="isCancel" value="A" <% if (isCancel = "A") then response.write "checked" end if %>> ��ü���
			<input type="radio" name="isCancel" value="C" <% if (isCancel = "C") then response.write "checked" end if %>> ����ֹ�
		</td>
	</tr>
	</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmview" method="get">
	<input type="hidden" name="iid" value="">
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="10">
			�˻���� : <b><%= oResend.FResultCount %></b> / �ֹ��Ǽ� : <b><%= oResend.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="70">�ֹ���ȣ</td>
        <td width="70">Site</td>
	    <td width="60">�ֹ���</td>
	    <td width="60">������</td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="100">�귣��</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="40">�ֹ�<br>����</td>
		<td width="40">�ֹ���</td>
		<td width="40">�����</td>
	</tr>
	<% if oResend.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF">
	  	<td colspan="10" align="center">�˻������ �����ϴ�.</td>
	</tr>
	<% else %>

	<% for i=0 to oResend.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td align="center">
		<%
			if (tmp <> oResend.FItemList(i).FOrderSerial) then
				tmp = oResend.FItemList(i).FOrderSerial
		%>
			<%= oResend.FItemList(i).FOrderSerial %>
	    <% end if %>
	    </td>
        <td><%= oResend.FItemList(i).FSiteName %></td>
		<td><%= oResend.FItemList(i).FBuyName %></td>
    	<td><%= oResend.FItemList(i).FReqName %></td>
	    <td><%= oResend.FItemList(i).FItemId %></td>
		<td><%= oResend.FItemList(i).Fmakerid %></td>
		<td align="left">
			<%= oResend.FItemList(i).FItemname %>
			<% if oResend.FItemList(i).FItemOptionName<>"" then %>
			<font color="blue">[<%= oResend.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= oResend.FItemList(i).FItemNo %></td>
		<td><%= left(oResend.FItemList(i).FRegDate,10) %></td>
		<td><%= left(oResend.FItemList(i).FRegDate,10) %></td>
	</tr>
  <% next %>
  <% end if %>
</table>


<%
set oResend = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
