<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ֹ�����(�ٹ�) ������ø���Ʈ
' History : 2007�� 11�� 29�� �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

dim nowdate,date1,date2,Edate
nowdate = now

if (yyyy1="") then
	''date1 = dateAdd("d",-4,nowdate)
	date1 = nowdate
	yyyy1 = Left(CStr(date1),4)
	mm1   = Mid(CStr(date1),6,2)
	dd1   = Mid(CStr(date1),9,2)

	yyyy2 = Left(CStr(nowdate),4)
	mm2   = Mid(CStr(nowdate),6,2)
	dd2   = Mid(CStr(nowdate),9,2)

	Edate = Left(CStr(nowdate+1),10)
else
	Edate = Left(CStr(dateserial(yyyy2, mm2 , dd2)+1),10)
end if

dim objbaljumakeonorder, balju_code

balju_code	= requestCheckVar(request("balju_code"),10)

set objbaljumakeonorder = New COldMiSend
objbaljumakeonorder.FPageSize = 500

if (balju_code <> "") then
	objbaljumakeonorder.FRectBaljuCode = balju_code
else
	objbaljumakeonorder.FRectStartDate = yyyy1 + "-" + mm1 + "-" + dd1
	objbaljumakeonorder.FRectEndDate = Left(CStr(Edate),10)
end if

objbaljumakeonorder.GetBaljuListMakeOnOrder

dim i, tmp
dim orgitemno, makeonorderitemno
%>
<script language='javascript'>

function misendmaster(v){
	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + v,"misendmaster","width=1200 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cOrderFin(detailidx){
    if (confirm('��� ó�� Ȯ�� �Ͻðڽ��ϱ�?')){
        var popwin = window.open("/admin/ordermaster/misendmaster_main_process.asp?detailidx=" + detailidx + "&mode=cancelFin","misendmaster_process","width=100 height=100 scrollbars=yes resizable=yes");
	    popwin.focus();
    }
}
</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			��������ڵ� :
			<input type="text" class="text" name="balju_code" value="<%= balju_code %>" size="10" maxlength="12">
			&nbsp;
			��ȸ�Ⱓ(�����������) : <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> (��������ڵ� ���� ���)
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>

<p>

* �ִ� 500������ ǥ�õ˴ϴ�.

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmview" method="get">
	<input type="hidden" name="iid" value="">
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%= objbaljumakeonorder.FResultCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	    <td width="70">�ֹ���ȣ</td>
        <td width="70">Site</td>
	    <td width="60">�ֹ���</td>
	    <td width="60">������</td>
		<td width="100">��ǰ����</td>
		<td width="50">��ǰ�ڵ�</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="40">�ֹ�<br>����</td>
	    <td>�ֹ�����<br>����</td>
		<td width="100">���</td>
	</tr>
	<% if objbaljumakeonorder.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF" height="25">
	  	<td colspan="16" align="center">�˻������ �����ϴ�.</td>
	</tr>
	<% else %>
	<% for i=0 to objbaljumakeonorder.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF" height="25">
	    <td align="center">
	    <%
	    if (tmp <> objbaljumakeonorder.FItemList(i).FOrderSerial) then
	      tmp = objbaljumakeonorder.FItemList(i).FOrderSerial

		  orgitemno = 0
		  makeonorderitemno = 0
	    %>
			<a href="javascript:misendmaster('<%= objbaljumakeonorder.FItemList(i).FOrderSerial %>');"><%= objbaljumakeonorder.FItemList(i).FOrderSerial %></a>
	    <% end if %>
		<%
		if (objbaljumakeonorder.FItemList(i).FisMakeOnOrderOrgItem) then
			orgitemno = orgitemno + objbaljumakeonorder.FItemList(i).FItemNo
		elseif (objbaljumakeonorder.FItemList(i).FisMakeOnOrderItem) then
			makeonorderitemno = makeonorderitemno + objbaljumakeonorder.FItemList(i).FItemNo
		end if
		%>
	    </td>
        <td><%= objbaljumakeonorder.FItemList(i).FSiteName %></td>
		<td><%= objbaljumakeonorder.FItemList(i).FBuyName %></td>
    	<td><%= objbaljumakeonorder.FItemList(i).FReqName %></td>
	    <td align="left">
			<% if (objbaljumakeonorder.FItemList(i).FisMakeOnOrderOrgItem) then %>
				<font color="blue">����ǰ</font>
			<% elseif (objbaljumakeonorder.FItemList(i).FisMakeOnOrderItem) then %>
				&nbsp; -&gt; <font color="green">�ֹ�����</font>
			<% end if %>
		</td>
		<td><%= objbaljumakeonorder.FItemList(i).FItemId %></td>
		<td align="left">
			<%= objbaljumakeonorder.FItemList(i).FItemname %>
			<% if objbaljumakeonorder.FItemList(i).FItemOptionName<>"" then %>
			<font color="blue">[<%= objbaljumakeonorder.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= objbaljumakeonorder.FItemList(i).FItemNo %></td>
		<td>
			<% if (objbaljumakeonorder.FItemList(i).FisMakeOnOrderItem) then %>
				<%= objbaljumakeonorder.FItemList(i).Frequiredetail %>
			<% end if %>
		</td>
	    <td>
			<% if (objbaljumakeonorder.FItemList(i).FisMakeOnOrderItem) then %>
				<% if (orgitemno = 0) then %>
					<font color="red"><b>����ǰ ����</b></font>
				<% elseif (orgitemno > 1) or (orgitemno <> makeonorderitemno) then %>
					<font color="red">��Ī �ʿ�</font>
				<% end if %>
			<% end if %>
		</td>
	</tr>
  <% next %>
  <% end if %>
</table>


<%
set objbaljumakeonorder = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
