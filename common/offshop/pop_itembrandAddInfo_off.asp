<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2013.02.21 �ѿ�� ����
' Description : �귣���ǰ �߰�
'				input - actionURL(db ó���� �ʿ��� �Ķ���ͱ��� ����) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim i, makerid, shopid, acURL, brandcount
	makerid    = RequestCheckVar(request("makerid"),32)
	shopid    = RequestCheckVar(request("shopid"),32)
	acURL	= request("acURL")

if shopid = "" then
	response.write "<script>alert('����ID �� �����ϴ�'); self.close();</script>"
	response.end
end if

if makerid<>"" then
	brandcount = getcontractbranditemcount(shopid,makerid)
end if
%>

<script language="javascript">

function jsSerach(){

	frm.target = "";
	frm.action = "";
	frm.submit();
}

function insertbranditem(){	
	
	if( confirm('�ش� ��ǰ�� ��� �߰� �Ͻðڽ��ϱ�?') ){
		frm.target = "FrameCKP";
		frm.action = "<%=acURL%>";
		frm.submit();

		opener.history.go(0);	
		//window.close();
	
	}else{
		return;	
	}	
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="mode" value="bi">
<input type="hidden" name="acURL" value="<%=acURL%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
	</td>
	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach('');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
	</td>
</tr>    
</table>

<br>

<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
<tr valign="bottom">       
    <td align="left">
    	�ظ���(<%=shopid%>)�� ���� ��ǰ�� �˻� �˴ϴ�.
    </td>
    <td align="right">
    </td>        
</tr>	
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="20">
		<% if brandcount<>"" then %>
			�˻���� : <b><%= brandcount %></b>�� ��ǰ�� �˻� �Ǿ����ϴ�.
			<% if brandcount <> 0 then %>
				<input type="button" value="����߰�(<%= brandcount %>��)" onClick="insertbranditem()" class="button">
			<% end if %>
		<% else %>
			�귣�带 �Է��� �ּ���
		<% end if %>
	</td>		
</tr>
</table>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="800" height="100"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
