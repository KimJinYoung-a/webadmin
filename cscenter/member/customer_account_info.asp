<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� >>[CS]��ȯ�Ұ�������
' History : 2020.12.01 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/member/customercls.asp"-->

<%
Dim oaccount,i,page,userid
	userid = requestcheckvar(request("userid"),32)
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
    page = requestcheckvar(getNumeric(request("page")),10)

if page = "" then page = 1

set oaccount = new CUserInfo
	oaccount.FPageSize = 20
	oaccount.FCurrPage = page
	oaccount.frectuserid = userid

    if userid<>"" then
	    oaccount.GetUser_accountinfo_List()
    end if
%>

<script type="text/javascript">

function getsubmit(page){
    frm.page.value=page;
    frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
    <td align="left">
        �����̵�(��ȸ���� �ֹ���ȣ) : <input type="text" name="userid" value="<%= userid%>" size="12" onKeyPress="if(window.event.keyCode==13) getsubmit('1');"> 			
    </td>	
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="getsubmit('1');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
	
    </td>
</tr>
</table>
<!-- �˻� �� -->

<br>
		
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        <% if userid="" then %>
            <font color="red">�����̵�(��ȸ���� �ֹ���ȣ)�� �Է��� �ּ���.</font>
        <% end if %>
    </td>
    <td align="right">	
    </td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="5">
        �˻���� : <b><%= oaccount.FTotalCount %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�����̵�<br>(��ȸ���� �ֹ���ȣ)</td>
    <td>�����</td>	
    <td>���¹�ȣ</td>	
    <td>���¸�</td>
    <td>���</td>
</tr>
<% if oaccount.FresultCount>0 then %>
<% for i=0 to oaccount.FresultCount-1 %>

<tr align="center" bgcolor="#FFFFFF">
    <td>
        <%= oaccount.FItemList(i).fuserid %>
    </td>		
    <td>
        <%= oaccount.FItemList(i).frebankname %>
    </td>	
    <td align="left">
        <%= oaccount.FItemList(i).fencaccount %>
    </td>
    <td>
        <%= oaccount.FItemList(i).frebankownername %>
    </td>
    <td></td>
</tr>   

<% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="5" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>

</table>

<%
set oaccount = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

