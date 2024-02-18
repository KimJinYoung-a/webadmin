<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->

<%
dim id, fin
id = RequestCheckvar(request("id"),10)
fin = RequestCheckvar(request("fin"),16)

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if


''Ȯ�ο�û���� :
dim OCsConfirm
set OCsConfirm = new CCSASList
OCsConfirm.FRectCsAsID = id

if id<>"" then
    OCsConfirm.GetOneCsConfirmItem
end if


if (ocsaslist.FResultCount<1) then
    response.write "<script>alert('��ȿ���� ���� �����Դϴ�.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

''Ȯ�ο�û������ ���� ���=> �������¸� ��� ����.
''Ȯ�ο�û������ �ִ� ���=> Ȯ�� ��û ���¿����� ����/�Ϸ� ����.
if (OCsConfirm.FResultCount<1) then
    if (ocsaslist.FOneItem.FCurrstate>="B006") then
        response.write "<script>alert('�Ϸ� ���� ���� ������ ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
else
    if (ocsaslist.FOneItem.FCurrstate>="B006") then
        response.write "<script>alert('�Ϸ� ���� ���� ������ ����/�Ϸ� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
end if


dim IsEditMode
IsEditMode = (OCsConfirm.FResultCount>0)

%>
<script language='javascript'>
function ActConfirmReg(frm){
    if (frm.confirmregmsg.value.length<1){
        alert('Ȯ�ο�û ������ �Է��� �ּ���.');
        frm.confirmregmsg.focus();
        return;
    }

    if (confirm('Ȯ�� ��û ���� �Է½� ���°� Ȯ�� ��û������ ����Ǹ�, ȯ�����Ͽ��� �����˴ϴ�. ����Ͻðڽ��ϱ�?')){
        frm.nextstate.value = "B005";
        frm.mode.value = "reg";
        frm.submit();
    }
}

function ActConfirmReEdit(frm){
    if (frm.confirmregmsg.value.length<1){
        alert('Ȯ�ο�û ������ �Է��� �ּ���.');
        frm.confirmregmsg.focus();
        return;
    }

    if (confirm('���û �Ͻðڽ��ϱ�?\n ���� Ȯ�� ������ �ʱ� ȭ �˴ϴ�.')){
        frm.nextstate.value = "B005";
        frm.mode.value = "reInput";
        frm.submit();
    }
}


function ActConfirmEdit(frm){
    if (frm.confirmregmsg.value.length<1){
        alert('Ȯ�ο�û ������ �Է��� �ּ���.');
        frm.confirmregmsg.focus();
        return;
    }

    if (confirm('������ ��û ������ ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "edit";
        frm.submit();
    }
}

function ActConfirmFinish(frm){
    if (frm.confirmfinishmsg.value.length<1){
        alert('Ȯ��ó�� ������ �Է��� �ּ���.');
        frm.confirmfinishmsg.focus();
        return;
    }

    if (confirm('Ȯ�� ��û ���� ó���� ���°� �������·� �� ����˴ϴ�. �Ϸ� ó�� �Ͻðڽ��ϱ�?')){
        frm.nextstate.value = "B001";
        frm.mode.value = "finish";
        frm.submit();
    }
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmConfirm" method="post" action="pop_ConfirmMsg_process.asp">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="nextstate" value="">
<input type="hidden" name="mode" value="">


<% if IsEditMode then %>

<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">Ȯ�ο�û �����</td>
    <td bgcolor="#FFFFFF">
        <%= OCsConfirm.FOneItem.Fconfirmreguserid %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">Ȯ�ο�û����</td>
    <td bgcolor="#FFFFFF">
        <textarea class="textarea" name="confirmregmsg" cols="48" rows="5"  ><%= OCsConfirm.FOneItem.Fconfirmregmsg %></textarea>
    </td>
</tr>
<% if fin<>"" then %>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">Ȯ�ο�û ó����</td>
    <td bgcolor="#FFFFFF">
        <%= session("ssBctID") %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">Ȯ��ó������</td>
    <td bgcolor="#FFFFFF">
        <textarea class="textarea" name="confirmfinishmsg" cols="48" rows="5"  ><%= OCsConfirm.FOneItem.Fconfirmfinishmsg %></textarea>
    </td>
</tr>
<% end if %>
<% else %>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">Ȯ�ο�û�����</td>
    <td bgcolor="#FFFFFF">
        <%= session("ssBctID") %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">Ȯ�ο�û����</td>
    <td bgcolor="#FFFFFF">
        <textarea class="textarea" name="confirmregmsg" cols="48" rows="5"  ></textarea>
    </td>
</tr>
<% end if %>
<tr height="25">
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <% if IsEditMode then %>
        <% if Not IsNULL(OCsConfirm.FOneItem.Fconfirmfinishdate) then %>
        ó�� �Ϸ�� �����Դϴ�. <br>: (��Ȯ�ο�û�� ����ó�������� �����˴ϴ�.)
        <br>
        <% if (fin<>"fin") then %>
        <!-- ��Ȯ�ο�û ����޴� �߰� ��� -->
        <input type="button" class="button" value=" �� Ȯ�ο�û  " onClick="ActConfirmReEdit(frmConfirm)">
        <% end if %>
        <% else %>
            <% if fin="" then %>
            <input type="button" class="button" value=" Ȯ�ο�û ���� " onClick="ActConfirmEdit(frmConfirm)">
            <% else %>
            <input type="button" class="button" value=" Ȯ�ο�û �Ϸ� " onClick="ActConfirmFinish(frmConfirm)">
            <% end if %>
        <% end if %>
    <% else %>
    <input type="button" class="button" value=" Ȯ�ο�û ��� " onClick="ActConfirmReg(frmConfirm)">

    <% end if %>
    </td>
</tr>
</form>
</table>
<%
set ocsaslist = Nothing
set OCsConfirm = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->