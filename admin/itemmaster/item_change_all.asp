<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : ��ǰ�ϰ�����[������]
' History : 2021.11.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
if not(C_ADMIN_AUTH or C_MD_AUTH or C_SYSTEM_Part) then
    response.write "<script type='text/javascript'>alert('������ �����ϴ�. MD��,������ ��Ʈ�� �̻� ���� ���� �մϴ�.');</script>"
    dbget.close() : response.end
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function chmakermwdiv(){
    if ( $('#makerid').val() == ''){
        alert('�����Ͻ� �귣��ID�� ������ �ּ���.');
        return;
    }
    $('#frmmakeritemch').attr('target', 'view');
    $('#frmmakeritemch').attr('action', '/admin/itemmaster/item_change_all_process.asp');
    $('#mode').val('makerchmwdiv');
    $('#frmmakeritemch').submit();
    $('#mode').val('');
}

function chmakermargin(){
    if ( $('#makerid').val() == ''){
        alert('�����Ͻ� �귣��ID�� ������ �ּ���.');
        return;
    }
    if ( $('#margin').val() == '' || $('#margin').val() == 0){
        alert('�����Ͻ� ���� % �� �Է��� �ּ���.');
        return;
    }
    $('#frmmakeritemch').attr('target', 'view');
    $('#frmmakeritemch').attr('action', '/admin/itemmaster/item_change_all_process.asp');
    $('#mode').val('makerchmargin');
    $('#frmmakeritemch').submit();
    $('#mode').val('');
}

function chmakersellyn_n(){
    if ( $('#makerid').val() == ''){
        alert('�����Ͻ� �귣��ID�� ������ �ּ���.');
        return;
    }
    $('#frmmakeritemch').attr('target', 'view');
    $('#frmmakeritemch').attr('action', '/admin/itemmaster/item_change_all_process.asp');
    $('#mode').val('makerchsellyn_n');
    $('#frmmakeritemch').submit();
    $('#mode').val('');
}

function chMoveMaker(){
    if ( $('#makerid').val() == ''){
        alert('�����Ͻ� �귣��ID�� ������ �ּ���.');
        return;
    }
    if ( $('#toMakerid').val() == ''){
        alert('�̵��� �귣��ID�� ������ �ּ���.');
        return;
    }
    $('#frmmakeritemch').attr('target', 'view');
    $('#frmmakeritemch').attr('action', '/admin/itemmaster/item_change_all_process.asp');
    $('#mode').val('MoveMaker');
    $('#frmmakeritemch').submit();
    $('#mode').val('');
}
</script>
<form name="frmmakeritemch" id="frmmakeritemch" method="post" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" id="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td>
		������ �귣��ID : 
		<input type="text" class="text" name="makerid" id="makerid" value="" size="15" maxlength=32 >
		<input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'makerid');" >
    </td>
</tr>
<tr class="a" height="25" bgcolor="#FFFFFF">
    <td>
        �ش� �귣�� ��ǰ ��� ��౸��(
		<label><input type="radio" name="mwdiv" value="M" checked>����</label>
		<label><input type="radio" name="mwdiv" value="W" >��Ź</label>
		<label><input type="radio" name="mwdiv" value="U" >��ü</label>
        )���� ����. �귣�� ��ǥ ������ ���� ���� �ϼž� �մϴ�.
        <input class="button" type="button" value="��౸�к����ϱ�" onClick="chmakermwdiv();">
    </td>
</tr>
<tr class="a" height="25" bgcolor="#FFFFFF">
    <td>
        �ش� �귣�� ��ǰ ��� ����(<input type="text" class="text" name="margin" id="margin" value="0" size="5" maxlength=5 >
        )% ���� ����.
        <input class="button" type="button" value="���������ϱ�" onClick="chmakermargin();">
    </td>
</tr>
<tr class="a" height="25" bgcolor="#FFFFFF">
    <td>
        �ش� �귣�� ��ǰ ��� <input class="button" type="button" value="�Ǹž������κ����ϱ�" onClick="chmakersellyn_n();">
    </td>
</tr>
<tr class="a" height="25" bgcolor="#FFFFFF">
    <td>
        �ش� �귣�� ��ǰ ���
		<input type="text" class="text" name="toMakerid" id="toMakerid" value="" size="15" maxlength="32" />
		<input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'toMakerid');" />
        <input class="button" type="button" value="�귣��� �����ϱ�" onClick="chMoveMaker();" />
        �� ���� �� ��౸���� ������ �����ؾ��մϴ�.
    </td>
</tr>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
    <iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" ></iframe>
<% else %>
    <iframe id="view" name="view" src="" width="100%" height=0 frameborder="0" ></iframe>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->