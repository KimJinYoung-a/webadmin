<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim falg
falg = request("falg")

%>
<script language='javascript'>
function MrOk(){
    if (document.all.rejectmsg.value.length<1){
        alert('������ �����ϼ���.');
        document.all.rejectmsg.focus();
        return;
    }else{
        document.all.ret.value = document.all.rejectmsg.value;
    }
    
    if (document.all.rejectmsg.value=="�����Է�"){
        if (document.all.rejectmsg_Text.value.length<1){
            alert('������ �Է� �ϼ���.');
            document.all.rejectmsg_Text.focus();
            return;
        }else{
            document.all.ret.value = document.all.rejectmsg_Text.value;
        }
    }
    window.close();
}

function MrCancel(){
    document.all.ret.value = '';
    window.close();
}

function ChgCombo(comp){
    if (comp.value=="�����Է�"){
        document.all.divtext.style.display = "inline";
    }else{
        document.all.divtext.style.display = "none";
    }
}

</script>
<BODY bgcolor="#ffffff" OnUnload="window.returnValue = document.all.ret.value;">
<INPUT type="hidden" name="ret">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" class="a">
<% if (falg="1") then %>
<!-- ��Ϻ��� -->
<tr height="30">
    <td align="center">��� ���� ���� ����</td>
</tr>
<tr height="30">
    <td align="center">
        <select name="rejectmsg" onChange="ChgCombo(this);">
        <option value="">����
        <option value="�̹��� ��� �ҷ�">�̹��� ��� �ҷ�
        <option value="��ǰ ���� ����">��ǰ ���� ����
        <option value="�����Է�">----�����Է�----
        </select>
    </td>
</tr>

<% elseif (falg="2") then %>
<!-- ��� �Ұ� ���� -->
<tr height="30">
    <td align="center">��� �Ұ� ���� ����</td>
</tr>
<tr height="30">
    <td align="center">
        <select name="rejectmsg" onChange="ChgCombo(this);">
        <option value="">����
        <option value="���ϻ�ǰ �Ǹ���">���ϻ�ǰ �Ǹ���
        <option value="�����Է�">----�����Է�----
        </select>
    </td>
</tr>

<% end if %>
<tr height="30">
    <td id="divtext" style="display=none;" align="center">
        <input type="text" name="rejectmsg_Text" size="30" maxlength="100">
    </td>
</tr>

<tr height="30">
    <td align="center">
        <input type="button" class="button" value="Ȯ��" onclick="MrOk()">
        <input type="button" class="button" value="���" onclick="MrCancel()">
    </td>
</tr>
</table>
</body>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
