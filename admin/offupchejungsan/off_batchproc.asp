<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->
<%
dim yyyy1, mm1
yyyy1 = request("yyyy1")
mm1 = request("mm1")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

''�ӽ�
''yyyy1 = "2006"
''mm1="12"
%>
<script language='javascript'>
function doOffbatch(comp, gubuncd){
//alert('�۾����Դϴ�. ���� �� �� �����ϴ�.');
//return;
<% if (session("ssbctId")<>"icommang") then %>
     alert('������ �� �����ϴ�.');
     return;
<% end if %>

    comp.disabled = true;
    var confirmStr = frmArrupdate.yyyy.value + '�� ' + frmArrupdate.mm.value + '�� ' + comp.value + ' ���� �Ͻðڽ��ϱ�?';
    if (confirm(confirmStr)){
        frmArrupdate.mode.value = "batchprocess";
        frmArrupdate.gubuncd.value = gubuncd;
        frmArrupdate.submit();
    }else{
        comp.disabled = false;
    }
}

function doOffNextbatch(comp){
//alert('�۾����Դϴ�. ���� �� �� �����ϴ�.');
//return;
    comp.disabled = true;
    if (confirm('��ü Ȯ���� ���·� �ϰ� ���� �Ͻðڽ��ϱ�? (������ ���¸� ����˴ϴ�.)')){
        frmArrupdate.mode.value = "batchnextstep";
        frmArrupdate.submit();
    }else{
        comp.disabled = false;
    }
}

function popOffConfirm(yyyymm,mode){
    var popwin = window.open('pop_off_jungsan_confirm.asp?yyyymm=' + yyyymm + '&mode='+mode,'pop_off_jungsan_confirm','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
    <tr height="10" valign="bottom" bgcolor="F4F4F4">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom" bgcolor="F4F4F4">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top" bgcolor="F4F4F4" width="730">
            �������� : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;
        </td>
        <td valign="top" bgcolor="F4F4F4" align="right">
            <a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
        <td>&nbsp;</td>
    </tr>
    <tr bgcolor="<%= adminColor("topbar") %>">
        <td>
            <input type="button" class="button" value="1.���� �����۾� " onClick="javascript:doOffbatch(this,'0001');">( �����귣�� �������ǻ���)
            <br><br>
            <input type="button" class="button" value="2.����-�ǸŽ� ��������" onClick="javascript:popOffConfirm('<%= yyyy1 %>-<%= mm1 %>','11');">
            <br><br>
            <input type="button" class="button" value="2.����-�귣�� ����" onClick="javascript:popOffConfirm('<%= yyyy1 %>-<%= mm1 %>','1');">
            <br>
            -------------------------&gt;��ü�������� ���ԼӼ� **����
            <br>
            <input type="button" class="button" value="3.���� �����۾�" onClick="javascript:doOffbatch(this,'0002');">( ��� Flag )
            <br><br>
            <input type="button" class="button" value="2.�ֹ� �ʰ� �ø�����" onClick="javascript:popOffConfirm('<%= yyyy1 %>-<%= mm1 %>','2');">
            <br><br>
            <!--
            <input type="button" class="button" value="��Ź�Ǹ� �ϰ��ۼ� (��ü->����->����)" onClick="javascript:doOffbatch(this,'B011');">
            <br><br>
            <input type="button" class="button" value="��ü��Ź�Ǹ� �ϰ��ۼ� (��ü->����)" onClick="javascript:doOffbatch(this,'B012');">
            <br><br>
            <input type="button" class="button" value="������ �ϰ��ۼ� (��ü->����->����)" onClick="javascript:doOffbatch(this,'B031');">
            <br><br>
            <input type="button" class="button" value="�������� �ϰ��ۼ� (��ü->����->����)" onClick="javascript:doOffbatch(this,'B021');">
            <br><br>
            <input type="button" class="button" value="������� �ϰ��ۼ� (��ü->����)" onClick="javascript:doOffbatch(this,'B022');">
            <br><br>
            <input type="button" class="button" value="��ü��� �ϰ��ۼ�" onClick="javascript:doOffbatch(this,'B077');">
            <br><br>
            <br><br>
            <br><br>
            -->
            <input type="button" class="button" value="������->��üȮ���� �ϰ�ó��" onClick="javascript:doOffNextbatch(this);">
            <br><br><br>
            <input type="button" class="button" value="�ߺ� ���� ����" onClick="javascript:popOffConfirm('<%= yyyy1 %>-<%= mm1 %>','90');">

        </td>
    </tr>
</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<form name="frmArrupdate" method="post" action="off_jungsan_process.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="gubuncd" value="">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->