<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%
dim id
id = requestcheckVar(request("id"),9)

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if

dim divcd
if (ocsaslist.FResultcount>0) then
    divcd = ocsaslist.FOneItem.FDivCd
end if


''ȯ������
dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = ocsaslist.FOneItem.FId
orefund.GetOneRefundInfo

response.write "<br><br>�ý����� ���� : ������� ������!!"
dbget.close()
response.end

%>

<script language='javascript'>
//��ü �߰� ���� ������
function clearAddUpchejungsan(frm){
    frm.add_upchejungsandeliverypay.value = "0";
    frm.add_upchejungsancause.value = "";
    frm.add_upchejungsancauseText.value = "";

    frm.buf_totupchejungsandeliverypay.value = frm.buf_refunddeliverypay.value*1 + frm.add_upchejungsandeliverypay.value*1;

}


//�߰� ���� ����
function conFirmSave(frm){
    if (frm.add_upchejungsandeliverypay){
        if (frm.add_upchejungsandeliverypay.value == ""){
            alert('�߰������ۺ� �Է��ϼ���.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*0 != 0){
            alert('���ڸ� �����մϴ�.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='�����Է�')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('�߰� ������� �ִ°�� �귣�� ���̵� �����Ǿ�� �մϴ�. ');
                return;
            }

            //�ֹ� ������ ���̵� �ִ� ��츸.

        }else{
            <% if (divcd="A700") then %>
            //alert('�߰� ������� �Է��ϼ���.');
            //frm.add_upchejungsandeliverypay.focus();
            //return;
            <% end if %>
        }
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

//�߰������ۺ�
function Change_add_upchejungsandeliverypay(comp){
    comp.form.buf_totupchejungsandeliverypay.value = comp.form.buf_refunddeliverypay.value*1 + comp.value*1;

    if (isNaN(comp.form.buf_totupchejungsandeliverypay.value)){
        comp.form.buf_totupchejungsandeliverypay.value = "0";
    }
}

//�߰������ۺ� ����
function Change_add_upchejungsancause(comp){
    if (comp.value=="�����Է�") {
        document.all.span_add_upchejungsancauseText.style.display = "inline";
    }else{
        document.all.span_add_upchejungsancauseText.style.display = "none";
    }
}
</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
    <form name="frmaction" method="post" action="pop_cs_action_process.asp">
    <input type="hidden" name="mode" value="addupchejungsanEdit">
    <input type="hidden" name="id" value="<%= id %>">
	<tr bgcolor="FFFFFF">
	    <td colspan="2"><strong>* ��ü �߰� ���� ����</strong></td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td width="100">�귣��ID</td>
	    <td ><input type="text" class="text_ro" name="buf_requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>" size="20" ReadOnly >
	    <% if (divcd="A700") then %>
	    <input type="button" class="button" value="�귣��ID�˻�" onclick="jsSearchBrandID(this.form.name,'buf_requiremakerid');" >
	    <% end if %>
	    </td>
	</tr>

	<tr bgcolor="FFFFFF">
	    <td width="100">ȸ����ۺ�</td>
	    <td ><input type="text" class="text_ro" name="buf_refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >��</td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td width="100">�߰������ۺ�</td>
	    <td ><input type="text" class="text" name="add_upchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay %>" size="10" onKeyUp="Change_add_upchejungsandeliverypay(this);">��
	    &nbsp;<select name="add_upchejungsancause" class="text" onChange='Change_add_upchejungsancause(this);'>
	    <option value="" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="","selected","") %>>��������
	    <option value="�߰���ۺ�" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="�߰���ۺ�","selected","") %> >�߰���ۺ�
	    <option value="�߰�����" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="�߰�����","selected","") %>>�߰�����
	    <option value="�����Է�" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","selected","") %>>�����Է�
	    </select>

	    <span name="span_add_upchejungsancauseText" id="span_add_upchejungsancauseText" style='display:<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","inline","none") %>'><input type="text" name="add_upchejungsancauseText" value="<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"",ocsaslist.FOneItem.Fadd_upchejungsancause,"") %>" size="10" maxlength="16" ></span>
	    <a href="javascript:clearAddUpchejungsan(frmaction);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	    </td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td width="100">�������ۺ�</td>
	    <td ><input type="text" class="text_ro" name="buf_totupchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >��</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
    <td align="center">
    <input type="button" value="����" onClick="conFirmSave(frmaction);">
    </td>
</tr>
</table>
<%
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
