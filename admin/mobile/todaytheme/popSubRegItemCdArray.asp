<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/mobile/todaythemeCls.asp" -->
<%
'###############################################
' PageName : popSubItemEdit.asp
' Discription : ���������� ��ǰ�ڵ� �ϰ� ���
' History : 2013.12.17 ����ȭ : �ű� ����
'###############################################

'// ���� ����
Dim listidx

'// �Ķ���� ����
listidx = request("listidx")

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
$(function(){
	//������ư
	$("#rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});

// ���˻�
function SaveForm(frm) {
	var selChk=true;
	if(frm.subItemidArray.value=="") {
		alert("�ϰ� ����Ͻ� ��ǰ�ڵ带 �Է����ּ���");
		frm.subItemidArray.focus();
		return;
	}

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}
</script>
<center>
<form name="frmSub" method="post" action="doSubRegItemCdArray.asp" style="margin:0px;">
<input type="hidden" name="listidx" value="<%=listidx%>" />
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>���� ���� - ��ǰ�ڵ� �ϰ� ���</b></td>
</tr>
<colgroup>
	<col width="100" />
	<col width="*" />
	<col width="100" />
	<col width="*" />
</colgroup>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��ǰ�ڵ�</td>
    <td colspan="3">
        <textarea name="subItemidArray" class="textarea" title="��ǰ�ڵ�" style="width:100%; height:80px;"></textarea>
        <p>�� ��ǰ�ڵ带 ��ǥ(,) �Ǵ� ���ͷ� �����Ͽ� �Է�</p>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���ļ���</td>
    <td>
        <input type="text" name="subSortNo" class="text" size="4" value="0" />
    </td>
    <td bgcolor="#DDDDFF">��뿩��</td>
    <td>
		<span id="rdoUsing">
		<input type="radio" name="subIsUsing" id="rdoUsing1" value="Y" checked /><label for="rdoUsing1">���</label>
		<input type="radio" name="subIsUsing" id="rdoUsing2" value="N" /><label for="rdoUsing2">����</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" �� �� " onClick="SaveForm(this.form);"></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->