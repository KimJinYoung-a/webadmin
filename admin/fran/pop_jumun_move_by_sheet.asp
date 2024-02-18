<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����� ���� ����̵�
' History : 2018.02.07 �̻� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%

dim i, j, k
dim idx

idx  = requestCheckVar(request("idx"), 3200)

%>
<script>
function jsChkForm(frm) {
	if (frm.moveshopid.value == "") {
		alert("����!!\n\n�������� �����ȵ�.");
		return false;
	}

	if (frm.scheduledt.value.length<1){
		alert('����̵����� �Է��ϼ���');
		calendarOpen3(frm.scheduledt,'����̵����� �Է��ϼ���','');
		return false;
	}

	if (frm.songjangdiv.value.length<1){
		alert('�ù�縦 ���� �ϼ���');
		frm.songjangdiv.focus();
		return false;
	}

	if (frm.songjangno.value.length<1){
		alert('���� ��ȣ�� �Է� �ϼ���');
		msfrm.songjangno.focus();
		return false;
	}

	return true;
}

function jsStockMove() {
	var frm = document.frm;

	if (jsChkForm(frm) != true) { return; }

	var ret = confirm('�Է��ϽŴ�� ��� �̵�ó�� �Ͻðڽ��ϱ�?');
	if (ret) {
		frm.mode.value = "saveorderbysheet";
		frm.method = "post";
		frm.action = "pop_jumun_move_process.asp";
		frm.submit();
	}
}
</script>
<form name="frm" method="get" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="<%= idx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
	    <td height="25" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	    <td>
		    <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("moveshopid", "", "21") %>
	    </td>
    </tr>
    <tr bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("tabletop") %>">
		    ����̵���
	    </td>
	    <td>
		    <input type="text" class="text" name="scheduledt" value="" size=10 readonly ><a href="javascript:calendarOpen(frm.scheduledt);">
		        <img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>

			    �ù�� :<% drawSelectBoxDeliverCompany "songjangdiv", "" %>
			    �����ȣ:<input type="text" class="text" name="songjangno" size=14 maxlength=16 value="" >
			    <br>
			    (�ù�� ������ ������� �ù��:��Ÿ���� �����ȣ:�����, ������� ���� �Է� �Ͻø� �˴ϴ�.)
	    </td>
    </tr>
    <tr bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("tabletop") %>" colspan="2" align="center">
	        <input type="button" value="����̵�ó��" onClick="jsStockMove()" class="button" id="btnMove">
	    </td>
    </tr>
</table>
</form>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
