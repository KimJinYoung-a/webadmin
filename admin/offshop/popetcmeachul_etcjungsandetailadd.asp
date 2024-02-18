<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ������
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%
dim idx, topidx, makerid, shopid

idx  = requestCheckVar(request("idx"),10)
topidx = requestCheckVar(request("topidx"),10)
makerid = requestCheckVar(request("makerid"),32)
shopid = requestCheckVar(request("shopid"),32)

%>
<script language='javascript'>
function AddValue(frm){
<% if (C_ADMIN_AUTH) then %>

<% else %>
	if (!ajaxBrandItem())
	{
		return;
	}
<% end if %>

	if (frm.itemno.value.length<1){
		alert('������ �Է� �ϼ���.');
		frm.itemno.focus();
		return;
	}

	if (frm.orgsellcash.value.length<1){
		alert('�Һ񰡸� �Է� �ϼ���.');
		frm.orgsellcash.focus();
		return;
	}


	if (frm.sellcash.value.length<1){
		alert('�ǸŰ��� �Է� �ϼ���.');
		frm.sellcash.focus();
		return;
	}

	if (frm.buycash.value.length<1){
		alert('���԰��� �Է� �ϼ���.');
		frm.buycash.focus();
		return;
	}

	if (frm.suplycash.value.length<1){
		alert('���ް��� �Է� �ϼ���.');
		frm.suplycash.focus();
		return;
	}

	frm.submit();
}

function checkBrandItem()
{
	if (event.keyCode==13)
		ajaxBrandItem("GET");
}

function ajaxBrandItem(mode)
{
	var f = document.frm;
	if (f.itemgubun.value.length!=2){
		alert('��ǰ������ �Է� �ϼ���.');
		f.itemgubun.focus();
		return false;
	}

	if (frm.itemid.value.length<1){
		alert('��ǰ��ȣ�� �Է� �ϼ���.');
		f.itemid.focus();
		return false;
	}

	if (frm.itemoption.value.length!=4){
		alert('��ǰ�ɼ��ڵ带 �Է� �ϼ���.');
		f.itemoption.value='0000';
		f.itemoption.focus();
		return false;
	}

	var url = "ajaxBrandItem.asp?shopid=<%=shopid%>&makerid=" + f.makerid.value + "&itemgubun=" + f.itemgubun.value + "&itemid=" + f.itemid.value + "&itemoption=" + f.itemoption.value;
	var xmlHttp = createXMLHttpRequest();
	xmlHttp.open("GET", url, false);
	xmlHttp.send('');
	if(xmlHttp.status == 200) {
		var arr = xmlHttp.responseText.split("|");
		//alert(xmlHttp.responseText);
		if (arr[0]=="Y")
		{
			if (mode=="GET")
			{
				f.itemname.value		= arr[1];
				f.itemoptionname.value	= arr[2];
				f.orgsellcash.value		= arr[3];
				f.sellcash.value		= arr[4];
				f.buycash.value		    = arr[5];
				f.suplycash.value		= arr[6];

			}
			return true;
		}
		else if (arr[0]=="N")
		{
			alert("������ ��ǰ�ڵ��Դϴ�.");
			f.itemid.focus();
			return false;
		}
		else
		{
			alert("�귣��ID�� ��ġ���� �ʴ� ��ǰ�ڵ��Դϴ�.");
			f.itemid.focus();
			return false;
		}
	}
}



// ajax�� ��ü �Լ�
function createXMLHttpRequest() {
    if (window.ActiveXObject) {
        xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    else if (window.XMLHttpRequest) {
        xmlHttp = new XMLHttpRequest();
    }
	return xmlHttp;
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frm method=post action="etc_meachul_process.asp">
	<input type=hidden name="mode" value="etcsubdetailadd">
	<input type=hidden name="topidx" value="<%= topidx %>">
	<input type=hidden name="idx" value="<%= idx %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="80">����</td>
		<td bgcolor="#FFFFFF" >
			<select class="select" name="linkbaljucode">
				<option value="">�Ϲݸ���
				<option value="witak">��Ź
				<option value="bojung" selected >���κ���
				<option value="etc2">��Ÿ
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text_ro" name='makerid' value="<%= makerid %>" size=32 maxlength=30 readonly>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name='itemgubun' value="" size=2 maxlength=2 onkeydown="checkBrandItem();">-
		<input type="text" class="text" name='itemid' value="" size=9 maxlength=9 onkeydown="checkBrandItem();">-
		<input type="text" class="text" name='itemoption' value="" size=4 maxlength=4 onkeydown="checkBrandItem();">
		<input type="button" class="button" value="�귣���ǰüũ" onclick="ajaxBrandItem('GET');">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='itemname' value="" size=32 maxlength=32></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ�</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='itemoptionname' value="" size=32 maxlength=32></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='itemno' value="" size=4 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�Һ�</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='orgsellcash' value="" size=10 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�ǸŰ�</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='sellcash' value="" size=10 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���԰�</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='buycash' value="" size=10 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���ް�</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='suplycash' value="" size=10 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" colspan=2 align=center>
			<input type="button" class="button" value="�����߰�" onclick="AddValue(frm)">
		</td>
	</tr>
	</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
