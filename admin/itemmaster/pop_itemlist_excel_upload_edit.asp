<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ��ǰ ���� ���ε� �ϰ� ����
' History : 2019.04.18 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemedit_temp_cls.asp"-->
<%
dim mode, i, failtype, chk_idx, chk_idx_fail
	mode 			= requestCheckVar(request("mode"),32)
	chk_idx 		= request("chk_idx")
	chk_idx_fail 		= request("chk_idx_fail")

dim oCManualMeachul
set oCManualMeachul = new Citemedit_templist
	oCManualMeachul.FPageSize = 1000
	oCManualMeachul.FCurrPage = 1
	oCManualMeachul.FRectRegAdminID = session("ssBctId")
	oCManualMeachul.FRectExcludeRegFinish = "Y"
	oCManualMeachul.GetsuccessitemList

dim oCFailManualMeachul
set oCFailManualMeachul = new Citemedit_templist
	oCFailManualMeachul.FPageSize = 1000
	oCFailManualMeachul.FCurrPage = 1
	oCFailManualMeachul.FRectRegAdminID = session("ssBctId")
	oCFailManualMeachul.GetFailList
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript">

function fnChkFile(sFile, arrExt){
    //���� ���ε� ����Ȯ��
     if (!sFile){
    	 return true;
    	}

    var blnResult = false;

    //���� Ȯ���� Ȯ��
	var pPoint = sFile.lastIndexOf('.');
	var fPoint = sFile.substring(pPoint+1,sFile.length);
	var fExet = fPoint.toLowerCase();

	for (var i = 0; i < arrExt.length; i++)
	   	{
	    	if (arrExt[i].toLowerCase() == fExet)
	    	{
	   			blnResult =  true;
	   		}
		}

	return blnResult;
}

function XLSumbit(){
	document.domain = '10x10.co.kr';
	var frm = document.frmFile;
    
	arrFileExt = new Array();			
	//arrFileExt[arrFileExt.length]  = "csv";     //xls
	arrFileExt[arrFileExt.length]  = "xls";     //csv
	
	if (frm.sFile.value==''){
		alert('������ �Է��� �ּ���');
		frm.sFile.focus();
		return;
	}
	
	//������ȿ�� üũ
	if (!fnChkFile(frm.sFile.value, arrFileExt)){
		//alert("������ csv���ϸ� ���ε� �����մϴ�.");   // xls
		alert("������ xls���ϸ� ���ε� �����մϴ�.");   // csv
		return;
	}

	frm.target='view';
	frm.submit();
}

function CheckAll(chk) {
	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			return;
		}

		if (v.disabled != true) {
			v.checked = chk.checked;
		}
	}
}

// ���ε� �����ϱ�
function delClick(mode) {
	var frm = document.frmdetail;

	if (mode=='delitem_fail'){
		if ($('input[name="chk_idx_fail"]:checked').length == 0) {
			alert('���� �������� �����ϴ�.');
			return;
		}
	}else{
		if ($('input[name="chk_idx"]:checked').length == 0) {
			alert('���� �������� �����ϴ�.');
			return;
		}
	}
	if (confirm("���� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value=mode;
		frm.action="/admin/itemmaster/pop_itemlist_excel_upload_process.asp"
		frm.target="view";
		frm.submit();
	}
}

function toggleChecked(status) {
    $('[name="chk_idx"]').each(function () {
        $(this).prop("checked", status);
    });
}
function toggleChecked_fail(status) {
    $('[name="chk_idx_fail"]').each(function () {
        $(this).prop("checked", status);
    });
}

// ��������
function saveClick() {
	var frm = document.frmdetail;

	if ($('input[name="chk_idx"]:checked').length == 0) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	if (confirm("���� �ϼ̽��ϱ�?\n���� ���� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value='edittemporder';
		frm.action="/admin/itemmaster/pop_itemlist_excel_upload_process.asp"
		frm.target="view";
		frm.submit();
	}
}

$(document).ready(function () {
    var checkAllBox = $("#chkall");
    checkAllBox.click(function () {
        var status = checkAllBox.prop('checked');
        toggleChecked(status);
    });
    var checkAllBox_fail = $("#chkall_fail");
    checkAllBox_fail.click(function () {
        var status_fail = checkAllBox_fail.prop('checked');
        toggleChecked_fail(status_fail);
    });
});

</script>

<form name="frmFile" method="post" action="<%= uploadUrl %>/linkweb/item/upload_itemlist_excel.asp"  enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>�¶��� ��ǰ ���� �ϰ� ���ε� ����</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">����</td>
	<td align="left">
		���ε� ����� ��ǰ��Ͽ� �ִ� "��ǰ�ٿ�ε�(����)" �Դϴ�. ������ <font color="red"><b>Save As Excel 97 -2003 ���չ���</b></font>�� ������ ���ε� ���ּ���.
		<% '<br>�ִ� <font color="red"><strong>1000��</strong></font> ��ǰ�� ���ε� ���� %>
		<!--<a href="<%'= uploadUrl %>/offshop/sample/item/item_list_sample_v1.csv" target="_blank">�ٿ�ε�</a>-->
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><font color="red"><b>���ǻ���</b></font></td>
	<td align="left">
		�� �������� ���������� <font color="red"><b>Save As Excel 97 -2003 ���չ���</b></font> ���¸� �ν� �մϴ�.
		<br><br>* �� ����� ��ǰ�ڵ�, �귣�� ���� �ִ� <font color="red"><b>1���� �״�� �μ���.</b></font>
		<br><br>* ��ǰ�ڵ带 �������� ������Ʈ �Ǳ� ������ <font color="red"><strong>��ǰ�ڵ�� ���� ����</strong></font>�̰ų�, Ʋ���� �ȵ˴ϴ�.
		<br><br><font color="red"><b>* ��ǰ��,�Һ��ڰ�(���԰��� �귣�� �⺻������ ���� �ڵ����),ISBN13,ǥ�ú귣��</b></font> �ʵ带 �Է��Ͻø� �Ǹ�, �״�� ���� �˴ϴ�.
		<br><br><font color="red"><strong>* ������</strong></font>�� ��ǰ�� ��ϵ��� �ʽ��ϴ�.
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>���ϸ�:</td>
	<td align="left"><input type="file" name="sFile" class="button"></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
	    <input type="button" class="button" value="���ε�" onClick="XLSumbit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	</td>
</tr>
</table>
</form>

<form name="frmdetail" method="post" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<% if oCManualMeachul.FResultCount > 0 then %>
	<Br>
	[���ε峻��]
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="chkall" id="chkall"></td>
		<td width="60">��ȣ</td>
		<td width="60">��ǰ�ڵ�</td>
		<td>�����ǰ��</td>
		<td>������ǰ��</td>
		<td width="60">����<br>�Һ��ڰ�</td>
		<td width="60">����<br>�Һ��ڰ�</td>
		<td width="100">����<br>isbn13</td>
		<td width="100">����<br>isbn13</td>
		<td width="100">����<br>ǥ�ú귣��</td>
		<td width="100">����<br>ǥ�ú귣��</td>
		<td width="80">�����</td>
		<td width="70">����</td>
		<td width="70">���</td>
	</tr>

	<% if oCManualMeachul.FResultCount > 0 then %>
		<% For i = 0 To oCManualMeachul.FResultCount - 1 %>
		<% if IsNull(oCManualMeachul.FItemList(i).Ffailtype) then %>
		<tr align="center" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" bgcolor="#CCCCCC">
		<% end if %>
			<td><input type="checkbox" name="chk_idx" value="<%= oCManualMeachul.FItemList(i).Fidx %>" ></td>
			<td><%= oCManualMeachul.FItemList(i).Fidx %></td>
			<td><%= oCManualMeachul.FItemList(i).Fitemid %></td>
			<td align="left"><%= oCManualMeachul.FItemList(i).Fitemname_10x10 %></td>
			<td align="left"><%= oCManualMeachul.FItemList(i).Fitemname %></td>
			<td align="right"><%= FormatNumber(oCManualMeachul.FItemList(i).forgprice_10x10, 0) %></td>
			<td align="right">
				<%= FormatNumber(oCManualMeachul.FItemList(i).forgprice, 0) %>

				<% if oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.1 then %>
					<Br><font color="red"><strong>90%�̻�����</strong></font>
				<% elseif oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.2 then %>
					<Br><font color="red"><strong>80%�̻�����</strong></font>
				<% elseif oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.3 then %>
					<Br><font color="red"><strong>70%�̻�����</strong></font>
				<% elseif oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.4 then %>
					<Br><font color="red"><strong>60%�̻�����</strong></font>
				<% elseif oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.5 then %>
					<Br><font color="red"><strong>50%�̻�����</strong></font>
				<% end if %>
			</td>
			<td align="left"><%= oCManualMeachul.FItemList(i).fisbn13_10x10 %></td>
			<td align="left"><%= oCManualMeachul.FItemList(i).fisbn13 %></td>
			<td><%= oCManualMeachul.FItemList(i).ffrontmakerid_10x10 %></td>
			<td><%= oCManualMeachul.FItemList(i).ffrontmakerid %></td>
			<td><%= oCManualMeachul.FItemList(i).Fregadminid %></td>
			<td><%= oCManualMeachul.FItemList(i).GetOrderTempStatusName %></td>
			<td>
				<%= oCManualMeachul.FItemList(i).GetFailTypeName %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="17">
				<input type="button" class="button" value="���û�ǰ ���� ��ǰ�� ����" onclick="saveClick()">	
				<input type="button" class="button" value="�����ϱ�" onclick="delClick('delitem');">
			</td>
		</tr>
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="17" height="35">
				���ε峻�� ����
			</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

<% if oCFailManualMeachul.FResultCount > 0 then %>
	<br>
	[���ε����]
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">��ȣ</td>
		<td width="20"><input type="checkbox" name="chkall_fail" id="chkall_fail"></td>
		<td width="60">��ǰ�ڵ�</td>
		<td>�����ǰ��</td>
		<td>������ǰ��</td>
		<td width="60">����<br>�Һ��ڰ�</td>
		<td width="80">����<br>�Һ��ڰ�</td>
		<td width="100">����<br>isbn13</td>
		<td width="100">����<br>isbn13</td>
		<td width="100">����<br>ǥ�ú귣��</td>
		<td width="100">����<br>ǥ�ú귣��</td>
		<td width="80">�����</td>
		<td width="70">����</td>
		<td width="70">���</td>
	</tr>

	<% if oCFailManualMeachul.FResultCount > 0 then %>
		<% For i = 0 To oCFailManualMeachul.FResultCount - 1 %>
		<% if IsNull(oCFailManualMeachul.FItemList(i).Ffailtype) then %>
		<tr align="center" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" bgcolor="#CCCCCC">
		<% end if %>
			<td><%= oCFailManualMeachul.FItemList(i).Fidx %></td>
			<td><input type="checkbox" name="chk_idx_fail" value="<%= oCFailManualMeachul.FItemList(i).Fidx %>" ></td>
			<td><%= oCFailManualMeachul.FItemList(i).Fitemid %></td>
			<td align="left"><%= oCFailManualMeachul.FItemList(i).Fitemname_10x10 %></td>
			<td align="left"><%= oCFailManualMeachul.FItemList(i).Fitemname %></td>
			<td align="right"><%= FormatNumber(oCFailManualMeachul.FItemList(i).forgprice_10x10, 0) %></td>
			<td align="right">
				<%= FormatNumber(oCFailManualMeachul.FItemList(i).forgprice, 0) %>

				<% if oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.1 then %>
					<Br><font color="red"><strong>90%�̻�����</strong></font>
				<% elseif oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.2 then %>
					<Br><font color="red"><strong>80%�̻�����</strong></font>
				<% elseif oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.3 then %>
					<Br><font color="red"><strong>70%�̻�����</strong></font>
				<% elseif oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.4 then %>
					<Br><font color="red"><strong>60%�̻�����</strong></font>
				<% elseif oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.5 then %>
					<Br><font color="red"><strong>50%�̻�����</strong></font>
				<% end if %>
			</td>
			<td align="left"><%= oCFailManualMeachul.FItemList(i).fisbn13_10x10 %></td>
			<td align="left"><%= oCFailManualMeachul.FItemList(i).fisbn13 %></td>
			<td><%= oCFailManualMeachul.FItemList(i).ffrontmakerid_10x10 %></td>
			<td><%= oCFailManualMeachul.FItemList(i).ffrontmakerid %></td>
			<td><%= oCFailManualMeachul.FItemList(i).Fregadminid %></td>
			<td><%= oCFailManualMeachul.FItemList(i).GetOrderTempStatusName %></td>
			<td>
				<%= oCFailManualMeachul.FItemList(i).GetFailTypeName %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="17">
				<input type="button" class="button" value="�����ϱ�" onclick="delClick('delitem_fail');">
			</td>
		</tr>
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="17" height="35">
				���ε峻�� ����
			</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

</form>

<% IF application("Svr_Info")="Dev" or C_ADMIN_AUTH THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->