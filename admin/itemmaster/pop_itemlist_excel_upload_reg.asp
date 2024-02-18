<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ��ǰ�űԵ��. ���� �ϰ�
' History : 2019.12.13 �ѿ�� ����
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
	oCManualMeachul.GetsuccessitemregList

dim oCFailManualMeachul
set oCFailManualMeachul = new Citemedit_templist
	oCFailManualMeachul.FPageSize = 1000
	oCFailManualMeachul.FCurrPage = 1
	oCFailManualMeachul.FRectRegAdminID = session("ssBctId")
	oCFailManualMeachul.GetregFailList
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
	arrFileExt[arrFileExt.length]  = "xls";     //csv
	
	if (frm.sFile.value==''){
		alert('������ �Է��� �ּ���');
		frm.sFile.focus();
		return;
	}
	
	//������ȿ�� üũ
	if (!fnChkFile(frm.sFile.value, arrFileExt)){
		alert("������ xls���ϸ� ���ε� �����մϴ�.");
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

	if (mode=='delregitem_fail'){
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
	if ($('input[name="chk_idx"]:checked').length > 100) {
		alert('�ѹ��� 100�پ� ���� ���� �մϴ�.');
		return;
	}

	if (confirm("���� �ϼ̽��ϱ�?\n���� ��ǰ���� ���� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value='regtemporder';
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

<form name="frmFile" method="post" action="<%= uploadUrl %>/linkweb/item/upload_itemlistreg_excel.asp"  enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>�¶��� ��ǰ �űԵ�� ���� �ϰ� ���ε�</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">����</td>
	<td align="left">
		���ε� ��� : <a href="<%= uploadUrl %>/offshop/sample/item/sample_product_new_v4.xls" target="_blank">�ٿ�ε�</a>
        <br>������ <font color="red"><b>Save As Excel 97 -2003 ���չ���</b></font>�� ������ ���ε� ���ּ���.
        <!--<br>�ִ� <font color="red"><strong>1000��</strong></font> ��ǰ�� ���ε� ����-->
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><font color="red"><b>���ǻ���</b></font></td>
	<td align="left">
		* �� ����� <font color="red"><b>ù��</b></font>�� �������� ������.
		<br><br>* <font color="red"><b>�ɼ��� �ִ� ��ǰ</b></font>�� ��� ���� ������ ���� �ϼż�, ���� ��ǰ �ɼ� ������� �ؿ��ٿ� �̾ �Է��� �ֽø� �˴ϴ�.
        <br>&nbsp;&nbsp;
        ������ǰ�� �ι�° �ɼǺ��ʹ� ��ǰ����(�귣��ID,��ǰ��,�⺻����ī�װ��ڵ�,�Һ��ڰ�,���԰�,�ŷ�����,��۱���,������,������,�˻�Ű����,ǥ�ú귣��)�� �������� ��� �νð�,
        <br>&nbsp;&nbsp;
        �ɼǸ�,������ڵ�,��ü�����ڵ�,��ǰ������,��ǰ����,[���Կ�]��ǰ��,[���Կ�]��ȭȭ��,[���Կ�]���԰�,[���Կ�]�ɼǸ� �� �Է��� �ֽø� �˴ϴ�.
		<br>&nbsp;&nbsp;
		�������� ����ν� ������ ������ ��ǰ������ �״�� ���� �˴ϴ�.
		<br><br>* ����ī�װ��ڵ带 �Է��Ұ��, �ڵ����� ��Ī�Ǿ� ����ī�װ��ڵ尡 �Է� �˴ϴ�.
		<br><br>* �� ���� : http://confluence.tenbyten.kr:8090/pages/editpage.action?pageId=59021677
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
		<td width="90">�ӽû�ǰ�ڵ�<Br>[�ӽÿɼ��ڵ�]</td>
		<td width="100">�귣��ID</td>
		<td width="100">ǥ�ú귣��</td>
		<td width="110">����ī�װ��ڵ�<br>[����ī�װ��ڵ�]</td>
		<td>��ǰ��<br>[�ɼǸ�]</td>
		<td width="60">�Һ��ڰ�<br>[���԰�]</td>
		<td width="90">�ŷ�����<br>[��۱���]</td>
		<td width="100">������ڵ�<br>[��ü�����ڵ�]</td>
		<td width="60">����</td>
		<td width="80">�����</td>
	</tr>

	<% if oCManualMeachul.FResultCount > 0 then %>
		<% For i = 0 To oCManualMeachul.FResultCount - 1 %>
		<% if IsNull(oCManualMeachul.FItemList(i).Ffailtype) then %>
		<tr align="center" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" bgcolor="#CCCCCC">
		<% end if %>
			<td><input type="checkbox" name="chk_idx" value="<%= oCManualMeachul.FItemList(i).Fidx %>" ></td>
			<td><%= oCManualMeachul.FItemList(i).Ftempitemid %><br>[<%= oCManualMeachul.FItemList(i).Ftempitemoption %>]</td>
			<td><%= oCManualMeachul.FItemList(i).Fmakerid %></td>
			<td><%= oCManualMeachul.FItemList(i).ffrontmakerid %></td>
			<td align="left">
				<%= oCManualMeachul.FItemList(i).Fdispcatecode %>
				<br>[<% if oCManualMeachul.FItemList(i).fcate_large<>"" and not(isnull(oCManualMeachul.FItemList(i).fcate_large)) then %>
					<%= oCManualMeachul.FItemList(i).fcate_large %>
				<% end if %>
				<% if oCManualMeachul.FItemList(i).fcate_mid<>"" and not(isnull(oCManualMeachul.FItemList(i).fcate_mid)) then %>
					<%= oCManualMeachul.FItemList(i).fcate_mid %>
				<% end if %>
				<% if oCManualMeachul.FItemList(i).fcate_small<>"" and not(isnull(oCManualMeachul.FItemList(i).fcate_small)) then %>
					<%= oCManualMeachul.FItemList(i).fcate_small %>
				<% end if %>]
			</td>
			<td align="left">
				<%= oCManualMeachul.FItemList(i).Fitemname %>
				<% if oCManualMeachul.FItemList(i).Fitemoptionname<>"" and not(isnull(oCManualMeachul.FItemList(i).Fitemoptionname)) then %>
					<br>[<%= oCManualMeachul.FItemList(i).Fitemoptionname %>]
				<% end if %>
			</td>
			<td align="left"><%= FormatNumber(oCManualMeachul.FItemList(i).forgprice, 0) %><br>[<%= FormatNumber(oCManualMeachul.FItemList(i).Fbuycash, 0) %>]</td>
			<td><%= mwdivName(oCManualMeachul.FItemList(i).Fmwdiv) %><br>[<%= getdeliverytypename(oCManualMeachul.FItemList(i).Fdeliverytype) %>]</td>
			<td align="left">
				<%= oCManualMeachul.FItemList(i).Fbarcode %>
				<% if oCManualMeachul.FItemList(i).Fupchemanagecode<>"" and not(isnull(oCManualMeachul.FItemList(i).Fupchemanagecode)) then %>
					<br>[<%= oCManualMeachul.FItemList(i).Fupchemanagecode %>]
				<% end if %>
			</td>
			<td>
				<%= oCManualMeachul.FItemList(i).GetOrderTempStatusName %>
				<% if oCManualMeachul.FItemList(i).GetFailTypeName<>"" then %>
					<br><%= oCManualMeachul.FItemList(i).GetFailTypeName %>
				<% end if %>
			</td>
			<td>
				<%= oCManualMeachul.FItemList(i).Fregadminid %>
				<Br><%= left(oCManualMeachul.FItemList(i).Fregdate,10) %>
				<Br><%= mid(oCManualMeachul.FItemList(i).Fregdate,11,12) %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="16">
				<input type="button" class="button" value="���û�ǰ ���� ��ǰ���� ���(100�پ�)" onclick="saveClick()">	
				<input type="button" class="button" value="�����ϱ�" onclick="delClick('delregitem');">
			</td>
		</tr>
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="16" height="35">
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
		<td width="20"><input type="checkbox" name="chkall_fail" id="chkall_fail"></td>
		<td width="90">�ӽ�>��ǰ�ڵ�<Br>[�ӽÿɼ��ڵ�]</td>
		<td width="100">�귣��ID</td>
		<td width="100">ǥ�ú귣��</td>
		<td width="110">����ī�װ��ڵ�<br>[����ī�װ��ڵ�]</td>
		<td>��ǰ��<br>[�ɼǸ�]</td>
		<td width="60">�Һ��ڰ�<br>[���԰�]</td>
		<td width="90">�ŷ�����<br>[��۱���]</td>
		<td width="100">������ڵ�<br>[��ü�����ڵ�]</td>
		<td width="60">����</td>
		<td width="80">�����</td>
	</tr>

	<% if oCFailManualMeachul.FResultCount > 0 then %>
		<% For i = 0 To oCFailManualMeachul.FResultCount - 1 %>
		<% if IsNull(oCFailManualMeachul.FItemList(i).Ffailtype) then %>
		<tr align="center" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" bgcolor="#CCCCCC">
		<% end if %>
			<td><input type="checkbox" name="chk_idx_fail" value="<%= oCFailManualMeachul.FItemList(i).Fidx %>" ></td>
			<td><%= oCFailManualMeachul.FItemList(i).Ftempitemid %><Br>[<%= oCFailManualMeachul.FItemList(i).Ftempitemoption %>]</td>
			<td><%= oCFailManualMeachul.FItemList(i).Fmakerid %></td>
			<td><%= oCFailManualMeachul.FItemList(i).ffrontmakerid %></td>
			<td align="left">
				<%= oCFailManualMeachul.FItemList(i).Fdispcatecode %>
				<br>[<% if oCFailManualMeachul.FItemList(i).fcate_large<>"" and not(isnull(oCFailManualMeachul.FItemList(i).fcate_large)) then %>
					<%= oCFailManualMeachul.FItemList(i).fcate_large %>
				<% end if %>
				<% if oCFailManualMeachul.FItemList(i).fcate_mid<>"" and not(isnull(oCFailManualMeachul.FItemList(i).fcate_mid)) then %>
					<%= oCFailManualMeachul.FItemList(i).fcate_mid %>
				<% end if %>
				<% if oCFailManualMeachul.FItemList(i).fcate_small<>"" and not(isnull(oCFailManualMeachul.FItemList(i).fcate_small)) then %>
					<%= oCFailManualMeachul.FItemList(i).fcate_small %>
				<% end if %>]
			</td>
			<td align="left">
				<%= oCFailManualMeachul.FItemList(i).Fitemname %>
				<% if oCFailManualMeachul.FItemList(i).Fitemoptionname<>"" and not(isnull(oCFailManualMeachul.FItemList(i).Fitemoptionname)) then %>
					<br>[<%= oCFailManualMeachul.FItemList(i).Fitemoptionname %>]
				<% end if %>
			</td>
			<td align="left"><%= FormatNumber(oCFailManualMeachul.FItemList(i).forgprice, 0) %><br>[<%= FormatNumber(oCFailManualMeachul.FItemList(i).Fbuycash, 0) %>]</td>
			<td><%= mwdivName(oCFailManualMeachul.FItemList(i).Fmwdiv) %><br>[<%= getdeliverytypename(oCFailManualMeachul.FItemList(i).Fdeliverytype) %>]</td>
			<td align="left">
				<%= oCFailManualMeachul.FItemList(i).Fbarcode %>
				<% if oCFailManualMeachul.FItemList(i).Fupchemanagecode<>"" and not(isnull(oCFailManualMeachul.FItemList(i).Fupchemanagecode)) then %>
					<br>[<%= oCFailManualMeachul.FItemList(i).Fupchemanagecode %>]
				<% end if %>
			</td>
			<td>
				<%= oCFailManualMeachul.FItemList(i).GetOrderTempStatusName %>
				<% if oCFailManualMeachul.FItemList(i).GetFailTypeName<>"" then %>
					<br><%= oCFailManualMeachul.FItemList(i).GetFailTypeName %>
				<% end if %>
			</td>
			<td>
				<%= oCFailManualMeachul.FItemList(i).Fregadminid %>
				<Br><%= left(oCFailManualMeachul.FItemList(i).Fregdate,10) %>
				<Br><%= mid(oCFailManualMeachul.FItemList(i).Fregdate,11,12) %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="16">
				<input type="button" class="button" value="�����ϱ�" onclick="delClick('delregitem_fail');">
			</td>
		</tr>
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="16" height="35">
				���ε峻�� ����
			</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width=1280 height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=1280 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
function getdeliverytypename(vdeliverytype)
    dim deliverytypename

    if vdeliverytype="1" then
        deliverytypename="�ٹ����ٹ��"
    elseif vdeliverytype="2" then
        deliverytypename="��ü���"
    else
        deliverytypename=""
    end if

    getdeliverytypename=deliverytypename
end function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->