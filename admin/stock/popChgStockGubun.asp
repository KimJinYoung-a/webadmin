<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<script language="javascript">

function jsSubmitReg(frm) {
	if (frm.skuCd.value == "") {
		alert("SKU �ڵ带 �Է��ϼ���.");
		return;
	}

	if (confirm("��ȯ �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsChangeStock2BULK() {
    var url;
    var brandArray;
    var skuCdArray;
    var frm = document.frm;

    <% if not(C_ADMIN_AUTH) then %>
    alert('�׽�Ʈ���Դϴ�.');
    return;
    <% end if %>

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'http://wapi.10x10.co.kr';
    <% END IF %>

    skuCdArray = frm.skuCd.value;
    skuCdArray = skuCdArray.replace(/\n/g, ',');
    url = url + '/agv/api.asp?mode=chgwarehouse2bulk&skuCdArray=' + skuCdArray;

    if (confirm('����� ��ũ��ȯ �Ͻðڽ��ϱ�?') != true) { return; }

    $.ajax({
        url: url,
        type: 'get',
        crossDomain: true,
        data: {},
        dataType: 'json',
        success: function(data) {
            if (data.resultCode == '200') {
                alert('������Ʈ�Ǿ����ϴ�.');
            } else {
                alert(data.resultMessage);
            }
        },
        error: function(jqXHR, textStatus, ex) {
            alert(textStatus + "," + ex + "," + jqXHR.responseText);
        }
    });
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="agvnewshortagestock_process.asp">
	<input type="hidden" name="mode" value="chgStockGubun">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>SKU�ڵ�</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td>
			<textarea class="textarea" name="skuCd" cols="24" rows="8"></textarea>
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value="�����ϱ�" onClick="jsChangeStock2BULK()">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
