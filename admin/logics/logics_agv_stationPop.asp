<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<%

dim stationCd
dim IsEditState : IsEditState = False

stationCd = requestCheckvar(request("stationCd"),32)

dim oAGVStation
Set oAGVStation = new CAGVItems

if (stationCd <> "") then
    IsEditState = True
    oAGVStation.FRectStationCd = stationCd
    oAGVStation.GetStationOne
else
    oAGVStation.GetStationOneEmpty
end if

%>
<script type="text/javascript">

function checkValue(obj, re, errMsg) {
    var str = obj.value;
    if (re.test(str) != true) {
        alert(errMsg);
        obj.focus();
        return false;
    }
    return true;
}

function SubmitFrm(frm) {
    var obj;
    var str;
    var re;

    if (checkValue(frm.stationCd, /^[0-9A-Z]{4}$/i, '�����̼��ڵ�� 4�ڸ� ���� �Ǵ� �����빮�ڸ� �����մϴ�.') != true) { return; }
    if (checkValue(frm.stationName, /.+/i, '�����̼Ǹ��� �Է��ϼ���.') != true) { return; }
    if (checkValue(frm.stationGubun, /.+/i, '�����̼Ǳ����� �Է��ϼ���.') != true) { return; }
    if (checkValue(frm.sortNo, /^[0-9]+$/i, 'ǥ�ü����� ���ڸ� �����մϴ�.') != true) { return; }

    if (confirm('�����Ͻðڽ��ϱ�?') == true) {
        frm.submit();
    }
}

function SubmitDel() {
    if (confirm('������ �����Ͻðڽ��ϱ�?') == true) {
        document.frm.mode.value = 'delstation';
        frm.submit();
    }
}

</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-left: 3px;">
<form name="frm" onsubmit="return false;" action="logics_agv_station_process.asp">
<input type="hidden" name="mode" value="<%= CHKIIF(IsEditState, "editstation", "addstation") %>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�����̼� ����</b>
			    </td>
			    <td align="right">

			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>" width="100">�����̼��ڵ�</td>
    <td>
        <input type="text" class="<%= CHKIIF(IsEditState, "text_ro", "text") %>" name="stationCd" value="<%= oAGVStation.FOneItem.FstationCd %>" size="4" <%= CHKIIF(IsEditState, "readOnly", "") %>>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>" width="100">�����̼Ǹ�</td>
    <td>
        <input type="text" class="text" name="stationName" value="<%= oAGVStation.FOneItem.FstationName %>" size="24">
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>" width="100">�����̼Ǳ���</td>
    <td>
        <select class="select" name="stationGubun">
            <option></option>
            <option value="PICK" <%= CHKIIF(oAGVStation.FOneItem.FstationGubun="PICK", "selected", "") %>>��ŷ �����̼�</option>
            <option value="IPGO" <%= CHKIIF(oAGVStation.FOneItem.FstationGubun="IPGO", "selected", "") %>>�԰� �����̼�</option>
        </select>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>" width="100">ǥ�ü���</td>
    <td>
        <input type="text" class="text" name="sortNo" value="<%= oAGVStation.FOneItem.FsortNo %>" size="2">
    </td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
    <td colspan="2" align="center">
        <input type="button" value="�����ϱ�" class="csbutton" onclick="SubmitFrm(document.frm);">
        <% if IsEditState = True then %>
        &nbsp;
        <input type="button" value="����" class="csbutton" onclick="SubmitDel(document.frm);">
        <% end if %>
    </td>
</tr>
</table>

<%
set oAGVStation = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
