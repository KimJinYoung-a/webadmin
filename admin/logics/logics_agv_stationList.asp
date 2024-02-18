<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : agv
' History : �̻� ����
'           2020.05.12 ������ ����
'           2020.05.20 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<%

dim research, page
dim stationGubun, i, j, k

research = requestCheckvar(request("research"),32)
page = requestCheckvar(request("page"),32)
stationGubun = requestCheckvar(request("stationGubun"),32)

if (research = "") then
    page = 1
    stationGubun = "PICK"
end if

dim oAGVStation
Set oAGVStation = new CAGVItems
    oAGVStation.FPageSize = 500
    oAGVStation.FCurrPage = page

    oAGVStation.GetStationList

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function SubmitFrm(frm) {
	frm.submit();
}

function jsAddStation() {
    var popwin = window.open('logics_agv_stationPop.asp','jsAddStation','width=400,height=170,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function jsModiStation(stationCd) {
    var popwin = window.open('logics_agv_stationPop.asp?stationCd=' + stationCd,'jsModiStation','width=400,height=170,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function uploadexcel(){
	document.domain = "10x10.co.kr";
	var popwin = window.open('/admin/logics/logics_agv_pickup_excel_upload.asp','adduploadexcel','width=1280,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>
<!-- ǥ ��ܹ� ����-->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="3" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
    <td align="left">
        �����̼Ǳ��� :
        <select class="select" name="stationGubun">
            <option></option>
            <option value="PICK" <%= CHKIIF(stationGubun="PICK", "selected", "") %>>��ŷ �����̼�</option>
            <option value="IPGO" <%= CHKIIF(stationGubun="IPGO", "selected", "") %>>�԰� �����̼�</option>
        </select>
    </td>
    <td rowspan="1" width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitFrm(document.frm);">
	</td>
</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->

<br />

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" value="�߰�" onclick="jsAddStation();" class="button" >
	</td>
</tr>
</table>
<!-- �׼� �� -->

<br />

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		�˻���� : <b><%= oAGVStation.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">�����̼��ڵ�</td>
	<td width="200">�����̼Ǹ�</td>
	<td width="100">�����̼Ǳ���</td>
    <td width="80">ǥ�ü���</td>
	<td width="150">�����</td>
    <td width="150">��������</td>
	<td>���</td>
</tr>
<% if oAGVStation.FResultCount >0 then %>
	<% for i=0 to oAGVStation.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=24>
		<td align=center>
		  	<%= oAGVStation.FItemList(i).FstationCd %>
		</td>
		<td align="center">
		  	<a href="javascript:jsModiStation('<%= oAGVStation.FItemList(i).FstationCd %>')"><%= oAGVStation.FItemList(i).FstationName %></a>
		</td>
		<td align="center">
		  	<%= oAGVStation.FItemList(i).getStationGubunName %>
		</td>
		<td align="center">
		  	<%= oAGVStation.FItemList(i).FsortNo %>
		</td>
		<td align="center">
		  	<%= oAGVStation.FItemList(i).Fregdate %>
		</td>
		<td align="center">
		  	<%= oAGVStation.FItemList(i).Fupdt %>
		</td>
		<td align=center>
	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="17" align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>
</table>

<%
set oAGVStation = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
