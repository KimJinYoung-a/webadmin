<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [CS]���>>[1:1���]���ϴ亯��
' Hieditor : 2021.07.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/board/customer_board_reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2, tmpDateStr, startDateStr, endDateStr, chkGrpByReplyUser
Dim replyUser, i, userlevel, sitename, totalcount, d0pro, d1pro, d2pro, d3pro, d4pro, d0_1pro, d2_3_4pro
	yyyy1 = requestcheckvar(getNumeric(request("yyyy1")),4)
	mm1 = requestcheckvar(getNumeric(request("mm1")),2)
	dd1 = requestcheckvar(getNumeric(request("dd1")),2)
	yyyy2 = requestcheckvar(getNumeric(request("yyyy2")),4)
	mm2 = requestcheckvar(getNumeric(request("mm2")),2)
	dd2 = requestcheckvar(getNumeric(request("dd2")),2)
	userlevel = requestcheckvar(getNumeric(request("userlevel")),10)
	sitename = requestcheckvar(request("sitename"),32)

chkGrpByReplyUser	= req("chkGrpByReplyUser","")
replyUser	= req("replyUser","")

if (yyyy1="") then
	chkGrpByReplyUser = "Y"
	startdateStr = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	startdateStr = DateSerial(yyyy1, mm1, dd1)
end if

yyyy1 = left(startdateStr,4)
mm1 = Mid(startdateStr,6,2)
dd1 = Mid(startdateStr,9,2)
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

endDateStr = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CReportMaster
	oreport.FRectuserlevel = userlevel
	oreport.FRectStart = startdateStr
	oreport.FRectEnd = endDateStr
	oreport.FRectReplyUser = replyUser
	oreport.FRectGroupByReplyUser = chkGrpByReplyUser
	oreport.FRectsitename = sitename
	oreport.FPageSize = 1000
	oreport.FCurrPage = 1
	oreport.getsameday_report

%>
<script type="text/javascript">

function searchSubmit(){
	//��¥ ��
	var startdate = frm.yyyy1.value + "-" + frm.mm1.value + "-" + frm.dd1.value;
	var enddate = frm.yyyy2.value + "-" + frm.mm2.value + "-" + frm.dd2.value;
    var diffDay = 0;
    var start_yyyy = startdate.substring(0,4);
    var start_mm = startdate.substring(5,7);
    var start_dd = startdate.substring(8,startdate.length);
    var sDate = new Date(start_yyyy, start_mm-1, start_dd);
    var end_yyyy = enddate.substring(0,4);
    var end_mm = enddate.substring(5,7);
    var end_dd = enddate.substring(8,enddate.length);
    var eDate = new Date(end_yyyy, end_mm-1, end_dd);

    diffDay = Math.ceil((eDate.getTime() - sDate.getTime())/(1000*60*60*24));

	if (diffDay > 31){
		alert('�˻��Ⱓ�� 1�� ������ �˻� ���� �մϴ�.');
		return;
	}

	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ(�亯�ϱ���) <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		* �亯��ID : <input type="text" class="text" name="replyUser" value="<%=replyUser%>" size="12" maxlength="32">
		&nbsp;
		<input type="checkbox" name="chkGrpByReplyUser" value="Y" <%if (chkGrpByReplyUser = "Y") then %>checked<% end if %> > �亯��ID ǥ��
		&nbsp;
		* ȸ����� : <% DrawselectboxUserLevel "userlevel", userlevel, "" %>
		&nbsp;
		* �Ǹ�ó : 
		<select name="sitename">
			<option value="" <% if sitename="" then response.write " selected" %>>��ü</option>
			<option value="10x10" <% if sitename="10x10" then response.write " selected" %>>10x10</option>
			<option value="10x10not" <% if sitename="10x10not" then response.write " selected" %>>���޸�</option>
		</select>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="searchSubmit();">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�� 1�ð� ���� ������ �Դϴ�. ���ϰ� �ִ� �Ŵ� �Դϴ�. �˻��Ͻ��� ���� ������ ���ð� ��ٷ� �ּ���.
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oreport.FresultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (chkGrpByReplyUser = "Y") then %>
		<td rowspan=2>�亯��ID</td>
	<% end if %>

	<td rowspan=2>�亯����</td>
	<td rowspan=2>����</td>
	<td colspan=2>��������</td>
	<td colspan=3>���ع̴�</td>
	<td rowspan=2>�հ�</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>D+0</td>
	<td>D+1</td>
	<td>D+2</td>
	<td>D+3</td>
	<td>D+4 �̻�</td>
</tr>

<% if oreport.FresultCount > 0 then %>
	<%
	for i=0 to oreport.FresultCount -1
	totalcount=oreport.FItemList(i).fd0+oreport.FItemList(i).fd1+oreport.FItemList(i).fd2+oreport.FItemList(i).fd3+oreport.FItemList(i).fd4
	d0pro=(oreport.FItemList(i).fd0 / totalcount)*100
	d1pro=(oreport.FItemList(i).fd1 / totalcount)*100
	d2pro=(oreport.FItemList(i).fd2 / totalcount)*100
	d3pro=(oreport.FItemList(i).fd3 / totalcount)*100
	d4pro=(oreport.FItemList(i).fd4 / totalcount)*100
	d0_1pro=((oreport.FItemList(i).fd0+oreport.FItemList(i).fd1)/totalcount)*100
	d2_3_4pro=((oreport.FItemList(i).fd2+oreport.FItemList(i).fd3+oreport.FItemList(i).fd4)/totalcount)*100
	%>
	<tr align="center" bgcolor="FFFFFF">
		<% if (chkGrpByReplyUser = "Y") then %>
			<td rowspan=3><%= oreport.FItemList(i).freplyuser %></td>
		<% end if %>

		<td rowspan=3><%= oreport.FItemList(i).freplydate %></td>
		<td>�亯�Ǽ�</td>
		<td><%= oreport.FItemList(i).fd0 %></td>
		<td><%= oreport.FItemList(i).fd1 %></td>
		<td><%= oreport.FItemList(i).fd2 %></td>
		<td><%= oreport.FItemList(i).fd3 %></td>
		<td><%= oreport.FItemList(i).fd4 %></td>
		<td><%= CurrFormat(totalcount) %></td>
	</tr>
	<tr align="center" bgcolor="FFFFFF">
		<td>����</td>
		<td><%= round(d0pro,2) %>%</td>
		<td><%= round(d1pro,2) %>%</td>
		<td><%= round(d2pro,2) %>%</td>
		<td><%= round(d3pro,2) %>%</td>
		<td><%= round(d4pro,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr align="center" bgcolor="FFFFFF">
		<td>�հ�</td>
		<td colspan=2><%= round(d0_1pro,2) %>%</td>
		<td colspan=3><%= round(d2_3_4pro,2) %>%</td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
