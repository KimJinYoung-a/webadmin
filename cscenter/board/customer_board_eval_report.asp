<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ [1:1���]�� ������ ���
' Hieditor : �̻� ����
'			 2021.03.03 �ѿ�� ����
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
Dim replyUser, i, userlevel
	yyyy1 = requestcheckvar(request("yyyy1"),4)
	mm1 = requestcheckvar(request("mm1"),2)
	dd1 = requestcheckvar(request("dd1"),2)
	yyyy2 = requestcheckvar(request("yyyy2"),4)
	mm2 = requestcheckvar(request("mm2"),2)
	dd2 = requestcheckvar(request("dd2"),2)
	userlevel = requestcheckvar(getNumeric(request("userlevel")),10)

chkGrpByReplyUser	= req("chkGrpByReplyUser","")
replyUser	= req("replyUser","")

if (yyyy1="") then
	chkGrpByReplyUser = "Y"

	yyyy1 = Year(now())
	mm1 = Month(now()) - 1
	dd1 = 1

	yyyy2 = Year(now())
	mm2 = Month(now())
	dd2 = 1

	startdateStr = Left(CStr(DateSerial(yyyy1,mm1,dd1)), 10)
	yyyy1 = Left(startdateStr, 4)
	mm1 = Right(Left(startdateStr, 7), 2)
	dd1 = Right(startdateStr, 2)

	endDateStr = Left(CStr(DateSerial(yyyy2,mm2,dd2)), 10)
	tmpDateStr = Left(CStr(DateSerial(yyyy2,mm2,dd2 - 1)), 10)

	yyyy2 = Left(tmpDateStr, 4)
	mm2 = Right(Left(tmpDateStr, 7), 2)
	dd2 = Right(tmpDateStr, 2)
else
	startdateStr = Left(CStr(DateSerial(yyyy1,mm1,dd1)), 10)
	endDateStr = Left(CStr(DateSerial(yyyy2,mm2,dd2 + 1)), 10)
end if


dim oreport
set oreport = new CReportMaster
	oreport.FRectuserlevel = userlevel
	oreport.FRectStart = startdateStr
	oreport.FRectEnd = endDateStr
	oreport.FRectReplyUser = replyUser
	oreport.FRectGroupByReplyUser = chkGrpByReplyUser
	oreport.FPageSize = 200
	oreport.FCurrPage = 1
	oreport.getQnaEvalReport

dim flashvar
flashvar = "startdate=" + startdateStr + "&enddate=" + endDateStr
%>
<script type="text/javascript">

/*
function popQna(qaDiv, replyDate, replyUser)
{
	if (replyDate)
	{
		var replyDate1 = replyDate;
		var replyDate2 = replyDate;
	}
	else
	{
		var f = document.frm;
		var replyDate1 = f.yyyy1.value + "-" + f.mm1.value + "-" + f.dd1.value;
		var replyDate2 = f.yyyy2.value + "-" + f.mm2.value + "-" + f.dd2.value;
	}
	var url = "/cscenter/board/cscenter_qna_board_list.asp?ckReplyDateDefault=on&qaDiv=" + qaDiv + "&writeid=" + replyUser + "&replyDate1=" + replyDate1 + "&replyDate2=" + replyDate2;
	var popwin = window.open(url,"PopMyQnaList","width=1024, height=768, left=50, top=50, scrollbars=yes, resizable=yes, status=yes");
	popwin.focus();
}
*/

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
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
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
		* �亯 ��� �� ������ 1:1 ��㵵 ��迡 �ջ�˴ϴ�.
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= oreport.FresultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="120">�亯��ID</td>
	<td width="80">�亯����</td>
	<td width="60">�Ѵ亯<br>�Ǽ�</td>
	<td width="60">��<br>�򰡼�</td>
	<td width="60">5��</td>
	<td width="60">4��</td>
	<td width="60">3��</td>
	<td width="60">2��</td>
	<td width="60">1��</td>
	<td width="60">�亯����</td>
	<td width="60">������</td>
	<td width="60">���</td>
	<td>���</td>
</tr>

<% if oreport.FresultCount > 0 then %>
	<% for i=0 to oreport.FresultCount -1 %>
	<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF"; height="25">
		<td><%= oreport.FItemList(i).Freplyuser %></td>
		<td><%= oreport.FItemList(i).Freplydate %></td>
		<td align="right"><%= oreport.FItemList(i).FtotCnt %></td>
		<td align="right"><%= oreport.FItemList(i).FtotEvalCnt %></td>
		<td align="right"><%= oreport.FItemList(i).FevalCnt5 %></td>
		<td align="right"><%= oreport.FItemList(i).FevalCnt4 %></td>
		<td align="right"><%= oreport.FItemList(i).FevalCnt3 %></td>
		<td align="right"><%= oreport.FItemList(i).FevalCnt2 %></td>
		<td align="right"><%= oreport.FItemList(i).FevalCnt1 %></td>
		<td align="right"><%= oreport.FItemList(i).FnoEvalCnt %></td>
		<td align="right"><%= oreport.FItemList(i).FevalSum %></td>
		<td align="right">
			<% if (oreport.FItemList(i).FtotEvalCnt = 0) then %>
			0
			<% else %>
			<%= FormatNumber((oreport.FItemList(i).FevalSum / oreport.FItemList(i).FtotEvalCnt), 2) %>
			<% end if %>

		</td>
		<td></td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
