<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : [����]��������ƮNEW
' History	:  ���ʻ����� ��
'              2017.07.07 �ѿ�� ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventPrizeCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim eventGubun, yyyy1,yyyy2,mm1,mm2,dd1,dd2, fromDate ,toDate, tmpDate
dim userid, eventCode, eventName, research, page, i
	menupos = requestCheckVar(Request("menupos"),32)
	research = requestCheckVar(Request("research"),32)
	page = requestCheckVar(Request("page"),32)
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	eventGubun = requestCheckVar(Request("eventGubun"),32)
	userid = requestCheckVar(Request("userid"),32)
	eventCode = requestCheckVar(Request("eventCode"),32)
	eventName = requestCheckVar(Request("eventName"),32)

if (research = "") then
	eventGubun = "tbl_event"
end if

if (userid = "") and (eventGubun = "") then
	eventGubun = "tbl_event"
end if

if (page = "") then
	page = "1"
end if

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) + 1), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

dim oCEventPrize
set oCEventPrize = new CEventPrize
	oCEventPrize.FRectEventGubun = eventGubun
	oCEventPrize.FRectUserid = userid
	oCEventPrize.FRectEventCode = eventCode
	oCEventPrize.FRectEventName = eventName
	oCEventPrize.FRectStartdate = fromDate
	oCEventPrize.FRectEndDate = toDate
	oCEventPrize.frectgubun="ONEVT"
	oCEventPrize.FPageSize = 20
	oCEventPrize.FCurrPage = page
	
	if (oCEventPrize.FRectUserid <> "") and (oCEventPrize.FRectEventGubun <> "") then
		oCEventPrize.GetUserEventJoinListNew
	else
		oCEventPrize.GetEventJoinListNew
	end if
%>

<script language='javascript'>

function fnGotoPage(page) {
	document.frm.page.value = page;
	document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		�̺�Ʈ���� :
		<select class="select" name="eventGubun">
			<option value=""></option>
			<option value="tbl_event" <% if (eventGubun = "tbl_event") then %>selected<% end if %> >�Ϲ� �̺�Ʈ</option>
			<option value="designfingers" <% if (eventGubun = "designfingers") then %>selected<% end if %> >�������ΰŽ� �̺�Ʈ</option>
			<option value="culturestation" <% if (eventGubun = "culturestation") then %>selected<% end if %> >��ó�����̼� �̺�Ʈ</option>
			<option value="tbl_event_etc" <% if (eventGubun = "tbl_event_etc") then %>selected<% end if %> >��Ÿ �̺�Ʈ</option>
		</select>
		&nbsp;
		���̵� :
		<input type="text" class="text" size="16" maxlength="32" name="userid" value="<%= userid %>">
		&nbsp;
		�̺�Ʈ�ڵ� :
		<input type="text" class="text" size="8" maxlength="32" name="eventCode" value="<%= eventCode %>">
		&nbsp;
		�̺�Ʈ�� :
		<input type="text" class="text" size="20" maxlength="32" name="eventName" value="<%= eventName %>">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		�������� :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
</tr>
</form>
</table>

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oCEventPrize.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oCEventPrize.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" width="100">����</td>
	<td align="center" width="60">�̺�Ʈ<br>�ڵ�</td>
	<td align="center" width="400">�̺�Ʈ��</td>
	<td align="center" width="80">������</td>
	<td align="center" width="80">������</td>
	<td align="center" width="80">��÷�ڹ�ǥ</td>
	<td align="center" width="100">���̵�</td>
	<td align="center">�ڸ�Ʈ</td>
	<td align="center" width="150">������</td>
	<td align="center" width="50">����</td>
	<td align="center">���</td>
</tr>
<% if oCEventPrize.FresultCount>0 then %>
<% for i=0 to oCEventPrize.FresultCount-1 %>
	<% if oCEventPrize.FItemList(i).finvaliduserid<>"" then %>
		<tr align="center" bgcolor="#e1e1e1" onmouseover=this.style.background="<%= adminColor("tabletop") %>"; onmouseout=this.style.background='#e1e1e1';>
	<% else %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="<%= adminColor("tabletop") %>"; onmouseout=this.style.background='#FFFFFF';>
	<% end if %>

	<td align="center" height="25">
		<%= oCEventPrize.FItemList(i).GetEventGubunName %>
	</td>
	<td align="center" height="25">
		<%= oCEventPrize.FItemList(i).Fevt_code %>
	</td>
	<td align="left">
		<%= oCEventPrize.FItemList(i).Fevt_name %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).Fevt_startdate %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).Fevt_enddate %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).Fevt_prizedate %>
	</td>
	<td align="left">
		<%= printUserId(oCEventPrize.FItemList(i).Fuserid, 2, "*") %>
	</td>
	<td align="left">
		<%= oCEventPrize.FItemList(i).Fcomment %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).Fregdate %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).GetIsUsingStr %>
	</td>
	<td align="center">
		<% if oCEventPrize.FItemList(i).finvaliduserid<>"" then %>
			Ư����������
		<% end if %>
	</td>
</tr>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if oCEventPrize.HasPreScroll then %>
			<span class="list_link"><a href="javascript:fnGotoPage(<%= oCEventPrize.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oCEventPrize.StartScrollPage to oCEventPrize.StartScrollPage + oCEventPrize.FScrollCount - 1 %>
			<% if (i > oCEventPrize.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oCEventPrize.FCurrPage) then %>
			<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
			<% else %>
			<a href="javascript:fnGotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
			<% end if %>
		<% next %>
		<% if oCEventPrize.HasNextScroll then %>
			<span class="list_link"><a href="javascript:fnGotoPage(<%= i %>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set oCEventPrize = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
