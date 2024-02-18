<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����� ����Ʈ
' Hieditor : 2018.03.07 �̻� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_deliverycls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim page, i, j, k
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, basedate, fromdate, todate
dim delaydiv, makerid, orderserial
dim research

page     = requestCheckVar(request("page"),10)
yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1		= requestCheckVar(request("mm1"),2)
dd1		= requestCheckVar(request("dd1"),2)
yyyy2	= requestCheckVar(request("yyyy2"),4)
mm2		= requestCheckVar(request("mm2"),2)
dd2		= requestCheckVar(request("dd2"),2)

research		= requestCheckVar(request("research"),3)
makerid			= requestCheckVar(request("makerid"),32)
orderserial		= requestCheckVar(request("orderserial"),32)
delaydiv		= requestCheckVar(request("delaydiv"),32)

If page = "" Then page = 1

if (yyyy1="") then
	basedate = Left(CStr(DateAdd("d", -14, now())),10)
	yyyy1 = Left(basedate,4)
	mm1   = Mid(basedate,6,2)
	dd1   = Mid(basedate,9,2)

	basedate = Left(CStr(DateAdd("d", -0, now())),10)
	yyyy2 = Left(basedate,4)
	mm2   = Mid(basedate,6,2)
	dd2   = Mid(basedate,9,2)
end if

fromdate = Left(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
todate = Left(CStr(DateSerial(yyyy2,mm2 ,dd2+1)),10)

dim oCCSDelivery
set oCCSDelivery = New CCSDelivery
oCCSDelivery.FCurrPage				= page
oCCSDelivery.FPageSize				= 50
oCCSDelivery.FRectStartDate			= fromdate
oCCSDelivery.FRectEndDate			= todate
oCCSDelivery.FRectMakerid			= makerid
oCCSDelivery.FRectOrderserial		= orderserial
oCCSDelivery.FRectDelayDiv			= delaydiv

if (makerid <> "") then
	oCCSDelivery.GetCSMemoDeliveryDelayByMakerid()
else
	oCCSDelivery.GetCSMemoDeliveryDelaySUM()
end if

%>
<script>

function jsSubmit(frm) {
	frm.submit();
}

function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

</script>
<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" height="60" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���Ƽ���� :
		<select class="select" name="delaydiv">
			<option></option>
			<option value="delay" <%= CHKIIF(delaydiv="delay", "selected", "") %>>�������</option>
			<option value="stockout" <%= CHKIIF(delaydiv="stockout", "selected", "") %>>ǰ��</option>
		</select>
		&nbsp;
		���Ƽ�ΰ��� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		�귣�� : <input type="text" class="text" name="makerid" value="<%= makerid %>">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(frm);">
	</td>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">

	</td>
</tr>
</table>
</form>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= FormatNumber(oCCSDelivery.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCCSDelivery.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">����</td>
	<td width="200">�귣��</td>
	<td width="100">�ֹ���ȣ</td>
	<td width="100">����<br />��ǰ�ڵ�</td>
	<td width="80">�ΰ�����</td>
	<td width="70">�ΰ��Ǽ�</td>
    <td>���</td>
</tr>
<% if (oCCSDelivery.FResultCount > 0) then %>
	<% for i = 0 to (oCCSDelivery.FResultCount - 1) %>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td><%= oCCSDelivery.FItemList(i).GetDelayDivName %></td>
		<td><%= oCCSDelivery.FItemList(i).Fmakerid %></td>
		<td><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= oCCSDelivery.FItemList(i).FOrderSerial %>')" class="zzz"><%= oCCSDelivery.FItemList(i).Forderserial %></a></td>
		<td><%= oCCSDelivery.FItemList(i).Fitemid %></td>
		<td><%= oCCSDelivery.FItemList(i).FDPlusNDay %></td>
		<td><%= oCCSDelivery.FItemList(i).FcheckCnt %></td>
    	<td></td>
	</tr>
	<% next %>
	<tr height="20">
	    <td colspan="19" align="center" bgcolor="#FFFFFF">
	        <% if oCCSDelivery.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCCSDelivery.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCCSDelivery.StartScrollPage to oCCSDelivery.FScrollCount + oCCSDelivery.StartScrollPage - 1 %>
	    		<% if i>oCCSDelivery.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCCSDelivery.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="12">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
