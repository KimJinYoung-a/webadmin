<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����α�
' Hieditor : 2013.11.14 �ѿ�� ����
'						2014.01.07 ������ ���� - �Ĺ� �ʵ��߰�, �Ⱓ�˻� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechulLogCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim research
Dim i, yyyy1,mm1,yyyy2,mm2, dd1, dd2, fromDate ,toDate ,oCMaechulLog, page, vatinclude, targetGbn, mwdiv_beasongdiv
dim searchfield, searchtext, makerid, dategbn, actDivCode, exceptSite
dim excTPL, isSum

	research = requestCheckvar(request("research"),10)
	actDivCode = requestCheckvar(request("actDivCode"),10)
	dategbn     = requestCheckvar(request("dategbn"),10)
	makerid   = requestcheckvar(request("makerid"),32)
	yyyy1   = requestcheckvar(request("yyyy1"),10)
	mm1     = requestcheckvar(request("mm1"),10)
	dd1     = requestcheckvar(request("dd1"),10)
	yyyy2   = requestcheckvar(request("yyyy2"),10)
	mm2     = requestcheckvar(request("mm2"),10)
	dd2     = requestcheckvar(request("dd2"),10)
	vatinclude     = requestcheckvar(request("vatinclude"),1)
	targetGbn     = requestcheckvar(request("targetGbn"),16)
	mwdiv_beasongdiv     = requestcheckvar(request("mwdiv_beasongdiv"),2)
	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
	page = requestcheckvar(request("page"),10)
    exceptSite = requestcheckvar(request("exceptSite"),10)

	excTPL 	= request("excTPL")
	isSum = requestcheckvar(request("isSum"),1)

if dategbn="" then dategbn="ActDate"
	if isSum = "" then isSum = "N"
if page = "" then page = 1
if (research = "") then
	excTPL = "Y"
end if

if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd1="") then dd1 = "01"
if (yyyy2="") then yyyy2 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm2="") then mm2 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd2="") then dd2 = "01"
'if (yyyy2="") then yyyy2 = Cstr(Year( dateadd("m",-1,date()) ))
'if (mm2="") then mm2 = Cstr(Month( dateadd("m",-1,date()) ))

'yyyy1=yyyy2
'mm1=mm2

fromDate = DateSerial(yyyy1, mm1,dd1)
toDate = DateSerial(yyyy2, mm2,dd2+1)

set oCMaechulLog = new CMaechulLog
	oCMaechulLog.FPageSize = 500       ''4000->500
	oCMaechulLog.FCurrPage = page
	oCMaechulLog.FRectDategbn = dategbn
	oCMaechulLog.FRectStartDate = fromDate
	oCMaechulLog.FRectEndDate = toDate
	oCMaechulLog.FRectvatinclude = vatinclude
	oCMaechulLog.FRecttargetGbn = targetGbn
	oCMaechulLog.FRectmwdiv_beasongdiv = mwdiv_beasongdiv
	oCMaechulLog.FRectSearchField = searchfield
	oCMaechulLog.FRectSearchText = searchtext
	oCMaechulLog.FRectmakerid = makerid
	oCMaechulLog.FRectActDivCode = actDivCode
	oCMaechulLog.FRectExceptSite = exceptSite

	oCMaechulLog.FRectExcTPL = excTPL
	oCMaechulLog.FRectIsSum = isSum
	oCMaechulLog.GetMaechul_month_brand_Log
%>

<script language="javascript">

function searchSubmit(page){
	frm.page.value=page;
	frm.submit();
}

function pop_detail_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, vatinclude, makerid){
	<% if dategbn="ActDate" then %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?actDate_yyyy1='+yyyy1+'&actDate_mm1='+mm1+'&actDate_dd1='+dd1+'&actDate_yyyy2='+yyyy2+'&actDate_mm2='+mm2+'&actDate_dd2='+dd2+'&chkActDate=Y&vatinclude='+vatinclude+'&makerid='+makerid+'&targetGbn=<%= targetGbn %>&actDivCode=<%= actDivCode %>&searchfield=<%=searchfield%>&searchtext=<%=searchtext%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% elseif (dategbn="chulgoDate") then %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?chulgoDate_yyyy1='+yyyy1+'&chulgoDate_mm1='+mm1+'&chulgoDate_dd1='+dd1+'&chulgoDate_yyyy2='+yyyy2+'&chulgoDate_mm2='+mm2+'&chulgoDate_dd2='+dd2+'&chkChulgoDate=Y&vatinclude='+vatinclude+'&makerid='+makerid+'&targetGbn=<%= targetGbn %>&actDivCode=<%= actDivCode %>&searchfield=<%=searchfield%>&searchtext=<%=searchtext%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% else %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?orgPay_yyyy1='+yyyy1+'&orgPay_mm1='+mm1+'&orgPay_dd1='+dd1+'&orgPay_yyyy2='+yyyy2+'&orgPay_mm2='+mm2+'&orgPay_dd2='+dd2+'&chkOrgPay=Y&vatinclude='+vatinclude+'&makerid='+makerid+'&targetGbn=<%= targetGbn %>&actDivCode=<%= actDivCode %>&searchfield=<%=searchfield%>&searchtext=<%=searchtext%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% end if %>

	pop_detail_list.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* ��¥ :
				<select name="dategbn">
					<option value="ipkumdate" <%=CHKIIF(dategbn="ipkumdate","selected","")%> >����������
					<option value="ActDate" <%=CHKIIF(dategbn="ActDate","selected","")%> >��������(ó������)
					<option value="chulgoDate" <%=CHKIIF(dategbn="chulgoDate","selected","")%> >�������
				</select>
				<%  DrawOneDateBoxdynamic "yyyy1",yyyy1,"mm1",mm1,"dd1",dd1,"", "", "", "" %> ~ <% DrawOneDateBoxdynamic "yyyy2",yyyy2,"mm2",mm2,"dd2",dd2,"", "", "", "" %>
				&nbsp;
				<input type="checkbox" name="excTPL" value="Y" <% if (excTPL = "Y") then %>checked<% end if %> >
				3PL ���� ����
				&nbsp;
				<input type="checkbox" name="isSum" value="S" <% if (isSum = "S") then %>checked<% end if %> >
				�귣�� �հ�
				<p>
				* ���ⱸ�� : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
				&nbsp;&nbsp;
				* �������� : <% drawSelectBoxVatYN "vatinclude", vatinclude %>
				&nbsp;&nbsp;
				* ���Ա��� : <% drawmwdiv_beasongdiv "mwdiv_beasongdiv", mwdiv_beasongdiv , "" %>
				&nbsp;&nbsp;
				* �ֹ����� :
				<select class="select" name="actDivCode">
					<option value=""></option>
					<option value="A" <% if (actDivCode = "A") then %>selected<% end if %> >���ֹ�</option>
					<option value="C" <% if (actDivCode = "C") then %>selected<% end if %> >����ֹ�</option>
					<option value="H" <% if (actDivCode = "H") then %>selected<% end if %> >��ǰ����</option>
					<option value="E" <% if (actDivCode = "E") then %>selected<% end if %> >��ȯ�ֹ�</option>
					<option value="M" <% if (actDivCode = "M") then %>selected<% end if %> >��ǰ�ֹ�</option>
					<option value="CC" <% if (actDivCode = "CC") then %>selected<% end if %> >�������ȭ�ֹ�</option>
					<option value="HH" <% if (actDivCode = "HH") then %>selected<% end if %> >��ǰ��������ֹ�</option>
					<option value="EE" <% if (actDivCode = "EE") then %>selected<% end if %> >��ȯ����ֹ�</option>
					<option value="MM" <% if (actDivCode = "MM") then %>selected<% end if %> >��ǰ����ֹ�</option>
				</select>
				<p>
				* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				&nbsp;&nbsp;
				* �˻����� :
				<select class="select" name="searchfield">
					<option value=""></option>
					<!-- option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >�ֹ���ȣ</option -->
					<option value="sitename" <% if (searchfield = "sitename") then %>selected<% end if %> >����ó</option>
				</select>
				<input type="text" class="text" name="searchtext" value="<%= searchtext %>">
				&nbsp;(<input type="checkbox" name="exceptSite" <%=CHKIIF(exceptSite="on","checked","")%> >�ش����ó����)
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit('');"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�� �ӵ��� ������ ��� ������ ���ð� ��ٷ� �ּ���. ���ϰ� ū ������ �Դϴ�.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!--<h5>�۾���</h5>-->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="40">
		�˻���� : <b><%= oCMaechulLog.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oCMaechulLog.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<%if isSum <> "S" then%>
	<% if (dategbn="ActDate") then %>
		<td rowspan="2">������<br>(ó����)</td>
	<% elseif (dategbn="chulgoDate") then %>
		<td rowspan="2">�����</td>
	<% else %>
		<td rowspan="2">��������</td>
	<% end if %>
	<%end if%>
	<td rowspan="2">�귣��ID</td>
	<td rowspan="2">��������</td>
	<td colspan="7">��޾�</td>
	<td colspan="4">ȸ�����</td>
	<td colspan="4">��ü�����</td>
	<td colspan="3">��ۺ������</td>
	<td rowspan="2">����<Br>���ϸ���</td>
	<td rowspan="2">���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td>��Ź</td>
	<td>��ü</td>
	<td>��ǰ��޾׼Ұ�</td>
	<td>�ٹ�</td>
	<td>����</td>
	<td>��ۺ���޾׼Ұ�</td>
	<td>����</td>
	<td>��Ź</td>
	<td>��ü</td>
	<td>�Ұ�</td>
	<td>����</td>
	<td>��Ź</td>
	<td>��ü</td>
	<td>�Ұ�</td>
	<td>�ٹ�</td>
	<td>����</td>
	<td>�Ұ�</td>
</tr>

<% if oCMaechulLog.FresultCount >0 then %>
<% for i=0 to oCMaechulLog.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<%if isSum <> "S" then%>
	<td>
		<a href="javascript:pop_detail_list('<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= dd1 %>','<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= dd2 %>','<%= oCMaechulLog.FItemList(i).fvatinclude %>','<%= oCMaechulLog.FItemList(i).fmakerid %>');" onfocus="this.blur()">
		<%= oCMaechulLog.FItemList(i).fyyyymm %></a>
	</td>
	<%end if%>
	<td><%= oCMaechulLog.FItemList(i).fmakerid %></td>
	<td><%= fnColor(oCMaechulLog.FItemList(i).fvatinclude,"tx") %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice_M, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice_W, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice_U, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice_M+oCMaechulLog.FItemList(i).ftotalMaechulPrice_W+oCMaechulLog.FItemList(i).ftotalMaechulPrice_U, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice_TT, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice_UU, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice_TT+oCMaechulLog.FItemList(i).ftotalMaechulPrice_UU, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).faccountMaechulPrice_M, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).faccountMaechulPrice_W, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).faccountMaechulPrice_U, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).faccountMaechulPrice_M+oCMaechulLog.FItemList(i).faccountMaechulPrice_W+oCMaechulLog.FItemList(i).faccountMaechulPrice_U, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash_M, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash_W, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash_U, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash_M+oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash_W+oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash_U, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).fbeasongUpcheJungsanCash_TT, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).fbeasongUpcheJungsanCash_UU, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).fbeasongUpcheJungsanCash_TT+oCMaechulLog.FItemList(i).fbeasongUpcheJungsanCash_UU, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMileage, 0) %></td>
	<td></td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="50" align="center">
       	<% if oCMaechulLog.HasPreScroll then %>
			<span class="list_link"><a href="javascript:searchSubmit('<%= oCMaechulLog.StartScrollPage-1 %>')">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oCMaechulLog.StartScrollPage to oCMaechulLog.StartScrollPage + oCMaechulLog.FScrollCount - 1 %>
			<% if (i > oCMaechulLog.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oCMaechulLog.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:searchSubmit('<%=i%>')" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oCMaechulLog.HasNextScroll then %>
			<span class="list_link"><a href="javascript:searchSubmit('<%=i%>')">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="40">�˻��� ������ �����ϴ�.</td>
</tr>
<% end if %>

<%
set oCMaechulLog = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
