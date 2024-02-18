<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �귣�庰���
' History : ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

dim research
dim makerid, onoffgubun, mwdiv, usingyn, centermwdiv, itemgubun, startMon, endMon, purchaseType
dim stocktype, limitrealstock
dim page, pagesize

dim i

research     	= RequestCheckVar(request("research"),32)
makerid     	= RequestCheckVar(request("makerid"),32)
onoffgubun     	= RequestCheckVar(request("onoffgubun"),32)
mwdiv     		= RequestCheckVar(request("mwdiv"),32)
usingyn     	= RequestCheckVar(request("usingyn"),32)
centermwdiv     = RequestCheckVar(request("centermwdiv"),32)
stocktype     	= RequestCheckVar(request("stocktype"),32)
limitrealstock 	= RequestCheckVar(request("limitrealstock"),32)
page     		= RequestCheckVar(request("page"),32)
pagesize     	= RequestCheckVar(request("pagesize"),32)
itemgubun     	= RequestCheckVar(request("itemgubun"),32)
startMon     	= RequestCheckVar(request("startMon"),32)
endMon     		= RequestCheckVar(request("endMon"),32)
purchaseType	= RequestCheckVar(request("purchaseType"),32)

if (research = "") then
	stocktype = "real"
end if

if (page = "") then
	page = "1"
end if
if (pagesize = "") then
	pagesize = "100"
end if

if itemgubun = "" then
	''itemgubun = "10"
end if


dim osummarystockbrand
set osummarystockbrand = new CSummaryItemStock

	osummarystockbrand.FPageSize = pagesize
	osummarystockbrand.FCurrPage = page

	osummarystockbrand.FRectMakerid = makerid
	osummarystockbrand.FRectOnlyIsUsing = usingyn
	osummarystockbrand.FRectMWDiv = mwdiv
	osummarystockbrand.FRectCenterMWDiv = centermwdiv
	osummarystockbrand.FRectStockType = stocktype
	osummarystockbrand.FRectlimitrealstock = limitrealstock
	osummarystockbrand.FRectItemGubun = itemgubun
	osummarystockbrand.FRectPurchaseType = purchaseType

	if IsNumeric(startMon) then
		osummarystockbrand.FRectStartDate = startMon
	elseif (startMon <> "") then
		response.write "<script>alert('������ ���ڸ� �����մϴ�. " & startMon & "')</script>"
	end if
	if IsNumeric(endMon) then
		osummarystockbrand.FRectEndDate = endMon
	elseif (endMon <> "") then
		response.write "<script>alert('������ ���ڸ� �����մϴ�. " & endMon & "')</script>"
	end if

	if itemgubun = "10" or itemgubun = "" then
		osummarystockbrand.GetRealStockByOnlineBrand
	else
		osummarystockbrand.GetRealStockByOfflineBrand
	end if

%>

<script language="JavaScript" src="/js/ttpbarcode.js"></script>
<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function SubmitSearch() {
    document.frm.submit();
}

function jsPopDetail(makerid, mwdiv, centermwdiv, itemgubun) {
	var onoffgubun = itemgubun;
	if (onoffgubun == "10") {
		onoffgubun = "on";
	} else if (itemgubun == "exc10") {
		onoffgubun = "off";
	} else {
		onoffgubun = "off" + itemgubun;
	}
	var url = "/admin/stock/brandcurrentstock.asp?menupos=708&onoffgubun=" + onoffgubun + "&makerid=" + makerid + "&usingyn=&mwdiv=" + mwdiv + "&centermwdiv=" + centermwdiv + "&stocktype=<%= stocktype %>&limitrealstock=<%= limitrealstock %>&startMon=<%= startMon %>&endMon=<%= endMon %>";

	var popwin = window.open(url, "jsPopDetail", "width=1500,height=800,scrollbars=yes,resizable=yes");
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �귣��:	<% drawSelectBoxDesignerwithName "makerid", makerid %>
		&nbsp;&nbsp;
		* ���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;&nbsp;
		* �ŷ����� :<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;&nbsp;
		* ���͸��Ա��� :
		<select class="select" name="centermwdiv">
			<option value="">��ü</option>
			<option value="M" <% if centermwdiv="M" then response.write "selected" %> >����</option>
			<option value="W" <% if centermwdiv="W" then response.write "selected" %> >Ư��</option>
			<option value="N" <% if centermwdiv="N" then response.write "selected" %> >������</option>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ��ǰ����: <% drawSelectBoxItemGubunForSearch "itemgubun", itemgubun %>
		&nbsp;&nbsp;
	    * ����� :
		<select name="stocktype" class="select">
			<option value="sys" <% if (stocktype = "sys") then %>selected<% end if %> >�ý������</option>
			<option value="real" <% if (stocktype = "real") then %>selected<% end if %> >��ȿ���</option>
		</select>
		: <% drawSelectBoxexistsstock "limitrealstock", limitrealstock, "" %>
		&nbsp;&nbsp;
		* ǥ�ð��� :
		<select class="select" name="pagesize">
			<option value="100" <% if (pagesize = "100") then %>selected<% end if %> >100 ��</option>
			<option value="500" <% if (pagesize = "500") then %>selected<% end if %> >500 ��</option>
			<option value="1000" <% if (pagesize = "1000") then %>selected<% end if %> >1000 ��</option>
		</select>
		&nbsp;&nbsp;
		* ������ :
		<input type="text" class="text" name="startMon" size="2" value="<%= startMon %>">
		~
		<input type="text" class="text" name="endMon" size="2" value="<%= endMon %>"> ����
		&nbsp;&nbsp;
	    * �������� :
		<% drawPartnerCommCodeBox true,"purchasetype","purchaseType",purchaseType,"" %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="29">
		�˻���� : <b><%= osummarystockbrand.FTotalCount %></b>
		&nbsp;
		������ :
		<% if osummarystockbrand.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>
		<b><%= page %> / <%= osummarystockbrand.FTotalPage %></b>
		<% if (osummarystockbrand.FTotalpage - osummarystockbrand.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"></td>
    <td width="200">�귣��ID</td>
	<td width="60">��ǰ����</td>
	<td width="60">���Ա���</td>
	<td width="60">����<br>���Ա���</td>
	<td width="70">��ǰǰ���</td>
	<td width="70">���>0<br>��ǰ��</td>
	<td width="70">���Ǹ�</td>

	<td width="70">�ý������</td>
	<td width="70">�ǻ����</td>
	<td width="70">�ǻ����</td>
	<td width="70">�ҷ�</td>
	<td width="70">��ȿ���</td>

	<td >���</td>
</tr>
<% if (osummarystockbrand.FResultCount = 0) then %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td colspan="14">������ �����ϴ�.</td>
</tr>
<% else %>
<% for i=0 to osummarystockbrand.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="25">
    <td></td>
    <td><%= osummarystockbrand.FItemList(i).Fmakerid %></td>
    <td><%= osummarystockbrand.FItemList(i).Fitemgubun %></td>
	<td><%= osummarystockbrand.FItemList(i).Fmwdiv %></td>
	<td><%= osummarystockbrand.FItemList(i).FCentermwdiv %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).FitemCnt,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).FitemPlusCnt,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Ftotsellno,0) %></td>

	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Ftotsysstock,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Ferrrealcheckno,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).getErrAssignStock,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Ferrbaditemno,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Frealstock,0) %></td>

	<td>
		<input type="button" class="button" value="��" onClick="jsPopDetail('<%= osummarystockbrand.FItemList(i).Fmakerid %>', '<%= osummarystockbrand.FItemList(i).Fmwdiv %>', '<%= osummarystockbrand.FItemList(i).Fcentermwdiv %>', '<%= osummarystockbrand.FItemList(i).Fitemgubun %>')" >
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="29" align="center">
		<% if osummarystockbrand.HasPreScroll then %>
		<a href="javascript:NextPage('<%= osummarystockbrand.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + osummarystockbrand.StartScrollPage to osummarystockbrand.FScrollCount + osummarystockbrand.StartScrollPage - 1 %>
			<% if i>osummarystockbrand.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if osummarystockbrand.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<%
set osummarystockbrand = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
