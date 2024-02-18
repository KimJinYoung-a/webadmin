<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
''response.write "������"
''dbget.close()	:	response.End


dim mode, yyyy, mm, makerid, cdl, cdm, cds
dim d, i, page
dim mstart, mend
dim OnlySellyn, OnlyIsUsing, danjongyn, mwdiv
dim research
dim monthDiff, grpby

mode = request("mode")

makerid = request("makerid")
yyyy = request("yyyy1")
mm = request("mm1")

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

page 		= request("page")
research 	= request("research")
OnlySellyn 	= request("OnlySellyn")
OnlyIsUsing = request("OnlyIsUsing")
danjongyn   = request("danjongyn")
mwdiv       = request("mwdiv")
monthDiff   = request("monthDiff")
grpby   	= request("grpby")

if (research="") and (OnlyIsUsing="") then OnlyIsUsing="Y"
if (research="") and (danjongyn="") then danjongyn="SN"
if (research="") and (mwdiv="") then mwdiv="MW"

if (mode = "") then mode = "out"

if (page = "") then
        page = 1
end if

if (yyyy = "") then
	d = CStr(dateadd("m" ,-1, now()))
	yyyy = Left(d,4)
	mm = Mid(d,6,2)
end if

if (monthDiff = "") then
	monthDiff = "1"
end if

if (grpby = "") then
	grpby = "itemid"
end if

mstart = "0000-00"
mend = "0000-00"

mstart = CStr(dateadd("m" ,-1, (yyyy + "-" + mm + "-01")))
mstart = Left(mstart,7)

mend = CStr(dateadd("m" ,-0, (yyyy + "-" + mm + "-01")))
mend = Left(mend,7)


dim olistforout

set olistforout = new CSummaryItemStock


olistforout.FRectEndDate = yyyy + "-" + mm
olistforout.FRectYYYYMM = Left(CStr(dateadd("m" ,-1, (olistforout.FRectEndDate + "-01"))),7)


olistforout.FRectMakerid = makerid
olistforout.FPageSize = 300
olistforout.FCurrPage = page
olistforout.FRectSearchMode = mode
olistforout.FRectCD1 = cdl
olistforout.FRectCD2 = cdm
olistforout.FRectCD3 = cds
olistforout.FRectOnlySellyn  = OnlySellyn
olistforout.FRectOnlyIsUsing = OnlyIsUsing
olistforout.FRectOnlyOldItem = "on"
olistforout.FRectOnlyOutItem = "on"
olistforout.FRectMwDiv = mwdiv
olistforout.FRectDanjongyn   = danjongyn
olistforout.FRectChulgoNo   = 1
olistforout.FRectMonthDiff   = monthDiff
olistforout.FRectGroupBy   = grpby

if (mode = "over") then
	olistforout.GetItemListOverStore
else
    '' ȸ������� �Լ��� ���� ��� (���� : GetItemListForOut)
    if (makerid<>"") then
	    olistforout.GetItemListTurnOver
	else
	    olistforout.GetBrandListTurnOver
	end if
end if
%>


<script language='javascript'>
function changecontent(){
	//dummy
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function popDetailByBrand(imakerid){
    var strUrl = '/admin/stock/outitem.asp?menupos=723';

    strUrl = strUrl + '&makerid=' + imakerid;
    strUrl = strUrl + '&research=on';
    strUrl = strUrl + '&yyyy1=' + frm.yyyy1.value;
    strUrl = strUrl + '&mm1=' + frm.mm1.value;
    strUrl = strUrl + '&mode=<%= mode %>';
	strUrl = strUrl + '&monthDiff=<%= monthDiff %>';
    strUrl = strUrl + '&OnlySellyn=' + frm.OnlySellyn.value;
    strUrl = strUrl + '&OnlyIsUsing=' + frm.OnlyIsUsing.value;
    strUrl = strUrl + '&mwdiv=' + frm.mwdiv.value;
    strUrl = strUrl + '&danjongyn=' + frm.danjongyn.value;
    strUrl = strUrl + '&cdl=' + frm.cdl.value;
    strUrl = strUrl + '&cdm=' + frm.cdm.value;
    strUrl = strUrl + '&cds=' + frm.cds.value;


    var popwin = window.open(strUrl,'popDetailByBrand','width=1200,height=800,scrollbars=yes,resizable=yes');

    popwin.focus();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<script>
function SubmitForm()
{
        document.frm.page.value = 1;
        document.frm.submit();
}
function GotoPage(pg)
{
        document.frm.page.value = pg;
        document.frm.submit();
}
</script>


	<form name="frm" method="get" action="" onsubmit="return false;">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�ش��� <% DrawYMBox yyyy, mm %> &nbsp;
        	<input type="radio" name="mode" value="out" <% if (mode = "out") then response.write "checked" end if %>>��������ǰ(�˻��Ⱓ : <%= mend %> ���� <select class="select" name="monthDiff"><option value="1" <% if (monthDiff = "1") then %>selected<% end if %> >1����</option><option value="6" <% if (monthDiff = "6") then %>selected<% end if %> >6����</option><option value="12" <% if (monthDiff = "12") then %>selected<% end if %> >12����</option><option value="18" <% if (monthDiff = "18") then %>selected<% end if %> >18����</option><option value="24" <% if (monthDiff = "24") then %>selected<% end if %> >24����</option></select>)
			&nbsp;&nbsp;
        	<input type="radio" name="mode" value="over" <% if (mode = "over") then response.write "checked" end if %>>��������ʰ���ǰ
        	<br>
        	�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
        	&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�Ǹ�:<% drawSelectBoxSellYN "OnlySellyn", OnlySellyn %>
	     	&nbsp;
	     	���:<% drawSelectBoxUsingYN "OnlyIsUsing", OnlyIsUsing %>
	     	&nbsp;
	     	����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	&nbsp;
	     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>

			<% if (mode <> "over") and (makerid<>"") then %>
			&nbsp;&nbsp;&nbsp;
			ǥ�ù��:
			<input type="radio" name="grpby" value="itemid" <% if (grpby = "itemid") then response.write "checked" end if %>> ��ǰ��
			&nbsp;
			<input type="radio" name="grpby" value="itemoption" <% if (grpby = "itemoption") then response.write "checked" end if %>> �ɼǺ�
			<% end if %>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<script language='javascript'>
document.onload = getOnload();

function getOnload(){
	startRequest('cdl','<%= cdl %>','<%= cdm %>','<%= cds %>');
}

function chkAll(v) {
	for (var i = 0;; i++) {
		var chk = document.getElementById("chk_" + i);
		if (chk == undefined) {
			break;
		}
		chk.checked = v.checked;
	}
}

function jsSetUseYN() {
	var frm = document.actfrm;

	frm.barcodeArr.value = "";
	for (var i = 0;; i++) {
		var chk = document.getElementById("chk_" + i);
		if (chk == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		frm.barcodeArr.value = frm.barcodeArr.value + "," + chk.value;
	}

	if (frm.barcodeArr.value == "") {
		alert("���� ��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm("���� ��ǰ�� ���������� ��ȯ�Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>

<p>

	<div align="right">
		<input type="button" class="button" value="���û�ǰ ������ ��ȯ" onClick="jsSetUseYN()" <% if (mode <> "out") or (makerid="") or (OnlyIsUsing <> "Y") or (OnlySellyn = "Y") or (OnlySellyn = "") or (grpby <> "itemid") then %>disabled<% end if %> >
		(��������ǰ+Ư���귣��+�����+�Ǹ��߾ƴ�+��ǰ�� �� ��� ��밡��)
	</div>

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<% if (mode = "over") then %>
		    ��������ʰ���ǰ :
		    <% else %>
		    ������� ��ǰ�� :
		    <% end if %>
		    <b><%= olistforout.FTotalCount %> </b>
		</td>
	</tr>

<% if (mode = "over") or (makerid<>"") then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="chkall" value="" onClick="chkAll(this);" <% if (mode <> "out") or (makerid="") or (OnlySellyn = "Y") or (OnlySellyn = "") or (grpby <> "itemid") then %>disabled<% end if %> ></td>
		<td width="100">�귣��ID</td>
		<td width="50">�̹���</td>
		<td width="40">��ǰ<br>�ڵ�</td>
		<td>��ǰ��(�ɼ�)</td>
		<td width="30">�ŷ�<br>����</td>

		<% if mode = "over" then %>
		<td width="60"><%= right(mend,2) %>��<br>�԰��ǰ</td>
		<td width="60"><%= right(mend,2) %>��<br>ON<br>�Ǹ�</td>
		<td width="60"><%= right(mend,2) %>��<br>OFF<br>���</td>
		<!-- <td width="30"><%= right(mend,2) %>��<br>��Ÿ<br>���</td>-->
		<% else %>
		<td width="60">�Ⱓ��<br>�԰��ǰ</td>
		<td width="60">�Ⱓ��<br>ON<br>�Ǹ�</td>
		<td width="60">�Ⱓ��<br>OFF<br>���</td>
		<!-- <td width="30">��Ÿ<br>���</td>-->
		<% end if %>
		<!--
		<td width="30">�ҷ�</td>
		<td width="30">����</td>
		-->
		<td width="60">�ش��<br>�������(�ǻ�)</td>
		<td width="60">�������(�ǻ�)</td>

		<td width="50">����<br>����</td>
		<td width="30">�Ǹ�<br>����</td>
		<td width="30">���<br>����</td>
		<td width="60">����<br>����</td>
	</tr>
<% for i=0 to olistforout.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" id="chk_<%= i %>" name="chk" value="<%= olistforout.FItemList(i).FItemID %>" <% if (mode <> "out") or (makerid="") or (OnlySellyn = "Y") or (OnlySellyn = "") or (grpby <> "itemid") then %>disabled<% end if %> ></td>
		<td align="left"><%= olistforout.FItemList(i).Fmakerid %></td>
		<td><img src="<%= olistforout.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<td><a href="javascript:PopItemSellEdit('<%= olistforout.FItemList(i).FItemID %>');"><%= olistforout.FItemList(i).FItemID %></a></td>
		<td align="left">
		    <a href="javascript:PopItemDetail('<%= olistforout.FItemList(i).Fitemid %>','<%= olistforout.FItemList(i).Fitemoption %>')"><%= olistforout.FItemList(i).Fitemname %></a>
		    <% if olistforout.FItemList(i).FitemoptionName <> "" then %>
		    <br>
		    <font color="blue">[<%= olistforout.FItemList(i).FitemoptionName %>]</font>
		    <% end if %>
		</td>
		<td><font color="<%= mwdivColor(olistforout.FItemList(i).Fmwdiv) %>"><%= mwdivName(olistforout.FItemList(i).Fmwdiv) %></font></td>

		<td><%= olistforout.FItemList(i).Freipgono %></td>
		<td><%= olistforout.FItemList(i).Fsellno %></td>
		<td><%= olistforout.FItemList(i).Foffchulgono %></td>
		<!--
		<td><%= olistforout.FItemList(i).Fetcchulgono %></td>
		<td><%= olistforout.FItemList(i).Ferrbaditemno %></td>
		<td><%= olistforout.FItemList(i).Ftoterrno %></td>
		-->
		<td><%= olistforout.FItemList(i).Frealstock %></td>
		<td><%= olistforout.FItemList(i).Fcurrrealstock %></td>
		<td>
	        <%= fnColor(olistforout.FItemList(i).Flimityn,"yn") %>
			<% if (olistforout.FItemList(i).Flimityn = "Y") then %>
	          	<br>(<%= olistforout.FItemList(i).GetLimitStr %>)
	        <% end if %>
<!--        	<font color="<%= ynColor(olistforout.FItemList(i).Flimityn) %>"><%= olistforout.FItemList(i).Flimityn %><% if (olistforout.FItemList(i).Flimityn = "Y") then response.write "(" + CStr(olistforout.FItemList(i).Flimitcount) + ")" end if %></font>-->
        </td>
		<td><%= fnColor(olistforout.FItemList(i).Fsellyn,"yn") %></td>
		<td><%= fnColor(olistforout.FItemList(i).Fisusing,"yn") %></td>
		<td><%= fnColor(olistforout.FItemList(i).Fdanjongyn,"dj") %></td>
	</tr>
<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
<% if olistforout.HasPreScroll then %>
		<a href="javascript:GotoPage(<%= olistforout.StartScrollPage-1 %>)">[pre]</a>
<% else %>
		[pre]
<% end if %>

<% for i=0 + olistforout.StartScrollPage to olistforout.FScrollCount + olistforout.StartScrollPage - 1 %>
        <% if i>olistforout.FTotalpage then Exit for %>
	<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
	<% else %>
		<a href="javascript:GotoPage(<%= i %>)">[<%= i %>]</a>
	<% end if %>
<% next %>

<% if olistforout.HasNextScroll then %>
		<a href="javascript:GotoPage(<%= i %>)">[next]</a>
<% else %>
		[next]
<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<% else %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="30">NO</td>
		<td>�귣��ID</td>
		<td width="80">����ǰ��</td>
		<td width="80">�Ⱓ��<br>�԰��ǰ</td>
		<td width="80">�Ⱓ��<br>ON�Ǹ����</td>
		<td width="80">�Ⱓ��<br>OFF���</td>
		<td width="80">�ش��<br>�������<br>(�ǻ�)</td>
		<td width="80">�������<br>(�ǻ�)</td>
		<td >&nbsp;</td>
    </tr>
    <% for i=0 to olistforout.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= (((page - 1) * olistforout.FPageSize) + i + 1) %></td>
        <td><%= olistforout.FItemList(i).Fmakerid %></td>
        <td><%= olistforout.FItemList(i).Fcnt %></td>
        <td><%= olistforout.FItemList(i).Freipgono %></td>
		<td><%= olistforout.FItemList(i).Fsellno %></td>
        <td><%= olistforout.FItemList(i).Foffchulgono %></td>
        <td><%= olistforout.FItemList(i).Frealstock %></td>
		<td><%= olistforout.FItemList(i).Fcurrrealstock %></td>
        <td align="left"><a href="javascript:popDetailByBrand('<%= olistforout.FItemList(i).Fmakerid %>');">��������&gt;&gt;</a></td>
    </tr>
    <% next %>
</table>
<% end if %>


<form name=actfrm method=post action="actoutbrand.asp">
	<input type=hidden name="mode" value="setuseyn">
	<input type=hidden name="barcodeArr" value="">
</form>

<%
set olistforout = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
