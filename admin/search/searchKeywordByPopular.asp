<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr, plat, searchKey
dim i
dim research

research = request("research")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
plat = request("plat")
searchKey = requestCheckVar(request("searchKey"),32)

nowdateStr = CStr(now())

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Format00(2, Cstr(Month(now())))
if (dd1="") then dd1 = Format00(2, Cstr(day(now())))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Format00(2, Cstr(Month(now())))
if (dd2="") then dd2 = Format00(2, Cstr(day(now())))

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = yyyy2 + "-" + mm2 + "-" + dd2


if (research = "") then
	''if (groupby = "") then groupby = "d"
end if

'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword
osearchKeyword.FRectStart 		= startdateStr
osearchKeyword.FRectEnd 		= nextdateStr
osearchKeyword.FRectPlatform	= plat
if (searchKey <> "") then
	osearchKeyword.FRectKeyword		= "%" & searchKey & "%"
end if

osearchKeyword.getReportByPopularEVT

%>

<script language='javascript'>

function popOpenTrand(yyyy1, yyyy2, mm1, mm2, dd1, dd2, currKeyword) {
	if ((yyyy1 == yyyy2) && (mm1 == mm2) && (dd1 == dd2)) {
		var startDate = new Date(yyyy1, (mm1 - 1), (dd1 - 7));
		yyyy1 = startDate.getFullYear();
		mm1 = startDate.getMonth() + 1;
		dd1 = startDate.getDate();
	}

	var popwin = window.open("/admin/search/searchKeywordByTrand.asp?yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&dd1=" + dd1 + "&yyyy2=" + yyyy2 + "&mm2=" + mm2 + "&dd2=" + dd2 + "&searchKeyword=" + currKeyword,"popOpenTrand","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popOpenRelated(yyyy1, yyyy2, mm1, mm2, dd1, dd2, currKeyword) {
	if ((yyyy1 == yyyy2) && (mm1 == mm2) && (dd1 == dd2)) {
		var startDate = new Date(yyyy1, (mm1 - 1), (dd1 - 7));
		yyyy1 = startDate.getFullYear();
		mm1 = startDate.getMonth() + 1;
		dd1 = startDate.getDate();
	}

	//var popwin = window.open("/admin/search/searchKeywordByRelated.asp?yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&dd1=" + dd1 + "&yyyy2=" + yyyy2 + "&mm2=" + mm2 + "&dd2=" + dd2 + "&searchKeyword=" + currKeyword,"popOpenRelated","width=800 height=600 scrollbars=yes resizable=yes");
	var popwin = window.open("/admin/search/manageRelatedKeywordNEW.asp?research=on&page=1&menupos=3970&orgkeyword="+currKeyword+"&relatedKeyword=&useYN=Y","popOpenRelated","width=800 height=600 scrollbars=yes resizable=yes");

	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left" height="30" >
			�Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			ä�� :
			<select class="select" name="plat">
				<option></option>
				<option value="App" <%= CHKIIF(plat="App", "selected", "") %>>��</option>
				<option value="Mob" <%= CHKIIF(plat="Mob", "selected", "") %>>�����</option>
				<option value="Web" <%= CHKIIF(plat="Web", "selected", "") %>>��</option>
			</select>
			&nbsp;
			�˻��� : <input type="text" name="searchKey" value="<%= searchKey %>">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="30">
			(1�ð� ���� ������)
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p />

* <a href="http://k.tenbyten.kr:5601/goto/1c9781fd5dcbd5dfeed988899efc828a" target="_blank">http://k.tenbyten.kr:5601/goto/1c9781fd5dcbd5dfeed988899efc828a</a> �� �����ϼż� �α�˻�� ��ȸ�Ͻ� �� �ֽ��ϴ�.

<p />

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50" height="30">����</td>
		<td width="250">�˻���</td>
		<td width="100">�˻�Ƚ��</td>
		<td width="100">�˻������<br />��ǰ��</td>
		<td width="100">�����˻���</td>
		<td>���</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FTotalCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= (i + 1) %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).FcurrKeyword %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).Fcount %>
		</td>
		<td align="center">
			<% if (osearchKeyword.FItemList(i).Favgmxrectcnt <= 200) then %>
			<font color=red><b><%= osearchKeyword.FItemList(i).Favgmxrectcnt %></b></font>
			<% else %>
			<%= osearchKeyword.FItemList(i).Favgmxrectcnt %>
			<% end if %>
		</td>
		<td align="center">
			<a href="javascript:popOpenRelated('<%= yyyy1 %>', '<%= yyyy2 %>', '<%= mm1 %>', '<%= mm2 %>', '<%= dd1 %>', '<%= dd2 %>', '<%= osearchKeyword.FItemList(i).FcurrKeyword %>')">����</a>
		</td>
		<td align="center">
		</td>
	</tr>
	<%
	next
	%>
	<% if (osearchKeyword.FTotalCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="4">
			�˻������ �����ϴ�.
		</td>
	</tr>
	<% end if %>
</table>
<%
set osearchKeyword = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
