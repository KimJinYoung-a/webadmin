<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/que/queItemCls.asp"-->
<%
Dim oOutItemSummary, page, i, sellsite, snapDate
Dim oQueSummary
page		= requestCheckvar(request("page"),10)
sellsite	= requestCheckvar(request("sellsite"),32)
snapDate	= requestCheckvar(request("snapDate"),10)

Dim research : research = requestCheckvar(request("research"),10)
Dim apiact : apiact = requestCheckvar(request("apiact"),32)
Dim isysyusr : isysyusr = requestCheckvar(request("isysyusr"),10)
'Dim showimage : showimage = requestCheckvar(request("showimage"),10)
Dim showsummary : showsummary = "on"
'Dim bygrp : bygrp = requestCheckvar(request("bygrp"),10)
Dim isiteMatch

if (research="") and (snapDate="") then snapDate=LEFT(dateadd("d",-1,NOW()),10)

If page = "" Then page = 1
if (showsummary="on") then
	set oOutItemSummary = new COutmallSummary
	oOutItemSummary.FRectSnapDate	= snapDate
	oOutItemSummary.getOutItemSummaryList
end if
SET oQueSummary = new COutmallSummary
	oQueSummary.FCurrPage		= page
	oQueSummary.FPageSize		= 50
	oQueSummary.FRectSellsite	= sellsite
	oQueSummary.FRectSnapDate	= snapDate
    ''oQueSummary.FRectIsSysUser  = isysyusr
    oQueSummary.FRectApiAction  = apiact
	oQueSummary.getOutSailyQueSummary
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function rePage(sellsite,param1,param2){
	var frm = document.frm;
	frm.sellsite.value=sellsite;
    frm.apiact.value=param1;

	frm.submit();
}


</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">

<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���޸� :
		<select class="select" name="sellsite">
			<option value="">-��ü-</option>
			<option value="ssg" <%= chkiif(sellsite = "ssg", "selected", "") %> >SSG</option>
			<option value="lfmall" <%= chkiif(sellsite = "lfmall", "selected", "") %> >LFmall</option>
			<option value="hmall1010" <%= chkiif(sellsite = "hmall1010", "selected", "") %> >hMall</option>
			<option value="boribori1010" <%= chkiif(sellsite = "boribori1010", "selected", "") %> >��������</option>
			<option value="auction1010" <%= chkiif(sellsite = "auction1010", "selected", "") %> >����</option>
			<option value="ezwel" <%= chkiif(sellsite = "ezwel", "selected", "") %> >���������</option>
			<option value="shintvshopping" <%= chkiif(sellsite = "shintvshopping", "selected", "") %> >�ż���TV����</option>
			<option value="skstoa" <%= chkiif(sellsite = "skstoa", "selected", "") %> >SKSTOA</option>
			<option value="gmarket1010" <%= chkiif(sellsite = "gmarket1010", "selected", "") %> >G����</option>
			<option value="gseshop" <%= chkiif(sellsite = "gseshop", "selected", "") %> >GSShop</option>
			<option value="benepia1010" <%= chkiif(sellsite = "benepia1010", "selected", "") %> >�����Ǿ�</option>
			<option value="wconcept1010" <%= chkiif(sellsite = "wconcept1010", "selected", "") %> >W����</option>
			<option value="interpark" <%= chkiif(sellsite = "interpark", "selected", "") %> >������ũ</option>
			<option value="nvstorefarm" <%= chkiif(sellsite = "nvstorefarm", "selected", "") %> >�������</option>
			<option value="nvstoregift" <%= chkiif(sellsite = "nvstoregift", "selected", "") %> >������� �����ϱ�</option>
			<option value="Mylittlewhoopee" <%= chkiif(sellsite = "Mylittlewhoopee", "selected", "") %> >������� Ĺ�ص�</option>
			<option value="WMP" <%= chkiif(sellsite = "WMP", "selected", "") %> >������</option>
			<option value="11st1010" <%= chkiif(sellsite = "11st1010", "selected", "") %> >11����</option>
			<option value="lotteon" <%= chkiif(sellsite = "lotteon", "selected", "") %> >�Ե�On</option>
			<option value="lotteimall" <%= chkiif(sellsite = "lotteimall", "selected", "") %> >�Ե����̸�</option>
			<option value="cjmall" <%= chkiif(sellsite = "cjmall", "selected", "") %> >CJMall</option>
			<option value="kakaostore" <%= chkiif(sellsite = "kakaostore", "selected", "") %> >īī���彺���</option>
			<option value="kakaogift" <%= chkiif(sellsite = "kakaogift", "selected", "") %> >īī������Ʈ</option>
		</select>&nbsp;&nbsp;
		������¥ :
		<input id="snapDate" name="snapDate" value="<%=snapDate%>" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />

		&nbsp;
        ����Action
        <input type="text" name="apiact" value="<%=apiact%>" size="10" maxlength="32">

        <% if (FALSE) then %>
        &nbsp;
        ���� �۾���
        <select name="isysyusr">
            <option value="">��ü
            <option value="1" <%=CHKIIF(isysyusr="1","selected","") %> >System
            <option value="0" <%=CHKIIF(isysyusr="0","selected","") %> >Human
        </select>
        <% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br />
<% if (showsummary="on") then %>
<%
	Dim iTTL, imallActive, imallWait, imallAVailSellY, iregWiat, iregFail, imallInActive
	for i=0 to oOutItemSummary.FResultCount - 1
		iTTL = iTTL + oOutItemSummary.FItemList(i).FTTL
		imallActive = imallActive + oOutItemSummary.FItemList(i).FmallActive
		imallWait = imallWait + oOutItemSummary.FItemList(i).FmallWait
		imallAVailSellY	 = imallAVailSellY + oOutItemSummary.FItemList(i).FmallAVailSellY
		iregWiat	 = iregWiat + oOutItemSummary.FItemList(i).FregWiat
        iregFail     = iregFail + oOutItemSummary.FItemList(i).FregFail
        imallInActive = imallInActive + oOutItemSummary.FItemList(i).FmallInActive
	next
%>
<table width="100%" align="center" cellpadding="3" cellspacing="5" class="a" bgcolor="#FFFFFF">
<tr>
	<td width="50%">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" style="cursor:pointer">
			<td >����Ʈ</td>
			<td width="10%">�ѻ�ǰ��</td>
			<td width="10%"><strong>Active</strong></td>
			<td width="10%">���δ��</td>
			<td width="10%">�Ǹž���</td>
			<td width="10%">��ϴ��</td>
			<td width="10%">��Ͻ���</td>
			<td width="10%">InActive<%=CLNG(oOutItemSummary.FResultCount/2)%></td>
		</tr>
		<% for i=0 to CLNG(oOutItemSummary.FResultCount/2) - 1 %>
		<% isiteMatch = oOutItemSummary.FItemList(i).Fsellsite=sellsite %>
		<tr align="right" bgcolor="#FFFFFF" style="cursor:pointer">
			<td align="center" <%=CHKIIF(isiteMatch,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=oOutItemSummary.FItemList(i).Fsellsite %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FTTL,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FmallActive,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FmallWait,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FmallAVailSellY,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FregWiat,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FregFail,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FmallInActive,0) %></td>
		</tr>
		<% next %>
		</table>
	</td>
	<td width="50%">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td >����Ʈ</td>
			<td width="10%">�ѻ�ǰ��</td>
			<td width="10%"><strong>Active</strong></td>
			<td width="10%">���δ��</td>
			<td width="10%">�Ǹž���</td>
			<td width="10%">��ϴ��</td>
			<td width="10%">��Ͻ���</td>
			<td width="10%">InActive</td>
		</tr>
		<% for i=CLNG(oOutItemSummary.FResultCount/2) to CLNG(oOutItemSummary.FResultCount/2)*2- 1 %>
        <% if oOutItemSummary.FResultCount>i then %>
		<% isiteMatch = oOutItemSummary.FItemList(i).Fsellsite=sellsite %>
		<tr align="right" bgcolor="#FFFFFF" style="cursor:pointer">
			<td align="center" <%=CHKIIF(isiteMatch,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=oOutItemSummary.FItemList(i).Fsellsite %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FTTL,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FmallActive,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FmallWait,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FmallAVailSellY,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FregWiat,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FregFail,0) %></td>
			<td <%=CHKIIF(isiteMatch ,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOutItemSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOutItemSummary.FItemList(i).FmallInActive,0) %></td>
		</tr>
        <% else %>
        <tr align="right" bgcolor="#FFFFFF" style="cursor:pointer">
            <td >&nbsp;</td>
            <td colspan="7"></td>
        </tr>
        <% end if %>
		<% next %>
		</table>
	</td>
</tr>
</table>
<br />
<% set oOutItemSummary = Nothing %>
<% end if %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= FormatNumber(oQueSummary.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oQueSummary.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if oQueSummary.FRectSellsite<>"" then %>
    <td width="100">��¥</td>
    <% else %>
	<td width="100">���޸�</td>
    <% end if %>
	<td width="80">apiAction</td>
	<td width="70">�ѰǼ�</td>
    <td width="70">�Ǽ�_S</td>
	<td width="70">����_S</td>
	<td width="70">����_S</td>
	<td width="70">�ߺ�SKIP_S</td>
	<td width="70">UnKnown_S</td>
	<td width="70">��������_S</td>
	<td width="70">NULL_S</td>
	<td ></td>
    <td width="70">�Ǽ�_H</td>
	<td width="70">����_H</td>
	<td width="70">����_H</td>
	<td width="70">�ߺ�SKIP_H</td>
	<td width="70">UnKnown_H</td>
	<td width="70">��������_H</td>
	<td width="70">NULL_H</td>
</tr>
<%
	Dim DiffStat
%>
<% For i=0 to oQueSummary.FResultCount - 1 %>
<%
	DiffStat = ""
%>
<tr align="center" bgcolor="#FFFFFF">
    <% if (oQueSummary.FRectSellsite<>"") then %>
    <td style="cursor:pointer;" onClick="rePage('<%=oQueSummary.FRectSellsite%>','','')"><%= oQueSummary.FItemList(i).Fyyyymmdd %></td>
    <% else %>
	<td style="cursor:pointer;" onClick="rePage('<%=oQueSummary.FItemList(i).FSellsite%>','','')"><%= oQueSummary.FItemList(i).FSellsite %></td>
    <% end if %>
	<td style="cursor:pointer;" onClick="rePage('<%= CHKIIF(oQueSummary.FRectSellsite<>"",oQueSummary.FRectSellsite,oQueSummary.FItemList(i).FSellsite)%>','<%= oQueSummary.FItemList(i).FapiAction %>','')"><%= oQueSummary.FItemList(i).FapiAction %></td>
	<td><%= FormatNumber(oQueSummary.FItemList(i).FTTL,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FTTL-oQueSummary.FItemList(i).FTTL_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_OK-oQueSummary.FItemList(i).FS_OK_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_ERR-oQueSummary.FItemList(i).FS_ERR_H,0) %></td>
	<td><%= FormatNumber(oQueSummary.FItemList(i).FS_DUPP-oQueSummary.FItemList(i).FS_DUPP_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_BLANK-oQueSummary.FItemList(i).FS_BLANK_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_NOREAD-oQueSummary.FItemList(i).FS_NOREAD_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_NULL-oQueSummary.FItemList(i).FS_NULL_H,0) %></td>
    <td></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FTTL_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_OK_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_ERR_H,0) %></td>
	<td><%= FormatNumber(oQueSummary.FItemList(i).FS_DUPP_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_BLANK_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_NOREAD_H,0) %></td>
    <td><%= FormatNumber(oQueSummary.FItemList(i).FS_NULL_H,0) %></td>
</tr>
<% Next %>
<tr height="21">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oQueSummary.HasPreScroll then %>
		<a href="javascript:goPage('<%= oQueSummary.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oQueSummary.StartScrollPage to oQueSummary.FScrollCount + oQueSummary.StartScrollPage - 1 %>
    		<% if i>oQueSummary.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oQueSummary.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<script language="javascript">
$(function() {
	var CAL_Start = new Calendar({
		inputField : "snapDate", trigger    : "sDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			//CAL_End.args.min = date;
			//CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->