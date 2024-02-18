<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/appmanage/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/enjoyeventCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : ����� ���� enjoybanner
' History : 2013.12.14 ����ȭ
'###############################################
	
	Dim isusing , dispcate
	dim page 
	Dim i
	dim oEnjoyeventlist
	Dim sDt , modiTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")

	if page="" then page=1

	set oEnjoyeventlist = new CMainbanner
	oEnjoyeventlist.FPageSize			= 20
	oEnjoyeventlist.FCurrPage			= page
	oEnjoyeventlist.Fisusing			= isusing
	oEnjoyeventlist.Fsdt				= sDt
	oEnjoyeventlist.GetContentsList()

%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script type='text/javascript'>
<!--
//����
function jsmodify(v){
	location.href = "enjoy_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
}
$(function() {
  	$("input[type=submit]").button();

  	// ������ư
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");

});

function RefreshCaFavKeyWordRec(term){
	if(confirm("�����- enjoyevent�� �����Ͻðڽ��ϱ�?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_enjoyevent_xml.asp?term=" + term;
			refreshFrm.submit();
	}
}

-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			* ��뿩�� :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			�������� <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:submit();">
		</td>
	</tr>
</form>	
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>������ �����Ͽ� <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�<a href="javascript:RefreshCaFavKeyWordRec(document.all.vTerm.value);"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>XML Real ����(����)</a></td>
    <td align="right">
		<!-- �űԵ�� -->
    	<a href="enjoy_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�� ��ϼ� : <b><%=oEnjoyeventlist.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oEnjoyeventlist.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <td width="10%">������ <br/>real ����ð�</td>
	<td width="20%">����̹���</td>	 
    <td width="15%">������/������</td>
    <td width="10%">�����</td>
    <td width="10%">�����</td>
    <td width="10%">����������</td>
    <td width="10%">��뿩��</td>
</tr>
<% 
	for i=0 to oEnjoyeventlist.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oEnjoyeventlist.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=oEnjoyeventlist.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oEnjoyeventlist.FItemList(i).Fidx%></td>
	<td>
		<%
			If oEnjoyeventlist.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(oEnjoyeventlist.FItemList(i).Fxmlregdate,10),"-",".") & " <br/> " & Num2Str(hour(oEnjoyeventlist.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(oEnjoyeventlist.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
    <td><img src="<%=oEnjoyeventlist.FItemList(i).Fimg1%>" height="100" alt="<%=oEnjoyeventlist.FItemList(i).Fimg1alt%>"/><br/><img src="<%=oEnjoyeventlist.FItemList(i).Fimg2%>" height="100" alt="<%=oEnjoyeventlist.FItemList(i).Fimg2alt%>"/><br/><img src="<%=oEnjoyeventlist.FItemList(i).Fimg3%>" height="100" alt="<%=oEnjoyeventlist.FItemList(i).Fimg3alt%>"/></td>
	<td>
		<% 
			Response.Write "����: "
			Response.Write replace(left(oEnjoyeventlist.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oEnjoyeventlist.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oEnjoyeventlist.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />����: "
			Response.Write replace(left(oEnjoyeventlist.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oEnjoyeventlist.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oEnjoyeventlist.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(oEnjoyeventlist.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(oEnjoyeventlist.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = oEnjoyeventlist.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(oEnjoyeventlist.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(oEnjoyeventlist.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if oEnjoyeventlist.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oEnjoyeventlist.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oEnjoyeventlist.StartScrollPage to oEnjoyeventlist.StartScrollPage + oEnjoyeventlist.FScrollCount - 1 %>
				<% if (i > oEnjoyeventlist.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oEnjoyeventlist.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oEnjoyeventlist.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set oEnjoyeventlist = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->