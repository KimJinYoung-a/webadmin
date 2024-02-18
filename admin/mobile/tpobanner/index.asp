<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/tpobanner.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : ����� ���� tpobanner
' History : 2013.12.14 ����ȭ
'###############################################
	
	Dim isusing , dispcate
	dim page 
	Dim i
	dim otpobannerlist
	Dim sDt , modiTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")

	if page="" then page=1

	set otpobannerlist = new CMainbanner
	otpobannerlist.FPageSize		= 20
	otpobannerlist.FCurrPage		= page
	otpobannerlist.Fisusing			= isusing
	otpobannerlist.Fsdt				= sDt
	otpobannerlist.GetContentsList()

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
	location.href = "tpo_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
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
	if(confirm("�����- TPOBANNER�� �����Ͻðڽ��ϱ�?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_new_tpobanner_xml.asp?term=" + term;
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
    	<a href="tpo_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�� ��ϼ� : <b><%=otpobannerlist.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=otpobannerlist.FtotalPage%></b>
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
    <td width="10%">���Ĺ�ȣ</td>	
    <td width="10%">��뿩��</td>
</tr>
<% 
	for i=0 to otpobannerlist.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(otpobannerlist.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=otpobannerlist.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=otpobannerlist.FItemList(i).Fidx%></td>
	<td>
		<%
			If otpobannerlist.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(otpobannerlist.FItemList(i).Fxmlregdate,10),"-",".") & " <br/> " & Num2Str(hour(otpobannerlist.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(otpobannerlist.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
    <td><img src="<%=otpobannerlist.FItemList(i).Fbgimg%>" width="100" /><img src="<%=otpobannerlist.FItemList(i).Flimg%>" width="100" height="100" alt="<%=otpobannerlist.FItemList(i).Flalt%>"/><img src="<%=otpobannerlist.FItemList(i).Frimg%>" width="100" height="100" alt="<%=otpobannerlist.FItemList(i).Fralt%>"/></td>
	<td>
		<% 
			Response.Write "����: "
			Response.Write replace(left(otpobannerlist.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(otpobannerlist.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(otpobannerlist.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />����: "
			Response.Write replace(left(otpobannerlist.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(otpobannerlist.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(otpobannerlist.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(otpobannerlist.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(otpobannerlist.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = otpobannerlist.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(otpobannerlist.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
	<td><%=otpobannerlist.FItemList(i).Fsortnum%></td>
    <td><%=chkiif(otpobannerlist.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if otpobannerlist.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= otpobannerlist.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + otpobannerlist.StartScrollPage to otpobannerlist.StartScrollPage + otpobannerlist.FScrollCount - 1 %>
				<% if (i > otpobannerlist.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(otpobannerlist.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if otpobannerlist.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set otpobannerlist = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->