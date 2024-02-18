<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/appmanage/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todaywishCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : ����� ���� todaywish
' History : 2013.12.14 ����ȭ
'###############################################
	
	Dim isusing , dispcate
	dim page 
	Dim i
	dim todaywishList
	Dim sDt , modiTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")

	if page="" then page=1
	If isusing = "" Then isusing ="Y"

	set todaywishList = new Ctodaywish
	todaywishList.FPageSize		= 20
	todaywishList.FCurrPage		= page
	todaywishList.Fisusing		= isusing
	todaywishList.Fsdt			= sDt
	todaywishList.GetContentsList()

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
	location.href = "todaywish_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
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
	if(confirm("�����- todaywish�� �����Ͻðڽ��ϱ�?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_todaywish_xml.asp";
			refreshFrm.submit();
	}
}

function jsquickadd(v){
	if(confirm("�Ϻ� ��������� ���� �Ͻðڽ��ϱ�?")) {
	location.href = "dotodaywish.asp?menupos=<%=menupos%>&mode=quickadd&prevDate="+v;
	}
}
-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
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
			<!-- ����� -->
			<% If sDt <> "" Then %>
			��<input type="button" onclick="jsquickadd(document.all.prevDate.value)" value="�������"/>
			<% End If %>
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="�˻�">
		</td>
	</tr>
</form>	
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>XML Real ����(����)</a></td>
    <td align="right">
		<!-- �űԵ�� -->
    	<a href="todaywish_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�� ��ϼ� : <b><%=todaywishList.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=todaywishList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <td width="10%">������ real ����ð�</td>
	<td width="22%">����</td>	 
    <td width="13%">������/������</td>
    <td width="10%">�����</td>
    <td width="10%">�����</td>
    <td width="10%">����������</td>
    <td width="10%">��뿩��</td>
</tr>
<% 
	for i=0 to todaywishList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(todaywishList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=todaywishList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=todaywishList.FItemList(i).Fidx%></td>
	<td>
		<%
			If todaywishList.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(todaywishList.FItemList(i).Fxmlregdate,10),"-",".") & " / " & Num2Str(hour(todaywishList.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(todaywishList.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
    <td onclick="jsmodify('<%=todaywishList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=todaywishList.FItemList(i).Fwishtitle%></td>
	<td onclick="jsmodify('<%=todaywishList.FItemList(i).Fidx%>');" style="cursor:pointer;">
		<% 
			If todaywishList.FItemList(i).Fstartdate <> "" And todaywishList.FItemList(i).Fenddate Then 
				Response.Write "����: "
				Response.Write replace(left(todaywishList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(todaywishList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(todaywishList.FItemList(i).Fstartdate),2,"0","R")
				Response.Write "<br />����: "
				Response.Write replace(left(todaywishList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(todaywishList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(todaywishList.FItemList(i).Fenddate),2,"0","R")
			End If 
		%>
	</td>
	<td><%=left(todaywishList.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(todaywishList.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = todaywishList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(todaywishList.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(todaywishList.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if todaywishList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= todaywishList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + todaywishList.StartScrollPage to todaywishList.StartScrollPage + todaywishList.FScrollCount - 1 %>
				<% if (i > todaywishList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(todaywishList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if todaywishList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set todaywishList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->