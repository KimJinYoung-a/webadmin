<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/productnoticeinfomationCls.asp" -->
<%
	Dim isusing , dispcate
	dim page 
	Dim i
	dim productNoticeInfomationList
	Dim sDt , modiTime, vplatform, mode

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	vplatform = "pc"
	mode = RequestCheckVar(request("mode"),5)

	if page="" then page=1
	If isusing = "" Then isusing ="Y"

	set productNoticeInfomationList = new CProductNoticeInfomation
	productNoticeInfomationList.FPageSize		= 100
	productNoticeInfomationList.FCurrPage		= page
	productNoticeInfomationList.GetInfomationList()

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
	<% if mode="copy" then %>
	location.href = "docopyjust1day.asp?idx="+v;
	<% else %>
	location.href = "just1day_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>&paramisusing=<%=isusing%>";
	<% end if %>
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
	if(confirm("�����- pick�� �����Ͻðڽ��ϱ�?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_new_pick_xml.asp?term=" + term;
			refreshFrm.submit();
	}
}

function jsquickadd(v){
	if(confirm("�Ϻ� ��������� ���� �Ͻðڽ��ϱ�?")) {
	location.href = "dopick.asp?menupos=<%=menupos%>&mode=quickadd&prevDate="+v;
	}
}
-->
</script>
<div style="float:right;clear:both;"><a href="productNoticeMainInfo_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&paramisusing=<%=isusing%>">
	<img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
</div>
<br><br>
<!--  ����Ʈ -->
<table width="30%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="10%">������� �ڵ�</td>
	<td width="20%">������� ���и�</td>
</tr>
<% 
	for i=0 to productNoticeInfomationList.FResultCount-1 
%>
<tr bgcolor="white" height="30" align="center">
    <td onclick="jsmodify('<%=productNoticeInfomationList.FItemList(i).FinfoDiv%>');" style="cursor:pointer;"><%=productNoticeInfomationList.FItemList(i).FinfoDiv%></td>
	<td onclick="jsmodify('<%=productNoticeInfomationList.FItemList(i).FinfoDiv%>');" style="cursor:pointer;"><%=productNoticeInfomationList.FItemList(i).FinfoDivName%></td>
</tr>
<% Next %>
</table>
<%
set productNoticeInfomationList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->