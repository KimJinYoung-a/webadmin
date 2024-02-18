<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/today_brandinfoCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : ����� ���� enjoybanner
' History : 2014.06.23 ����ȭ
'###############################################
	
	Dim isusing , dispcate , validdate , research
	dim page 
	Dim i
	dim oBrandinfo
	Dim sDt , modiTime , sedatechk
	Dim addtype

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	validdate= request("validdate")
	research= request("research")
	sedatechk = request("sedatechk")
	addtype = request("addtype")

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end if
	
	if page="" then page=1

	set oBrandinfo = new CMainbanner
	oBrandinfo.FPageSize			= 20
	oBrandinfo.FCurrPage			= page
	oBrandinfo.Fisusing			= isusing
	oBrandinfo.Fsdt				= sDt
	oBrandinfo.FRectvaliddate		= validdate
	oBrandinfo.FRectsedatechk		= sedatechk '//������ ���� üũ
	oBrandinfo.GetContentsList()
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
	location.href = "brandinfo_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
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
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_todayenjoy_xml.asp";
			refreshFrm.submit();
	}
}

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}
function addContents(){
	var dateOptionParam
	var frm = document.frm
	dateOptionParam = frm.prevDate.value

	document.location.href="brandinfo_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&dateoption="+dateOptionParam
}
-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������&nbsp;
			* ��뿩�� :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			�����ϱ��� <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
			�������� <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			&nbsp;
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
<!-- 	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>XML Real ����(����)</a></td> -->
    <td align="right">
		<!-- �űԵ�� -->
    	<a href="javascript:addContents();"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�� ��ϼ� : <b><%=oBrandinfo.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oBrandinfo.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="20%">�귣���̹���</td>	 
	<td width="10%">�������̹���<br/>��ǰ�ڵ�</td>
    <td width="15%">������/������</td>
    <td width="10%">�����</td>
    <td width="10%">�����</td>
    <td width="10%">����������</td>
    <td width="10%">��뿩��</td>
</tr>
<% 
	for i=0 to oBrandinfo.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oBrandinfo.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=oBrandinfo.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oBrandinfo.FItemList(i).Fidx%></td>
    <td>
		<img src="<%=oBrandinfo.FItemList(i).Fmainimg%>" width="200" alt=""/>
	</td>
	<td>
		<img src="<%=oBrandinfo.FItemList(i).Fmoreimg%>" width="100" alt=""/>
		<br/><br/>
		��ǰ�ڵ�1 : <%=oBrandinfo.FItemList(i).Fitemid1%><br/>��ǰ�ڵ�2 : <%=oBrandinfo.FItemList(i).Fitemid2%>
	</td>
	<td>
		<% 
			Response.Write "����: "
			Response.Write replace(left(oBrandinfo.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oBrandinfo.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oBrandinfo.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />����: "
			Response.Write replace(left(oBrandinfo.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oBrandinfo.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oBrandinfo.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(oBrandinfo.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(oBrandinfo.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = oBrandinfo.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(oBrandinfo.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(oBrandinfo.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td colspan="11" align="center">
		<% if oBrandinfo.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oBrandinfo.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oBrandinfo.StartScrollPage to oBrandinfo.StartScrollPage + oBrandinfo.FScrollCount - 1 %>
			<% if (i > oBrandinfo.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oBrandinfo.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oBrandinfo.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set oBrandinfo = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->