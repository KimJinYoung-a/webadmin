<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/justonedayCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 메인 justoneday
' History : 2013.12.14 이종화
'###############################################
	
	Dim isusing , dispcate
	dim page 
	Dim i
	dim oJustonedayList
	Dim sDt , modiTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")

	if page="" then page=1
	If isusing = "" Then isusing = "Y"

	set oJustonedayList = new Cjustoneday
	oJustonedayList.FPageSize		= 20
	oJustonedayList.FCurrPage		= page
	oJustonedayList.Fisusing			= isusing
	oJustonedayList.Fsdt				= sDt
	oJustonedayList.GetContentsList()

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
//수정
function jsmodify(v){
	location.href = "justoneday_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
}
$(function() {
  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
	
});

function RefreshCaFavKeyWordRec(term){
	if(confirm("모바일- JUSTONEDAY에 적용하시겠습니까?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_justoneday_xml.asp?term=" + term;
			refreshFrm.submit();
	}
}

function jsquickadd(v){
	if(confirm("일별 빠른등록을 실행 하시겠습니까?")) {
	location.href = "dojustoneday.asp?menupos=<%=menupos%>&mode=quickadd&prevDate="+v;
	}
}
-->
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			<!-- 퀵등록 -->
			<% If sDt <> "" Then %>
			일<input type="button" onclick="jsquickadd(document.all.prevDate.value)" value="빠른등록"/>
			<% End If %>
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="검색">
		</td>
	</tr>
</form>	
</table>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>오늘을 포함하여 <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">일간<a href="javascript:RefreshCaFavKeyWordRec(document.all.vTerm.value);"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>XML Real 적용(예약)</a></td>
    <td align="right">
		<!-- 신규등록 -->
    	<a href="justoneday_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		총 등록수 : <b><%=oJustonedayList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oJustonedayList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="10%">마지막 real 적용시간</td>
	<td width="20%">제목</td>	 
    <td width="15%">시작일/종료일</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="10%">사용여부</td>
</tr>
<% 
	for i=0 to oJustonedayList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oJustonedayList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=oJustonedayList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oJustonedayList.FItemList(i).Fidx%></td>
	<td>
		<%
			If oJustonedayList.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(oJustonedayList.FItemList(i).Fxmlregdate,10),"-",".") & " / " & Num2Str(hour(oJustonedayList.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(oJustonedayList.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
    <td onclick="jsmodify('<%=oJustonedayList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oJustonedayList.FItemList(i).Fonairtitle%></td>
	<td onclick="jsmodify('<%=oJustonedayList.FItemList(i).Fidx%>');" style="cursor:pointer;">
		<% 
			Response.Write "시작: "
			Response.Write replace(left(oJustonedayList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oJustonedayList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oJustonedayList.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />종료: "
			Response.Write replace(left(oJustonedayList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oJustonedayList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oJustonedayList.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(oJustonedayList.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(oJustonedayList.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = oJustonedayList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(oJustonedayList.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(oJustonedayList.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if oJustonedayList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oJustonedayList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oJustonedayList.StartScrollPage to oJustonedayList.StartScrollPage + oJustonedayList.FScrollCount - 1 %>
				<% if (i > oJustonedayList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oJustonedayList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oJustonedayList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set oJustonedayList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->