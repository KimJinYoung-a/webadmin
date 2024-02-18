<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/main_play.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 메인 play
' History : 2015-05-13 이종화
'###############################################
	
	Dim isusing , dispcate
	dim page 
	Dim i
	dim playBannerlist
	Dim sDt , modiTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	
	If isusing = "" Then isusing ="Y"

	sDt = request("prevDate")

	if page="" then page=1

	set playBannerlist = new CMainbanner
		playBannerlist.FPageSize		= 20
		playBannerlist.FCurrPage		= page
		playBannerlist.Fisusing			= isusing
		playBannerlist.Fsdt				= sDt
		playBannerlist.GetContentsList()

	'// 그라운드 이름
	Function gubuntitle(v)
		Select Case v
			Case "1"
				Response.write "GROUND"
			Case "2"
				Response.write "STYLE+"
			Case "3"
				Response.write "DESIGN FINGERS"
			Case "4"
				Response.write "T-EPISODE"
			Case "5"
				Response.write "GIFT"
			Case Else
				Response.write ""
		End Select
	End Function

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
	location.href = "play_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
}
$(function() {
  	$("input[type=submit]").button();

  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");

});

function RefreshCaFavKeyWordRec(){
	if(confirm("모바일 , 앱 hot keyword 에 적용하시겠습니까?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_play_xml.asp";
			refreshFrm.submit();
	}
}

-->
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
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
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:submit();">
		</td>
	</tr>
</form>	
</table>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>XML Real 적용(예약)</a></td>
    <td align="right">
		<!-- 신규등록 -->
    	<a href="play_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		총 등록수 : <b><%=playBannerlist.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=playBannerlist.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <td width="10%">마지막 <br/>real 적용시간</td>
	<td width="20%">등록내용</td>	 
    <td width="15%">시작일/종료일</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="10%">정렬번호</td>	
    <td width="10%">사용여부</td>
</tr>
<% 
	for i=0 to playBannerlist.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(playBannerlist.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=playBannerlist.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=playBannerlist.FItemList(i).Fidx%></td>
	<td>
		<%
			If playBannerlist.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(playBannerlist.FItemList(i).Fxmlregdate,10),"-",".") & " <br/> " & Num2Str(hour(playBannerlist.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(playBannerlist.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
    <td align="left"><img src="<%=playBannerlist.FItemList(i).Fpimg%>" alt="" width="100"/><br/><br/>구분텍스트: <%=gubuntitle(playBannerlist.FItemList(i).Fgubun)%><br/>제목 : <%=playBannerlist.FItemList(i).Ftitle%></td>
	<td>
		<% 
			Response.Write "시작: "
			Response.Write replace(left(playBannerlist.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(playBannerlist.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(playBannerlist.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />종료: "
			Response.Write replace(left(playBannerlist.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(playBannerlist.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(playBannerlist.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(playBannerlist.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(playBannerlist.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = playBannerlist.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(playBannerlist.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
	<td><%=playBannerlist.FItemList(i).Fsortnum%></td>
    <td><%=chkiif(playBannerlist.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if playBannerlist.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= playBannerlist.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + playBannerlist.StartScrollPage to playBannerlist.StartScrollPage + playBannerlist.FScrollCount - 1 %>
				<% if (i > playBannerlist.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(playBannerlist.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if playBannerlist.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set playBannerlist = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->