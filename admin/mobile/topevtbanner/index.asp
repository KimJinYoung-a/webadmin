<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/submenu/inc_subhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/topeventCls.asp" -->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 카테고리 top 2 event
' History : 2015-09-16 이종화
'###############################################
	
	Dim isusing , gcode , validdate , research
	dim page 
	Dim i
	dim oTopevtList
	Dim sDt , modiTime , sedatechk

	page = request("page")
	gcode = request("gcode")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	validdate= request("validdate")
	research= request("research")
	sedatechk = request("sedatechk")

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end if
	
	if page="" then page=1

	set oTopevtList = new CMainbanner
	oTopevtList.FPageSize			= 20
	oTopevtList.FCurrPage			= page
	oTopevtList.Fisusing			= isusing
	oTopevtList.Fsdt				= sDt
	oTopevtList.FRectvaliddate		= validdate
	oTopevtList.FRectgnbcode		= gcode

	oTopevtList.FRectsedatechk		= sedatechk '//시작일 기준 체크

	oTopevtList.GetContentsList()

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
	location.href = "enjoy_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
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

function RefreshCaFavKeyWordRec(v){
	if(confirm("모바일, 앱 TOP 2 EVENT에 적용하시겠습니까?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.gcode.value = v;
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_topeventbanner_xml.asp";
			refreshFrm.submit();
	}
}

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}
-->
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			* 주의 사항 : <span style="font-size:13px;"><strong>GNB 메뉴 검색후 XML 등록 버튼이 생성 됩니다. (개별 생성)</strong></span></br>
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전&nbsp;
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			시작일기준 <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
			지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			* GNB 메뉴 : 
			<% Call drawSelectBoxGNB("gcode" , gcode) %>
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
	<% If gcode <> "" Then %>
	<td><a href="javascript:RefreshCaFavKeyWordRec('<%=gcode%>');"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>XML Real 적용(예약)</a></td>
	<% End If %>
    <td align="right">
		<!-- 신규등록 -->
    	<a href="enjoy_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&gcode=<%=gcode%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		총 등록수 : <b><%=oTopevtList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oTopevtList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <td width="10%">마지막 <br/>real 적용시간</td>
    <td width="10%">등록 GNB/제목</td>
	<td width="15%">등록이미지</td>	 
    <td width="15%">시작일/종료일</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="10%">우선순위</td>
    <td width="10%">사용여부</td>
</tr>
<% 
	for i=0 to oTopevtList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oTopevtList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=oTopevtList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oTopevtList.FItemList(i).Fidx%></td>
	<td>
		<%
			If oTopevtList.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(oTopevtList.FItemList(i).Fxmlregdate,10),"-",".") & " <br/> " & Num2Str(hour(oTopevtList.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(oTopevtList.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
	<td>GNB : <%=oTopevtList.FItemList(i).Fgnbname%><br/><br/><%=oTopevtList.FItemList(i).Fevttitle%></br><%=oTopevtList.FItemList(i).Fevttitle2%></td>
    <td align="left">
		<% If oTopevtList.FItemList(i).Flinktype = "2" then %>
		<img src="<%=oTopevtList.FItemList(i).Fevtimg%>" width="200" alt="<%=oTopevtList.FItemList(i).Fevtalt%>"/>
		<% Else %>
		<img src="<%=oTopevtList.FItemList(i).Fevtmolistbanner%>" width="200" height="90" alt="<%=oTopevtList.FItemList(i).Fevtalt%>"/>
		<% End If %>
	</td>
	<td>
		<% 
			Response.Write "시작: "
			Response.Write replace(left(oTopevtList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oTopevtList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oTopevtList.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />종료: "
			Response.Write replace(left(oTopevtList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oTopevtList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oTopevtList.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(oTopevtList.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(oTopevtList.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = oTopevtList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(oTopevtList.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=oTopevtList.FItemList(i).Fsortnum%></td>
    <td><%=chkiif(oTopevtList.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center">
		<% if oTopevtList.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oTopevtList.StarScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oTopevtList.StartScrollPage to oTopevtList.StartScrollPage + oTopevtList.FScrollCount - 1 %>
			<% if (i > oTopevtList.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oTopevtList.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oTopevtList.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set oTopevtList = Nothing
%>
<form name="refreshFrm" method="post">
<input type="hidden" name="gcode" />
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->