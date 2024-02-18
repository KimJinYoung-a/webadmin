<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : PC메인관리 저스트원데이
' History : 서동석 생성
'			2022.07.04 한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/just1DayCls2018New.asp" -->
<%
	Dim isusing , dispcate
	dim page 
	Dim i
	dim just1DayList
	Dim sDt , modiTime, vplatform, mode

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	vplatform = "pc"
	mode = RequestCheckVar(request("mode"),5)

	if page="" then page=1
	If isusing = "" Then isusing ="Y"

	set just1DayList = new Cjust1Day
	just1DayList.FPageSize		= 20
	just1DayList.FCurrPage		= page
	just1DayList.Fisusing			= isusing
	just1DayList.Fsdt					= sDt
	just1DayList.FRectPlatform			= vplatform
	just1DayList.GetContentsList()

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
	<% if mode="copy" then %>
	location.href = "docopyjust1day.asp?idx="+v;
	<% else %>
	location.href = "just1day_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>&paramisusing=<%=isusing%>";
	<% end if %>
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

function RefreshCaFavKeyWordRec(term){
	if(confirm("모바일- pick에 적용하시겠습니까?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_new_pick_xml.asp?term=" + term;
			refreshFrm.submit();
	}
}

function jsquickadd(v){
	if(confirm("일별 빠른등록을 실행 하시겠습니까?")) {
	location.href = "dopick.asp?menupos=<%=menupos%>&mode=quickadd&prevDate="+v;
	}
}
-->
</script>
<!-- 검색 시작 -->
<form name="frm" method="post" style="margin:0px;" action="/admin/sitemaster/just1day2018/index.asp">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
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
			<input type="submit" class="button_s" value="검 색">
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<div style="float:right;clear:both;"><a href="just1day_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&paramisusing=<%=isusing%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a></div>
<br><br>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		총 등록수 : <b><%=just1DayList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=just1DayList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="10%">구분</td>
	<td width="22%">제목</td>	 
    <td width="18%">시작일/종료일</td>
	<td width="10%">최대할인율</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="10%">사용여부</td>
</tr>
<% 
	for i=0 to just1DayList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(just1DayList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=just1DayList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=just1DayList.FItemList(i).Fidx%></td>
	<td onclick="jsmodify('<%=just1DayList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=Replace(Replace(just1DayList.FItemList(i).FType,"event","기획전"),"just1day","JUST1DAY")%></td>
    <td onclick="jsmodify('<%=just1DayList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%= ReplaceBracket(just1DayList.FItemList(i).Ftitle) %></td>
	<td onclick="jsmodify('<%=just1DayList.FItemList(i).Fidx%>');" style="cursor:pointer;" align="left">
		<% 
			If just1DayList.FItemList(i).Fstartdate <> "" And just1DayList.FItemList(i).Fenddate Then 
				Response.Write "시작: "
				Response.Write replace(left(just1DayList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(just1DayList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(just1DayList.FItemList(i).Fstartdate),2,"0","R")
				Response.Write "<br />종료: "
				Response.Write replace(left(just1DayList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(just1DayList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(just1DayList.FItemList(i).Fenddate),2,"0","R")
				If cInt(datediff("d", now() , just1DayList.FItemList(i).Fenddate)) < 0 Or cInt(datediff("h", now() , just1DayList.FItemList(i).Fenddate )) < 0  Then 
					Response.write " <span style=""color:red"">(종료)</span>"
				ElseIf cInt(datediff("d", just1DayList.FItemList(i).Fenddate , now())) = 0  Then '종료일이 오늘날짜
					If cInt(datediff("h", now() , just1DayList.FItemList(i).Fenddate )) >= 0 And cInt(datediff("h", now() , just1DayList.FItemList(i).Fenddate )) < 24 Then ' 오늘
						Response.write " <span style=""color:red"">(약 "& cInt(datediff("h", now() , just1DayList.FItemList(i).Fenddate )) &" 시간후 종료)</span>"
					Else  ' 시작전
						Response.write " <span style=""color:red"">(시작전)</span>"					
					End If 
				'// 시작일이 오늘날짜이고 종료일이 오늘이 아니면
				ElseIf cInt(datediff("d", just1DayList.FItemList(i).Fstartdate , now()))>=0 And cInt(datediff("d", just1DayList.FItemList(i).Fenddate , now())) < 0 Then
					Response.write " <span style=""color:red"">(약 "&cInt(datediff("d", now() , just1DayList.FItemList(i).Fenddate ))&"일 " &cInt(datediff("h", now() , just1DayList.FItemList(i).Fenddate ))-cInt(datediff("d", now() , just1DayList.FItemList(i).Fenddate ))*24 &"시간후 종료)</span>"
				ElseIf cInt(datediff("d", just1DayList.FItemList(i).Fenddate , now())) < 0 Then
					If cInt(datediff("h", now() , just1DayList.FItemList(i).Fenddate )) >= 0 And cInt(datediff("h", now() , just1DayList.FItemList(i).Fenddate )) < 24 Then ' 오늘
						Response.write " <span style=""color:red"">(약 "& cInt(datediff("h", now() , just1DayList.FItemList(i).Fenddate )) &" 시간후 종료)</span>"
					Else  ' 시작전
						Response.write " <span style=""color:red"">(시작전)</span>"
					End If 
				End If
			End If 
		%>
	</td>
    <td onclick="jsmodify('<%=just1DayList.FItemList(i).Fidx%>');" style="cursor:pointer;">
		<%=just1DayList.FItemList(i).Fsaleper%>
	</td>
	<td><%=left(just1DayList.FItemList(i).Fregdate,10)%></td>
	<td><%=just1DayList.FItemList(i).Fusername%></td>
	<td>
		<%
			modiTime = just1DayList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write just1DayList.FItemList(i).Fusername2 & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(just1DayList.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if just1DayList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= just1DayList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + just1DayList.StartScrollPage to just1DayList.StartScrollPage + just1DayList.FScrollCount - 1 %>
				<% if (i > just1DayList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(just1DayList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if just1DayList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set just1DayList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->