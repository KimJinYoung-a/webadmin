<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todayenjoyCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 메인 enjoybanner
' History : 2014.06.23 이종화
'		  : 2018.11.28 최종원 메인상단 기획전 추가
'###############################################
	
	Dim isusing , dispcate , validdate , research
	dim page 
	Dim i
	dim oEnjoyeventlist
	Dim sDt , modiTime , sedatechk
	Dim addtype , prevTime, setTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	validdate= request("validdate")
	research= request("research")
	sedatechk = request("sedatechk")
	addtype = request("addtype")
	prevTime = request("prevTime")	

	response.write sDt

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end If
	
	If addtype = "" Then addtype = 1	

	if page="" then page=1

	set oEnjoyeventlist = new CMainbanner
	oEnjoyeventlist.FPageSize			= 20
	oEnjoyeventlist.FCurrPage			= page
	oEnjoyeventlist.Fisusing			= isusing
	oEnjoyeventlist.Fsdt				= sDt
	oEnjoyeventlist.FRectvaliddate		= validdate
	'oEnjoyeventlist.FRectsedatechk		= sedatechk '//시작일 기준 체크
	oEnjoyeventlist.FRecttype			= addtype '//이벤트 타입
	'oEnjoyeventlist.FRectSelDateTime	= prevTime
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
//수정
function jsmodify(v, addtype){
	if(addtype == 3){
		location.href = "mainTopExhibition_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>&indexparam=1";
	}else{
		location.href = "enjoy_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
	}
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
	if(confirm("모바일- enjoyevent에 적용하시겠습니까?")) {
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
function controlExhibition(){			
	var popwin; 		
	popwin = window.open("/admin/pcmain/multievent/exhibition_ctrl.asp", "popup_item", "width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}
function addContent(){
	var contentType = document.frm.addtype.value;
	var dateOptionParam = document.frm.prevDate.value

	if(contentType == "3"){				
		document.location.href= "mainTopExhibition_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&dispOption="+contentType+"&dateoption="+dateOptionParam;	
	}else{
		document.location.href="enjoy_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"+"&dateoption="+dateOptionParam;
	}
	
}
function fnTrendEventCopy() {
    var popwin = window.open("/admin/pcmain/multievent/index.asp?mode=copy","popTemplateManage","width=1200,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
}
-->

function popTodayEasyReg(){
    let popTodayEasyReg = window.open('/admin/mobile/popTodayEasyReg.asp?type=enjoy','mainposcodeedit','width=800,height=400,scrollbars=yes,resizable=yes');
    popTodayEasyReg.focus();
}

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
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전&nbsp;
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			<!--시작일기준 <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />-->
			지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />			
			<!--시간 <input type="input" name="prevTime" value="<%=prevTime%>" class="text" size="10" maxlength="10" /> 시~			-->
			&nbsp; 타입 선택 : 
			<select name="addtype" class="select">
				<option value="">2016버전</option>
				<option value="1" <%=chkiif(addtype="1"," selected","")%>>기본형</option>
				<option value="2" <%=chkiif(addtype="2"," selected","")%>>기본형 + 상품 3개</option>
				<option value="3" <%=chkiif(addtype="3"," selected","")%>>메인상단기획전</option>
			</select>
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			</div>
		</td>
		<td width="120" bgcolor="<%= adminColor("gray") %>">
			<button sytle="float:left" type="button" onclick="controlExhibition();">메인상단기획전관리</button>
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
<!-- 	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>XML Real 적용(예약)</a></td> -->
    <td align="right">
        <input type="button" class="button" value="간편등록" onClick="popTodayEasyReg();" />
		<!-- 신규등록 -->
		<button onClick="fnTrendEventCopy();">불러오기</button>&nbsp;&nbsp;
    	<a href="javascript:void(0)" onclick="addContent();"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		총 등록수 : <b><%=oEnjoyeventlist.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oEnjoyeventlist.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="5%">타입</td>
	<td width="20%">등록이미지</td>
	<td width="10%">기획전코드</td>
    <td width="15%">시작일/종료일</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="5%">우선순위</td>
    <td width="10%">사용여부</td>
</tr>
<% 
	for i=0 to oEnjoyeventlist.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oEnjoyeventlist.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td style="cursor:pointer;">
		<a href="" onclick="jsmodify('<%=oEnjoyeventlist.FItemList(i).Fidx%>', '<%=oEnjoyeventlist.FItemList(i).Faddtype%>');return false;"><%=oEnjoyeventlist.FItemList(i).Fidx%></a>
		<p>&nbsp;</p>
		<% if oEnjoyeventlist.FItemList(i).Faddtype <> 3 then %>
		<a href="" onclick="window.open('enjoy_preview.asp?idx=<%=oEnjoyeventlist.FItemList(i).Fidx%>','enjoypreview', 'width=733, height=900');return false;">[미리보기]</a>
		<% end if %>		
	</td>
	<td>
		<%
			If oEnjoyeventlist.FItemList(i).Faddtype = 1 Then
				Response.write "기본형"
			ElseIf oEnjoyeventlist.FItemList(i).Faddtype = 2 Then
				Response.write "기본형<br/>+상품 3개"
			ElseIf oEnjoyeventlist.FItemList(i).Faddtype = 3 Then
				Response.write "메인상단<br/>기획전"				
			Else
				Response.write "2016버전"
			End If 
		%>
	</td>
    <td>	
	<% If oEnjoyeventlist.FItemList(i).Faddtype = 3 Then %>
		<img src="<%=oEnjoyeventlist.FItemList(i).Fevtimg%>" width="200" alt="<%=oEnjoyeventlist.FItemList(i).Fevtalt%>"/>
	<% else %>
		<% If oEnjoyeventlist.FItemList(i).Flinktype = "2" then %>
		<img src="<%=oEnjoyeventlist.FItemList(i).Fevtimg%>" width="200" alt="<%=oEnjoyeventlist.FItemList(i).Fevtalt%>"/>
		<% Else %>
		<img src="<%=oEnjoyeventlist.FItemList(i).Fevtmolistbanner%>" width="200" height="90" alt="<%=oEnjoyeventlist.FItemList(i).Fevtalt%>"/>
		<% End If %>	
	<% end if %>
	</td>
	<td><%=oEnjoyeventlist.FItemList(i).Fevt_code%></td>
	<td>
		<% 
			Response.Write "시작: "
			Response.Write replace(left(oEnjoyeventlist.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oEnjoyeventlist.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oEnjoyeventlist.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />종료: "
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
    <td><%=oEnjoyeventlist.FItemList(i).Fsortnum%></td>
    <td><%=chkiif(oEnjoyeventlist.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center">
		<% if oEnjoyeventlist.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oEnjoyeventlist.StarScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oEnjoyeventlist.StartScrollPage to oEnjoyeventlist.StartScrollPage + oEnjoyeventlist.FScrollCount - 1 %>
			<% if (i > oEnjoyeventlist.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oEnjoyeventlist.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oEnjoyeventlist.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
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