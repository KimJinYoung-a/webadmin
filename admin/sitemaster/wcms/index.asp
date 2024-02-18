<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/mainWCMSCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 사이트 메인 관리
' History : 2013.03.28 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim siteDiv, pageDiv, isusing, tplIdx, selDt, sDt, eDt 
Dim oTemplate, oMainCont, lp, modiTime
Dim page

'// 파라메터 접수
siteDiv = request("site")
pageDiv = request("pDiv")
isusing = request("isusing")
tplIdx = request("tplIdx")
sDt = request("sDt")
eDt = request("eDt")
page = request("page")
if siteDiv="" then siteDiv="P"		'기본값 PC웹(P:PC웹, M:모바일)
if pageDiv="" then pageDiv="10"		'기본값 사이트메인(10:사이트메인, 20:이벤트메인...)
if isusing="" then isusing="Y"
if sDt="" then sDt=cStr(date)
if eDt="" then eDt=cStr(date)
if sDt=eDt then selDt=sDt
if page="" then page="1"


'// 템플릿 목록
	set oTemplate = new CCMSContent
	oTemplate.FPageSize = 50
	oTemplate.FRectSiteDiv = siteDiv
	oTemplate.FRectPageDiv = pageDiv
    oTemplate.GetTemplateList

'// 페이지정보 목록
	set oMainCont = new CCMSContent
	oMainCont.FPageSize = 20
	oMainCont.FRectTplIdx = tplIdx
	oMainCont.FRectSiteDiv = siteDiv
	oMainCont.FRectPageDiv = pageDiv
	oMainCont.FRectIsUsing = isusing
	oMainCont.FRectStartDate = sDt
	oMainCont.FRectEndDate = eDt
    oMainCont.GetMainPageList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function() {
  	$("input[type=submit]").button();

  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// 캘린더
	var arrDayMin = ["일","월","화","수","목","금","토"];
	var arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
      	showOn: "button",
    	onClose: function() {
    		if($("#sDt").datepicker("getDate")>$("#eDt").datepicker("getDate")) {
    			$("#eDt").datepicker("setDate",$("#sDt").datepicker("getDate"));
    		}
    	}
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
      	showOn: "button",
    	onClose: function() {
    		if($("#eDt").datepicker("getDate")<$("#sDt").datepicker("getDate")) {
    			$("#sDt").datepicker("setDate",$("#eDt").datepicker("getDate"));
    		}
    	}
    });

	// 행 정렬
	$( "#mainList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="30" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});

function popTemplateManage(){
    var popwin = window.open('/admin/sitemaster/wcms/popTemplateEdit.asp?site=<%=siteDiv%>&pDiv=<%=pageDiv%>','popTemplateManage','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function goMainContent(idx) {
	location.href="/admin/sitemaster/wcms/mainPageManage.asp?mainIdx="+idx+"&site=<%=siteDiv%>&pDiv=<%=pageDiv%>&menupos=<%=request("menupos")%>&sDt=<%=sDt%>&eDt=<%=eDt%>";
}

function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 템플릿을 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.target="_self";
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}


//프론트 미리보기
function previewPage() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("미리보실 템플릿을 선택해주세요.");
		return;
	}

	if($("select[name='site'], input[name='site']").val()=="M") {
		var url;
		switch($("select[name='pDiv']").val()) {
			case "10" :
				url = "<%=mobileUrl%>/chtml/preview/previewMainIndex.asp?sDt=<%=sDt%>";
				break;
			case "20" :
				url = "<%=mobileUrl%>/chtml/preview/previewEventBanner.asp?sDt=<%=sDt%>";
				break;
		}

		 if(confirm("[<%=sDt%>]부터 나흘간의 미리보기를 하시겠습니까?")) {
			 var popwin = window.open('','refreshFrm_Main','width=350,height=600,scrollbars=yes,resizable=yes');
			 popwin.focus();
			 frmList.target = "refreshFrm_Main";
			 frmList.action = url;
			 frmList.submit();
		}
	} else {
		alert("PC웹은 서비스 준비중입니다.\n기존 페이지관리를 사용해주세요.");
		return;
	}
}


//프론트 페이지에 선택 적용
function assignPage() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("적용하실 템플릿을 선택해주세요.");
		return;
	}

	if($("select[name='site'], input[name='site']").val()=="M") {
		var msg, url;
		switch($("select[name='pDiv']").val()) {
			case "10" :
				msg = "선택하신 항목을 \"모바일\" 사이트 메인 페이지를 프론트에 즉시 적용하시겠습니까?\n\n※한 번 적용된 내용은 되돌릴 수 없습니다.";
				url = "<%=mobileUrl%>/chtml/make_main_xml.asp";
				break;
			case "20" :
				msg = "선택하신 항목을 \"모바일\" 이벤트 메인 페이지를 프론트에 즉시 적용하시겠습니까?\n\n※오늘부터 5일분이 저장되며, 한 번 적용된 내용은 되돌릴 수 없습니다.";
				url = "<%=mobileUrl%>/chtml/make_event_xml.asp";
				break;
		}

		if(confirm(msg)) {
			 var popwin = window.open('','refreshFrm_Main','');
			 popwin.focus();

			if($("input[name='cTrm']").val()!="0") {
				frmList.sTrm.value = $("input[name='cTrm']").val();
			}
			 frmList.chkAll.value="N";
			 frmList.target = "refreshFrm_Main";
			 frmList.action = url;
			 frmList.submit();
		}
	} else {
		if(confirm("현재 \"PC웹\" 프론트페이지를 즉시 적용하시겠습니까?\n\n※한 번 적용된 내용은 되돌릴 수 없습니다.")) {
			alert("PC웹은 서비스 준비중입니다.\n기존 페이지관리를 사용해주세요.");
		}
	}
}

//프론트 페이지에 전체 적용
function assignPageALL() {
	if($("select[name='site'], input[name='site']").val()=="M") {
		var msg, url;
		switch($("select[name='pDiv']").val()) {
			case "10" :
				msg = "앞으로 4일간의 전체 항목을 \"모바일\" 사이트 메인 페이지를 프론트에 즉시 적용하시겠습니까?\n\n※한 번 적용된 내용은 되돌릴 수 없습니다.";
				url = "<%=mobileUrl%>/chtml/make_main_xml.asp";
				break;
			case "20" :
				msg = "앞으로 5일간의 전체 항목을 \"모바일\" 이벤트 메인 페이지를 프론트에 즉시 적용하시겠습니까?\n\n※한 번 적용된 내용은 되돌릴 수 없습니다.";
				url = "<%=mobileUrl%>/chtml/make_event_xml.asp";
				break;
		}

		if(confirm(msg)) {
			 var popwin = window.open('','refreshFrm_Main','');
			 popwin.focus();

			if($("input[name='cTrm']").val()!="0") {
				frmList.sTrm.value = $("input[name='cTrm']").val();
			}
			 frmList.chkAll.value="Y";
			 frmList.target = "refreshFrm_Main";
			 frmList.action = url;
			 frmList.submit();
		}
	} else {
		if(confirm("현재 \"PC웹\" 프론트페이지를 즉시 적용하시겠습니까?\n\n※한 번 적용된 내용은 되돌릴 수 없습니다.")) {
			alert("PC웹은 서비스 준비중입니다.\n기존 페이지관리를 사용해주세요.");
		}
	}
}

// 빠른 등록
function fnQuitReg(oTpl) {
	var tplId = oTpl.value;
	var tplNm = oTpl[oTpl.selectedIndex].text;
	var tplDt = document.frm.sDt.value;

	var chk = confirm("일자 ["+tplDt+"]에 \""+tplNm+"\"를 등록하시겠습니까?\n\n※ 템플릿 기본값으로 설정되며 이후 내용을 반드시 변경하셔야됩니다.");
	if(chk) {
		var frm = document.frmQuitReg;
		frm.tplIdx.value=tplId;
		frm.StartDate.value=tplDt;
		frm.EndDate.value=tplDt;
		frm.mainTitle.value="*** 빠른등록 > 수정해주세요";
		frm.submit();
	} else {
		return;
	}
}
</script>

<!-- 상단 검색폼 시작 -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
	    <% if C_ADMIN_AUTH then %>
	    사이트:
		<select name="site" class="select">
			<option value="P" <%=chkIIF(siteDiv="P","selected","")%> >PC웹</option>
			<option value="M" <%=chkIIF(siteDiv="M","selected","")%> >모바일</option>
		</select>
		&nbsp;/&nbsp;
		<% else %>
		<input type="hidden" name="site" value="<%=siteDiv%>" />
		<% end if %>
	    사용처:
		<select name="pDiv" class="select">
			<option value="10" <%=chkIIF(pageDiv="10","selected","")%> >사이트 메인</option>
			<option value="20" <%=chkIIF(pageDiv="20","selected","")%> >이벤트 메인</option>
		</select>
		&nbsp;/&nbsp;
	    사용구분:
		<select name="isusing" class="select">
			<option value="A">전체</option>
			<option value="Y" <%=chkIIF(isusing="Y","selected","")%> >사용함</option>
			<option value="N" <%=chkIIF(isusing="N","selected","")%> >사용안함</option>
		</select>
		&nbsp;/&nbsp;
		템플릿:
		<select name="tplIdx" class="select">
			<option value="">전체</option>
			<%
				if oTemplate.FResultCount>0 then
					for lp=0 to (oTemplate.FResultCount-1)
						Response.Write "<option value='" & oTemplate.FItemList(lp).FtplIdx & "' " & chkIIF(cStr(oTemplate.FItemList(lp).FtplIdx)=tplIdx,"selected","") & ">" & oTemplate.FItemList(lp).FtplName & "</option>"
					next
				end if
			%>
		</select>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="검색" />
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
	<td>
		조회기간:
		<span id="rdoDtPreset">
		<input type="radio" name="selDatePreset" id="rdoDtOpt1" value="<%=dateadd("d",-1,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",-1,date)),"checked","")%> /><label for="rdoDtOpt1">-1</label><input type="radio" name="selDatePreset" id="rdoDtOpt2" value="<%=date%>" <%=chkIIF(selDt=cStr(date),"checked","")%> /><label for="rdoDtOpt2">오늘</label><input type="radio" name="selDatePreset" id="rdoDtOpt3" value="<%=dateadd("d",1,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",1,date)),"checked","")%> /><label for="rdoDtOpt3">+1</label><input type="radio" name="selDatePreset" id="rdoDtOpt4" value="<%=dateadd("d",2,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",2,date)),"checked","")%> /><label for="rdoDtOpt4">+2</label><input type="radio" name="selDatePreset" id="rdoDtOpt5" value="<%=dateadd("d",3,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",3,date)),"checked","")%> /><label for="rdoDtOpt5">+3</label><input type="radio" name="selDatePreset" id="rdoDtOpt6" value="<%=dateadd("d",4,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",4,date)),"checked","")%> /><label for="rdoDtOpt6">+4</label><input type="radio" name="selDatePreset" id="rdoDtOpt7" value="<%=dateadd("d",5,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",5,date)),"checked","")%> /><label for="rdoDtOpt7">+5</label><input type="radio" name="selDatePreset" id="rdoDtOpt8" value="<%=dateadd("d",6,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",6,date)),"checked","")%> /><label for="rdoDtOpt8">+6</label><input type="radio" name="selDatePreset" id="rdoDtOpt9" value="<%=dateadd("d",7,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",7,date)),"checked","")%> /><label for="rdoDtOpt9">+7</label>
		</span>
		<input type="text" id="sDt" name="sDt" size="10" value="<%=sDt%>" style="height:22px;" /> ~
		<input type="text" id="eDt" name="eDt" size="10" value="<%=eDt%>" style="height:22px;" />
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td align="left">
    	<input type="button" value="전체선택" class="button" onClick="chkAllItem()">
    	<input type="button" value="상태저장" class="button" onClick="saveList()" title="우선순위 및 선노출여부를 일괄저장합니다.">
    	/
		<input type="button" value="미리보기" class="button" onClick="previewPage()" style="background-color:#F0FFF0" title="예상 프론트페이지를 미리봅니다.">
    	/
    	<% if C_ADMIN_AUTH then %>
		<input type="text" class="text" name="cTrm" value="0" size="1" style="text-align:center;" title="오늘이전 날짜를 지정합니다.(관리자용)">
		<% else %>
		<input type="hidden" name="cTrm" value="0">
		<% end if %>
    	<% if siteDiv="M" and pageDiv="10" then %><input type="button" value="선택 적용" class="button" onClick="assignPage()" style="background-color:#F8F8E8" title="프론트페이지에 선택 적용합니다."><% end if %>
    	<input type="button" value="전체 적용" class="button" onClick="assignPageALL()" style="background-color:#FFF0F0" title="프론트페이지에 전체 적용합니다.">
    </td>
    <td align="right">
    	<% if C_ADMIN_AUTH then %>
		<input type="button" class="button" value="템플릿관리" onClick="popTemplateManage();">&nbsp;
		<% end if %>
		<select class="select" onchange="fnQuitReg(this);">
			<option value="">-- 빠른등록 --</option>
			<%
				if oTemplate.FResultCount>0 then
					for lp=0 to (oTemplate.FResultCount-1)
						Response.Write "<option value='" & oTemplate.FItemList(lp).FtplIdx & "' " & chkIIF(cStr(oTemplate.FItemList(lp).FtplIdx)=tplIdx,"selected","") & ">" & oTemplate.FItemList(lp).FtplName & "</option>"
					next
				end if
			%>
		</select>&nbsp;
    	<input type="button" value="컨텐츠 등록" class="button" onClick="goMainContent('');">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 목록 시작 -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="main">
<input type="hidden" name="sTrm" value="0">
<input type="hidden" name="chkAll" value="N">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		검색결과 : <b><%=oMainCont.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oMainCont.FtotalPage%></b>
	</td>
</tr>
<colgroup>
    <col width="30" />
    <col width="50" />
    <col width="80" />
    <col width="120" />
    <col width="*" />
    <col width="150" />
    <col width="60" />
    <col width="110" />
    <col width="80" />
    <col width="70" />
</colgroup>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>&nbsp;</td>
    <td>번호</td>
    <td>진행상태</td>
    <td>템플릿</td>
    <td>제목</td>
    <td>노출기간</td>
    <td>우선<br>순위</td>
    <td>선노출여부</td>
    <td>등록자</td>
    <td>작업자</td>
</tr>
<tbody id="mainList">
<%	for lp=0 to oMainCont.FResultCount - 1 %>
<tr align="center" bgcolor="<%=chkIIF(oMainCont.FItemList(lp).IsExpired,"#DDDDDD","#FFFFFF")%>">
    <td><input type="checkbox" name="chkIdx" value="<%=oMainCont.FItemList(lp).FmainIdx%>" /></td>
    <td><a href="javascript:goMainContent(<%=oMainCont.FItemList(lp).FmainIdx%>)"><%=oMainCont.FItemList(lp).FmainIdx%></a></td>
    <td><%=oMainCont.FItemList(lp).getMainStat%></td>
    <td><%=oMainCont.FItemList(lp).FtplName%></td>
    <td align="left"><a href="javascript:goMainContent(<%=oMainCont.FItemList(lp).FmainIdx%>)"><%=oMainCont.FItemList(lp).FmainTitle%></a></td>
    <td>
    <%
    	Response.Write "시작: "
    	Response.Write replace(left(oMainCont.FItemList(lp).FmainStartDate,10),"-",".") & " / " & Num2Str(hour(oMainCont.FItemList(lp).FmainStartDate),2,"0","R") & ":" &Num2Str(minute(oMainCont.FItemList(lp).FmainStartDate),2,"0","R")
    	Response.Write "<br />종료: "
    	Response.Write replace(left(oMainCont.FItemList(lp).FmainEndDate,10),"-",".") & " / " & Num2Str(hour(oMainCont.FItemList(lp).FmainEndDate),2,"0","R") & ":" & Num2Str(minute(oMainCont.FItemList(lp).FmainEndDate),2,"0","R")
    %>
    </td>
    <td><input type="text" name="sort<%=oMainCont.FItemList(lp).FmainIdx%>" size="3" class="text" value="<%=oMainCont.FItemList(lp).FmainSortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oMainCont.FItemList(lp).FmainIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oMainCont.FItemList(lp).FmainIsPreOpen="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">노출</label><input type="radio" name="use<%=oMainCont.FItemList(lp).FmainIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oMainCont.FItemList(lp).FmainIsPreOpen="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">없음</label>
		</span>
    </td>
    <td><%=oMainCont.FItemList(lp).FmainRegUsername%></td>
    <td>
    <%
    	modiTime = oMainCont.FItemList(lp).FmainLastModiDate
    	if Not(modiTime="" or isNull(modiTime)) then
	    		Response.Write getStaffUserName(oMainCont.FItemList(lp).FmainLastModiUserid) & "<br />"
	    		Response.Write left(modiTime,10)
	    end if
    %>
    </td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="10" align="center">
    <% if oMainCont.HasPreScroll then %>
		<a href="javascript:goPage('<%= oMainCont.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oMainCont.StartScrollPage to oMainCont.FScrollCount + oMainCont.StartScrollPage - 1 %>
		<% if lp>oMainCont.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oMainCont.HasNextScroll then %>
		<a href="javascript:goPage('<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<form name="frmQuitReg" method="POST" action="doQuickMainCont.asp" style="margin:0;">
<input type="hidden" name="site" value="<%=siteDiv%>" />
<input type="hidden" name="pDiv" value="<%=pageDiv%>" />
<input type="hidden" name="tplIdx" value="" />
<input type="hidden" name="StartDate" value="" />
<input type="hidden" name="EndDate" value="" />
<input type="hidden" name="sTm" value="00:00:00" />
<input type="hidden" name="eTm" value="23:59:59" />
<input type="hidden" name="mainTitle" value="" />
</form>
<!-- 목록 끝 -->
<%
	set oTemplate = Nothing
	set oMainCont = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->