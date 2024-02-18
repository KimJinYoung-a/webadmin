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
<!-- #include virtual="/lib/classes/appmanage/eventBannerCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : APP 메인 이벤트 배너 관리
' History : 2014.03.27 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim appName, bannerType, isUsing, selDt, sDt, eDt 
Dim oEvtBanner, lp
Dim page

'// 파라메터 접수
appName = request("appName")
isusing = request("isusing")
bannerType = request("bannerType")
sDt = request("sDt")
eDt = request("eDt")
page = request("page")
if appName="" then appName="wishapp"		'기본값 wishApp (wishapp:위시, hitchhiker:히치하이커, calapp:캘린더)
if isusing="" then isusing="A"				'기본값 전체
if sDt="" then sDt=cStr(date)
if eDt="" then eDt=cStr(date)
if sDt=eDt then selDt=sDt
if page="" then page="1"

'// 페이지정보 목록
	set oEvtBanner = new CEvtBanner
	oEvtBanner.FPageSize = 20
	oEvtBanner.FRectAppName = appName
	oEvtBanner.FRectIsUsing = isusing
	oEvtBanner.FRectType = bannerType
	oEvtBanner.FRectStartDate = sDt
	oEvtBanner.FRectEndDate = eDt
    oEvtBanner.GetEvtBannerList
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
	/*
	$( "#mainList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="30" colspan="11" style="border:1px solid #F9BD01;">&nbsp;</td>');
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
	*/
});

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
		alert("수정하실 배너를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.target="_self";
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}

function goEvtBannerent(idx) {
    var popwin = window.open('/admin/appmanage/eventbanner/pop_EvtBannerEdit.asp?idx='+idx+'&appName=<%=appName%>','popEvtBanner','width=750,height=420,scrollbars=yes,resizable=yes');
    popwin.focus();
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
	    사용처:
		<select name="appName" class="select">
			<option value="wishapp" <%=chkIIF(appName="wishapp","selected","")%> >위시 APP</option>
			<option value="calapp" <%=chkIIF(appName="calapp","selected","")%> >캘린더 APP</option>
			<option value="hitchhiker" <%=chkIIF(appName="hitchhiker","selected","")%> >히치하이커</option>
		</select>
		&nbsp;/&nbsp;
	    사용구분:
		<select name="isusing" class="select">
			<option value="A">전체</option>
			<option value="Y" <%=chkIIF(isusing="Y","selected","")%> >사용함</option>
			<option value="N" <%=chkIIF(isusing="N","selected","")%> >사용안함</option>
		</select>
		&nbsp;/&nbsp;
	    배너형태:
		<select name="bannerType" class="select">
			<option value="">전체</option>
			<option value="F" <%=chkIIF(bannerType="F","selected","")%> >풀배너</option>
			<option value="H" <%=chkIIF(bannerType="H","selected","")%> >하프배너</option>
		</select>
	</td>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="검색" />
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
	<td>
		조회기간:
		<span id="rdoDtPreset">
		<input type="radio" name="selDatePreset" id="rdoDtOpt1" value="<%=dateadd("d",-1,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",-1,date)),"checked","")%> /><label for="rdoDtOpt1">-1</label><input type="radio" name="selDatePreset" id="rdoDtOpt2" value="<%=date%>" <%=chkIIF(selDt=cStr(date),"checked","")%> /><label for="rdoDtOpt2">오늘</label><input type="radio" name="selDatePreset" id="rdoDtOpt3" value="<%=dateadd("d",1,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",1,date)),"checked","")%> /><label for="rdoDtOpt3">+1</label><input type="radio" name="selDatePreset" id="rdoDtOpt4" value="<%=dateadd("d",2,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",2,date)),"checked","")%> /><label for="rdoDtOpt4">+2</label><input type="radio" name="selDatePreset" id="rdoDtOpt5" value="<%=dateadd("d",3,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",3,date)),"checked","")%> /><label for="rdoDtOpt5">+3</label><input type="radio" name="selDatePreset" id="rdoDtOpt6" value="<%=dateadd("d",4,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",4,date)),"checked","")%> /><label for="rdoDtOpt6">+4</label><input type="radio" name="selDatePreset" id="rdoDtOpt7" value="<%=dateadd("d",5,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",5,date)),"checked","")%> /><label for="rdoDtOpt7">+5</label>
		</span>
		<input type="text" id="sDt" name="sDt" size="10" value="<%=sDt%>" style="height:22px;" /> ~
		<input type="text" id="eDt" name="eDt" size="10" value="<%=eDt%>" style="height:22px;" />
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
    <td align="left">
    	<input type="button" value="전체선택" class="button" onClick="chkAllItem()">
    	<input type="button" value="상태저장" class="button" onClick="saveList()" title="우선순위 및 노출여부를 일괄저장합니다.">
    </td>
    <td align="right">
    	<input type="button" value="컨텐츠 등록" class="button" onClick="goEvtBannerent('');">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 목록 시작 -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="chkAll" value="N">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		검색결과 : <b><%=oEvtBanner.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oEvtBanner.FtotalPage%></b>
	</td>
</tr>
<colgroup>
    <col width="30" />
    <col width="50" />
    <col width="50" />
    <col width="100" />
    <col width="*" />
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
    <td>타입</td>
    <td>이미지</td>
    <td>제목</td>
    <td>링크</td>
    <td>노출기간</td>
    <td>우선<br>순위</td>
    <td>노출여부</td>
    <td>등록자</td>
    <td>작업자</td>
</tr>
<tbody id="mainList">
<%	for lp=0 to oEvtBanner.FResultCount - 1 %>
<tr align="center" bgcolor="<%=chkIIF(oEvtBanner.FItemList(lp).IsExpired,"#DDDDDD","#FFFFFF")%>">
    <td><input type="checkbox" name="chkIdx" value="<%=oEvtBanner.FItemList(lp).Fidx%>" /></td>
    <td><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><%=oEvtBanner.FItemList(lp).Fidx%></a></td>
    <td><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><%=oEvtBanner.FItemList(lp).getBannerTypeNm%></a></td>
    <td><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><img src="<%=oEvtBanner.FItemList(lp).FbannerImg%>" alt="배너" style="width:94px; border:1px solid #606060;" /></a></td>
    <td align="left"><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><%=oEvtBanner.FItemList(lp).FeventName%></a></td>
    <td align="left"><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><%=oEvtBanner.FItemList(lp).FbannerLink%></a></td>
    <td>
    <%
    	Response.Write "시작: "
    	Response.Write replace(left(oEvtBanner.FItemList(lp).FstartDate,10),"-",".") & " / " & Num2Str(hour(oEvtBanner.FItemList(lp).FstartDate),2,"0","R") & ":" &Num2Str(minute(oEvtBanner.FItemList(lp).FstartDate),2,"0","R")
    	Response.Write "<br />종료: "
    	Response.Write replace(left(oEvtBanner.FItemList(lp).FendDate,10),"-",".") & " / " & Num2Str(hour(oEvtBanner.FItemList(lp).FendDate),2,"0","R") & ":" & Num2Str(minute(oEvtBanner.FItemList(lp).FendDate),2,"0","R")
    %>
    </td>
    <td><input type="text" name="sort<%=oEvtBanner.FItemList(lp).Fidx%>" size="3" class="text" value="<%=oEvtBanner.FItemList(lp).FsortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oEvtBanner.FItemList(lp).Fidx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oEvtBanner.FItemList(lp).FisUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">노출</label><input type="radio" name="use<%=oEvtBanner.FItemList(lp).Fidx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oEvtBanner.FItemList(lp).FisUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">안함</label>
		</span>
    </td>
    <td><%=oEvtBanner.FItemList(lp).FregUsername%></td>
    <td><%=getStaffUserName(oEvtBanner.FItemList(lp).FlastUpdateUser)%>
    </td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="center">
    <% if oEvtBanner.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEvtBanner.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oEvtBanner.StartScrollPage to oEvtBanner.FScrollCount + oEvtBanner.StartScrollPage - 1 %>
		<% if lp>oEvtBanner.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oEvtBanner.HasNextScroll then %>
		<a href="javascript:goPage('<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<div style="text-align:right;">※ [노출함]으로 지정된 배너는 노출기간에 맞춰 자동으로 오픈되며, 풀배너 형태는 하프배너보다 우선적으로 표시됩니다.</div>
<!-- 목록 끝 -->
<%
	set oEvtBanner = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->