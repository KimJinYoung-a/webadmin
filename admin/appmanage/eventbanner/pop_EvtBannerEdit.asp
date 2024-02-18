<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/appmanage/eventBannerCls.asp" -->
<%
'###############################################
' PageName : pop_EvtBannerEdit.asp
' Discription : 이벤트 매너 등록/수정
' History : 2014.03.28 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim i, page
Dim oEvtBanner
Dim Idx, appName, startDate, endDate, eventName, sortNo, bannerType, bannerImg, bannerLink, isUsing, regUserid, regdate, lastUpdateUser, lastUpdate, workComment
Dim startTime, endTime


'// 파라메터 접수
idx = request("idx")
appName = request("appName")
if appName="" then appName="wishapp"

'// 템플릿 내용
	set oEvtBanner = new CEvtBanner
	oEvtBanner.FRectIdx = idx
    if idx<>"" then
    	oEvtBanner.GetOneEvtBanner()
		if oEvtBanner.FResultCount>0 then
            appName			= oEvtBanner.FOneItem.FappName
			startDate		= left(oEvtBanner.FOneItem.FstartDate,10)
            endDate			= left(oEvtBanner.FOneItem.FendDate,10)
            eventName		= oEvtBanner.FOneItem.FeventName
            sortNo			= oEvtBanner.FOneItem.FsortNo
            bannerType		= oEvtBanner.FOneItem.FbannerType
            bannerImg		= oEvtBanner.FOneItem.FbannerImg
            bannerLink		= oEvtBanner.FOneItem.FbannerLink
            isUsing			= oEvtBanner.FOneItem.FisUsing
            regUserid		= oEvtBanner.FOneItem.FregUserid
            regdate			= oEvtBanner.FOneItem.Fregdate
            lastUpdateUser	= oEvtBanner.FOneItem.FlastUpdateUser
            lastUpdate		= oEvtBanner.FOneItem.FlastUpdate
            workComment		= oEvtBanner.FOneItem.FworkComment

            startTime		= Num2Str(hour(oEvtBanner.FOneItem.FstartDate),2,"0","R") & ":" & Num2Str(minute(oEvtBanner.FOneItem.FstartDate),2,"0","R") & ":" & Num2Str(second(oEvtBanner.FOneItem.FstartDate),2,"0","R")
            endTime			= Num2Str(hour(oEvtBanner.FOneItem.FendDate),2,"0","R") & ":" & Num2Str(minute(oEvtBanner.FOneItem.FendDate),2,"0","R") & ":" & Num2Str(second(oEvtBanner.FOneItem.FendDate),2,"0","R")
		end if
    else
    	startDate		= date
    	EndDate			= date
    	bannerType		= "H"
    	startTime		= "00:00:00"
    	endTime			= "23:59:59"
    	regdate = now()
    	sortNo = "50"
    	isUsing="N"
    end if
    set oEvtBanner = Nothing
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function(){
	//라디오 버튼
	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// 캘린더
	var arrDayMin = ["일","월","화","수","목","금","토"];
	var arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
    $("#startDate").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 1,
      	showOn: "button",
    	onClose: function() {
    		if($("#startDate").datepicker("getDate")>$("#endDate").datepicker("getDate")) {
    			$("#endDate").datepicker("setDate",$("#startDate").datepicker("getDate"));
    		}
    	}
    });
    $("#endDate").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 1,
      	showOn: "button",
    	onClose: function() {
    		if($("#endDate").datepicker("getDate")<$("#startDate").datepicker("getDate")) {
    			$("#startDate").datepicker("setDate",$("#endDate").datepicker("getDate"));
    		}
    	}
    });
});

// 폼검사
function SaveEvtBanner(frm) {
	var selChk=true;
	$("select").each(function(){
		if($(this).val()=="") {
			alert($(this).attr("title")+"을(를) 선택해주세요");
			$(this).focus();
			selChk=false;
			return false;
		}
	});
	if(!selChk) return;

	if($("input[name='eventName']").val()=="") {
		alert("제목을 입력해주세요.");
		$("input[name='eventName']").focus();
		selChk=false;
	}

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}

function jsSetImg(sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('pop_evtBanner_upload.asp?yr=<%=Year(regdate)%>&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   $("#"+sName).val('');
	   $("#"+sSpan).fadeOut();
	}
}

function fnDefaulUrl(v) {
	if(v=="e") {
		$("input[name='bannerLink']").val("/event/eventmain.asp?eventid=");
	} else if(v=="i") {
		$("input[name='bannerLink']").val("/category/category_itemPrd.asp?itemid=");
	}
}
</script>
<center>
<form name="frmEvtBanner" method="post" action="doEvtBanner.asp" style="margin:0px;">
<table width="690" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>이벤트 배너 등록/수정</b></td>
</tr>
<% if idx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">번호</td>
    <td width="610" colspan="3">
        <%=idx %>
        <input type="hidden" name="idx" value="<%=idx %>" />
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">사용처</td>
    <td width="230">
        <select name="appName" class="select" title="사용처">
        	<option value="">::선택::</option>
			<option value="wishapp" <%=chkIIF(appName="wishapp","selected","")%> >위시 APP</option>
			<option value="calapp" <%=chkIIF(appName="calapp","selected","")%> >캘린더 APP</option>
			<option value="hitchhiker" <%=chkIIF(appName="hitchhiker","selected","")%> >히치하이커</option>
        </select>
    </td>
    <td width="100" bgcolor="#DDDDFF">배너형태</td>
    <td width="230">
		<select name="bannerType" class="select" title="배너형태">
			<option value="">전체</option>
			<option value="F" <%=chkIIF(bannerType="F","selected","")%> >풀배너</option>
			<option value="H" <%=chkIIF(bannerType="H","selected","")%> >하프배너</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">노출여부</td>
    <td width="230">
		<% if idx<>"" then %>
		<span class="rdoUsing">
		<input type="radio" name="isusing" id="rdoUsing_1" value="Y" <%=chkIIF(isUsing="Y","checked","")%> /><label for="rdoUsing_1">노출</label><input type="radio" name="isusing" id="rdoUsing_2" value="N" <%=chkIIF(isUsing="N","checked","")%> /><label for="rdoUsing_2">안함</label>
		</span>
		<% else %>
		<input type="hidden" name="isusing" value="N">
		노출안함<br><span style="color:#D03030;font-size:11px;">※ 최초 등록시 지정불가 (등록 후 변경 요망)</span>
		<% end if %>
    </td>
    <td width="100" bgcolor="#DDDDFF">정렬순서</td>
    <td width="230">
		<input type="text" name="sortNo" class="text" size="4" value="<%=sortNo%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">제목</td>
    <td width="610" colspan="3">
        <input type="text" name="eventName" value="<%= eventName %>" maxlength="64" size="64" title="제목">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">노출기간</td>
    <td width="610" colspan="3">
		<input type="text" id="startDate" name="startDate" size="10" value="<%=startDate%>" style="height:22px;" />
		<input type="text" name="startTime" size="8" value="<%=startTime%>" style="height:22px;"> ~
		<input type="text" id="endDate" name="endDate" size="10" value="<%=endDate%>" style="height:22px;" />
		<input type="text" name="endTime" size="8" value="<%=endTime%>" style="height:22px;">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">작업전달사항</td>
    <td width="610" colspan="3">
		<textarea name="workcomment" style="width:100%; height:90px;"><%=workcomment%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">배너이미지</td>
    <td width="610" colspan="3">
		<input type="hidden" name="bannerImg" id="bannerImg" value="<%=bannerImg%>">
		<input type="button" value="이미지 등록" onClick="jsSetImg('<%=bannerImg%>','bannerImg','spanBanner')" class="button"> <span style="color:#D03030;font-size:11px;">※ 600×282px</span>
		<div id="spanBanner" style="padding: 5 5 5 5">
			<%IF bannerImg <> "" THEN %>
			<img src="<%=bannerImg%>" width="400" border="0">
			<a href="javascript:jsDelImg('bannerImg','spanBanner');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">배너 링크</td>
    <td width="610" colspan="3">
        <input type="text" name="bannerLink" value="<%= bannerLink %>" maxlength="64" size="64"><br>
        <span onclick="fnDefaulUrl('e')" style="color:#606060;font-size:11px;cursor:pointer;">ex #1) /event/eventmain.asp?eventid=이벤트번호</span><br>
        <span onclick="fnDefaulUrl('i')" style="color:#606060;font-size:11px;cursor:pointer;">ex #2) /category/category_itemPrd.asp?itemid=상품번호</span><br>
        <span style="color:#D03030;font-size:11px;">(APP내 경로로 전환되니 절대로 Full URL을 입력하지 마세요)</span>
    </td>
</tr>



<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" 저 장 " onClick="SaveEvtBanner(this.form);"></td>
</tr>
</table>
</form>
</center>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->