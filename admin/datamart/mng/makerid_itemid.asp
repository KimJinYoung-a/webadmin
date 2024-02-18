<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include file="./makerid_itemid_cls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp, c
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// 즐겨찾기
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<style type="text/css">
.fontstrong {font-weight:bold;}
.fontred {color:#FF0000 !important;}
.fontblue {color:#0000FF !important;}
.bgitemtt {background-color:#FaFaFa;}
.bgred {background-color:#FFD6D6;}
.bgblue {background-color:#BFE9FF;}
.bgredtt {background-color:#FF5F5F;}
.bgbluett {background-color:#39A5FD;}
.bggraytt {background-color:#F3F3F3;}
<% For c=2 To 15 %>
.bgcolor<%=c%> {}
<% Next %>
</style>
<script language='javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;

	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "즐겨찾기에서 제외하시겠습니까?";
	} else {
		msg = "즐겨찾기에 추가하시겠습니까?";
	}

	ret = confirm(msg);

	if (ret) {
		frm.submit();
	}
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<%
	Dim cMI, vArr0, vArr1, vArr2, i, k, vGubun, vDateType, vCompareKey, vSDate, vEDate, vSDate2, vEDate2, vRowTot, vRowTotB, vComP, vTmpCom, vRowBody
	Dim vTot101, vTot102, vTot103, vTot104, vTot124, vTot121, vTot122, vTot120, vTot112, vTot119, vTot117, vTot116, vTot125, vTot118, vTot115, vTot110, vTot000, vTotAll
	Dim vTot101B, vTot102B, vTot103B, vTot104B, vTot124B, vTot121B, vTot122B, vTot120B, vTot112B, vTot119B, vTot117B, vTot116B, vTot125B, vTot118B, vTot115B, vTot110B, vTot000B, vTotAllB
	Dim vUpDownPercent, vUpVal, vDownVal, vUpValTmp, vDownValTmp, vItemTotalCount


	vTotTmpB = vTot101B & "," & vTot102B & "," & vTot103B & "," & vTot104B & "," & vTot124B & "," & vTot121B & "," & vTot122B & "," & vTot120B & ","
	vTotTmpB = vTotTmpB & vTot112B & "," & vTot119B & "," & vTot117B & "," & vTot116B & "," & vTot125B & "," & vTot118B & "," & vTot115B & "," & vTot110B

	vGubun = NullFillWith(requestCheckVar(request("gubun"),1),"1")
	vCompareKey = NullFillWith(requestCheckVar(request("comparekey"),1),"")
	vSDate = NullFillWith(requestCheckVar(request("sdate"),10),DateAdd("m",-1,date()))
	vEDate = NullFillWith(requestCheckVar(request("edate"),10),date())
	vSDate2 = requestCheckVar(request("sdate2"),10)
	vEDate2 = requestCheckVar(request("edate2"),10)

	If vCompareKey = "o" Then
		If DateDiff("m",vSDate,vEDate) > 3 Then
			Response.Write "<script>alert('비교옵션은 3개월 안으로 검색해주세요.\n느려지는 부담이 있습니다.');history.back();</script>"
			dbget.close
			Response.End
		End If
	End If


	Set cMI = New CMIS
	vArr0 = cMI.fnGetItemTotalCountByDisp
	vItemTotalCount = cMI.FItemTotalCount

	cMI.FRectSDate = vSDate
	cMI.FRectEDate = vEDate

	If vGubun = "1" Then
		vArr1 = cMI.fnGetItemSellStdateByDisp
	ElseIf vGubun = "2" Then
		vArr1 = cMI.fnGetUserCitemSellStdateByDisp
	End If

	If vCompareKey = "o" Then

		cMI.FRectSDate = vSDate2
		cMI.FRectEDate = vEDate2

		If vGubun = "1" Then
			vArr2 = cMI.fnGetItemSellStdateByDisp
		ElseIf vGubun = "2" Then
			vArr2 = cMI.fnGetUserCitemSellStdateByDisp
		End If
	Else
	End If
	Set cMI = Nothing

%>
<script>
$(function() {
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:12px;");
});

function goScheduleReg(jobidx){
	location.href = "db_schedule_write.asp?jobidx="+jobidx+"";
}
function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=400, height=300');
	winCal.focus();
}
function jsDateCompare(a,b){
	var startArray = a.split("-");
	var endArray = b.split("-");

	var start_date = new Date(startArray[0], startArray[1], startArray[2]);
	var end_date = new Date(endArray[0], endArray[1], endArray[2]);

	if(start_date.getTime() > end_date.getTime()){
		return "x";
	}else{
		return "o";
	}
}
function jsGoSearch(){
	if(jsDateCompare($("#sdate").val(),$("#edate").val()) == "x"){
		alert("시작일은 종료일의 이전 날짜로 선택해주세요.");
		return;
	}

	if($("#comparekey").is(":checked")){
		if($("#sdate2").val() == ""){
			alert("시작일을 선택 후 검색하세요.");
			return;
		}
		if($("#edate2").val() == ""){
			alert("종료일을 선택 후 검색하세요.");
			return;
		}

		if(jsDateCompare($("#sdate2").val(),$("#edate2").val()) == "x"){
			alert("비교시작일은 비교종료일의 이전 날짜로 선택해주세요.");
			return;
		}
	}

	$("#realbtn").hide();
	$("#loadingbtn").show();
	frm1.action = "makerid_itemid.asp";
	frm1.submit();
}
function jsGoSearchGubun(g){
	frm1.action = "makerid_itemid.asp";
	frm1.submit();
}
function jsCompareKey(){
	if($("#comparekey").is(":checked")){
		$("#sdate2").prop("disabled", false);
		$("#edate2").prop("disabled", false);
	}else{
		$("#sdate2").prop("disabled", true);
		$("#edate2").prop("disabled", true);
	}
}
function jsExcelDown(){
	//alert($("#excel_view").html());
	$("#excel_val").val($("#excel_view").html());
	 $("#exe").submit();
	//frm1.action = "makerid_itemid_xls.asp";
	//frm1.submit();
}
</script>


<form action="makerid_itemid_xls.asp" method="post" target="_blank" id="exe">
<input type="hidden" name="excel_val" id="excel_val" />
</form>
<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2>DATAMART &gt; <strong>텐바이텐 신규입점/상품등록 현황</strong></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="3942">
			</form>
			<% if (menupos > 1) then %>
				<% if (IsMenuFavoriteAdded) then %>
					<a href="javascript:fnMenuFavoriteAct('delonefavorite')">즐겨찾기</a> l
				<% else %>
					<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l
				<% end if %>
			<% end if %>
			<!-- 마스터이상 메뉴권한 설정 //-->
			<% if C_ADMIN_AUTH then %>
			<a href="Javascript:PopMenuEdit('3942');">권한변경</a> l
			<% end if %>
			<!-- Help 설정 //-->
			<% if (imenuposhelp<>"") or (C_ADMIN_AUTH) then %>
			<a href="Javascript:PopMenuHelp('3942');">HELP</a>
			<% end if %>
		</div>
	</div>

	<!-- 상단 검색폼 시작 -->
	<form name="frm1" id="frm1" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap2">
		<div class="search rowSum1">
			<ul>
				<li>
					<span class="rdoUsing">
						<input type="radio" name="gubun" id="gubun_1" value="1" onClick="jsGoSearchGubun('1');" <%=CHKIIF(vGubun="1"," checked","")%> /><label for="gubun_1">상품등록수</label>
						<input type="radio" name="gubun" id="gubun_2" value="2" onClick="jsGoSearchGubun('2');" <%=CHKIIF(vGubun="2"," checked","")%> /><label for="gubun_2">신규입점수</label>
					</span>
				</li>
				<li class="lMar10 rMar10">
					상품등록일 :
					<input type="text" name="sdate" id="sdate" value="<%=vSDate%>" style="text-align:center;height:35px;" size="10" maxlength="10" readonly>
					<strong>&nbsp;~&nbsp;</strong>
					<input type="text" name="edate" id="edate" value="<%=vEDate%>" style="text-align:center;height:35px;" size="10" maxlength="10" readonly>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "sdate", trigger    : "sdate",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "edate", trigger    : "edate",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
				<li class="lMar10">
					<span class="lMar10"><input type="checkbox" name="comparekey" id="comparekey" value="o" style="width: 1.5em; height: 1.5em;" onClick="jsCompareKey();" <%=CHKIIF(vCompareKey="o","checked","")%> /> 비교</span>
					<input type="text" name="sdate2" id="sdate2" value="<%=vSDate2%>" placeholder="시작일선택" style="text-align:center;height:35px;" size="10" maxlength="10" class="lMar10" readonly <%=CHKIIF(vCompareKey="o","","disabled")%>>
					<strong>&nbsp;~&nbsp;</strong>
					<input type="text" name="edate2" id="edate2" value="<%=vEDate2%>" placeholder="종료일선택" style="text-align:center;height:35px;" size="10" maxlength="10" readonly <%=CHKIIF(vCompareKey="o","","disabled")%>>
					<script type="text/javascript">
						var CAL_Start2 = new Calendar({
							inputField : "sdate2", trigger    : "sdate2",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End2.args.min = date;
								CAL_End2.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End2 = new Calendar({
							inputField : "edate2", trigger    : "edate2",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start2.args.max = date;
								CAL_Start2.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
			</ul>
		</div>
		<span id="realbtn"><input type="button" class="schBtn2" value="검색" onClick="jsGoSearch();" /></span>
		<span id="loadingbtn" style="display:none;"><input type="button" class="schBtn2" value="검색" onClick="alert('실행중입니다.\n조금만 기다려주세요.');" /></span>
		<input type="button" class="resetBtn" value="초기화" onClick="location.href='/admin/datamart/mng/makerid_itemid.asp';" />
	</div>
	</form>

	<div class="pad20">
		<div class="tPad10">
			<div class="overHidden pad10">
				<% If vCompareKey = "o" Then %>
				<div class="ftLt">비교일 대비 일별 증감율이 높을 경우 <strong class="fontred">붉은색</strong>으로, 낮을 경우 <strong class="fontblue">푸른색</strong>으로 표기</div>
				<% End If %>
				<div class="ftRt"><input type="image" src="/images/excel_download.png" onClick="jsExcelDown();" /></div>
			</div>
			<span  id="excel_view">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div></div></th>
					<th><div>디자인<br />문구</div></th>
					<th><div>디지털/<br />핸드폰</div></th>
					<th><div>캠핑/<br />트래블</div></th>
					<th><div>&nbsp;토이&nbsp;</div></th>
					<th><div>디자인가전</div></th>
					<th><div>가구/<br />수납</div></th>
					<th><div>데코/<br />조명</div></th>
					<th><div>패브릭/<br />생활</div></th>
					<th><div>&nbsp;키친&nbsp;</div></th>
					<th><div>&nbsp;푸드&nbsp;</div></th>
					<th><div>패션의류</div></th>
					<th><div>패션잡화</div></th>
					<th><div>주얼리/<br />시계</div></th>
					<th><div>&nbsp;뷰티&nbsp;</div></th>
					<th><div>베이비/<br />키즈</div></th>
					<th><div>Cat&Dog</div></th>
					<th><div>미지정</div></th>
					<th><div>합계</div></th>
				</tr>
				</thead>
				<tbody>
				<tr>
					<td class="ct">전체 상품 등록 수</td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(0,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(1,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(2,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(3,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(4,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(5,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(6,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(7,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(8,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(9,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(10,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(11,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(12,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(13,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(14,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(15,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vArr0(16,0),0)%></td>
					<td class="ct fontstrong bgitemtt"><%=FormatNumber(vItemTotalCount,0)%></td>
				</tr>
				<tr>
					<td class="ct">통계합계</td>
					<td class="ct fontstrong bggraytt" id="tot101"></td>
					<td class="ct fontstrong bggraytt" id="tot102"></td>
					<td class="ct fontstrong bggraytt" id="tot103"></td>
					<td class="ct fontstrong bggraytt" id="tot104"></td>
					<td class="ct fontstrong bggraytt" id="tot124"></td>
					<td class="ct fontstrong bggraytt" id="tot121"></td>
					<td class="ct fontstrong bggraytt" id="tot122"></td>
					<td class="ct fontstrong bggraytt" id="tot120"></td>
					<td class="ct fontstrong bggraytt" id="tot112"></td>
					<td class="ct fontstrong bggraytt" id="tot119"></td>
					<td class="ct fontstrong bggraytt" id="tot117"></td>
					<td class="ct fontstrong bggraytt" id="tot116"></td>
					<td class="ct fontstrong bggraytt" id="tot125"></td>
					<td class="ct fontstrong bggraytt" id="tot118"></td>
					<td class="ct fontstrong bggraytt" id="tot115"></td>
					<td class="ct fontstrong bggraytt" id="tot110"></td>
					<td class="ct fontstrong bggraytt" id="tot000"></td>
					<td class="ct fontstrong bggraytt" id="totall"></td>
				</tr>
				<%
					'0~9 : yyyymm, weekname, d101cnt, d102cnt, d103cnt, d104cnt, d124cnt, d121cnt, d122cnt, d120cnt,
					'10~17 : d112cnt, d119cnt, d117cnt, d116cnt, d125cnt, d118cnt, d115cnt, d110cnt
					If isArray(vArr1) Then
						For i=0 To UBound(vArr1,2)



							Response.Write "<tr>" & vbCrLf
							Response.Write "	<td"
							If vArr1(1,i) = "일" Then
								Response.Write " class='fontred'"
							ElseIf vArr1(1,i) = "토" Then
								Response.Write " class='fontblue'"
							End If

							Response.Write ">" & vArr1(0,i) & " [" & vArr1(1,i) & "]"
							If vSDate2 <> "" AND vEDate2 <> "" Then
								Response.Write "<br /><span class='cGy1 fs11'>" & CHKIIF(DateAdd("d",-(i),vEDate2)<CDate(vSDate2),"","(" & DateAdd("d",-(i),vEDate2) & " " & fnWeekNameReturn(DatePart("w",DateAdd("d",-(i),vEDate2))) & ")</span>")
							End If
							Response.Write "</td>" & vbCrLf

							For k=2 To 18
								If vCompareKey = "o" Then
									If i <= UBound(vArr2,2) Then

										vComP = vArr2(k,i)
										vComP = "<br /><span class='cGy1 fs11'>(" & vComP & ")</span>"

										vUpDownPercent = fnCompareValue(vArr1(k,i), vArr2(k,i)) & "|" & k

										If k = 2 Then
											vUpVal = vUpDownPercent
											vDownVal = vUpDownPercent
										Else
											vUpVal = fnCompareUpDownValue("up",vUpDownPercent,vUpValTmp)
											vDownVal = fnCompareUpDownValue("down",vUpDownPercent,vDownValTmp)
										End If

										vUpValTmp = vUpVal
										vDownValTmp = vDownVal

									End If
								End If
								vRowBody = vRowBody & "	<td class='ct bgcolor"&k&"'>" & FormatNumber(vArr1(k,i),0) & vComP & "</td>" & vbCrLf
								vComP = ""
							Next

							If vCompareKey = "o" Then
								If vUpVal <> "" Then
									vRowBody = Replace(vRowBody,"bgcolor"&Split(vUpVal,"|")(1)&"","bgred")
								End If
								If vDownVal <> "" Then
									vRowBody = Replace(vRowBody,"bgcolor"&Split(vDownVal,"|")(1)&"","bgblue")
								End If
								Response.Write vRowBody
							Else
								Response.Write vRowBody
							End If
							vRowBody = ""


							vRowTot = vArr1(2,i) + vArr1(3,i) + vArr1(4,i) + vArr1(5,i) + vArr1(6,i) + vArr1(7,i) + vArr1(8,i) + vArr1(9,i) + vArr1(10,i) + vArr1(11,i) + vArr1(12,i) + vArr1(13,i) + vArr1(14,i) + vArr1(15,i) + vArr1(16,i) + vArr1(17,i) + vArr1(18,i)
							If vCompareKey = "o" Then
								If i <= UBound(vArr2,2) Then
								vRowTotB = vArr2(2,i) + vArr2(3,i) + vArr2(4,i) + vArr2(5,i) + vArr2(6,i) + vArr2(7,i) + vArr2(8,i) + vArr2(9,i) + vArr2(10,i) + vArr2(11,i) + vArr2(12,i) + vArr2(13,i) + vArr2(14,i) + vArr2(15,i) + vArr2(16,i) + vArr2(17,i) + vArr2(18,i)
								Else
									vRowTotB = ""
								End If
							End If

							Response.Write "	<td class='ct'>" & FormatNumber(vRowTot,0) & CHKIIF(vRowTotB<>"","<br />("&vRowTotB&")","") & "</td>" & vbCrLf
							Response.Write "</tr>" & vbCrLf
							vUpVal = ""
							vDownVal = ""
							vUpValTmp = ""
							vDownValTmp = ""

							vTot101 = vTot101 + vArr1(2,i)
							vTot102 = vTot102 + vArr1(3,i)
							vTot103 = vTot103 + vArr1(4,i)
							vTot104 = vTot104 + vArr1(5,i)
							vTot124 = vTot124 + vArr1(6,i)
							vTot121 = vTot121 + vArr1(7,i)
							vTot122 = vTot122 + vArr1(8,i)
							vTot120 = vTot120 + vArr1(9,i)
							vTot112 = vTot112 + vArr1(10,i)
							vTot119 = vTot119 + vArr1(11,i)
							vTot117 = vTot117 + vArr1(12,i)
							vTot116 = vTot116 + vArr1(13,i)
							vTot125 = vTot125 + vArr1(14,i)
							vTot118 = vTot118 + vArr1(15,i)
							vTot115 = vTot115 + vArr1(16,i)
							vTot110 = vTot110 + vArr1(17,i)
							vTot000 = vTot000 + vArr1(18,i)
							vTotAll = vTotAll + vRowTot

							If vCompareKey = "o" Then
								If i <= UBound(vArr2,2) Then
									vTot101B = vTot101B + vArr2(2,i)
									vTot102B = vTot102B + vArr2(3,i)
									vTot103B = vTot103B + vArr2(4,i)
									vTot104B = vTot104B + vArr2(5,i)
									vTot124B = vTot124B + vArr2(6,i)
									vTot121B = vTot121B + vArr2(7,i)
									vTot122B = vTot122B + vArr2(8,i)
									vTot120B = vTot120B + vArr2(9,i)
									vTot112B = vTot112B + vArr2(10,i)
									vTot119B = vTot119B + vArr2(11,i)
									vTot117B = vTot117B + vArr2(12,i)
									vTot116B = vTot116B + vArr2(13,i)
									vTot125B = vTot125B + vArr2(14,i)
									vTot118B = vTot118B + vArr2(15,i)
									vTot115B = vTot115B + vArr2(16,i)
									vTot110B = vTot110B + vArr2(17,i)
									vTot000B = vTot000B + vArr2(18,i)
									vTotAllB = vTotAllB + vRowTotB
								End If
							End If

						Next
					End If
				%>
				</tr>
				</tbody>
			</table>
			</span>
			<br />
			<div class="ct tPad20 cBk1">

			</div>
		</div>
	</div>
</div>
<% If vTotAll <> "" Then %>
<script>
$("#tot101").html("<%=FormatNumber(vTot101,0)%><%=CHKIIF(vTot101B<>"","<br />("&FormatNumber(vTot101B,0)&")","")%>");
$("#tot102").html("<%=FormatNumber(vTot102,0)%><%=CHKIIF(vTot102B<>"","<br />("&FormatNumber(vTot102B,0)&")","")%>");
$("#tot103").html("<%=FormatNumber(vTot103,0)%><%=CHKIIF(vTot103B<>"","<br />("&FormatNumber(vTot103B,0)&")","")%>");
$("#tot104").html("<%=FormatNumber(vTot104,0)%><%=CHKIIF(vTot104B<>"","<br />("&FormatNumber(vTot104B,0)&")","")%>");
$("#tot124").html("<%=FormatNumber(vTot124,0)%><%=CHKIIF(vTot124B<>"","<br />("&FormatNumber(vTot124B,0)&")","")%>");
$("#tot121").html("<%=FormatNumber(vTot121,0)%><%=CHKIIF(vTot121B<>"","<br />("&FormatNumber(vTot121B,0)&")","")%>");
$("#tot122").html("<%=FormatNumber(vTot122,0)%><%=CHKIIF(vTot122B<>"","<br />("&FormatNumber(vTot122B,0)&")","")%>");
$("#tot120").html("<%=FormatNumber(vTot120,0)%><%=CHKIIF(vTot120B<>"","<br />("&FormatNumber(vTot120B,0)&")","")%>");
$("#tot112").html("<%=FormatNumber(vTot112,0)%><%=CHKIIF(vTot112B<>"","<br />("&FormatNumber(vTot112B,0)&")","")%>");
$("#tot119").html("<%=FormatNumber(vTot119,0)%><%=CHKIIF(vTot119B<>"","<br />("&FormatNumber(vTot119B,0)&")","")%>");
$("#tot117").html("<%=FormatNumber(vTot117,0)%><%=CHKIIF(vTot117B<>"","<br />("&FormatNumber(vTot117B,0)&")","")%>");
$("#tot116").html("<%=FormatNumber(vTot116,0)%><%=CHKIIF(vTot116B<>"","<br />("&FormatNumber(vTot116B,0)&")","")%>");
$("#tot125").html("<%=FormatNumber(vTot125,0)%><%=CHKIIF(vTot125B<>"","<br />("&FormatNumber(vTot125B,0)&")","")%>");
$("#tot118").html("<%=FormatNumber(vTot118,0)%><%=CHKIIF(vTot118B<>"","<br />("&FormatNumber(vTot118B,0)&")","")%>");
$("#tot115").html("<%=FormatNumber(vTot115,0)%><%=CHKIIF(vTot115B<>"","<br />("&FormatNumber(vTot115B,0)&")","")%>");
$("#tot110").html("<%=FormatNumber(vTot110,0)%><%=CHKIIF(vTot110B<>"","<br />("&FormatNumber(vTot110B,0)&")","")%>");
$("#tot000").html("<%=FormatNumber(vTot000,0)%><%=CHKIIF(vTot000B<>"","<br />("&FormatNumber(vTot000B,0)&")","")%>");
$("#totall").html("<%=FormatNumber(vTotAll,0)%><%=CHKIIF(vTotAllB<>"","<br />("&FormatNumber(vTotAllB,0)&")","")%>");
<%
If vCompareKey = "o" Then
	Dim vTotTmp, vTotTmpB, vTotID
	vUpValTmp = ""
	vDownValTmp = ""
	vTotID = "101,102,103,104,124,121,122,120,112,119,117,116,125,118,115,110"

	vTotTmp = vTot101 & "," & vTot102 & "," & vTot103 & "," & vTot104 & "," & vTot124 & "," & vTot121 & "," & vTot122 & "," & vTot120 & ","
	vTotTmp = vTotTmp & vTot112 & "," & vTot119 & "," & vTot117 & "," & vTot116 & "," & vTot125 & "," & vTot118 & "," & vTot115 & "," & vTot110 & "," & vTot000

	vTotTmpB = vTot101B & "," & vTot102B & "," & vTot103B & "," & vTot104B & "," & vTot124B & "," & vTot121B & "," & vTot122B & "," & vTot120B & ","
	vTotTmpB = vTotTmpB & vTot112B & "," & vTot119B & "," & vTot117B & "," & vTot116B & "," & vTot125B & "," & vTot118B & "," & vTot115B & "," & vTot110B & "," & vTot000B

	For k=0 To 16
		vUpDownPercent = fnCompareValue(CDbl(Split(vTotTmp,",")(k)), CDbl(Split(vTotTmpB,",")(k))) & "|" & k

		If k = 0 Then
			vUpVal = vUpDownPercent
			vDownVal = vUpDownPercent
		Else
			vUpVal = fnCompareUpDownValue("up",vUpDownPercent,vUpValTmp)
			vDownVal = fnCompareUpDownValue("down",vUpDownPercent,vDownValTmp)
		End If
'response.write vUpDownPercent & ",,," & vUpValTmp & "=" & vUpVal & vbCrLf

		vUpValTmp = vUpVal
		vDownValTmp = vDownVal
	Next

	'' 뭔가 비교시에 split 오류가 있는득. Response.Write "$('#tot"&Split(vTotID,",")(split(vUpVal,"|")(1))&"').removeClass('bggraytt').addClass('bgredtt');" & vbCrLf
	''?? Response.Write "$('#tot"&Split(vTotID,",")(split(vDownVal,"|")(1))&"').removeClass('bggraytt').addClass('bgbluett');" & vbCrLf
End If
%>
</script>
<% End If %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
