<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 배송비 부담금액 수정 팝업
' Hieditor : 2020.08.27 원승현 추가
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/halfdeliverypay/halfdeliverypaycls.asp"-->
<%
Dim i, mode
Dim startdate, enddate, starttime, endtime
Dim idx
dim oHalfDeliveryView, loginUserId, dateModifyCheck

idx = requestCheckvar(request("idx"), 50)

loginUserId = session("ssBctId")

if Trim(idx) = "" then
	response.write "<script>alert('정상적인 경로로 접근해주세요.');window.close();</script>"
	response.end
end If

dateModifyCheck = false

'// halfdeliverypay View 데이터를 가져온다.
set oHalfDeliveryView = new CgetHalfDeliveryPay
	oHalfDeliveryView.FRectIdx = idx
	oHalfDeliveryView.getHalfDeliveryPayview()


if Not(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate="" or isNull(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate)) Then
	starttime = Num2Str(hour(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate),2,"0","R") &":"& Num2Str(minute(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate),2,"0","R") &":"& Num2Str(second(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate),2,"0","R")
else
	starttime = "00:00:00"
end if

if Not(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate="" or isNull(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate)) Then
	endtime = Num2Str(hour(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate),2,"0","R") &":"& Num2Str(minute(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate),2,"0","R") &":"& Num2Str(second(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate),2,"0","R")
else
	endtime = "00:00:00"
end if

'// 수정하려는 상품이 현재 진행되고 있으면 시작일 수정이 안됨
'// 아직 시작 전 이거나 종료된 상품일 경우에만 시작일 수정이 가능
If Cdate(left(now(), 10)) >= Cdate(left(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate,10)) And Cdate(left(now(),10)) < Cdate(dateadd("d", 1, left(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate,10))) Then
	dateModifyCheck = false
Else
	dateModifyCheck = true
End If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
body {background-color:#fff;}
</style>
</head>
<body>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script type='text/javascript'>
document.domain = "10x10.co.kr";

function formatDate(date) { 
	var d = new Date(date), 
	month = '' + (d.getMonth() + 1), 
	day = '' + d.getDate(), 
	year = d.getFullYear(); 
	if (month.length < 2) month = '0' + month; 
	if (day.length < 2) day = '0' + day; 
	return [year, month, day].join('-'); 
}

function frmedit(){
	var frm  = document.frm;
	var today = new Date();

	<% If dateModifyCheck Then %>
		if(frm.startdate.value=="")
		{
			alert('시작일을 입력해 주세요');
			frm.startdate.focus();
			return;
		}

		if(formatDate(today) >= formatDate(new Date(frm.startdate.value))) {
			alert('시작일은 오늘 기준 다음 일자부터 설정하실 수 있습니다.\n오늘 이후의 일자로 입력해주세요.');
			frm.startdate.focus();
			return;
		}
	<% End If %>

	if(frm.enddate.value=="")
	{
		alert('종료일을 입력해 주세요');
		frm.enddate.focus();
		return;
	}

	if(formatDate(new Date(frm.enddate.value)) <= formatDate(new Date(frm.startdate.value))) {
		alert('종료일은 시작일 이후로만 설정하실 수 있습니다.');
		frm.enddate.focus();
		return;
	}
	/*
	if(formatDate(today) >= formatDate(new Date(frm.startdate.value))) {
		alert('수정시 종료일은 오늘 기준 다음 일자부터 설정하실 수 있습니다.\n오늘 이후의 일자로 입력해주세요.');
		frm.enddate.focus();
		return;
	}
	*/


	if(frm.halfdeliverypay.value=="")
	{
		alert('배송비 부담금액을 입력해 주세요');
		frm.halfdeliverypay.focus();
		return;
	}	

	if(!IsDigit(frm.halfdeliverypay.value)){
		alert("배송비 부담금액은 숫자만 입력 가능합니다.");
		document.frm.halfdeliverypay.focus();
		return;
	}

	if(confirm("수정하시겠습니까?")) {
		frm.submit();
	} else {
		return false;
	}
}

$(function()
{
	//달력대화창 설정
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
		<% if idx<>"" then %>maxDate: "<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
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
		<% if idx<>"" then %>minDate: "<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});

function checkLength(objname, maxlength)
{
	var objstr = objname.value;
	var objstrlen = objstr.length

	var maxlen = maxlength;
	var i = 0;
	var bytesize = 0;
	var strlen = 0;
	var onechar = "";
	var objstr2 = "";

	for (i = 0; i < objstrlen; i++)
	{
		onechar = objstr.charAt(i);

		if (escape(onechar).length > 4)
		{
			bytesize += 2;
		}
		else
		{
			bytesize++;
		}

		if (bytesize <= maxlen)
		{
			strlen = i + 1;
		}
	}

	if (bytesize > maxlen)
	{
		alert("허용된 문자열을 초과하였습니다.\n한글 기준 최대 "+maxlength/2+"자 까지 작성할 수 있습니다.");
		objstr2 = objstr.substr(0, strlen);
		objname.value = objstr2;
	}
	objname.focus();
}
</script>
<%' 팝업 사이즈 : 750*800 %>
<form name="frm" method="post" action="halfdeliverypay_proc.asp">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="adminid" value="<%=loginUserId%>">
<input type="hidden" name="idx" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fidx%>">
<input type="hidden" name="defaultdeliveryType" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultDeliveryType%>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultFreeBeasongLimit%>">
<input type="hidden" name="defaultDeliverPay" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultDeliverPay%>">

	<div class="popWinV17">
		<h1>수정</h1>
		<div class="popContainerV17 pad30">
			<table class="tbType1 writeTb">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>번호(idx) <strong class="cRd1"></strong></div></th>
					<td><%=oHalfDeliveryView.FOneHalfDeliveryPay.Fidx%></td>
				</tr>
				<% If dateModifyCheck Then %>
					<tr>
						<th><div>시작일 <strong class="cRd1">*</strong></div></th>
						<td>
							<input type="text" id="sDt" name="startdate" size="10" readonly value="<%=left(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate,10)%>"/> <input type="hidden" name="starttime" value="<%=starttime%>" /><br/>
							<p class="tPad05 fs11 cGy1">- 수정시 시작일 지정이 종료일 이후로 선택이 안되신다면 종료일을 먼저 입력 후 시작일을 입력해주세요.</p>
						</td>
					</tr>
				<% Else %>
					<tr>
						<th><div>시작일 <strong class="cRd1">*</strong></div></th>
						<td>
							<%=left(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate,10)%>
							<input type="hidden" name="startdate" value="<%=left(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate,10)%>"/> <input type="hidden" name="starttime" value="<%=starttime%>" />
						</td>
					</tr>
				<% End If %>
				<tr>
					<th><div>종료일 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="eDt" name="enddate" size="10" readonly value="<%=left(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate,10)%>" /> <input type="hidden" name="endtime" value="<%=endtime%>" />
					</td>
				</tr>				
				<tr>
					<th><div>배송비 부담금액 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="halfdeliverypay" name="halfdeliverypay" size="10" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.FHalfDeliveryPay%>"/>
						<p class="tPad05 fs11 cGy1">- 콤마없이 숫자만 넣어주세요.</p>
						<p class="tPad05 fs11 cRd1">- 텐배 : 업체가 부담하는 배송비입니다.(정산차감)</p>
						<p class="tPad05 fs11 cRd1">- 업배 : 업체에 지원하는 배송비입니다.(추가정산)</p>						
					</td>
				</tr>
				<tr>
					<th><div>사용여부 <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="N" <% If oHalfDeliveryView.FOneHalfDeliveryPay.Fisusing="N" Then %> checked <% End If %> /> 사용안함</label>
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="Y" <% If oHalfDeliveryView.FOneHalfDeliveryPay.Fisusing="Y" Then %> checked <% End If %> /> 사용함</label>
						</span>
					</td>
				</tr>
				<tr>
					<th><div>등록된 상품</div></th>
					<td>
						<table class="tbType2 writeTb">
							<tr>
								<th align="center" width="12%">이미지</th>
								<th align="center" width="11%">상품코드</th>
								<th align="center" width="15%">브랜드아이디</th>
								<th align="center" width="28%">상품명</th>
								<th align="center" width="13%">조건배송여부</th>
								<th align="center" width="13%">무료배송기준금액</th>
								<th align="center" width="8%">배송비</th>
							</tr>
							<tbody>
								<tr>
									<td width='11%'>
										<img src='<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fsmallimage%>'>
									</td>
									<td width='10%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fitemid%>
									</td>
									<td width='14%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fbrandid%>
									</td>
									<td width='27%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fitemname%>
									</td>
									<td width='12%'>
										<%=getBeadalDivname(oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultDeliveryType)%>
									</td>
									<td width='12%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultFreeBeasongLimit%>
									</td>
									<td width='8%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultDeliverPay%>
									</td>
								</tr>
							</tbody>
						</table>
					</td>
				</tr>									
				<tr>
					<th><div>등록정보</div></th>
					<td>
						<span class="tPad05 col2"><%=oHalfDeliveryView.FOneHalfDeliveryPay.Fadminid%>(<%=fnGetMyname(oHalfDeliveryView.FOneHalfDeliveryPay.Fadminid)%>)<br/><%=oHalfDeliveryView.FOneHalfDeliveryPay.Fregdate%></span>
					</td>
				</tr>
				<% If oHalfDeliveryView.FOneHalfDeliveryPay.Flastadminid <> "" Then %>
				<tr>
					<th><div>최종수정</div></th>
					<td>
						<span class="tPad05 col2 cRd1"><%=oHalfDeliveryView.FOneHalfDeliveryPay.Flastadminid%>(<%=fnGetMyname(oHalfDeliveryView.FOneHalfDeliveryPay.Flastadminid)%>)<br/><%=oHalfDeliveryView.FOneHalfDeliveryPay.Flastupdate%></span>
					</td>
				</tr>
				<% End If %>
				</tbody>
			</table>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
			<input type="button" value="수정" onclick="frmedit();" class="cRd1" style="width:100px; height:30px;" />
		</div>
	</div>
</form>
</body>
</html>
<%
	set oHalfDeliveryView = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
