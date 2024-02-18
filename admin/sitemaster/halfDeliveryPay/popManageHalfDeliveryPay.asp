<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 배송비 부담금액 등록 팝업
' Hieditor : 2020.08.27 원승현 추가
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i, loginUserId
loginUserId = session("ssBctId")
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

function frmHalfDeliveryPaySubmit(){
	var frm  = document.frm;
	var today = new Date();

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

	if(frm.enddate.value=="")
	{
		alert('종료일을 입력해 주세요');
		frm.enddate.focus();
		return;
	}

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

	if(typeof frm.iid == "undefined")
	{
		alert('상품을 등록해주세요.');
		return;
	}

	if(confirm("입력하신 시작일,종료일,배송비부담금액,사용여부는\n등록창에서 등록하신 상품에 일괄적용 됩니다.\n기존에 등록하셨던 상품이 있을경우엔 해당 상품은 등록/수정되지 않습니다.\n\n시작일:"+frm.startdate.value+" "+frm.starttime.value+"\n종료일:"+frm.enddate.value+" "+frm.endtime.value+"\n배송비 부담금액:"+frm.halfdeliverypay.value+"원\n사용여부:"+$('input:radio[name=isusing]:checked').val()+"\n해당 정보로 등록하시겠습니까?")) {
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
	var today = new Date();
	$("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
		numberOfMonths: 2,
		showCurrentAtPos: 1,
		showOn: "button",
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
		onClose: function( selectedDate ) {
			$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});

function jsAddItemData() {
	document.domain ="10x10.co.kr";
	var winAddItem;
	winAddItem = window.open('pop_additemlist.asp','popAddItem','width=1000,height=600');
	winAddItem.focus();
}

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

function userAreaDeleteItem(trid) {
	var tr = $(trid).parent().parent();
	tr.remove();
}

function viewUserAddItemListData() {
	var str_array = $("#viewitemdataparent").val().split(',');
	var str_array_detail;
	for(var i = 0; i < str_array.length; i++) {
		str_array_detail = str_array[i].split('|');
		if($("#itemListArea").html().indexOf("viewdata"+str_array_detail[1]) == -1) {
			$("#itemListArea").append("<tr id=viewdata"+str_array_detail[1]+"><td width='11%'><img src='"+str_array_detail[0]+"'></td><td width='10%'>"+str_array_detail[1]+"</td><td width='14%'>"+str_array_detail[2]+"</td><td width='27%'>"+str_array_detail[3]+"</td><td width='12%'>"+str_array_detail[4]+"</td><td width='12%'>"+str_array_detail[5]+"</td><td width='8%'>"+str_array_detail[6]+"</td><td width='10%'><button onclick='userAreaDeleteItem(this);'>삭제</button></td><input type='hidden' name='iid' value='"+str_array_detail[1]+"'>");
		}
	}
}

</script>
<%' 팝업 사이즈 : 750*800 %>
<form name="frm" method="post" action="halfdeliverypay_proc.asp">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="adminid" value="<%=loginUserId%>">
<input type="hidden" name="viewitemdataparent" id="viewitemdataparent">
	<div class="popWinV17">
		<h1>등록</h1>
		<div class="popContainerV17 pad30">
			<table class="tbType1 writeTb">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>시작일 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="sDt" name="startdate" size="10" readonly /> <input type="hidden" name="starttime" value="00:00:00" />
					</td>
				</tr>
				<tr>
					<th><div>종료일 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="eDt" name="enddate" size="10" readonly /> <input type="hidden" name="endtime" value="23:59:59" />
					</td>
				</tr>
				<tr>
					<th><div>배송비 부담금액 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="halfdeliverypay" name="halfdeliverypay" size="10" />
						<p class="tPad05 fs11 cGy1">- 콤마없이 숫자만 넣어주세요.</p>
						<p class="tPad05 fs11 cRd1">- 텐배 : 업체가 부담하는 배송비입니다.(정산차감)</p>
						<p class="tPad05 fs11 cRd1">- 업배 : 업체에 지원하는 배송비입니다.(추가정산)</p>
					</td>
				</tr>
				<tr>
					<th><div>사용여부 <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="N" checked /> 사용안함</label>
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="Y" /> 사용함</label>
						</span>
					</td>
				</tr>
				<tr>
					<th><div>상품등록 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="button" value="상품등록" onclick="jsAddItemData();" style="width:100px; height:30px;" />
						<p>&nbsp;</p>
						<table class="tbType2 writeTb">
							<tr>
								<th align="center" width="12%">이미지</th>
								<th align="center" width="11%">상품코드</th>
								<th align="center" width="15%">브랜드아이디</th>
								<th align="center" width="28%">상품명</th>
								<th align="center" width="13%">조건배송여부</th>
								<th align="center" width="13%">무료배송기준금액</th>
								<th align="center" width="8%">배송비</th>
								<th width="10%"></th>
							</tr>
							<tbody id="itemListArea">
							</tbody>
						</table>
					</td>
				</tr>				
				</tbody>
			</table>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
			<input type="button" value="등록" onclick="frmHalfDeliveryPaySubmit();" class="cRd1" style="width:100px; height:30px;" />
		</div>
	</div>
</form>
</body>
</html>
<%
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
