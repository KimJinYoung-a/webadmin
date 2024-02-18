<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 보너스 쿠폰 적용 제외 상품or브랜드 등록 팝업
' Hieditor : 2021.02.02 원승현 추가
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

function frmExcludingCouponSubmit(){
	var frm  = document.frm;
	var today = new Date();

	if($("#mode").val() == "additem") {
		if (itemid == "") {
			alert('상품을 등록해주세요.');
			return;
		}

		if(confirm("기존에 등록하셨던 상품이 있을경우엔 해당 상품은 등록/수정되지 않습니다.\n사용여부:"+$('input:radio[name=isusing]:checked').val()+"\n해당 정보로 등록하시겠습니까?")) {
			frm.submit();
		} else {
			return false;
		}		
	}

	if($("#mode").val() == "addbrand") {
		if (makerid == "") {
			alert('브랜드를 등록해주세요.');
			return;
		}
		if(confirm("기존에 등록하셨던 브랜드가 있을경우엔 해당 브랜드는 등록/수정되지 않습니다.\n사용여부:"+$('input:radio[name=isusing]:checked').val()+"\n해당 정보로 등록하시겠습니까?")) {
			frm.submit();
		} else {
			return false;
		}
	}
}


function jsAddItemData() {
	document.domain ="10x10.co.kr";
	var winAddItem;
	winAddItem = window.open('/common/pop_singleItemSelect.asp?target=frm&ptype=excludingcoupon','popAddItem','width=1000,height=600');
	winAddItem.focus();
}

function jsAddBrandData() {
	document.domain ="10x10.co.kr";
	var winAddItem;
	winAddItem = window.open('/admin/member/popBrandSearch.asp?frmName=frm&compName=makerid&isjsdomain=o','popAddBrand','width=1000,height=600');
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

function excludingCouponTypeChange(v) {
	if (v=='I') {
		$("#mode").val("additem");
		$("#itemAddArea").show();
		$("#brandAddArea").hide();
	} else if(v=="B") {
		$("#mode").val("addbrand");
		$("#itemAddArea").hide();
		$("#brandAddArea").show();
	}
}

</script>
<%' 팝업 사이즈 : 750*800 %>
<form name="frm" method="post" action="excludingCoupon_proc.asp">
<input type="hidden" name="mode" id="mode" value="additem">
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
					<th><div>구분 <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="excludingCouponType" class="formRadio" value="I" checked onclick="excludingCouponTypeChange(this.value)"/> 상품</label>
							<label class="rMar20"><input type="radio" name="excludingCouponType" class="formRadio" value="B" onclick="excludingCouponTypeChange(this.value)" /> 브랜드</label>
						</span>
					</td>
				</tr>				
				<tr id="itemAddArea">
					<th><div>상품등록 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="itemid" name="itemid" size="10" />
						<input type="button" value="상품검색" onclick="jsAddItemData();" style="width:100px;" />
					</td>
				</tr>				
				<tr id="brandAddArea" style="display:none;">
					<th><div>브랜드 등록 <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="makerid" name="makerid" size="10" />
						<input type="button" value="브랜드 검색" onclick="jsAddBrandData();" style="width:100px;" />
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
				</tbody>
			</table>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
			<input type="button" value="등록" onclick="frmExcludingCouponSubmit();" class="cRd1" style="width:100px; height:30px;" />
		</div>
	</div>
</form>
</body>
</html>
<%
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
