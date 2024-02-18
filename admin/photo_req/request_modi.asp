<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 촬영 요청 수정 & 뷰 페이지
' History : 2012.03.15 김진영 생성
'			2015.07.28 한용민 수정(수기로 박혀 있는 부분 디비에서 가져옴. 디비구조 변경. 기능 개선&추가. 디자인 신규로 변경)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim gub, gubnm, i, udate, k
Dim cPhotoreq, rno, arrFileList, sMode2, isUpdateDate
Dim PhotoCnt

rno = request("req_no")
gub = request("gub")
sMode2 = request("sMode2")
udate = request("udate")

set cPhotoreq = new Photoreq
	cPhotoreq.FReq_no = rno
	cPhotoreq.fnPhotoreqUpdate
	PhotoCnt = cPhotoreq.fnGetPhotoUser
	arrFileList = cPhotoreq.fnGetFileList

if cPhotoreq.FTotalCount = 0 or cPhotoreq.FTotalCount="" or isnull(cPhotoreq.FTotalCount) then
	Call Alert_move("해당 정보가 없습니다","request_list.asp?menupos="&menupos)
end if

	isUpdateDate = CDate("2016-12-19 18:30:00")
If cPhotoreq.FPhotoreqList(0).FReq_use = "" Then
	Call Alert_move("해당 정보가 없습니다","request_list.asp?menupos="&menupos)
End If

dim copendata
set copendata = new Photoreq
	copendata.FReq_no = rno
	copendata.fnphoto_opendata
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function jsChkSubj(chk){
	if(chk=='5') {
		document.getElementById('detail').style.display = "block";
	} else {
		document.getElementById('detail').style.display = "none";
	}
}
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function fileupload(){
	window.open('request_popupload2.asp','worker','width=420,height=200,scrollbars=yes');
}
function jsDefaultOpt(){
	if ($("#dopt").attr("checked")){
		$("#defaultOptTr").show();
	}else{
		$("#defaultOptTr").hide();
	}
}

// 완성링크 및 참조 URL
function AutoOpenurlInsert() {
	var f = document.all;

	var rowLen = f.divopenurl.rows.length;
//	if(rowLen > 5){
//		alert('더 이상 늘릴 수 없습니다.');
//		return;
//	}
	var r  = f.divopenurl.insertRow(rowLen++);
	var c0 = r.insertCell(0);

	var Html;
	c0.innerHTML = "";

	var inHtml = "<input type='hidden' name='openidx' value=''>"
	inHtml = inHtml + "<input type='text' name='openurl' value='' size=50 maxlength=512>"
	inHtml = inHtml + " <a href='' onclick='dateclearRow(this); return false;'>삭제</a>";
	//alert(inHtml)
	c0.innerHTML = inHtml;

	document.itemreg.lineopenurlCnt.value = f.divopenurl.rows.length;
}

function AutoInsert() {
	// 포토그래퍼 가져오기
	var vreq_photo = $.ajax({
		type: "POST",
		contentType: "application/x-www-form-urlencoded;charset=euc-kr",
		url: "/admin/photo_req/popsearchselect.asp",
		data: "searchtype=req_photo",
		dataType: "text",
		async: false
	}).responseText;

	// 스타일리스트 가져오기
	var vreq_Stylist = $.ajax({
		type: "GET",
		contentType: "application/x-www-form-urlencoded;charset=euc-kr",
		url: "/admin/photo_req/popsearchselect.asp",
		data: "searchtype=req_Stylist",
		dataType: "text",
		async: false
	}).responseText;

	var f = document.all;

	var rowLen = f.div1.rows.length;
	if(rowLen > 5){
		alert('더 이상 늘릴 수 없습니다.');
		return;
	}
	var r  = f.div1.insertRow(rowLen++);
	var c0 = r.insertCell(0);

	var Html;
	c0.innerHTML = "";

	var inHtml = "<input type='hidden' name='tmpcnt'>"
	inHtml = inHtml + "<select class='select' name='yyyy' >"
	inHtml = inHtml + "<option value=<%= year(date()) %> selected><%= year(date()) %></option>"
	<% for i=2002 to Year(now)+1 %>
	inHtml = inHtml + "<option value=<%= CStr(i) %> ><%= CStr(i) %></option>"
	<% next %>
	inHtml = inHtml + "</select>"
	inHtml = inHtml + "<select class='select' name='mm' >"
	inHtml = inHtml + "<option value='<%= month(date()) %>' selected><%= month(date()) %></option>"
	<% for i=1 to 12 %>
	inHtml = inHtml + "<option value='<%= Format00(2,i) %>' ><%= Format00(2,i) %></option>"
	<% next %>
	inHtml = inHtml + "</select>"
	inHtml = inHtml + "<select class='select' name='dd' >"
	inHtml = inHtml + "<option value='<%= day(date()) %>' selected><%= day(date()) %></option>"
	<% for i=1 to 31 %>
	inHtml = inHtml + "<option value='<%= Format00(2,i) %>' ><%= Format00(2,i) %></option>"
	<% next %>
	inHtml = inHtml + "</select> "
	inHtml = inHtml + "<select name='req_day_start' class='select' onchange=document.getElementById('sca1').value=this.value >"
	inHtml = inHtml + "<option value=''>-선택-</option><option value='8'>10:00</option><option value='9'>10:30</option><option value='10'>11:00</option>"
	inHtml = inHtml + "<option value='11'>11:30</option><option value='12'>12:00</option><option value='13'>12:30</option><option value='14'>13:00</option>"
	inHtml = inHtml + "<option value='15'>13:30</option><option value='16'>14:00</option><option value='17'>14:30</option><option value='18'>15:00</option>"
	inHtml = inHtml + "<option value='19'>15:30</option><option value='20'>16:00</option><option value='21'>16:30</option><option value='22'>17:00</option>"
	inHtml = inHtml + "<option value='23'>17:30</option><option value='24'>18:00</option>"
	inHtml = inHtml + "</select> ~ "
	inHtml = inHtml + "<select name='req_day_end' class='select' onchange=document.getElementById('sca2').value=this.value>"
	inHtml = inHtml + "<option value=''>-선택-</option><option value='8'>10:00</option><option value='9'>10:30</option><option value='10'>11:00</option>"
	inHtml = inHtml + "<option value='11'>11:30</option><option value='12'>12:00</option><option value='13'>12:30</option><option value='14'>13:00</option>"
	inHtml = inHtml + "<option value='15'>13:30</option><option value='16'>14:00</option><option value='17'>14:30</option><option value='18'>15:00</option>"
	inHtml = inHtml + "<option value='19'>15:30</option><option value='20'>16:00</option><option value='21'>16:30</option><option value='22'>17:00</option>"
	inHtml = inHtml + "<option value='23'>17:30</option><option value='24'>18:00</option>"
	inHtml = inHtml + "</select>"
	inHtml = inHtml + " <input type='text' name='comment' size=25>"
	inHtml = inHtml + "&nbsp;&nbsp;" + vreq_photo + "&nbsp;&nbsp;" + vreq_Stylist
	inHtml = inHtml + " <a href='' onclick='dateclearRow(this); return false;'>삭제</a>";
	//alert(inHtml)
	c0.innerHTML = inHtml;

	document.itemreg.lineCnt.value = f.div1.rows.length;
}
function goURL(){
	var uu = document.getElementById('lurl').value;
	window.open(uu);
}
function pop_print(){
	window.open('request_print.asp?req_no=<%=rno%>');
}
function onlyNumber(){
	if((event.keyCode<48)||(event.keyCode>57))
		event.returnValue=false;
}
function form_check(){
	var frm = document.itemreg;

	if(frm.req_use.value == 0) {
		alert("촬영용도 구분을 선택하세요");
		frm.req_use.focus();
		return false;
	}

	if(frm.req_use.value == 1) {
		if(frm.req_use_detail.value == 0) {
			alert("기본 상세페이지를 선택하세요");
			frm.req_use_detail.focus();
			return false;
		}
	}
	if (frm.prd_name.value == ""){
		alert("상품명을 입력하세요.");
		frm.prd_name.focus();
		return;
	}
//	if (frm.prd_type.value == ""){
//		alert("상품군을 입력하세요.");
//		frm.prd_type.focus();
//		return;
//	}
	if (frm.prd_type2.value == ""){
		alert("총 상품수량을 입력하세요.");
		frm.prd_type2.focus();
		return;
	}
	if (frm.import_level.value == "0"){
		alert("중요도를 입력하세요.");
		frm.import_level.focus();
		return;
	}
	<% if not(session("ssAdminLsn")<="3") then %>
		if (frm.import_level.value == "4" || frm.import_level.value == "5"){
			alert('중요도 S와 A는 팀장,파트장,선임만 선택 가능 합니다.');
			return;
		}
	<% end if %>
	if(frm.req_department.value == "") {
		alert("요청부서를 선택하세요");
		frm.req_department.focus();
		return false;
	}
	if(frm.req_cdl_disp.value == "") {
		alert("카테고리를 선택하세요");
		frm.req_cdl_disp.focus();
		return false;
	}
//	if(frm.MDid.value == "00") {
//		alert("담당MD를 선택하세요");
//		frm.MDid.focus();
//		return false;
//	}

	if (frm.req_etc1.value == ""){
		alert("상품 특징 및 주요 전달 사항을 입력하세요.");
		frm.req_etc1.focus();
		return;
	}
	if(parseInt(document.getElementById('sca1').value) >= parseInt(document.getElementById('sca2').value)){
		alert("스케쥴 시간이 잘 못 되었습니다.");
		document.getElementById('sca1').value = "1";
		document.getElementById('sca2').value = "2";
		return false;
	}
	if(frm.itemid.value!=''){
		if (!IsNumbers(frm.itemid.value)){
			alert('상품코드를 정확하게 입력해 주세요.');
			frm.itemid.focus();
			return;
		}
	}
	<% If sMode2 <> "I" and (C_ADMIN_AUTH or C_CONTENTS_part) Then %>
		if (frm.tmpcnt != undefined){
			if (frm.tmpcnt.length>1){
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.yyyy[i].value == "") {
						alert('촬영일시 년도를 입력해 주세요.');
						frm.yyyy[i].focus();
						return false;
					}
				}
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.mm[i].value == "") {
						alert('촬영일시 월을 입력해 주세요.');
						frm.mm[i].focus();
						return false;
					}
				}
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.dd[i].value == "") {
						alert('촬영일시 일을 입력해 주세요.');
						frm.dd[i].focus();
						return false;
					}
				}
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.req_day_start[i].value == "") {
						alert('촬영일시 시작 시간을 입력해 주세요.');
						frm.req_day_start[i].focus();
						return false;
					}
				}
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.req_day_end[i].value == "") {
						alert('촬영일시 종료 시간을 입력해 주세요.');
						frm.req_day_end[i].focus();
						return false;
					}
				}
			}else{
				if(frm.yyyy.value == "") {
					alert('촬영일시 년도를 입력해 주세요.');
					frm.yyyy.focus();
					return false;
				}
				if(frm.mm.value == "") {
					alert('촬영일시 월을 입력해 주세요.');
					frm.mm.focus();
					return false;
				}
				if(frm.dd.value == "") {
					alert('촬영일시 일을 입력해 주세요.');
					frm.dd.focus();
					return false;
				}
				if(frm.req_day_start.value == "") {
					alert('촬영일시 시작 시간을 입력해 주세요.');
					frm.req_day_start.focus();
					return false;
				}
				if(frm.req_day_end.value == "") {
					alert('촬영일시 종료 시간을 입력해 주세요.');
					frm.req_day_end.focus();
					return false;
				}
			}
		}
	<% end if %>

	//frm.lineCnt.value = document.all.div1.rows.length;
	frm.action = "/admin/photo_req/request_proc.asp";
	frm.submit();
}
function filedownload(idx){
	filefrm.file_idx.value = idx;
	filefrm.submit();
}
function clearRow(tdObj) {
	if(confirm("선택하신 파일을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}

function viewSche(vdate){
	window.open('request_cal_day.asp?getday='+vdate);
}

function dateclearRow(tdObj) {
	if ( itemreg.lineCnt.value < 2 ){
		alert('최소 1개의 스케줄은 입력하셔야 합니다.');
		return;
	}

	if(confirm("선택하신 줄을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
		var f = document.all;
		document.itemreg.lineCnt.value = f.div1.rows.length;
	} else {
		return false;
	}
}

function popdepartment(){
	var popwin = window.open('popdepartmentselect.asp','addreg','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function downphotoitemlist_sample(){
	var popwin = window.open('http://imgstatic.10x10.co.kr/offshop/sample/photo/photoitemlist_sample.xlsx','exceldown','width=100,height=100,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 새상품 추가 팝업
function addnewItem(){
	var popwin;
	popwin = window.open("/admin/photo_req/pop_itemAddInfo.asp?smode=E", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>

<!-- 표 상단바 시작-->
<form name="itemreg" method="post">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="mode2" value="<%=sMode2%>">
<input type="hidden" name="userFont" value="<%=cPhotoreq.FPhotoreqList(0).FFontColor%>">
<input type="hidden" name="req_no" value="<%=rno%>">
<input type="hidden" id="sca1" name = "sca1" value = "1">
<input type="hidden" id="sca2" name = "sc12" value = "2">
<input type="hidden" name="req_id" value="<%=cPhotoreq.FPhotoreqList(0).FReq_id%>">
<input type="hidden" name="lineCnt" value="<%= cPhotoreq.FResultcount %>">

<!-- 1.일반정보 등록 폼 시작-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td bgcolor="#FFFFFF" height="30" colspan="5">1.일반정보</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">촬영 구분 *</td>
	<td bgcolor="#FFFFFF">
	<%
		If gub <> "" Then
			Select Case gub
				Case "2"	gubnm = "추가촬영"
				Case "3"	gubnm = "재촬영"
				Case "4"	gubnm = "추가 기입 요청건"
			End Select
			response.write gubnm
			response.write "<input type='hidden' name='req_gubunS' value='"&gub&"'>"
			response.write "<input type='hidden' name='req_gubun' value='"&gubnm&"'>"
		Else
			response.write cPhotoreq.FPhotoreqList(0).FReq_gubun
			response.write "<input type='hidden' name='req_gubun' value='"&cPhotoreq.FPhotoreqList(0).FReq_gubun&"'>"
		End If
	%>
	</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<%If sMode2 <> "I" Then %>
			<%= chkIIF(cPhotoreq.FPhotoreqList(0).FLoad_req <> "","불러온 요청서 No : "&cPhotoreq.FPhotoreqList(0).FLoad_req&"","") %>
		<%End If%>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="4">
		<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0" id="lyRequre" style="display:none;">
		<tr>
			<td align="left">
				촬영요청 no. <input type="text" name="request_no" value="" size="10" class="text">
				<input type="button" value="확인" class="button" onclick="request_modi();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">촬영용도 구분 *</td>
	<td bgcolor="#FFFFFF" ><% call DrawPicGubun2("req_use", "doc_status", cPhotoreq.FPhotoreqList(0).FReq_use) %></td>
	<td bgcolor="#FFFFFF" colspan="3">
		<div id = "detail" bgcolor="#FFFFFF" <% If cPhotoreq.FPhotoreqList(0).FReq_use_detail = "" Then response.write "style=display:none;" End If %>><% call DrawPicGubun2("req_use_detail", "doc_status_detail" ,cPhotoreq.FPhotoreqList(0).FReq_use_detail) %></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품명(기획전명) *</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_name" value="<%=cPhotoreq.FPhotoreqList(0).FPrd_name%>" size="64" maxlength="30" class="text"></td>
</tr>
<input type="hidden" name="prd_type" value="<%=cPhotoreq.FPhotoreqList(0).FPrd_type%>" size="60" maxlength="128" class="text">
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품군 *</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_type" value="<%'=cPhotoreq.FPhotoreqList(0).FPrd_type%>" size="60" maxlength="128" class="text"></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">총 상품수량 *</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="text" style="IME-MODE:disabled;" name="prd_type2" value="<%=cPhotoreq.FPhotoreqList(0).FPrd_type2%>" size="10" maxlength="10" class="text" onkeypress="onlyNumber();">&nbsp;(15종 이상 시 스케줄 문의 필수)
		<!--<select name="prd_type2">
			<% for i = 1 to 10 %> 
			<option value="<%= i %>"><%= i %></option>
			<% next %>
		</select>
		(수량이 10개 이상일 경우 요청서를 나눠서 올려주세요)-->
	</td>
</tr>
<input type="hidden" name="prd_price" value="<%=cPhotoreq.FPhotoreqList(0).FPrd_price%>" size="10" maxlength="128" class="text">
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매가(소비자가)</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_price" value="<%'=cPhotoreq.FPhotoreqList(0).FPrd_price%>" size="10" maxlength="128" class="text"></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">중요도 </td>
	<td bgcolor="#FFFFFF" colspan="4">
		<select name="import_level" class="select">
			<option value="0">--중요도 선택--</option>

			<% if session("ssAdminLsn")<="3" then %>
				<option value="5" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="5","selected","") %>>S</option>
				<option value="4" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="4","selected","") %>>A</option>
			<% end if %>

			<option value="3" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="3","selected","") %>>B</option>
			<option value="2" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="2","selected","") %>>C</option>
			<option value="1" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="1","selected","") %>>D</option>
		</select>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">요청부서/카테고리 *</td>
	<td bgcolor="#FFFFFF" >
		<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1">
		<tr>
			<td align="left">
				<input type="hidden" name="req_department" value="<%= cPhotoreq.FPhotoreqList(0).FReq_department %>">
				<input type="hidden" name="MDid" value="<%= cPhotoreq.FPhotoreqList(0).FMDid %>">
				<div name="divdepartmentname" id="divdepartmentname"><%= getDepartmentALL(cPhotoreq.FPhotoreqList(0).FReq_department) %></div>
				<div name="divMDidname" id="divMDidname"><%= cPhotoreq.FPhotoreqList(0).FMDname %></div>
				<input type="button" onclick="popdepartment();" value="부서검색" class="button" >
				<!--<select name="req_department" class="select">
					<option value="">--부서 선택--</option>
					<option value="MD" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="MD","selected","") %>>MD</option>
					<option value="MKT" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="MKT","selected","") %>>MKT</option>
					<option value="ithinkso" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="ithinkso","selected","") %>>ithinkso</option>
					<option value="off" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="off","selected","") %>>off</option>
					<option value="JR" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="JR","selected","") %>>전략기획</option>
					<option value="WD" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="WD","selected","") %>>WD</option>
					<option value="CT" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="CT","selected","") %>>컨텐츠</option>
				</select>-->
			</td>
		</tr>
		</table>
	</td>
	<td bgcolor="#FFFFFF">
		<!--관리카테 : <% call DrawCategoryLarge("req_category", cPhotoreq.FPhotoreqList(0).FReq_category) %><br>-->
		전시카테고리 : <%= fnStandardDispCateSelectBox(1,cPhotoreq.FPhotoreqList(0).freq_cdl_disp, "req_cdl_disp", cPhotoreq.FPhotoreqList(0).freq_cdl_disp, "")%>
		<% 'call DrawCategoryLarge_disp("req_cdl_disp", cPhotoreq.FPhotoreqList(0).freq_cdl_disp) %>
	</td>
	<td bgcolor="#FFFFFF" colspan=2>
		촬영요청자 : <%= cPhotoreq.FPhotoreqList(0).FReq_name %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">브랜드ID :</td>
	<td bgcolor="#FFFFFF" colspan="4"><%	drawSelectBoxDesignerWithName "makerid", cPhotoreq.FPhotoreqList(0).FMakerid %></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드<br>(쉼표로 복수입력가능)</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="text" name="itemid" size="60" maxlength="128" class="text" value="<%=cPhotoreq.FPhotoreqList(0).FItemid%>" >
		<input type="button" value="상품추가" onclick="addnewItem();" class="button">
<%
if cPhotoreq.FPhotoreqList(0).FItemid <> "" then
dim oitem
set oitem = new CItem
oitem.FPageSize = 1000
oitem.FCurrPage = 1
oitem.FRectItemid = cPhotoreq.FPhotoreqList(0).FItemid
oitem.GetItemList
%>
<% if oitem.FresultCount > 0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">상품ID</td>
	<td align="center">이미지</td>
	<td align="center">브랜드</td>
	<td align="center">상품명</td>
</tr>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
	<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
</tr>
<% next %>
</table>
<% end if %>
<%
set oitem = nothing
end if
%>
	</td>
</tr>
<!-- 업체등록시에는 사용안함(MD만 등록가능) -->
<input type="hidden" name="req_date" size="10" value="<%=left(cPhotoreq.FPhotoreqList(0).FReq_date,10)%>" >
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">희망 촬영일자 </td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="req_date" size="10" value="<%'=left(cPhotoreq.FPhotoreqList(0).FReq_date,10)%>" onClick="jsPopCal('req_date');" style="cursor:hand;"></td>
</tr>-->
<tr>
	<td height="30" width="15%" bgcolor="#DDDDFF">
		첨부파일등록
	</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="button" onclick="downphotoitemlist_sample();" value="촬영요청상품리스트 샘플 다운로드" class="button">
		<br><br><input type="button" value="파일업로드" class="button" onclick="fileupload();">
		(최대20mb까지 업로드 가능하며, 문서에 주민번호 혹은 전화번호같은 개인정보가 들어갈 경우 방화벽에서 막히니 주의)
		<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
<%
	IF isArray(arrFileList) THEN
		For i =0 To UBound(arrFileList,2)
%>
		<tr>
			<td>
				<input type='hidden' name='doc_file' value='<%=arrFileList(1,i)%>'>
				<input type='hidden' name='doc_realfile' value='<%=arrFileList(3,i)%>'>
				<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
				<span class="a" onClick="filedownload(<%=arrFileList(0,i)%>)" style="cursor:pointer"><%=Split(Replace(arrFileList(1,i),"http://",""),"/")(4)%></span>
			</td>
		</tr>
<%
		Next
		Response.Write "<input type='hidden' name='isfile' value='o'>"
	Else
		Response.write "<tr><td></td></tr>"
	End If
%>
		</table>
	</td>
</tr>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td bgcolor="#FFFFFF" height="30" colspan="2">2.촬영컨셉(중복선택 가능)</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">필요 촬영군</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>
				<% call CheckBoxUseType("doc_use_type", rno, "1") %>
			</td>
		<%
			Dim odefault, isOptExist, defaultoptArr, dOptdispYn
			set odefault = new Photoreq
				isOptExist = odefault.getdefaultOpt(rno)

				IF isArray(isOptExist) THEN
					dOptdispYn = "Y"
				Else
					dOptdispYn = "N"
				End If

				IF dOptdispYn = "Y" THEN
					For k =0 To UBound(isOptExist,2)
						defaultoptArr = defaultoptArr & isOptExist(1,k) & "," 
					Next
				Else
					defaultoptArr = ""
				End If
		%>
			<tr id="defaultOptTr" <%= Chkiif(dOptdispYn="Y", "style='display:block;'", "style='display:none;'") %>style="display:block;">
				<td>
					<input type="checkbox" name="defaultOpt" value="901" <%= Chkiif(instr(defaultoptArr, "901") > 0, "checked", "") %>>정면
					<input type="checkbox" name="defaultOpt" value="902" <%= Chkiif(instr(defaultoptArr, "902") > 0, "checked", "") %>>후면
					<input type="checkbox" name="defaultOpt" value="903" <%= Chkiif(instr(defaultoptArr, "903") > 0, "checked", "") %>>측면
					<input type="checkbox" name="defaultOpt" value="904" <%= Chkiif(instr(defaultoptArr, "904") > 0, "checked", "") %>>단체컷
					<input type="checkbox" name="defaultOpt" value="905" <%= Chkiif(instr(defaultoptArr, "905") > 0, "checked", "") %>>패키지컷
				</td>
			</tr>
		<%
				SET odefault = nothing
		%>
		</tr>
		</table>
	</td>
</tr>
<% If isUpdateDate >= cPhotoreq.FPhotoreqList(0).FReq_regdate Then %>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">메인 촬영 컨셉</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr><% call CheckBoxUseType("doc_use_concept", rno, "2") %></tr>
		</table>
	</td>
</tr>
<% End If %>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 특징 및 주요 전달 사항 *</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="req_etc1" rows="18" class="textarea" style="width:100%"><%=cPhotoreq.FPhotoreqList(0).FReq_etc1%></textarea></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">관련 링크 및 참조 URL </td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="req_url" size="60" maxlength="200" class="text" value="<%=cPhotoreq.FPhotoreqList(0).FReq_url%>" id="lurl">
		<!--<input type="button" class="button" value="바로가기" onclick="goURL();">-->
	</td>
</tr>
<input type="hidden" name="req_etc2" size="10" value="<%=cPhotoreq.FPhotoreqList(0).FReq_etc2%>" >
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">촬영 시 유의사항</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="req_etc2" rows="5" class="textarea" style="width:100%"><%=cPhotoreq.FPhotoreqList(0).FReq_etc2%></textarea></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">게시글 사용여부</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="radio" name="use_yn" value="Y" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FUse_yn="Y","checked","") %>>Y
		<input type="radio" name="use_yn" value="N" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FUse_yn="N","checked","") %>>N
	</td>
</tr>
</table>

<% If sMode2 <> "I" and (C_ADMIN_AUTH or C_CONTENTS_part) Then %>
	<br>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left">
		<td bgcolor="#FFFFFF" height="30" colspan="4">3. 진행상황 <input type="button" value="촬영 스케줄 보기" class="button" onclick="viewSche('<%=Left(now(),10)%>');"></td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">상태</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select name="req_status" class="select">
				<option value="0">--진행상태선택--</option>
				<option value="4" <%= chkIIF(gub = "4","selected","") %> <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "4","selected","") %> >추가기입 요청</option>
				<option value="1" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "1","selected","") %> >촬영스케줄 지정</option>
				<option value="2" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "2","selected","") %> >촬영중</option>
				<option value="3" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "3","selected","") %> >촬영완료</option>
				<option value="9" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "9","selected","") %> >최종오픈</option>
			</select>
			<input type="checkbox" name="fontColor" value="R" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FFontColor = "R","checked","") %>>진행상태 텍스트 색상 바뀜(요청자에게 추가기입 요청 알림시에만 체크하세요!)
		</td>
	</tr>

	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">촬영 확정 일시</td>
		<td bgcolor="#FFFFFF" width="70%">
			<div id="lyRequre1">
				<table id="div1" class="a">
				<%
				dim vstart_date, vend_date

				For i = 0 to cPhotoreq.FResultcount -1
				
				if cPhotoreq.FPhotoreqList(i).FStart_date="" or isnull(cPhotoreq.FPhotoreqList(i).FStart_date) then
					vstart_date=date()
				else
					vstart_date=cPhotoreq.FPhotoreqList(i).FStart_date
				end if
				%>
					<tr>
						<td>
							<input type="hidden" name="tmpcnt">
							<% DrawOneDateBoxdynamic "yyyy", year(vstart_date), "mm", month(vstart_date), "dd", day(vstart_date), "", "", "", "" %>
							<select name="req_day_start" class="select" onchange="document.getElementById('sca1').value=this.value">
								<option value="">-선택-</option>
								<option value="8" <% if cstr("10:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>10:00</option>
								<option value="9" <% if cstr("10:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>10:30</option>
								<option value="10" <% if cstr("11:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>11:00</option>
								<option value="11" <% if cstr("11:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>11:30</option>
								<option value="12" <% if cstr("12:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>12:00</option>
								<option value="13" <% if cstr("12:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>12:30</option>
								<option value="14" <% if cstr("13:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>13:00</option>
								<option value="15" <% if cstr("13:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>13:30</option>
								<option value="16" <% if cstr("14:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>14:00</option>
								<option value="17" <% if cstr("14:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>14:30</option>
								<option value="18" <% if cstr("15:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>15:00</option>
								<option value="19" <% if cstr("15:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>15:30</option>
								<option value="20" <% if cstr("16:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>16:00</option>
								<option value="21" <% if cstr("16:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>16:30</option>
								<option value="22" <% if cstr("17:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>17:00</option>
								<option value="23" <% if cstr("17:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>17:30</option>
								<option value="24" <% if cstr("18:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>18:00</option>
							</select>
							~
							<select name="req_day_end" class="select" onchange="document.getElementById('sca2').value=this.value">
								<option value="">-선택-</option>
								<option value="8" <% if cstr("10:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>10:00</option>
								<option value="9" <% if cstr("10:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>10:30</option>
								<option value="10" <% if cstr("11:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>11:00</option>
								<option value="11" <% if cstr("11:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>11:30</option>
								<option value="12" <% if cstr("12:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>12:00</option>
								<option value="13" <% if cstr("12:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>12:30</option>
								<option value="14" <% if cstr("13:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>13:00</option>
								<option value="15" <% if cstr("13:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>13:30</option>
								<option value="16" <% if cstr("14:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>14:00</option>
								<option value="17" <% if cstr("14:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>14:30</option>
								<option value="18" <% if cstr("15:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>15:00</option>
								<option value="19" <% if cstr("15:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>15:30</option>
								<option value="20" <% if cstr("16:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>16:00</option>
								<option value="21" <% if cstr("16:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>16:30</option>
								<option value="22" <% if cstr("17:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>17:00</option>
								<option value="23" <% if cstr("17:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>17:30</option>
								<option value="24" <% if cstr("18:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>18:00</option>
							</select>
							<input type='text' name='comment' value='<%= cPhotoreq.FPhotoreqList(i).fcomment %>' size=25>
							&nbsp;<% call SelectUser("1", "req_photo", ""&cPhotoreq.FPhotoreqList(i).FReq_photo&"") %>
							&nbsp;<% call SelectUser("2", "req_Stylist", ""&cPhotoreq.FPhotoreqList(i).FReq_stylist&"") %>
							<a href='' onclick='dateclearRow(this); return false;'>삭제</a>
						</td>
					</tr>
				<%
				Next
				%>
				</table>
			</div>
		</td>
		<td bgcolor="#FFFFFF"><input type="button" value="촬영일시추가" onClick="AutoInsert()" class="button"></td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">코멘트 작성</td>
		<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="req_comment" size="60" maxlength="128" class="text" value="<%=cPhotoreq.FPhotoreqList(0).FReq_comment%>"></td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">SMS 전송</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="checkbox" name="req_SMS" value="Y">담당MD에게 SMS전송 (MD지정 안 된 경우, 촬영요청자에게 전송)
		</td>
	</tr>
	</table>

<% else %>
	<br>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left">
		<td bgcolor="#FFFFFF" height="30" colspan="4">3. 진행상황</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">상태</td>
		<td bgcolor="#FFFFFF" colspan="3">
	<%
		Select Case cPhotoreq.FPhotoreqList(0).FReq_status
			Case "1"	response.write "촬영스케줄 지정"
			Case "2"	response.write "촬영중"
			Case "3"	response.write "촬영완료"
			Case "4"	response.write "추가기입 요청"
			Case "10"	response.write "최종오픈"
		End Select
	%>
		</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">촬영 확정 일시</td>
		<td bgcolor="#FFFFFF" colspan=3>
				<table class="a">
	<%
		For i = 0 to cPhotoreq.FResultcount -1
	%>
				<tr>
					<td>
						<font color="BLUE">시작 : <%=cPhotoreq.FPhotoreqList(i).FStart_date%></font> ~  <font color="RED">종료 : <%=cPhotoreq.FPhotoreqList(i).FEnd_date%></font>
						&nbsp;포토 : <% call SelectUser2("1", ""&cPhotoreq.FPhotoreqList(i).FReq_photo&"") %>
						&nbsp;스타일 : <% call SelectUser2("2", ""&cPhotoreq.FPhotoreqList(i).FReq_stylist&"") %>
					</td>
				</tr>
	<%
		Next
	%>
				</table>
		</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">코멘트 작성</td>
		<td bgcolor="#FFFFFF" colspan="3"><%=cPhotoreq.FPhotoreqList(0).FReq_comment%></td>
	</tr>
	</table>
<% End If %>

<%
' 상태값이 촬영완료 이거나 최종오픈 일경우
if cPhotoreq.FPhotoreqList(0).FReq_status="3" or cPhotoreq.FPhotoreqList(0).FReq_status="9" then
%>
	<Br>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left">
		<td bgcolor="#FFFFFF" height="30" colspan="3">4. 최종오픈정보</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">완성링크 및 참조 URL</td>
		<td bgcolor="#FFFFFF">
			<input type="hidden" name="lineopenurlCnt" value="0">
			<div id="lyopenurl">
				<table id="divopenurl" class="a">
				<%
				For i = 0 to copendata.FResultcount -1
				%>
					<tr>
						<td>
							<input type='hidden' name='openidx' value='<%= copendata.FPhotoreqList(i).fopenidx %>'>
							<input type='text' name='openurl' value='<%= copendata.FPhotoreqList(i).fopenurl %>' size=50 maxlength=512>
							<a href='' onclick='dateclearRow(this); return false;'>삭제</a>
						</td>
					</tr>
				<%
				Next
				%>
				</table>
			</div>
		</td>
		<td bgcolor="#FFFFFF"><input type="button" value="완성링크추가" onClick="AutoOpenurlInsert()" class="button"></td>
	</tr>
	</table>
<% end if %>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td bgcolor="#FFFFFF" height="30">
		<input type="button" class="button" value=" 저 장 " onclick="form_check();">
		<input type="button" class="button" value=" 취 소 " onClick="window.location='/admin/photo_req/request_list.asp?menupos=<%=menupos%>'">
		<input type="button" class="button" value=" 인 쇄 " onClick="pop_print();">		
	</td>
</tr>
</table>

</form>
<form name="filefrm" method="post" action="<%=uploadImgUrl%>/linkweb/photo_req/photo_req_download2.asp" target="fileiframe">
<input type="hidden" name="brd_sn" value="<%=rno%>">
<input type="hidden" name="file_idx" value="">
</form>
<%
set cPhotoreq=nothing
set copendata = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->