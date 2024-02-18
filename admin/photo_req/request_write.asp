<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 촬영 요청 등록페이지
' History : 2012.03.13 김진영 생성
'			2015.07.28 한용민 수정(수기로 박혀 있는 부분 디비에서 가져옴. 디비구조 변경. 기능 개선&추가. 디자인 신규로 변경
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
<%
Dim i, cdl, makerid, arrFileList, MDid, cdl_disp

Dim cPhotoreq, isUpdateDate
set cPhotoreq = new Photoreq
	cPhotoreq.fnReqno

	isUpdateDate = CDate("2016-12-19")
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>

function request_modi(){
	var qqq;
	var frm = document.itemreg;
	var chk = 0;

	if (frm.request_no.value==''){
		alert('촬영요청 번호를 입력하세요.');
		frm.request_no.focus();
		return;
	}

	for(var i=0; i<frm.req_gubun.length; i++) {
		if(frm.req_gubun[i].checked){
			qqq = frm.req_gubun[i].id;
			chk++;
		}
	}
	if(confirm("저장된 내용을 불러오시겠습니까?") == true) {
		location.href= 'request_modi.asp?req_no='+document.getElementById('request_no').value+'&gub='+qqq+'&sMode2=I&menupos=<%= menupos %>';
	} else {
		return false;
	}
}
function jsChkSubj(chk){

	if(chk=='5') {
		document.getElementById('detail').style.display = "block";
	} else {
		document.getElementById('detail').style.display = "none";
	}
}

function jsDefaultOpt(){
	if ($("#dopt").attr("checked")){
		$("#defaultOptTr").show();
	}else{
		$("#defaultOptTr").hide();
	}
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function fileupload()
{
	window.open('request_popupload2.asp','worker','width=420,height=200,scrollbars=yes');
}
function onlyNumber(){
	if((event.keyCode<48)||(event.keyCode>57))
		event.returnValue=false;
}
function form_check(){
	var frm = document.itemreg;
	var chk = 0;
	for(var i=0; i<frm.req_gubun.length; i++) {
		if(frm.req_gubun[i].checked) chk++;
	}
	if(chk == "0"){
		alert("촬영 구분을 선택하세요");
		return false;
	}

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
	if(frm.prd_type2.value!=''){
		if (!IsDouble(frm.prd_type2.value)){
			alert('총 상품수량은 숫자만 가능합니다.');
			frm.prd_type2.focus();
			return;
		}
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

	if(frm.itemid.value!=''){
		if (!IsNumbers(frm.itemid.value)){
			alert('상품코드를 정확하게 입력해 주세요.');
			frm.itemid.focus();
			return;
		}
	}

	frm.action = "request_proc.asp";
	frm.submit();
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
	popwin = window.open("/admin/photo_req/pop_itemAddInfo.asp", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>

<form name="itemreg" method="post">
<input type="hidden" name="mode" value="I">
<input type = "hidden" name = "req_no" value = "<%=cPhotoreq.Freq_no + 1%>">

<!-- 1.일반정보 등록 폼 시작-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td bgcolor="#FFFFFF" height="30" colspan="5">1.일반정보</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">촬영 구분 *</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="req_gubun" id="1" value="신규" onClick="document.getElementById('lyRequre').style.display='none';">신규</label>
		&nbsp;<label><input type="radio" name="req_gubun" id="2" value="추가촬영" onClick="document.getElementById('lyRequre').style.display='block';">추가촬영</label>
		&nbsp;<label><input type="radio" name="req_gubun" id="3" value="재촬영" onClick="document.getElementById('lyRequre').style.display='block';">재촬영</label>
		&nbsp;<label><input type="radio" name="req_gubun" id="4" value="추가 기입 요청건" onClick="document.getElementById('lyRequre').style.display='block';">추가 기입 요청건</label>
	</td>
	<td bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="4">
		<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0" id="lyRequre" style="display:none;">
		<tr>
			<td align="left">
				촬영요청 no. <input type="text" name="request_no" value="" id="request_no" size="10" class="text">
				<input type="button" value="확인" class="button" onclick="request_modi();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">촬영용도 구분 *</td>
	<td bgcolor="#FFFFFF" ><% call DrawPicGubun("req_use", "doc_status", "1") %></td>
	<td bgcolor="#FFFFFF" colspan="3">
		<div id = "detail" bgcolor="#FFFFFF" style="display:none;" ><% call DrawPicGubun("req_use_detail", "doc_status_detail", "1") %></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품명(기획전명) *</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_name" size="60" maxlength="30" class="text"></td>
</tr>
<input type="hidden" name="prd_type" size="60" maxlength="128" class="text">
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품군 *</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_type" size="60" maxlength="128" class="text"></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">총 상품수량 *</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="text" style="IME-MODE:disabled;" name="prd_type2" size="10" maxlength="10" class="text" onkeypress="onlyNumber();">&nbsp;(15종 이상 시 스케줄 문의 필수)
		<!--<select name="prd_type2">
			<% for i = 1 to 10 %> 
			<option value="<%= i %>"><%= i %></option>
			<% next %>
		</select>
		(수량이 10개 이상일 경우 요청서를 나눠서 올려주세요)-->
	</td>
</tr>
<input type="hidden" name="prd_price" size="10" maxlength="128" class="text">
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매가(소비자가)</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_price" size="10" maxlength="128" class="text"></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">중요도 </td>
	<td bgcolor="#FFFFFF" colspan="4">
		<select name="import_level" class="select">
			<option value="0">--중요도 선택--</option>

			<% if session("ssAdminLsn")<="3" then %>
				<option value="5">S</option>
				<option value="4">A</option>
			<% end if %>

			<option value="3">B</option>
			<option value="2">C</option>
			<option value="1">D</option>
		</select>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">요청부서/카테고리 *</td>
	<td bgcolor="#FFFFFF" >
		<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1">
		<tr>
			<td align="left">
				<input type="hidden" name="req_department" value="">
				<input type="hidden" name="MDid" value="">
				<div name="divdepartmentname" id="divdepartmentname"></div>
				<div name="divMDidname" id="divMDidname"></div>
				<input type="button" onclick="popdepartment();" value="부서검색" class="button" >
				<!--<select name="req_department" class="select">
					<option value="">--부서 선택--</option>
					<option value="MD">MD</option>
					<option value="MKT">MKT</option>
					<option value="ithinkso">ithinkso</option>
					<option value="off">off</option>
					<option value="JR">전략기획</option>
					<option value="WD">WD</option>
					<option value="CT">컨텐츠</option>
				</select>-->
			</td>
		</tr>
		</table>
	</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<!--관리카테 : <% call DrawCategoryLarge("req_category", cdl) %><br>-->
		전시카테고리 : <%= fnStandardDispCateSelectBox(1,cdl_disp, "req_cdl_disp", cdl_disp, "")%>
		<% 'call DrawCategoryLarge_disp("req_cdl_disp", cdl_disp) %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">브랜드ID :</td>
	<td bgcolor="#FFFFFF" colspan="4"><%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드<br>(쉼표로 복수입력가능)</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="text" name="itemid" size="60" maxlength="128" class="text">
		<input type="button" value="상품추가" onclick="addnewItem();" class="button">
	</td>
</tr>
<!-- 업체등록시에는 사용안함(MD만 등록가능) -->
<input type="hidden" name="req_date" size="10" >
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">희망 촬영일자 </td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="req_date" size="10" onClick="jsPopCal('req_date');" style="cursor:hand;"></td>
</tr>-->
<tr>
	<td height="30" width="15%" bgcolor="#DDDDFF">첨부파일등록</td>
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
				<input type='hidden' name='doc_file' value='<%=arrFileList(0,i)%>'>
				<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
				<a href='<%=arrFileList(0,i)%>' target='_blank'>
				<%=Split(Replace(arrFileList(0,i),"http://",""),"/")(4)%></a>
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
<Br>
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
				<% call CheckBoxUseType("doc_use_type", "", "") %>
				<!-- <br />기본컷인 경우 정면, 후면, 측면, 단체컷, 패키지컷 중에 필요하신 항목을 기입해주세요. -->
			</td>
		</tr>
		<tr id="defaultOptTr" style="display:none;">
			<td>
				<input type="checkbox" name="defaultOpt" value="901">정면
				<input type="checkbox" name="defaultOpt" value="902">후면
				<input type="checkbox" name="defaultOpt" value="903">측면
				<input type="checkbox" name="defaultOpt" value="904">단체컷
				<input type="checkbox" name="defaultOpt" value="905">패키지컷
			</td>
		</tr>
		</table>
	</td>
</tr>

<% If isUpdateDate > Date() Then %>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">메인 촬영 컨셉</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr><% call CheckBoxUseType("doc_use_concept", "", "") %></tr>
		</table>
	</td>
</tr>
<% End If %>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 특징 및 주요 전달 사항 *</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="req_etc1" rows="18" class="textarea" style="width:100%"></textarea></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">관련 링크 및 참조 URL </td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="req_url" size="60" maxlength="200" class="text" value="http://">
	</td>
</tr>
<input type="hidden" name="req_etc2" size="10" >
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">촬영 시 유의사항</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="req_etc2" rows="5" class="textarea" style="width:100%"></textarea></td>
</tr>-->
</table>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td bgcolor="#FFFFFF" height="30">
		<input type = "button" class="button" value="저장" onclick="form_check();">
		<input type = "button" class="button" value="취소" onClick="window.location='/admin/photo_req/request_list.asp?menupos=<%=menupos%>'">
	</td>
</tr>
</table>

</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->