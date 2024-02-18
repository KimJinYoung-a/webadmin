<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual ="/lib/classes/partnerusercls.asp" -->
<!-- #include virtual="/lib/classes/lecture_itemregcls.asp"-->
<%
Sub SelectBoxDesignerItem1()
	dim query1
	%>
	<select name="tempid" onchange="TnDesignerNMargineAppl(this.value);">
	<option value=''>-- 업체선택 --</option>
	<%

	query1 = "select c.userid, c.coname from [db_user].[dbo].tbl_user_c c" + vbcrlf
	query1 = query1 + " where c.userdiv='14'" + vbcrlf

	rsget.Open query1,dbget,1

	if  not rsget.EOF  then
		rsget.Movefirst

		do until rsget.EOF
			response.write ("<option value='" & rsget("userid") & "," &rsget("coname") & "'>" & rsget("userid") & " (" & rsget("coname") & ")</option>")
			rsget.MoveNext
		loop

	end if

	rsget.close

	response.write("</select>")
End Sub

%>
<script language="JavaScript">
<!--

function checkform(form) {
//alert('금일(4월 22일) 서비스 점검 관계로 아이템 업로드가 불가능 합니다. 잠시만 기다려 주세요');
//return;

	var limitynv = "";
	var optionv="";
	var aa="";
	var bb="";
	var cc="";
	var dd="";
	var ee="";
	var ff="";
	var gg="";
	var hh="";

	aa=document.getElementById("imgmainload");
	bb=document.getElementById("imgbasicload");
	dd=document.getElementById("imgadd1load");
	ee=document.getElementById("imgadd2load");
	ff=document.getElementById("imgadd3load");
	gg=document.getElementById("imgadd4load");
	hh=document.getElementById("imgadd5load");

	for (var i = 0; i < form.limityn.length; i++) {
	if ( form.limityn[i].checked) {
		 limitynv = form.limityn[i].value
	   }
	}

	//for(var i=0; i<document.itemreg.realopt.options.length; i++) {
		//optionv += (document.itemreg.itemoptionnameno.value + document.itemreg.itemoptioncode.options[i].value + ",")
	//	optionv += (document.itemreg.realopt.options[i].value + ",")
	 //}

	if (form.cd1.value == ""){
	  alert("카테고리를 선택해주세요!");
	  form.cd1.focus();
	  return;
	}

	if (form.cd2.value == ""){
	  alert("카테고리를 선택해주세요!");
	  form.cd2.focus();
	  return;
	}

	if (form.cd3.value == ""){
	  alert("카테고리를 선택해주세요!");
	  form.cd3.focus();
	  return;
	}

//	if (i2ndcate.style.display=="inline"){
//		if (form.stylegubun.value == ""){
//		  alert("스타일 구분을 선택해주세요!");
//		  return;
//		}
//
//		if (form.itemstyle.value == ""){
//		  alert("스타일을 선택해주세요!");
//		  return;
//		}
//	}

	if(form.itemname.value == ""){
	  alert("상품명을 입력해주세요!");
	  form.itemname.focus();
	  return;
	}
//	else if(form.itemsource.value.length<1){
//	  alert("상품재질을 입력해주세요!");
//	  form.itemsource.focus();
//	  return;
//	}
//	else if(form.itemsize.value.length<1){
//	  alert("상품사이즈를 입력해주세요!");
//	  form.itemsize.focus();
//	  return;
//	}
//	else if(form.sourcearea.value == ""){
//	  alert("원산지를 입력해주세요!");
//	  form.sourcearea.focus();
//	  return;
//	}
//	else if(form.makename.value == ""){
//	  alert("제조사를 입력해주세요!");
//	  form.makename.focus();
//	  return;
//	}
//	else if(form.keywords.value == ""){
//	  alert("검색 키워드를 입력해주세요!");
//	  form.keywords.focus();
//	  return;
//	}
	else if(!IsDigit(form.sellcash.value)){
	  alert("판매가는 숫자만 가능합니다.");
	  form.sellcash.focus();
	  return;
	}
	else if(form.buycash.value == ""){
	  alert("공급가를 입력해주세요!");
	  form.buycash.focus();
	  return;
	}
	else if(!IsDigit(form.buycash.value)){
	  alert("공급가는 숫자만 가능합니다.");
	  form.buycash.focus();
	  return;
	}
	else if(limitynv == "Y" && form.limitno.value == ""){
	  alert("한정수량을 입력해주세요!");
	  form.limitno.focus();
	  return;
	}
	else if(limitynv == "Y" && !IsDigit(form.limitno.value)){
	  alert("한정수량은 숫자만 가능합니다.");
	  form.limitno.focus();
	  return;
	}
	//else if(form.itemoptionname.value == "" || optionv == ""){
	//  alert("옵션을 선택해주세요!");
	//  form.itemoptionname.focus();
	//}

	else if(aa.fileSize > 150000){
		alert("파일사이즈는 150Kbyte를 넘기실 수 없습니다...");
		form.imgmain.focus();
	}
	else if(aa.width > 610){
		alert("가로폭은 600픽셀을 넘기실 수 없습니다...");
		form.imgmain.focus();
	}
	else if(aa.height > 410){
		alert("세로폭은 400픽셀을 넘기실 수 없습니다...");
		form.imgmain.focus();
	}

//---------------------------------------------------------
	else if(bb.fileSize > 150000){
		alert("파일사이즈는 150Kbyte를 넘기실 수 없습니다...");
		form.imgbasic.focus();
	}
	else if(bb.width > 410){
		alert("가로폭은 400픽셀을 넘기실 수 없습니다...");
		form.imgbasic.focus();
	}
	else if(bb.height > 410){
		alert("세로폭은 400픽셀을 넘기실 수 없습니다...");
		form.imgbasic.focus();
	}
//---------------------------------------------------------
	else if(dd.fileSize > 150000){
		alert("파일사이즈는 150Kbyte를 넘기실 수 없습니다...");
		form.imgadd1.focus();
	}
	else if(dd.width > 610){
		alert("가로폭은 600픽셀을 넘기실 수 없습니다...");
		form.imgadd1.focus();
	}
	else if(dd.height > 410){
		alert("세로폭은 400픽셀을 넘기실 수 없습니다...");
		form.imgadd1.focus();
	}
//---------------------------------------------------------
	else if(ee.fileSize > 150000){
		alert("파일사이즈는 150Kbyte를 넘기실 수 없습니다...");
		form.imgadd2.focus();
	}
	else if(ee.width > 610){
		alert("가로폭은 600픽셀을 넘기실 수 없습니다...");
		form.imgadd2.focus();
	}
	else if(ee.height > 410){
		alert("세로폭은 400픽셀을 넘기실 수 없습니다...");
		form.imgadd2.focus();
	}
//---------------------------------------------------------
	else if(ff.fileSize > 150000){
		alert("파일사이즈는 150Kbyte를 넘기실 수 없습니다...");
		form.imgadd3.focus();
	}
	else if(ff.width > 610){
		alert("가로폭은 600픽셀을 넘기실 수 없습니다...");
		form.imgadd3.focus();
	}
	else if(ff.height > 410){
		alert("세로폭은 400픽셀을 넘기실 수 없습니다...");
		form.imgadd3.focus();
	}
//---------------------------------------------------------
	else if(gg.fileSize > 150000){
		alert("파일사이즈는 150Kbyte를 넘기실 수 없습니다...");
		form.imgadd4.focus();
	}
	else if(gg.width > 610){
		alert("가로폭은 600픽셀을 넘기실 수 없습니다...");
		form.imgadd4.focus();
	}
	else if(gg.height > 410){
		alert("세로폭은 400픽셀을 넘기실 수 없습니다...");
		form.imgadd4.focus();
	}
//---------------------------------------------------------
	else if(hh.fileSize > 150000){
		alert("파일사이즈는 150Kbyte를 넘기실 수 없습니다...");
		form.imgadd5.focus();
	}
	else if(hh.width > 610){
		alert("가로폭은 600픽셀을 넘기실 수 없습니다...");
		form.imgadd5.focus();
	}
	else if(hh.height > 410){
		alert("세로폭은 400픽셀을 넘기실 수 없습니다...");
		form.imgadd5.focus();
	}

//---------------------------------------------------------
	//else if(form.imglist.value == ""){
	//  alert("리스트이미지를 선택해주세요!");
	//  form.imglist.focus();
	//}
	//else if(form.imgsmall.value == ""){
	//  alert("스몰이미지를 선택해주세요!");
	//  form.imgsmall.focus();
	//}

    else{
		if(confirm("상품을 올리시겠습니까?") == true){
		//form.itemoptioncode2.value=optionv;
		//alert(form.itemoptioncode2.value);
<!--		form.submit();-->
		}
	}
}


function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var isellvat, ibuyvat, imileage;
	imargin = frm.margin.value;
	isellcash = frm.sellcash.value;

	isvatinclude = frm.vatinclude.value;

	if (imargin.length<1){
		alert('마진을 입력하세요.');
		frm.margin.focus();
		return;
	}

	if (isellcash.length<1){
		alert('판매가를 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDouble(imargin)){
		alert('마진은 숫자로 입력하세요.');
		frm.margin.focus();
		return;
	}

	if (!IsDigit(isellcash)){
		alert('판매가는 숫자로 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (isvatinclude=='Y'){
		isellvat = parseInt(parseInt(1/11 * parseInt(isellcash)));
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = parseInt(parseInt(1/11 * parseInt(ibuycash)));
		imileage = parseInt(isellcash*0.01) ;
	}else{
		isellvat = 0;
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = 0;
		imileage = parseInt(isellcash*0.01) ;
	}

	frm.sellvat.value = isellvat;
	frm.buycash.value = ibuycash;
	frm.buyvat.value = ibuyvat;
	frm.mileage.value = imileage;
}

function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.lecturerid.value = varArray[0];
	document.itemreg.lecturer.value = varArray[1];


}

function TnBasicItemInfo(){
	window.open("/admin/lecture/basic_lecture_info.asp","option_win","width=300,height=200,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes");
}

//-->
</script>

<form name="itemreg" method="post" action="http://partner.10x10.co.kr/admin/shopmaster/lecture_itemreg_upload_bywebadmin.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="designerid" value="">

<!-- itemreg.asp 고정값임-->
<input type="hidden" name="itemsource" value=""><!-- 상품재질 -->
<input type="hidden" name="itemsize" value=""><!-- 상품사이즈 -->
<input type="hidden" name="sourcearea" value=""><!-- 원산지 -->
<input type="hidden" name="makename" value=""><!-- 제조사 -->
<input type="hidden" name="mwdiv" value="U"><!-- 매입위탁구분, 업체무료배송(U)-->
<input type="hidden" name="vatinclude" value="N"><!-- 과세, 면세 여부, N-->
<input type="hidden" name="deliverytype" value="5"><!-- 배송구분, N-->
<input type="hidden" name="limityn" value="Y"><!-- 한정판매구분, Y-->
<input type="hidden" name="pojangok" value="Y"><!-- 포장가능여부, N-->
<input type="hidden" name="sellyn" value="N"><!-- 판매여부, Y-->
<input type="hidden" name="dispyn" value="N"><!-- 전시여부, N-->
<input type="hidden" name="isusing" value="N"><!-- 사용여부, N-->
<input type="hidden" name="usinghtml" value="N"><!-- HTML사용유무, N-->
<input type="hidden" name="itemcontent" value=""><!-- 아이템 설명-->
<input type="hidden" name="ordercomment" value=""><!-- 주문시 유의사항 -->
<input type="hidden" name="designercomment" value=""><!-- 업체코멘트 -->
<!-- itemreg.asp -->

<table width="750" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#3d3d3d">

<tr>
	<td width="100%">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td colspan="4" style="padding-left:20"><a href="javascript:TnBasicItemInfo();"><font color="red">기본틀생성</font></a></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120" style="spacing-left:1px">강좌 월 구분 <font color="red">(*)</font></td>
				<td colspan="3"><input type="text" name="yyyymm" value="" size="7" maxlength="7" class="input_b">(<%= Left(now(),7) %>)</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">브랜드 <font color="red">(*)</font></td>
				<td><% SelectBoxDesignerItem1 %></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강좌 카테고리 <font color="red">(*)</font></td>
				<td>
					<select name="cd1">
						<option value="95">전시안함[95]</option>
					</select>
					<select name="cd2">
						<option value="20">College[20]</option>
					</select>
					<select name="cd3">
						<option value="10">College[10]</option>
					</select>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강좌명 <font color="red">(*)</font></td>
				<td><input type="text" name="itemname" maxlength="64" size="50" class="input_b"></td>
			</tr>
		</table>
	</td>
</tr>


<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">검색키워드 <font color="red">(*)</font></td>
		<td colspan="3"><input type="text" name="keywords" maxlength="50" size="50" class="input_b" value="강좌,아카데미,수공예,컬리지">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">마진</td>
				<td colspan="3" >
					<input type="text" name="margin" maxlength="32" size="5" class="input_b" value="50">%
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">판매가(소비자가) <font color="red">(*)</font></td>
				<td colspan="3"><input type="text" name="sellcash" maxlength="16" size="16" class="input_b">원&nbsp;&nbsp;<input type="text" name="sellvat" maxlength="32" size="10" class="input_b">&nbsp;&nbsp;<font color="red"><input type="button" value="공급가 자동 계산" class="button" onclick="CalcuAuto(itemreg);"></font></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">매입가 <font color="red">(*)</font></td>
				<td colspan="3">
					<input type="text" name="buycash" maxlength="16" size="16" class="input_b">원&nbsp;&nbsp;<input type="text" name="buyvat" maxlength="32" size="10" class="input_b"> (<b>부가세 포함가</b>로 입력해 주세요.)
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">마일리지 <font color="red">(*)</font></td>
				<td colspan="3"><input type="text" name="mileage" maxlength="32" size="10" class="input_b"> (기본 판매가의 1%)</td>
			</tr>
		</table>
	</td>
</tr>

<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">소속아이디</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecturerid" value=""  class="input_b"size="30" maxlength="32"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강사명</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecturer" value=""  class="input_b"size="30" maxlength="32"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강좌비</td>
				<td bgcolor="#FFFFFF" width="250">
					<input type="text" name="lecsum" value="" class="input_b" size="12" maxlength="12">
					<input type="checkbox" name="matinclude">재료비포함
				</td>
				<td bgcolor="#DDDDFF" width="120">재료비</td>
				<td bgcolor="#FFFFFF"><input type="text" name="matsum" value=""  class="input_b" size="12" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">재료비설명</td>
				<td bgcolor="#FFFFFF"><input type="text" name="matdesc" value=""  class="input_b" size="90" maxlength="128"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">장소</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecspace" size="30" value=""  class="input_b"maxlength="64"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강좌횟수</td>
				<td bgcolor="#FFFFFF"><input type="text" name="leccount" value=""  class="input_b"size="6" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강의시간</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lectime" value=""  class="input_b"size="20" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">총강의시간</td>
				<td bgcolor="#FFFFFF"><input type="text" name="tottime" value=""  class="input_b"size="6" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강의기간(주기)</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecperiod" value=""  class="input_b"size="30" maxlength="64">(ex : 매주 금요일 몇시~몇시)</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120"><font color="red">*</font>한정수량<font color="red">*</font></td>
				<td bgcolor="#FFFFFF" width="160"><input type="text" name="limitno" maxlength="32" style="background-color:#FFFFFF;" class="input_b">(개)</td>
				<td bgcolor="#DDDDFF" width="120">적정인원</td>
				<td bgcolor="#FFFFFF"><input type="text" name="properperson" value="" class="input_b" size="6" maxlength="12"></td>
				<td bgcolor="#DDDDFF" width="120">최소인원</td>
				<td bgcolor="#FFFFFF" ><input type="text" name="minperson" value="" class="input_b" size="6" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">예약등록일</td>
				<td bgcolor="#FFFFFF" width="250"><input type="text" name="reservestart" value="" class="input_b" size="15" maxlength="10" onclick="calender_open('reservestart');"></td>
				<td bgcolor="#DDDDFF" width="120">예약마감일</td>
				<td bgcolor="#FFFFFF"><input type="text" name="reserveend" value="" class="input_b" size="15" maxlength="10" onclick="calender_open('reserveend');"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강좌내용<br>(커리큘럼)</td>
				<td bgcolor="#FFFFFF">
					<table border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a" >
						<tr bgcolor="#DDDDFF">
							<td>1주</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate01" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate01');">~<input type="text" name="lecdate01_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate01_end');">(2004-06-06 14:00:00)</td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>2주</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate02" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate02');">~<input type="text" name="lecdate02_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate02_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>3주</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate03" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate03');">~<input type="text" name="lecdate03_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate03_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>4주</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate04" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate04');">~<input type="text" name="lecdate04_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate04_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>5주</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate05" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate05');">~<input type="text" name="lecdate05_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate05_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>6주</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate06" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate06');">~<input type="text" name="lecdate06_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate06_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>7주</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate07" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate07');">~<input type="text" name="lecdate07_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate07_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>8주</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate08" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate08');">~<input type="text" name="lecdate08_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate08_end');"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">강좌개요</td>
				<td bgcolor="#FFFFFF"><textarea name="leccontents" class="input_b" rows="10" cols="80"></textarea></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">커리큘럼소개</td>
				<td bgcolor="#FFFFFF"><textarea name="leccurry" class="input_b" rows="10" cols="80"></textarea></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">기타사항</td>
				<td bgcolor="#FFFFFF"><textarea name="lecetc" class="input_b" rows="10" cols="80"></textarea></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">접수종료</td>
				<td bgcolor="#FFFFFF">
				&nbsp;&nbsp;&nbsp;
				<input type="radio" name="regfinish" value="N" > 접수중
				<input type="radio" name="regfinish" value="Y" checked > 접수종료
				</td>
			</tr>
		</table>
	</td>
</tr>

<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">사용여부</td>
				<td bgcolor="#FFFFFF">
				&nbsp;&nbsp;&nbsp;
				<input type="radio" name="isusing" value="Y" checked > 사용중(전시함)
				<input type="radio" name="isusing" value="N"  > 사용안함(전시안함)
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">사용여부</td>
				<td bgcolor="#FFFFFF"><input type="button" value="내용저장" onclick="CheckForm();return false;">&nbsp;&nbsp;&nbsp;</td>
			</tr>
		</table>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->