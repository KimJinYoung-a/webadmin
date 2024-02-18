<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->

<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<%
Dim code_large , code_mid , weclass , chgstyle , classtype

	code_large  = RequestCheckvar(request("code_large"),3)
	code_mid   = RequestCheckvar(request("code_mid"),3)
	weclass = RequestCheckvar(request("weclass"),1)

	If weclass = "Y" Then
		chgstyle = "style='display:none;'"
		classtype = "<font color='red'><strong>WeClass 단체 수강 입력</storng></font>"
	End If 

'강사 아이디,강사별 마진 표시
'<option value="강사아이디,강사이름(소속),마진>아이디,강사이름(소속)</option>
'''db_academy.dbo.tbl_lec_user ??
public Sub SelectLecturerId()
	dim sqlStr,i
''	sqlStr = "select  c.userid,p.company_name,c.defaultmargine, c.regdate, u.lec_margin, u.mat_margin"
''	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
''	sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
''	sqlStr = sqlStr + "     left join [ACADEMYDB].db_academy.dbo.tbl_lec_user u on c.userid=u.lecturer_id"
''	sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
''	sqlStr = sqlStr + " and c.userdiv < 22" + vbcrlf
''	sqlStr = sqlStr + " and c.userdiv='14'" + vbcrlf

	sqlStr = "select  u.lecturer_id, g.lecturer_name, u.lec_margin, u.mat_margin, u.regdate, u.lecturer_name as brandName"
	sqlStr = sqlStr + " from db_academy.dbo.tbl_lec_User u"
	sqlStr = sqlStr + "     left join db_academy.dbo.tbl_corner_good g"
	sqlStr = sqlStr + "     on u.lecturer_id=g.lecturer_id"
	sqlStr = sqlStr + " where u.lec_yn='Y'" + vbCrlf
	sqlStr = sqlStr + " order by u.lecturer_id"
	
    rsAcademyget.CursorLocation = adUseClient
    rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
	if not rsAcademyget.eof then
			response.write "<select name='temp_lec_id' onchange='javascript:FnLecturerApp(this.value);'>"
			response.write "<option value=''>선택</option>"
		for i=0 to rsAcademyget.recordcount-1
			response.write "<option value='" & db2html(rsAcademyget("lecturer_id")) & "," & db2html(rsAcademyget("lecturer_name")) & "," & rsAcademyget("lec_margin") & "," & rsAcademyget("mat_margin") & "," & left(rsAcademyget("regdate"),10) & "'>" & db2html(rsAcademyget("lecturer_id")) & "(" & db2html(rsAcademyget("lecturer_name")) & ") - "&db2html(rsAcademyget("brandName"))&"</option>"
		rsAcademyget.movenext
		next
			response.write "</select>"
	end if
    rsAcademyget.Close
end sub
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--

//submit 기본사항 체크
function frmsub(frm){

	//var frm=document.lecfrm;

	if (frm.CateCD1.value.length < 1){
		alert('클래스 구분을 선택해 주세요');
		frm.CateCD1.focus();
		return;
	}

//	if (frm.CateCD2.value.length < 1){
//		alert('강좌 구분을 선택해 주세요');
//		frm.CateCD2.focus();
//		return;
//	}

	if (frm.CateCD3.value.length < 1){
		alert('장소 구분을 선택해 주세요');
		frm.CateCD3.focus();
		return;
	}

	if (frm.classlevel.value.length < 1){
		alert('강좌 등급을 선택해 주세요');
		frm.classlevel.focus();
		return;
	}
	<% if weclass="" or weclass="N" then%>
	if (frm.lec_date.value.length < 1){
		alert('강좌월을 입력해 주세요.');
		frm.lec_date.focus();
		return;
	}
	<% end if %>

	if (frm.lec_title.value.length < 1){
		alert('강좌명을 입력해 주세요.');
		frm.lec_title.focus();
		return;
	}

	if (frm.lecturer_id.value.length < 1){
		alert('강사를 선택해 주세요.');
		frm.lecturer_id.focus();
		return;
	}


	//if (frm.lec_cost.value.length < 1|| frm.lec_cost.value==0){
	if (frm.lec_cost.value.length < 1){
		alert('수강료를 입력해 주세요.');
		frm.lec_cost.focus();
		return;
	}
	
	//if (frm.buying_cost.value.length < 1|| frm.buying_cost.value==0){
	if (frm.buying_cost.value.length < 1){
		alert('매입가 자동계산을 해주세요.');
		frm.buying_cost.focus();
		return;
	}
//2016/12/13 셀프워크샵 정산구분 생성
    if (frm.code_large.value=="76"){
        if (frm.lecjgubun.value.length<1){
            alert('카테고리가 셀프워크샵 인 경우 정산방식을 선택하세요.');
            frm.lecjgubun.focus();
            return;
        }
    }else{
        if (frm.lecjgubun.value=="1"){
            alert('카테고리가 셀프워크샵이 아닌경우 정산방식을 기본이나 선택(안함) 으로 선택하세요.');
            frm.lecjgubun.focus();
            return;
        }
    }
	
//가격은 추후 정리 -_-;

	if (frm.mileage.value.length < 1){
		alert('마일리지를 입력해 주세요.');
		frm.mileage.focus();
		return;
	}

	//if (frm.mat_contents.value.length < 1){
	//	alert('재료비 설명을 입력해 주세요.');
	//	frm.mat_contents.focus();
	//	return;
	//}

	// 기타내용 에디터에서 폼으로 입력
//	if (sector_1.chk==0){
//		frm.lec_etccontents.value = editor.document.body.innerHTML;
//	} else if(sector_1.chk!=3){
//		frm.lec_etccontents.value = editor.document.body.innerText;
//	}

	frm.submit();
}

//재료비 포함될때 input box disable 시킴
function Fnmat(){
var frm=document.lecfrm;

	if(frm.matinclude_yn.checked){
		frm.matinclude_yn.value='';
		frm.mat_cost.disabled='on';
	}else{
		frm.matinclude_yn.value='on';
		frm.mat_cost.disabled='';
	}
}

//강사별 마진,아이디,소속 표시
function FnLecturerApp(str){
	var varArray;
	varArray = str.split(',');
    
    if (varArray[0]){
    	document.lecfrm.lecturer_id.value = varArray[0];
    	document.lecfrm.lecturer_name.value = varArray[1];
    	document.lecfrm.margin.value = varArray[2];
    	document.lecfrm.mat_margin.value = varArray[3];
    	document.lecfrm.lecturer_regdate.value = varArray[4];
    }else{
        document.lecfrm.lecturer_id.value = "";
        document.lecfrm.lecturer_name.value = "";
    	document.lecfrm.margin.value = 0;
    	document.lecfrm.mat_margin.value = 0;
    	document.lecfrm.lecturer_regdate.value = "";
    }
    
	CalcuAuto(document.lecfrm);
}

//매입가 자동 계산 표시
function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var isellvat, ibuyvat, imileage;

	imargin = frm.margin.value;
	isellcash = frm.lec_cost.value;

	if (imargin.length<1){
		alert('마진을 입력하세요.');
		frm.margin.focus();
		return;
	}

	if (isellcash.length<1){
		alert('판매가를 입력하세요.');
		frm.lec_cost.focus();
		return;
	}

	if (!IsDouble(imargin)){
		alert('마진은 숫자로 입력하세요.');
		frm.margin.focus();
		return;
	}

	if (!IsDigit(isellcash)){
		alert('판매가는 숫자로 입력하세요.');
		frm.lec_cost.focus();
		return;
	}

	isellvat = 0;
	ibuycash = isellcash - parseInt(isellcash*imargin/100);
	ibuyvat = 0;
	imileage = parseInt((isellcash*1 + frm.mat_cost.value*1)*0.01) ;


	//frm.sellvat.value = isellvat;
	//frm.lec_cost.value = isellvat;
	frm.buying_cost.value=ibuycash;
	//frm.buyvat.value = ibuyvat;
	frm.mileage.value = imileage;
}

function CalcuAutoMaT(frm){

	if (frm.mat_margin.value.length<1){
		alert('재료비 마진을 입력하세요.');
		frm.mat_margin.focus();
		return;
	}

	if (frm.mat_cost.value.length<1){
		alert('재료비를 입력하세요.');
		frm.mat_cost.focus();
		return;
	}

	if (!IsDouble(frm.mat_margin.value)){
		alert('마진은 숫자로 입력하세요.');
		frm.mat_margin.focus();
		return;
	}

	if (!IsDigit(frm.mat_cost.value)){
		alert('재료비는 숫자로 입력하세요.');
		frm.mat_cost.focus();
		return;
	}

	var ibuycash = frm.mat_cost.value*1 - parseInt(frm.mat_cost.value*frm.mat_margin.value/100);

	frm.mat_buying_cost.value=ibuycash;
	frm.mileage.value = parseInt((frm.lec_cost.value*1 + frm.mat_cost.value*1)*0.01) ;
}


//지난 강좌 이미지 사용 여부 보여주기.
function showimgyn(){
	var frm = imagetag.style
	frm.display='block';
}

//강좌시간 추가시 input box 추가 
function addtime(){
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<div>"));
	timetbl.insertAdjacentText("BeforeEnd","시작일시 ");
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Sday' value=''>"));
	timetbl.insertAdjacentText("BeforeEnd"," ");
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_STime' value='00:00'>"));
 	timetbl.insertAdjacentText("BeforeEnd"," ~ 종료일시 ");
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Eday' value=''>"));
	timetbl.insertAdjacentText("BeforeEnd"," ");
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_ETime' value='00:00'>"));
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input type='hidden' name='lecOption' value=''>"));
}

//약도 선택 팝업창
function popmap(){
	popwin = window.open('/academy/lecture/lib/pop_lec_mapimg.asp','popMap','width=720,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//기존 강좌 불러오기 팝업(지난 강좌에서 불러오기)
function PopOldLectureList(frm){

	popwin = window.open('/academy/lecture/lib/pop_lec_list.asp?lecturer='+ frm.lecturer_id.value ,'Listwin','width=720,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}


function InsertCate(){
	popwin = window.open('/academy/lecture/lib/pop_lec_cate.asp' ,'Listwin','width=370,height=30,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//-->
</script>

<table width="800" border="0" align="center"  class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="lecfrm" method="post" action="<%=UploadImgFingers%>/linkweb/doFingerLecture.asp">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="oldidx" value="">
<input type="hidden" name="hidweclass" value="<%=weclass%>">
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">강좌코드</td>
		<td width="500" bgcolor="#FFFFFF" align="left">
			<input type="button" value="지난강좌에서 불러오기" onclick="PopOldLectureList(lecfrm);">&nbsp;&nbsp;<%=classtype%>
		</td>
	</tr>
	</tr>

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">강좌 구분</td>
		<td bgcolor="#FFFFFF" align="left">
			<select name="lec_gubun">
				<option value="0">일반</option>
				<option value="1">단체</option>
			</select>
		</td>
	</tr>

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">클래스 구분</td>
		<td bgcolor="#FFFFFF" align="left"><%=makeCateSelectBox("CateCD1","") %></td>
	</tr>
		<tr align="center" bgcolor="#DDDDFF">
		<td width="80">장소 구분</td>
		<td bgcolor="#FFFFFF" align="left"><%=makeCateSelectBox("CateCD3","")%></td>
	</tr>
	</tr>
		<tr align="center" bgcolor="#DDDDFF">
		<td width="80">등급 구분</td>
		<td bgcolor="#FFFFFF" align="left">
			<select name="classlevel">
				<option value="">::선택::</option>
				<option value="1">초급</option>
				<option value="2">중급</option>
				<option value="3">고급</option>
			</select>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">카테고리구분(New)</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="hidden" name="code_large" value="">
			<input type="hidden" name="code_mid" value="">
			<input type="text" name="large_name" value="" readonly size="20"  class="text_ro">
			<input type="text" name="mid_name" value="" readonly size="20"  class="text_ro">
			<input type="button" value="카테고리 선택" onclick="InsertCate()">
		</td>
	</tr>
    <tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
    	<td width="80">정산방식</td>
    	<td bgcolor="#FFFFFF" align="left">
    		<select name="lecjgubun" >
    		    <option value="">선택
    		    <option value="0">기본(강좌 원천징수 정산)
    		    <option value="1">수수료정산
    		</select>
    	</td>
    </tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">강좌월</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="7" size="7" name="lec_date" value="">
			(입력예시:2016-08)
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">강좌명</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="64" name="lec_title" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">강사</td>
		<td bgcolor="#FFFFFF" align="left">
		<% SelectLecturerId() %> (강사리스트에 강좌사용여부를 설정하셔야 나옵니다.)
			<input type="hidden" name="lecturer_id" value="">
			<input type="hidden" name="lecturer_name" value="">
			<input type="hidden" name="lecturer_regdate" value="">
		</td>
	</tr>
	
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">약도</td>
		<td bgcolor="#FFFFFF" align="left">
		    <input class="text_ro" type="text" maxlength="3" size="3" name="map_idx" value="" readOnly >
			<input class="input_a" type="text" maxlength="128" size="64" name="lec_mapimg" value="">
			<input type="button" value="약도찾기" onclick="javascript:popmap();">
		</td>
	</tr>
    <tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">강좌장소</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="64" name="lec_space" value="">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="5"></td>
	</tr>

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">가격 설정</td>
		<td bgcolor="#FFFFFF" align="left">
			<table width="600" border="0" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td width="60">&nbsp;</td>
					<td align="Center">수강료</td>
					<td align="Center" width="10">X</td>
					<td align="Center">기본마진</td>
					<td align="Center" width="10">=</td>
					<td align="Left">매입가</td>
				</tr>
				<tr>
				    <td>&nbsp;</td>
					<td align="Center"><input class="input_a" type="text" maxlength="10" size="8" name="lec_cost" value="0"></td>
					<td align="Center">&nbsp;</td>
					<td align="Center"><input class="input_a" type="text" maxlength="8" size="4" name="margin"  value="50">%</td>
					<td align="Center">&nbsp;</td>
					<td align="left">
						<input class="input_a" type="text" maxlength="10" size="8" name="buying_cost"  value="">
						<input type="button" value="매입가 자동 계산" class="button" onclick="javascript:CalcuAuto(lecfrm);">
					</td>
				</tr>
				<tr>
				    <td height="10" colspan="6"></td>
				</tr>
				<tr>
				    <td align="Center"></td>
					<td align="Center">재료비</td>
					<td align="Center" width="10">X</td>
					<td align="Center">기본마진</td>
					<td align="Center" width="10">=</td>
					<td align="Left">매입가</td>
				</tr>
				<!-- 2010 리뉴얼시 변경 기존 matinclude_yn="Y"인 내역 없음 재료비0인 내역 matinclude_yn="X"로 변경 -->
				<tr>
				    <td align="Center">
					    <select name="matinclude_yn" onChange="">
					    <option value="X"  >재료비 없음
					    <option value="C" >재료비 함께결제
					    
						<!-- <option value="N"  style='color:#999999'>재료비 현장결제(기존방식) -->
					    </select>
					</td>
					<td align="Center">    
					    <input class="input_a" type="text" maxlength="10" size="8" name="mat_cost" value="0">
					</td>
					<td align="Center">&nbsp;</td>
					<td align="Center"><input class="input_a" type="text" maxlength="8" size="4" name="mat_margin"  value="0">%</td>
					<td align="Center">&nbsp;</td>
					<td align="left">
					    <input class="input_a" type="text" maxlength="10" size="8" name="mat_buying_cost"  value="0">
					    <input type="button" value="매입가 자동 계산" class="button" onclick="javascript:CalcuAutoMaT(lecfrm);">
					</td>
				</tr>
				
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">마일리지</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="15" size="10" name="mileage" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">작품구성</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="200" size="90" name="lec_attribute" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">작품크기</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="200" size="90" name="lec_size" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">재료설명</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="64" name="mat_contents" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">키워드등록</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="64" name="keyword" value="">
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="5"></td>
	</tr>

	<!--추가 입력 -->

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80"><%=chkiif(weclass="Y","최대인원","한정인원")%></td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="8" size="4" name="limit_count" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">최소인원</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="8" size="4" name="min_count" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">강좌접수시작</td>
		<td bgcolor="#FFFFFF" align="left">
	        <input id="reg_startday" name="reg_startday" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="reg_startday_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">강좌접수종료</td>
		<td bgcolor="#FFFFFF" align="left">
	        <input id="reg_endday" name="reg_endday" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="reg_endday_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "reg_startday", trigger    : "reg_startday_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "reg_endday", trigger    : "reg_endday_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">강좌 횟수 <br>/ 시간</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="8" size="4" name="lec_count" value=""> 회
			&nbsp;&nbsp;
			총<input class="input_a" type="text" maxlength="8" size="4" name="lec_time" value="">시간
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">강좌기간</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="32" name="lec_period" value="">(ex : 매주 금요일 몇시~몇시)
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">강좌시간</td>
		<td bgcolor="#FFFFFF" align="left">
			<table width="500" class="a" border="0" cellpadding="0" cellspacing="0" >
			<tr>
				<td>
					<div class="a" id="timetbl">
						<div>
						시작일시 <input class="input_a" type="text" maxlength="10" size="10" name="lec_Sday" value="<%=Date()%>"> <input class="input_a" type="text" maxlength="5" size="5" name="lec_STime" value="00:00"> ~
						종료일시 <input class="input_a" type="text" maxlength="10" size="10" name="lec_Eday" value="<%=Date()%>"> <input class="input_a" type="text" maxlength="5" size="5" name="lec_ETime" value="00:00">
						<input type="hidden" name="lecOption" value="">
						</div>
					</div>
				</td>
				<td><div class="a"><input type="button" value="시간추가" onclick="javascript:addtime();"></div></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">강좌소개</td>
		<td bgcolor="#FFFFFF" align="left">
			<textarea name="lec_outline" cols="76" rows="7"></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">강좌내용</td>
		<td bgcolor="#FFFFFF" align="left">
			<textarea name="lec_contents" cols="76" rows="7"></textarea>
		</td>
	</tr>

	<!--2016-05-20 유태욱 추가 -->
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">커리큘럼</td>
		<td bgcolor="#FFFFFF" align="left">
			<textarea name="lec_curriculum" cols="76" rows="10">[day]</textarea>
			<br><font color="red">※ 일자 구분은 [day] 로 구분함.</font>
		</td>
	</tr>

	<!--2016-05-20 유태욱 추가(모바일 주의사항) -->
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">모바일 주의사항</td>
		<td bgcolor="#FFFFFF" align="left">
			<textarea name="lec_mocaution" cols="76" rows="10"></textarea>
		</td>
	</tr>
	
	<!-- 2016-05-19 유태욱 추가(동영상url)-->
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">동영상URL</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" size="90" name="lec_movie" value="">
			<br>
			<font color="red">
				<!--※ 비메오 : copy embed code 복사 (예 :</font><font color="blue"> //player.vimeo.com/video/102309330</font><font color="red"> ) http: 제외<br>-->
				※ 유튜브 : 소스코드 복사 (예 : </font><font color="blue">http://www.youtube.com/embed/qj4rn1I_dC8 </font><font color="red">)
				※ 유튜브 동영상 URL복사 아님!
			</font>
		</td>
	</tr>

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">개인준비물</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="200" size="90" name="lec_prepare" value="">
		</td>
	</tr>
	<% if (FALSE) then %>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">기타내용</td>
		<td bgcolor="#FFFFFF" align="left">
			<% 
				'에디터의 너비와 높이를 설정
				dim editor_width, editor_height, brd_content
				editor_width = "100%"
				editor_height = "250"
			%>
			<!-- INCLUDE Virtual="/lib/util/editor.asp" -->
			<input type="hidden" name="lec_etccontents" value="">
			<font color="#8c7301">
			<br>※1. 문단나누기 - 엔터 (Enter Key)
			<br>※2. 행나누기 - 시프트 + 엔터 (Shift + Enter Key)
			</font>
		</td>
	</tr>
    <% end if %>
    <tr align="center" bgcolor="#DDDDFF">
		<td width="80">기타내용</td>
		<td bgcolor="#FFFFFF" align="left">
		    <textarea name="lec_etccontents" cols="76" rows="10"></textarea>
		</td>
	</tr>
	    
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">접수여부</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="radio" name="reg_yn" value="Y">Y
			<input type="radio" name="reg_yn" value="N" checked>N
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">전시여부</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="radio" name="disp_yn" value="Y">Y
			<input type="radio" name="disp_yn" value="N" checked>N
		</td>
	</tr>
	
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="2">
			<span id="imagetag"  style="display:none">
			이미지 사용여부   :  <input type="checkbox" name="image_saveas_yn">(기존이미지를 사용하고자 할때 체크해 주세요)
			</span>
		</td>
	</tr>
	
	<tr align="center" bgcolor="#DDDDFF">
		<td colspan="2">
			<input  type="button" value="저장" onclick="javascript:frmsub(lecfrm);">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="취소">
		</td>
	</tr>
	
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->