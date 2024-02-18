<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스
' History : 2009.04.07 서동석 생성
'			2010.05.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim lec_idx ,hidweclass ,chgstyle,classtype
	lec_idx= RequestCheckvar(request("lec_idx"),10)

dim onelec
set onelec = new CLecture
	onelec.FRectidx=lec_idx
	onelec.GetOneLecture

dim lectime,i,j,oldOpCd
set lectime = new CLectime
	lectime.getlectime lec_idx

dim oLectoption
set oLectoption = new CLectOption
	oLectoption.FRectidx = lec_idx
	
	if lec_idx<>"" then
		oLectoption.GetLectOptionInfo
	end if

if (onelec.FOneItem.isWeClass) then
    hidweclass = "Y"
else
    hidweclass = "N"
end if

If hidweclass = "Y" Then
	chgstyle = "style='display:none;'"
	classtype = "<font color='red'><strong>WeClass 단체 수강 수정</storng></font>"
End If 

public Sub SelectLecturerId(byval lecturer_id)
	dim sqlStr,i
	
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
			response.write "<option value='" & db2html(rsAcademyget("lecturer_id")) & "," & db2html(rsAcademyget("lecturer_name")) & "," & rsAcademyget("lec_margin") & "," & rsAcademyget("mat_margin") & "," & left(rsAcademyget("regdate"),10) & "' "&CHKIIF(lecturer_id=(rsAcademyget("lecturer_id")),"selected","") &">" & db2html(rsAcademyget("lecturer_id")) & "(" & db2html(rsAcademyget("lecturer_name")) & ") - "&db2html(rsAcademyget("brandName"))&"</option>"
		rsAcademyget.movenext
		next
			response.write "</select>"
	end if
    rsAcademyget.Close
end Sub
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>

function popLecDateEdit1(lec_idx){
	var popwin = window.open('popLecOptionEdit.asp?lec_idx='+lec_idx,'popLecDateEdit','width=700,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function frmsub(frm){

	//var frm=document.lecfrm;

	if (frm.CateCD1.value.length < 1){
		alert('클래스 구분을 선택해 주세요');
		frm.CateCD1.focus();
		return;
	}

	if (frm.CateCD3.value.length < 1){
		alert('장소 구분을 선택해 주세요');
		frm.CateCD3.focus();
		return;
	}
	<% if hidweclass="" or hidweclass="N" then%>
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
	
	if ((frm.lec_count.value.length < 1)||(!IsDigit(frm.lec_count.value))){
	    alert('강좌 횟수는 숫자만 가능합니다.');
		frm.lec_count.focus();
		return;
	}
	
	if ((frm.lec_time.value.length < 1)||(!IsDouble(frm.lec_time.value))){
	    alert('강좌 시간은 숫자만 가능합니다.');
		frm.lec_time.focus();
		return;
	}
	
	
    //재료비
    
    
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
    
	// 기타내용 에디터에서 폼으로 입력
//	if (sector_1.chk==0){
//		frm.lec_etccontents.value = editor.document.body.innerHTML;
//	} else if(sector_1.chk!=3){
//		frm.lec_etccontents.value = editor.document.body.innerText;
//	}

    if (confirm('저장하시겠습니까?')){
	    frm.submit();
	}
}

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
	//imileage = parseInt(isellcash*0.01) ;
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


function showtime(){
	var frm=timetbl.style;
	if(frm.display=="none"){
		frm.display='';
	} else {
		frm.display='none';
	}
}

//강좌시간 추가시 input box 추가 
function addtime(tgt,opCd){
	var tfrm = document.all[tgt];
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<div>"));
	tfrm.insertAdjacentText("BeforeEnd","시작일시 ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Sday' value=''>"));
	tfrm.insertAdjacentText("BeforeEnd"," ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_STime' value='00:00'>"));
 	tfrm.insertAdjacentText("BeforeEnd"," ~ 종료일시 ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Eday' value=''>"));
	tfrm.insertAdjacentText("BeforeEnd"," ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_ETime' value='00:00'>"));
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input type='hidden' name='lecOption' value='" + opCd + "'>"));
}

//커리큘럼 추가시 input box 추가 
function addtime(tgt,opCd){
	var tfrm = document.all[tgt];
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<div>"));
	tfrm.insertAdjacentText("BeforeEnd","시작일시 ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Sday' value=''>"));
	tfrm.insertAdjacentText("BeforeEnd"," ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_STime' value='00:00'>"));
 	tfrm.insertAdjacentText("BeforeEnd"," ~ 종료일시 ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Eday' value=''>"));
	tfrm.insertAdjacentText("BeforeEnd"," ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_ETime' value='00:00'>"));
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input type='hidden' name='lecOption' value='" + opCd + "'>"));
}

function popmap(){
	popwin = window.open('/academy/lecture/lib/pop_lec_mapimg.asp','popMap','width=720,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//카테고리
function InsertCate(){
	popwin = window.open('/academy/lecture/lib/pop_lec_cate.asp' ,'Listwin','width=370,height=30,scrollbars=yes,resizable=yes');
	popwin.focus();
}

<% ''/리뉴얼 신규카테고리 미리 등록을 위한 임시. 리뉴얼후 삭제 %>
function tmpInsertCate(){
	popwin = window.open('/academy/lecture/lib/pop_lec_cate_tmp.asp' ,'Listwin','width=400,height=100,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<form name="lecfrm" method="post" action="<%=UploadImgFingers%>/linkweb/doFingerLecture.asp" style="margin:0px;">
<input type="hidden" name="mode" value="modi">
<input type="hidden" name="hidweclass" value="<%=hidweclass%>">
<input type="hidden" name="idx" value="<%=onelec.FOneItem.Fidx %>">

<table width="800" border="0" align="center"  class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌코드</td>
	<td width="500" bgcolor="#FFFFFF" align="left"><%=onelec.FOneItem.Fidx %>&nbsp;&nbsp;<%=classtype%></td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌 구분</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="lec_gubun">
			<option value="0" <%=chkiif(onelec.FOneItem.Flec_gubun = "0","selected","")%>>일반</option>
			<option value="1" <%=chkiif(onelec.FOneItem.Flec_gubun = "1","selected","")%>>단체</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">클래스 구분</td>
	<td bgcolor="#FFFFFF" align="left"><%=makeCateSelectBox("CateCD1",onelec.FOneItem.FCateCD1) %></td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">장소 구분</td>
	<td bgcolor="#FFFFFF" align="left"><%=makeCateSelectBox("CateCD3",onelec.FOneItem.FCateCD3)%></td>
</tr>
</tr>
	<tr align="center" bgcolor="#DDDDFF">
	<td width="80">등급 구분</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="classlevel">
			<option value="" <%=chkiif(onelec.FOneItem.Fclasslevel = "","selected","")%>>::선택::</option>
			<option value="1" <%=chkiif(onelec.FOneItem.Fclasslevel = "1","selected","")%>>초급</option>
			<option value="2" <%=chkiif(onelec.FOneItem.Fclasslevel = "2","selected","")%>>중급</option>
			<option value="3" <%=chkiif(onelec.FOneItem.Fclasslevel = "3","selected","")%>>고급</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">카테고리구분</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="hidden" name="code_large" value="<%=onelec.FOneItem.Fcode_large%>">
		<input type="hidden" name="code_mid" value="<%=onelec.FOneItem.Fcode_mid%>">
		<input type="text" name="large_name" value="<%=onelec.FOneItem.Fcode_large_nm%>" readonly size="20"  class="text_ro">
		<input type="text" name="mid_name" value="<%=onelec.FOneItem.Fcode_mid_nm%>" readonly size="20"  class="text_ro">
		<input type="button" value="카테고리 선택" onclick="InsertCate();" class="button">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
	<td width="80">정산방식</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="lecjgubun" >
		    <option value="">선택
		    <option value="0" <%=chkiif(onelec.FOneItem.Flecjgubun = "0","selected","")%>>기본(강좌 원천징수 정산)
		    <option value="1" <%=chkiif(onelec.FOneItem.Flecjgubun = "1","selected","")%>>수수료정산
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
	<td width="80">강좌월</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="7" size="7" name="lec_date" value="<%= onelec.FOneItem.Flec_date %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌명</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="128" size="64" name="lec_title" value="<%= onelec.FOneItem.Flec_title %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강사</td>
	<td bgcolor="#FFFFFF" align="left">
	<% SelectLecturerId(onelec.FOneItem.Flecturer_id) %> (강사리스트에 강좌사용여부를 설정하셔야 나옵니다.)
		<input type="hidden" name="lecturer_id" value="<%= onelec.FOneItem.Flecturer_id %>">
		<input type="hidden" name="lecturer_regdate" value="<%= left(onelec.FOneItem.Flecturer_regdate,10) %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강사명</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" name="lecturer_name" value="<%= onelec.FOneItem.Flecturer_name %>" size="10" maxlength="16">
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
	<td width="80">약도</td>
	<td bgcolor="#FFFFFF" align="left">
	    <input class="text_ro" type="text" maxlength="3" size="3" name="map_idx" value="<%= onelec.FOneItem.Fmap_idx %>" readOnly >
		<input class="input_a" type="text" maxlength="128" size="64" name="lec_mapimg" value="<%= onelec.FOneItem.Flec_mapimg %>">
		<input type="button" value="약도찾기" onclick="javascript:popmap();" class="button">
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
	<td width="80">강좌장소</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="128" size="64" name="lec_space" value="<%= onelec.FOneItem.Flec_space %>">
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
				<td align="Center"><input class="input_a" type="text" maxlength="10" size="8" name="lec_cost" value="<%= onelec.FOneItem.Flec_cost %>"></td>
				<td align="Center">&nbsp;</td>
				<td align="Center"><input class="input_a" type="text" maxlength="8" size="4" name="margin"  value="<%= onelec.FOneItem.Fmargin %>">%</td>
				<td align="Center">&nbsp;</td>
				<td align="left">
					<input class="input_a" type="text" maxlength="10" size="8" name="buying_cost"  value="<%= onelec.FOneItem.Fbuying_cost %>">
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
			<!-- 기존
			<input type="checkbox" name="matinclude_yn" onclick="javascript:Fnmat();" <% if onelec.FOneItem.Fmatinclude_yn="Y" then response.write "checked" %>>(재료비포함)
			<input class="input_a" type="text" maxlength="16" size="8" name="mat_cost" value="<%= onelec.FOneItem.Fmat_cost %>">
			-->
			<!-- 2010 리뉴얼시 변경 기존 matinclude_yn="Y"인 내역 없음 재료비0인 내역 matinclude_yn="X"로 변경 -->
			
			<tr>
			    <td align="Center">
			    <!-- option value="N" onelec.FOneItem.Fmatinclude_yn="N" 지울것 -->
				    <select name="matinclude_yn" onChange="">
				    <option value="X"  <%= CHKIIF(onelec.FOneItem.Fmatinclude_yn="X" or onelec.FOneItem.Fmat_cost=0,"selected","") %>>재료비 없음
				    <option value="C" <%= CHKIIF(onelec.FOneItem.Fmatinclude_yn="C","selected","") %> >재료비 함께결제
				    
					<!-- <option value="N" <%= CHKIIF(onelec.FOneItem.Fmatinclude_yn="N" or (onelec.FOneItem.Fmatinclude_yn="N" and onelec.FOneItem.Fmat_cost>0),"selected","") %> style='color:#999999'>재료비 현장결제(기존방식) -->
				    </select>
				</td>
				<td align="Center">    
				    <input class="input_a" type="text" maxlength="10" size="8" name="mat_cost" value="<%= onelec.FOneItem.Fmat_cost %>">
				</td>
				<td align="Center">&nbsp;</td>
				<td align="Center"><input class="input_a" type="text" maxlength="8" size="4" name="mat_margin"  value="<%= onelec.FOneItem.Fmat_margin %>">%</td>
				<td align="Center">&nbsp;</td>
				<td align="left">
				    <input class="input_a" type="text" maxlength="10" size="8" name="mat_buying_cost"  value="<%= onelec.FOneItem.Fmat_buying_cost %>">
				    <input type="button" value="매입가 자동 계산" class="button" onclick="javascript:CalcuAutoMaT(lecfrm);">
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">마일리지</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="15" size="10" name="mileage" value="<%= onelec.FOneItem.Fmileage %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">작품구성</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="200" size="90" name="lec_attribute" value="<%= onelec.FOneItem.Flec_attribute %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">작품크기</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="200" size="90" name="lec_size" value="<%= onelec.FOneItem.Flec_size %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">재료설명</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="128" size="64" name="mat_contents" value="<%= onelec.FOneItem.Fmat_contents %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">키워드등록</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="128" size="64" name="keyword" value="<%= onelec.FOneItem.Fkeyword %>">
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td colspan="2" height="5"></td>
</tr>

<!--추가 입력 -->

<tr align="center" bgcolor="#DDDDFF">
	<td width="80">총 한정인원</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="8" size="4" name="limit_count" value="<%= onelec.FOneItem.Flimit_count %>" readonly style="background-color='#EEEEEE'">
		※ 한정인원의 수정은 일정(옵션)수정에서 해주세요.
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">총 접수인원</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="8" size="4" name="limit_sold" value="<%= onelec.FOneItem.Flimit_sold %>" readonly style="background-color='#EEEEEE'">
		※ 접수인원의 수정은 일정(옵션)수정에서 해주세요.
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌당<br>최소인원</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="8" size="4" name="min_count" value="<%= onelec.FOneItem.Fmin_count %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌접수시작</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="reg_startday" name="reg_startday" value="<%=onelec.FOneItem.Freg_startday%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="reg_startday_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌접수종료</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="reg_endday" name="reg_endday" value="<%=onelec.FOneItem.Freg_endday%>" class="text" size="10" maxlength="10" />
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
	<td width="80">강좌 횟수 /<br> 시간</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="8" size="4" name="lec_count" value="<%= onelec.FOneItem.Flec_count %>"> 회
		&nbsp;&nbsp;
		총<input class="input_a" type="text" maxlength="8" size="4" name="lec_time" value="<%= onelec.FOneItem.Flec_time %>">시간
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강의일정<br><input type="button" value="일정수정" onclick="popLecDateEdit1('<%= lec_idx %>');" class="button"></td>
	<td bgcolor="#FFFFFF" align="left">
	<!-- 수정시에는 옵션 목록.. -->
	<table width="100%" border="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        	<td width="40">코드</td>
        	<td width="100">접수기간</td>
        	<td >강좌일시</td>
        	<td width="30">사용<br>여부</td>
        	<td width="30">정원</td>
        	<td width="30">신청<br>인원</td>
        	<td width="30">남은<br>인원</td>
        	<td width="40">마감<br>여부</td>
        </tr>
	<% for i=0 to oLectoption.FResultCount -1 %>
	
	    <tr align="center" bgcolor="<%= ChkIIF(oLectoption.FItemList(i).Fisusing="Y","#FFFFFF","#DDDDDD") %>">
        	<td><%= oLectoption.FItemList(i).FlecOption %></td>
        	<td>
        		시작: <%= FormatDateTime(oLectoption.FItemList(i).FRegStartDate,2) %><br>
        		종료: <%= FormatDateTime(oLectoption.FItemList(i).FRegEndDate,2) %>
        	</td>
        	<td align="left">
        		<%=oLectoption.FItemList(i).FlecOptionName%><br>
        		<%= FormatDateTime(oLectoption.FItemList(i).FlecStartDate,2) %>&nbsp;
        		<%= FormatDateTime(oLectoption.FItemList(i).FlecStartDate,4) %>~
        		<%= FormatDateTime(oLectoption.FItemList(i).FlecEndDate,4) %>
        	</td>
        	<td>
        	    <%= oLectoption.FItemList(i).Fisusing %>
        	</td>
        	<td><%= oLectoption.FItemList(i).Flimit_count %></td>
        	<td><%= oLectoption.FItemList(i).Flimit_sold %></td>
        	<td><%= oLectoption.FItemList(i).Flimit_count-oLectoption.FItemList(i).Flimit_sold %></td>
        	<td><% if oLectoption.FItemList(i).IsOptionSoldOut then %><font color="red">마감</font><% end if %></td>
        </tr>
	<% next %>
	</table>
	<!-- lec_period 사용안함...
		<input class="input_a" type="text" maxlength="128" size="40" name="lec_period" value="<%= onelec.FOneItem.Flec_period %>" readonly style="background-color='#EEEEEE'">(ex : 매주 금요일 몇시~몇시)
	    -->
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌시간</td>
	<td bgcolor="#FFFFFF" align="left">
		<table width="100%" class="a" border="0" cellpadding="2" cellspacing="2" >
		<%
			for i = 0 to lectime.FResultcount -1
				if oldOpCd<>lectime.FlecOption(i) then
					j=j+1
		%>
		<tr align="center">
			<td width="50" bgcolor="#E8E8E8"><%=lectime.FlecOption(i)%></td>
			<td bgcolor="#F2F2F2"><div class="a" id="timetbl<%=j%>">
		<%
				end if
				oldOpCd = lectime.FlecOption(i)
		%>
					<div>
					시작일시 <input class="input_a" type="text" maxlength="10" size="10" name="lec_Sday" value="<% = formatdatetime(lectime.FStartDate(i),2) %>"> <input class="input_a" type="text" maxlength="5" size="5" name="lec_STime" value="<% = formatdatetime(lectime.FStartDate(i),4) %>"> ~
					종료일시 <input class="input_a" type="text" maxlength="10" size="10" name="lec_Eday" value="<% = formatdatetime(lectime.FEndDate(i),2) %>"> <input class="input_a" type="text" maxlength="5" size="5" name="lec_ETime" value="<% = formatdatetime(lectime.FEndDate(i),4) %>">
					<input type="hidden" name="lecOption" value="<%=lectime.FlecOption(i)%>">
					</div>
			<% if (i>=(lectime.FResultcount-1) ) or (i<lectime.FResultcount and oldOpCd<>lectime.FlecOption(i+1)) then %>
				</div>
			</td>
			<td width="110" bgcolor="#F2F2F2"><div class="a"><input type="button" value="시간추가 #<%=j%>" onclick="javascript:addtime('timetbl<%=j%>','<%=lectime.FlecOption(i)%>');" class="button"></div></td>
		</tr>
			<%
				end if
			%>
		<% next %>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌소개</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="lec_outline" cols="76" rows="7"><%= onelec.FOneItem.Flec_outline %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">강좌내용</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="lec_contents" cols="76" rows="7"><%= onelec.FOneItem.Flec_contents %></textarea>
	</td>
</tr>

<!--2016-05-20 유태욱 추가 -->
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">커리큘럼</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="lec_curriculum" cols="76" rows="10"><%=chkiif(onelec.FOneItem.Flec_curriculum = "","[day]1",onelec.FOneItem.Flec_curriculum)%></textarea>
		<br><font color="red">※ 일자 구분은 [day] 로 구분함.</font>
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF">
	<td width="80">개인준비물</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="200" size="90" name="lec_prepare" value="<%= onelec.FOneItem.Flec_prepare %>">
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

			brd_content = onelec.FOneItem.Flec_etccontents
			if inStr(brd_content,"<br>")=0 and inStr(brd_content,"<P>")=0 then
				brd_content = replace(brd_content,vbCrLf,"<br>")
			end if
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
	    <textarea name="lec_etccontents" cols="76" rows="10"><%= onelec.FOneItem.Flec_etccontents %></textarea>
	</td>
</tr>    
<!--2016-05-20 유태욱 추가(모바일 주의사항) -->
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">모바일 주의사항</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="lec_mocaution" cols="76" rows="10"><%= onelec.FOneItem.Flec_mocaution %></textarea>
	</td>
</tr>

<!-- 2016-05-19 유태욱 추가(동영상url)-->
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">동영상URL</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" size="90" name="lec_movie" value="<%= onelec.FOneItem.Flec_movie %>">
		<br>
		<font color="red">
			<!--※ 비메오 : copy embed code 복사 (예 :</font><font color="blue"> //player.vimeo.com/video/102309330</font><font color="red"> ) http: 제외<br>-->
			※ 유튜브 : 소스코드 복사 (예 : </font><font color="blue">http://www.youtube.com/embed/qj4rn1I_dC8 </font><font color="red">)
			※ 유튜브 동영상 URL복사 아님!
		</font>
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF">
	<td width="80">접수여부</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="radio" name="reg_yn" value="Y" <% if onelec.FOneItem.Freg_yn="Y" then response.write "checked" %>>Y
		<input type="radio" name="reg_yn" value="N" <% if onelec.FOneItem.Freg_yn="N" then response.write "checked" %>>N
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">전시여부</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="radio" name="disp_yn" value="Y" <% if onelec.FOneItem.Fdisp_yn="Y" then response.write "checked" %>>Y
		<input type="radio" name="disp_yn" value="N" <% if onelec.FOneItem.Fdisp_yn="N" then response.write "checked" %>>N
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td colspan="2">
		<input  type="button" value="저장" onclick="javascript:frmsub(lecfrm);" class="button">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="취소" class="button">
	</td>
</tr>
</table>

</form>
<%
set onelec = nothing
set lectime = nothing
set oLectoption = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->