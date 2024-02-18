<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원별 급여 기본정보 등록
' History : 2010.12.23 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function_v3.asp"-->
<%
Dim sEmpno ,cMember, clsPayForm
Dim susername, susermail, sdirect070, djoinday, blnstatediv, spart_name, sposit_name, sjob_name
Dim startdate, enddate,defaultpay ,foodpay,jobpay ,inBreakTime  ,holidaywdtime	,regdate    ,lastupdate ,adminid,iposit_sn,dretireday,sjuminno,suserphone,susercell,szipcode,szipaddr,suseraddr
Dim StartHour(8), StartMinute(8), EndHour(8), EndMinute(8), BreakSHour(8), BreakSMinute(8),  BreakEHour(8), BreakEMinute(8),DutyTime(8) ,NightTime(8), iworktype(8)
Dim totDutyTime,iOverTime,iPatternSeq,part_sn,spatternname,totNightTime, iHolidayTime,avgWeek,totPaySum
Dim iTotCnt,iPageSize, iTotalPage,page
Dim arrList, intLoop
Dim ino, lastno
Dim realWeekWorkDay
dim idepartment_id,sdepartmentNameFull
dim iRectposit_sn, sRectposit_name, iRectdp_id, sRectdpNmFull, jobkind, placekind

sEmpno =   requestCheckvar(request("sEN"),14)
ino =requestCheckvar(request("ino"),10)
lastno =requestCheckvar(request("Lno"),10)
iPatternSeq = requestCheckvar(request("iPS"),10)


'// 4.345238095 == 월 평균 WEEK 수 = (365일 / 7일 / 12개월)
avgWeek = 4.345238095

realWeekWorkDay = 0

iPageSize = 5
page = requestCheckvar(Request("page"),10)
if page ="" then page = 1

	'// ========================================================================
	'사원 기본정보 가져오기
	'// ========================================================================
	Set cMember  = new CTenByTenMember
	cMember.Fempno		= sEmpno
	cMember.fnGetMemberData

	susername		= cMember.Fusername
	djoinday	  	= cMember.Fjoinday
	blnstatediv 	= cMember.Fstatediv
	iRectposit_sn		= cMember.Fposit_sn
	spart_name  	= cMember.Fpart_name
	sRectposit_name 	= cMember.Fposit_name
	sjob_name		= cMember.Fjob_name
	dretireday		= cMember.Fretireday
	sjuminno		= cMember.Fjuminno
	suserphone		= cMember.FuserPhone
	susercell		= cMember.Fusercell
	szipcode		= cMember.Fzipcode
	szipaddr		= cMember.Fzipaddr
	suseraddr		= cMember.Fuseraddr
	iRectdp_id	= cMember.Fdepartment_id
	sRectdpNmFull = cMember.FdepartmentNameFull
	Set cMember 	= nothing

	'// ========================================================================
	'사번에 해당하는 기입력 계약정보 리스트 가져오기
	'// ========================================================================
	Set clsPayForm = new CPayForm

	clsPayForm.Fempno= sEmpno
	clsPayForm.FPageSize= iPageSize
	clsPayForm.FCurrPage= page
	arrList = clsPayForm.fnGetDefaultPayList
	iTotCnt = clsPayForm.FTotCnt
dim sRectNo
	IF 	ino <>"" or lastno <> "" THEN
		'// ====================================================================
		'선택 계약정보 가져오기
		'// ====================================================================
		sRectNo = ino
		if ino ="" and lastno <> "" THEN '기존계약서 가져와야 할 경우 마지막 회차 계약서 정보 가져온다
			sRectNo = lastno
		end if
		clsPayForm.Fempno= sEmpno
		clsPayForm.Fino = sRectNo
		clsPayForm.fnGetDefaultPayData

		startdate		= clsPayForm.Fstartdate
		enddate			= clsPayForm.Fenddate

		defaultpay    	= clsPayForm.Fdefaultpay
		foodpay	    	= clsPayForm.Ffoodpay
		jobpay			= clsPayForm.Fjobpay

		inBreakTime		= clsPayForm.FinBreakTime
		iOverTime		= clsPayForm.FOverTime

		realWeekWorkDay = 0
		For intLoop = 1 To 7
			StartHour(intLoop) 		= clsPayForm.FStartHour(intLoop)
			StartMinute(intLoop)  	= clsPayForm.FStartMinute(intLoop)
			EndHour(intLoop)       	= clsPayForm.FEndHour(intLoop)
			EndMinute(intLoop)      = clsPayForm.FEndMinute(intLoop)
			BreakSHour(intLoop)     = clsPayForm.FBreakSHour(intLoop)
			BreakSMinute(intLoop)   = clsPayForm.FBreakSMinute(intLoop)
			BreakEHour(intLoop)     = clsPayForm.FBreakEHour(intLoop)
			BreakEMinute(intLoop)   = clsPayForm.FBreakEMinute(intLoop)
			DutyTime(intLoop)		= clsPayForm.FDutyTime(intLoop)
			iworktype(intLoop)		= clsPayForm.Fworktype(intLoop)
			NightTime(intLoop)		= clsPayForm.FNightTime(intLoop)
			if (iworktype(intLoop) = "1") then
				'// TODO : 한시간 미만 근무는 고려하지 않는다.(스크립트에서 처리함)
				realWeekWorkDay = realWeekWorkDay + 1
			end if
		Next

		totDutyTime  	= clsPayForm.FTotDutyTime
		totNightTime	= clsPayForm.FtotNightTime
		totPaySum		= clsPayForm.FTotPaySum

		holidaywdtime	= clsPayForm.Fholidaywdtime
		regdate         = clsPayForm.Fregdate
		lastupdate      = clsPayForm.Flastupdate
		adminid         = clsPayForm.Fadminid

		iposit_sn		= clsPayForm.Fposit_sn
		sposit_name 	= clsPayForm.Fposit_name
		idepartment_id	= clsPayForm.Fdepartment_id
		sdepartmentNameFull = clsPayForm.FdepartmentNameFull
		jobkind		= clsPayForm.Fjobkind
		placekind		= clsPayForm.Fplacekind
	END IF

'신규등록이나 기존계약내용가져오기 할때는 현재의 계약형태, 부서 가져온다
		if iposit_sn = "" or isNull(iposit_sn) or (ino ="" and lastno <> "") then iposit_sn = iRectposit_sn
		if sposit_name = "" or isNull(sposit_name) or (ino ="" and lastno <> "") then sposit_name = sRectposit_name
		if idepartment_id = "" or isNull(idepartment_id) or (ino ="" and lastno <> "") then idepartment_id = iRectdp_id
		if sdepartmentNameFull = "" or isNull(sdepartmentNameFull) or (ino ="" and lastno <> "") then sdepartmentNameFull = sRectdpNmFull

	'// ====================================================================
	'패턴 가져오기
	'// ====================================================================
	IF iPatternSeq <> "" THEN
		clsPayForm.Fpatternseq= ipatternseq
		clsPayForm.fnGetPayPatternData

		part_sn		= clsPayForm.Fpart_sn
		spatternname	= clsPayForm.Fpatternname

		defaultpay    	= clsPayForm.Fdefaultpay
		foodpay	    	= clsPayForm.Ffoodpay
		jobpay			= clsPayForm.Fjobpay
		inBreakTime		= clsPayForm.FinBreakTime
		iOverTime		= clsPayForm.FOverTime

		realWeekWorkDay = 0
		For intLoop = 1 To 7
			StartHour(intLoop) 		= clsPayForm.FStartHour(intLoop)
			StartMinute(intLoop)  	= clsPayForm.FStartMinute(intLoop)
			EndHour(intLoop)       	= clsPayForm.FEndHour(intLoop)
			EndMinute(intLoop)      = clsPayForm.FEndMinute(intLoop)
			BreakSHour(intLoop)     = clsPayForm.FBreakSHour(intLoop)
			BreakSMinute(intLoop)   = clsPayForm.FBreakSMinute(intLoop)
			BreakEHour(intLoop)     = clsPayForm.FBreakEHour(intLoop)
			BreakEMinute(intLoop)   = clsPayForm.FBreakEMinute(intLoop)
			DutyTime(intLoop)		= clsPayForm.FDutyTime(intLoop)
			iworktype(intLoop)		= clsPayForm.Fworktype(intLoop)

			if (iworktype(intLoop) = "1") then
				'// TODO : 한시간 미만 근무는 고려하지 않는다.(스크립트에서 처리함)
				realWeekWorkDay = realWeekWorkDay + 1
			end if
		Next

		totDutyTime  	= clsPayForm.FTotDutyTime
		totPaySum		= clsPayForm.FTotPaySum
		holidaywdtime	= clsPayForm.Fholidaywdtime
		regdate         = clsPayForm.Fregdate
		lastupdate      = clsPayForm.Flastupdate
		adminid         = clsPayForm.Fadminid
	END IF
	Set clsPayForm = nothing

 	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

	if defaultpay ="" then defaultpay =0
	if foodpay ="" then foodpay =0
	if jobpay ="" then jobpay =0
	if inBreakTime ="" then inBreakTime = 0
	if iOverTime = "" or isNull(iOverTime) then iOverTime = 0
	if totDutyTime = "" then totDutyTime = 0
	if totNightTime = "" then totNightTime = 0
	if totPaySum ="" then totPaySum =0
%>
<html>
<head>
<title>계약정보 등록</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script type="text/javascript" src="/js/jsPayCal.js"></script>
<script type="text/javascript" src="/js/datetime.js"></script>
<script type="text/javascript">
<!--
  	//계약일 포맷
  function jsDateform(obj) {
	var tmp;
	tmp = obj.value;
	tmp = tmp.replace(/\-/g, "");

	if (isNaN(tmp) == true) {
		alert("계약일은 숫자이외에 입력할 수 없습니다.");
		obj.value = "";
		obj.focus();
		return ;
	}

	if (tmp.length <8) {
		alert("년월일형태로 입력해주세요(ex:20101230)");
	//	obj.value = "";
	//	obj.focus();
		return;
	}

	obj.value = tmp.replace(/([0-9]{4})([0-9]+)([0-9]{2})/,"$1-$2-$3");


	var arrValue = obj.value.split("-");
	if(arrValue[1] < 1 || arrValue[1] > 12){
		alert("월은 1~12까지 등록가능합니다.");
		obj.focus();
		return;
	}
	if(arrValue[2] < 1 || arrValue[2] > 31){
		alert("일은 1~31까지 등록가능합니다.");
		obj.focus();
		return ;
	}

}


//폼 체크 후 submit 처리
	function jsChkform(frm){
		var dJD  = "<%=djoinday%>";
		if(frm.dSD.value ==""){
			alert("계약 시작일을 입력해주세요");
			frm.dSD.focus();
			return false;
		}

		if(frm.dSD.value < dJD ){
			alert("계약 시작일이 입사일보다 빠릅니다. 계약시작일을 다시 입력해주세요");
			frm.dSD.focus();
			return false;
		}

		if(frm.dED.value ==""){
			alert("계약 종료일을 입력해주세요");
			frm.dED.focus();
			return false;
		}

		if(frm.dED.value <= frm.dSD.value){
			alert("계약종료일이 시작일보다 빠릅니다. 계약종료일을 다시 입력해주세요");
			frm.dED.focus();
			return false;
		}

		if(!IsDigit(frm.iHP.value)){
			alert("시급은 숫자만 입력가능합니다.");
			frm.iHP.focus();
			return false;
		}

		if(!IsDigit(frm.iEP.value)){
			alert("식대는 숫자만 입력가능합니다.");
			frm.iEP.focus();
			return false;
		}

		if(frm.iEP.value*1 > 4000){
			alert("식대는 4000원을 초과할 수 없습니다.");
			frm.iEP.focus();
			return false;
		}

	   	var selWH = 0;
		if(frm.selWH1.value == "3") { selWH = selWH + 1; }
		if(frm.selWH2.value == "3") { selWH = selWH + 1; }
		if(frm.selWH3.value == "3") { selWH = selWH + 1; }
		if(frm.selWH4.value == "3") { selWH = selWH + 1; }
		if(frm.selWH5.value == "3") { selWH = selWH + 1; }
		if(frm.selWH6.value == "3") { selWH = selWH + 1; }
		if(frm.selWH7.value == "3") { selWH = selWH + 1; }

		var totDuty =document.all.totDuty.innerHTML;
		 totDuty = jsFormToTime(totDuty);

		 if(totDuty < 900 && selWH > 0){
		 alert("총근무 시간이 15시간이하일 경우 주휴일 설정은 불가능합니다.  ");
		 return false;
		 }

		 if(totDuty >= 900 && selWH == 0){
		 alert("주휴일을 설정해주세요");
		 return false;
		 }

		if( selWH > 1){
		alert("주휴일 설정은 하루만 가능합니다.");
		return false;
		}


		//계약일 폼 설정 확인
		//근무시간 폼 설정 확인
		if(confirm("계약을 설정하시겠습니까? ")){
			frm.iEP.disabled = false;
			frm.hidMode.value="modify";
			return true;
		}
		return false;
	}



	//패턴가져오기
	function jsGetPattern(){
		var winGP = window.open("pop_payform_pattern.asp?sEN=<%=sEmpno%>&ino=<%=ino%>","popGP"," width=700, height=800, scrollbars=yes");
		winGP.focus();
	}

	//지난계약서 가져오기
	function jsGetLastPay(Lno){
		location.href = "pop_payform.asp?sEN=<%=sEmpno%>&Lno="+Lno;
	}

	//신규등록
	function jsNewReg(){
		location.href = "pop_payform.asp?sEN=<%=sEmpno%>";
	}

	//선택 내용보기
	function jsViewPay(ino){
		location.href = "pop_payform.asp?sEN=<%=sEmpno%>&ino="+ino+"&page=<%=page%>";
	}

	// 페이지 이동
	function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

	//계약서 프린트
	function jsPRint(){
		var juminno = "<%=sjuminno%>";
		var userphone = "<%=suserphone%>";
		var usercell ="<%=susercell%>";
		var saddr = "<%=szipaddr&suseraddr%>";

		//if(juminno=="" ||(userphone=="" && usercell =="")||saddr==""){
		//alert("필수 사원정보가 입력되어있지 않습니다. 주민등록번호, 전화번호 또는 핸드폰번호 , 주소를 사원정보에서 입력해주세요 ");
		//return;
		//}

		if((userphone=="" && usercell =="")||saddr==""){
		alert("필수 사원정보가 입력되어있지 않습니다. 전화번호 또는 핸드폰번호 , 주소를 사원정보에서 입력해주세요 ");
		return;
		}
        alert('주민번호 뒷자리는 출력후 직접 적어 주시기 바랍니다.');
		var winCP = window.open("print_pay.asp?sEN=<%=sEmpno%>&ino=<%=ino%>","popCP"," width=850, height=800, scrollbars=yes");
		winCP.focus();
	}

	//계약서 삭제
	function jsDeletePayFrom() {
		var frm = document.frmPay;
		var ino = frm.ino.value;
		var sDt = frm.dSD.value;
		var eDt = frm.dED.value;

		var nowDt = new Date();
		nowDt = toDateString(nowDt)
		if(nowDt>eDt){
			alert("계약이 종료된 계약서는 삭제할 수 없습니다.");
			return;
		}

		if(confirm(ino+'회차 ('+sDt+'~'+eDt+') 계약서를 삭제하시겠습니까?')) {
			frm.hidMode.value="delete";
			frm.submit();
		}
	}

// 식대지원
	function jsChkFoodPay(v) {
		var frm = document.frmPay;
		if (v.checked == true) {
			frm.iEP.disabled = false;
		} else {
			frm.iEP.disabled = true;
			frm.iEP.value = 0;
		}
	}

//-->
</script>
</head>
<body leftmargin="10" topmargin="10">
	<form name="frm" method="post" action="pop_payform.asp">
		<input type="hidden" name="page" value="">
		<input type="hidden" name="sEN" value="<%=sempno%>">
		<input type="hidden" name="ino" value="<%=ino%>">
	</form>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<form name="frmPay" method="post" action="procPayform.asp" onsubmit="return jsChkform(this)">
<input type="hidden" name="hidMode" value="modify">
<input type="hidden" name="hidEN" value="<%=sempno%>">
<input type="hidden" name="hidPSN" value="<%=iposit_sn%>">
<input type="hidden" name="hidDid" value="<%=idepartment_id%>">
<input type="hidden" name="hidPSNm" value="<%=sposit_name%>">
<input type="hidden" name="hidDPNm" value="<%=sdepartmentNameFull%>">
<tr>
	<td><strong>계약직사원 계약정보 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">사번</td>
			<td bgcolor="#FFFFFF" width="180"><%=sempno%> <%IF blnstatediv ="N" THEN%><font color="red">[퇴사]</font><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">입사일</td>
			<td bgcolor="#FFFFFF"><%=formatdate(djoinday,"0000-00-00")%></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">이름</td>
			<td bgcolor="#FFFFFF"><%=susername%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">계약구분</td>
			<td bgcolor="#FFFFFF"><%=sRectposit_name%></td>

		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>회차</td>
			<td>계약구분</td>
			<td>부서</td>
			<td>계약시작일</td>
			<td>계약종료일</td>
			<td>시급(원)</td>
			<td>총급여(원)</td>
	    	</tr>
		<% if isArray(arrList) then %>
		<% for intLoop=0 to ubound(arrList,2)
				lastno = arrList(5,0)
		%>
		<tr height=30 align="center" bgcolor=<%IF Cstr(ino) = Cstr(arrList(5,intLoop)) THEN%>"<%=adminColor("green")%>"<%ELSE%>"#FFFFFF"<%END IF%>>
			<td><a href="javascript:jsViewPay('<%=arrList(5,intLoop)%>');"><%=arrList(5,intLoop)%></a></td>
			<td><a href="javascript:jsViewPay('<%=arrList(5,intLoop)%>');"><%=arrList(7,intLoop)%></a></td>
			<td><a href="javascript:jsViewPay('<%=arrList(5,intLoop)%>');"><%=arrList(9,intLoop)%></a></td>
			<td><a href="javascript:jsViewPay('<%=arrList(5,intLoop)%>');"><%=formatdate(arrList(1,intLoop),"0000-00-00")%></a></td>
			<td><a href="javascript:jsViewPay('<%=arrList(5,intLoop)%>');"><%=formatdate(arrList(2,intLoop),"0000-00-00")%></a></td>
			<td align="right"><a href="javascript:jsViewPay('<%=arrList(5,intLoop)%>');"><%=formatnumber(arrList(3,intLoop),0)%></a></td>
			<td align="right"><a href="javascript:jsViewPay('<%=arrList(5,intLoop)%>');"><%=formatnumber(arrList(4,intLoop),0)%></a></td>
		</tr>
		<% next %>
		<% else %>
		<tr>
			<td colspan="65" align="center" bgcolor="#FFFFFF">등록된 계약정보가 없습니다.</td>
		</tr>
		<% end if %>
		</table>
	</td>
</tr>
<!-- 페이지 시작 -->
<%
Dim iStartPage,iEndPage,iX,iPerCnt
iPerCnt = 10

iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

If (page mod iPerCnt) = 0 Then
	iEndPage = page
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
		    <tr valign="bottom" height="25">
		        <td valign="bottom" align="center">
		         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
				<% else %>[pre]<% end if %>
		        <%
					for ix = iStartPage  to iEndPage
						if (ix > iTotalPage) then Exit for
						if Cint(ix) = Cint(page) then
				%>
					<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
				<%		else %>
					<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
				<%
						end if
					next
				%>
		    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
				<% else %>[next]<% end if %>
		        </td>
		    </tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center"><hr width="100%"></td>
</tr>
<tr>
	<td align="right">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" class="a" >
		<tr>
			<td align="left">
					<%IF ino <>"" THEN%> <input type="button" value="계약서 프린트" onClick="jsPRint();" class="button"><%END IF%>
		  </td>
			<td align="right">
					<%IF blnstatediv ="Y" and   (C_PSMngPart or C_ADMIN_AUTH)   THEN%>
			<input type="button" value="신규등록" onClick="jsNewReg();" class="button" style="width:80px;">
			<%END IF%>
				<input type="button" value="기존 계약내용 가져오기" onClick="jsGetLastPay('<%=lastno%>');" class="button" style="width:150px;">
				<input type="button" value="패턴가져오기" onClick="jsGetPattern();" class="button"  style="width:100px;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="right"></td>
		</tr>
		</table>
	</td>
</tr>

<tr>
	<td><div style="color:red;">+ 계약형태(직위)및 부서는 사원관리에서 수정해주세요</div>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">회차</td>
			<td bgcolor="#FFFFFF"><input type="text" name="ino" value="<%=ino%>" style="border:0" readonly size="10"></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">계약형태</td>
			<td bgcolor="#FFFFFF"><%=sposit_name%></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">부서</td>
			<td bgcolor="#FFFFFF">
				<%=sdepartmentNameFull%>&nbsp;&nbsp;&nbsp;
				직무 : <%sbGetOptEventCodeValue "jobkind",jobkind,False,""%>
				근무지 : <%sbGetOptEventCodeValue "placekind",placekind,False,""%>
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">계약일</td>
			<td bgcolor="#FFFFFF">
			시작일: <input type="text" name="dSD" size="10" maxlength="10" value="<%=startdate%>" onFocusOut="jsDateform(this)"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('dSD');"  style="cursor:hand;">
			~ 종료일: <input type="text" name="dED" size="10"  value="<%=enddate%>"  maxlength="10" onFocusOut="jsDateform(this)"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('dED');"  style="cursor:hand;">
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시급</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iHP" size="10" style="text-align:right;" value="<%=defaultpay%>" onKeyUp="jsSetMonthlypay();"> 원</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">식대</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="iEP" size="10" style="text-align:right;" value="<%=foodpay%>" <% if (foodpay = 0) then %>disabled<% end if %>  onKeyUp="jsSetMonthlypay();"> 원
				&nbsp;
				<input type="checkbox" name="binEP" value="1" <% if (foodpay <> 0) then %>checked<% end if %> onClick="jsChkFoodPay(this)"> 식대지원
			</td>
		</tr>
		<input type="hidden" name="blnBT" value="">
		<!--
		* 휴게시간에 근무할 수 없다.(근로기준법)
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">휴계시간</td>
			<td bgcolor="#FFFFFF"><input type="checkbox" name="blnBT" value="1" onClick="jsSetInBreakTime();" <%IF inBreakTime THEN%>checked<%END IF%>>근무시간 포함 </td>
		</tr>
		-->
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시간외 수당</td>
			<td bgcolor="#FFFFFF"><input type="checkbox" name="blnOT" value="1"  <%IF iOverTime > 0  and iposit_sn =12 THEN%>checked<%END IF%> onClick="jsSetOverTime();" <%IF iposit_sn = 13 THEN%>disabled<%END IF%>>지급
				<span id="spanOT" style="display:<%IF  iOverTime = 0 OR  iposit_sn = 13 THEN%>none<%END IF%>;"><input type="text" size="5" maxlength="10" style="text-align:right;" name="iot" value="<%=iOverTime%>" onKeyUp="jsSetOverTimePay();"> 시간</span> </td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td><!-- 요일별 근무시간 지정 -->
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">구분</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">주휴일</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">근무시간</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">휴계시간</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">총근무시간</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">주휴시간</td>
		</tr>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>" >시작</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >종료</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >시작</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >종료</td>
		</tr>
		<%
		For intLoop = 1 To 7%>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>"><%=fnGetStringWD(intLoop)%></td>
			<td  bgcolor="#FFFFFF">
				<select name="selWH<%=intLoop%>"  onChange="jsSetWH(<%=intLoop%>);">
				<option value="1" <%IF iworktype(intLoop) ="1"  THEN%>selected<%END IF%>>근무일</option>
				<option value="2" <%IF iworktype(intLoop) ="2" THEN%>selected<%END IF%> style="color:blue">무급휴일</option>
				<option value="3" <%IF iworktype(intLoop) ="3" THEN%>selected<%END IF%> style="color:red">주휴일</option>
				<option value="4" <%IF iworktype(intLoop) ="4" THEN%>selected<%END IF%>>유급휴일</option>
				</select>
			</td>
			<td  bgcolor="#FFFFFF">
				<input type="text" name="iSH<%=intLoop%>" value="<%=StartHour(intLoop)%>" size="2" maxlength="2" style="text-align:right" <%IF iworktype(intLoop) ="3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iSH<%=intLoop%>','iSM<%=intLoop%>',2);">
				:
			 	<input type="text" name="iSM<%=intLoop%>" value="<%=StartMinute(intLoop)%>" size="2"  maxlength="2" style="text-align:right" <%IF iworktype(intLoop) ="3"  THEN%>disabled<%END IF%>  onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iSM<%=intLoop%>','iEH<%=intLoop%>',2);">
			</td>
			<td  bgcolor="#FFFFFF">
				<input type="text" name="iEH<%=intLoop%>" value="<%=EndHour(intLoop)%>" size="2"  maxlength="2" style="text-align:right" <%IF iworktype(intLoop) ="3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iEH<%=intLoop%>','iEM<%=intLoop%>',2);">
				:
			 	<input type="text" name="iEM<%=intLoop%>" value="<%=EndMinute(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) ="3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iEM<%=intLoop%>','iBSH<%=intLoop%>',2);">
			</td>
			<td  bgcolor="#FFFFFF">
				<input type="text" name="iBSH<%=intLoop%>" value="<%=BreakSHour(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) ="3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iBSH<%=intLoop%>','iBSM<%=intLoop%>',2);">
				:
			 	<input type="text" name="iBSM<%=intLoop%>" value="<%=BreakSMinute(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) ="3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iBSM<%=intLoop%>','iBEH<%=intLoop%>',2);">
			</td>
			<td  bgcolor="#FFFFFF">
				<input type="text" name="iBEH<%=intLoop%>" value="<%=BreakEHour(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) ="3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iBEH<%=intLoop%>','iBEM<%=intLoop%>',2);">
				:
			 	<input type="text" name="iBEM<%=intLoop%>" value="<%=BreakEMinute(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) ="3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);<%IF (intLoop+1)<8 THEN%>TnTabNumber('iBEM<%=intLoop%>','iSH<%=intLoop+1%>',2);<%END IF%>">
			</td>
			<td  bgcolor="#FFFFFF"><input type="text" name="iD<%=intLoop%>" size="5" value="<%=DutyTime(intLoop)%>" readonly style="border:0;" <%IF iworktype(intLoop) ="3" THEN%>disabled<%END IF%>></td>
			<td  bgcolor="#FFFFFF"><input type="text" name="iWHT<%=intLoop%>" size="5" value="<%IF iworktype(intLoop) ="3"  THEN%><%=format00(2,Fix(holidaywdtime/60))&":"&format00(2,holidaywdtime mod 60)%><%END IF%>"  style="border:0;" ></td>
				<input type="hidden" name="intd<%=intLoop%>" size="5" value="<%=NightTime(intLoop)%>">
		</tr>
		<%
		Next %>
		<tr  align="center">
			<td colspan="6" bgcolor="<%= adminColor("tabletop") %>">주간 총 근무시간</td>
			<td bgcolor="<%=adminColor("sky")%>"><span id="totDuty"><%=format00(2,Fix(totDutyTime/60))&":"&format00(2,totDutyTime mod 60)%></span></td>
			<td bgcolor="<%=adminColor("sky")%>"><span id="totWHT"><%=format00(2,Fix(holidaywdtime/60))&":"&format00(2,holidaywdtime mod 60)%></span></td>
		</tr>
 		</table>
 	</td>
 </tr>
 <tr>
 	<td>
 		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
 		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="4" align="center">월 합계</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">기본급</td>
			<td bgcolor="#FFFFFF"><input type="text" name="idp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(totDutyTime/60*avgWeek)%>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">기본근무시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totdt" value="<%=ceilValue(totDutyTime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">주휴수당</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iwhdp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(holidaywdtime/60*avgWeek)%>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">주휴시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totwhdt" value="<%=ceilValue(holidaywdtime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">식대</td>
			<td bgcolor="#FFFFFF"><input type="text" name="ifp"  size="10" style="text-align:right;" value="<%= foodpay * ceilValue(realWeekWorkDay * avgWeek) %>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">근무일수</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totd" value="<%= ceilValue(realWeekWorkDay * avgWeek) %>" size="5" style="text-align:right;border:0;" ></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시간외수당</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iotp"  size="10" style="text-align:right;" value="<%=defaultpay*iOverTime*1.5%>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시간외근무시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totot" value="<%=iOverTime%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">야간근무수당</td>
			<td bgcolor="#FFFFFF"><input type="text" name="inp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(totNightTime/60*avgWeek)*0.5%>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">야간근무시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totnt" value="<%=ceilValue(totNightTime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">휴일근무수당</td>
			<td bgcolor="#FFFFFF"><input type="text" name="ihdp"  size="10" style="text-align:right;" value="0"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">휴일근무시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="tothdt" value="0" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center"><b>월급여합계</b></td>
			<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="itotp"  size="10" style="text-align:right;"value="<%=totPaySum%>"> 원</td>
		</tr>

		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" class="a">
		<tr>
			<td width="40%">
			<%if C_PSMngPart or C_ADMIN_AUTH then%>
				<input type="button" class="button" value="삭제" style="width:50px;background-color:#FFAAAA;" onClick="jsDeletePayFrom();">
			<%end if%>
			</td>
			<td width="60%">
			<%if C_PSMngPart or C_ADMIN_AUTH then%>
				<input type="submit" class="button" value="확인" style="width:50px;">&nbsp;
			<%end if%>
			<input type="button" class="button" value="취소" onClick="self.close()" style="width:50px;">&nbsp;&nbsp;
			<input type="button"  class="button" value="재계산" onClick="jsCalDutyTime(7);">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
