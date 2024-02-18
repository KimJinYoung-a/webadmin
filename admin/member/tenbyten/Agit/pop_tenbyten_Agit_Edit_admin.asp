<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 아지트 예약등록
' History : 2011.03.10 허진원 수정
'           2012.02.14 허진원 - 미니달력 교체
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<%
  dim empno, cMember
	dim idx, sDt, sTm, eDt, eTm
	dim AreaDiv,userid,susername,iposit_sn,ipart_sn,spart_name, sposit_name,sdepartmentNameFull
	dim suserPhone,susercell,ChkStart,ChkEnd,usePersonNo,etcComment
	dim department_id, inc_subdepartment, nodepartonly
	dim usepoint, usemoney,blnUsing
	dim maxeday
	dim regType
	
if mid(date(),6,10) >="04-01" and mid(date(),6,10) <"09-01" then 
		maxeday =dateserial(year(date()),12,31)
	elseif mid(date(),6,10) >="09-01"  then
		 maxeday =dateserial(year(dateadd("yyyy",1,date())),06,30)	
	elseif 	 mid(date(),6,10) <"04-01"  then
		 maxeday =dateserial(year(date()),06,30)	
	end if	
	idx = request("idx")
 empno = requestCheckvar(Request("sEn"),32)
 sDt		= requestCheckvar(Request("ChkStart"),10)
 regType= requestCheckvar(Request("regType"),1)
 
 '// 내용 접수
	if idx<>"" then
		dim oAgitCal
		Set oAgitCal = new CAgitCalendarDetail
		oAgitCal.read(idx)

		AreaDiv			= oAgitCal.FAreaDiv
		empno				= oAgitCal.Fempno
		userid			= oAgitCal.Fuserid
		susername		= oAgitCal.Fusername
		iposit_sn		= oAgitCal.Fposit_sn
		ipart_sn			= oAgitCal.Fpart_sn
		department_id =  oAgitCal.Fdepartment_id
		suserPhone		= oAgitCal.FuserPhone
		susercell			= oAgitCal.FuserHP
		ChkStart		= oAgitCal.FChkStart
		ChkEnd			= oAgitCal.FChkEnd
		usePersonNo		= oAgitCal.FusePersonNo
		etcComment		= oAgitCal.FetcComment
    usepoint      = oAgitCal.FusePoint
    usemoney      = oAgitCal.FuseMoney
    blnUsing				= oAgitCal.FUsing
		sDt = left(ChkStart,10)
		eDt = left(ChkEnd,10)
		sTm = Num2Str(Hour(ChkStart),2,"0","R") & ":" & Num2Str(Minute(ChkStart),2,"0","R")
		eTm = Num2Str(Hour(ChkEnd),2,"0","R") & ":" & Num2Str(Minute(ChkEnd),2,"0","R")

		Set oAgitCal = Nothing
		
	 
		if empno ="00000000000000" and userid ="admin" then
			regType ="1"
		else
			regType="2"	
		end if
	end if
	
 if regType = "" then regType ="1"
 if empno <> "" then
Set cMember = new CTenByTenMember
	cMember.Fempno = empno
	cMember.fnGetMemberData
	
	empno   		= cMember.Fempno
	userid			= cMember.Fuserid 
	susername      	= cMember.Fusername 
	suserphone     	= cMember.Fuserphone
	susercell      	= cMember.Fusercell   
	ipart_sn       	= cMember.Fpart_sn
	iposit_sn     	= cMember.Fposit_sn 
	spart_name     	= cMember.Fpart_name
	sposit_name     = cMember.Fposit_name 
	sdepartmentNameFull	= cMember.FdepartmentNameFull
Set cMember = nothing


dim clsap
dim totap, useap, reqap,payap, ispenalty, psdate, pedate, pkind, pCause, pPoint
set clsap = new CMyAgit
		clsap.FRectEmpno = empno
		clsap.FRectChkStart = sDt
		clsap.fnGetMyAgit
		totap = clsap.FtotPoint
		useap = clsap.FusePoint 
		pkind = clsap.Fpenaltykind
		psdate = clsap.Fpenaltysdate
		pedate = clsap.Fpenaltyedate
		pCause = clsap.FpenaltyCause
		pPoint = clsap.FpenaltyPoint
set clsap = nothing
else
	if regType="1" then
	empno = "00000000000000"
	userid ="admin"
	end if
end if

	

	''if AreaDiv="" then AreaDiv="1"
	if sDt="" then sDt=date
	if sTm="" then sTm="15:00" 
	if eDt="" then eDt=dateAdd("d",1,sDt)
	if eTm="" then eTm="13:00"
	if usePersonNo="" then usePersonNo=1
		
		dim chkdate
		chkdate = datediff("d",date(),sDt)
	 
%>
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
.tabArea .btnTabTop {border-bottom: 10px solid #DADADA; border-right: 10px solid transparent; height: 0; width: 100px; font-size:1px; line-height:1;}
.tabArea .btnTabBody {background: #E0E0E0; height: 18px; width: 100px; text-align:center; cursor: pointer;}

.tabArea .currentTop1 {border-bottom: 10px solid #BABAFF;}
.tabArea .currentBody1 {background: #C8C8FF;}
.tabArea .currentTop2 {border-bottom: 10px solid #FFBABA;}
.tabArea .currentBody2 {background: #FFC8C8;}
</style>
<script type="text/javascript">
// 등록 확인 및 처리
function jsSubmit()	{
	var frm = document.frm;

	if(frm.AreaDiv.value == "") {
		alert("예약할 아지트를 선택해주세요.");
		frm.AreaDiv.focus();
		return;
	}  	

	if(getDayInterval(toDate(frm.ChkStart.value), toDate('<%=date%>'))>0) {
		alert("지나간 날짜는 등록하실 수 없습니다. 날짜를 확인해주세요.");
		return;
	}

	if(getDayInterval(toDate(frm.ChkStart.value), toDate(frm.ChkEnd.value))<0) {
		alert("기간이 잘못되어 있습니다. 날짜를 확인해주세요.");
		return;
	}

	//상,하반기 입력가능 기간 체크
	if(frm.ChkEnd.value > "<%=maxeday%>") {
	<% if getEditAble then %>
		alert('[관리자 권한]예약 가능 날짜가 아닙니다.');
	<% else %>
		alert("예약 가능 날짜가 아닙니다.");
		return;
	<% end if %>
	}

	if(frm.uTerm.value>5){
		alert("이용기간은 최대 5박까지 가능합니다.");
		return;
	}

	if(frm.regType.value=="2"){
		if(frm.userPhone.value==""&&frm.userHP.value=="") {
			alert("비상연락처를 입력해주세요.");
			return;
		}

		if(frm.chkp.value!=1){
			alert("포인트/ 금액확인 버튼을 눌러주세요");
			return;
		}

		if(frm.iPoint.value==""){
			alert("포인트를 확인해주세요");
			return;
		}

		if(frm.sMoney.value==""){
			alert("금액을 확인해주세요");
			return;
		}

		if(parseInt(frm.iPoint.value)>parseInt(frm.avPoint.value)){
			alert("사용예정포인트가 사용가능 포인트보다 많습니다. 신청 불가능합니다.");
			return;
		}
	}

	if(confirm(frm.ChkStart.value +"~"+frm.ChkEnd.value +"기간에 ("+frm.uTerm.value+"박)\n등록하시겠습니까?")) {
		document.frm.submit();
	}
}

//이용 기간 확인 및 박수 자동입력
function chkTerm() {
	var frm = document.frm;
	var startday = frm.ChkStart;
	var endday = frm.ChkEnd;

	var startdate = toDate(startday.value);
	var enddate = toDate(endday.value);

	if ((startday.value == "") || (endday.value == "")) {
		alert("기간을 입력해주십시요.");
		return;
	}

	if (getDayInterval(startdate, enddate) < 0) {
		//alert("잘못된 기간입니다.");
		//return;
	}

	frm.uTerm.value = getDayInterval(startdate, enddate);
	frm.chkp.value = 0;
}

//직원 아이디 검사 및 관련내용 자동입력
function chkTenMember(sEn) {
	if(!sEn) {
		alert("검사할 사번 또는 이름을 입력해주세요.");
		frm.sEn.focus();
		return;
	}

	document.getElementById("ifmProc").src="actionTenUser.asp?sEn="+sEn+"&sDt="+ frm.ChkStart.value;
}

//삭체처리
function delBook() {
	var chkdate = "<%=chkdate%>";
	var sMsg ="";
	/*
	if (chkdate <6  && chkdate >0) {
		sMsg="투숙일 5일전~1일전 취소는 3개월간 이용이 불가능하고 신청 포인트가 차감됩니다.";
	} else if(chkdate ==0) {
		sMsg="투숙일 당일 취소는 6개월간 이용이 불가능하고 신청 포인트가 차감되며 환불이 불가합니다.";
	}
	*/

	if(confirm(sMsg+" 신청을 취소 하시겠습니까?"))	{
		frm.mode.value = "del";
		frm.submit();
	}
}

//포인트, 금액 세팅
function jsSetPoint(){
	var frm = document.frm;

	var startday = frm.ChkStart.value;
	var endday = frm.ChkEnd.value;
	var usePersonNo = frm.usePersonNo.value; 
	var empno = frm.sEn.value;
	document.getElementById("ifmProc").src="procCalPoint.asp?ChkStart="+startday+"&ChkEnd="+endday+"&usePersonNo="+usePersonNo+"&empno="+empno;
}

function jschkPoint(){
	document.frm.chkp.value =0 ;
}

function jsSetRegType(ivalue){ 
	document.frmType.regType.value= ivalue;
	document.frmType.submit();
}

function jsPopSetPenalty(){
	var p = window.open("/admin/member/agit/popRegAgitPenalty.asp?idx=<%=idx%>","popPAgit","width=600,height=500,scrollbars=yes,resize=yes");
	p.focus();
}

/* 2018-05-10; 허진원 - 팝업으로 교체
function jsSetPenalty(){
	if(confirm("관리자 패널티를 등록하시겠습니까? 등록시 이후 1년간 아지트 이용이 불가합니다.")){
		document.frm.mode.value = "pt";
		document.frm.submit();
	}
}
*/
</script>
<form name="frmType"  method="post" action="">
<input type="hidden" name="regType" value="<%=regType%>">	
</form>
<form name="frm" method="post" action="tenbyten_agit_Process.asp" >
<input type="hidden" name="mode" value="<%=chkIIF(idx="","add","modi")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="regType" value="<%=regType%>">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><b>텐바이텐 아지트 예약</b><br><hr width="100%"></td>
</tr>
<tr>
	<td> 
		<div >
			<table cellpadding="0" cellspacing="0" class="tabArea">
		<tr>
			<td class="btnTabTop <%=chkIIF(regType="1","currentTop1","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(regType="2","currentTop2","")%>">&nbsp;</td> 
		</tr>
		<tr>
			<td class="btnTabBody <%=chkIIF(regType="1","currentBody1","")%>" onclick="jsSetRegType('1')">관리자 등록</td>
			<td class="btnTabBody <%=chkIIF(regType="2","currentBody2","")%>" onclick="jsSetRegType('2')">사번 등록</td> 
		</tr>
		</table>		
		</div>
		<div id="dvEn" style="display:<%if regType="1" then%>none<%end if%>;">
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
			 
			<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>사번</b></td>
			<td> 
				<input type="text" name="sEn" size="20" class="text" value="<%=empno%>"> 
				<input type="button" value="사번확인" class="button_s" style="width:70px;text-align:center;" onclick="chkTenMember(frm.sEn.value)">				
				<font color=gray>(사번 또는 이름 입력)</font>
				<input type="hidden" name="chkCfm" value="<%=chkIIF(idx="","N","Y")%>">
				
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>직원ID</b></td>
			<td>  
				<input type="text" name="userid" size="20" class="text" value="<%=userid%>">
				<!--<input type="button" value="ID확인" class="button_s" style="width:55px;text-align:center;" onclick="chkTenMember(frm.userid.value)">
				<font color=gray>(ID 또는 이름 입력)</font>
				<input type="hidden" name="chkCfm" value="<%=chkIIF(idx="","N","Y")%>">-->
			</td>
		</tr>
		<% if C_ADMIN_AUTH or C_PSMngPart then %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>이름/직급</b></td>
			<td>  <input type="text" name="username" value="<%=susername%>" class="text">/<input type="text" class="text" name="posit_nm" value="<%=sposit_name%>"></td>
		</tr>
		<% else %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>이름</b></td>
			<td>  <input type="text" name="username" value="<%=susername%>" class="text"> <input type="hidden" name="posit_nm" value="<%=sposit_name%>"></td>
		</tr>
		<% end if %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>소속부서</b></td>
			<td><input type="text" name="department_nm" class="text" value="<%=sdepartmentNameFull %>" size="40"></td>
		</tr> 
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>비상연락처</b></td>
			<td>
				전화 <input type="text" name="userPhone" size="18" maxlength="18" class="text" value="<%=suserPhone%>"> /
				휴대폰 <input type="text" name="userHP" size="18" maxlength="18" class="text" value="<%=suserCell%>">
			</td>
		</tr>	
	</table>
</div>
<div id="dvAd" style="display:<%if regType="2" then%>none<%end if%>;">
<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
	<tr  bgcolor="#FFFFFF">
		<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>ID</b></td>			
		<td>Admin  </td>
	</tr>
</table>
</div>			
<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>아지트</b></td>
			<td>
				<select name="AreaDiv" class="select">
					<option value="">::선택::</option>
					<!--<option value="1">제주도</option>-->
					<!--<option value="2">양평</option>-->
					<option value="3">속초</option>
				</select>
				<script type="text/javascript">
					document.frm.AreaDiv.value="<%=AreaDiv%>";
				</script>
			</td>
		</tr>
		
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>이용기간</b></td>
			<td style="line-height:18px;">
				<input id="ChkStart" name="ChkStart" value="<%=sDt%>" class="text" onchange="chkTerm()" size="10" maxlength="10" <%if idx <> "" then%>readonly<%end if%>/><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkStart_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<input type="text" name="ChkSTime" size="5" maxlength="5" class="text" value="<%=sTm%>"> ~
				<input id="ChkEnd" name="ChkEnd" value="<%=eDt%>" class="text" onchange="chkTerm()" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkEnd_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<input type="text" name="ChkETime" size="5" maxlength="5" class="text" value="<%=eTm%>">
		    	<font color=gray>(<input type="text" name="uTerm" readonly class="text" value="<%=DateDiff("d",sDt,eDt)%>" style="text-align:right; width:20px; border:0px; color:gray;">박)</font>
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "ChkStart", trigger    : "ChkStart_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							chkTerm();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "ChkEnd", trigger    : "ChkEnd_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							chkTerm();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</td>
		</tr>
		<%dim iPoint, sMoney,avPoint%>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>예약인원</b></td>
			<td><input type="text" name="usePersonNo" size="4" class="text" value="<%=usePersonNo%>" onKeyPress="jschkPoint();" style="text-align:right;padding-right:3px;">명 <font color=gray>(본인포함 총인원수)</font>
				
				</td>
		</tr> 
	</table>
	<div id="dvEn2" style="display:<%if regType="1" then%>none<%end if%>;">
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
		<%if idx ="" then%>
		<tr>
			<td colspan="2" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value="포인트/ 금액 확인" onClick="jsSetPoint();">
				<input type="hidden" name="chkp" id="chkp" value="0"></td>
		</tr>
		<%end if%>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>포인트</b></td>
			<td> 
				<input type="text" name="iPoint" id="iPoint" size="4" class="text" value="<%=usepoint%>" style="text-align:right;padding-right:3px;">
					<%if idx ="" then%>
				/<input type="text" name="avPoint" id="avPoint" value="<%=totap-useap%>" style="border:0px;width:30px;" class="text" readonly>
				<font color=gray>(사용예정포인트/사용가능포인트)</font>
				<%end if%>
				<div style="color:blue;font-size:11px;padding:3px">주중 1포인트, 금,토, 공휴일 전일 , 연휴전날, 성수기 2포인트 차감</div>
				</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>금액</b></td>
			<td><input type="text" name="sMoney" id="sMoney" size="10" class="text" value="<%=useMoney%>" style="text-align:right;padding-right:3px;">원
				<span style="color:blue;font-size:11px;">1박:15,000원/ 5인이상 1만원 추가</span>
				</td>
			
		</tr>
		</table>
		</div>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">		
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>비고</b></td>
			<td><textarea name="etcComment" class="textarea" style="width:100%; height:50px;"><%=etcComment%></textarea></td>
		</tr>		
		<% if idx<>""  then %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>신청상태</b></td>
			<td><%if blnUsing ="Y" then%>신청<%else%>신청취소<%end if%></td>
		</tr> 
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>패널티</b></td>
			<td><%if pkind = "1" then%>
				 투숙일 5일전~1일전 취소, <%=psdate%>~<%=pedate%> 3개월간 이용 불가, 신청 Point차감
				 <%elseif pkind="2" then%>
				 당일 취소, <%=psdate%>~<%=pedate%> 6개월간 이용 불가, 신청 Point차감, 환불불가
				 <%elseif pkind="3" then%> 
				 No-show, <%=psdate%>~<%=pedate%> 1년간 이용 불가, 신청 Point차감, 환불불가
				 <%elseif pkind="4" then%> 
				 관리자 패널티, <%=psdate%>~<%=pedate%> 이용 불가
				 <%=chkIIF(pPoint>0," ("&pPoint&"pt 차감)","")%>
				 <%=chkIIF(pCause="" or isNull(pCause),"","<br>"&pCause)%>
				 <%else%>
				 -
				 <!--<input type="button" class="button" value="관리자 패널티 등록" onClick="jsPopSetPenalty();">//-->
				 <%end if%>
			</td>
		</tr> 
	<% end if%>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="center">	 
	<% if idx<>""  then
					if    Cstr(sDt) >= Cstr(date()) and blnUsing="Y" then %>
				<input type="button" value="신청취소" class="button" style="width:60px;text-align:center;" onclick="delBook()">
 <% 			end if
		else %>
				<input type="botton" value="신청" class="button" style="width:60px;text-align:center;" onClick="jsSubmit();">
	<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<iframe id="ifmProc" src="" width="0" height="0" frameborder="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->