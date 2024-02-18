<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 아지트 예약등록
' History : 2011.03.10 허진원 수정
'           2012.02.14 허진원 - 미니달력 교체
'           2018.03.26 허진원 - 속초 아지트 추가
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
	dim AreaDiv,userid,susername,iposit_sn,ipart_sn,spart_name, sposit_name, sdepartmentNameFull
	dim suserPhone,susercell,ChkStart,ChkEnd,usePersonNo,etcComment
	dim department_id, inc_subdepartment, nodepartonly
	dim usepoint, usemoney
	dim nowDateShort, availSdate, availEdate, nowYearShort
	
	'// 신청 가능 기간 지정
	nowDateShort = mid(date(),6,10)
	nowYearShort = year(date)
	if nowDateShort>="02-15" and nowDateShort<="04-14" then		'1차 신청
		availSdate = dateserial(nowYearShort,2,17)
		availEdate = dateserial(nowYearShort,4,30)
	elseif nowDateShort>="04-15" and nowDateShort<="06-14" then	'2차 신청
		availSdate = dateserial(nowYearShort,4,17)
		availEdate = dateserial(nowYearShort,6,30)
	elseif nowDateShort>="06-15" and nowDateShort<="08-14" then	'3차 신청
		availSdate = dateserial(nowYearShort,6,15)
		availEdate = dateserial(nowYearShort,8,31)
	elseif nowDateShort>="08-15" and nowDateShort<="10-14" then	'4차 신청
		availSdate = dateserial(nowYearShort,8,17)
		availEdate = dateserial(nowYearShort,10,31)
	elseif nowDateShort>="10-15" and nowDateShort<="12-14" then	'5차 신청
		availSdate = dateserial(nowYearShort,10,17)
		availEdate = dateserial(nowYearShort,12,31)
	elseif nowDateShort>="12-15" or nowDateShort<="02-14" then	'6차 신청(익년)
		availSdate = dateserial(chkIIF(nowDateShort<="02-14",nowYearShort-1,nowYearShort),12,12)
		availEdate = dateserial(chkIIF(nowDateShort<="02-14",nowYearShort,nowYearShort+1),2,28)
	end if

	idx = request("idx")
	empno = session("ssBctSn")
	sDt		= requestCheckvar(Request("ChkStart"),10)
 
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
    
		sDt = left(ChkStart,10)
		eDt = left(ChkEnd,10)
		sTm = Num2Str(Hour(ChkStart),2,"0","R") & ":" & Num2Str(Minute(ChkStart),2,"0","R")
		eTm = Num2Str(Hour(ChkEnd),2,"0","R") & ":" & Num2Str(Minute(ChkEnd),2,"0","R")

		Set oAgitCal = Nothing
	end if
	
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
dim totap, useap, reqap,payap, ispenalty, psdate, pedate, pkind
set clsap = new CMyAgit
		clsap.FRectEmpno = empno
		clsap.FRectChkStart = sDt
		clsap.fnGetMyAgit
		totap = clsap.FtotPoint
		useap = clsap.FusePoint 
		pkind = clsap.Fpenaltykind
		psdate = clsap.Fpenaltysdate
		pedate = clsap.Fpenaltyedate
set clsap = nothing
end if

	

	''if AreaDiv="" then AreaDiv="1"
	if sDt="" then sDt=date
	'if sTm="" then
	 sTm="15:00" 
	if eDt="" then eDt=dateAdd("d",1,sDt)
'	if eTm="" then
		 eTm="13:00"
	if usePersonNo="" then usePersonNo=1
		
		dim chkdate
		chkdate = datediff("d",date(),sDt)
	 
%>
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--
	// 등록 확인 및 처리
	function jsSubmit()	{
		var frm = document.frm;
		if(frm.AreaDiv.value == "") {
			alert("예약할 아지트를 선택해주세요.");
			frm.AreaDiv.focus();
			return ;
		}  
		
		if(frm.userPhone.value==""&&frm.userHP.value=="") {
			alert("비상연락처를 입력해주세요.");
			return;
		}

		if(getDayInterval(toDate(frm.ChkStart.value), toDate('<%=date%>'))>0) {
			alert("지나간 날짜는 등록하실 수 없습니다. 날짜를 확인해주세요.");
			return ;
		}

		if(getDayInterval(toDate(frm.ChkStart.value), toDate(frm.ChkEnd.value))<0) {
			alert("기간이 잘못되어 있습니다. 날짜를 확인해주세요.");
			return ;
		}
		
		//입력가능 기간 체크
		if(frm.ChkStart.value < "<%=availSdate%>") {
			alert("예약 가능 날짜가 아닙니다.\n※ <%=availSdate%>부터 가능");
			return;
		}
		if(frm.ChkEnd.value > "<%=availEdate%>") {
			alert("예약 가능 날짜가 아닙니다.\n※ <%=availEdate%>까지 가능");
			return;
		}

		if(frm.uTerm.value>5){
			alert("이용기간은 최대 5박까지 가능합니다.");
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
		
		if(confirm(frm.ChkStart.value +"~"+frm.ChkEnd.value +"기간에 ("+frm.uTerm.value+"박)   등록하시겠습니까?"))	{
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
	function chkTenMember(uid) {
		if(!uid) {
			alert("검사할 ID를 입력해주세요.");
			frm.userid.focus();
			return;
		}

		document.getElementById("ifmProc").src="actionTenUser.asp?uid="+uid;
	}

	//삭체처리
	function delBook() {
		var chkdate = "<%=chkdate%>";
		var sMsg ="";
		if (chkdate <6  && chkdate >0){
			 //sMsg="투숙일 5일전~1일전 취소는 3개월간 이용이 불가능하고 신청 포인트가 차감됩니다."  
		}else if(chkdate ==0){
			//sMsg="투숙일 당일 취소는 6개월간 이용이 불가능하고 신청 포인트가 차감되며 환불이 불가합니다."  
		}
		
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
		document.getElementById("ifmProc").src="procCalPoint.asp?ChkStart="+startday+"&ChkEnd="+endday+"&usePersonNo="+usePersonNo+"&empno=<%=empno%>";
	}
	
	function jschkPoint(){
		document.frm.chkp.value =0 ;
	}
	
 
//-->
</script>
 
<form name="frm" method="post" action="tenbyten_agit_Process.asp" >
<input type="hidden" name="mode" value="<%=chkIIF(idx="","add","modi")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><b>텐바이텐 아지트 예약</b><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
			<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>사번</b></td>
			<td><%=empno%>
				<input type="hidden" name="sEn" size="20" class="text" value="<%=empno%>"> 
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>직원ID</b></td>
			<td> <%=userid%>
				<input type="hidden" name="userid" size="20" class="text" value="<%=userid%>">
				<!--<input type="button" value="ID확인" class="button_s" style="width:55px;text-align:center;" onclick="chkTenMember(frm.userid.value)">
				<font color=gray>(ID 또는 이름 입력)</font>
				<input type="hidden" name="chkCfm" value="<%=chkIIF(idx="","N","Y")%>">-->
			</td>
		</tr>
		<% if C_ADMIN_AUTH or C_PSMngPart then %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>이름/직급</b></td>
			<td> <%=susername%>/<%=sposit_name%>  <input type="hidden" name="username" value="<%=susername%>"></td>
		</tr>
		<% else %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>이름</b></td>
			<td> <%=susername%> <input type="hidden" name="username" value="<%=susername%>"></td>
		</tr>
		<% end if %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>소속부서</b></td>
			<td><%=sdepartmentNameFull %></td>
		</tr> 
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>비상연락처</b></td>
			<td>
				전화 <input type="text" name="userPhone" size="18" maxlength="18" class="text" value="<%=suserPhone%>"> /
				휴대폰 <input type="text" name="userHP" size="18" maxlength="18" class="text" value="<%=suserCell%>">
			</td>
		</tr>	
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>아지트</b></td>
			<td>
				<select name="AreaDiv">
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
				<input type="text" name="ChkSTime" size="5" maxlength="5" class="text" value="<%=sTm%>" readonly > ~
				<input id="ChkEnd" name="ChkEnd" value="<%=eDt%>" class="text" onchange="chkTerm()" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkEnd_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<input type="text" name="ChkETime" size="5" maxlength="5" class="text" value="<%=eTm%>" readonly >
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
		<%if idx ="" then%>
		<tr>
			<td colspan="2" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value="포인트/ 금액 확인" onClick="jsSetPoint();">
				<input type="hidden" name="chkp" id="chkp" value="0"></td>
		</tr>
		<%end if%>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>포인트</b></td>
			<td> 
				<input type="text" name="iPoint" id="iPoint" size="4" class="text_ro" value="<%=usepoint%>" style="text-align:right;padding-right:3px;" readonly>
				<%if idx ="" then%>
				/ 
				<input type="text" name="avPoint" id="avPoint" value="<%=totap-useap%>" readonly style="border:0px;width:30px;" class="text">
				<font color=gray>(사용예정포인트/사용가능포인트)</font>
				<%end if%>
				<div style="color:blue;font-size:11px;padding:3px">주중 1포인트, 금,토, 공휴일 전일 , 연휴전날 2포인트 차감</div>
				</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>금액</b></td>
			<td><input type="text" name="sMoney" id="sMoney" size="10" class="text_ro" value="<%=useMoney%>" style="text-align:right;padding-right:3px;" readonly>원
				<span style="color:blue;font-size:11px;">1박:15,000원/ 5인이상 1만원 추가</span>
				</td>
			
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>비고</b></td>
			<td><textarea name="etcComment" class="textarea" style="width:100%; height:50px;"><%=etcComment%></textarea></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="center">		 
	<% if idx<>"" then
					if  empno = session("ssBctSn") and sDt >= date() then %>
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