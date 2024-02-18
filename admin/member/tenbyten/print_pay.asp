<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  계약직사원 계약서
' History : 2011.01.12 정윤정  생성
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
Dim cMember,clsPayForm
Dim sEmpno , ino
Dim susername, susermail, sdirect070, djoinday, blnstatediv, spart_name, sposit_name, sjob_name
Dim startdate, enddate,defaultpay ,foodpay,jobpay ,inBreakTime  , holidaywdtime	,regdate    ,lastupdate ,adminid,iposit_sn,dretireday,sjuminno,suserphone,susercell,szipcode,szipaddr,suseraddr
Dim StartHour(8), StartMinute(8), EndHour(8), EndMinute(8), BreakSHour(8), BreakSMinute(8),  BreakEHour(8), BreakEMinute(8),DutyTime(8) ,NightTime(8), iworktype(8)
Dim totDutyTime,iOverTime,iPatternSeq,part_sn,spatternname,totNightTime, iHolidayTime,avgWeek,totPaySum
Dim iTotCnt,iPageSize, iTotalPage,page
Dim arrList, intLoop, jobkind, placekind

avgWeek = 4.345238095
sEmpno =   requestCheckvar(request("sEN"),14)
ino =requestCheckvar(request("ino"),10)

'사원 요약정보 가져오기-----------------
Set cMember  = new CTenByTenMember
	cMember.Fempno		= sEmpno
	cMember.fnGetMemberData
	susername	= cMember.Fusername
	sjuminno		= cMember.Fjuminno
	suserphone	= cMember.FuserPhone
	susercell		= cMember.Fusercell
	szipcode		= cMember.Fzipcode
	szipaddr		= cMember.Fzipaddr
	suseraddr	= cMember.Fuseraddr
	djoinday	  	= cMember.Fjoinday
	blnstatediv 	= cMember.Fstatediv
	iposit_sn		= cMember.Fposit_sn
	spart_name  	= cMember.Fpart_name
	sposit_name 	= cMember.Fposit_name
	sjob_name	= cMember.Fjob_name
	dretireday		= cMember.Fretireday
Set cMember = nothing
'---------------------------------------
'사원 계약정보 가져오기-----------------
Set clsPayForm = new CPayForm
	clsPayForm.Fempno= sEmpno
	clsPayForm.Fino = ino
	clsPayForm.fnGetDefaultPayData

	startdate		= clsPayForm.Fstartdate
	enddate		= clsPayForm.Fenddate

	defaultpay    	= clsPayForm.Fdefaultpay
	foodpay	    	= clsPayForm.Ffoodpay
	jobpay		= clsPayForm.Fjobpay

	inBreakTime	= clsPayForm.FinBreakTime
	iOverTime		= clsPayForm.FOverTime

	For intLoop = 1 To 7
	StartHour(intLoop) 		= clsPayForm.FStartHour(intLoop)
	StartMinute(intLoop)  	= clsPayForm.FStartMinute(intLoop)
	EndHour(intLoop)       	= clsPayForm.FEndHour(intLoop)
	EndMinute(intLoop)       = clsPayForm.FEndMinute(intLoop)
	BreakSHour(intLoop)     	= clsPayForm.FBreakSHour(intLoop)
	BreakSMinute(intLoop)     = clsPayForm.FBreakSMinute(intLoop)
	BreakEHour(intLoop)     	= clsPayForm.FBreakEHour(intLoop)
	BreakEMinute(intLoop)     = clsPayForm.FBreakEMinute(intLoop)
	DutyTime(intLoop)		=  clsPayForm.FDutyTime(intLoop)
	iworktype(intLoop)		= clsPayForm.Fworktype(intLoop)
	Next

	totDutyTime  = clsPayForm.FTotDutyTime
	totNightTime	= clsPayForm.FtotNightTime
	totPaySum	=clsPayForm.FTotPaySum

	holidaywdtime	  = clsPayForm.Fholidaywdtime
	regdate        =clsPayForm.Fregdate
	lastupdate     =clsPayForm.Flastupdate
	adminid        =clsPayForm.Fadminid
	jobkind		= clsPayForm.Fjobkind
	placekind		= clsPayForm.Fplacekind
Set clsPayForm = nothing
'---------------------------------------
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
<!--
	 document.body.onload=function(){window.print();}
//-->
</script>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a">
<tr>
	<td align="center" style="font-family: 맑은고딕,Verdana;font-size:18px;"><strong>계 약 직 근 로 계 약 서 [ <%IF iposit_sn =13 THEN%>시 급<%ELSE%>월 급<%END IF%> ]</strong></td>
</tr>
<tr>
	<td>
		<table width="100%" border="1" cellpadding="3" cellspacing="0" align="center" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center">
			<td rowspan="2" bgcolor="#FFFFFF" valign="top"><b>사용자<br>(갑)</b></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>상호</b></td>
			<td bgcolor="#FFFFFF">㈜텐바이텐</td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>사업자번호</b></td>
			<td bgcolor="#FFFFFF">211-87-00620</td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>대표자</b></td>
			<td bgcolor="#FFFFFF">최은희</td>
		</tr>
		<tr  align="center">
			<td bgcolor="<%= adminColor("tabletop") %>"><b>주소</b></td>
			<td colspan="6" bgcolor="#FFFFFF">(03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐</td>
		</tr>
		<tr align="center">
			<td rowspan="3" bgcolor="#FFFFFF" valign="top"><b>근로자<br>(을)</b></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>성명</b></td>
			<td bgcolor="#FFFFFF"><%=susername%></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>주민등록번호</b></td>
			<td bgcolor="#FFFFFF"><%=LEFT(sjuminno,8)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><!-- eastone 수정-->
			<td bgcolor="<%= adminColor("tabletop") %>"><b>전화번호</b></td>
			<td bgcolor="#FFFFFF"><%IF susercell <> "" THEN%><%=susercell%><%ELSE%><%=suserphone%><%END IF%></td>
		</tr>
		<tr  align="center">
			<td bgcolor="<%= adminColor("tabletop") %>"><b>주소</b></td>
			<td colspan="6" bgcolor="#FFFFFF"><%=szipaddr%> <%=suseraddr%></td>
		</tr>
		<tr  align="center">
			<td bgcolor="<%= adminColor("tabletop") %>"><b>직무</b></td>
			<td colspan="2" bgcolor="#FFFFFF"><% GetEvnetKindName "jobkind", jobkind %></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>근무지</b></td>
			<td colspan="2" bgcolor="#FFFFFF"><% GetEvnetKindName "placekind", placekind %></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td  align="center">위의 ㈜텐바이텐(이하 "갑")과 근로자(이하 "을")(은)는 상호 동등한 지위에서 자유의사에 따라<br />
		다음과 같이 근로계약을 체결하고 상호 성실히 이행 및 준수할 것을 서약합니다.
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="2" cellspacing="0" align="center" class="a">
		<tr>
			<td valign="top" width="85"><b>1.계약기간</b></td>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" class="a">
				<tr>
					<td>계약시작일:</td>
					<td><b><%=year(startdate)%>년 <%=month(startdate)%>월 <%=day(startdate)%>일</b></td>
					<td>계약종료일:</td>
					<td><b><%=year(enddate)%>년 <%=month(enddate)%>월 <%=day(enddate)%>일</b></td>
				</tr>
				<tr>
					<td colspan="4">
						근로계약 만료 1개월 전 양 당사자 간의 특별한 의사가표시가 없는 한 자동해지된다.<br />
						임금 변동 사유 발생 시(최저임금 상승, 근로조건 변동 등), 갱신 계약을 진행한다.
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td><b>2.시용기간</b></td>
			<td>신규 입사자의 경우 최초 입사일로부터 3개월간은 시용기간으로 한다.</td>
		</tr>
		<tr>
			<td><b>3.직무/근무지</b></td>
			<td>갑의 인사발령에 의해 지정하는 직무와 근무지로 한다. </td>
		</tr>
		<tr>
			<td valign="top"><b>4.근로조건</b></td>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center" class="a">
				<tr>
					<td  valign="top" width="70"><b>1) 급여조건</b></td>
					<td width="500">
				<%IF iposit_sn =13 THEN%>
					시급액( <b><%=formatnumber(defaultpay,0)%></b>원) X 근무시간 의 식으로 산정하며, 주휴수당이 발생하는 경우,<br>
					이를 추가하여 지급한다.<br>
					단, 근로기준법에 의거 4주를 평균한 1주 소정근로시간이 15시간 미만인 경우에는<br>
					주휴일을 부여하지 아니한다.
				<%END IF%>
					</td>
				</tr>
				<%IF iposit_sn=12 or iposit_sn=15 THEN%>
				<tr>
					<td  colspan="2">
						<table width="100%" border="1" cellpadding="3" cellspacing="0" align="center" class="a" bgcolor=#BABABA>
						<tr align="center">
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">기본급</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">주휴수당</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">시간외수당</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">야간근무수당</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">&nbsp;</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="25%">월급여</td>
						</tr>
						<tr  align="center">
							<td  bgcolor="#FFFFFF"><%IF totDutyTime =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(defaultpay*ceilValue(totDutyTime/60*avgWeek),0)%><%END IF%></td>
							<td  bgcolor="#FFFFFF"><%IF holidaywdtime =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(defaultpay*ceilValue(holidaywdtime/60*avgWeek),0)%><%END IF%></td>
							<td  bgcolor="#FFFFFF"><%IF iOverTime =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(defaultpay*iOverTime*1.5,0)%><%END IF%></td>
							<td  bgcolor="#FFFFFF"><%IF totNightTime =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(defaultpay*ceilValue(totNightTime/60*avgWeek)*0.5,0)%><%END IF%></td>
							<td  bgcolor="#FFFFFF">&nbsp;</td>
							<td  bgcolor="#FFFFFF"><%IF totPaySum =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(totPaySum,0)%><%END IF%></td>
						</tr>
						</table>
						<% if sEmpno ="90201704030065"or sEmpno="90201702010023" or sEmpno= "90201602150020" or sEmpno ="90201602150021" or sEmpno = "90201702010024" or sEmpno = "90201702010025" or sEmpno = "90201704120084" or sEmpno = "90201702010026" then '//수기예외처리 2017.12 정유정차장 요청%>
						<p style="padding:1px">본 계약이 월 22시간의 연장근로수당이 합산된 포괄산정임금이 아님을 확인하였으며 을은 이에 동의한다.단 연장근로 등에 대해서는 별도의 가산수당을 지급한다.</p>
					 <%end if%> 
					</td>
				</tr>
				<%END IF%>
				<tr>
					<td  colspan="2"><b>2) 근로시간</b><br>
						<table width="100%" border="1" cellpadding="3" cellspacing="0" align="center" class="a" bgcolor=#BABABA>
						<tr align="center">
							<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">구분</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">근무시간</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">휴게시간</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">비고</td>
						</tr>
						<tr align="center">
							<td  bgcolor="<%= adminColor("tabletop") %>" >시작</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" >종료</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" >시작</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" >종료</td>
						</tr>
						<%
						For intLoop = 1 To 7%>
						<tr align="center" bgcolor="#FFFFFF">
							<td><%=fnGetStringWD(intLoop)%></td>
							<td><%IF StartHour(intLoop) ="00" and StartMinute(intLoop) ="00" THEN%>&nbsp;<%ELSE%><%=StartHour(intLoop)%>:<%=StartMinute(intLoop)%><%END IF%></td>
							<td><%IF EndHour(intLoop) ="00" and EndMinute(intLoop) ="00" THEN%>&nbsp;<%ELSE%><%=EndHour(intLoop)%>:<%=EndMinute(intLoop)%><%END IF%></td>
							<td><%IF BreakSHour(intLoop) ="00" and BreakSMinute(intLoop) ="00" THEN%>&nbsp;<%ELSE%><%=BreakSHour(intLoop)%>:<%=BreakSMinute(intLoop)%><%END IF%></td>
							<td><%IF BreakEHour(intLoop) ="00" and BreakEMinute(intLoop) ="00" THEN%>&nbsp;<%ELSE%><%=BreakEHour(intLoop)%> : <%=BreakEMinute(intLoop)%><%END IF%> </td>
							<td><%IF iworktype(intLoop)= "1"  THEN%>
								근무일
								<%ELSEIF  iworktype(intLoop)= "2"  THEN%>
									무급휴일
								<%ELSEIF  iworktype(intLoop)= "3"  THEN%>
									주휴일
								<%ELSEIF iworktype(intLoop)  = "4" THEN%>
						 		유급휴일
								<%END IF%>
							</td>
						</tr>
						<%  	Next %>
						</table>
					<td>
				</tr>
				<tr>
					<td valign="top"><b>3) 지급시기</b></td>
					<td><%IF iposit_sn =12 THEN%>
						매월 1일부터 말일까지의 임금을 당월 말일
						<%ELSE%>
						당월분의 임금을 익월 5일
						<%END IF%>
						에 지급하며,<br>
						제세공과금을 원천징수 한 후 지급한다.
					</td>
				</tr>
				<tr>
					<td valign="top"><b>4) 퇴직급여</b></td>
					<td>근로기준법에 따라 계속근속년수 1년에 대하여 30일분의 평균임금으로 지급한다.<br>
						당해 퇴직과 관련하여 발생하는 모든 금품은 퇴직일이 속한 달의 익월 15내에 지급받는 것에 합의한다.
					</td>
				</tr>
				<tr>
					<td valign="top"><b>5) 복무규율</b></td>
					<td>무단결근 3회 이상 발생시 회사는 해고조치 할 수 있다.기타 다른 사항은 취업규칙 및 사규에 의거한다.</td>
				</tr>
				<tr>
					<td valign="top"><b>6) 유급휴일</b></td>
					<td>근로자의 날(5/1), 주휴일(발생시)로 하고, 연차휴가는 1년이상 근무시 15일을<br>
						부여하며, 그 외 사항은 근로기준법에 따른다.
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td valign="top"><b>5. 기타조건</b></td>
			<td>본 계약서에 명시되지 않은 사항에 대해서는 근로기준법 등의 관계법령 또는 취업규칙,사규등 갑이<br>
			  별도로 정한 규정에 따른다.
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">위와 같이 근로계약을 체결함.</td>
</tr>
<tr>
	<td align="center"> <%=year(startdate)%> 년 &nbsp;&nbsp; <%=month(startdate)%> 월 &nbsp;&nbsp; <%=day(startdate)%> 일</td>
</tr>
<tr>
	<td align="center"  >
		<table width="100%" border="1" cellpadding="3" cellspacing="0" align="center" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" align="center">
			<td width="10%"><b>사용자<br>(갑)</b></td>
			<td width="40%">(주)텐바이텐 대표이사 최은희 <img src="http://scm.10x10.co.kr/images/seal1.gif" width="80" align="absmiddle"></td>
			<td width="10%"><b>근로자<br>(을)</b></td>
			<td align="right">(인)&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</html>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->