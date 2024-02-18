<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메일 통계
' History : 2007.08.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
Dim omd
Dim idx,mode
	idx = requestcheckvar(getNumeric(request("idx")),10)
	mode = requestcheckvar(request("mode"),32)

If idx = "" Then idx=0

set omd = New CMailzineOne
	omd.GetMailingOne idx
%>
<link href="/css/report.css" rel="stylesheet" type="text/css">
<script type="text/javascript">

function TnMailDataReg(frm){
	if(frm.title.value == ""){
		alert("발송이름을 적어주세요");
		frm.title.focus();
	}
	else if(frm.gubun.value == ""){
		alert("발송구분을 적어주세요");
		frm.gubun.focus();
	}
	else if(frm.startdate.value == ""){
		alert("발송시작시간을 적어주세요");
		frm.startdate.focus();
	}
	else if(frm.enddate.value == ""){
		alert("발송종료시간을 적어주세요");
		frm.enddate.focus();
	}
	else if(frm.reenddate.value == ""){
		alert("재발송종료시간을 적어주세요");
		frm.reenddate.focus();
	}
	else if(frm.totalcnt.value == ""){
		alert("총대상자수를 적어주세요");
		frm.totalcnt.focus();
	}
	else if(frm.realcnt.value == ""){
		alert("실발송통수를 적어주세요");
		frm.realcnt.focus();
	}
	else if(frm.realpct.value == ""){
		alert("실발송비율을 적어주세요");
		frm.realpct.focus();
	}
	else if(frm.filteringcnt.value == ""){
		alert("필터링통수를 적어주세요");
		frm.filteringcnt.focus();
	}
	else if(frm.filteringpct.value == ""){
		alert("필터링비율을 적어주세요");
		frm.filteringpct.focus();
	}
	else if(frm.successcnt.value == ""){
		alert("성공발송통수를 적어주세요");
		frm.successcnt.focus();
	}
	else if(frm.successpct.value == ""){
		alert("성공율을 적어주세요");
		frm.successpct.focus();
	}
	else if(frm.failcnt.value == ""){
		alert("실패발송통수를 적어주세요");
		frm.failcnt.focus();
	}
	else if(frm.failpct.value == ""){
		alert("실패율을 적어주세요");
		frm.failpct.focus();
	}
	else if(frm.opencnt.value == ""){
		alert("오픈통수를 적어주세요");
		frm.opencnt.focus();
	}
	else if(frm.openpct.value == ""){
		alert("오픈율을 적어주세요");
		frm.openpct.focus();
	}
	else if(frm.noopencnt.value == ""){
		alert("미오픈통수를 적어주세요");
		frm.noopencnt.focus();
	}
	else if(frm.noopenpct.value == ""){
		alert("미오픈율을 적어주세요");
		frm.noopenpct.focus();
	}
	else{
		frm.submit();
	}
}

</script>

<form method="post" name="sform" action="/admin/report/domaildata.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<table cellpadding="0" cellspacing="0" border="0">
<tr>
	<td>
		<table width="660" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td height="2" bgcolor="C6E7EA"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="1" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>발송구분</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="gubun" size="35" class="box" type="text" value="<% = omd.fgubun %>"> ex) mailzine , mailzine_not , mailzine_event</td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>발송이름</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="title" size="65" class="box" type="text" value="<% = omd.Ftitle %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>발송시작시간</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="startdate" size="65" class="box" type="text" value="<% = omd.Fstartdate %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>발송종료시간</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="enddate" size="65" class="box" type="text" value="<% = omd.Fenddate %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>재발송종료시간</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="reenddate" size="65" class="box" type="text" value="<% = omd.Freenddate %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>총대상자수</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="totalcnt" size="30" class="box" type="text" value="<% = omd.Ftotalcnt %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>실발송통수(발송비율)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="realcnt" size="20" class="box" type="text" value="<% = omd.Frealcnt %>">&nbsp;&nbsp;<input name="realpct" size="20" class="box" type="text" value="<% = omd.Frealpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>필터링 통수(필터링 비율)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="filteringcnt" size="20" class="box" type="text" value="<% = omd.Ffilteringcnt %>">&nbsp;&nbsp;<input name="filteringpct" size="20" class="box" type="text" value="<% = omd.Ffilteringpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>성공발송 통수(성공률)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="successcnt" size="20" class="box" type="text" value="<% = omd.Fsuccesscnt %>">&nbsp;&nbsp;<input name="successpct" size="20" class="box" type="text" value="<% = omd.Fsuccesspct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>실패발송 통수(실패율)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="failcnt" size="20" class="box" type="text" value="<% = omd.Ffailcnt %>">&nbsp;&nbsp;<input name="failpct" size="20" class="box" type="text" value="<% = omd.Ffailpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>오픈 통수(오픈율)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="opencnt" size="20" class="box" type="text" value="<% = omd.Fopencnt %>">&nbsp;&nbsp;<input name="openpct" size="20" class="box" type="text" value="<% = omd.Fopenpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>미오픈 통수(미오픈율)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="noopencnt" size="20" class="box" type="text" value="<% = omd.Fnoopencnt %>">&nbsp;&nbsp;<input name="noopenpct" size="20" class="box" type="text" value="<% = omd.Fnoopenpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<!--2016-12-07 유태욱 추가-->
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>클릭 수(클릭율)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="clickcnt" size="20" class="box" type="text" value="<% = omd.Fclickcnt %>">&nbsp;&nbsp;<input name="clickpct" size="20" class="box" type="text" value="<% = omd.Fclickpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<!-- //-2016-12-07 유태욱 추가-->
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>메일러</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" >
								<%= omd.fmailergubun %>
								<Br>KEY : <%= omd.fmailer_key_maeching %>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" style="padding-bottom:1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff" align="right"><a href="javascript:TnMailDataReg(sform);">저장하기</a>&nbsp;&nbsp;&nbsp;</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<script type="text/javascript">

function autochk(){
	var arrFrm = new Array();  //찾는 정보

	arrFrm[0] = new Array() //제목 공백 없이 입력
	arrFrm[1] = new Array()	//필드

	arrFrm[0][0]	=	'발송이름';
	arrFrm[0][1]	=	'발송시작시간';
	arrFrm[0][2]	=	'발송종료시간';
	arrFrm[0][3]	=	'재발송종료시간';
	arrFrm[0][4]	=	'총대상자수';
	arrFrm[0][5]	=	'실발송통수';
	arrFrm[0][6]	=	'필터링통수';
	arrFrm[0][7]	=	'성공발송통수';
	arrFrm[0][8]	=	'실패발송통수';
	arrFrm[0][9]	=	'오픈통수';
	arrFrm[0][10]	=	'미오픈통수';

	arrFrm[1][0]	=	'title';
	arrFrm[1][1]	=	'startdate';
	arrFrm[1][2]	=	'enddate';
	arrFrm[1][3]	=	'reenddate';
	arrFrm[1][4]	=	'totalcnt';
	arrFrm[1][5]	=	'realcnt';
	arrFrm[1][6]	=	'filteringcnt';
	arrFrm[1][7]	=	'successcnt';
	arrFrm[1][8]	=	'failcnt';
	arrFrm[1][9]	=	'opencnt';
	arrFrm[1][10]	=	'noopencnt';

	var strCont = document.autofrm.testtxt.value;
	var tmpValue;

	strCont = strCont.replace(/\s{2,}/g,'\n'); 	//2칸이상의 공백을 "\n" 으로 변경
	strCont = strCont.replace(/\n/g,'/');				//위에서 변경한 "\n" 을 "/" 으로 변경

	var Wcount = strCont.length - strCont.replace(/\//g,'').length; // "/" 의 갯수 구함- 루프 돌리기 위함

	var i = 0;

	while (i < Wcount){
		tmpValue=getTmpValue(strCont);
		strCont=getStrCont(strCont);

		tmpValue		= tmpValue.replace(/\s/g,'');			//추출된 문자에서 공백 제거

		for(k=0;k<11;k++){  //tmpValue 에서 찾는 정보의 값이 있으면 적용

				if(tmpValue.indexOf(arrFrm[0][k])==0){

						tmpValue=getTmpValue(strCont);
						strCont=getStrCont(strCont);

						var frm =eval('document.sform.' + arrFrm[1][k]);

						frm.value=tmpValue.replace(/[통](\W\S*)*/,'');
						i=i+1;
				}
		}
	i=i+1;
	}
	//비율 구하기
	TnMailDataPercent();
}

// 입력된 값에서 첫 "/" 부터 문장끝까지 반환
function getStrCont(strCont){

	var index	=	strCont.indexOf('/'); 					// "/"의 위치를 찾는다
	var len		= strCont.length;									// 전체 문장의길이를 구한다
	strCont 	= strCont.substring(index+1,len);	// "/" 다음 부분 부터 문장끝까지 저장

	return strCont;
}

// 입력된 값에서 처음부터  첫 "/"까지의 문장 반환
function getTmpValue(strCont){

	var index	=	strCont.indexOf('/'); 			// "/" 의 위치를 찾는다
	tmpValue 	= strCont.substring(0,index);	// 전체 문장에서 처음부터 "/" 까지의 문자열을 추출

	return tmpValue;
}

function TnMailDataPercent(){
	//실발송 통수		:
	document.sform.realpct.value = Math.round(eval(document.sform.realcnt.value/document.sform.totalcnt.value)*10000)/100;
	//필터링 통수		:
	document.sform.filteringpct.value = Math.round(eval(document.sform.filteringcnt.value/document.sform.totalcnt.value)*10000)/100;
	//성공 발송 통수:
	document.sform.successpct.value = Math.round(eval(document.sform.successcnt.value/document.sform.totalcnt.value)*10000)/100;
	//실패 발송 통수:
	document.sform.failpct.value = Math.round(eval(document.sform.failcnt.value/document.sform.totalcnt.value)*10000)/100;
	//오픈통수			:
	document.sform.openpct.value = Math.round(eval(document.sform.opencnt.value/document.sform.totalcnt.value)*10000)/100;
	//미오픈 통수		:
	document.sform.noopenpct.value = Math.round(eval(document.sform.noopencnt.value/document.sform.totalcnt.value)*10000)/100;
}

</script>

<form name="autofrm" style="margin:0px;">
	<textarea name="testtxt" cols="40" rows="5"></textarea>
	<input type="button" value="추출" onclick="autochk();" />
</form>

<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->