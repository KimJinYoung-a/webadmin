<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCompanyOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/company/recruit_cls.asp"-->
<%
	Dim page
%>
<!-- 상단띠 시작 -->
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
	// 폼검사 및 실행
	function submitForm() {
		var form = document.frm_upload;

		if(!form.rcb_startdate.value) {
			alert("채용시작일자를 선택해주세요.");
			form.rcb_startdate.focus();
			return;
		}
		if(!form.rcb_enddate.value) {
			alert("채용마감일자를 선택해주세요.");
			form.rcb_enddate.focus();
			return;
		}
		if(!form.rcb_jobtype.value) {
			alert("채용직무를 입력해 주세요.");
			form.rcb_jobtype.focus();
			return;
		}


		if(!dateChk(form.rcb_startdate.value,form.rcb_enddate.value)) {
			alert("마감일은 시작일보다 같거나 빠를 수 없습니다.\n\n채용기간을 확인해주세요.");
			form.rcb_enddate.focus();
			return;
		}

		if(!form.rcb_subject.value) {
			alert("제목을 입력해주세요.");
			form.rcb_subject.focus();
			return;
		}

		//2017-02-16 유태욱추가(경력여부, 채용직무)
		form.rcb_career.value=0
		form.rcb_career1.value=0
		form.rcb_career2.value=0
	    var chk1 = $("#rcb_career1").is(":checked");
	    var chk2 = $("#rcb_career2").is(":checked");
	    if(chk1) $("#rcb_career1").val(1);
	    if(chk2) $("#rcb_career2").val(2);

		form.rcb_career.value = Number(form.rcb_career1.value)+Number(form.rcb_career2.value);

		if(form.rcb_career.value==0) {
			alert("경력 여부를 선택해주세요.");
			form.rcb_career.focus();
			return;
		}

		var personalchk = $("#rcb_personalchk").is(":checked");
	    if(personalchk){
	    	$("#rcb_personal").val(1);
	    }else{
	    	$("#rcb_personal").val(0);
	    }

		if(confirm("입력한 내용으로 저장하시겠습니까?")) {
			form.submit();
		} else {		
			return;
		}
	}

	function dateChk(dt1,dt2) {
		//구분자로 나누어 배열로 변환
		v0=dt1.split("-");
		v1=dt2.split("-");

		//일자에 해당하는 타임스탬프로 변환
		v0=new Date(v0[0],v0[1],v0[2]).valueOf();
		v1=new Date(v1[0],v1[1],v1[2]).valueOf();

		//일차이를 구한뒤 하루에 해당하는 값으로 곱하여, 초단위를 일단위로 변환
		cha=(v1-v0)/(1000*60*60*24);

		if(cha>0)
			return true;
		else
			return false;
	}

	function fnalways(){
		var Now = new Date();
		var Nowyear = Now.getFullYear();
		var inpuyNowyear = Nowyear+1;
		var alwayschk1 = $("#rcb_alwayschkbox").is(":checked");
	    if(alwayschk1){
	    	$("#rcb_always").val(1);
			$("#rcb_enddate").val(inpuyNowyear+'-12-31');
			$("#rcb_enddate").hide();
			$("input[name=rcb_enddate]").attr("readonly",true);
			$("#rcb_enddate_trigger").hide();
	    }else{
	    	$("#rcb_always").val(0);
			$("#rcb_enddate").val("");
			$("#rcb_enddate").show();
			$("#rcb_enddate_trigger").show();	    	
	    }
	}
	

//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<table width="780" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<form method="post" name="frm_upload" action="<%=uploadUrl%>/linkweb/company/Recruit_process.asp" onsubmit="return false" enctype="multipart/form-data">
<input type="hidden" name="retURL" value="<%=manageUrl%>/company/intro/recruit_list.asp?menupos=<%= menupos %>">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<tr height="10" valign="bottom">
	<td background="/images/tbl_blue_round_02.gif"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><b>채용공고 신규 작성</b></td>
	<td align="right">&nbsp;</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 상단띠 끝 -->
<!-- 메인 내용 시작 -->
<table width="900" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	<tr>
		<td width="70" bgcolor="#E6E6E6" align="center">기간</td>
		<td width="320" bgcolor="#FFFFFF">
			<input id="rcb_startdate" name="rcb_startdate" value="" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="rcb_startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			<input id="rcb_enddate" name="rcb_enddate" value="" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="rcb_enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "rcb_startdate", trigger    : "rcb_startdate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "rcb_enddate", trigger    : "rcb_enddate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			<input type="hidden" name="rcb_always" id="rcb_always" value=0 />
			&nbsp;&nbsp;<input type="checkbox" name="rcb_alwayschkbox" id="rcb_alwayschkbox"  onclick="fnalways();" />상시
		</td>
		<td width="70" bgcolor="#E6E6E6" align="center">상태</td>
		<td width="180" colspan="8" bgcolor="#FFFFFF">
			<select name="rcb_state">
				<option value="0" selected>일반</oprion>
				<option value="1">조기마감</oprion>
			</select>
		</td>
	</tr>

	<tr>
		<td bgcolor="#E6E6E6" align="center">채용직무</td>
		<td bgcolor="#FFFFFF"><input type="text" name="rcb_jobtype" size="80" value=""><Br>
			<p>[대표직무] MD, 오프라인, 매장, 마케팅, 서비스 기획, 개발, 디자인, 컨텐츠 제작, 경영, 인사법무, CS, 물류 </p>
			<p>두가지 이상의 직무를 같이 올릴 경우. ex) MD / 마케팅 </p>
		</td>
		<td width="70" bgcolor="#E6E6E6" align="center">경력여부</td>
		<td width="180" colspan="8" bgcolor="#FFFFFF">
			<input type="hidden" name="rcb_career" value="0" >
			신입<input type="checkbox" name="rcb_career1" id="rcb_career1" value="0" >
			경력<input type="checkbox" name="rcb_career2" id="rcb_career2" value="0" >
		</td>
	</tr>

	<tr>
		<td bgcolor="#E6E6E6" align="center">제목</td>
		<td bgcolor="#FFFFFF" colspan="10"><input type="text" name="rcb_subject" size="80" value=""></td>
	</tr>


	<tr>
		<td bgcolor="#E6E6E6" align="center">지원사이트 URL</td>
		<td bgcolor="#FFFFFF" colspan="10"><input type="text" name="rcb_recruit_url" size="80" value=""></td>
	</tr>

	<tr>
		<td bgcolor="#E6E6E6" align="center">모집부문 및 자격요건 (이미지)</td>
		<td bgcolor="#FFFFFF" colspan="10">
			<input type="file" name="uploadFile" size="50" /><br />
			<input type="file" name="uploadFile" size="50" /><br />
			<input type="file" name="uploadFile" size="50" />
			<div id="moreFiles" style="display:none;">
			<input type="file" name="uploadFile" size="50" /><br />
			<input type="file" name="uploadFile" size="50" /><br />
			<input type="file" name="uploadFile" size="50" /><br />
			<input type="file" name="uploadFile" size="50" /><br />
			<input type="file" name="uploadFile" size="50" /><br />
			<input type="file" name="uploadFile" size="50" /><br />
			<input type="file" name="uploadFile" size="50" />
			</div>
			<span onclick="$('#moreFiles').show();$(this).hide();" style="cursor:pointer;"><br />...파일 더 추가하기</span>
		</td>
	</tr>

	<tr>
		<td bgcolor="#E6E6E6" align="center">내용</td>
		<td bgcolor="#FFFFFF" colspan="10">
			<textarea cols="110" rows="20" name="rcb_content"  id="rcb_content" >
▶ 제출서류 					
- 이력서 (사진첨부, 희망연봉 필수기재)
 (개인 블로그와 인스타계정이 있는 경우, 이력서에 주소기재(선택))
- 경력기술서(상세기술)-경력자만 선택 제출 
- 자기소개서
- 포트폴리오(선택)				
					
					
▶전형방법										 
[정규직]					
- 1차 : 서류전형 (이력서, 자기소개서, 경력기술서, 포트폴리오)
- 2차 : 실무진 면접
- 3차 : 임원진 면접 (웹디자이너 : 1차 서류전형 합격자에게 과제물이 나갈 예정)
					
[계약직]					
- 1차 : 서류전형 (이력서, 자기소개서, 경력기술서)					
- 2차 : 실무면접 (업무능력 면접)					
					
					
▶접수방법					
- 이메일 접수만 가능 (insa@10x10.co.kr)					
- 메일제목: [모집분야_경력년차_홍길동]  					
  ex [패션 AMD_경력 1년_박보검] or [리테일 일산_경력 2년_유아인]  or [웹디자이너_신입_김태희] or [전산장비_경력 3년_정우성] or [스타일리스트_경력2년_공유] or [재무회계_경력1년_김고은]					
- 파일명 : 모집분야_경력년차_홍길동.zip					
- 메일제목에 모집분야 및 담당업무를 필수 기재, 미 기재시 서류전형 탈락될수 있습니다.
				
					
▶지원방식					
자유이력서 (이메일 접수)					
					
					
▶근무조건					
- 급여 : 면접후 결정

- 기본복지 
: 4대보험
: 10x10/ GS shop 임직원 할인          
: 자유로운 연차사용
: 자기개발비, 경조금, 동호회 지원
: 직원 휴양시설(제주 아파트, 양평 펜션)
: 야근식대 및 귀가비 지원
: 장기근속(3년) 휴가 및 휴가비 지급
: 장기근속(2년) 별도 건강검진 지원

- 기본업무시간 : 오전 9시~오후 6시, 주5일근무

- 근무지 : 서울시 종로구 동숭동 (4호선 혜화역 100미터 거리)

* 매장 기본 업무시간 : 매장 영업시간 내 8시간, 주5일근무
  매장 근무지 : 일산벨라시타점 - 경기도 고양시 일산동구 백석동 1237 요진벨라시타 1층 텐바이텐

(계약직의 경우, 최장 2년까지 계약 연장 가능/ 정규직 TO 발생 시, 채용면접에 합격해야만 전환 가능)					
	
				
▶합격자발표										
1차 서류전형 합격자에 한해 2차 면접 일정을 개별 통보합니다. 					
					
					
인터넷으로 제출한 서류는 일체 반환하지 않으며 채용절차의 공정화에 관한 법률의거 일정기간 경과 후 폐기합니다. 					

			</textarea>
		</td>
	</tr>

	<tr>
		<td bgcolor="#E6E6E6" align="center">개인정보 수집 및 이용 동의</td>
		<td width="180" colspan="10" bgcolor="#FFFFFF">
			<input type="hidden" name="rcb_personal" id="rcb_personal" value="0" >
			<input type="checkbox" name="rcb_personalchk" id="rcb_personalchk" >&nbsp;'개인정보 수집 및 이용 동의' 다운로드 사용(이메일로 접수 받을때 사용)
		</td>
	</tr>

</table>
<!-- 메인 내용 끝 -->
<table width="900" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
				<a href="javascript:submitForm();"><img src="/images/icon_confirm.gif" width="45" border="0" align="absbottom"></a>
				<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" width="45" border="0" align="absbottom"></a>
			</td>
		</tr>
		</table>
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
	<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</form>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCompanyClose.asp" -->