<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 통합알림스케줄
' Hieditor : 2022.12.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/noti/IntegrateNotificationCls.asp" -->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
<%
dim sIdx, notiType, linkCode, startDate, endDate, reserveTime, pushIsusing, kakaoAlrimIsusing, pushtitle, pushcontents
dim pushurl, templateCode, contents, button_name, button_url_mobile, button_name2, button_url_mobile2, failed_type
dim failed_subject, failed_msg, etc_template_code, member_smsok_checkyn, member_kakaoalrimyn_checkyn, regDate
dim lastUpdate, adminUserid, lastUserid, isusing, time1, time2
dim oSchedule, mode
	sIdx = requestcheckvar(getNumeric(trim(request("sIdx"))),10)
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)

If sIdx = "0" Or sIdx = "" Or isnull(sIdx) Then 
    mode = "mInsert"
Else
    mode = "mEdit"
End If 

set oSchedule = new cNotiList
    oSchedule.FrectsIdx = sIdx

    If mode = "mEdit" then
		oSchedule.fIntegrateNotificationScheduleOne()
        if oSchedule.FResultCount > 0 then
            notiType=oSchedule.FOneItem.FnotiType
            linkCode=oSchedule.FOneItem.flinkCode
            startDate = Left(oSchedule.FOneItem.fstartDate,10)
            endDate = Left(oSchedule.FOneItem.FendDate,10)
            reserveTime=oSchedule.FOneItem.FreserveTime
			time1=left(oSchedule.FOneItem.FreserveTime,2)
			time2=right(oSchedule.FOneItem.FreserveTime,2)
            pushIsusing=oSchedule.FOneItem.fpushIsusing
            kakaoAlrimIsusing=oSchedule.FOneItem.fkakaoAlrimIsusing
            pushtitle=oSchedule.FOneItem.fpushtitle
            if oSchedule.FOneItem.Fpushcontents<>"" then
                pushcontents = replace(oSchedule.FOneItem.Fpushcontents,"\n",vbcrlf)
            end if
            pushurl=oSchedule.FOneItem.Fpushurl
            templateCode=oSchedule.FOneItem.FtemplateCode
            contents=oSchedule.FOneItem.Fcontents
            button_name=oSchedule.FOneItem.Fbutton_name
            button_url_mobile=oSchedule.FOneItem.fbutton_url_mobile
            button_name2=oSchedule.FOneItem.fbutton_name2
            button_url_mobile2=oSchedule.FOneItem.fbutton_url_mobile2
            failed_type=oSchedule.FOneItem.Ffailed_type
            failed_subject=oSchedule.FOneItem.Ffailed_subject
            failed_msg=oSchedule.FOneItem.Ffailed_msg
            etc_template_code=oSchedule.FOneItem.fetc_template_code
			' 수기템플릿
			if not(etc_template_code="" or isnull(etc_template_code)) then
				templateCode = "etc-9999"
			end if
            member_smsok_checkyn=oSchedule.FOneItem.Fmember_smsok_checkyn
            member_kakaoalrimyn_checkyn=oSchedule.FOneItem.fmember_kakaoalrimyn_checkyn
            regDate=oSchedule.FOneItem.fregDate
            lastUpdate=oSchedule.FOneItem.flastupdate
            adminUserid=oSchedule.FOneItem.fadminUserid
            lastUserid=oSchedule.FOneItem.flastUserid
            isusing=oSchedule.FOneItem.fisusing
        end if
	else
		pushtitle="제목입력하세요"
		pushcontents="(광고) 내용입력하세요"&vbcrlf&"※ 수신거부 : 마이텐바이텐 > 설정"
    end if
set oSchedule = Nothing

if notiType="" then notiType="EVENT"
if pushIsusing="" then pushIsusing="N"
if kakaoAlrimIsusing="" then kakaoAlrimIsusing="N"
if isusing="" then isusing="Y"
if mode = "mInsert" then
	if member_kakaoalrimyn_checkyn="" then member_kakaoalrimyn_checkyn="Y"
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function putLinkText(key) {
	var frm = document.inputfrm;
	var urllink = frm.pushurl;
	switch(key) {
		case 'event':
			urllink.value='http://m.10x10.co.kr/apps/appCom/wish/web2014/event/eventmain.asp?eventid=이벤트번호';
			break;
		case 'itemid':
			urllink.value='http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=상품코드';
			break;
		case 'etc':
			urllink.value='http://m.10x10.co.kr/apps/appCom/wish/web2014/';
			break;
	}
}

function chDivPushIsusing(pushIsusing){
	if (pushIsusing=="Y"){
		document.getElementById("divPushIsusing").style.display="inline";
		inputfrm.pushIsusing[0].checked = true;
	}else{
		document.getElementById("divPushIsusing").style.display="none";
		inputfrm.pushIsusing[1].checked = true;
	}
}

function chDivKakaoAlrimIsusing(kakaoAlrimIsusing){
	if (kakaoAlrimIsusing=="Y"){
		document.getElementById("divKakaoAlrimIsusing").style.display="inline";
		inputfrm.kakaoAlrimIsusing[0].checked = true;
	}else{
		document.getElementById("divKakaoAlrimIsusing").style.display="none";
		inputfrm.kakaoAlrimIsusing[1].checked = true;
	}
}

//저장
function checkSubmit(){
	var pushIsusing='';
	var kakaoAlrimIsusing='';
	var frm=document.inputfrm;
	for(var i=0;i<frm.pushIsusing.length;i++){
		if (frm.pushIsusing[i].checked==true){
			pushIsusing=frm.pushIsusing[i].value;
			break;
		}
	}
	for(var i=0;i<frm.kakaoAlrimIsusing.length;i++){
		if (frm.kakaoAlrimIsusing[i].checked==true){
			kakaoAlrimIsusing=frm.kakaoAlrimIsusing[i].value;
			break;
		}
	}
	if (pushIsusing==''){ 
		alert('푸시 사용 여부를 선택해 주세요.');
		return;
	}
	if (kakaoAlrimIsusing==''){ 
		alert('카카오 알림톡 사용 여부를 선택해 주세요.');
		return;
	}
	if (frm.notiType.value==''){ 
		alert('구분을 선택해 주세요.');
		frm.notiType.focus();
		return;
	}
	if (frm.linkCode.value==''){ 
		alert('관련코드를 등록해 주세요.');
		frm.linkCode.focus();
		return;
	}
	if (frm.startDate.value==''){ 
		alert('기간 시작일을 등록해 주세요.');
		frm.startDate.focus();
		return;
	}
	if (frm.endDate.value==''){ 
		alert('기간 종료일을 등록해 주세요.');
		frm.endDate.focus();
		return;
	}
	if (frm.time1.value==''){ 
		alert('발송시간 시간을 등록해주세요');
		frm.time1.focus();
		return;
	}
	if (frm.time2.value==''){ 
		alert('발송시간 분을 등록해주세요');
		frm.time2.focus();
		return;
	}

	var timeCheck = false;
	var reserveTime=frm.time2.value;
	if ( (reserveTime == 00 || reserveTime == 10 || reserveTime == 20 || reserveTime == 30 || reserveTime == 40 || reserveTime == 50) ){
		timeCheck = true
	}
	if ( !(timeCheck)){
		alert('발송은 10분 단위로 등록 하실수 있습니다.');

		<% if C_ADMIN_AUTH then %>
			if (!confirm('[관리자]계속 하시겠습니까?')){
				return;
			}
		<% else %>
			return;
		<% end if %>
	}
	if (frm.isusing.value==''){ 
		alert('알림사용여부를 선택해 주세요.');
		frm.isusing.focus();
		return;
	}

	if (kakaoAlrimIsusing=='Y'){
		if (frm.contents.value==''){ 
			alert('카카오 알림톡 내용을 등록해 주세요.');
			frm.contents.focus();
			return;
		}
		if (frm.template_code.value==''){
			alert('카카오톡 알림톡 템플릿코드가 지정되어 있지 않습니다.');
			frm.template_code.focus();
			return;
		}
		// 알림톡 수기템플릿
		if ($('#inputfrm select[name="template_code"] option:selected').val()=='etc-9999'){
			if(frm.etc_template_code.value==''){ 
				alert("수기템플릿코드를 입력해 주세요.");
				frm.etc_template_code.focus();
				return;
			}
		}
	}

	if (pushIsusing=='Y'){
		if (frm.pushtitle.value==''){ 
			alert('푸시제목을 등록해주세요.');
			frm.pushtitle.focus();
			return;
		}
		if (frm.pushcontents.value==''){ 
			alert('푸시내용을 등록해주세요.');
			frm.pushcontents.focus();
			return;
		}
		if (frm.pushurl.value==''){ 
			alert('푸시링크을 등록해주세요');
			frm.pushurl.focus();
			return;
		}
	}

	//frm.target="_blank";
	frm.submit();
}

//타켓대상
function setComp(comp){
	var mode='<%= mode %>';

	if (comp.name=="failed_type"){
		if (comp.value!=''){
			document.getElementById("divfailed_subject").style.display="";
			document.getElementById("divfailed_msg").style.display="";
		}else{
			document.getElementById("divfailed_subject").style.display="none";
			document.getElementById("divfailed_msg").style.display="none";
		}
	}
}

// 전체템플릿 가져오기. 아작스
function calltemplateajax(sendmethod,template_code){
	str = $.ajax({
		type: "POST",
		url: "/admin/appmanage/lms/lmstemplate_act.asp",
		data: "sendmethod="+sendmethod+"&template_code="+template_code+"&mode=templateajax",
		dataType: "html",
		async: false
	}).responseText;
	if(str!="") {
		$("#templateCode").empty().html(str);
	}
}

// 템플릿내용 가져오기. 아작스
function calltemplatecontentsajax(sendmethod,templateCode){
	// 알림톡 수기템플릿
	if (templateCode=="etc-9999"){
		document.getElementById("spanetc_template_code").style.display="";
	}else{
		document.getElementById("spanetc_template_code").style.display="none";
		$("#etc_template_code").val("");
	}

	$.ajax({
		type: "POST",
		url: "/admin/appmanage/lms/lmstemplate_act.asp",
		data: "sendmethod="+sendmethod+"&template_code="+templateCode+"&mode=templatecontentsajax",
		dataType: "text",
		async:false,
		cache:true,
		success : function(Data){
			var result = jQuery.parseJSON(Data);
			if (result.resultcode=="00"){
				inputfrm.contents.value=result.contents.replace(/!@#/gi,"\n");
				inputfrm.button_name.value=result.button_name;
				inputfrm.button_url_mobile.value=result.button_url_mobile;
				inputfrm.button_name2.value=result.button_name2;
				inputfrm.button_url_mobile2.value=result.button_url_mobile2;
				$("#failed_type").val(result.failed_type).prop("selected", true);
				inputfrm.failed_subject.value=result.failed_subject;
				inputfrm.failed_msg.value=result.failed_msg.replace(/!@#/gi,"\n");
			}
		},
		//ajax error
		error: function(err){
			alert("ERR: " + err.responseText);
			return;
		}
	});
}

// 타켓 치환코드 가져오기. 아작스
function callreplacetagcodeajaxPush(targetkey,istargetMsg){
	$("#replacetagcodePush").empty().html("");
	if (istargetMsg=='1'){
		str = $.ajax({
			type: "POST",
			url: "/admin/appmanage/push/msg/pushtargetquery_act.asp",
			data: "targetkey="+targetkey+"&mode=replacetagcode",
			dataType: "html",
			async: false
		}).responseText;
		if(str!="") {
			$("#replacetagcodePush").empty().html(str);
		}
	}else{
		$("#replacetagcode").empty().html("<br><br>※ 실제 고객 데이터로 치환되는코드 (제목,내용)<br><font color='red'>${CUSTOMERID},${CUSTOMERNAME},${CUSTOMERLEVELNAME}</font><br>비회원의 경우 아이디와 성함 모두 '<font color='red'>고객</font>' 으로 표시 되며, 회원등급은 '<font color='red'>비회원</font>' 으로 표시 됩니다.");
	}
}

// 타켓 치환코드 가져오기. 아작스
function callreplacetagcodeajax(targetkey){
	str = $.ajax({
		type: "POST",
		url: "/admin/appmanage/lms/lmstargetquery_act.asp",
		data: "targetkey="+targetkey+"&mode=replacetagcode",
		dataType: "html",
		async: false
	}).responseText;
	if(str!="") {
		$("#replacetagcode").empty().html(str);
	}
}

</script>

<form name="inputfrm" id="inputfrm" method="post" action="/admin/appmanage/noti/IntegrateNotificationScheduleProcess.asp" style="margin:0px;">
<input type="hidden" name="sIdx" value="<%= sIdx %>">
<input type="hidden" name="mode" value="<%= mode %>">
<table width="100%" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>통합알림스케줄 등록/수정</b></font>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF">
		<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<% If mode = "mEdit" then %>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				스케줄번호
			</td>
			<td bgcolor="FFFFFF" align="left">
				<%= sIdx %>
			</td>	
		</tr>
		<% end if %>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				구분
			</td>
			<td bgcolor="FFFFFF" align="left">
				<% DrawNotiType "notiType",notiType,"" %>
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				관련코드
			</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="linkCode" value="<%= linkCode %>" size=8 maxlength=10 >
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				기간
			</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" id="termstartDate" name="startDate" size="7" maxlength=10 value="<%= startDate %>" onClick="jsPopCal('startDate');"  style="cursor:pointer;" />
				<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkstartDate_trigger" onclick="return false;" />
				<script type="text/javascript">
					var CAL_Start = new Calendar({
						inputField : "termstartDate", trigger    : "ChkstartDate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							//CAL_End.args.min = date;
							//CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(startDate<>"",", max: " & replace(startDate,"-",""),"")%>
					});
				</script>
				~
				<input type="text" id="termendDate" name="endDate" size="7" maxlength=10 value="<%= endDate %>" onClick="jsPopCal('endDate');"  style="cursor:pointer;" />
				<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkendDate_trigger" onclick="return false;" />
				<script type="text/javascript">
					var CAL_Start = new Calendar({
						inputField : "termendDate", trigger    : "ChkendDate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							//CAL_End.args.min = date;
							//CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(endDate<>"",", max: " & replace(endDate,"-",""),"")%>
					});
				</script>
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				발송시간
			</td>
			<td bgcolor="FFFFFF" align="left">
				<% DrawTimeBoxdynamic "time1", time1, "time2", time2, "", "", "", "N" %>
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				알림사용여부
			</td>
			<td bgcolor="FFFFFF" align="left">
				<% drawSelectBoxisusingYN "isusing",isusing, "" %>
			</td>	
		</tr>
		<% If mode = "mEdit" then %>
			<tr>
				<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">최초등록</td>
				<td bgcolor="FFFFFF" align="left">
					<%= regDate %><br><%= adminUserid %>
				</td>
			</tr>
			<tr>
				<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">마지막수정</td>
				<td bgcolor="FFFFFF" align="left">
					<%= lastUpdate %><br><%= lastUserid %>
				</td>
			</tr>
		<% end if %>
		</table>
	</td>
</tr>
<tr align="center">
	<td width="50%" bgcolor="<%= adminColor("tabletop") %>">
		푸시 사용
		<input type="radio" value="Y" name="pushIsusing" onclick="chDivPushIsusing(this.value);">Y
		<input type="radio" value="N" name="pushIsusing" onclick="chDivPushIsusing(this.value);">N
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">
		카카오 알림톡
		<input type="radio" value="Y" name="kakaoAlrimIsusing" onclick="chDivKakaoAlrimIsusing(this.value);">Y
		<input type="radio" value="N" name="kakaoAlrimIsusing" onclick="chDivKakaoAlrimIsusing(this.value);">N
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" width="50%" valign="top">
		<div id="divPushIsusing">
		<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>">
				푸시제목
			</td>
			<td bgcolor="FFFFFF">
				<input type="text" class="text" name="pushtitle" value="<%= pushtitle %>" size="60"/>
			</td>	
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">
				푸시내용
			</td>
			<td bgcolor="FFFFFF">
				<textarea name="pushcontents" cols=50 rows=8><%= pushcontents %></textarea>
				<br><br>맨앞에 <font color="red">(광고)</font> 꼭! 넣어주세요.
				<br>맨뒤에 <font color="red">※ 수신거부 : 마이텐바이텐 > 설정</font> 꼭! 넣어주세요.
				<span id="replacetagcodePush"></span>
			</td>	
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">
				푸시링크
			</td>
			<td bgcolor="FFFFFF">
				<input type="text" class="text" name="pushurl" value="<%= pushurl %>" size="60"/><br/>
				<br/>ex) 전체 주소로 입력<br>
				<font color="#707070">
				- <span style="cursor:pointer" onClick="putLinkText('event')">
				이벤트 링크 : http://m.10x10.co.kr/apps/appCom/wish/web2014/event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('itemid')">
				상품코드 링크 : http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=<font color="darkred">상품코드</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('etc')">
				기타 링크 : <font color="darkred">http://m.10x10.co.kr/apps/appCom/wish/web2014/기타</font></span><br>

				<br>- 링크주소에 .asp 반드시 넣어 주세요.
				<br>&nbsp;&nbsp;잘못된주소 : http://m.10x10.co.kr/apps/appCom/wish/web2014/brand/?gaparam=push_7284_0
				<br>&nbsp;&nbsp;정상주소 : http://m.10x10.co.kr/apps/appCom/wish/web2014/brand/index.asp?gaparam=push_7284_0
				</font>
			</td>	
		</tr>
		</table>
		</div>
	</td>
	<td bgcolor="#FFFFFF" valign="top">
		<div id="divKakaoAlrimIsusing">
		<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>">
				발송방법
			</td>
			<td bgcolor="FFFFFF">
				카카오알림톡
				<input type="hidden" name="sendmethod" value="KAKAOALRIM">
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>">
				제외타켓
			</td>
			<td bgcolor="FFFFFF">
				<span id="exceptionmember_kakaoalrimyn_checkyn" >
					<input type="checkbox" name="member_kakaoalrimyn_checkyn" value="Y" <%=CHKIIF(member_kakaoalrimyn_checkyn="Y","checked","")%>>알림톡광고알림거부자제외
				</span>
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>">
				내용
			</td>
			<td bgcolor="FFFFFF">
				<span id="templateCode" ></span>
				<span id="spanetc_template_code" <%=CHKIIF(templateCode="etc-9999","","style='display:none'")%>>
				템플릿코드 : <input type="text" class="text" name="etc_template_code" id="etc_template_code" value="<%= etc_template_code %>" maxlength="32" size="10" /><br><br>
				</span>
				<textarea name="contents" cols=50 rows=8><%= contents %></textarea>
				<span id="replacetagcode"></span>
				<span id="template_comment" >
					<br><br>템플릿 내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.
				</span>
			</td>
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>">
				KAKAO
			</td>
			<td bgcolor="FFFFFF">
				버튼이름1 : 
				<Br><input type="text" class="text" name="button_name" value="<%= button_name %>" size="60" maxlength=64 />
				<Br>예) 확인하러 가기
				<Br>
				버튼모바일주소1 : 
				<Br><input type="text" class="text" name="button_url_mobile" value="<%= button_url_mobile %>" size="60" maxlength=64 />
				<Br>예) https://tenten.app.link/J3xFnMMFT4
				<Br><Br>
				버튼이름2 : 
				<Br><input type="text" class="text" name="button_name2" value="<%= button_name2 %>" size="60" maxlength=64 />
				<Br>
				버튼모바일주소2 : 
				<Br><input type="text" class="text" name="button_url_mobile2" value="<%= button_url_mobile2 %>" size="60" maxlength=64 />
				<Br><Br>
				실패시문자발송여부 : <% Drawfailed_type "failed_type", failed_type, " onChange='setComp(this);'" %>
				<span id="divfailed_subject" <%=CHKIIF(failed_type<>"",""," style='display:none'")%>>
					<br><br>실패시문자제목:
					<br><input type="text" class="text" name="failed_subject" value="<%= failed_subject %>" size="55" maxlength=50 />
				</span>
				<span id="divfailed_msg" <%=CHKIIF(failed_type<>"",""," style='display:none'")%>>
					<br><br>실패시문자내용:
					<br><textarea name="failed_msg" cols=50 rows=8><%= failed_msg %></textarea>
				</span>
			</td>
		</tr>
		</table>
		</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">
	    <input type="button" value=" 저장 " class="button" onclick="checkSubmit();"/>
	</td>
</tr>
</table>
</form>

<% if (application("Svr_Info")="Dev") then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="100%" height="500"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<script type="text/javascript">
	chDivPushIsusing('<%= pushIsusing %>');
	chDivKakaoAlrimIsusing('<%= kakaoAlrimIsusing %>');
	callreplacetagcodeajaxPush('999','1');
	calltemplateajax('KAKAOALRIM','<%= templateCode %>')
	callreplacetagcodeajax('999');
</script>
<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->