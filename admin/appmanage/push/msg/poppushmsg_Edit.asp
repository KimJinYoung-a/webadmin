<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 예약 푸시 메시지 작성
' Hieditor : 서동석 생성
'			 2017.03.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->
<%
Dim idx , listimg , state , reservedate , stitle, admcomment, targetKey, baseIdx, targetState, targetStateName, mayTargetCnt, iIsTargetActionValid
Dim viewno , textimg , worktext, pushimg, oPush , subtitle , mode, date1 , time1 , time2 , isusing, istargetMsg, makeridarr, itemidarr, keywordarr
dim bonuscouponidxarr, notclickyn, noduppDate, noduppDate2, noduppDate3, regadminid, lastadminid, regdate, lastupdate, pushcontents
dim pushimg2, pushimg3, pushimg4, pushimg5, sendranking, privateYN
	idx = requestcheckvar(request("idx"),10)

	If idx = "0" Or idx = "" Then 
		mode = "mInsert"
	Else
		mode = "mEdit"
	End If 
	
	noduppDate  =0
	noduppDate2  =0
	noduppDate3  =0
	istargetMsg =0
	iIsTargetActionValid = false
	'//db 1row

If idx <> "0" then
	set oPush = new cpush_msg_list
			 oPush.FRectIdx = idx
			
			if idx <> "" Then
				oPush.pushmsgtest_getrow()

				if oPush.FResultCount > 0 then			
					stitle			= oPush.FOneItem.fpushtitle
					subtitle		= oPush.FOneItem.fpushurl
					state			= oPush.FOneItem.fstate
					reservedate		= oPush.FOneItem.freservedate
					isusing			= oPush.FOneItem.fisusing
					pushimg			= oPush.FOneItem.fpushimg
					pushimg2		= oPush.FOneItem.fpushimg2
					pushimg3		= oPush.FOneItem.fpushimg3
					pushimg4		= oPush.FOneItem.fpushimg4
					pushimg5		= oPush.FOneItem.fpushimg5
					istargetMsg    = oPush.FOneItem.fistargetMsg
					noduppDate     = oPush.FOneItem.fnoduppDate
					noduppDate2     = oPush.FOneItem.fnoduppDate2
					noduppDate3     = oPush.FOneItem.fnoduppDate3
					targetKey      = oPush.FOneItem.ftargetKey
					admcomment     = oPush.FOneItem.fadmcomment
					baseIdx        = oPush.FOneItem.fbaseIdx
					targetState     = oPush.FOneItem.ftargetState
					targetStateName = oPush.FOneItem.getTargetStateName
					mayTargetCnt    = oPush.FOneItem.fmayTargetCnt
					iIsTargetActionValid = oPush.FOneItem.IsTargetActionValid
					privateYN    = oPush.FOneItem.fprivateYN

					if trim(oPush.FOneItem.fmakeridarr) <> "" then
						makeridarr = replace(oPush.FOneItem.fmakeridarr,"""","")
					end if
					if trim(oPush.FOneItem.fkeywordarr) <> "" then
						keywordarr = replace(oPush.FOneItem.fkeywordarr,"""","")
					end if

					itemidarr = oPush.FOneItem.fitemidarr
					bonuscouponidxarr = oPush.FOneItem.fbonuscouponidxarr
					notclickyn = oPush.FOneItem.fnotclickyn
					regadminid = oPush.FOneItem.fregadminid
					lastadminid = oPush.FOneItem.flastadminid
					regdate = oPush.FOneItem.fregdate
					lastupdate = oPush.FOneItem.flastupdate
					if oPush.FOneItem.fpushcontents<>"" then
						pushcontents = replace(oPush.FOneItem.fpushcontents,"\n",vbcrlf)
					end if
					sendranking = oPush.FOneItem.fsendranking
				end if	
			end if
	set oPush = Nothing

	date1 = Left(reservedate,10)
	time1 = Mid(FormatDateTime(reservedate,4),1,2)
	time2 = Mid(FormatDateTime(reservedate,4),4,2)
else
    stitle="제목입력하세요"
	pushcontents="(광고) 내용입력하세요"&vbcrlf&"※ 수신거부 : 마이텐바이텐 > 설정"
	'stitle="(광고) 제목입력하세요 ※ 수신거부 : 마이텐바이텐 > 설정"
End If

if sendranking="" then sendranking="6"
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

	function jsgolist(){
		opener.location.reload();
		self.close();
	}

	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
		var wImgView;
		wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
		wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsSetImg(sImg, sName, sSpan){	
		document.domain = '10x10.co.kr';

		var winImg;
		winImg = window.open('/admin/mobile/lib/pop_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	//저장
	function subcheck(){
		var frm=document.inputfrm;
        
        if ((frm.istargetMsg[1].checked)&&(frm.targetKey.value.length<1)){
            alert('타겟을 선택해주세요');
			frm.targetKey.focus();
			return;
        }

		var noduppDatecnt = 0;
		if (frm.noduppDate.checked){
			noduppDatecnt = noduppDatecnt + 1;
		}
		if (frm.noduppDate2.checked){
			noduppDatecnt = noduppDatecnt + 1;
		}
		if (frm.noduppDate3.checked){
			noduppDatecnt = noduppDatecnt + 1;
		}
		if (noduppDatecnt > 1){
			alert('금일 이전 광고 발송 대상자 제한은 하나만 체크 하셔야 합니다.');
			return;
		}

		if (frm.targetKey.value=='9100' || frm.targetKey.value=='9110' || frm.targetKey.value=='9111' || frm.targetKey.value=='9300' || frm.targetKey.value=='9310'){
			if (frm.makeridarr.value=='' && frm.itemidarr.value==''){
				alert('브랜드ID나 상품코드 둘중 하나는 반드시 입력 하셔야 합니다.');
				frm.makeridarr.focus();
				return;
			}
		}
		if (frm.targetKey.value=='9200'){
			if (frm.makeridarr.value=='' && frm.itemidarr.value=='' && frm.keywordarr.value==''){
				alert('브랜드ID,상품코드,키워드 셋중 하나는 반드시 입력 하셔야 합니다.');
				frm.makeridarr.focus();
				return;
			}
		}
		if (frm.targetKey.value=='9120'){
			if (frm.makeridarr.value==''){
				alert('브랜드ID는 반드시 입력 하셔야 합니다.');
				frm.makeridarr.focus();
				return;
			}
		}
		if (frm.targetKey.value=='9400' || frm.targetKey.value=='9410'){
			if (frm.bonuscouponidxarr.value==''){
				alert('보너스쿠폰번호는 반드시 입력 하셔야 합니다.');
				frm.bonuscouponidxarr.focus();
				return;
			}
		}

		if (frm.stitle.value==''){ 
			alert('제목을 등록해주세요.');
			frm.stitle.focus();
			return;
		}

		if (frm.pushcontents.value==''){ 
			alert('내용을 등록해주세요.');
			frm.pushcontents.focus();
			return;
		}

		if (frm.subtitle.value==''){ 
			alert('링크을 등록해주세요');
			frm.subtitle.focus();
			return;
		}

		if (frm.time1.value==''){ 
			alert('발송 시간을 등록해주세요');
			frm.time1.focus();
			return;
		}

		if (frm.time2.value==''){ 
			alert('발송 분을 등록해주세요');
			frm.time2.focus();
			return;
		}

		if(frm.state.value==''){ 
			alert("상태를 선택해주세요");
			frm.state.focus();
			return;
		}

		if(frm.pushimg5.value!=''){ 
			if(frm.pushimg.value=='' || frm.pushimg2.value=='' || frm.pushimg3.value=='' || frm.pushimg4.value==''){ 
				alert("이미지를 차례대로 입력해주세요");
				return;
			}
		}
		if(frm.pushimg4.value!=''){ 
			if(frm.pushimg.value=='' || frm.pushimg2.value=='' || frm.pushimg3.value==''){ 
				alert("이미지를 차례대로 입력해주세요");
				return;
			}
		}
		if(frm.pushimg3.value!=''){ 
			if(frm.pushimg.value=='' || frm.pushimg2.value==''){ 
				alert("이미지를 차례대로 입력해주세요");
				return;
			}
		}
		if(frm.pushimg2.value!=''){ 
			if(frm.pushimg.value==''){ 
				alert("이미지를 차례대로 입력해주세요");
				return;
			}
		}
		if(frm.sendranking.value==''){ 
			alert("발송우선순위를 선택해주세요");
			frm.sendranking.focus();
			return;
		}

		///////////////////// 요일 시간대 발송 시간 체크 ///////////////////////	//2017.03.27 한용민 생성
		var reservationdate = frm.reservationdate.value;
		var yyyy = reservationdate.substr(0, 4);
		var mm = reservationdate.substr(5, 2);
		var dd = reservationdate.substr(8, 2);
		var week = new Array('일요일','월요일','화요일','수요일','목요일','금요일','토요일')
		var rweek = week[new Date(yyyy,mm,dd).getDay()]

		var tmp_targetMsg_all_hour = false;
		var tmp_targetMsg_all_time = false;
		//비 타겟 전체 발송
        if (frm.istargetMsg[0].checked){
        	//오전 8시 ~ 오후8시
			if(rweek=='일요일' || rweek=='월요일' || rweek=='화요일' || rweek=='수요일' || rweek=='목요일' || rweek=='금요일' || rweek=='토요일'){
				if ( (frm.time1.value >= 08 && frm.time1.value <= 20) ){
					tmp_targetMsg_all_hour = true
				}
			}

			//10분 단위 발송
			if(rweek=='일요일' || rweek=='월요일' || rweek=='화요일' || rweek=='수요일' || rweek=='목요일' || rweek=='금요일' || rweek=='토요일'){
				if ( (frm.time2.value == 00 || frm.time2.value == 10 || frm.time2.value == 20 || frm.time2.value == 30 || frm.time2.value == 40 || frm.time2.value == 50) ){
					tmp_targetMsg_all_time = true
				}
			}
	        if ( !(tmp_targetMsg_all_hour && tmp_targetMsg_all_time)){
		        alert('전체 발송은 월~일 오전 8시 ~ 오후8시 10분 단위로 등록 하실수 있습니다.');

				<% if C_ADMIN_AUTH then %>
					if (!confirm('[관리자]계속 하시겠습니까?')){
						return;
					}
				<% else %>
					return;
				<% end if %>
			}
        }

		var tmp_targetMsg_multi_hour = false;
		var tmp_targetMsg_multi_time = false;
		//멀티 타겟
        if (frm.istargetMsg[1].checked){
        	//수기 타켓이 아니라면
        	if (frm.targetKey.value != '9999'){
				//오전 8시 ~ 오후8시  10분단위
				if(rweek=='일요일' || rweek=='월요일' || rweek=='화요일' || rweek=='수요일' || rweek=='목요일' || rweek=='금요일' || rweek=='토요일'){
					if ( (frm.time1.value >= 08 && frm.time1.value <= 20) ){
						tmp_targetMsg_multi_hour = true
					}
				}

        		//월~금 8시
        		//if(rweek=='일요일' || rweek=='월요일' || rweek=='화요일' || rweek=='수요일' || rweek=='목요일' || rweek=='금요일' || rweek=='토요일'){
        		//	if ( (frm.time1.value == 08 && frm.time2.value == 00) ){
        		//		tmp_targetMsg_multi_hour = true
        		//	}
        		//}

				//월~일 10시30분
    			//if ( (frm.time1.value == 10 && frm.time2.value == 30) ){
    			//	tmp_targetMsg_multi_hour = true
    			//}

				//10분 단위 발송
				if(rweek=='일요일' || rweek=='월요일' || rweek=='화요일' || rweek=='수요일' || rweek=='목요일' || rweek=='금요일' || rweek=='토요일'){
					if ( (frm.time2.value == 00 || frm.time2.value == 10 || frm.time2.value == 20 || frm.time2.value == 30 || frm.time2.value == 40 || frm.time2.value == 50) ){
						tmp_targetMsg_multi_time = true
					}
				}

		        if ( !(tmp_targetMsg_multi_hour && tmp_targetMsg_multi_time)){
			        alert('멀티 타겟 발송은 월~일 오전 8시 ~ 오후8시 10분 단위로 등록 하실수 있습니다.');

					<% if C_ADMIN_AUTH then %>
						if (!confirm('[관리자]계속 하시겠습니까?')){
						    return;
						}
				    <% else %>
						return;
				    <% end if %>
				}
        	}
        }
		///////////////////// 요일 시간대 발송 시간 체크 ///////////////////////

		//frm.target="_blank";
		frm.submit();
	}

	function chgstate(v){
		var frm = document.frmstate;
		if ( v == "I" ){
			frm.state.value = 1;
		}else{
			frm.state.value = 0;
		}

		frm.target = "FrameCKP";
        frm.submit();
	}

	function putLinkText(key) {
		var frm = document.inputfrm;
		var urllink = frm.subtitle;
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

	function chgusing(){
		var frm = document.frmdel

		frm.target = "FrameCKP";
		frm.submit();
	}

	//타켓대상
	function setComp(comp){
		var istargetMsg="";
		if (inputfrm.istargetMsg[0].checked){
			istargetMsg="0";
		}else{
			istargetMsg="1";
		}
		var targetkey=$('#inputfrm select[name="targetKey"] option:selected').val();

	    if (comp.name=="istargetMsg"){
    	    if (comp.value=="1"){
    	        document.getElementById("itargetcmt").style.display="inline";
    	    }else{
    	        document.getElementById("itargetcmt").style.display="none";
    	    }
			callreplacetagcodeajax(targetkey,istargetMsg);
    	}
    	
    	if (comp.name=="targetKey"){
    	    if (comp.value>1){
    	        document.getElementById("baseIdx").style.display="inline";

				if (comp.value=='9100' || comp.value=='9110' || comp.value=='9111' || comp.value=='9300' || comp.value=='9310'){
					document.getElementById("makeridarr").style.display="inline";
    	        	document.getElementById("itemidarr").style.display="inline";
    	        	document.getElementById("keywordarr").style.display="none";
    	        	document.getElementById("bonuscouponidxarr").style.display="none";
    	    	}else if(comp.value=='9200'){
					document.getElementById("makeridarr").style.display="inline";
    	        	document.getElementById("itemidarr").style.display="inline";
    	        	document.getElementById("keywordarr").style.display="inline";
    	        	document.getElementById("bonuscouponidxarr").style.display="none";
    	    	}else if(comp.value=='9120'){
					document.getElementById("makeridarr").style.display="inline";
    	        	document.getElementById("itemidarr").style.display="none";
    	        	document.getElementById("keywordarr").style.display="none";
    	        	document.getElementById("bonuscouponidxarr").style.display="none";
    	    	}else if(comp.value=='9400' || comp.value=='9410'){
					document.getElementById("makeridarr").style.display="none";
    	        	document.getElementById("itemidarr").style.display="none";
    	        	document.getElementById("keywordarr").style.display="none";
    	        	document.getElementById("bonuscouponidxarr").style.display="inline"; 
    	    	}else{
					document.getElementById("makeridarr").style.display="none";
		        	document.getElementById("itemidarr").style.display="none";
		        	document.getElementById("keywordarr").style.display="none";
		        	document.getElementById("bonuscouponidxarr").style.display="none";
    	    	}
				callreplacetagcodeajax(comp.value,istargetMsg);
    	    }else{
    	        document.getElementById("baseIdx").style.display="none";
    	    }
    	}
	}

	// 직접 타게팅 입력
	function csvtarget(idx, mode){
		var popwin = window.open('/admin/appmanage/push/msg/poppushmsg_file.asp?idx='+idx+'&mode='+mode+'&menupos=<%=menupos%>','addreg','width=600,height=400,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function acttarget(){
		var frm=document.inputfrm;
		if (frm.targetKey.value=='9310'){
			if (frm.reservationdate.value!='<%= date() %>'){
				alert('당일구매안한사람 타게팅은 당일만 가능 합니다.');
				return;
			}

			alert('[주의]당일구매안한사람 타게팅은\n발송 대상자가 타게팅 현재 시간에 안산사람으로 저장 됩니다.\n당일날 발송 시간에 임박해서 등록 하고 사용하세요.');
		}

	    if (!confirm('타게팅을 작성하시겠습니까?')) return;
		var frm = document.frmtarget
        frm.mode.value="target";
        
		frm.target = "FrameCKP";
		frm.submit();
	}

	function retarget(){
		var frm=document.inputfrm;
		if (frm.targetKey.value=='9310'){
			if (frm.reservationdate.value!='<%= date() %>'){
				alert('당일구매안한사람 타게팅은 당일만 가능 합니다.');
				return;
			}

			alert('[주의]당일구매안한사람 타게팅은\n발송 대상자가 타게팅 현재 시간에 안산사람으로 저장 됩니다.\n당일날 발송 시간에 임박해서 등록 하고 사용하세요.');
		}

	    if (!confirm('타게팅을 (재) 작성하시겠습니까?')) return;
		var frm = document.frmtarget
        frm.mode.value="retarget";
        
		frm.target = "FrameCKP";
		frm.submit();
	}
	
	function ABtarget(){
	    if (!confirm('타겟을 분리 작성하시겠습니까?')) return;
		var frm = document.frmtarget
        frm.mode.value="abtarget";
        
		frm.target = "FrameCKP";
		frm.submit();
	}

	// 타켓 치환코드 가져오기. 아작스
	function callreplacetagcodeajax(targetkey,istargetMsg){
		$("#replacetagcode").empty().html("");
		if (istargetMsg=='1'){
			str = $.ajax({
				type: "POST",
				url: "/admin/appmanage/push/msg/pushtargetquery_act.asp",
				data: "targetkey="+targetkey+"&mode=replacetagcode",
				dataType: "html",
				async: false
			}).responseText;
			if(str!="") {
				$("#replacetagcode").empty().html(str);
			}
		}else{
			$("#replacetagcode").empty().html("<br><br>※ 실제 고객 데이터로 치환되는코드 (제목,내용)<br><font color='red'>${CUSTOMERID},${CUSTOMERNAME},${CUSTOMERLEVELNAME}</font><br>비회원의 경우 아이디와 성함 모두 '<font color='red'>고객</font>' 으로 표시 되며, 회원등급은 '<font color='red'>비회원</font>' 으로 표시 됩니다.");
		}
	}

</script>

<form name="inputfrm" id="inputfrm" method="post" action="/admin/appmanage/push/msg/doPushmsg_proc.asp" style="margin:0px;">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mode" value="<%=mode%>">
<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="4" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>푸시메시지 등록/수정</b></font><br/><br/>
	</td>
</tr>

<% If idx <> "0" Then %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">번호</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<b><%=idx%></b>
		</td>
	</tr>
<% End If %>

<tr>
	<td width="150" align="center" bgcolor="<%= adminColor("tabletop") %>">발송일</td>
	<td colspan="3" bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			<%=date1%><input type="hidden" name="reservationdate" size=20 maxlength=10 value="<%= date1 %>"/>
   		<%ELSE%>
			<input type="text" id="termSdt" name="reservationdate" size="7" maxlength=10 value="<%= date1 %>" />
			<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "termSdt", trigger    : "ChkStart_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						//CAL_End.args.min = date;
						//CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d" <%'=chkIIF(date1<>"",", max: " & replace(date1,"-",""),"")%>
				});
			</script>
   		<%END IF%>
		예) (<%=Left(Now(),10)%>)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">발송시간</td>
	<td colspan="3" bgcolor="#FFFFFF">
   		<% DrawTimeBoxdynamic "time1", time1, "time2", time2, "", "", "", "N" %>
		<br><br>
		금일 이전 광고 발송 대상자에 대해 발송에서 제외
		&nbsp;&nbsp;
		<input type="checkbox" name="noduppDate" <%=CHKIIF(noduppDate=1,"checked","")%>> (광고 1일 1회)
		&nbsp;&nbsp;
		<input type="checkbox" name="noduppDate2" <%=CHKIIF(noduppDate2=1,"checked","")%>> (광고 1일 2회)
		&nbsp;&nbsp;
		<input type="checkbox" name="noduppDate3" <%=CHKIIF(noduppDate3=1,"checked","")%>> (광고 1일 3회)
		<br>
		<input type="checkbox" name="notclickyn" <%=CHKIIF(notclickyn="Y","checked","")%>> 금일 이전 광고 발송 대상자에 대해 클릭한 사람 발송에서 제외
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">타게팅여부</td>
	<td colspan="3" bgcolor="#FFFFFF">
	    <input type="radio" name="istargetMsg" value="0" <%=CHKIIF(istargetMsg=1,"","checked")%> onClick="setComp(this);">전체
	    &nbsp;
	    <input type="radio" name="istargetMsg" value="1" <%=CHKIIF(istargetMsg=1,"checked","")%> onClick="setComp(this);">타게팅
	    &nbsp; <font color=gray>수기 타케팅인 경우 개발팀에 타케팅 내용을 요청하셔야 발송됩니다.</font>

	    <div id="itargetcmt" <%=CHKIIF(istargetMsg=1,"","style='display:none'")%> >
		    <br>--------------------------------------------------------------------------
		    <br>타겟 대상:<% call drawSelectBoxTarget("targetKey",targetKey," onChange='setComp(this);'", "", "") %>
		    <span id="baseIdx" <%=CHKIIF(targetKey>1,"","style='display:none'")%>>
				&nbsp;타겟메인Idx:<input type="text" name="baseIdx" value="<%=baseIdx%>" size="10" maxlength="10">
			</span>
		    <span id="makeridarr" <%=CHKIIF(targetKey="9100" or targetKey="9110" or targetKey="9111" or targetKey="9200" or targetKey="9300" or targetKey="9120" or targetKey="9310",""," style='display:none'")%>>
		    	<br><br>브랜드ID:<textarea name="makeridarr" cols=40 rows=3><%= makeridarr %></textarea> EX) ithinkso,7321
		    </span>
		    <span id="itemidarr" <%=CHKIIF(targetKey="9100" or targetKey="9110"  or targetKey="9111" or targetKey="9200" or targetKey="9300" or targetKey="9310",""," style='display:none'")%>>
		    	<br><br>상품코드:<textarea name="itemidarr" cols=40 rows=3><%= itemidarr %></textarea> EX) 12334,432132
		    </span>
		    <span id="keywordarr" <%=CHKIIF(targetKey="9200",""," style='display:none'")%>>
		    	<br><br>키워드:<textarea name="keywordarr" cols=40 rows=3><%= keywordarr %></textarea> EX) 우산,책상
		    </span>
		    <span id="bonuscouponidxarr" <%=CHKIIF(targetKey="9400" or targetKey="9410",""," style='display:none'")%>>
		    	<br><br>보너스쿠폰번호:<textarea name="bonuscouponidxarr" cols=40 rows=3><%= bonuscouponidxarr %></textarea> EX) 652,671
		    </span>
		    <br>
		    <br><br><font color=gray>수기타겟인경우 아래 코멘트에 타겟 대상을 간략히 적으시기 바랍니다.</font>
		    <br>타게팅 코멘트:<input type="text" name="admcomment" value="<%=admcomment%>" size="150"/>
	    </div>
    </td>
</tr>
<tr height="60">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="stitle" value="<%= stitle %>" size="160"/>
	</td>
</tr>
<tr height="60">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">내용</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<textarea name="pushcontents" cols=100 rows=8><%= pushcontents %></textarea>
		<br><br>맨앞에 <font color="red">(광고)</font> 꼭! 넣어주세요.
		<br>맨뒤에 <font color="red">※ 수신거부 : 마이텐바이텐 > 설정</font> 꼭! 넣어주세요.
		<span id="replacetagcode"></span>
	</td>
</tr>
<!--<input type="hidden" name="pushcontents" value="<%'= pushcontents %>">-->
<tr height="100">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">링크</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="subtitle" value="<%=subtitle%>" size="160"/><br/>
		<br/>ex) 전체 주소로 입력<br>
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event')">
		이벤트 링크 : http://m.10x10.co.kr/apps/appCom/wish/web2014/event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid')">
		상품코드 링크 : http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=<font color="darkred">상품코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('etc')">
		기타 링크 : <font color="darkred">http://m.10x10.co.kr/apps/appCom/wish/web2014/기타</font></span><br>

		<% if (idx="0" or idx="") then %>
			- ELK 링크 &gaparam=push_발송번호_타겟구분 (?가 있으면 &로 추가 파라메터 시작 , ?없으면 ?로 추가 파라메터 시작) 
	    <% else %>
			- ELK 링크 &gaparam=push_<%=idx%>_<%=targetKey%> (?가 있으면 &로 추가 파라메터 시작 , ?없으면 ?로 추가 파라메터 시작) 
	    <% end if %>

		<br>- 링크주소에 .asp 반드시 넣어 주세요.
		<br>&nbsp;&nbsp;잘못된주소 : http://m.10x10.co.kr/apps/appCom/wish/web2014/brand/?gaparam=push_7284_0
		<br>&nbsp;&nbsp;정상주소 : http://m.10x10.co.kr/apps/appCom/wish/web2014/brand/index.asp?gaparam=push_7284_0
		</font>
	</td>
</tr>

<%
' 멀티타켓일경우
IF (istargetMsg=1) THEN
%>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">타겟상태</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%= targetStateName %> 
			타겟수량(<%=FormatNumber(mayTargetCnt,0)%>)

			<% if targetKey=10 or targetKey=11 then %>
				<% If state = 0 or C_ADMIN_AUTH Then %>
					&nbsp;<input type="button" onClick="csvtarget('<%= idx %>','csvtarget')" value="CSV타게팅" class="button">
				<% end if %>
			<% else %>
				<% if iIsTargetActionValid then %>
					&nbsp;<input type="button" value="타게팅" onClick="acttarget()" class="button">
				<% else %>
					<% if (C_ADMIN_AUTH) and (targetKey<>"9999") then %>
						&nbsp;<input type="button" value="관리자 - RE타게팅" onClick="retarget()" class="button">
						&nbsp;&nbsp;
						<input type="button" value="관리자 - ABTEST용분리" onClick="ABtarget()" class="button">
					<% end if %>
				<% end if %>
			<% end if %>
	    </td>
	</tr>
<% end if %>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상태</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% If idx = "0" Or idx = "" Then %>
			<input type="hidden" name="state" value="0" />작성중
		<% Else %>
			<input type="hidden" name="state" value="<%=state%>" />
			<% If state = 0 Then %>
				<strong>작성중</strong>
				&nbsp;<input type="button" value="발송예약으로 변경" onclick="<%=chkiif(isusing ="Y","chgstate('I')","alert('사용중이 아닙니다.');")%>;" class="button" />
				&nbsp;
				<% If isusing="Y" then%>
					<span style="float:right;clear:both;"><input type="button" value="사용중지" onclick="chgusing();" class="button" /></span>
				<% Else %>
					<strong>&lt;&nbsp;<font color="red">사용중이 아닙니다.</font>&nbsp;&gt;</strong>
				<% End If %>
			<% ElseIf state = 1 then %>
				<strong>발송예약</strong>
				&nbsp;<input type="button" value="작성중으로 변경" onclick="<%=chkiif(isusing ="Y","chgstate('R')","alert('사용중이 아닙니다.');")%>;" class="button" />
				&nbsp;
				<% If isusing="Y" then%>
					<span style="float:right;clear:both;"><input type="button" value="사용중지" onclick="chgusing();" class="button" /></span>
				<% Else %>
					<strong>&lt;&nbsp;<font color="red">사용중이 아닙니다.</font>&nbsp;&gt;</strong>
				<% End If %>
			<% Else %>
				<%= pushmsgstate(state) %>
			<% End If %>
		<% End If %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동시발송우선순위</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% Drawsendranking "sendranking",sendranking,"" %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지</td>
	<td bgcolor="#FFFFFF" width="40%">
		<% '(최대 1000x1000) %>
		가로사이즈 : 단말기가로폭/이미지수
		<br>세로사이즈 : 560
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지1</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg" value="<%=pushimg%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg" value="이미지등록" onClick="jsSetImg('<%= pushimg %>','pushimg','simgdiv1')" class="button"/>
		<div id="simgdiv1" style="padding: 5 5 5 5">
			<% IF pushimg <> "" THEN %>			
				<img src="<%=pushimg%>" border="0" height=100 onclick="jsImgView('<%=pushimg%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg','simgdiv1');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지2</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg2" value="<%=pushimg2%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg2" value="이미지등록" onClick="jsSetImg('<%= pushimg2 %>','pushimg2','simgdiv2')" class="button"/>
		<div id="simgdiv2" style="padding: 5 5 5 5">
			<% IF pushimg2 <> "" THEN %>			
				<img src="<%=pushimg2%>" border="0" height=100 onclick="jsImgView('<%=pushimg2%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg2','simgdiv2');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지3</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg3" value="<%=pushimg3%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg3" value="이미지등록" onClick="jsSetImg('<%= pushimg3 %>','pushimg3','simgdiv3')" class="button"/>
		<div id="simgdiv3" style="padding: 5 5 5 5">
			<% IF pushimg3 <> "" THEN %>			
				<img src="<%=pushimg3%>" border="0" height=100 onclick="jsImgView('<%=pushimg3%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg3','simgdiv3');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지4</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg4" value="<%=pushimg4%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg4" value="이미지등록" onClick="jsSetImg('<%= pushimg4 %>','pushimg4','simgdiv4')" class="button"/>
		<div id="simgdiv4" style="padding: 5 5 5 5">
			<% IF pushimg4 <> "" THEN %>			
				<img src="<%=pushimg4%>" border="0" height=100 onclick="jsImgView('<%=pushimg4%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg4','simgdiv4');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지5</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg5" value="<%=pushimg5%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg5" value="이미지등록" onClick="jsSetImg('<%= pushimg5 %>','pushimg5','simgdiv5')" class="button"/>
		<div id="simgdiv5" style="padding: 5 5 5 5">
			<% IF pushimg5 <> "" THEN %>			
				<img src="<%=pushimg5%>" border="0" height=100 onclick="jsImgView('<%=pushimg5%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg5','simgdiv5');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>

<% if mode = "mEdit" then %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">개인화푸시여부</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%= privateYN %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">최초등록</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%=regdate%><br><%=regadminid%>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">마지막수정</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%=lastupdate%><br><%=lastadminid%>
		</td>
	</tr>
<% end if %>

<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center">
	    <% If (state < 7) Then %>
			<input type="button" value=" 저장 " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
	    <% end if %>

		<input type="button" value=" 취소(main) " class="button" onclick="jsgolist();"/>
	</td>
</tr>
</table>
</form>
<form name="frmtarget" method="post" action="/admin/appmanage/push/msg/doPushmsg_proc.asp" style="margin:0px;">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="target">
</form>
<form name="frmdel" method="get" action="/admin/appmanage/push/msg/doPushmsg_proc.asp" style="margin:0px;">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="del">
</form>
<form name="frmstate" method="get" action="/admin/appmanage/push/msg/doPushmsg_proc.asp" style="margin:0px;">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="state">
</form>

<% if (application("Svr_Info")="Dev") then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="100%" height="500"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<script type="text/javascript">
	callreplacetagcodeajax('<%= targetkey %>','<%= istargetMsg %>');
</script>
<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->