<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS발송관리
' Hieditor : 2020.03.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
<%
dim ridx,sendmethod,title,contents,state,testsend,isusing,reservedate,exception7dayyn,targetkey,targetstate
dim targetcnt,regadminid,lastadminid,regdate,lastupdate,repeatlmsyn,member_pushyn_checkyn, targetStateName
dim mode,i, olms, page, iIsTargetActionValid, date1, time1, time2, makeridarr, itemidarr, keywordarr, bonuscouponidxarr
dim button_name, button_url_mobile, button_name2, button_url_mobile2, failed_type, failed_subject, failed_msg, orderitemidexceptionarr
dim template_code, etc_template_code, exceptionlogin, exceptionuserlevelarr, eventcodearr, member_kakaoalrimyn_checkyn
	ridx = requestcheckvar(getNumeric(request("ridx")),10)
    menupos = requestcheckvar(getNumeric(request("menupos")),10)
    page = requestcheckvar(getNumeric(request("page")),10)

	If ridx = "0" Or ridx = "" or isnull(ridx) Then
		mode = "mInsert"
        ridx=0
	Else
		mode = "mEdit"
	End If 

	iIsTargetActionValid = false

if ridx <> 0 Then
    set olms = new clms_msg_list
        olms.FRectrIdx = ridx
        olms.lmsmsg_getrow()

        if olms.FResultCount > 0 then			
            ridx			= olms.FOneItem.fridx
            sendmethod			= olms.FOneItem.fsendmethod
            title			= olms.FOneItem.ftitle
            contents			= olms.FOneItem.fcontents
            'if contents<>"" then
            '    contents = replace(contents,"\n",vbcrlf)
            'end if
            state			= olms.FOneItem.fstate
            testsend			= olms.FOneItem.ftestsend
            isusing			= olms.FOneItem.fisusing
            reservedate			= olms.FOneItem.freservedate
            exception7dayyn			= olms.FOneItem.fexception7dayyn
            targetkey			= olms.FOneItem.ftargetkey
            targetstate			= olms.FOneItem.ftargetstate
            targetStateName = olms.FOneItem.getTargetStateName
            targetcnt			= olms.FOneItem.ftargetcnt
            regadminid			= olms.FOneItem.fregadminid
            lastadminid			= olms.FOneItem.flastadminid
            regdate			= olms.FOneItem.fregdate
            lastupdate			= olms.FOneItem.flastupdate
            repeatlmsyn			= olms.FOneItem.frepeatlmsyn
			member_pushyn_checkyn = olms.FOneItem.fmember_pushyn_checkyn
            iIsTargetActionValid = olms.FOneItem.IsTargetActionValid

			if trim(olms.FOneItem.fmakeridarr) <> "" then
				makeridarr = replace(olms.FOneItem.fmakeridarr,"""","")
			end if
			if trim(olms.FOneItem.fkeywordarr) <> "" then
				keywordarr = replace(olms.FOneItem.fkeywordarr,"""","")
			end if

			itemidarr = olms.FOneItem.fitemidarr
			bonuscouponidxarr = olms.FOneItem.fbonuscouponidxarr
			button_name = olms.FOneItem.fbutton_name
			button_url_mobile = olms.FOneItem.fbutton_url_mobile
			button_name2 = olms.FOneItem.fbutton_name2
			button_url_mobile2 = olms.FOneItem.fbutton_url_mobile2
			failed_type = olms.FOneItem.ffailed_type
			failed_subject = olms.FOneItem.ffailed_subject
			failed_msg = olms.FOneItem.ffailed_msg
			orderitemidexceptionarr = olms.FOneItem.forderitemidexceptionarr
			template_code = olms.FOneItem.ftemplate_code
			etc_template_code = olms.FOneItem.fetc_template_code
			' 수기템플릿
			if sendmethod="KAKAOALRIM" then
				if not(etc_template_code="" or isnull(etc_template_code)) then
					template_code = "etc-9999"
				end if
			end if
			exceptionlogin = olms.FOneItem.fexceptionlogin
			exceptionuserlevelarr = olms.FOneItem.fexceptionuserlevelarr
			eventcodearr = olms.FOneItem.feventcodearr
			member_kakaoalrimyn_checkyn = olms.FOneItem.fmember_kakaoalrimyn_checkyn

			date1 = Left(reservedate,10)
			time1 = Mid(FormatDateTime(reservedate,4),1,2)
			time2 = Mid(FormatDateTime(reservedate,4),4,2)
        end if
    set olms = Nothing
else
	if sendmethod="LMS" then
		title="제목입력하세요"
		contents="(광고) 내용입력하세요"&vbcrlf&"(무료수신거부) 080-851-6030"
		'stitle="(광고) 제목입력하세요 (무료수신거부) 080-851-6030"
	end if
End If
if mode = "mInsert" then
	if member_kakaoalrimyn_checkyn="" then member_kakaoalrimyn_checkyn="Y"
end if

if sendmethod="" then sendmethod="LMS"
if repeatlmsyn="" then repeatlmsyn="N"
if isusing="" then isusing="Y"
if targetcnt="" then targetcnt=0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

	//저장
	function subcheck(){
		var frm=document.inputfrm;

        if (frm.sendmethod.value.length<1){
            alert('발송방법을 선택해주세요');
			frm.sendmethod.focus();
			return;
        }
        if (frm.targetkey.value.length<1){
            alert('타겟을 선택해주세요');
			frm.targetkey.focus();
			return;
        }

		// 60 장바구니담은사람, 61 위시담은사람, 62 장바구니orWish담은사람, 65 최근6개월구매한사람, 66 당일구매안한사람
		if (frm.targetkey.value=='60' || frm.targetkey.value=='61' || frm.targetkey.value=='62' || frm.targetkey.value=='65' || frm.targetkey.value=='66'){
			if (frm.makeridarr.value=='' && frm.itemidarr.value==''){
				alert('브랜드ID나 상품코드 둘중 하나는 반드시 입력 하셔야 합니다.');
				frm.makeridarr.focus();
				return;
			}
		}
		// 63 브랜드를찜한사람
		if (frm.targetkey.value=='63'){
			if (frm.makeridarr.value==''){
				alert('브랜드ID는 반드시 입력 하셔야 합니다.');
				frm.makeridarr.focus();
				return;
			}
		}
		// 64 클릭이나키워드검색한사람
		if (frm.targetkey.value=='64'){
			if (frm.makeridarr.value=='' && frm.itemidarr.value=='' && frm.keywordarr.value==''){
				alert('브랜드ID,상품코드,키워드 셋중 하나는 반드시 입력 하셔야 합니다.');
				frm.makeridarr.focus();
				return;
			}
		}
		// 67 보너스쿠폰다운후모든쿠폰안쓴사람, 68 보너스쿠폰다운안받은사람
		if (frm.targetkey.value=='67' || frm.targetkey.value=='68'){
			if (frm.bonuscouponidxarr.value==''){
				alert('보너스쿠폰번호는 반드시 입력 하셔야 합니다.');
				frm.bonuscouponidxarr.focus();
				return;
			}
		}
		// 67 마일리지소멸예정잔액
		if (frm.targetkey.value=='69'){
			if (frm.eventcodearr.value==''){
				alert('이벤트번호는 반드시 입력 하셔야 합니다.');
				frm.eventcodearr.focus();
				return;
			}
		}

		if (frm.sendmethod.value=='LMS'){
			if (frm.title.value==''){ 
				alert('제목을 등록해 주세요.');
				frm.title.focus();
				return;
			}
			if (GetByteLength(frm.title.value) > 120){
				alert("제목이 제한길이를 초과하였습니다. 120자 까지 작성 가능합니다.");
				frm.title.focus();
				return;
			}	
		}
		if (frm.contents.value==''){ 
			alert('내용을 등록해 주세요.');
			frm.contents.focus();
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

        if (frm.sendmethod.value=='KAKAOFRIEND' || frm.sendmethod.value=='KAKAOALRIM'){
			//if (frm.button_name.value.length<1){
			//	alert('카카오톡 버튼 이름을 입력해 주세요.');
			//	frm.button_name.focus();
			//	return;
			//}
			//if (frm.button_url_mobile.value.length<1){
			//	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');
			//	frm.button_url_mobile.focus();
			//	return;
			//}
			if (frm.failed_type.value=='LMS'){
				if (frm.failed_subject.value==''){
					alert('카카오톡 실패시 문자제목를 입력해 주세요.');
					frm.failed_subject.focus();
					return;
				}
				if (GetByteLength(frm.failed_subject.value) > 50){
					alert("카카오톡 실패시 문자제목이 제한길이를 초과하였습니다. 50자 까지 작성 가능합니다.");
					frm.failed_subject.focus();
					return;
				}
				if (frm.failed_msg.value==''){
					alert('카카오톡 실패시 문자내용을 입력해 주세요.');
					frm.failed_msg.focus();
					return;
				}
			}
			if (frm.sendmethod.value=='KAKAOALRIM'){
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
        }

		///////////////////// 요일 시간대 발송 시간 체크 ///////////////////////	//2017.03.27 한용민 생성
		var reservationdate = frm.reservationdate.value;
		var yyyy = reservationdate.substr(0, 4);
		var mm = reservationdate.substr(5, 2);
		var dd = reservationdate.substr(8, 2);
		var week = new Array('일요일','월요일','화요일','수요일','목요일','금요일','토요일')
		var rweek = week[new Date(yyyy,mm,dd).getDay()]

		var tmp_targetMsg_0 = false;
		//오전 8시 ~ 오후8시 10분단위
		if ( (frm.time1.value >= 08 && frm.time1.value <= 20) ){
			tmp_targetMsg_0 = true
		}

		if ( !tmp_targetMsg_0 ){
			alert('발송은 오전 8시 ~ 오후8시 \n\n로 등록 하실수 있습니다.');
			return;
		}

		var tmp_targetMsg_1 = false;
		//수기 타켓이 아니라면
		if (frm.targetkey.value != '9999'){
			//10분 단위 발송
			if(rweek=='일요일' || rweek=='월요일' || rweek=='화요일' || rweek=='수요일' || rweek=='목요일' || rweek=='금요일' || rweek=='토요일'){
				if ( (frm.time2.value == 00 || frm.time2.value == 10 || frm.time2.value == 20 || frm.time2.value == 30 || frm.time2.value == 40 || frm.time2.value == 50) ){
					tmp_targetMsg_1 = true
				}
			}

			//월~일 10시30분
			//if ( (frm.time1.value == 10 && frm.time2.value == 30) ){
			//	tmp_targetMsg_1 = true
			//}

			if ( !tmp_targetMsg_1 ){
				alert('발송은 월~일 아침8시~저녁8시 까지 10분 단위로 등록 하실수 있습니다.');
				
				<% '' 2017/05/10 추가  %>
				<% if (C_ADMIN_AUTH or session("ssBctID")="fotoark") then %>
				if (!confirm('[관리자]계속 하시겠습니까?')){
					return;
				}
				<% else %>
				return;
				<% end if %>
			}
		}
		///////////////////// 요일 시간대 발송 시간 체크 ///////////////////////

		//frm.target="_blank";
		frm.submit();
	}

	function chgstate(v){
		if ( v == "I" ){
			frmstate.state.value = 1;
			if (inputfrm.targetcnt.value<1){
				alert("먼저 타켓팅을 해주세요.");
				return;
			}
		}else{
			frmstate.state.value = 0;
		}

		frmstate.target = "FrameCKP";
        frmstate.submit();
	}

	function chgusing(){
		var frm = document.frmdel

		frm.target = "FrameCKP";
		frm.submit();
	}

	//타켓대상
	function setComp(comp){
		var mode='<%= mode %>';
    	if (comp.name=="targetkey"){
    	    if (comp.value>1){
				// 타켓 CSV타켓(휴대폰번호):1 , CSV타켓(텐바이텐고객번호):2 , CSV타켓(텐바이텐고객아이디):3
    	        if (comp.value=='1' || comp.value=='2' || comp.value=='3'){
					if (comp.value=='1'){
						document.getElementById("member_pushyn_checkyn").style.display="none";
						document.getElementById("exceptionlogin").style.display="none";
						document.getElementById("exceptionuserlevelarr").style.display="none";
						document.getElementById("replacetagcode").style.display="none";
						document.getElementById("exceptionmember_kakaoalrimyn_checkyn").style.display="";
					}else{
						document.getElementById("member_pushyn_checkyn").style.display="inline";
						document.getElementById("exceptionlogin").style.display="inline";
						document.getElementById("exceptionuserlevelarr").style.display="inline";
						document.getElementById("replacetagcode").style.display="inline";
						if (inputfrm.sendmethod.value=="KAKAOALRIM"){
							document.getElementById("exceptionmember_kakaoalrimyn_checkyn").style.display="inline";
						}
						callreplacetagcodeajax(comp.value);
					}

					document.getElementById("makeridarr").style.display="none";
					document.getElementById("itemidarr").style.display="none";
					document.getElementById("keywordarr").style.display="none";
					document.getElementById("bonuscouponidxarr").style.display="none";
					document.getElementById("orderitemidexceptionarr").style.display="none";
					document.getElementById("eventcodearr").style.display="none";
    	    	}else{
					 // 60 장바구니담은사람, 61 위시담은사람, 62 장바구니orWish담은사람, 65 최근6개월구매한사람, 66 당일구매안한사람
					if (comp.value=='60' || comp.value=='61' || comp.value=='62' || comp.value=='65' || comp.value=='66'){
						document.getElementById("makeridarr").style.display="inline";
						document.getElementById("itemidarr").style.display="inline";
						document.getElementById("keywordarr").style.display="none";
						document.getElementById("bonuscouponidxarr").style.display="none";
						document.getElementById("eventcodearr").style.display="none";
					// 63 브랜드를찜한사람
					}else if(comp.value=='63'){
						document.getElementById("makeridarr").style.display="inline";
						document.getElementById("itemidarr").style.display="none";
						document.getElementById("keywordarr").style.display="none";
						document.getElementById("bonuscouponidxarr").style.display="none";
						document.getElementById("eventcodearr").style.display="none";
					// 64 클릭이나키워드검색한사람
					}else if(comp.value=='64'){
						document.getElementById("makeridarr").style.display="inline";
						document.getElementById("itemidarr").style.display="inline";
						document.getElementById("keywordarr").style.display="inline";
						document.getElementById("bonuscouponidxarr").style.display="none";
						document.getElementById("eventcodearr").style.display="none";
					// 67 보너스쿠폰다운후모든쿠폰안쓴사람, 68 보너스쿠폰다운안받은사람
					}else if(comp.value=='67' || comp.value=='68'){
						document.getElementById("makeridarr").style.display="none";
						document.getElementById("itemidarr").style.display="none";
						document.getElementById("keywordarr").style.display="none";
						document.getElementById("bonuscouponidxarr").style.display="inline";
						document.getElementById("eventcodearr").style.display="none";
					// 69 마일리지소멸예정잔액
					}else if(comp.value=='69'){
						document.getElementById("makeridarr").style.display="none";
						document.getElementById("itemidarr").style.display="none";
						document.getElementById("keywordarr").style.display="none";
						document.getElementById("bonuscouponidxarr").style.display="none";
						document.getElementById("eventcodearr").style.display="inline";
					}else{
						document.getElementById("makeridarr").style.display="none";
						document.getElementById("itemidarr").style.display="none";
						document.getElementById("keywordarr").style.display="none";
						document.getElementById("bonuscouponidxarr").style.display="none";
						document.getElementById("eventcodearr").style.display="none";
					}
					document.getElementById("member_pushyn_checkyn").style.display="inline";
					document.getElementById("exceptionlogin").style.display="inline";
					document.getElementById("exceptionuserlevelarr").style.display="inline";
					document.getElementById("replacetagcode").style.display="inline";
					document.getElementById("orderitemidexceptionarr").style.display="inline";
					if (inputfrm.sendmethod.value=="KAKAOALRIM"){
						document.getElementById("exceptionmember_kakaoalrimyn_checkyn").style.display="inline";
					}
					callreplacetagcodeajax(comp.value);
    	    	}

    	    }else{
    	        document.getElementById("member_pushyn_checkyn").style.display="none";
				document.getElementById("exceptionlogin").style.display="none";
				document.getElementById("exceptionuserlevelarr").style.display="none";
				document.getElementById("replacetagcode").style.display="none";
				document.getElementById("makeridarr").style.display="none";
				document.getElementById("itemidarr").style.display="none";
				document.getElementById("keywordarr").style.display="none";
				document.getElementById("bonuscouponidxarr").style.display="none";
				document.getElementById("orderitemidexceptionarr").style.display="none";
				document.getElementById("eventcodearr").style.display="none";
				document.getElementById("exceptionmember_kakaoalrimyn_checkyn").style.display="none";
    	    }

    	}else if (comp.name=="sendmethod"){
    	    if (comp.value=='LMS'){
				document.getElementById("divtitle").style.display="";
				document.getElementById("divbutton").style.display="none";
				document.getElementById("divadvertising_comment1").style.display="";
				document.getElementById("template_code").style.display="none";
				document.getElementById("exceptionmember_kakaoalrimyn_checkyn").style.display="none";
    	    }else{
				if (comp.value=='KAKAOALRIM'){
					document.getElementById("template_comment").style.display="";
					document.getElementById("template_code").style.display="";
					document.getElementById("exceptionmember_kakaoalrimyn_checkyn").style.display="";
					calltemplateajax(comp.value,"")
				}else{
					document.getElementById("template_comment").style.display="none";
					document.getElementById("template_code").style.display="none";
					document.getElementById("exceptionmember_kakaoalrimyn_checkyn").style.display="none";
				}
				document.getElementById("divtitle").style.display="none";
				document.getElementById("divbutton").style.display="";
				document.getElementById("divadvertising_comment1").style.display="none";
			}

    	}else if (comp.name=="failed_type"){
    	    if (comp.value!=''){
				document.getElementById("divfailed_subject").style.display="";
				document.getElementById("divfailed_msg").style.display="";
    	    }else{
				document.getElementById("divfailed_subject").style.display="none";
				document.getElementById("divfailed_msg").style.display="none";
			}
		}
	}

	// 직접 타게팅 입력
	function csvtarget(ridx, mode){
		var popcsvtarget = window.open('/admin/appmanage/lms/poplmsmsg_file.asp?ridx='+ridx+'&mode='+mode+'&menupos=<%=menupos%>','addcsvtarget','width=1400,height=600,scrollbars=yes,resizable=yes');
		popcsvtarget.focus();
	}

	function acttarget(){
		var frm=document.inputfrm;

		// 66 당일구매안한사람
		if (frm.targetkey.value=='66'){
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

	function deltarget(){
		var frm=document.inputfrm;

	    if (!confirm('타게팅을 리셋 하시겠습니까?')) return;
		var frm = document.frmtarget
        frm.mode.value="targetdel";
        
		frm.target = "FrameCKP";
		frm.submit();
	}

	function retarget(){
		var frm=document.inputfrm;

		// 66 당일구매안한사람
		if (frm.targetkey.value=='66'){
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
			$("#template_code").empty().html(str);
		}
	}

	// 템플릿내용 가져오기. 아작스
	function calltemplatecontentsajax(sendmethod,template_code){
		if (sendmethod=="KAKAOALRIM"){
			// 알림톡 수기템플릿
			if (template_code=="etc-9999"){
				document.getElementById("spanetc_template_code").style.display="";
			}else{
				document.getElementById("spanetc_template_code").style.display="none";
				$("#etc_template_code").val("");
			}
		}else{
			document.getElementById("spanetc_template_code").style.display="none";
			$("#etc_template_code").val("");
		}

		$.ajax({
			type: "POST",
			url: "/admin/appmanage/lms/lmstemplate_act.asp",
			data: "sendmethod="+sendmethod+"&template_code="+template_code+"&mode=templatecontentsajax",
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
<form name="inputfrm" id="inputfrm" method="post" action="/admin/appmanage/lms/dolmsmsg_proc.asp" style="margin:0px;">
<input type="hidden" name="ridx" value="<%= ridx %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="repeatlmsyn" value="<%= repeatlmsyn %>">
<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="4" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>메세지 등록/수정</b></font><br/><br/>
	</td>
</tr>

<% If ridx <> "0" Then %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">번호</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<b><%=ridx%></b>
		</td>
	</tr>
<% End If %>

<tr>
	<td width="150" align="center" bgcolor="<%= adminColor("tabletop") %>">발송방법</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% if ridx <> 0 Then %>
			<%= Selectsendmethodname(sendmethod) %>
			<input type="hidden" name="sendmethod" value="<%= sendmethod %>">
		<% else %>
			<% Drawsendmethod "sendmethod",sendmethod, " onChange='setComp(this);'","" %>
		<% end if %>
	</td>
</tr>
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
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">타켓</td>
	<td colspan="3" bgcolor="#FFFFFF">
        타겟 대상:<% call drawSelectBoxlmsTarget("targetkey",targetkey," onChange='setComp(this);'", "") %>
		<span id="makeridarr" <%=CHKIIF(targetKey="60" or targetKey="61" or targetKey="62" or targetKey="63" or targetKey="64" or targetKey="65" or targetKey="66",""," style='display:none'")%>>
			<br><br>브랜드ID:<textarea name="makeridarr" cols=40 rows=2><%= makeridarr %></textarea> 예) ithinkso,7321
		</span>
		<span id="itemidarr" <%=CHKIIF(targetKey="60" or targetKey="61"  or targetKey="62" or targetKey="64" or targetKey="65" or targetKey="66",""," style='display:none'")%>>
			<br><br>상품코드:<textarea name="itemidarr" cols=40 rows=2><%= itemidarr %></textarea> 예) 12334,432132
		</span>
		<span id="keywordarr" <%=CHKIIF(targetKey="64",""," style='display:none'")%>>
			<br><br>키워드:<textarea name="keywordarr" cols=40 rows=2><%= keywordarr %></textarea> 예) 우산,책상
		</span>
		<span id="bonuscouponidxarr" <%=CHKIIF(targetKey="67" or targetKey="68",""," style='display:none'")%>>
			<br><br>보너스쿠폰번호:<textarea name="bonuscouponidxarr" cols=40 rows=2><%= bonuscouponidxarr %></textarea> 예) 652,671
		</span>
		<span id="eventcodearr" <%=CHKIIF(targetKey="69",""," style='display:none'")%>>
			<br><br>이벤트번호:<textarea name="eventcodearr" cols=40 rows=2><%= eventcodearr %></textarea> 예) 100001,100002
		</span>
    </td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제외타켓</td>
	<td colspan="3" bgcolor="#FFFFFF">
		이전발송대상자제외 : 
		<input type="radio" name="exception7dayyn" value="" <% if exception7dayyn="" or isnull(exception7dayyn) then response.write " checked" %>>선택안함
		<input type="radio" name="exception7dayyn" value="currentday" <% if exception7dayyn="currentday" then response.write " checked" %>>오늘
		<input type="radio" name="exception7dayyn" value="before1day" <% if exception7dayyn="before1day" then response.write " checked" %>>1일전
		<input type="radio" name="exception7dayyn" value="before2day" <% if exception7dayyn="before2day" then response.write " checked" %>>2일전
		<input type="radio" name="exception7dayyn" value="before3day" <% if exception7dayyn="before3day" then response.write " checked" %>>3일전
		<input type="radio" name="exception7dayyn" value="before7day" <% if exception7dayyn="before7day" then response.write " checked" %>>7일전
		<input type="radio" name="exception7dayyn" value="before14day" <% if exception7dayyn="before14day" then response.write " checked" %>>14일전
		<input type="radio" name="exception7dayyn" value="before21day" <% if exception7dayyn="before21day" then response.write " checked" %>>21일전
		<input type="radio" name="exception7dayyn" value="currentmonth" <% if exception7dayyn="currentmonth" then response.write " checked" %>>당월
		<input type="radio" name="exception7dayyn" value="before1_1month" <% if exception7dayyn="before1_1month" then response.write " checked" %>>1달전(1일기준)
		<input type="radio" name="exception7dayyn" value="before2_1month" <% if exception7dayyn="before2_1month" then response.write " checked" %>>2달전(1일기준)
		<input type="radio" name="exception7dayyn" value="before3_1month" <% if exception7dayyn="before3_1month" then response.write " checked" %>>3달전(1일기준)
		<span id="orderitemidexceptionarr" <%=CHKIIF(targetkey>3,""," style='display:none'")%>>
			<br><br>해당상품구매한사람제외 : <textarea name="orderitemidexceptionarr" cols=40 rows=2><%= orderitemidexceptionarr %></textarea> 예) 5555,5556
		</span>
        <span id="member_pushyn_checkyn" <%=CHKIIF(targetkey>1,"","style='display:none'")%>>
			<br><br>푸시수신제외 : <% drawSelectBoxisusingYN "member_pushyn_checkyn",member_pushyn_checkyn,"" %>
			예) Y선택 : 푸시수신여부 N인사람이 제외 됩니다. / N선택 : 푸시수신여부 Y인사람이 제외 됩니다.
        </span>
		<span id="exceptionlogin" <%=CHKIIF(targetkey>1,"","style='display:none'")%>>
			<br><br>로그인한사람제외 : 
			<input type="radio" name="exceptionlogin" value="" <% if exceptionlogin="" or isnull(exceptionlogin) then response.write " checked" %>>선택안함
			<input type="radio" name="exceptionlogin" value="currentday" <% if exceptionlogin="currentday" then response.write " checked" %>>오늘
			<input type="radio" name="exceptionlogin" value="before1day" <% if exceptionlogin="before1day" then response.write " checked" %>>1일전
			<input type="radio" name="exceptionlogin" value="before2day" <% if exceptionlogin="before2day" then response.write " checked" %>>2일전
			<input type="radio" name="exceptionlogin" value="before3day" <% if exceptionlogin="before3day" then response.write " checked" %>>3일전
			<input type="radio" name="exceptionlogin" value="before7day" <% if exceptionlogin="before7day" then response.write " checked" %>>7일전
			<input type="radio" name="exceptionlogin" value="before14day" <% if exceptionlogin="before14day" then response.write " checked" %>>14일전
			<input type="radio" name="exceptionlogin" value="before21day" <% if exceptionlogin="before21day" then response.write " checked" %>>21일전
			<input type="radio" name="exceptionlogin" value="currentmonth" <% if exceptionlogin="currentmonth" then response.write " checked" %>>당월
			<input type="radio" name="exceptionlogin" value="before1_1month" <% if exceptionlogin="before1_1month" then response.write " checked" %>>1달전(1일기준)
			<input type="radio" name="exceptionlogin" value="before2_1month" <% if exceptionlogin="before2_1month" then response.write " checked" %>>2달전(1일기준)
			<input type="radio" name="exceptionlogin" value="before3_1month" <% if exceptionlogin="before3_1month" then response.write " checked" %>>3달전(1일기준)
		</span>
		<span id="exceptionuserlevelarr" <%=CHKIIF(targetkey>1,"","style='display:none'")%>>
		<%= exceptionuserlevelarr %>
			<br><br>회원등급제외 : 
			<input type="checkbox" name="exceptionuserlevelarr" value="0" <%=CHKIIF(instr(exceptionuserlevelarr,"0")>0," checked","")%>>WHITE
			<input type="checkbox" name="exceptionuserlevelarr" value="1" <%=CHKIIF(instr(exceptionuserlevelarr,"1")>0," checked","")%>>RED
			<input type="checkbox" name="exceptionuserlevelarr" value="2" <%=CHKIIF(instr(exceptionuserlevelarr,"2")>0," checked","")%>>VIP
			<input type="checkbox" name="exceptionuserlevelarr" value="3" <%=CHKIIF(instr(exceptionuserlevelarr,"3")>0," checked","")%>>VIPGOLD
			<input type="checkbox" name="exceptionuserlevelarr" value="4" <%=CHKIIF(instr(exceptionuserlevelarr,"4")>0," checked","")%>>VVIP
			<input type="checkbox" name="exceptionuserlevelarr" value="7" <%=CHKIIF(instr(exceptionuserlevelarr,"7")>0," checked","")%>>STAFF
			<input type="checkbox" name="exceptionuserlevelarr" value="8" <%=CHKIIF(instr(exceptionuserlevelarr,"8")>0," checked","")%>>FAMILY
		</span>
		<span id="exceptionmember_kakaoalrimyn_checkyn" <%=CHKIIF(sendmethod="KAKAOALRIM" and targetkey>1,"","style='display:none'")%>>
			<br><br><input type="checkbox" name="member_kakaoalrimyn_checkyn" value="Y" <%=CHKIIF(member_kakaoalrimyn_checkyn="Y","checked","")%>>알림톡광고알림거부자제외
		</span>
    </td>
</tr>
<tr height="60" id="divtitle" <%=CHKIIF(sendmethod="LMS","","style='display:none'")%>>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="title" value="<%= title %>" size="160"/>
	</td>
</tr>
<tr height="60">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">내용</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<span id="template_code" <%=CHKIIF(sendmethod="KAKAOALRIM","","style='display:none'")%>>
		</span>
		<span id="spanetc_template_code" <%=CHKIIF(sendmethod="KAKAOALRIM" and template_code="etc-9999","","style='display:none'")%>>
		템플릿코드 : <input type="text" class="text" name="etc_template_code" id="etc_template_code" value="<%= etc_template_code %>" maxlength="32" size="10" /><br><br>
		</span>
		<textarea name="contents" cols=100 rows=8><%= contents %></textarea>
		<span id="divadvertising_comment1" <%=CHKIIF(sendmethod="LMS","","style='display:none'")%>>
			<br><br>맨앞에 <font color="red">(광고)</font> 꼭! 넣어주세요.
			<br>맨뒤에 <font color="red">(무료수신거부) 080-851-6030</font> 꼭! 넣어주세요.
		</span>
		<span id="replacetagcode" <%=CHKIIF(targetkey>1,"","style='display:none'")%>></span>
		<span id="template_comment" <%=CHKIIF(sendmethod="KAKAOALRIM","","style='display:none'")%>>
			<br><br>템플릿 내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.
		</span>
	</td>
</tr>
<tr height="60" id="divbutton" <%=CHKIIF(sendmethod<>"LMS","","style='display:none'")%>>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">KAKAO</td>
	<td colspan="3" bgcolor="#FFFFFF">
		버튼이름1 : 
		<Br><input type="text" class="text" name="button_name" value="<%= button_name %>" size="64" maxlength=64 />
		예) 확인하러 가기
		<Br>
		버튼모바일주소1 : 
		<Br><input type="text" class="text" name="button_url_mobile" value="<%= button_url_mobile %>" size="120" maxlength=256 />
		예) https://tenten.app.link/J3xFnMMFT4
		<Br><Br>
		버튼이름2 : 
		<Br><input type="text" class="text" name="button_name2" value="<%= button_name2 %>" size="64" maxlength=64 />
		<Br>
		버튼모바일주소2 : 
		<Br><input type="text" class="text" name="button_url_mobile2" value="<%= button_url_mobile2 %>" size="120" maxlength=256 />
		<Br><Br>
		실패시문자발송여부 : <% Drawfailed_type "failed_type", failed_type, " onChange='setComp(this);'" %>
		<span id="divfailed_subject" <%=CHKIIF(sendmethod<>"LMS" and failed_type<>"",""," style='display:none'")%>>
			<br><br>실패시문자제목:
			<br><input type="text" class="text" name="failed_subject" value="<%= failed_subject %>" size="55" maxlength=50 />
		</span>
		<span id="divfailed_msg" <%=CHKIIF(sendmethod<>"LMS" and failed_type<>"",""," style='display:none'")%>>
			<br><br>실패시문자내용:
			<br><textarea name="failed_msg" cols=100 rows=8><%= failed_msg %></textarea>
		</span>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">타겟상태</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="hidden" name="targetcnt" value="<%= targetcnt %>">
		<% if ridx <> 0 Then %>
			<%= targetStateName %> 
			타겟수량(<%=FormatNumber(TargetCnt,0)%>)

			<% if targetkey=1 or targetkey=2 or targetkey=3 then %>
				<% If state = 0 or C_ADMIN_AUTH Then %>
					&nbsp;<input type="button" onClick="csvtarget('<%= ridx %>','csvtarget')" value="CSV타게팅" class="button">
					<% If targetState > 0 Then %>
						&nbsp;<input type="button" value="타게팅리셋" onClick="deltarget()" class="button">
					<% end if %>
				<% end if %>
			<% else %>
				<% if targetkey<>9999 then %>
					<% if iIsTargetActionValid then %>
						&nbsp;<input type="button" value="타게팅" onClick="acttarget()" class="button">
					<% else %>
						<% if state<9 then %>
							&nbsp;<input type="button" value="재타게팅" onClick="retarget()" class="button">
							<% If targetState > 0 Then %>
								&nbsp;<input type="button" value="타게팅리셋" onClick="deltarget()" class="button">
							<% end if %>
						<% end if %>
					<% end if %>
				<% end if %>
			<% end if %>
		<% else %>
			<font color="red">신규등록 저장후 수정에서 타켓팅을 해주세요.</font>
		<% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상태</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% If ridx = "0" Or ridx = "" Then %>
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
				<%= lmsmsgstate(state) %>
			<% End If %>
		<% End If %>
	</td>
</tr>

<% if mode = "mEdit" then %>
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
	</td>
</tr>
</table>
</form>
<form name="frmtarget" method="post" action="/admin/appmanage/lms/dolmsmsg_proc.asp" style="margin:0px;">
	<input type="hidden" name="ridx" value="<%= ridx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="target">
</form>
<form name="frmdel" method="get" action="/admin/appmanage/lms/dolmsmsg_proc.asp" style="margin:0px;">
	<input type="hidden" name="ridx" value="<%= ridx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="del">
</form>
<form name="frmstate" method="get" action="/admin/appmanage/lms/dolmsmsg_proc.asp" style="margin:0px;">
	<input type="hidden" name="ridx" value="<%= ridx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="state">
</form>

<% if (application("Svr_Info")="Dev") then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="100%" height="500"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<script type="text/javascript">
	<% if ridx <> 0 Then %>
		<% if sendmethod="KAKAOALRIM" then %>
			calltemplateajax('<%= sendmethod %>','<%= template_code %>')
		<% end if %>

		<% if targetkey<>"1" then %>
			callreplacetagcodeajax('<%= targetkey %>');
		<% end if %>
	<% end if %>
</script>
<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
