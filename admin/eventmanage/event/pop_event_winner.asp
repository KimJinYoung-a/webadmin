<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_event_winner.asp
' Description :  이벤트 당첨등록
' History : 2007.02.22 정윤정 생성
'           2009.08.06 허진원 SMS/이메일 발송 추가
'			2020.04.09 한용민 수정(사은품구분 체크 추가)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode, cEvtCont
Dim egKindCode, ename, ebrand
	eCode		= requestCheckVar(Request("eC"),10)
 	egKindCode 	= requestCheckVar(Request("egKC"),10) 	

if eCode<>"" and eCode<>"0" then
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	
	cEvtCont.fnGetEventCont	 '이벤트 내용 가져오기
	ename 		= db2html(cEvtCont.FEName)
	ebrand 		= cEvtCont.FEBrand
	set cEvtCont = nothing
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

	function jsChType(iVal){	
		$("#span1").hide();

		if(iVal == "2"){
			$("#div1").hide();
			$("#div2").show();
			$("#div3").hide();
		}else if	(iVal == "3"){
			$("#div1").show();
			$("#div2").hide();
			$("#div3").hide();
			$("#div3_1").show();
		}else if	(iVal == "5"){
			$("#div1").show();
			$("#div2").hide();
			$("#div3").show();
			$("#div3_1").hide();
		}else if	(iVal == "1"){
			$("#div1").hide();
			$("#div2").hide();
			$("#div3").hide();
			$("#span1").show();
		}else{
			$("#div1").hide();
			$("#div2").hide();
			$("#div3").hide();
		}	
	}
	
	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	
	function jsWinnerSubmit(){
		var frm=document.frmWin;
		if(!frm.sR.value){
			alert("등수를 입력해주세요");
			frm.sR.focus();
			return false;
		}
		
		if(!IsDigit(frm.sR.value)){
			alert("등수는 숫자만 입력가능합니다.");
			frm.sR.focus();
			return false;
		}

		if(frm.uploadtype[0].checked){
			if(!frm.sW.value){
				alert("당첨자를 입력해주세요");
				frm.sW.focus();
				return false;
			}
		}
		
		if(frm.evtprizetype.value == "3"){
			if(!frm.sGKN.value){
				alert("사은품명을  입력해 주세요");
				frm.sGKN.focus();
				return false;
			}

			if(!frm.iGK.value){
				alert("사은품명을 확인 버튼을 눌러서 확인해 주세요");
				return false;
			}

			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('출고 요청일을 선택하세요.');
			    return false;
			}
			
			if (frm.reqdeliverdate.value<=frm.dAEDate.value){
			    // 특정 상품배송 이벤트로 당첨확인 기간 내에도 출고요청일 지정 가능(2015.04.28; 허진원)
			    if(!confirm('출고 요청일이 당첨확인기간 이전으로 지정되어있습니다.\n\n확인하셨습니까?')){
				    frm.reqdeliverdate.focus();
				    return false;
			    }
			    //alert('출고 요청일은 당첨확인기간 이후로 입력하셔야 합니다.');
			    //frm.reqdeliverdate.focus();
			    //return false;
			}

			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('배송 구분을 선택하세요.');
        		return false;
        	}
			if(frm.isupchebeasong[1].checked&&(frm.jungsan.checked)&&(frm.jungsanValue.value=="")){
			    alert('정산액(매입가)를 입력하세요');
			    frm.jungsanValue.focus();
			    return false;
			}
            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('업체 아이디를 선택하세요.');
        		return false;
            }
            
            
		}
		
		if(frm.evtprizetype.value == "2"){
			if(!frm.couponvalue.value){
				alert("쿠폰금액 또는 할인율을 입력해주세요!");
				frm.couponvalue.focus();
				return false;
			}
			
			if(!frm.minbuyprice.value){
				alert("최소금액을 입력해주세요!");
				frm.minbuyprice.focus();
				return false;
			}
			
			 if(!frm.sDate.value || !frm.eDate.value ){
			  	alert("기간을 입력해주세요");
			  	frm.sDate.focus();
			  	return false;
			  }
		
			  if(frm.sDate.value > frm.eDate.value){
			  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.sDate.focus();
			  	return false;
			  }	  		
		}
		
		if(frm.evtprizetype.value == "5"){
			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('출고 요청일을 선택하세요.');
			    return false;
			}
			
			if (frm.reqdeliverdate.value<=frm.dAEDate.value){
			    alert('출고 요청일은 당첨확인기간 이후로 입력하셔야 합니다.');
			    frm.reqdeliverdate.focus();
			    return false;
			}
			
			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('배송 구분을 선택하세요.');
        		return false;
        	}
            
            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('업체 아이디를 선택하세요.');
        		return false;
            }
            
			 if(frm.itemuse_itemid.value == ""){
			  	alert("테스터상품을 선택해주세요");
			  	return false;
			  }
			  
			 if(frm.itemuse.value == ""){
			  	alert("고객님께 보여지는 테스터상품명을 입력해주세요");
			  	frm.itemuse.focus();
			  	return false;
			  }
            
            if(GetByteLength(frm.itemuse.value) > 100)
            {
			  	alert("테스터상품명을 100 Byte 이내로 입력해주세요");
			  	frm.itemuse.focus();
			  	return false;
            }
            
			 if(!frm.itemuse_sDate.value || !frm.itemuse_eDate.value ){
			  	alert("테스터상품사용기간을 입력해주세요");
			  	frm.itemuse_sDate.focus();
			  	return false;
			  }
		
			  if(frm.itemuse_sDate.value > frm.itemuse_eDate.value){
			  	alert("테스터상품사용기간 종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.itemuse_sDate.focus();
			  	return false;
			  }
			  
			 if(!frm.usewrite_sDate.value || !frm.usewrite_eDate.value ){
			  	alert("테스터후기등록기간을 입력해주세요");
			  	frm.usewrite_sDate.focus();
			  	return false;
			  }
		
			  if(frm.usewrite_sDate.value > frm.usewrite_eDate.value){
			  	alert("테스터후기등록기간 종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.usewrite_sDate.focus();
			  	return false;
			  }	 
		}
		

		if(confirm("등록하신 내용은 수정 또는 삭제가 불가능하며 고객에게 바로 적용됩니다.\n\n등록 하시겠습니까? ")){
			if(frm.uploadtype[1].checked){
				frm.target = "excelframe";
				$("#normalSubmit").hide();
				$("#excelprocing").show();
			} else {
				frm.target = "";
			}
			frm.action="eventprize_process.asp";
			frm.submit();
			return true;
		}else{
		    return false;
		}
	}

	function jsAutoWinnerSubmit(){
		var frm=document.frmWin;
		if(!frm.sR.value){
			alert("등수를 입력해주세요");
			frm.sR.focus();
			return false;
		}
		
		if(!IsDigit(frm.sR.value)){
			alert("등수는 숫자만 입력가능합니다.");
			frm.sR.focus();
			return false;
		}

		if(frm.evtprizetype.value == "3"){
			if(!frm.sGKN.value){
				alert("사은품명을  입력해 주세요");
				frm.sGKN.focus();
				return false;
			}

			if(!frm.iGK.value){
				alert("사은품명을 확인 버튼을 눌러서 확인해 주세요");
				return false;
			}

			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('출고 요청일을 선택하세요.');
			    return false;
			}
			
			if (frm.reqdeliverdate.value<=frm.dAEDate.value){
			    // 특정 상품배송 이벤트로 당첨확인 기간 내에도 출고요청일 지정 가능(2015.04.28; 허진원)
			    if(!confirm('출고 요청일이 당첨확인기간 이전으로 지정되어있습니다.\n\n확인하셨습니까?')){
				    frm.reqdeliverdate.focus();
				    return false;
			    }
			    //alert('출고 요청일은 당첨확인기간 이후로 입력하셔야 합니다.');
			    //frm.reqdeliverdate.focus();
			    //return false;
			}

			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('배송 구분을 선택하세요.');
        		return false;
        	}
			if(frm.isupchebeasong[1].checked&&(frm.jungsan.checked)&&(frm.jungsanValue.value=="")){
			    alert('정산액(매입가)를 입력하세요');
			    frm.jungsanValue.focus();
			    return false;
			}
            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('업체 아이디를 선택하세요.');
        		return false;
            }
            
            
		}
		
		if(frm.evtprizetype.value == "2"){
			if(!frm.couponvalue.value){
				alert("쿠폰금액 또는 할인율을 입력해주세요!");
				frm.couponvalue.focus();
				return false;
			}
			
			if(!frm.minbuyprice.value){
				alert("최소금액을 입력해주세요!");
				frm.minbuyprice.focus();
				return false;
			}
			
			 if(!frm.sDate.value || !frm.eDate.value ){
			  	alert("기간을 입력해주세요");
			  	frm.sDate.focus();
			  	return false;
			  }
		
			  if(frm.sDate.value > frm.eDate.value){
			  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.sDate.focus();
			  	return false;
			  }	  		
		}
		
		if(frm.evtprizetype.value == "5"){
			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('출고 요청일을 선택하세요.');
			    return false;
			}
			
			if (frm.reqdeliverdate.value<=frm.dAEDate.value){
			    alert('출고 요청일은 당첨확인기간 이후로 입력하셔야 합니다.');
			    frm.reqdeliverdate.focus();
			    return false;
			}
			
			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('배송 구분을 선택하세요.');
        		return false;
        	}
            
            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('업체 아이디를 선택하세요.');
        		return false;
            }
            
			 if(frm.itemuse_itemid.value == ""){
			  	alert("테스터상품을 선택해주세요");
			  	return false;
			  }
			  
			 if(frm.itemuse.value == ""){
			  	alert("고객님께 보여지는 테스터상품명을 입력해주세요");
			  	frm.itemuse.focus();
			  	return false;
			  }
            
            if(GetByteLength(frm.itemuse.value) > 100)
            {
			  	alert("테스터상품명을 100 Byte 이내로 입력해주세요");
			  	frm.itemuse.focus();
			  	return false;
            }
            
			 if(!frm.itemuse_sDate.value || !frm.itemuse_eDate.value ){
			  	alert("테스터상품사용기간을 입력해주세요");
			  	frm.itemuse_sDate.focus();
			  	return false;
			  }
		
			  if(frm.itemuse_sDate.value > frm.itemuse_eDate.value){
			  	alert("테스터상품사용기간 종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.itemuse_sDate.focus();
			  	return false;
			  }
			  
			 if(!frm.usewrite_sDate.value || !frm.usewrite_eDate.value ){
			  	alert("테스터후기등록기간을 입력해주세요");
			  	frm.usewrite_sDate.focus();
			  	return false;
			  }
		
			  if(frm.usewrite_sDate.value > frm.usewrite_eDate.value){
			  	alert("테스터후기등록기간 종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.usewrite_sDate.focus();
			  	return false;
			  }	 
		}

		if(confirm("등록하신 내용은 수정 또는 삭제가 불가능하며 고객에게 바로 적용됩니다.\n\n등록 하시겠습니까? ")){
			if(frm.uploadtype[1].checked){
				frm.target = "excelframe";
				$("#normalSubmit").hide();
				$("#excelprocing").show();
			} else {
				frm.target = "";
			}
			frm.action="eventprize_auto_process.asp";
			frm.submit();
			return true;
		}else{
		    return false;
		}
	}

	function disabledBox(comp){
        var frm = comp.form;
        if (comp.value=="Y"){
            frm.makerid.disabled = false;
            frm.jungsan.disabled = false;

			frm.jungsanValue.disabled = false;
	        //frm.jungsan.checked = true;
        }else{
            frm.makerid.selectedIndex = 0;
            frm.makerid.value = '';
            frm.makerid.disabled = true;
            frm.jungsan.disabled = true;

	        frm.jungsanValue.value = '';
	        frm.jungsanValue.disabled = true;
	        frm.jungsan.checked = false;
        }
    }

	//사은품 종류 등록
	function jsSetGiftKind(){
		var gift_delivery, isupchebeasong;
		var sGKN;
		var makerid;

		for (var i=0;i < frmWin.isupchebeasong.length; i++){
			if (frmWin.isupchebeasong[i].checked){
				isupchebeasong=frmWin.isupchebeasong[i].value;
			}
		}

		sGKN=frmWin.evt_name.value
		makerid=frmWin.ebrand.value

		if (isupchebeasong==""){
			alert("배송구분을 선택해 주세요.");
			return;
		}
		gift_delivery=isupchebeasong

		var winkind;
		winkind = window.open('/admin/shopmaster/gift/popgiftKindReg.asp?gift_delivery='+gift_delivery+'&makerid='+makerid+'&sGKN='+sGKN,'popkind','width=1280px, height=960px, scrollbars=yes');
		winkind.focus();
	}

	//SMS내용 확인
	function chkSMSTextLength(cont) {
		if(GetByteLength(cont)>80) {
			alert("SMS은 80 Byte까지만 발송 가능합니다.");
		}
		$("#smsCnt").html(GetByteLength(cont));
	}

	function swSMS() {
		if(frmWin.chkSMS.checked) {
			frmWin.smsCont.className="textarea";
			frmWin.smsCont.disabled=false;
		} else {
			frmWin.smsCont.className="textarea_ro";
			frmWin.smsCont.disabled=true;
		}
	}
	function swEmail() {
		if(frmWin.chkEmail.checked) {
			frmWin.emailCont.className="textarea";
			frmWin.emailCont.disabled=false;
		} else {
			frmWin.emailCont.className="textarea_ro";
			frmWin.emailCont.disabled=true;
		}
	}
	
	function GetByteLength(val){
	 	var real_byte = val.length;
	 	for (var ii=0; ii<val.length; ii++) {
	  		var temp = val.substr(ii,1).charCodeAt(0);
	  		if (temp > 127) { real_byte++; }
	 	}
	
	   return real_byte;
	}
	
	function ViewByteLength()
	{
		frmWin.bytecheck.value = GetByteLength(frmWin.itemuse.value);
	}

	function jungsanYN(){
		var frm = document.frmWin;
		if(frm.jungsan.checked==true){
			frm.jungsanValue.disabled = false;
		}else{
			frm.jungsanValue.value = '';
			frm.jungsanValue.disabled = true;
		}
	}
	function checkover1(obj) {
		var val = obj.value;
		if (val) {
			if (val.match(/^\d+$/gi) == null) {
				alert("숫자만 넣으세요!");
				document.frmWin.jungsanValue.value = '';
				obj.select();
				return;
			}
		}
	}
	
	function jsUploadType(a){
		if(a == "direct"){
			$("#spandirect").show();
			$("#spanexcel").hide();
		} else {
			$("#spanexcel").show();
			$("#spandirect").hide();
		}
	}
	
	function jsGoExcelUp(){
		var winexcel;
		winexcel = window.open('/admin/eventmanage/event/pop_event_winner_excelupload.asp?eventid=<%=eCode%>','winexcel','width=400px, height=150px');
		winexcel.focus();
	}
	
	function jsPageReload(){
		opener.location.reload();
	}
</script>

<script type="text/javascript">
var speed = 350 //깜빡이는 속도 - 1000은 1초

function doBlink(){
var blink = $("blink");
for (var i=0; i < blink.length; i++)
blink[i].style.visibility = blink[i].style.visibility == "" ? "hidden" : ""
}

function startBlink() {
setInterval("doBlink()",speed)
}
window.onload = startBlink;
</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 당첨자 등록&nbsp;&nbsp;&nbsp;<font color="red"><b>※ <blink><u>테스터 이벤트 당첨 등록시</u></blink> 반드시 <blink><u>구분을 테스터 이벤트로</u></blink> 하세요.</b></font></div>
<table border="0" width="100%" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmWin" method="post">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="egKC" value="<%=egKindCode%>">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="evt_name" value="<%= ename %>">
<input type="hidden" name="ebrand" value="<%= ebrand %>">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="130" align="center" bgcolor="<%= adminColor("tabletop") %>">구분</td>
				<td bgcolor="#FFFFFF">
					<%sbGetOptCommonCodeArr "evtprizetype", "", False,True,"onChange=jsChType(this.value);"%>
				</td>
			</tr>

			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">등수</td>
				<td bgcolor="#FFFFFF"><input type="text" size="2" name="sR"></td>
			</tr>	
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">등수별칭</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="sRN" size="20" maxlength="32">
					<span id="span1" style="display:none;color:darkred;">(당첨확인페이지의 이벤트명에 추가되어 표시됨)</span>
				</td>
			</tr>	
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">당첨확인기간</td>
				<td bgcolor="#FFFFFF"><input type="text" name="dASDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('dASDate');" style="cursor:hand;">
					~<input type="text" name="dAEDate" size="10"  maxlength="10" value="<%=dateadd("d",14,date())%>" onClick="jsPopCal('dAEDate');" style="cursor:hand;"></td>
			</tr>
			<% If eCode="4" Then %>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">당첨인원</td>
				<td bgcolor="#FFFFFF"><input type="text" name="prizecnt" value=""  size="2" maxlength="2" style="cursor:hand;"></td>
			</tr>
			<% End If %>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">당첨자</td>
				<td bgcolor="#FFFFFF">
					<label id="labeldirect" style="cursor:pointer;"><input type="radio" id="labeldirect" name="uploadtype" value="direct" onClick="jsUploadType('direct');" checked>ID 직접입력(100명 이하)</label>
					<label id="labelexcel" style="cursor:pointer;"><input type="radio" id="labelexcel" name="uploadtype" value="excel" onClick="jsUploadType('excel');">Excel로 등록(100명 이상)</label>
					<strong><a href="" onClick="$('#excelexplain').show();return false;"><font color="red">[필독!!]</font></a></strong>
					<div style="display:none;padding:5 0 5 0;" id="excelexplain">
					<strong><a href="" onClick="$('#excelexplain').hide();return false;"><font size="3" color="blue">[ 닫 기 ]</font></a></strong><br>
					* 다운이 안될때는 아래 주소를 긁어 복사한 뒤 인터넷 창 주소에 붙여 넣어 실행.<br>
					* 저장 도중 창이 <strong>다운됐을 경우</strong> 저장된 리스트를 먼저 확인하고, 저장된게 <strong>하나도 없으면 엑셀을 다시 올려서</strong> 진행 하면되고, 저장된게 <strong>있는 경우 등록창을 열어 "Excel로 등록" 을 선택하고 업로드 하지 않고 나머지 내용을 등록 후 저장</strong> 하면 됩니다.
					</div>
					<div style="padding:5 0 5 0;">
						<span id="spandirect" style="display:block;">
							콤머로 구분, 공백없이 (예: aaa,bbb,ccc)<br>
							<textarea name="sW" rows="5" style="width:100%"></textarea>
						</span>
						<span id="spanexcel" style="display:none;">
							<input type="button" value="Excel 등록" onClick="jsGoExcelUp();"> 정해진 폼을 다운받아 등록 [<a href="/admin/eventmanage/event/event_winner_userlist.xls" target="_blank"><u><strong>Download ↓ </strong></u></a>] <%=manageUrl%>/admin/eventmanage/event/event_winner_userlist.xls<br>
							
						</span>
					</div>
				</td>
			</tr>	
		</table>
	</td>
		
</tr>
<tr>
	<td>
		<div id="div1" style="display:;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">							
			<tr>
				<td align="center" width="130"  bgcolor="<%= adminColor("tabletop") %>">배송지 등록구분</td>
				<td bgcolor="#FFFFFF">
					<input type=radio name=rdgubun value="U">User가 배송지 입력
					<input type=radio name=rdgubun value="F" checked>User 기본 주소 사용 <font color="blue">[가능한 기본 주소지 사용]</font>
				</td>
			</tr>				
			<!-- 배송 구분 추가 : 서동석 -->
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">출고요청일</td>
            	<td bgcolor="#FFFFFF">
            		<input type="text" name="reqdeliverdate" size="10" maxlength="10"  value="" >
		            <a href="javascript:jsPopCal('reqdeliverdate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
            	</td>
            </tr>
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">배송구분</td>
            	<td bgcolor="#FFFFFF">
            		<input type=radio name="isupchebeasong" value="N" checked onClick="disabledBox(this);">텐바이텐배송
            		<input type=radio name="isupchebeasong" value="Y" onClick="disabledBox(this);">업체직접배송
            	</td>
            </tr>
			<tr id="div3_1" style="display:block;">
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">사은품명</td>
				<td bgcolor="#FFFFFF"><input type="hidden" name="iGK" >
					<input type="text" name="sGKN" size="10" onkeyup="document.frmWin.iGK.value='';"> 
					<input type="button" class="button" value="확인" onClick="jsSetGiftKind();">				
					<div id="spanImg"></div>	
				</td>
			</tr>				
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">정산여부</td>
            	<td bgcolor="#FFFFFF">
					<input type="checkbox" class="checkbox" name="jungsan" id="jungsan" onclick="javascript:jungsanYN();">정산함&nbsp;&nbsp;
					정산액(매입가) : <input type="text" class="text" id="jungsanValue" name="jungsanValue" onkeyup="checkover1(this)">
            	</td>
            </tr>
            <tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">업체배송시<br>업체ID</td>
            	<td bgcolor="#FFFFFF">
            	    <% drawSelectBoxDesignerwithName "makerid","" %>
            	    <script language='javascript'>
            	    document.frmWin.makerid.disabled=true;
            	    document.frmWin.jungsan.disabled=true;
            	    </script>
            	</td>
            </tr>
		</table>	
		</div>	
		<div id="div2" style="display:none;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">							
			<tr>
				<td align="center" width="130" bgcolor="<%= adminColor("tabletop") %>">쿠폰타입</td>
				<td bgcolor="#FFFFFF">
					<input type=text name=couponvalue maxlength=7 size=10>
					<input type=radio name=coupontype value="1" onclick="alert('% 할인 쿠폰입니다.');">%할인
					<input type=radio name=coupontype value="2" checked >원할인
					(금액 또는 % 할인)
				</td>
			</tr>						
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">최소구매금액</td>
				<td bgcolor="#FFFFFF"><input type=text name=minbuyprice maxlength=7 size=10>원 이상 구매시 사용가능(숫자)</td>
			</tr>	
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">유효기간</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="sDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('sDate');" style="cursor:hand;">
					~<input type="text" name="eDate" size="10"  maxlength="10" onClick="jsPopCal('eDate');" style="cursor:hand;">
				</td>
			</tr>	
		</table>	
		</div>
	</td>
</tr>
<tr id="div3" style="display:none;">
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td align="center" width="130" bgcolor="<%= adminColor("tabletop") %>">테스터상품</td>
				<td bgcolor="#FFFFFF">
					※ 옵션이 있으면 상품명 뒤에 입력하면 됩니다.<br>&nbsp;&nbsp;&nbsp;&nbsp;이곳 입력란은 실제 고객님께 보여지는 테스트 상품명입니다.<br>
					<input type="button" value="상품" onClick="window.open('/admin/eventmanage/event/pop_CateItemList.asp','popWinn','width=800, height=500, scrollbars=yes');">
					<input type="text" name="itemuse" value="" size="50" onkeyup="ViewByteLength()">
					<input type="text" name="bytecheck" value="" size="2">
					<input type="hidden" name="itemuse_itemid" value="">
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">테스터상품사용기간</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="itemuse_sDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('itemuse_sDate');" style="cursor:hand;">
					~<input type="text" name="itemuse_eDate" size="10"  maxlength="10" onClick="jsPopCal('itemuse_eDate');" style="cursor:hand;">
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">테스터후기등록기간</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="usewrite_sDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('usewrite_sDate');" style="cursor:hand;">
					~<input type="text" name="usewrite_eDate" size="10"  maxlength="10" onClick="jsPopCal('usewrite_eDate');" style="cursor:hand;">
				</td>
			</tr>
		</table>
		
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" width="130"  bgcolor="<%= adminColor("tabletop") %>">당첨자 SMS<br>보내기</td>
			<td bgcolor="#FFFFFF">
				<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td><textarea name="smsCont" rows="2" style="width:100%" class="textarea" onkeyup="chkSMSTextLength(this.value)">[텐바이텐] 이벤트당첨을 축하합니다. 공지사항 및 마이텐바이텐을 확인해주세요.</textarea></td>
					<td width="110" valign="bottom"><input type=checkbox name=chkSMS value="Y" checked onClick="swSMS()">SMS동시발송</td>
				</tr>
				<tr>
					<td align="right">※ 80Byte까지 입력가능(현재 <span id="smsCnt">76</span>Byte)</td>
					<td></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" width="130"  bgcolor="<%= adminColor("tabletop") %>">당첨자 이메일<br>보내기</td>
			<td bgcolor="#FFFFFF">
				<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td><textarea name="emailCont" rows="5" style="width:100%" class="textarea_ro" disabled></textarea></td>
					<td width="110" valign="bottom"><input type=checkbox name=chkEmail value="Y" onClick="swEmail()">이메일 동시발송</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<span id="normalSubmit">
		<% If eCode="4" Then %>
		<input type="button" class="button" value="자동당첨" onclick="jsAutoWinnerSubmit();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<% End If %>

		<a href="" onclick="jsWinnerSubmit();return false;"><img src="/images/icon_confirm.gif" border="0"></a>
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</span>
		&nbsp;<br>
		<span id="excelprocing" style="display:none;"><blink><strong>* 처리 되는 중입니다. 창을 닫지말고 최종 완료까지 처리해주세요.</strong></blink></span>
		<span id="excelprocdetail"></span>
		<span id="excelSubmit" style="display:none;"><input type="submit" value="다음 100개 실행" style="height:30px;"></span>
	</td>
</tr>	
</form>	
</table>
<iframe id="excelframe" src="about:blank" name="excelframe" width="0" height="0"></iframe>
<!-- #include virtual="/lib/db/dbclose.asp" -->