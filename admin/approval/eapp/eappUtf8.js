 
/* 전자결재 공통 스크립트 */

//전결규정보기
function popDecision(){
    var popwin = window.open('/admin/approval/eapp/popDecision.asp','popDecision','width=900, height=900, scrollbars=yes,resizable=yes');
}

//파일첨부
function jsAttachFile(sP){
	var winAF = window.open('/admin/approval/eapp/popRegFile.asp?sp='+sP,'popAF','width=400, height=300');
	winAF.focus();
}

//파일삭제
function jsFileDel(sName){
	$("#dF"+sName).remove(); 
}

//파일 다운로드
    function jsDownload(sDownURL, sRFN, sFN){
    var winFD = window.open(sDownURL+"/linkweb/eapp/procDownload.asp?sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
 }


//부서별 자금구분 금액별, 퍼센트별 값 자동 변경
function jsSetMoney(sType,iValue,iPageType){
	var iMoney, iPercent, iRectValue;
	var mRequestPay;
	if (typeof(eval("document.all."+sType+"PM").length) == "undefined" ){
		iMoney = eval("document.all.mPM");
		iPercent = eval("document.all.iPM");
		iRectValue = eval("document.all."+sType+"PM");
	}else{
		iMoney = eval("document.all.mPM["+iValue+"]");
		iPercent = eval("document.all.iPM["+iValue+"]");
		iRectValue = eval("document.all."+sType+"PM["+iValue+"]");
	}

	if(iPageType==1){ // 품의서용
		mRequestPay = document.all.mRP.value.replace(/\,/g,"");
		if(mRequestPay == "" || mRequestPay == 0){
			alert("품의금액을 먼저 입력해주세요");
			iRectValue.value ="";
			document.all.mRP.focus();
			return;
		}
		if(iRectValue.value !=""){
			if (sType =="m"){
				iPercent.value = (parseInt(iMoney.value.replace(/\,/g,""),10)/parseInt(mRequestPay,10)*100).toFixed(1);
			}else{
				iMoney.value = jsSetComma(parseInt(mRequestPay,10)*(parseInt(iPercent.value,10)/100));
			}
		}
	}else{//결제요청서용
		 mRequestPay = document.all.mprp.value.replace(/\,/g,"");
		if(mRequestPay == "" || mRequestPay == 0){
				alert("결제요청금액을 먼저 입력해주세요");
				iRectValue.value ="";
				document.all.mprp.focus();
				return;
			}

			if(iRectValue.value !=""){
				if (sType =="m"){
					iPercent.value = (parseInt(iMoney.value.replace(/\,/g,""),10)/parseInt(mRequestPay,10)*100).toFixed(1);
				}else{
					iMoney.value = jsSetComma(parseInt(mRequestPay,10)*(parseInt(iPercent.value,10)/100));
				}
			}
	}
}


//달력보기
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

// 페이지 이동
function jsGoPage(iCP)
	{
		document.frmList.iCP.value=iCP;
		document.frmList.submit();
	}



// tr 색상변경
var pre_selected_row = null;
var pre_selected_row_color = null;

var pre_selected_row_0  = null;
var pre_selected_row_color_0  = null;
var pre_selected_row_1  = null;
var pre_selected_row_color_1 = null;

function ChangeColor(e, selcolor, defcolor){
	if (pre_selected_row_color != null) {
	        pre_selected_row.bgColor = pre_selected_row_color;
        }

        pre_selected_row = e;
        pre_selected_row_color = defcolor;

       e.bgColor = selcolor;
}

 function evalChangeColor(e, selcolor, defcolor){
	if (pre_selected_row_color_0  != null) {
	        pre_selected_row_0.bgColor = pre_selected_row_color_0;
	        pre_selected_row_1.bgColor = pre_selected_row_color_1;
        }

        pre_selected_row_0 = eval(e+"[0]");
        pre_selected_row_color_0  = defcolor;
        pre_selected_row_1  = eval(e+"[1]");
        pre_selected_row_color_1  = defcolor;

        eval(e+"[0]").bgColor = selcolor;
        eval(e+"[1]").bgColor = selcolor;
}


//결재라인, 참조 아이디 등록
function jsRegID(iMode){ 
		var ilastApprovalid = document.frm.iLAID.value;
		var idpid = document.frm.hidPS.value;
		var icid1 = document.frm.hidcid1.value;
		var icid2 = document.frm.hidcid2.value;
		var icid3 = document.frm.hidcid3.value;
		var icid4 = document.frm.hidcid4.value;
		var winRI = window.open('/admin/approval/eapp/popSetAuthID.asp?iM='+iMode+'&iLAID='+ilastApprovalid+'&sjn='+document.frm.hidJN.value+'&idpid='+idpid+'&icid1='+icid1+'&icid2='+icid2+'&icid3='+icid3+'&icid4='+icid4 ,'popAL','width=800, height=650, resizable=yes, scrollbars=yes');
		winRI.focus();
}

//합의자 추가 2013/10/21
function jsRegID_H(iMode){
		var iAuthPosition= '0';
		var ilastApprovalid = document.frm.hidAI_H.value;
		var idpid = document.frm.hidPS_H.value;	
		if ((idpid=='')||(idpid==undefined)) {
		     idpid = '8';
		}

		var winRI = window.open('/admin/approval/eapp/popSetAuthID.asp?iM='+iMode+'&iLAI='+ilastApprovalid+'&iAP='+iAuthPosition+'&idpid='+idpid ,'popAL','width=650, height=550, resizable=yes, scrollbars=yes');
		winRI.focus();
}


function jsIsHundred()
{
	if(document.frm.iaidx.value == "172" || document.frm.iaidx.value == "173")  //iAIdx = > iaidx
	{
		var mRPtmp;
		mRPtmp = document.frm.mRP.value.replace(/\,/gi, '');
		mRPtmp = parseInt(mRPtmp);

		if(mRPtmp < 1000000)
		{
			alert("수지항목이\n[172] 비품(공통) 또는 [173] 사무실/매장공사/인테리어비용외\n인 경우 품의금액이 100만원 미만시일 경우엔\n다른 수지항목을 선택해야 합니다.\n\n이에 대한 문의는 경영지원팀에 문의하시기 바랍니다.");
			return false;
		}
		else
		{
			return true;
		}
	}
	else
	{
		return true;
	}
}

var blockChar=["&lt;script","<scrip","<form","&lt;form"];  
 function chkContent(p) {
 for (var i=0; i<blockChar.length; i++) {
  if (p.indexOf(blockChar[i])>=0) {
   return blockChar[i];
  }
 }
 return null;
} 

//결재등록(품의서)
function jsEappSubmit(iState){    
	var arrRfI;	 
  	var mRequestPay = document.all.mRP.value.replace(/\,/g,""); 
	if(!jsIsHundred()){
		return;
	}
//	if (document.frm.tmpisAgreeNeed.value == "Y"){
//		if(document.frm.tmpisAgreeNeedTarget.value != document.frm.hidAI_H.value){
//			alert('합의자를 수정할 수 없는 문서입니다.');
//			return;
//		}
//	}
	
	if (iState == 1){
	    //if(jsChkBlank(document.frm.hidAI.value) ){
			//alert("결재자를 등록해주세요");
			//return;
		//}
			if (document.frm.tmpisAgreeNeed.value == "Y" && jsChkBlank(document.frm.hidAI_H.value)){			
			alert('합의자를 등록해주세요');
			return;			
		}
		
		if(jsChkBlank(document.frm.hidALI.value) ){
			alert("최종결재자를 등록해주세요");
			return;
		}
		
		if(jsChkBlank(document.frm.sRN.value) ){
			alert("품의서명을 입력해주세요");
			return;
		}

        //2013/10/28추가
        if ((document.all.hidPE.value=="True")&&(mRequestPay == "" || mRequestPay == 0)){
			alert("품의금액을 먼저 입력해주세요");
			document.all.mRP.focus();
			return;
		}

        if(!jsIsHundred()){
    		return;
    	}

		//결재자아이디와 참조자, 합의자 중복 확인
		if (document.frm.hidAI.value != ""){
			var	arrAI  = document.frm.hidAI.value.split(",");
			var sLastAI = document.frm.hidALI.value;
				for(j=0;j<arrAI.length;j++){
					if(document.frm.hidAI_H.value !=""){
						if(document.frm.hidAI_H.value ==arrAI[j] || document.frm.hidAI_H.value == sLastAI){
						alert("합의자와 결재자는 중복될 수 없습니다. 다시 선택해주세요");
						return;
						} 
					}
					
					if(sLastAI ==arrAI[j]  ){
						alert("최종결재자와 결재자는 중복될 수 없습니다. 다시 선택해주세요");
						return;
					}
						
					if(document.frm.hidRfI.value !=""){
							arrRfI = document.frm.hidRfI.value.split(",");
							for(i=0;i<arrRfI.length;i++){
								if(arrRfI[i] ==arrAI[j]|| arrRfI[i] == sLastAI){
									alert("참조자와 결재자는 중복될 수 없습니다. 다시 선택해주세요");
									return;
								}
							}
					} 
				}
			
		}
	 

		var totPM = 0;
		if(document.frm.hidPE.value=="True"){  // iAIdx => iaidx
		    if (document.all.iP){ //조건추가
    			if(jsChkBlank(document.all.iP.value) ){
    				alert("부서를 등록해주세요");
    				return;
    			}

    			if(jsChkBlank(mRequestPay) ){
    				alert("부서를 등록해주세요");
    				return;
    			}

    			if(typeof(document.all.mPM) !="undefined"){
    			  	if(typeof(document.all.mPM.length)!="undefined"){
    			  		for(i=0;i<document.all.mPM.length;i++){
    						totPM = totPM + parseInt(document.all.mPM[i].value.replace(/\,/g,""));
    						}
    					}else{
    						totPM = document.all.mPM.value.replace(/\,/g,"");
    					}

    				if (parseInt(mRequestPay) != parseInt(totPM)){
    					alert("자금구분 금액이 품의금액과 다릅니다. 재설정해주세요");
    					return;
    				}
    			}
		    }
		}

 
    	 var content = Editor.getContent();
    	 	var str = chkContent(content); 
			  if (str) {
			   alert("script태그 또는 form태그는 사용할 수 없는 문자열 입니다.  html을 체크해서 내용을 수정해주세요");
			   return ;  
			  } 
		  
         document.getElementById("editor").value = content; 
         
//        var conChk =  document.getElementById("editor").value.indexOf('<form');  
//           if (conChk !=-1){
//             alert("내용에 유효하지 않은 <form 태그가 존재합니다. html을 체크해서 내용을 수정해주세요");
//             return ;
//           }

    	if(confirm("결재등록하시겠습니까?")){
    			document.all.mRP.value = document.all.mRP.value.replace(/\,/g,"");
    			if(typeof(document.all.mPM) !="undefined"){
    		  	if(typeof(document.all.mPM.length)!="undefined"){
    		  		for(i=0;i<document.all.mPM.length;i++){
    				 		document.all.mPM[i].value = document.all.mPM[i].value.replace(/\,/g,"");
    				 		if(document.frm.mP.value ==""){
    				 			document.frm.mP.value =document.all.mPM[i].value;
    				 		}else{
    				 			document.frm.mP.value = document.frm.mP.value+","+document.all.mPM[i].value;
    				 		}
    					}
    				}else{
    						document.all.mPM.value = document.all.mPM.value.replace(/\,/g,"");
    						document.frm.mP.value =	document.all.mPM.value;
    				}
    			}
    			
	   	$("input[name='sFileP[]']").each( function(index,elem) {  
			     var a = $(elem).val();  
			     if( document.frm.sFile.value==""){
			     	document.frm.sFile.value = a;
			    }else{
			     document.frm.sFile.value = document.frm.sFile.value + ","+a;
			   }
			  }); 
    		document.frm.hidRS.value = 1;
    		document.frm.submit();
    	}
	}else if(iState ==-1) {
		if(confirm("삭제하시겠습니까?")){
	    	document.frm.hidM.value = "D";
			document.frm.submit();
		}
	}else  if(iState ==0) {   
	    
        var content = Editor.getContent();
         document.getElementById("editor").value = content; 
      
        var conChk =  document.getElementById("editor").value.indexOf('<form');  
           if (conChk !=-1){
             alert("내용에 유효하지 않은 <form 태그가 존재합니다. html을 체크해서 내용을 수정해주세요");
             return ;
           }
               
        if(confirm("임시저장하시겠습니까?")){
				document.all.mRP.value = document.all.mRP.value.replace(/\,/g,"");
				if(typeof(document.all.mPM) !="undefined"){
		    	  	if(typeof(document.all.mPM.length)!="undefined"){
			      		for(i=0;i<document.all.mPM.length;i++){
				 		 	document.all.mPM[i].value = document.all.mPM[i].value.replace(/\,/g,"");
				 			if(document.frm.mP.value ==""){
					 			document.frm.mP.value =document.all.mPM[i].value;
					 		}else{
					 			document.frm.mP.value = document.frm.mP.value+","+document.all.mPM[i].value;
					 		}
					    }
				    }else{
						document.all.mPM.value = document.all.mPM.value.replace(/\,/g,"");
						document.frm.mP.value =	document.all.mPM.value;
				    }			
			    }
			    
			    $("input[name='sFileP[]']").each( function(index,elem) {  
		            var a = $(elem).val();  
		            if( document.frm.sFile.value==""){
		     	        document.frm.sFile.value = a;
		            }else{
		                document.frm.sFile.value = document.frm.sFile.value + ","+a;
		            }
		        });  
           
			document.frm.hidRS.value = 0;
			document.frm.submit(); 
		} 
	}    

}

	function validForm(editor) {
		// Place your validation logic here

		// sample : validate that content exists
		var validator = new Trex.Validator();
		var content = editor.getContent();
		if (!validator.exists(content)) {
			alert('내용을 입력하세요');
			return false;
		}

		return true;
	}
	
	function setForm(editor) { 
        var form = editor.getForm();
        var content = editor.getContent();

        var field = document.getElementById("content");
        field.value = content;
 
        return true;
    }     
	
//내용 복사
function jsContsCopy(){ 
     window.clipboardData.setData('Text', frm["editor"].value);  
}

//문서 팝업창으로 띄우기
function jsPopView(sPage){
	 	 var winNew = window.open(sPage,"popNew","width=880, height=600,scrollbars=yes, resizable=yes");
		 winNew.focus();
	}

//문서 프린트
function jsPopModPrint(ireportidx){
	 	 var winNew = window.open("printmodeapp.asp?iridx="+ireportidx,"popNew","width=880, height=600, scrollbars=yes, resiziable=yes");
		 winNew.focus();
	}
function jsPopConfirmPrint(ireportidx){
	 	 var winNew = window.open("printconfirmeapp.asp?iridx="+ireportidx,"popNew","width=1024, height=600, scrollbars=yes, resiziable=yes");
		 winNew.focus();
	}
function jsPopMPPrint(ireportidx,ipayrequestidx){
	 var winNew = window.open("printregpayrequest.asp?iridx="+ireportidx+"&ipridx="+ipayrequestidx,"popNew","width=1024, height=600, scrollbars=yes, resiziable=yes");
		 winNew.focus();
}
function jsPopCPPrint(ireportidx,ipayrequestidx,iauthstate){
	 var winNew = window.open("printconfirmpayrequest.asp?iridx="+ireportidx+"&ipridx="+ipayrequestidx+"&ias="+iauthstate,"popNew","width=1024, height=600, scrollbars=yes, resiziable=yes");
		 winNew.focus();
}


//결재등록
	function jsEappConfirm(iState){ 
		var iAuthposition = document.frm.iRAP.value;
		var sRectAuthType = document.frm.iRAT.value;
		if (iState == 1){
			if (sRectAuthType=="L"){
			    document.frm.hidRS.value = 7; //최종승인.
			}else{
    			//if(jsChkBlank(document.frm.hidAI.value) ){
    			//	alert("결재자를 등록해주세요");
    			//	return;
    		//	}

						if(!(document.frm.hidALI.value) ){
						alert("최종결재자를 등록해주세요");
						return;
					}
		
    			//참조자와 결재자 아이디 체크
    			if(document.frm.hidRfI.value !="" && document.frm.hidAI.value != ""){
    				arrRfI = document.frm.hidRfI.value.split(",")
    				for(i=0;i<arrRfI.length;i++){
    					if(arrRfI[i] ==document.frm.hidAI.value){
    						alert("참조자와 결재자는 중복될 수 없습니다. 다시 선택해주세요");
    						return;
    					}
    				}
    			}
    			document.frm.hidRS.value = 1;
			}

            var appStr = "승인";
            if (sRectAuthType == "L"){
                appStr = "최종"+appStr;
            }else if(sRectAuthType=="A"){ //합의승인
                appStr = "합의"+appStr;
            }

			if(confirm(appStr+"하시겠습니까?")){
			    //if (document.frm.hidM_H.value=="1"){ //2013/10/28추가
			    //    document.frm.blnL.value = 1;
			    //}

    			document.frm.hidAS.value = 1;
    			document.frm.submit();
			}

		}else   {
			if(iAuthposition==1 || iState==5){
				document.frm.hidRS.value = iState;
			}else{
				document.frm.hidRS.value = 1;
			}
			var strMsg;
			if(iState==5){
				strMsg = "반려"
			}else{
				strMsg = "보류"  //3
			}

			if(confirm(strMsg+" 하시겠습니까?")){
			document.frm.hidAI.value ="";
			document.frm.hidAS.value = iState;
			document.frm.submit();
		}
		}

	}


//코멘트 삭제
function jsCommDel(commentidx){
	if(confirm("삭제하시겠습니까?")){
	document.frmCD.iCidx.value = 	commentidx;
	document.frmCD.submit();
}
}


//결제요청서 결재컨펌
function jsPayEappConfirm(iState){
	var strMsg;
	if (iState==5){
		document.frm.hidAS.value = iState;
		strMsg = "결제반려";
	}else if(iState==9){
		if(jsChkBlank(document.frm.dprld.value) ){
			alert("결제(입금)일을 입력해주세요");
			return;
		}
		document.frm.hidAS.value = 9;
		strMsg = "결제완료";
	}else if(iState==1){
	document.frm.hidAS.value = 1;
		strMsg = "결제승인";
	}else{
		strMsg = "결제확인";
	 var ichkVal=0;
	 for(i=0;i<document.frm.rdoDK.length;i++){
	 	if(document.frm.rdoDK[i].checked){
	 		 ichkVal = document.frm.rdoDK[i].value;
	 	}
	 }

	if (ichkVal ==0){
		alert("서류종류를 선택해주세요");
		return;
	}


	if(ichkVal == 1 || ichkVal == 2){
	 	if(jsChkBlank(document.frm.sEK.value)){
	 		alert("세금계산서 검색버튼을 눌러 증빙서류 내용을 등록해주세요");
	 		document.frm.btnB1.focus();
	 		return;
	 	}
		if(ichkVal  !=0 && ichkVal != 8 && ichkVal != 9){
			if(document.frm.mTP.value.replace(/\,/g,"") != document.frm.mprp.value.replace(/\,/g,"")){
		 		alert("결제요청금액과 증빙서류의 총금액이 다릅니다.확인 후 다시 등록해주세요")
		   	return;
			}
		 }
	}
		var mTotPrice = 0;

 		if(typeof(document.all.mPM) =="undefined"){
 			alert("자금구분 부서를 등록해주세요");
 			return;
 		}

		if(typeof(document.all.mPM.length)!="undefined"){
			for(i=0;i<document.all.mPM.length;i++){
				mTotPrice = mTotPrice + parseInt(document.all.mPM[i].value.replace(/\,/g,""));
			}
		}else{
			 mTotPrice = document.all.mPM.value.replace(/\,/g,"");
		}

		if(mTotPrice !=document.frm.mprp.value.replace(/\,/g,"")){
			alert(mTotPrice+"/"+document.frm.mprp.value.replace(/\,/g,"")+"부서별 자금구분금액과 결제요청금액이 다릅니다.");
			return;
		}
		if(jsChkBlank(document.frm.dPD.value) ){
			alert("결제예정일을 입력해주세요");
			return;
		}

		//if(jsChkBlank(document.frm.selOB.value) ){
		//	alert("출금은행을 선택해주세요");
		//	return;
		//}

		document.frm.hidAS.value =7;
	}

	if(typeof(document.all.mPM) !="undefined"){
			  	if(typeof(document.all.mPM.length)!="undefined"){
			  		for(i=0;i<document.all.mPM.length;i++){
					 		document.all.mPM[i].value = document.all.mPM[i].value.replace(/\,/g,"");
					 		if(document.frm.mP.value ==""){
							 			document.frm.mP.value =document.all.mPM[i].value;
							 }else{
							 			document.frm.mP.value = document.frm.mP.value+","+document.all.mPM[i].value;
							 }
						}
					}else{
							document.all.mPM.value = document.all.mPM.value.replace(/\,/g,"");
							document.frm.mP.value =	document.all.mPM.value;
					}
				}

		document.frm.mTP.value= document.frm.mTP.value.replace(/\,/g,"");
		document.frm.mSP.value= document.frm.mSP.value.replace(/\,/g,"");
		document.frm.mVP.value= document.frm.mVP.value.replace(/\,/g,"");


	if(document.frm.rdoDK[0].checked == true || document.frm.rdoDK[1].checked == true)
	{
		document.frm.rdoTD[0].checked = true;
	}


	if(confirm(strMsg+" 하시겠습니까?")){
	document.frm.hidPRS.value = iState;
	 document.frm.hidM.value ="C";
	 document.frm.submit();
	}
}

function jsonlyEngNDigit(obj)
{
	var regexp = /^[A-Za-z0-9]+$/;

	if(!regexp.test(obj))
		return false;
	else
		return true;
}

//결제요청서 등록
function jsPayEappSubmit(	sMode, reportprice ,arapcd){
		if(jsChkBlank(document.frm.dprd.value) ){
			alert("결제요청일을 입력해주세요");
			return;
		}

		if(jsChkBlank(document.frm.mprp.value) ){
			alert("결제요청금액을 입력해주세요");
			document.frm.mprp.focus();
			return;
		}

		var addPrice = reportprice*0.1;
		if(parseInt(document.frm.mprp.value.replace(/\,/g,""),0)+parseInt(document.frm.hidTP.value,0) > parseInt(reportprice,0)+parseInt(addPrice,0)){
			alert("결제요청금액이 품의금액보다 많습니다. 다시 입력해주세요\n\n결제요청은 (품의금액+품의금액의 10%)까지 요청가능합니다.");
			document.frm.mprp.value ="";
			document.frm.mprp.focus();
			return;
		}

		if(jsChkBlank(document.frm.sprt.value) ){
			alert("자금용도를 입력해주세요");
			document.frm.sprt.focus();
			return;
		}

	if(arapcd!=351){	//수지항목-비타민제도 아닐경우
		if(jsChkBlank(document.frm.hidcustcd.value) ){
			alert("거래처를 선택해주세요");
			return;
		}
	}

	 var ichkVal=0;
	 for(i=0;i<document.frm.rdoDK.length;i++){
	 	if(document.frm.rdoDK[i].checked){
	 		 ichkVal = document.frm.rdoDK[i].value;
	 	}
	 }

	if (ichkVal ==0){
		alert("서류종류를 선택해주세요");
		return;
	}


	if(ichkVal == 1 || ichkVal == 2){
	 	if(jsChkBlank(document.frm.sEK.value)){
	 		alert("세금계산서 검색버튼을 눌러 증빙서류 내용을 등록해주세요");
	 		document.frm.btnB1.focus();
	 		return;
	 	}
		if(ichkVal  !=0 && ichkVal != 8 && ichkVal != 9){
			if(document.frm.mTP.value.replace(/\,/g,"") != document.frm.mprp.value.replace(/\,/g,"")){
		 		alert("결제요청금액과 증빙서류의 총금액이 다릅니다.확인 후 다시 등록해주세요")
		   	return;
			}
		 }
	}
		var mTotPrice = 0;

    //alert(document.all.iaidx.value); //수지항목check

        if ((document.all.iaidx.value=="0")||(document.all.iaidx.value=="")){
            alert("수지 항목을 선택해주세요");
 			return;
        }

 		if(typeof(document.all.mPM) =="undefined"){
 			alert("자금구분 부서를 등록해주세요");
 			return;
 		}

		if(typeof(document.all.mPM.length)!="undefined"){
			for(i=0;i<document.all.mPM.length;i++){
				mTotPrice = mTotPrice + parseInt(document.all.mPM[i].value.replace(/\,/g,""));
			}
		}else{
			 mTotPrice = document.all.mPM.value.replace(/\,/g,"");
		}

		if(mTotPrice !=document.frm.mprp.value.replace(/\,/g,"")){
			alert(mTotPrice+"/"+document.frm.mprp.value.replace(/\,/g,"")+"부서별 자금구분금액과 결제요청금액이 다릅니다.");
			return;
		}

	if(confirm("결제요청하시겠습니까?\n국세청승인번호를 잘못 등록하거나, 수기계산서 일 경우 결재일 전날까지 증빙서류를 제출하지 않으면 결재완료가 되지 않습니다.")){
		document.frm.hidM.value =sMode;
		document.all.mprp.value = document.all.mprp.value.replace(/\,/g,"");
				if(typeof(document.all.mPM) !="undefined"){
			  	if(typeof(document.all.mPM.length)!="undefined"){
			  		for(i=0;i<document.all.mPM.length;i++){
					 		document.all.mPM[i].value = document.all.mPM[i].value.replace(/\,/g,"");
					 		if(document.frm.mP.value ==""){
							 			document.frm.mP.value =document.all.mPM[i].value;
							 }else{
							 			document.frm.mP.value = document.frm.mP.value+","+document.all.mPM[i].value;
							 }
						}
					}else{
							document.all.mPM.value = document.all.mPM.value.replace(/\,/g,"");
							document.frm.mP.value =	document.all.mPM.value;
					}
				}

		document.frm.mTP.value= document.frm.mTP.value.replace(/\,/g,"");
		document.frm.mSP.value= document.frm.mSP.value.replace(/\,/g,"");
		document.frm.mVP.value= document.frm.mVP.value.replace(/\,/g,"");
		
		$("input[name='sFileP[]']").each( function(index,elem) {  
     var a = $(elem).val();  
     if( document.frm.sFile.value==""){
     	document.frm.sFile.value = a;
    }else{
     document.frm.sFile.value = document.frm.sFile.value + ","+a;
   }
  }); 
		document.frm.hidPRS.value = 1;
		document.frm.submit();
	}
}

//자금구분 부서 등록
function jsSetPartMoney(iType, sAUCD,sACCGRP){
	var winPart = window.open("/admin/linkedERP/biz/popGetBiz.asp?iType="+iType+"&sAUCD="+sAUCD+"&sACCGRP="+sACCGRP,"popPart","width=600, height=800, resizable=yes, scrollbars=yes");
	winPart.focus();
}

//scm 링크 새창연결
function jsGoScm(sURL, iIdx){
	var winScm = window.open(sURL+iIdx,"newScm","");
	winScm.focus();
}

//관련 링크 새창 이동
function jsFileLink(sURL){
	var winNew = window.open(sURL);
	winNew.focus();
}

//서류제출여부 수정
 function jsModTakeDoc(payrequestidx,isTakeDoc){
   	var winTD = window.open("/admin/approval/eapp/popModTakeDoc.asp?ipridx="+payrequestidx+"&blnTD="+isTakeDoc,"popPart","width=600, height=200, resizable=yes, scrollbars=yes");
   	winTD.focus();
   }

//메뉴이동
	function jsGoMenuSetIdx(iRMenu,reportidx, payrequestidx){
		top.location.href = "/admin/approval/eapp/popIndex.asp?iRM="+iRMenu+"&iridx="+reportidx+"&ipridx="+payrequestidx;
	}


//금액 숫자 콤마 자동으로 넣기
function auto_amount(frm,num) {
//if (navigator.userAgent.indexOf("MSIE") != -1) {
var keyCode = event.keyCode;
 
	if ( ((keyCode >= 48) && (keyCode <=57)) || ((keyCode>=96) && (keyCode<=105))|| keyCode ==13 || keyCode==36 || keyCode==35 || keyCode==46 || keyCode==8 || keyCode==109 || keyCode==189) {
	//숫자키, 엔터키, 홈키, 앤더키, 백스페이스키, 딜리터키 일 때
	var str=num.value;
	str=str_number(str); //숫자가 아닌 문자 제거
  
		if ( (str != null) && (str != "") && (str != "0") ) {
		//3자리마다 콤마넣기
		//str = parseInt(str,10);//십진수로 변환
			var str = "" + str;
			var objRegExp = new RegExp("(-?[0-9]+)([0-9]{3})");
				while (objRegExp.test(str)) {
				str = str.replace(objRegExp, "$1,$2");
				}

			if(str=="NaN"){
			    str=0; 
			}else if(str=="-NaN"){
			    str="-";
			  } 
			num.value = str;
		} else {
			num.value = "0";
		}
	}
//} else {
//return false;
//}
}
function str_number(str) {
	//문자열에서 숫자만 가져가기
	var val = str;
	var temp = "";
	var num = "";
	var i = 0;
	var chkm = "";
	if (val.substr(0,1)=="-"){
	    chkm = ""+"-";
	}
	
	for (i=0; i<val.length; i++) {
	temp = ""+val.substr(i,1);
 
		if ( (temp >= "0" && temp <= "9") ) {
		num = num + temp;
		}
	}
	num = ""+parseInt(num,10);//십진수로 변환
	return chkm+num;
}
function num_check() {
//숫자만 입력받기.
//if (navigator.userAgent.indexOf("MSIE") != -1) {
	var keyCode = event.keyCode;
 
	if ((keyCode < 48 && keyCode !=45) || keyCode > 57  ) {
	event.returnValue=false;
	}
//}
}

function jsSetComma(str){
	if ( (str != null) && (str != "") && (str != "0") ) {
//3자리마다 콤마넣기
//str = parseInt(str,10);//십진수로 변환
var str = "" + str;
var objRegExp = new RegExp("(-?[0-9]+)([0-9]{3})");
while (objRegExp.test(str)) {
str = str.replace(objRegExp, "$1,$2");
}
}

return str;
}


function jsNewRegXML(){
    var winD = window.open("/admin/tax/popRegfileXML.asp","popDXML","width=600, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}


function jsNewRegHand(){
    var winD = window.open("/admin/tax/popRegfileHand.asp","popDHand","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}


function jsGetTax(){
var sSearchText = document.frm.sbizno.value;
if(sSearchText=="undefined"){sSearchText=""};
var totSum = document.frm.mprp.value.replace(/\,/g,"");
var winTax = window.open("/admin/tax/popSetEseroTax.asp?sST="+sSearchText+"&totSum="+totSum,"popT","width=1200, height=800, resizable=yes, scrollbars=yes");
winTax.focus();
}



  //총급액 입력으로 공급가, 부가세 변경
  function jsSetPrice(){
  	if(document.frm.rdoVK[0].checked){
  		document.frm.mSP.value= 	parseInt((document.frm.mTP.value/1.1).toFixed(5)) ;
  		document.frm.mVP.value= document.frm.mTP.value-document.frm.mSP.value;
  	}else{
  		document.frm.mSP.value= document.frm.mTP.value;
  		document.frm.mVP.value= 0;
  	}
  }

//수지항목 불러오기
 	function jsGetARAP(){
 			var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp","popARAP","width=800,height=600,resizable=yes, scrollbars=yes");
 			winARAP.focus();
 	}

 	//선택 수지항목 가져오기
 	function jsSetARAP(dAC, sANM,sACC,sACCNM){
 		document.frm.iaidx.value = dAC;
 		document.frm.sANM.value = "["+dAC+"]"+sANM;
 		document.frm.sACC.value = sACC;
 		document.frm.sACCNM.value = "["+sACC+"]"+sACCNM;
 	}


 		//서류종류에 따라 disable처리
	function jsSetDocDis(iTypeValue){
		document.frm.dID.value= "";
		document.frm.sINm.value= "";
		document.frm.mTP.value= "";
		document.frm.mSP.value= "";
		document.frm.mVP.value= "";
	  if(iTypeValue==1){
			document.all.dSel1.style.display = "";
			document.all.dView1.style.display = "none";
			//document.all.spB2.style.display ="";
			document.all.spB3.style.display ="none";
		}else if(iTypeValue==2){
			document.all.dSel1.style.display = "";
			document.all.dView1.style.display = "none";
			//document.all.spB2.style.display ="none";
			document.all.spB3.style.display ="";
		}else{
			document.all.dSel1.style.display = "none";
			document.all.dView1.style.display = "none";
		}
	}


		//결제방법에 따른 화폐등록 view 여부
	function jsChFC(){
		if(document.frm.selPT.options[document.frm.selPT.selectedIndex].value == 1){
			document.all.spCurr.style.display = "";
		}else{
			document.all.spCurr.style.display = "none";
		};
	}



 	//거래처 정보 보기
	function jsGetCust(cust_cd){
		var Strparm="";
		if (cust_cd!=""){
			Strparm = "?selSTp=1&sSTx="+ cust_cd;
		}
		var winC = window.open("/admin/linkedERP/cust/popGetCust.asp"+Strparm,"popC","width=1200, height=600,resizable=yes, scrollbars=yes");
		winC.focus();
	}

	 //거래처 선택
   function jsSetCust(custcd, custnm,banknm, accno, snm, sbizno ){
   document.frm.hidcustcd.value = custcd;
   document.frm.scustnm.value = custnm;
   document.frm.selIB.value = banknm;
   document.frm.san.value = accno;
   document.frm.sah.value = snm;
   document.frm.sbizno.value = sbizno;
  }

