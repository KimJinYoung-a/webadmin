/* ���ڰ��� ���� ��ũ��Ʈ */

//�����������
function popDecision(){
    var popwin = window.open('/admin/approval/eapp/epop/popDecision.asp','popDecision','width=900, height=900, scrollbars=yes,resizable=yes');
}

//����÷��
function jsAttachFile(sP){
	var winAF = window.open('/admin/approval/eapp/popRegFile.asp?sp='+sP,'popAF','width=400, height=300');
	winAF.focus();
}

//���ϻ���
function jsFileDel(sName){
	$("#dF"+sName).remove(); 
}

//���� �ٿ�ε�
    function jsDownload(sDownURL, sRFN, sFN){
    var winFD = window.open(sDownURL+"/linkweb/eapp/procDownload.asp?sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
 }


//�μ��� �ڱݱ��� �ݾ׺�, �ۼ�Ʈ�� �� �ڵ� ����
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

	if(iPageType==1){ // ǰ�Ǽ���
		mRequestPay = document.all.mRP.value.replace(/\,/g,"");
		if(mRequestPay == "" || mRequestPay == 0){
			alert("ǰ�Ǳݾ��� ���� �Է����ּ���");
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
	}else{//������û����
		 mRequestPay = document.all.mprp.value.replace(/\,/g,"");
		if(mRequestPay == "" || mRequestPay == 0){
				alert("������û�ݾ��� ���� �Է����ּ���");
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


//�޷º���
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

// ������ �̵�
function jsGoPage(iCP)
	{
		document.frmList.iCP.value=iCP;
		document.frmList.submit();
	}



// tr ���󺯰�
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


//�������, ���� ���̵� ���
function jsRegID(iMode){ 
		var ilastApprovalid = document.frm.iLAID.value;
		var idpid = document.frm.hidPS.value;
		var icid1 = document.frm.hidcid1.value;
		var icid2 = document.frm.hidcid2.value;
		var icid3 = document.frm.hidcid3.value;
		var icid4 = document.frm.hidcid4.value;
		var winRI = window.open('/admin/approval/eapp/epop/popSetAuthID.asp?iM='+iMode+'&iLAID='+ilastApprovalid+'&sjn='+document.frm.hidJN.value+'&idpid='+idpid+'&icid1='+icid1+'&icid2='+icid2+'&icid3='+icid3+'&icid4='+icid4 ,'popAL','width=800, height=550, resizable=yes, scrollbars=yes');
		winRI.focus();
}

//������ �߰� 2013/10/21
function jsRegID_H(iMode){
		var iAuthPosition= '0';
		var ilastApprovalid = document.frm.hidAI_H.value;
		var idpid = document.frm.hidPS_H.value;	
		if ((idpid=='')||(idpid==undefined)) {
		     idpid = '8';
		}

		var winRI = window.open('/admin/approval/eapp/epop/popSetAuthID.asp?iM='+iMode+'&iLAI='+ilastApprovalid+'&iAP='+iAuthPosition+'&idpid='+idpid ,'popAL','width=650, height=550, resizable=yes, scrollbars=yes');
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
			alert("�����׸���\n[172] ��ǰ(����) �Ǵ� [173] �繫��/�������/���׸������\n�� ��� ǰ�Ǳݾ��� 100���� �̸����� ��쿣\n�ٸ� �����׸��� �����ؾ� �մϴ�.\n\n�̿� ���� ���Ǵ� �濵�������� �����Ͻñ� �ٶ��ϴ�.");
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


//������(ǰ�Ǽ�)
function jsEappSubmit(iState){    
	var arrRfI;	 
  	var mRequestPay = document.all.mRP.value.replace(/\,/g,""); 
	if(!jsIsHundred()){
		return;
	}
//	if (document.frm.tmpisAgreeNeed.value == "Y"){
//		if(document.frm.tmpisAgreeNeedTarget.value != document.frm.hidAI_H.value){
//			alert('�����ڸ� ������ �� ���� �����Դϴ�.');
//			return;
//		}
//	}

	if (iState == 1){
	    //if(jsChkBlank(document.frm.hidAI.value) ){
			//alert("�����ڸ� ������ּ���");
			//return;
		//}

		if (document.frm.tmpisAgreeNeed.value == "Y" && jsChkBlank(document.frm.hidAI_H.value) && jsChkBlank(document.frm.hidAHI.value)){
			alert('�����ڸ� ������ּ���');
			return;			
		}

		if(jsChkBlank(document.frm.hidALI.value) && jsChkBlank(document.frm.hidAHI.value)){
			alert("���������ڸ� ������ּ���.");
			return;
		}
		
		if(jsChkBlank(document.frm.sRN.value) ){
			alert("ǰ�Ǽ����� �Է����ּ���");
			return;
		}

        //2013/10/28�߰�
        if ((document.all.hidPE.value=="True")&&(mRequestPay == "" || mRequestPay == 0)){
			alert("ǰ�Ǳݾ��� ���� �Է����ּ���");
			document.all.mRP.focus();
			return;
		}

        if(!jsIsHundred()){
    		return;
    	}

		//�����ھ��̵�� ������, ������ �ߺ� Ȯ��
		if (document.frm.hidAI.value != ""){
			var	arrAI  = document.frm.hidAI.value.split(",");
			var sLastAI = document.frm.hidALI.value;
			var sLastAH = document.frm.hidAHI.value;
				for(j=0;j<arrAI.length;j++){
					if(document.frm.hidAI_H.value !=""){
						if(document.frm.hidAI_H.value ==arrAI[j] || document.frm.hidAI_H.value == sLastAI || document.frm.hidAI_H.value == sLastAH){
						alert("�����ڿ� �����ڴ� �ߺ��� �� �����ϴ�. �ٽ� �������ּ���");
						return;
						} 
					}
					
					if(sLastAI ==arrAI[j] || sLastAH ==arrAI[j] ){
						alert("���������ڿ� �����ڴ� �ߺ��� �� �����ϴ�. �ٽ� �������ּ���");
						return;
					}
						
					if(document.frm.hidRfI.value !=""){
							arrRfI = document.frm.hidRfI.value.split(",");
							for(i=0;i<arrRfI.length;i++){
								if(arrRfI[i] ==arrAI[j]|| arrRfI[i] == sLastAI){
									alert("�����ڿ� �����ڴ� �ߺ��� �� �����ϴ�. �ٽ� �������ּ���");
									return;
								}
							}
					} 
				}
			
		}

		var totPM = 0;
		if(document.frm.hidPE.value=="True"){  // iAIdx => iaidx
		    if (document.all.iP){ //�����߰�
    			if(jsChkBlank(document.all.iP.value) ){
    				alert("�μ��� ������ּ���");
    				return;
    			}

    			if(jsChkBlank(mRequestPay) ){
    				alert("�μ��� ������ּ���");
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
    					alert("�ڱݱ��� �ݾ��� ǰ�Ǳݾװ� �ٸ��ϴ�. �缳�����ּ���");
    					return;
    				}
    			}
		    }
		}

 
    	 var content = Editor.getContent();
         document.getElementById("editor").value = content; 
         
        var conChk =  document.getElementById("editor").value.indexOf('<form');  
           if (conChk !=-1){
             alert("���뿡 ��ȿ���� ���� <form �±װ� �����մϴ�. html�� üũ�ؼ� ������ �������ּ���");
             return ;
           }

    	if(confirm("�������Ͻðڽ��ϱ�?")){
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

			// ��Ϲ�ư �Ͻ� ����
			$("#btnSm").attr('disabled',true).css("color","#CCC");
			setTimeout(function() {
				$("#btnSm").attr('disabled',false).css("color","blue");
			}, 5000);

			document.frm.hidRS.value = 1;
			document.frm.submit();
    	}
	}else if(iState ==-1) {
		if(confirm("�����Ͻðڽ��ϱ�?")){
	    	document.frm.hidM.value = "D";
			document.frm.submit();
		}
	}else  if(iState ==0) {   
	    
        var content = Editor.getContent();
         document.getElementById("editor").value = content; 
      
        var conChk =  document.getElementById("editor").value.indexOf('<form');  
           if (conChk !=-1){
             alert("���뿡 ��ȿ���� ���� <form �±װ� �����մϴ�. html�� üũ�ؼ� ������ �������ּ���");
             return ;
           }
               
        if(confirm("�ӽ������Ͻðڽ��ϱ�?")){
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
			alert('������ �Է��ϼ���');
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
	
//���� ����
function jsContsCopy(){ 
     window.clipboardData.setData('Text', frm["editor"].value);  
}

//���� �˾�â���� ����
function jsPopView(sPage){
	 	 var winNew = window.open(sPage,"popNew","width=880, height=600,scrollbars=yes, resizable=yes");
		 winNew.focus();
	}

//���� ����Ʈ
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


//������
	function jsEappConfirm(iState){ 
		var iAuthposition = document.frm.iRAP.value;
		var sRectAuthType = document.frm.iRAT.value;
		if (iState == 1){
			if (sRectAuthType=="L" || sRectAuthType=="F"){
			    document.frm.hidRS.value = 7; //��������.
			}else{
    			//if(jsChkBlank(document.frm.hidAI.value) ){
    			//	alert("�����ڸ� ������ּ���");
    			//	return;
    		//	}

					if(jsChkBlank(document.frm.hidALI.value) && jsChkBlank(document.frm.hidAHI.value)){
						alert("���������ڸ� ������ּ���");
						return;
					}
		
    			//�����ڿ� ������ ���̵� üũ
    			if(document.frm.hidRfI.value !="" && document.frm.hidAI.value != ""){
    				arrRfI = document.frm.hidRfI.value.split(",")
    				for(i=0;i<arrRfI.length;i++){
    					if(arrRfI[i] ==document.frm.hidAI.value){
    						alert("�����ڿ� �����ڴ� �ߺ��� �� �����ϴ�. �ٽ� �������ּ���");
    						return;
    					}
    				}
    			}
    			document.frm.hidRS.value = 1;
			}

            var appStr = "����";
            if (sRectAuthType=="L"){
                appStr = "����"+appStr;
            }else if(sRectAuthType=="F"){ //�ּ����ǽ���
                appStr = "��������"+appStr;
            }else if(sRectAuthType=="A"){ //���ǽ���
                appStr = "����"+appStr;
            }

			if(confirm(appStr+"�Ͻðڽ��ϱ�?")){
			    //if (document.frm.hidM_H.value=="1"){ //2013/10/28�߰�
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
				strMsg = "�ݷ�"
			}else{
				strMsg = "����"  //3
			}

			if(confirm(strMsg+" �Ͻðڽ��ϱ�?")){
			document.frm.hidAI.value ="";
			document.frm.hidAS.value = iState;
			document.frm.submit();
		}
		}

	}


//�ڸ�Ʈ ����
function jsCommDel(commentidx){
	if(confirm("�����Ͻðڽ��ϱ�?")){
	document.frmCD.iCidx.value = 	commentidx;
	document.frmCD.submit();
}
}


//������û�� ��������
function jsPayEappConfirm(iState){
	var strMsg;
	if (iState==5){
		document.frm.hidAS.value = iState;
		strMsg = "�����ݷ�";
	}else if(iState==9){
		if(jsChkBlank(document.frm.dprld.value) ){
			alert("����(�Ա�)���� �Է����ּ���");
			return;
		}
		document.frm.hidAS.value = 9;
		strMsg = "�����Ϸ�";
	}else if(iState==1){
	document.frm.hidAS.value = 1;
		strMsg = "��������";
	}else{
		strMsg = "����Ȯ��";
	 var ichkVal=0;
	 for(i=0;i<document.frm.rdoDK.length;i++){
	 	if(document.frm.rdoDK[i].checked){
	 		 ichkVal = document.frm.rdoDK[i].value;
	 	}
	 }

	if (ichkVal ==0){
		alert("���������� �������ּ���");
		return;
	}


	if(ichkVal == 1 || ichkVal == 2){
	 	if(jsChkBlank(document.frm.sEK.value)){
	 		alert("���ݰ�꼭 �˻���ư�� ���� �������� ������ ������ּ���");
	 		document.frm.btnB1.focus();
	 		return;
	 	}
		if(ichkVal  !=0 && ichkVal != 8 && ichkVal != 9){
			if(document.frm.mTP.value.replace(/\,/g,"") != document.frm.mprp.value.replace(/\,/g,"")){
		 		alert("������û�ݾװ� ���������� �ѱݾ��� �ٸ��ϴ�.Ȯ�� �� �ٽ� ������ּ���")
		   	return;
			}
		 }
	}
		var mTotPrice = 0;

 		if(typeof(document.all.mPM) =="undefined"){
 			alert("�ڱݱ��� �μ��� ������ּ���");
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
			alert(mTotPrice+"/"+document.frm.mprp.value.replace(/\,/g,"")+"�μ��� �ڱݱ��бݾװ� ������û�ݾ��� �ٸ��ϴ�.");
			return;
		}
		if(jsChkBlank(document.frm.dPD.value) ){
			alert("������������ �Է����ּ���");
			return;
		}

		//if(jsChkBlank(document.frm.selOB.value) ){
		//	alert("��������� �������ּ���");
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


	if(confirm(strMsg+" �Ͻðڽ��ϱ�?")){
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

//������û�� ���
function jsPayEappSubmit(	sMode, reportprice ,arapcd){ 
 
	 
	  
		if(jsChkBlank(document.frm.dprd.value) ){
			alert("������û���� �Է����ּ���");
			return;
		}

		if(jsChkBlank(document.frm.mprp.value) ){
			alert("������û�ݾ��� �Է����ּ���");
			document.frm.mprp.focus();
			return;
		}

		var addPrice = reportprice*0.1;
		if(parseInt(document.frm.mprp.value.replace(/\,/g,""),0)+parseInt(document.frm.hidTP.value,0) > parseInt(reportprice,0)+parseInt(addPrice,0)){
			alert("������û�ݾ��� ǰ�Ǳݾ׺��� �����ϴ�. �ٽ� �Է����ּ���\n\n������û�� (ǰ�Ǳݾ�+ǰ�Ǳݾ��� 10%)���� ��û�����մϴ�.");
			document.frm.mprp.value ="";
			document.frm.mprp.focus();
			return;
		}
		
	if (document.frm.selPT.options[document.frm.selPT.selectedIndex].value ==2){//������ü 
	  	 if(!document.frm.san.value){
	  	 	alert("��������� ������ü�� ��� ���¹�ȣ�� �Է����ּ���");
	  	 	return;
	  	 }
	  } 

		if(jsChkBlank(document.frm.sprt.value) ){
			alert("�ڱݿ뵵�� �Է����ּ���");
			document.frm.sprt.focus();
			return;
		}

	if(arapcd!=351){	//�����׸�-��Ÿ������ �ƴҰ��
		if(jsChkBlank(document.frm.hidcustcd.value) ){
			alert("�ŷ�ó�� �������ּ���");
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
		alert("���������� �������ּ���");
		return;
	}


	if(ichkVal == 1 || ichkVal == 2){
	 	if(jsChkBlank(document.frm.sEK.value)){
	 		alert("���ݰ�꼭 �˻���ư�� ���� �������� ������ ������ּ���");
	 		document.frm.btnB1.focus();
	 		return;
	 	}
		if(ichkVal  !=0 && ichkVal != 8 && ichkVal != 9){
			if(document.frm.mTP.value.replace(/\,/g,"") != document.frm.mprp.value.replace(/\,/g,"")){
		 		alert("������û�ݾװ� ���������� �ѱݾ��� �ٸ��ϴ�.Ȯ�� �� �ٽ� ������ּ���")
		   	return;
			}
		 }
	}
		var mTotPrice = 0;

    //alert(document.all.iaidx.value); //�����׸�check

        if ((document.all.iaidx.value=="0")||(document.all.iaidx.value=="")){
            alert("���� �׸��� �������ּ���");
 			return;
        }

 		if(typeof(document.all.mPM) =="undefined"){
 			alert("�ڱݱ��� �μ��� ������ּ���");
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
			alert(mTotPrice+"/"+document.frm.mprp.value.replace(/\,/g,"")+"�μ��� �ڱݱ��бݾװ� ������û�ݾ��� �ٸ��ϴ�.");
			return;
		}

		if(confirm("������û�Ͻðڽ��ϱ�?\n����û���ι�ȣ�� �߸� ����ϰų�, �����꼭 �� ��� ������ �������� ���������� �������� ������ ����Ϸᰡ ���� �ʽ��ϴ�.")){
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
			}
		); 

		// ��Ϲ�ư �Ͻ� ����
		$("#btnSm").attr('disabled',true).css("color","#CCC");
		setTimeout(function() {
			$("#btnSm").attr('disabled',false).css("color","blue");
		}, 5000);

		document.frm.hidPRS.value = 1;
		document.frm.submit();
	}
}

//�ڱݱ��� �μ� ���
function jsSetPartMoney(iType, sAUCD,sACCGRP){
	var winPart = window.open("/admin/linkedERP/biz/popGetBiz.asp?iType="+iType+"&sAUCD="+sAUCD+"&sACCGRP="+sACCGRP,"popPart","width=600, height=800, resizable=yes, scrollbars=yes");
	winPart.focus();
}

//scm ��ũ ��â����
function jsGoScm(sURL, iIdx){
	var winScm = window.open(sURL+iIdx,"newScm","");
	winScm.focus();
}

//���� ��ũ ��â �̵�
function jsFileLink(sURL){
	var winNew = window.open(sURL);
	winNew.focus();
}

//�������⿩�� ����
 function jsModTakeDoc(payrequestidx,isTakeDoc){
   	var winTD = window.open("/admin/approval/eapp/popModTakeDoc.asp?ipridx="+payrequestidx+"&blnTD="+isTakeDoc,"popPart","width=600, height=200, resizable=yes, scrollbars=yes");
   	winTD.focus();
   }

//�޴��̵�
	function jsGoMenuSetIdx(iRMenu,reportidx, payrequestidx){
		top.location.href = "/admin/approval/eapp/popIndex.asp?iRM="+iRMenu+"&iridx="+reportidx+"&ipridx="+payrequestidx;
	}


//�ݾ� ���� �޸� �ڵ����� �ֱ�
function auto_amount(frm,num) {
//if (navigator.userAgent.indexOf("MSIE") != -1) {
var keyCode = event.keyCode;
 
	if ( ((keyCode >= 48) && (keyCode <=57)) || ((keyCode>=96) && (keyCode<=105))|| keyCode ==13 || keyCode==36 || keyCode==35 || keyCode==46 || keyCode==8 || keyCode==109 || keyCode==189) {
	//����Ű, ����Ű, ȨŰ, �ش�Ű, �齺���̽�Ű, ������Ű �� ��
	var str=num.value;
	str=str_number(str); //���ڰ� �ƴ� ���� ����
  
		if ( (str != null) && (str != "") && (str != "0") ) {
		//3�ڸ����� �޸��ֱ�
		//str = parseInt(str,10);//�������� ��ȯ
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
	//���ڿ����� ���ڸ� ��������
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
	num = ""+parseInt(num,10);//�������� ��ȯ
	return chkm+num;
}
function num_check() {
//���ڸ� �Է¹ޱ�.
//if (navigator.userAgent.indexOf("MSIE") != -1) {
	var keyCode = event.keyCode;
 
	if ((keyCode < 48 && keyCode !=45) || keyCode > 57  ) {
	event.returnValue=false;
	}
//}
}

function jsSetComma(str){
	if ( (str != null) && (str != "") && (str != "0") ) {
//3�ڸ����� �޸��ֱ�
//str = parseInt(str,10);//�������� ��ȯ
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



  //�ѱ޾� �Է����� ���ް�, �ΰ��� ����
  function jsSetPrice(){
  	if(document.frm.rdoVK[0].checked){
  		document.frm.mSP.value= 	parseInt((document.frm.mTP.value/1.1).toFixed(5)) ;
  		document.frm.mVP.value= document.frm.mTP.value-document.frm.mSP.value;
  	}else{
  		document.frm.mSP.value= document.frm.mTP.value;
  		document.frm.mVP.value= 0;
  	}
  }

//�����׸� �ҷ�����
 	function jsGetARAP(){
 			var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp","popARAP","width=800,height=600,resizable=yes, scrollbars=yes");
 			winARAP.focus();
 	}

 	//���� �����׸� ��������
 	function jsSetARAP(dAC, sANM,sACC,sACCNM){
 		document.frm.iaidx.value = dAC;
 		document.frm.sANM.value = "["+dAC+"]"+sANM;
 		document.frm.sACC.value = sACC;
 		document.frm.sACCNM.value = "["+sACC+"]"+sACCNM;
 	}


 		//���������� ���� disableó��
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


		//��������� ���� ȭ���� view ����
	function jsChFC(){
		if(document.frm.selPT.options[document.frm.selPT.selectedIndex].value == 1){
			document.all.spCurr.style.display = "";
		}else{
			document.all.spCurr.style.display = "none";
		};
	}



 	//�ŷ�ó ���� ����
	function jsGetCust(cust_cd){
		var Strparm="";
		if (cust_cd!=""){
			Strparm = "?selSTp=1&sSTx="+ cust_cd;
		}
		var winC = window.open("/admin/linkedERP/cust/popGetCust.asp"+Strparm,"popC","width=1200, height=600,resizable=yes, scrollbars=yes");
		winC.focus();
	}

	 //�ŷ�ó ����
   function jsSetCust(custcd, custnm,banknm, accno, snm, sbizno ){
   document.frm.hidcustcd.value = custcd;
   document.frm.scustnm.value = custnm;
   document.frm.selIB.value = banknm;
   document.frm.san.value = accno;
   document.frm.sah.value = snm;
   document.frm.sbizno.value = sbizno;
  }

	
	function jsTexSetting(){
		$("#sprt").val($("#sINm").val());
		if($("#sprt").val() != ""){
			$("#sprt").prop('readonly', true);
		}else{
			$("#sprt").prop('readonly', false);
		}
	}