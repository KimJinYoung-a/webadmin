<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결재라인 등록
' History : 2011.03.16 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
dim oMember, arrList,intLoop
dim sUsername, oldpart_sn
Dim imgJoin,part_sn
Dim iLastApprovalID,sjob_name
Dim pid, oldpid
Dim idepartment_id, user_cid1, user_cid2, user_cid3, user_cid4
idepartment_id =  requestCheckvar(Request("idpid"),10)
user_cid1=  requestCheckvar(Request("icid1"),10)
user_cid2=  requestCheckvar(Request("icid2"),10)
user_cid3=  requestCheckvar(Request("icid3"),10)
user_cid4=  requestCheckvar(Request("icid4"),10)
sUsername =  requestCheckvar(Request("sUN"),32)
iLastApprovalID =requestCheckvar(Request("iLAID"),10)
sjob_name	 =requestCheckvar(Request("sjn"),32)
 
'사원리스트 가져오기
	Set oMember = new CTenByTenMember
	oMember.Fusername 		= sUsername
	arrList = oMember.fnGetUserTreeListNew
	set oMember = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"> </script> 
<script type="text/javascript">
	//트리모양 클릭에 따라 변경
	function jsOpenClose(cValue,iValue){ 
		if(eval("document.all.divB"+cValue+iValue).style.display=="none"){
			eval("document.all.Fimg"+iValue).src = "/images/dtree/openfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tminus.png";
			eval("document.all.divB"+cValue+iValue).style.display=""; 
		}else{
			eval("document.all.Fimg"+iValue).src = "/images/dtree/closedfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tplus.png";
			eval("document.all.divB"+cValue+iValue).style.display="none"; 
		}
	}
	
	
	//사원선택
	function  jsSetId(iValue, sUid, sUname, ijobsn,sJobname){  
	 if(document.all.divU.length== undefined ){  
		if (eval("document.all.divU").style.background =="white"){
				eval("document.all.divU").style.background = "yellow";
				if (document.frmS.hidUI.value==""){
					document.frmS.hidUI.value = sUid;
					document.frmS.hidUN.value = sUname;
					document.frmS.hidJS.value = ijobsn;
					document.frmS.hidJN.value = sJobname;
				}else{
					document.frmS.hidUI.value = document.frmS.hidUI.value +','+ sUid;
					document.frmS.hidUN.value = document.frmS.hidUN.value +','+ sUname;
					document.frmS.hidJS.value = document.frmS.hidJS.value + ',' + ijobsn;
					document.frmS.hidJN.value = document.frmS.hidJN.value +','+ sJobname;
			 }
		}else{
			eval("document.all.divU").style.background = "white";
		var arrUI = document.frmS.hidUI.value.split(",");
		var arrUN = document.frmS.hidUN.value.split(","); 
		var arrJS = document.frmS.hidJS.value.split(","); 
		var arrJN = document.frmS.hidJN.value.split(",");
		 document.frmS.hidUI.value ="";
		 document.frmS.hidUN.value =""; 
		 document.frmS.hidJS.value = "";
		 document.frmS.hidJN.value ="";
	 		for(i=0;i<arrUI.length;i++){ 
	  			if(arrUI[i]!=sUid){
						if(document.frmS.hidUI.value==""){
							document.frmS.hidUI.value = arrUI[i];
							document.frmS.hidUN.value = arrUN[i];
							document.frmS.hidJS.value = arrJS[i];
							document.frmS.hidJN.value = arrJN[i];
						}else{
						document.frmS.hidUI.value = document.frmS.hidUI.value +','+ arrUI[i];
						document.frmS.hidUN.value = document.frmS.hidUN.value +','+ arrUN[i];
						document.frmS.hidJS.value = document.frmS.hidJS.value + ',' + arrJS[i];
						document.frmS.hidJN.value = document.frmS.hidJN.value +','+ arrJN[i];
						}
	 			}
 			}
		}
	}else{
		if (eval("document.all.divU["+iValue+"]").style.background =="white"){
				eval("document.all.divU["+iValue+"]").style.background = "yellow";
				if (document.frmS.hidUI.value==""){
					document.frmS.hidUI.value = sUid;
					document.frmS.hidUN.value = sUname;
					document.frmS.hidJS.value = ijobsn;
					document.frmS.hidJN.value = sJobname;
				}else{
					document.frmS.hidUI.value = document.frmS.hidUI.value +','+ sUid;
					document.frmS.hidUN.value = document.frmS.hidUN.value +','+ sUname;
					document.frmS.hidJS.value = document.frmS.hidJS.value + ',' + ijobsn;
					document.frmS.hidJN.value = document.frmS.hidJN.value +','+ sJobname;
			 }
		}else{
			eval("document.all.divU["+iValue+"]").style.background = "white";
		var arrUI = document.frmS.hidUI.value.split(",");
		var arrUN = document.frmS.hidUN.value.split(","); 
		var arrJS = document.frmS.hidJS.value.split(","); 
		var arrJN = document.frmS.hidJN.value.split(",");
		 document.frmS.hidUI.value ="";
		 document.frmS.hidUN.value =""; 
		 document.frmS.hidJS.value = "";
		 document.frmS.hidJN.value ="";
	 		for(i=0;i<arrUI.length;i++){ 
	  			if(arrUI[i]!=sUid){
						if(document.frmS.hidUI.value==""){
							document.frmS.hidUI.value = arrUI[i];
							document.frmS.hidUN.value = arrUN[i];
							document.frmS.hidJS.value = arrJS[i];
							document.frmS.hidJN.value = arrJN[i];
						}else{
						document.frmS.hidUI.value = document.frmS.hidUI.value +','+ arrUI[i];
						document.frmS.hidUN.value = document.frmS.hidUN.value +','+ arrUN[i];
						document.frmS.hidJS.value = document.frmS.hidJS.value + ',' + arrJS[i];
						document.frmS.hidJN.value = document.frmS.hidJN.value +','+ arrJN[i];
						}
	 			}
 			}
		}
	}		 
}
	
	//선택사원 결재(itype =1), 합의(itype =2), 참조(itype =3), 최종결재(itype=4) 에 각각등록
	function jsProcId(iType){
		var chkU;
		var arrUI = document.frmS.hidUI.value.split(",");
		var arrUN = document.frmS.hidUN.value.split(","); 
		var arrJS = document.frmS.hidJS.value.split(",");
		var arrJN = document.frmS.hidJN.value.split(",");
			
		if(iType==2){
			if(arrUI.length > 1 || $("#selAU option").size()>0){
				alert("합의는 한명만 선택가능합니다.");
				return;
			}else{ 
				if(document.frmS.hidUI.value!=""){
					$("#selAU").append("<option value='"+document.frmS.hidUI.value+"-"+document.frmS.hidJS.value+"'>"+document.frmS.hidUN.value+" "+ document.frmS.hidJN.value+" ["+document.frmS.hidUI.value+"]</option>")
				}
			 }
		}else if(iType==3){
			for(i=0; i<arrUI.length;i++){
				chkU = 0;
				for(j=0; j<$("#selCU option").size();j++){ 
					if(($("#selCU option:eq("+j+")").val().split("-"))[0]==arrUI[i]){
					chkU = 1;
					}
				}
		 
				if (chkU ==0){
					if(arrUI[i]!=""){
					$("#selCU").append("<option value='"+arrUI[i]+"-"+arrJS[i]+"'>"+arrUN[i]+" "+arrJN[i]+" ["+arrUI[i]+"]</option>")
				}
				}
			}
		}else if(iType==4){
			if(arrUI.length > 1 || $("#selLPU option").size()>0){
				alert("최종결재는 한명만 선택가능합니다.");
				return;
			}else{ 
				if(document.frmS.hidUI.value!=""){
					$("#selLPU").append("<option value='"+document.frmS.hidUI.value+"-"+document.frmS.hidJS.value+"'>"+document.frmS.hidUN.value+" "+ document.frmS.hidJN.value+" ["+document.frmS.hidUI.value+"]</option>")
				}
			 }
			 $("#tRowLHU").hide();
		}else if(iType==5){
			if(arrUI.length > 1 || $("#selLHU option").size()>0){
				alert("최종합의는 한명만 선택가능합니다.");
				return;
			}else{ 
				if(document.frmS.hidUI.value!=""){
					$("#selLHU").append("<option value='"+document.frmS.hidUI.value+"-"+document.frmS.hidJS.value+"'>"+document.frmS.hidUN.value+" "+ document.frmS.hidJN.value+" ["+document.frmS.hidUI.value+"]</option>")
				}
			 }
			 $("#tRowLPU").hide();
		}else{
			
			for(i=0; i<arrUI.length;i++){
				chkU = 0;
				for(j=0; j<$("#selPU option").size();j++){ 
					if(($("#selPU option:eq("+j+")").val().split("-"))[0]==arrUI[i]){
					chkU = 1;
					}
				} 
				if (chkU ==0){
					if(arrUI[i]!=""){
					$("#selPU").append("<option value='"+arrUI[i]+"-"+arrJS[i]+"'>"+arrUN[i]+" "+arrJN[i]+" ["+arrUI[i]+"]</option>")
				}
				}
			}
		 }
		 
		 for(i=0;i<document.all.divU.length;i++){	
		 	document.all.divU[i].style.background = "white";
		 }	
		 document.frmS.hidUI.value ="";
		 document.frmS.hidUN.value =""; 
		 document.frmS.hidJS.value =""; 
		 document.frmS.hidJN.value ="";
	}
	
	//선택삭제
	function jsSelDel(iType){
		if (iType==2){
		 $("#selAU option:selected").remove();
		}else if(iType==3){
			$("#selCU option:selected").remove();
		}else if(iType==4){
			$("#selLPU option:selected").remove();
			if($("#selLPU option").size()==0) {
				$("#tRowLHU").show();
			}
		}else if(iType==5){
			$("#selLHU option:selected").remove();
			if($("#selLHU option").size()==0) {
				$("#tRowLPU").show();
			}
		}else{
			$("#selPU option:selected").remove();
		}
	}
	
 	// 위로 올리기.
 function jsSelectMoveUp(){
 		var opt = selPU.options.selectedIndex;  
     if(opt> 0) {
     	var tmpText = selPU.options[opt].text;
     	var tmpValue = selPU.options[opt].value;
     	selPU.options[opt].text= selPU.options[opt-1].text;
     	selPU.options[opt].value= selPU.options[opt-1].value;
     	selPU.options[opt-1].text= tmpText;
     	selPU.options[opt-1].value= tmpValue;
     	 selPU.options.selectedIndex = opt-1; 
     } 
} 

	// 아래로 내리기.
function jsSelectMoveDown(){
		var opt = selPU.options.selectedIndex;  
     if(opt<selPU.options.length-1) {
     	var tmpText = selPU.options[opt].text;
     	var tmpValue = selPU.options[opt].value;
     	selPU.options[opt].text= selPU.options[opt+1].text;
     	selPU.options[opt].value= selPU.options[opt+1].value;
     	selPU.options[opt+1].text= tmpText;
     	selPU.options[opt+1].value= tmpValue;
     	 selPU.options.selectedIndex = opt+1; 
     } 
	} 
	
//품의서에 결재선 등록 시작	-------------------------------------------------------------
var ilastApprovalID = "<%=ilastApprovalID%>";//최종승인자 직급
function jsSetReport(){
	var strMsg0 = "";  
	var strMsg1 = ""; 
	var strMsg2 = "";
	var strMsg = ""
	var arrValue = "";
	var arrID = "";
	var arrJobsn = "";
	var arrText = "";
	var arrIDC = ""; 
	var arrTextC = "";
	var isLast  = 0;
	var iAuthPosition = 0;
	var strSMS = "&nbsp;"
	
	//결재선 
	strMsg = ""
	opener.document.frm.hidALI.value = ""
	opener.document.frm.hidALTxt.value = ""
	opener.document.frm.hidALJ.value = ""
	opener.document.frm.hidAI.value = ""
	opener.document.frm.hidAJ.value = ""
	opener.document.frm.hidATxt.value = ""
	opener.document.frm.hidAI_H.value = ""
	opener.document.frm.hidATxt_H.value = ""
	opener.document.frm.sRfN.value = ""
	opener.document.frm.hidRfI.value = ""
	
	if(document.all.selPU.length>0){	 
		for(i=0;i<document.all.selPU.length;i++){   
			arrValue = document.all.selPU[i].value.split("-");   
			strSMS = "&nbsp;"; 
				iAuthPosition = iAuthPosition + 1; //결재순서
				if (i ==0){strSMS =  "	<input type='checkbox' value='1' name='chkSms' checked> SMS전송"};
					strMsg1 = strMsg1 
								+ "<td valign='top'  height= '100%' width='180'>"
				 				+ "<div id='dAP"+iAuthPosition+"'>"
					 			+ "<table width='100%' height='100%'  cellpadding='5' cellspacing='0' class='a' border=0>" 
								+ "<tr><td align='Center' bgcolor='#E6E6E6' height='20'>"+iAuthPosition+"차 검토</td></tr>"
								+ "<tr><td align='Center'>승인대기</td></tr>"
								+ "<tr><td align='Center'>"+document.all.selPU[i].text+"</td></tr>"
								+ "<tr><td align='Center'>&nbsp;</td></tr>"
								+ "<tr><td align='Center'>"+strSMS+"</td></tr>"
								+ "</table>" 
								+ "</div>"
								+ "</td>" 
								
								if (iAuthPosition == 1){
									arrID = arrValue[0];
									arrJobsn = arrValue[1] ;
									arrText = document.all.selPU[i].text;
								}else{
									arrID = arrID + "," + arrValue[0];
									arrJobsn =arrJobsn  + "," + arrValue[1] ;
									arrText = arrText+","+document.all.selPU[i].text;
							 }
							
							opener.document.frm.hidAI.value = arrID
							opener.document.frm.hidAJ.value = arrJobsn
							opener.document.frm.hidATxt.value = arrText 
			}	 
		}else{
					strMsg1 = "<td valign='top'>"
							+ "<div id='dAP1'>"
							+ "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>" 
							+	"<tr><td align='Center' bgcolor='#E6E6E6' height='20'>&nbsp;</td></tr>"
							+	"<tr><td align='Center'>&nbsp;</td></tr>" 
							+	"</table>"
							+	"</div>"
							+	"</td>"
		} 	
	
	
	 
	//합의 
	if(document.all.selAU.length>0){
		strSMS = "&nbsp;"
		if (arrID ==""){strSMS =  "	<input type='checkbox' value='1' name='chkSms_H' checked> SMS전송"};
	  strMsg2 = "<td valign='top'  width='180'  height='100%'>"
	  			+ "<div id='dAP_H'>"
	  			+ "<table width='100%' height='100%' cellpadding='5' cellspacing='0' class='a' border=0>"
					+ "<tr><td align='Center' bgcolor='#E6E6E6' height='20'>합의</td></tr>"
					+ "<tr><td align='Center'>승인대기</td></tr>"
					+ "<tr><td align='Center'>"+document.all.selAU[0].text+"</td></tr>"
					+	"<tr><td align='Center'>&nbsp;</td></tr>" 
					+ "<tr><td align='Center'>"+strSMS+"</td></tr>"
					+ "</table> "
					+"</div>"
					+"</td>" 
					opener.document.frm.hidAI_H.value = (document.all.selAU[0].value.split("-"))[0];
					opener.document.frm.hidATxt_H.value = document.all.selAU[0].text;
	}else{
		 strMsg2 = "<td valign='top'  width='180'>"
					+ "<div id='dAP_H'>"
					+ "<table width='100%' cellpadding='5' cellspacing='0' class='a'>"
					+ "	<tr><td align='Center' bgcolor='#E6E6E6' height='20'>합의</td></tr>"
					+ "	<tr><td align='Center'>&nbsp;</td></tr>"
						+	"<tr><td align='Center'>&nbsp;</td></tr>" 
					+ "	<tr><td align='Center'></td></tr>"
					+ "	</table>"
					+ "</div>"
					+ "</td>"
	}
	
	
	//최종승인  
	if(document.all.selLPU.length>0){ 
		strSMS = "&nbsp;" 
		 	if (arrID=="" && document.all.selAU.length<=0){strSMS =  "	<input type='checkbox' value='1' name='chkSms' checked> SMS전송"};
		 	
		 		strMsg0 ="<td valign='top'  width='180' height= '100%'>"
		 					+ "<div id='dAP0'>"
		 					+ "<table width='100%' height='100%' cellpadding='5' cellspacing='0' class='a' border=0>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6' height='20'>최종승인</td></tr>"
							+ "<tr><td align='Center'>승인대기</td></tr>"
							+ "<tr><td align='Center'>"+document.all.selLPU[0].text+"</td></tr>"
							+ "<tr><td align='Center'>&nbsp;</td></tr>"
							+ "<tr><td align='Center'>"+strSMS+"</td></tr>"
							+ "</table> " 
							+ "</div>"
							+ "</td>"
						 
							opener.document.frm.hidALI.value 	 = (document.all.selLPU[0].value.split("-"))[0];
							opener.document.frm.hidALTxt.value = document.all.selLPU[0].text;
							opener.document.frm.hidALJ.value   = (document.all.selLPU[0].value.split("-"))[1];
				isLast = 1;
	}else{
		strMsg0 = "<td valign='top'  width='180'>"
							+ "<div id='dAP0'>"
							+ "<table width='100%' height='100%' cellpadding='5' cellspacing='0' class='a'>" 
							+	"<tr><td align='Center' bgcolor='#E6E6E6' height='20'>최종승인</td></tr>"
							+ "<tr><td align='Center'>승인대기</td></tr>"
							+	"<tr><td align='Center'>&nbsp;</td></tr>"
							+	"<tr><td align='Center'><%=sjob_name%></td></tr>"
							+	"</table>"
							+	"</div>"
							+	"</td>"
	} 

	//최종합의
	if(document.all.selLHU.length>0){ 
		strSMS = "&nbsp;" 
		 	if (arrID=="" && document.all.selAU.length<=0){strSMS =  "	<input type='checkbox' value='1' name='chkSms' checked> SMS전송"};
		 	
		 		strMsg0 ="<td valign='top'  width='180' height= '100%'>"
		 					+ "<div id='dAP0'>"
		 					+ "<table width='100%' height='100%' cellpadding='5' cellspacing='0' class='a' border=0>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6' height='20'>최종합의</td></tr>"
							+ "<tr><td align='Center'>승인대기</td></tr>"
							+ "<tr><td align='Center'>"+document.all.selLHU[0].text+"</td></tr>"
							+ "<tr><td align='Center'>&nbsp;</td></tr>"
							+ "<tr><td align='Center'>"+strSMS+"</td></tr>"
							+ "</table> " 
							+ "</div>"
							+ "</td>"
						 
							opener.document.frm.hidAHI.value	= (document.all.selLHU[0].value.split("-"))[0];
							opener.document.frm.hidAHTxt.value	= document.all.selLHU[0].text;
							opener.document.frm.hidAHJ.value	= (document.all.selLHU[0].value.split("-"))[1];
				isLast = 1;
	} 

 //결재선 + 합의 등록						
		strMsg= "<table width='520' align='left' cellpadding='0' cellspacing='1' class='a' border='0'>"
					+	"<tr> " 
					+ strMsg1
					+ strMsg2
					+ strMsg0 
					+ "</tr>" 
					+"</table>" 
		opener.document.all.dAP.innerHTML = strMsg;  
		opener.document.frm.blnL.value = isLast; //최종승인자 등록여부 확인 
		
 //참조등록
 	for(i=0;i<document.all.selCU.length;i++){
				if(i==0){
				arrTextC = document.all.selCU[i].text ;
				arrIDC =document.all.selCU[i].value.split("-")[0] ;
				}else{
				arrTextC = arrTextC +"," + document.all.selCU[i].text ;
				arrIDC =arrIDC +","+ document.all.selCU[i].value.split("-")[0] ;
				}
			} 
			if(arrTextC!=""){
			opener.document.frm.sRfN.value = arrTextC
			opener.document.frm.hidRfI.value =arrIDC
} 
			self.close();
}
//품의서에 결재선 등록 끝-------------------------------------------------------------

// 페이지 로드시 기존 선택 결재선 가져오기
$(window).load(function(){  
	//결재선  
	 var arrAI = opener.document.frm.hidAI.value.split(",");
	 var arrAJ = opener.document.frm.hidAJ.value.split(",");
	 var arrATxt = opener.document.frm.hidATxt.value.split(","); 
	 for(i=0;i<arrAI.length;i++){
	 	if(arrAI[i]!=""){
	 	$("#selPU").append("<option value='"+arrAI[i]+"-"+arrAJ[i]+"'>"+arrATxt[i]+"</option>");
	}
	}
	
	//최종결재
	if(opener.document.frm.hidALI.value!=""){
		$("#selLPU").append("<option value='"+opener.document.frm.hidALI.value+"-"+opener.document.frm.hidALJ.value+"'>"+opener.document.frm.hidALTxt.value+"</option>");  
	}

	//최종합의
	if(opener.document.frm.hidAHI.value!=""){
		$("#selLHU").append("<option value='"+opener.document.frm.hidAHI.value+"-"+opener.document.frm.hidAHJ.value+"'>"+opener.document.frm.hidAHTxt.value+"</option>");  
	}

	//합의 
	if(opener.document.frm.hidAI_H.value !=""){ 
		$("#selAU").append("<option value='"+opener.document.frm.hidAI_H.value+"'>"+opener.document.frm.hidATxt_H.value+"</option>");  
	}
	//참조
	 var arrRI = opener.document.frm.hidRfI.value.split(",") ;
	 var arrRTxt = opener.document.frm.sRfN.value.split(",") ;  
	 for(i=0;i<arrRI.length;i++){
	 	if(arrRI[i]!=""){
		 	$("#selCU").append("<option value='"+arrRI[i]+"'>"+arrRTxt[i]+"</option>");
		}
	}

	if($("#selLPU option").size()>0) {
		$("#tRowLHU").hide();
	}
	if($("#selLHU option").size()>0) {
		$("#tRowLPU").hide();
	}
});

//검색 
$(document).ready(function(){
	$("#btnSearch").click(function(){
		 document.frmS.hidUI.value ="";
		 document.frmS.hidUN.value =""; 
		 document.frmS.hidJS.value =""; 
		 document.frmS.hidJN.value ="";
		var username = escape($("#sUN").val()); 
		 var url="ajaxUserList.asp";
		 var params = "sUN="+username; 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){ 
		 		$("#divUL").html(args);
		 	},

		 	error:function(e){
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 });
	}); 
	
	//등록자 부서를 default로 뿌려주기
	 <%if user_cid1 <> "" and not isnull(user_cid1) then%>  
	 		jsOpenClose(1,<%=user_cid1%>); 
	 <%end if%>		
	 <%if user_cid2 <> "" and not isnull(user_cid2) then%> 
	   jsOpenClose(2,<%=user_cid2%>); 
	 <%end if%>  
	 <%if user_cid3 <> "" and not isnull(user_cid3) then%> 
	   jsOpenClose(3,<%=user_cid3%>); 
	 <%end if%>  
	 <%if user_cid4 <> "" and not isnull(user_cid4) then%> 
	   jsOpenClose(4,<%=user_cid4%>);   
	 <%end if%>  
});
</script> 
<style>
	FORM {display:inline;}
</style>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"> 
	<tr>
		<td>결재선 선택 <hr width="100%"></td>
	</tr>
	<tr>
		<td> 
			<table width="650" align="center" cellpadding="3" cellspacing="1" class="a" border="0"> 
				<tr>
					<td>
						<!-- 사원 리스트-------------------->
						<form name="frmS" id="frmS" method="post" onsubmit="return false;">
							<input type="hidden" name="hidUI" value="">
							<input type="hidden" name="hidUN" value=""> 
							<input type="hidden" name="hidJS" value=""> 
							<input type="hidden" name="hidJN" value=""> 
						<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
							<tr>
								<td>사원명: <input type="text" name="sUN" id="sUN" value="<%=sUserName%>"> <input type="button" class="button" value="검색" id="btnSearch"></td>
							</tr>
							<tr>
								<td><%dim cid1, cid2, cid3, cid4, oldcid1, oldcid2, oldcid3, oldcid4 %> 
									<div id="divUL" style="padding:5px;border:gray solid 1px;width:350px;height:350px;overflow-y:auto;"> 
									<%IF isArray(arrList) THEN
											For intLoop = 0 To UBound(arrList,2) 
												part_sn = arrList(0,intLoop)
												cid1 = arrList(10,intLoop)
												cid2 = arrList(11,intLoop)
												cid3 = arrList(12,intLoop)
												cid4 = arrList(13,intLoop)
												
												if isnull(cid2) then cid2 = 0 
												if isnull(cid3) then cid3 = 0 
												if isnull(cid4) then cid4 = 0 	
												'//부서가 틀려지면 라인이미지가 틀려진다
												if intLoop < UBound(arrList,2)  THEN 
													IF part_sn <> arrList(0,intLoop+1)  THEN
														imgJoin = "joinbottom.gif"
													ELSE	
														imgJoin = "join.gif"
													END IF
												else
													imgJoin = "joinbottom.gif"
												end if	 
											
											if cid1 <> oldcid1 then	
										%>		
											<% if intloop <> 0 then%> 
												<% if oldcid2 <> 0 then%>
														<% if oldcid3 <> 0 then%> 
															<% if oldcid4 <> 0 then%>
															</div>
															<%end if%>
														</div>
														<%end if%>
												</div>
												<%end if%>
											</div>
											<% end if%> 
											<div id="divP1<%=cid1%>" style="cursor:hand;" onClick="jsOpenClose('1','<%=cid1%>');"><img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid1%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid1%>"> <%=arrList(1,intLoop)%></div>  	 
											<div id="divB1<%=cid1%>" style="display:none;cursor:hand;">	 
										<%end if%>	 
								 	 	<% if cid2  <>  oldcid2 and cid2 <> 0 then	 
								 	 				 if oldcid2<> 0 then
										%>	 
												<% if oldcid3 <> 0 then%> 
													<% if oldcid4 <> 0 then%>
															</div>
															<%end if%>
														</div>
														<%end if%>
													</div>  
										<% 			end if%>		 
												<div id="divP2<%=cid2%>" style="cursor:hand;padding:0 0 0 1;" onClick="jsOpenClose('2','<%=cid2%>');">
													<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
													<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid2%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid2%>"> <%=arrList(1,intLoop)%></div>  	 
												<div id="divB2<%=cid2%>" style="display:none;cursor:hand;">		
										<%end if%>	
											<% if cid3  <>  oldcid3 and cid3 <> 0 then	 
								 	 				 if oldcid3<> 0 then
										%>	 
													<% if oldcid4 <> 0 then%>
															</div>
															<%end if%>
													</div>  
										<% 			end if%>		 
												<div id="divP3<%=cid3%>" style="cursor:hand;padding:0 0 0 1;" onClick="jsOpenClose('3','<%=cid3%>');">
													<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
													<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
													<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid3%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid3%>"> <%=arrList(1,intLoop)%></div>  	 
												<div id="divB3<%=cid3%>" style="display:none;cursor:hand;">		
										<%end if%>	
									<% if cid4  <>  oldcid4 and cid4 <> 0 then	 
								 	 				 if oldcid4<> 0 then
										%>	 
													</div>  
										<% 			end if%>		 
												<div id="divP4<%=cid4%>" style="cursor:hand;padding:0 0 0 1;" onClick="jsOpenClose('4','<%=cid4%>');">
													<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
													<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
													<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
													<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid4%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid4%>"> <%=arrList(1,intLoop)%></div>  	 
												<div id="divB4<%=cid4%>" style="display:none;cursor:hand;">		
										<%end if%>	 
											<div id="divU" style="padding:0 0 0 1;background:white;" onClick="jsSetId('<%=intLoop%>','<%=arrList(4,intLoop)%>','<%=arrList(5,intLoop)%>','<%=arrList(6,intLoop)%>','<%=arrList(8,intLoop)%>');">
													<%if not isnull(arrList(4,intLoop)) then %>		
												<%if cid4<> 0 then%>
												<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
												<%end if%>
												<%if cid3<> 0 then%>
												<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
												<%end if%>
												<%if cid2<> 0 then%>
												<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
												<%end if%>
												<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
												<img src="/images/dtree/<%=imgJoin%>" align="absmiddle"> 
												<%=arrList(5,intLoop)%>&nbsp;<%=arrList(8,intLoop)%> <font color="gray">[<%=arrList(4,intLoop)%>]</font> 
												<%end if%>
											</div>
											
									<%		
											oldcid1 = cid1
											oldcid2 = cid2
											oldcid3 = cid3
											oldcid4 = cid4
										Next
									END IF
								%>
									</div>
								</td>
							</tr>
						</table>
					</form>
					<!-- //사원 리스트-------------------->
					</td> 
					<td valign="top"> 
						<table width="200" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
							<!--- 결재선 선택 리스트-------------------->
							<tr>	
								<td>
									<input type="button" id="btnAdd" value="결재 ▶" style="color:blue;" class="button" onClick="jsProcId(1);" style="cursor:hand;">
								</td>
								<td>
										<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" border="0">
											<tr>
												<td colspan="2"> + 결재선 (결재순서 ↓)</td>
											</tr>
											<tr>
												<td colspan="2"> 
														<select  name="selPU" id="selPU" multiple size="6" style="width:200px" class="textarea">
				
														</select> 
												</td>
											</tr>
											<tr>
												<td><img src="/images/mbtn_up.gif" align="absmiddle" id="sel_up" onClick="jsSelectMoveUp();"> <img src="/images/mbtn_down.gif" align="absmiddle" onClick="jsSelectMoveDown();"style="cursor:hand;"></td>
												<td align="right"><img src="/images/mbtn_selectDel.gif" align="absmiddle" onClick="jsSelDel(1)" style="cursor:hand;"></td>
											</tr>
										</table>
								</td>
							</tr>
							<!---  //결재선 선택 리스트-------------------->
							<!--- 합의 리스트-------------------->
							<tr>
								<td>
									<input type="button" id="btnAdd" value="합의 ▶"   class="button" onClick="jsProcId(2);" style="cursor:hand;">
								</td>
								<td style="padding-top:10px;"> 
										<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" border="0">
										<tr>
											<td>+ 합의</td>
										</tr>	
										<tr>
											<td align="center">
												<select  name="selAU" id="selAU" multiple size="1" style="width:200px" class="textarea">

												</select>
											</td>
										</tr>
											<tr>
													<td align="right"><img src="/images/mbtn_selectDel.gif" align="absmiddle"  onClick="jsSelDel(2)" style="cursor:hand;"></td>
											</tr>
									</table>
								</td>
							</tr>
							<!--- //합의 리스트-------------------->
							<!--- 최종결재  리스트-------------------->
							<tr id="tRowLPU">
								<td>
									<input type="button" id="btnAddLPU" value="최종결재 ▶"   class="button" style="color:blue;" onClick="jsProcId(4);" style="cursor:hand;">
								</td>
								<td style="padding-top:10px;"> 
										<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" border="0">
										<tr>
											<td>+ 최종결재</td>
										</tr>
										<tr>
											<td align="center">
												<select  name="selLPU" id="selLPU" multiple size="1" style="width:200px" class="textarea">

												</select>
											</td>
										</tr>
											<tr>
													<td align="right"><img src="/images/mbtn_selectDel.gif" align="absmiddle"  onClick="jsSelDel(4)" style="cursor:hand;"></td>
											</tr>
									</table>
								</td>
							</tr>
							<tr id="tRowLHU">
								<td>
									<input type="button" id="btnAddLHU" value="최종합의 ▶"   class="button" style="color:blue;" onClick="jsProcId(5);" style="cursor:hand;">
								</td>
								<td style="padding-top:10px;"> 
										<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" border="0">
										<tr>
											<td>+ 최종합의</td>
										</tr>
											<td align="center">
												<select  name="selLHU" id="selLHU" multiple size="1" style="width:200px" class="textarea">

												</select>
											</td>
										</tr>
											<tr>
													<td align="right"><img src="/images/mbtn_selectDel.gif" align="absmiddle"  onClick="jsSelDel(5)" style="cursor:hand;"></td>
											</tr>
									</table>
								</td>
							</tr>
								<!--- //최종결재 리스트-------------------->
							<tr>
								<td colspan="2"><hr width="100%"></td>
							</tr>
								<!--- 참조 리스트-------------------->
							<tr>
								<td>
									<input type="button" id="btnAdd" value="참조 ▶"  class="button" onClick="jsProcId(3);">
								</td>
								<td> 
										<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" border="0" >
										<tr>
											<td>+ 참조</td>
										</tr>	
										<tr>
											<td align="center">
												<select  name="selCU" id="selCU" multiple size="4" style="width:200px" class="textarea">

												</select>
											</td>
										</tr>
											<tr>
													<td align="right"><img src="/images/mbtn_selectDel.gif" align="absmiddle" onClick="jsSelDel(3)" style="cursor:hand;"></td>
											</tr>
									</table>
								</td>
							</tr>	<!--- //참조 리스트-------------------->
						</table>
					</td>
				</tr>
					<tr>
	<td align="center" colspan="3"><hr width="100%"><input type="button" class="button" value="등록" onClick="jsSetReport();" style="cursor:hand;"></td>
</tr>
			</table>
		</td>
	</tr>

</table>
