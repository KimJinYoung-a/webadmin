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
	Dim part_sn, ireportidx
	Dim iMode, ilastApprovalID,iAuthPosition,sjob_name

	part_sn 		= requestCheckvar(Request("part_sn"),10)
	ireportidx 		= requestCheckvar(Request("iridx"),10)
	iMode			= requestCheckvar(Request("iM"),1)
	ilastApprovalID	= requestCheckvar(Request("iLAI"),10)
	iAuthPosition	= requestCheckvar(Request("iAP"),10)
  	sjob_name		= requestCheckvar(Request("sjn"),30)
	IF part_sn = "" THEN part_sn =0
	'// 직원정보 리스트
	dim oMember, arrList,intLoop
	Set oMember = new CTenByTenMember
	oMember.Fpart_sn 		= part_sn
	arrList = oMember.fnGetPartUserList
	set oMember = nothing

%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript">
<!--
$(document).ready(function(){
	$("#part_sn").change(function(){
		var part_sn = $("#part_sn").val();
		 var url="ajaxPartUserList.asp";
		 var params = "part_sn="+part_sn;

		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){
		 		$("#devU").html(args);
		 	},

		 	error:function(e){
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 });
	});
});

var mode = "<%=iMode%>";
var ilastApprovalID = "<%=ilastApprovalID%>";
var iAuthPosition ="<%=iAuthPosition%>";
$(function(){
	//추가버튼 클릭시 이벤트
	$("#btnAdd").click(function(){
	 var sValue;
	 var sText;
	if( $("#selUL option:selected").size() < 1){return}; //선택값 없는 경우 return 처리
	if ($("#selUL option:selected").size()> 1) {
	  alert("값은 한개씩 선택해주세요");
	  return;
	 	}

	sValue = $("#selUL option:selected").val();  //선택값
	sText  = $("#selUL option:selected").text(); //선택값의 텍스트


	for(j=0; j<$("#selUC option").size();j++){

		if($("#selUC option:eq("+j+")").val()==sValue){
			alert("이미 등록되어있습니다.");
			return;
		}
	}
 if (mode !=2){
	if ($("#selUC option").size() > 0){
		alert("결재자는 한명만 등록가능합니다.");
		return;
	}
}
	$("#selUC").append("<option value='"+sValue+"'>"+sText+"</option>"); //추가처리
	});

	//삭제
	$("#btnDel").click(function(){
		 $("#selUC option:selected").remove();
	});
});


//opener 등록폼에 추가
	function jsSetId(){
		var strMsg = "";
		var strUser ="";
		if(mode==1){
			strUser = "결재자";
		}else if(mode==3){
			strUser = "담당자";
		}else if(mode==4){
			strUser = "합의자";
		}else{
			strUser = "참조자";
		}

		if(document.frm.selUC.length==0){
 			alert(strUser+"을 추가해주세요");
		 return;
	 	}
		var arrValue =  document.frm.selUC[0].value.split("-");
		if (mode==1){ //지출품의서 결재선 등록
			if( arrValue[1]<=ilastApprovalID && arrValue[1] > 0){	//최종승인 직급선택시
				if(confirm("최종승인 직급입니다. 최종승인자로 지정하시겠습니까?\n\n합의가 필요한 경우 합의자를 추가 등록하시기 바랍니다.")){
					//검토자 div null 처리
					strMsg = "<table width='100%' cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>&nbsp;</td></tr>"
							+ "<tr><td align='Center'>&nbsp;</td></tr>"
							+ "</table>"
					opener.eval("document.all.dAP"+iAuthPosition).innerHTML = strMsg;

					//최종승인 div에 내용등록
					strMsg = "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>최종승인자</td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sASD' style='border:0;text-align:center;' value='승인대기'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sALN' id='sALN' value='"+document.frm.selUC[0].text+"' style='border:0;text-align:center;' readonly size='20'>"
							+ "<input type='hidden' name='hidAJ' id='hidAJ' value='"+ arrValue[1]+"'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sADD' value='' style='border:0;text-align:center;'></td></tr>"
							+ "<tr><td align='Center'><input type='button' class='button' value='결재자 등록' onClick='jsRegID(1,0);'><br><input type='checkbox' value='1' name='chkSms' checked> SMS전송</td></tr>"
							+ "</table> "
					opener.document.all.dAP0.innerHTML = strMsg;
					opener.document.frm.hidAI.value = arrValue[0] ;
					opener.document.frm.hidPS.value = document.frm.part_sn.value;
					opener.document.frm.blnL.value = 1;

                    //합의자 추가 (최종 직급 선택시만 가능)
					strMsg = "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>"
					       + "<tr><td align='Center' bgcolor='#E6E6E6'>합의</td></tr>"
					       + "<tr><td align='Center'>&nbsp;</td></tr>"
						   + "<tr><td align='Center'><input type='button' class='button' value='합의자 등록' onClick='jsRegID_H(4);'></td></tr>"
						   + "</table>"

					opener.document.all.dAP_H.innerHTML = strMsg;
                    opener.document.frm.hidAI_H.value = '';//초기화
			        opener.document.frm.hidPS_H.value = '';
					self.close();
				}
			}else{   //일반검토자 선택시
				if(opener.document.frm.blnL.value==1){ //기존에 최종승인자 선택 하고 수정할 경우 최종승인 div null 처리
			 	strMsg = "<table width='100%' cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>최종승인자</td></tr>"
							+ "<tr><td align='Center'>&nbsp;</td></tr>"
							+ "<tr><td align='Center'><%=sjob_name%></td></tr>"
							+ "</table>"
				opener.document.all.dAP0.innerHTML = strMsg;
				}
				//내용 넣기
				strMsg = "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>"+iAuthPosition+"차 검토</td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sASD' style='border:0;text-align:center;' value='승인대기'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sALN' id='sALN' value='"+document.frm.selUC[0].text+"' style='border:0;text-align:center;' readonly size='20'>"
							+ "<input type='hidden' name='hidAJ' id='hidAJ' value='"+ arrValue[1]+"'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sADD' value='' style='border:0;text-align:center;'></td></tr>"
							+ "<tr><td align='Center'><input type='button' class='button' value='결재자 등록' onClick='jsRegID(1,"+iAuthPosition+");'><br><input type='checkbox' value='1' name='chkSms' checked> SMS전송</td></tr>"
							+ "</table> "
				opener.eval("document.all.dAP"+iAuthPosition).innerHTML = strMsg;
				opener.document.frm.hidAI.value = arrValue[0] ;
				opener.document.frm.hidPS.value = document.frm.part_sn.value;
				opener.document.frm.blnL.value = 0;
				self.close();
			}
		}else if(mode==3){	//재무회계 담당자
			opener.document.frm.hidAI.value=arrValue[0];
			opener.document.frm.sALN.value=document.frm.selUC[0].text;
			opener.document.frm.hidAJ.value= arrValue[1];
			self.close();
        }else if(mode==4){	//합의자
            //합의자 추가 (최종 직급 선택시만 가능)
            strMsg = "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>합의</td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sASD_H' style='border:0;text-align:center;' value='승인대기'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sALN_H' id='sALN_H' value='"+document.frm.selUC[0].text+"' style='border:0;text-align:center;' readonly size='20'>"
							+ "<input type='hidden' name='hidAJ_H' id='hidAJ_H' value='"+ arrValue[1]+"'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sADD_H' value='' style='border:0;text-align:center;'></td></tr>"
							+ "<tr><td align='Center'><input type='button' class='button' value='합의자 등록' onClick='jsRegID_H(4);'><br><input type='checkbox' value='1' name='chkSms_H' checked> SMS전송</td></tr>"
							+ "</table> "
			opener.document.all.dAP_H.innerHTML = strMsg;
			opener.document.frm.hidAI_H.value = arrValue[0] ;
			opener.document.frm.hidPS_H.value = document.frm.part_sn.value;
			self.close();
		}else{ //참조자 등록
			for(i=0;i<frm.selUC.length;i++){
				if(i==0){
				opener.document.frm.sRfN.value = document.frm.selUC[i].text ;
				opener.document.frm.hidRfI.value =arrValue[0] ;
				}else{
				opener.document.frm.sRfN.value = opener.document.frm.sRfN.value +"," + document.frm.selUC[i].text ;
				opener.document.frm.hidRfI.value =opener.document.frm.hidRfI.value +","+ document.frm.selUC[i].value.split("-")[0] ;
				}
			}
			opener.document.frm.hidPS.value = document.frm.part_sn.value;
			self.close();
		}
	}

$(window).load(function(){ //페이지 로드시

	if ((mode!=2)&&(mode!=4)){
		if( ($("#hidAI",window.opener.document).val() != "") && ($("#hidAI",window.opener.document).val() !=  undefined )){ //기존 선택값 있을 경우
		var sText = $("#sALN",window.opener.document).val();
		var sValue = $("#hidAI",window.opener.document).val()+"-"+$("#hidAJ",window.opener.document).val();

		$("#selUC").append("<option value='"+sValue+"'>"+sText+"</option>"); //옵션값 추가
		}
	}else if (mode==4){ //합의자
	    if( ($("#hidAI_H",window.opener.document).val() != "") && ($("#hidAI_H",window.opener.document).val() !=  undefined )){ //기존 선택값 있을 경우
		var sText = $("#sALN_H",window.opener.document).val();
		var sValue = $("#hidAI_H",window.opener.document).val()+"-"+$("#hidAJ_H",window.opener.document).val();

		$("#selUC").append("<option value='"+sValue+"'>"+sText+"</option>"); //옵션값 추가
		}
	}else{ //2 : 참조
		if($("#hidRfI",window.opener.document).val() != "" && typeof($("#hidRfI",window.opener.document).val()) !=  undefined){ //기존 선택값 있을 경우
    		var sText = $("#sRfN",window.opener.document).val();
    		var sValue = $("#hidRfI",window.opener.document).val();
            if (sValue!=undefined){
        		var arrN = sText.split(",");
        		var arrI = sValue.split(",");
        		for(i=0;i<arrI.length;i++){
        		$("#selUC").append("<option value='"+arrI[i]+"'>"+arrN[i]+"</option>"); //옵션값 추가
        		}
    		}
		}
	}
});


//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
<form name="frm" method="post">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
		<tr>
			<td  align="center"> <%=printPartOptionAddEtc("part_sn", part_sn, "id=part_sn")%></td>
		</tr>
		<tr>
			<td align="center">
			<div id="devU">
				<select  name="selUL" id="selUL" multiple size="20" style="width:200px">
				<%IF isArray(arrList) THEN
					For intLoop = 0 To UBOund(arrList,2)
				%>
					<option value="<%=arrList(2,intLoop)&"-"&arrList(4,intLoop)%>"><%=arrList(1,intLoop)%>&nbsp;<%=arrList(7,intLoop)%> <%=arrList(2,intLoop)%>   </option>
				<%	Next
				END IF%>
				</select>
			</div>
			</td>
		</tr>
		</table>
	</td>
	<td>
		<input type="button" id="btnAdd" value="추가▶" class="button"> <br><br>
		<input type="button" id="btnDel" value="삭제◀" class="button">
	</td>
	<td  valign="bottom">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td>
				<select  name="selUC" id="selUC" multiple size="20" style="width:200px">

				</select>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" colspan="3"><input type="button" class="button" value="등록" onClick="jsSetId();"></td>
</tr>

</form>
</table>
