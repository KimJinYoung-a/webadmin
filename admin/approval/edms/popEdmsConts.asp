<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 문서관리 문서내용
' History : 2011.02.24 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
Dim clsedms ,sMode
Dim iedmsidx ,icateidx1,icateidx2,iserialNum,sedmsname,sedmscode,iviewNo,sedmsFile,blnApproval,blnScmApproval
Dim slastApprovalid,sscmLink,sscmsubmitLink,dregdate,sadminid,blnUsing,blnPay,blnCfoAgree
Dim sReFileName,sFileName, iagreeyn, iagreeyntarget, iagreeyntargetname
	iedmsidx = requestCheckvar(Request("ieidx"),10)
 	icateidx1 = requestCheckvar(Request("selC1"),10)
 	icateidx2 = requestCheckvar(Request("selC2"),10)

 	if icateidx1="" then icateidx1 = 0
 	if icateidx2="" then icateidx2 = 0
 	sMode = "I"
Set clsedms = new Cedms
IF iedmsidx <> "" THEN
	sMode ="U"
	clsedms.Fedmsidx 	= iedmsidx
	clsedms.fnGetEdmsData

	iedmsidx         	= clsedms.Fedmsidx
	icateidx1        	= clsedms.Fcateidx1
	icateidx2        	= clsedms.Fcateidx2
	iserialNum    		= clsedms.FserialNum
	sedmsname  			= clsedms.Fedmsname
	sedmscode   		= clsedms.Fedmscode
	iviewNo          	= clsedms.FviewNo
	sedmsFile        	= clsedms.FedmsFile
	blnApproval      	= clsedms.FisApproval
	blnScmApproval 		= clsedms.FisScmApproval
	slastApprovalid  	= clsedms.FlastApprovalid
	sscmLink         	= clsedms.FscmLink
	sscmsubmitLink		= clsedms.FscmsubmitLink
	dregdate         	= clsedms.Fregdate
	sadminid         	= clsedms.Fadminid
	blnPay				= clsedms.FPayEApp
	blnUsing			= clsedms.FisUsing
	blnCfoAgree         = clsedms.FCfoAgree
	iagreeyn      	    = clsedms.FisAgreeNeed
	iagreeyntarget      = clsedms.FisAgreeNeedTarget
	iagreeyntargetname	= clsedms.FisAgreeNeedTargetName
	if sedmsFile <> "" Then
 	sReFileName = sedmscode&"_"&sedmsname&"."&split(sedmsFile,".")(ubound(split(sedmsFile,".")))
 	sFileName = split(sedmsFile,"/")(ubound(split(sedmsFile,"/")))
	end if
END IF

 	if blnPay = "" THEN blnPay = "False"
 %>
 <script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
 <script type="text/javascript" src="/js/ajax.js"></script>
<script language="javascript">
<!--
    // 카테고리 ajax =========================================================================================================
    initializeReturnFunction("processAjax()");
    initializeErrorFunction("onErrorAjax()");

    var _divName = "C2";

    function processAjax(){
        var reTxt = xmlHttp.responseText;
        eval("document.all.div"+_divName).innerHTML = reTxt;

        if (_divName=="SN"){
        jsSetSDC();
       	}
    }

    function onErrorAjax() {
            alert("ERROR : " + xmlHttp.status);
    }

    //선택한 카테고리에 대한 하위 카테고리 리스트 가져오기 Ajax
    function jsSetCategory(sMode){
        var ipcidx  = document.frmReg.selC1.value;
        var icidx   = $("#selC2").val();
        if(sMode=="C2"){
        document.frmReg.sSN.value="";
        document.frmReg.sDC.value="";
        }
        _divName = sMode;
        initializeURL('ajaxCategory.asp?sMode='+sMode+'&ipcidx='+ipcidx+'&icidx='+icidx);
    	startRequest();

    }



//파일명 보여주기
function jsSetFile(sfilepath, sfilename, sfilelocation){
 eval("document.all."+sfilelocation).style.display= "";
 eval("document.all."+sfilelocation).innerHTML = "파일명: "+sfilename + " <a href=javascript:jsDelFile('"+sfilelocation+"');>[x]</a>" ;
 document.frmReg.hidAF.value = sfilepath + sfilename;

}

//파일 지우기
function jsDelFile(sfilelocation){
eval("document.all."+sfilelocation).innerHTML = "";
eval("document.all."+sfilelocation).style.display= "none";
document.frmReg.hidAF.value = ""
}

function jsSetSDC(){
	var sC1 = document.frmReg.selC1.options[document.frmReg.selC1.selectedIndex].text.split("-");
	var sC2 =$("#selC2 option:selected").text().split("-") ;
	document.frmReg.sDC.value = sC1[0]+"-"+sC2[0]+"-"+$("#sSN").val();
}

//등록 필드체크
function jsSubmit(){
	var frm = document.frmReg;
	frm.hidC2.value = $("#selC2").val();
	frm.hidSN.value = $("#sSN").val();

	if(frm.selC1.value=="0"){
	alert("대 카테고리를 선택하세요");
	return false;
	}

	if(frm.hidC2.value=="0"){
	alert("중 카테고리를 선택하세요");
	return false;
	}

	if(frm.hidSN.value.length <3){
	alert("일련번호는 세자리 숫자만 가능합니다.");
	return false;
	}

	if(frm.sDN.value==""){
	alert("문서명을 입력해주세요");
	return false;
	}

	if(frm.isAgreeNeed[0].checked == true) {
		if(frm.sNm.value == ''){
			alert('합의자를 선택 하세요')
			return false;
		}
	} 

//	if(frm.selJN.value==""){
//	alert("최종결재자 직책을  입력해주세요");
//	return false;
//	}

/*
	if((!frm.rdoH[0].checked)&&(!frm.rdoH[1].checked)){
	alert("CFO 합의 구분을 선택해 주세요.");
	return false;
	}
*/

	return true;
}

function isAgreeNeed_tr() {
	var frm = document.frmReg;

	if(frm.isAgreeNeed[0].checked == true) {
		document.getElementById("isAgreeNeedtr").style.display = "block";
	} else {
		document.getElementById("isAgreeNeedtr").style.display = "none";
		eval("document.frmReg.sId").value = "";
		eval("document.frmReg.sNm").value = ""; 
	}
}

function jsGetID(iCid, sUserID){
	var openWorker = window.open('/admin/approval/edms/PopWorkerList_edm.asp?department_id='+iCid+'&sUserid='+sUserID,'openWorker','width=350,height=570,scrollbars=yes');
	openWorker.focus();
}

	function jsDelID(sType){ 
		eval("document.frmReg.sId").value = "";
		eval("document.frmReg.sNm").value = ""; 
	}

 //파일 다운로드
    function jsDownload(ieidx, sRFN, sFN){
    var winFD = window.open("<%=uploadImgUrl%>/linkweb/edms/procDownload.asp?ieidx="+ieidx+"&sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
    }
//-->
</script>

<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td><strong>문서관리 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
		<form name="frmReg" method="post" action="procedms.asp" onSubmit="return jsSubmit();">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="hidAF" value="<%=sedmsFile%>">
		<input type="hidden" name="ieidx" value="<%=iedmsidx%>">
		<input type="hidden" name="hidC2" value="<%=icateidx2%>">
		<input type="hidden" name="hidSN" value="<%=iserialNum%>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">대 카테고리</td>
			<td bgcolor="#FFFFFF">
				<div id="divC1">
					<select name="selC1" id="selC1" onChange="jsSetCategory('C2');">
					<option value="0">--선택--</option>
					<%clsedms.sbGetOptedmsCategory 1,0,icateidx1 %>
					</select>
				</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">중 카테고리</td>
			<td bgcolor="#FFFFFF">
			<div id="divC2">
					<select name="selC2" id="selC2" onChange="jsSetCategory('SN');">
					<option value="0">-- 선택 --</option>
				<% 	IF icateidx1 > 0 THEN	'대카테고리 선택 후 중카테고리 선택가능하게
						clsedms.sbGetOptedmsCategory 2,icateidx1,icateidx2
					END IF
				%>
					</select>
			</div>
			</td>
		</tr>

		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">중카테고리 일련번호</td>
			<td bgcolor="#FFFFFF"> <div id="divSN"><input type="text" name="sSN" id="sSN" size="3" maxlenght="3" value="<%=iserialNum%>" onkeyup="jsSetSDC();"></div> </td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">문서코드</td>
			<td bgcolor="#FFFFFF"> <input type="text" name="sDC" size="10" maxlenght="10"  value="<%=sedmscode%>" style="border:0; readonly"> </td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">문서명</td>
			<td bgcolor="#FFFFFF"> <input type="text" name="sDN" size="30" maxlenght="64"  value="<%=sedmsname%>"> </td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">표시순서</td>
			<td bgcolor="#FFFFFF"> <input type="text" name="iVN" size="5" maxlenght="10" style="text-align:right;" value="<%=iviewno%>"> </td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">양식파일</td>
			<td bgcolor="#FFFFFF"> <%=sReFileName%>  <br> (양식파일의 등록 및 수정은 문서 등록후  리스트에서 해주세요)
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">결재유무</td>
			<td bgcolor="#FFFFFF"> <input type="radio" name="rdoA" value="1" <%IF blnApproval THEN%>checked<%END IF%> >Y <input type="radio" name="rdoA" value="0" <%IF not blnApproval THEN%>checked<%END IF%> >N</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">전자결재여부</td>
			<td bgcolor="#FFFFFF"> <input type="radio" name="rdoEA" value="1" <%IF blnScmApproval THEN%>checked<%END IF%> >Y <input type="radio" name="rdoEA" value="0" <%IF not blnScmApproval THEN%>checked<%END IF%> >N</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">합의자필요여부</td>
			<td bgcolor="#FFFFFF">
				<label id="isAgreeNeed1"><input type="radio"  name="isAgreeNeed" value="Y" onClick="isAgreeNeed_tr();" <% If iagreeyn = "Y" Then %>checked<% end if %>>Y</label>&nbsp;
				<label id="isAgreeNeed2"><input type="radio"  name="isAgreeNeed" value="N" onClick="isAgreeNeed_tr();" <% If isnull(iagreeyn) or iagreeyn <> "Y" Then %>checked<% end if %>>N</label>&nbsp;
			</td>
		</tr>
		<% If iagreeyn = "Y" Then %>
			<tr id="isAgreeNeedtr">
		<% Else %>
			<tr id="isAgreeNeedtr" style='display:none'>
		<% End If %>
			<td bgcolor="<%= adminColor("tabletop") %>">합의자</td>
			<td bgcolor="#FFFFFF">
				<div style="padding:1px; <%= adminColor("tablebg") %>;">
					<input type="hidden" name="sId" value="<%=iagreeyntarget%>">
					<input type="name" name="sNm" value="<%=iagreeyntargetname%>"class="text_ro" readonly size="10">&nbsp;
					<input type="button" class="button" value="선택" onClick="jsGetID('25','<%=iagreeyntarget%>');"> 
					<input type="button" value="&times"  class="button" onClick="jsDelID();"  />
				</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">최종결재자</td>
			<td bgcolor="#FFFFFF"> <%=printJobOption("selJN", slastApprovalid)%> </td>
		</tr>
		<!--
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">CFO합의필요</td>
			<td bgcolor="#FFFFFF"><input type="radio" name="rdoH" value="1" <%IF blnCfoAgree THEN%>checked<%END IF%> >합의필요  <input type="radio" name="rdoH" value="0" <%IF not blnCfoAgree THEN%>checked<%END IF%> >필요없음</td>
		</tr>
		-->
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">SCM연결링크</td>
			<td bgcolor="#FFFFFF"> <input type="text" name="sSL" size="50" maxlenght="100" value="<%=sscmLink%>"> </td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">SCMsubmit링크</td>
			<td bgcolor="#FFFFFF"> <input type="text" name="sSSL" size="50" maxlenght="100" value="<%=sscmsubmitLink%>"> </td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">결제요청서 사용유무</td>
			<td bgcolor="#FFFFFF"><input type="radio" name="rdoP" value="1" <%IF blnPay THEN%>checked<%END IF%> >사용  <input type="radio" name="rdoP" value="0" <%IF not blnPay THEN%>checked<%END IF%> >사용안함</td>
		</tr>
		<%IF iedmsidx <> "" THEN%>
			<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">사용유무</td>
			<td bgcolor="#FFFFFF"><input type="radio" name="rdoU" value="1" <%IF blnUsing THEN%>checked<%END IF%> >사용  <input type="radio" name="rdoU" value="0" <%IF not blnUsing THEN%>checked<%END IF%> >사용안함</td>
		</tr>
		<%END IF%>
		</table>
	</td>
</tr>
 <tr>
 	<td align="center"><input type="submit" value="등록" class="button"></td>
 	</tr>
 </form>
</table>
<!-- 페이지 끝 -->
</body>
</html>
<%
Set clsedms = nothing
 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->




