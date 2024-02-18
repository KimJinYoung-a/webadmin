<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_basicinfo.asp
' Discription : I형(통합형) 이벤트 작업자 정보 등록 창
' History : 2019.01.22 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim cEvtCont
Dim eCode, blnReqPublish, emdid, emdnm, edgid, edgstat1, edgstat2, edgid2, epsid
dim epsnm, edpnm, edpid, sWorkTag, efwd, fromlist


eCode = Request("eC")
fromlist = Request("fromlist")

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'이벤트 코드
    '이벤트 화면설정 내용 가져오기
    cEvtCont.fnGetEventDisplay
    blnReqPublish = cEvtCont.FisReqPublish
    emdid = cEvtCont.FEMdId
    emdnm = cEvtCont.FEMdName
    edgid = cEvtCont.FEDgId
    edgid2 = cEvtCont.FEDgId2
    edgstat1 = cEvtCont.FEDgStat1
    edgstat2 = cEvtCont.FEDgStat2
    epsid = cEvtCont.FEPsId
    epsnm = cEvtCont.FEPsName
    edpnm = cEvtCont.FEDpName
    edpid = cEvtCont.FEDpId
    sWorkTag = cEvtCont.FWorkTag
    efwd = db2html(cEvtCont.FEFwd)
    set cEvtCont = nothing
    if isnull(edgid) then edgid=""
    if isnull(edgid2) then edgid2=""
else
    edgid=""
    edgid2=""
end if 

if emdid = "" then 
    emdid = session("ssBctId")
    emdnm = session("ssBctCname")
end if

dim idepartmentid, sdepartmentname, clsMem
'부서명 가져오기
set clsMem = new CTenByTenMember
clsMem.Fuserid = emdid
clsMem.fnGetDepartmentInfo
idepartmentid = clsMem.Fdepartment_id
sdepartmentname = clsMem.FdepartmentNameFull 
set clsMem = Nothing
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
    frm.action="workerinfo_process.asp";
    frm.submit();
}

function jsGetID(sType, iCid, sUserID){
    var openWorker = window.open('/admin/eventmanage/event/V5/popup/popWorkerList.asp?sType='+sType+'&department_id='+iCid+'&sUserid='+sUserID,'openWorker','width=350,height=570,scrollbars=yes');
    openWorker.focus();
}

function jsDelID(sType){ 
    eval("document.frmEvt.s"+sType+"Id").value = "";
    eval("document.frmEvt.s"+sType+"Nm").value = ""; 
}

function jsChkMBReq(){ 
    if(document.frmEvt.chkMB.checked){  
            document.frmEvt.sWorkTag.value = "★★" + document.frmEvt.sWorkTag.value; 
    }else{
            document.frmEvt.sWorkTag.value =  document.frmEvt.sWorkTag.value.replace("★★", "");
    }
}

function jsPubHelp(){ 
    var winPop = window.open("pop_publishing.asp","popHelp","width=500,height=500,scrollbars=yes,resizable=yes");
    winPop.focus();
}

function jsAddByte(obj){ 
    var realText = obj.value; 
    var textBit = '';
    var textLen = 0;
    for (var i = 0 ; i < realText.length ; i++) {
        textBit = realText.charAt(i); 
        if(escape(textBit).length > 4) {
            textLen = textLen + 2;
        } else {
            textLen = textLen + 1;
        }

        if (textLen >= 32){
            realText = realText.substr(0,i);
            obj.value = realText;
            break;
        }
    }
	$("#WorkTag").html(textLen);
}
</script>

<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="imod" value="WU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="fromlist" value="<%=fromlist%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>담당자 정보</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <!-- for dev msg 기본 작성자를 본인으로 지정, th/td 클릭시 '찾기' 팝업 노출 -->
                    <th>기획자</th>
                    <td>
                        <input type="hidden" name="sMdId" value="<%=emdid%>">
                        <input type="text" class="formControl formControl150" name="sMdNm" placeholder="기획자" value="<%=eMDnm%>" readonly>
						<button class="btn4 btnBlue1 lMar05" onClick="jsGetID('Md','<%=idepartmentid%>','<%=emdid%>');return false;">찾기</button>
						<button class="btn4 btnGrey1 lMar05" onClick="jsDelID('Md');return false;">삭제</button>
                    </td>
                </tr>
                <tr>
                    <th>디자이너</th>
                    <td>
                        <%sbGetDesignerid "sDgId", edgid, ""%>
                        <input type="hidden" name="designerstatus" value="20">
                    </td>
                </tr>
                <tr>
                    <th>퍼블리셔</th>
                    <td>
                        <input type="hidden" name="sPsId" value="<%=epsid%>">
                        <input type="text" name="sPsNm" value="<%=epsnm%>" class="formControl formControl150" readonly size="10">
                        <button class="btn4 btnBlue1 lMar05" onClick="jsGetID('Ps','157','<%=epsid%>');return false;">찾기</button>
                        <button class="btn4 btnGrey1 lMar05" onClick="jsDelID('Ps');return false;">삭제</button>
                        <div class="formInline lMar05">
                            <label class="formCheckLabel">
                                <input type="checkbox" class="formCheckInput" name="chkReqP" value="1"<% if blnReqPublish then %> checked<%end if%>>
                                퍼블리싱 요청
                                <i class="inputHelper"></i>
                            </label>
                            <span class="mdi mdiBlue mdi-help-circle-outline cBl4" onClick="jsPubHelp();"></span>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>개발자</th>
                    <td>
                        <input type="hidden" name="sDpId" value="<%=edpid%>">
                        <input type="text" name="sDpNm" value="<%=edpnm%>" class="formControl formControl150" readonly size="10">
                        <button class="btn4 btnBlue1 lMar05" onClick="jsGetID('Dp','130','<%=epsid%>');return false;">찾기</button>
                        <button class="btn4 btnGrey1 lMar05" onClick="jsDelID('Dp');return false;">삭제</button>
                    </td>
                </tr>
                <tr>
                    <th>작업구분</th>
                    <td>
                        <input type="text" class="formControl formControl550" name="sWorkTag" placeholder=""  value="<%= sWorkTag %>" OnKeyUp="jsAddByte(this);">
                        <span class="lMar05 cGy1 fs12 vBtm"><span class="cPk1 vBtm" id="WorkTag">32</span><span class="cPk1 vBtm">byte</span>/32byte</span>
						<script type="text/javascript">
							jsAddByte(frmEvt.sWorkTag);
						</script>
                    </td>
                </tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">저장</button>
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->