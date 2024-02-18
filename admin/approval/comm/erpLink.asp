<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/db/dbiTMSHelper.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode, param1


mode = requestCheckVar(request("mode"),20)
param1 = requestCheckVar(request("param1"),20)

dim prcName, returnValue
dim objCmd
SELECT CASE mode
    CASE "getCust" '//erp 거래처 목록 재수신 // 전체 데이터수신.

        prcName = "db_partner.[dbo].sp_Ten_TMS_BA_CUST_getAllData"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing
    CASE "setCust" '//erp 거래처 목록 재수신 // 전체 데이터수신. (SCM=>ERP)

        prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_Update"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbiTms_dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing


	CASE "divmake"
        prcName = "db_SCM_LINK.[dbo].sp_SCM2ERP_payreqDIV_MAKE('"&param1&"')"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbiTms_dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing
 ''---------------------------------------------------------------------------------------------   
    CASE "getArapCD_sERP"
        prcName = "db_partner.[dbo].sp_TMS_get_BA_ARAP_CD_sERP"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing
    CASE "getBIZSECTION_sERP"
        prcName = "db_partner.[dbo].sp_TMS_get_BA_BIZSECTION_sERP"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing
   CASE "getCommCD_sERP"
        prcName = "db_partner.[dbo].sp_TMS_get_BA_COM_CD_sERP"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing 	
    CASE "getACCCD_sERP"
        prcName = "db_partner.[dbo].sp_TMS_get_SL_ACC_CD_sERP"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing 		
    CASE "getACCCDGRP_sERP"
        prcName = "db_partner.[dbo].sp_TMS_get_SL_ACC_CD_GRP_sERP"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing

    CASE "setCUST_sERP" '//erp 거래처 목록 재수신 // 전체 데이터수신. (SCM=>ERP)

        prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_Update_sERP"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
    		With objCmd
    			.ActiveConnection = dbiTms_dbget
    			.CommandType = adCmdText
    			.CommandText = "{?= call "&prcName&"}"
    			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    			.Execute, , adExecuteNoRecords
    			End With
    	returnValue = objCmd(0).Value
    	Set objCmd = nothing
    CASE ""
        rw "미지정 : [" & mode & "]"

	CASE ELSE
        rw "미지정 : [" & mode & "]"
END SELECT

rw "returnValue:"&returnValue
%>
<script type="text/javascript" src="http://webadmin.10x10.co.kr/admin/approval/eapp/eapp.js"></script>
<script language='javascript'>
function jsSetErpAnbun(){
    var frmact = document.frmAct;

    if (confirm('안분 ? ')){
        frmact.mode.value="divmake";
        frmact.param1.value=frmA.yyyymm.value;
        frmact.submit();
    }

}



function jsSetErpCUST(){
    var frmact = document.frmAct;

    if (confirm('ERP목록전송 - 거래처 ? ')){
        frmact.mode.value="setCust";
        frmact.param1.value="";
        frmact.submit();
    }

}
function jsGetErpCUST(){
   // alert('수정중');
   // return;
    var frmact = document.frmAct;

    if (confirm('ERP목록수신 - 거래처 ? ')){
        frmact.mode.value="getCust";
        frmact.param1.value="";
        frmact.submit();
    }

}


//----------------------------------

//수지항목 가져오기.
function jsGetArapCD_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP 목록수신 - 수지항목 ')){
        frmact.mode.value="getArapCD_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}

function jsGetBIZSECTION_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP 목록수신 - 사업부문 ')){
        frmact.mode.value="getBIZSECTION_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}

function jsGetCommCD_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP 목록수신 - 공통코드 ')){
        frmact.mode.value="getCommCD_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}

function jsGetACCCD_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP 목록수신 - 계정과목 ')){
        frmact.mode.value="getACCCD_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}

function jsGetACCCDGRP_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP 목록수신 - 계정과목 그룹 ')){
        frmact.mode.value="getACCCDGRP_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}


function jsSetErpCUST_sERP(){
//alert(1)
//return;
    var frmact = document.frmAct;

    if (confirm('sERP 목록전송 - 거래처 ')){
        frmact.mode.value="setCUST_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}


</script>
<!--
YYYYMM : <input type="text" name="yyyymm" value="<%=LEFT(dateadd("m",-1,now()),7)%>" size=7 maxlength=7>
<input type="button" type="button" value="공통안분 적용" onClick="jsSetErpAnbun();">
-->


<br>

<br><br>
<input type="button" class="button" value="ERP목록전송 - 거래처" onClick="jsSetErpCUST_sERP();">
<br><br>
<input type="button" class="button" value="ERP목록수신 - 거래처" onClick="jsGetErpCUST();">
<br><br>
<input type="button" class="button" value="거래처POP" onClick="jsGetCust('');">

<hr>

sERP
<br><br>

    ..............................................
    ..............................................
    <input type="button" onClick="document.location.href='/admin/approval/comm/erpLink.asp?menupos=1635'" value="reload">
    <br><br>
    
    수지항목 연동 : db_partner.dbo.tbl_TMS_BA_ARAP_CD_sERP<br>
    
    <input type="button" class="button" value="sERP : 수지항목 수신" onClick="jsGetArapCD_sERP();">
    <br><br>
    
    사업부문 연동 : db_partner.dbo.sp_TMS_get_BA_BIZSECTION_sERP<br>
    
    <input type="button" class="button" value="sERP : 사업부문 수신" onClick="jsGetBIZSECTION_sERP();">
    <br><br>
    
    공통코드 연동 : db_partner.dbo.sp_TMS_get_BA_COM_CD_sERP<br>
    
    <input type="button" class="button" value="sERP : 공통코드 수신" onClick="jsGetCommCD_sERP();">
    <br><br>
    
    계정과목 연동 : db_partner.dbo.sp_TMS_get_SL_ACC_CD_sERP<br>
    
    <input type="button" class="button" value="sERP : 계정과목 수신" onClick="jsGetACCCD_sERP();">
    <br><br>
    
    계정과목 그룹 연동 : db_partner.dbo.sp_TMS_get_SL_ACC_CD_GRP_sERP<br>
    
    <input type="button" class="button" value="sERP : 계정과목 그룹 수신" onClick="jsGetACCCDGRP_sERP();">
    <br><br>
    
    거래처 전송 : db_SCM_LINK.[dbo].sp_BA_CUST_Update_sERP<br>
    <input type="button" class="button" value="sERP : ERP목록전송 - 거래처" onClick="jsSetErpCUST_sERP();">
    <br><br>
    
    거래처 등록 : <a href="http://scm.10x10.co.kr/admin/linkedERP/cust/popGetCust.asp" target="_blank">http://scm.10x10.co.kr/admin/linkedERP/cust/popGetCust.asp</a>
    
    <br><br>
    ..............................................
    ..............................................
    <br><strong>리턴값 numeric이 아님 @SLTRKEY / 결의서 타는부분..</strong>
    <br><strong>[toDo]</strong>-수지항목:계정과목 검토. /*거래처 자동등록*/
    <br><br>
    sERP 계산서 전송 메뉴<br>
    <input type="button" class="button" value="sERP : [경영]재무회계>>이세로전자자료" onClick="location.href='/admin/tax/?menupos=1395'">
    <br><br>
    
    sERP 지출 전송 메뉴, <strong>[toDo]</strong>sERP 결제결과 수신메뉴.<br>
    <input type="button" class="button" value="sERP : [경영]재무회계>>결제요청서 리스트 전송" onClick="location.href='/admin/approval/payreqList/?menupos=1383'">
    <br><br>
    
    
    sERP 법인카드 개별등록. / sERP 카드승인내역 수신.<br>
    <input type="button" class="button" value="sERP : [경영]운영비관리>>법인카드-승인리스트" onClick="location.href='/admin/expenses/card/preDailyOpExp.asp?menupos=1451'">
    <br><br>
    
    sERP 현금운영비<br>
    <input type="button" class="button" value="sERP : [경영]운영비관리>>현금운영비관리" onClick="location.href='/admin/expenses/opexp/?menupos=1340'">
    <br><br>
    
     <!-- [toDo] 거래처<br> 공제Y인거는 미지급금.(25300)/대변 10300 보통예금 (거래처 : 통장)-->
    sERP 법인카드-월별내역<br>
    <input type="button" class="button" value="sERP : [경영]운영비관리>>법인카드-월별내역" onClick="location.href='/admin/expenses/card/?menupos=1450'">
    <br><br>
    
    <strong>[toDo]</strong><br>
    sERP 전표 등록 결과 업데이트 db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE //현금운영비, 법인카드 상세.
    <br><br>
    
    
    sERP 대량이체전송.<br>
    <input type="button" class="button" value="sERP : [경영]재무회계>>입금확정File" onClick="location.href='/admin/upchejungsan/jungsanfinishNew.asp?menupos=1388'">
    
    <br><br>

<form name="frmAct" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="param1" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->