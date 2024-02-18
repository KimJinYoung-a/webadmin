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
    CASE "getCust" '//erp �ŷ�ó ��� ����� // ��ü �����ͼ���.

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
    CASE "setCust" '//erp �ŷ�ó ��� ����� // ��ü �����ͼ���. (SCM=>ERP)

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

    CASE "setCUST_sERP" '//erp �ŷ�ó ��� ����� // ��ü �����ͼ���. (SCM=>ERP)

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
        rw "������ : [" & mode & "]"

	CASE ELSE
        rw "������ : [" & mode & "]"
END SELECT

rw "returnValue:"&returnValue
%>
<script type="text/javascript" src="http://webadmin.10x10.co.kr/admin/approval/eapp/eapp.js"></script>
<script language='javascript'>
function jsSetErpAnbun(){
    var frmact = document.frmAct;

    if (confirm('�Ⱥ� ? ')){
        frmact.mode.value="divmake";
        frmact.param1.value=frmA.yyyymm.value;
        frmact.submit();
    }

}



function jsSetErpCUST(){
    var frmact = document.frmAct;

    if (confirm('ERP������� - �ŷ�ó ? ')){
        frmact.mode.value="setCust";
        frmact.param1.value="";
        frmact.submit();
    }

}
function jsGetErpCUST(){
   // alert('������');
   // return;
    var frmact = document.frmAct;

    if (confirm('ERP��ϼ��� - �ŷ�ó ? ')){
        frmact.mode.value="getCust";
        frmact.param1.value="";
        frmact.submit();
    }

}


//----------------------------------

//�����׸� ��������.
function jsGetArapCD_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP ��ϼ��� - �����׸� ')){
        frmact.mode.value="getArapCD_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}

function jsGetBIZSECTION_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP ��ϼ��� - ����ι� ')){
        frmact.mode.value="getBIZSECTION_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}

function jsGetCommCD_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP ��ϼ��� - �����ڵ� ')){
        frmact.mode.value="getCommCD_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}

function jsGetACCCD_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP ��ϼ��� - �������� ')){
        frmact.mode.value="getACCCD_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}

function jsGetACCCDGRP_sERP(){
    var frmact = document.frmAct;

    if (confirm('sERP ��ϼ��� - �������� �׷� ')){
        frmact.mode.value="getACCCDGRP_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}


function jsSetErpCUST_sERP(){
//alert(1)
//return;
    var frmact = document.frmAct;

    if (confirm('sERP ������� - �ŷ�ó ')){
        frmact.mode.value="setCUST_sERP";
        frmact.param1.value="";
        frmact.submit();
    }

}


</script>
<!--
YYYYMM : <input type="text" name="yyyymm" value="<%=LEFT(dateadd("m",-1,now()),7)%>" size=7 maxlength=7>
<input type="button" type="button" value="����Ⱥ� ����" onClick="jsSetErpAnbun();">
-->


<br>

<br><br>
<input type="button" class="button" value="ERP������� - �ŷ�ó" onClick="jsSetErpCUST_sERP();">
<br><br>
<input type="button" class="button" value="ERP��ϼ��� - �ŷ�ó" onClick="jsGetErpCUST();">
<br><br>
<input type="button" class="button" value="�ŷ�óPOP" onClick="jsGetCust('');">

<hr>

sERP
<br><br>

    ..............................................
    ..............................................
    <input type="button" onClick="document.location.href='/admin/approval/comm/erpLink.asp?menupos=1635'" value="reload">
    <br><br>
    
    �����׸� ���� : db_partner.dbo.tbl_TMS_BA_ARAP_CD_sERP<br>
    
    <input type="button" class="button" value="sERP : �����׸� ����" onClick="jsGetArapCD_sERP();">
    <br><br>
    
    ����ι� ���� : db_partner.dbo.sp_TMS_get_BA_BIZSECTION_sERP<br>
    
    <input type="button" class="button" value="sERP : ����ι� ����" onClick="jsGetBIZSECTION_sERP();">
    <br><br>
    
    �����ڵ� ���� : db_partner.dbo.sp_TMS_get_BA_COM_CD_sERP<br>
    
    <input type="button" class="button" value="sERP : �����ڵ� ����" onClick="jsGetCommCD_sERP();">
    <br><br>
    
    �������� ���� : db_partner.dbo.sp_TMS_get_SL_ACC_CD_sERP<br>
    
    <input type="button" class="button" value="sERP : �������� ����" onClick="jsGetACCCD_sERP();">
    <br><br>
    
    �������� �׷� ���� : db_partner.dbo.sp_TMS_get_SL_ACC_CD_GRP_sERP<br>
    
    <input type="button" class="button" value="sERP : �������� �׷� ����" onClick="jsGetACCCDGRP_sERP();">
    <br><br>
    
    �ŷ�ó ���� : db_SCM_LINK.[dbo].sp_BA_CUST_Update_sERP<br>
    <input type="button" class="button" value="sERP : ERP������� - �ŷ�ó" onClick="jsSetErpCUST_sERP();">
    <br><br>
    
    �ŷ�ó ��� : <a href="http://scm.10x10.co.kr/admin/linkedERP/cust/popGetCust.asp" target="_blank">http://scm.10x10.co.kr/admin/linkedERP/cust/popGetCust.asp</a>
    
    <br><br>
    ..............................................
    ..............................................
    <br><strong>���ϰ� numeric�� �ƴ� @SLTRKEY / ���Ǽ� Ÿ�ºκ�..</strong>
    <br><strong>[toDo]</strong>-�����׸�:�������� ����. /*�ŷ�ó �ڵ����*/
    <br><br>
    sERP ��꼭 ���� �޴�<br>
    <input type="button" class="button" value="sERP : [�濵]�繫ȸ��>>�̼��������ڷ�" onClick="location.href='/admin/tax/?menupos=1395'">
    <br><br>
    
    sERP ���� ���� �޴�, <strong>[toDo]</strong>sERP ������� ���Ÿ޴�.<br>
    <input type="button" class="button" value="sERP : [�濵]�繫ȸ��>>������û�� ����Ʈ ����" onClick="location.href='/admin/approval/payreqList/?menupos=1383'">
    <br><br>
    
    
    sERP ����ī�� �������. / sERP ī����γ��� ����.<br>
    <input type="button" class="button" value="sERP : [�濵]������>>����ī��-���θ���Ʈ" onClick="location.href='/admin/expenses/card/preDailyOpExp.asp?menupos=1451'">
    <br><br>
    
    sERP ���ݿ��<br>
    <input type="button" class="button" value="sERP : [�濵]������>>���ݿ�����" onClick="location.href='/admin/expenses/opexp/?menupos=1340'">
    <br><br>
    
     <!-- [toDo] �ŷ�ó<br> ����Y�ΰŴ� �����ޱ�.(25300)/�뺯 10300 ���뿹�� (�ŷ�ó : ����)-->
    sERP ����ī��-��������<br>
    <input type="button" class="button" value="sERP : [�濵]������>>����ī��-��������" onClick="location.href='/admin/expenses/card/?menupos=1450'">
    <br><br>
    
    <strong>[toDo]</strong><br>
    sERP ��ǥ ��� ��� ������Ʈ db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE //���ݿ��, ����ī�� ��.
    <br><br>
    
    
    sERP �뷮��ü����.<br>
    <input type="button" class="button" value="sERP : [�濵]�繫ȸ��>>�Ա�Ȯ��File" onClick="location.href='/admin/upchejungsan/jungsanfinishNew.asp?menupos=1388'">
    
    <br><br>

<form name="frmAct" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="param1" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->