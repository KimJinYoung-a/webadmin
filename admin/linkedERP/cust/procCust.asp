<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 거래처  등록
' History : 2011.12.08 정윤정  생성
'			2017.01.19 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sMode, objCmd, returnValue
Dim sCustcd,sCustNm,sBRNTYPE, sARYN, sAPYN, sCoYN, sBizNo, sCeoNm, sTaxType, sBSCD, sINTP, sTelNo, sFaxNo, sEmail, sDispSeq, sPostCd, sAddr
Dim sEmpno,sEmpNm, sPos,sDeptNM, sEmpTel, sEmpHP, sEmpEmail
Dim sModUser, sBigo
Dim arrBankNm, arrBankCd, arrSavMN,arrAcctNo, intLoop,sBankcd,sAcctNo,sSavMN,sARAPTYPE,sDEFACCTYN, sPSGB
dim srectEmpno,srectBankno,srectAcctno
sCustcd =	requestCheckvar(Request("hidCcd"),13)
 
	srectEmpno= requestCheckvar(Request("hidEno"),10)
	srectBankno= requestCheckvar(Request("hidBNo"),8)
	srectAcctno= requestCheckvar(Request("hidANo"),30)
sPSGB	= requestCheckvar(Request("rdoRT"),1)
IF sPSGB = "2" THEN
	sCustNm= requestCheckvar(Request("scnm7"),30)
	sBizNo= requestCheckvar(Request("sBno71"),6)&requestCheckvar(Request("sBno72"),7)
	sPos= requestCheckvar(Request("sEP7"),50)
	sDeptNM= requestCheckvar(Request("sDNm7"),50)
	sTelNo = requestCheckvar(Request("sTNo7"),30)
	sEmail= requestCheckvar(Request("sE7"),70)
	sEmpNm= requestCheckvar(Request("sem7"),50)
	sEmpTel= sTelNo
	sEmpEmail=sEmail
	sEmpHP= requestCheckvar(Request("sEHp7"),12)
	sDispSeq= requestCheckvar(Request("sDS7"),5)
	sBRNTYPE	= requestCheckvar(Request("selBRNT7"),1)
ELSE
	sCustNm= requestCheckvar(Request("scnm"),50)
	sBizNo= requestCheckvar(Request("sBno"),13)
	sCeoNm= requestCheckvar(Request("sceonm"),30)
	sTelNo= requestCheckvar(Request("sTNo"),12)
	sEmail= requestCheckvar(Request("sE"),70)
	sEmpNm= requestCheckvar(Request("sENm"),50)
	sEmpTel= requestCheckvar(Request("sETN"),30)
	sEmpHP= requestCheckvar(Request("sEHp"),30)
	sEmpEmail= requestCheckvar(Request("sEE"),70)
	sPos= requestCheckvar(Request("sEP"),50)
	sDeptNM= requestCheckvar(Request("sDNm"),50)
	sDispSeq= requestCheckvar(Request("sDS"),5)
	sBRNTYPE	= requestCheckvar(Request("selBRNT"),1)
END IF

sCustNm = trim(replace(sCustNm," ","")) '중간,앞뒤 공백제거
sBizNo  = trim(replace(replace(sBizNo,"-","")," ",""))

sARYN= requestCheckvar(Request("chkAR"),3)
sAPYN= requestCheckvar(Request("chkAP"),3)
sCoYN= requestCheckvar(Request("rdoCo"),3)

sTaxType= requestCheckvar(Request("selTType"),3)
sBSCD= requestCheckvar(Request("sBS"),35)
sINTP= requestCheckvar(Request("sIN"),35)
sFaxNo= requestCheckvar(Request("sFNo"),30)

sPostCd= requestCheckvar(Request("sPCd"),6)
sAddr= requestCheckvar(Request("sAddr"),200)
sEmpno= requestCheckvar(Request("hidENo"),10)

sBankcd= requestCheckvar(Request("selBC"),8)
sAcctNo= requestCheckvar(Request("sAN"),30)
sSavMN= requestCheckvar(Request("sSN"),50)
sModUser = session("ssBctId")
sARAPTYPE = 2 '1 입금처 계좌,2 지급처 계좌
sDEFACCTYN ="Y" '기본계좌 세팅여부
sMode		= requestCheckvar(Request("hidM"),2)
arrBankcd = split(sBankcd,",")

IF sARYN = "" THEN sARYN = "N"
IF sAPYN = "" THEN sAPYN = "N"
IF sDispSeq = "" THEN sDispSeq = 0

Dim prcName

SELECT CASE sMode
Case "I"
	'기본정보등록
	prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_ContsInsert_sERP"
	''if (session("ssBctID")="icommang") then prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_ContsInsert_sERP"
	    
	IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbiTms_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('"&sBRNTYPE&"', '"&sCoYN&"' ,'"&sARYN&"', '"&sAPYN&"'"&_
						+",'"&sCustNm&"','"&sBizNo&"','"&sCeoNm&"','"&sBSCD&"','"&sINTP&"','"&sPostCd&"','"&sAddr&"','"&sEmail&"','"&sTelNo&"'"&_
						+",'"&sFaxNo&"','"&sTaxType&"','"&sDispSeq&"','"&sModUser&"','"&sBIGO&"'"&_
						+", '"&sEmpNm&"' ,'"&sPos&"', '"&sDeptNM&"','"&sEmpTel&"','"&sEmpHP&"','"&sEmpEmail&"'"&_
						+", '"&sBankcd&"' ,'"&sAcctNo&"', '"&sARAPTYPE&"','"&sSavMN&"','"&sDEFACCTYN&"','"&sPSGB&"'"&_
						+")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

    IF 	returnValue =0 THEN  		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
    IF 	returnValue =-1 THEN  		Call Alert_return ("등록된 동일한 사업자 번호가 존재 합니다."&sBizNo&" 등록 불가.")
    IF 	returnValue <> 1 then response.end
    
    ''if (session("ssBctID")="icommang") then response.end
%>
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->
<!-- #include virtual="/lib/db/dbOpen.asp" -->
<%
    prcName = "db_partner.[dbo].sp_Ten_TMS_BA_CUST_getAllData"
    IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"

		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('','"&sBizNo&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
	Set objCmd = nothing
%>
<!-- #include virtual="/lib/db/dbClose.asp" -->
<script language="javascript">
	alert("등록되었습니다.");
	opener.location.reload();
	self.close();
	</script>
<%
	response.end
Case "U"
	'기본정보등록
	prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_ContsUpdate_sERP"
	''if (session("ssBctID")="icommang") then prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_ContsUpdate_sERP"
	IF application("Svr_Info")="Dev" THEN prcName = prcName & "_TEST"

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbiTms_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('"&sCustCD&"','"&sBRNTYPE&"', '"&sCoYN&"' ,'"&sARYN&"', '"&sAPYN&"'"&_
						+",'"&sCustNm&"','"&sBizNo&"','"&sCeoNm&"','"&sBSCD&"','"&sINTP&"','"&sPostCd&"','"&sAddr&"','"&sEmail&"','"&sTelNo&"'"&_
						+",'"&sFaxNo&"','"&sTaxType&"','"&sDispSeq&"','"&sModUser&"','"&sBIGO&"'"&_
						+",'"&sEmpNo&"', '"&sEmpNm&"' ,'"&sPos&"', '"&sDeptNM&"','"&sEmpTel&"','"&sEmpHP&"','"&sEmpEmail&"'"&_
						+", '"&sBankcd&"' ,'"&sAcctNo&"', '"&sARAPTYPE&"','"&sSavMN&"','"&sDEFACCTYN&"','"&sPSGB&"'"&_
						+")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF 	returnValue =0 THEN  		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	IF 	returnValue <> 1 then response.end
	
	''if (session("ssBctID")="icommang") then response.end
%>
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->
<!-- #include virtual="/lib/db/dbOpen.asp" -->
<%
    prcName = "db_partner.[dbo].sp_Ten_TMS_BA_CUST_getAllData"
    IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('"&sCustCD&"','"&sBizNo&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
	Set objCmd = nothing
%>
<!-- #include virtual="/lib/db/dbClose.asp" -->
<script language="javascript">
	alert("등록되었습니다.");
	opener.location.reload();
	self.close();
	</script>
<%
	response.end
CASE "DA" '//계좌번호 사용안함처리
 
	prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_AcctDelete_sERP" 
	IF application("Svr_Info")="Dev" THEN prcName = prcName & "_TEST"
 
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbiTms_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('"&sCustCD&"' , '"&srectBankno&"' ,'"&srectAcctno&"' )}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF 	returnValue =0 THEN  		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	IF 	returnValue <> 1 then response.end
	
	''if (session("ssBctID")="icommang") then response.end
%>
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->
<!-- #include virtual="/lib/db/dbOpen.asp" -->
<%
    prcName = "db_partner.[dbo].sp_Ten_TMS_BA_CUST_getAllData"
    IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('"&sCustCD&"','"&sBizNo&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
	Set objCmd = nothing
%>
<!-- #include virtual="/lib/db/dbClose.asp" -->
<script language="javascript">
	alert("계좌번호가 삭제되었습니다."); 
	parent.location.reload();
	</script>
<%
	response.end
CASE "IA"

	prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_ACCT_INSERT_sERP" 
	IF application("Svr_Info")="Dev" THEN prcName = prcName & "_TEST"
 response.write "{?= call "&prcName&"('"&sCustCD&"' , '"&sBankcd&"' ,'"&sAcctNo&"', '"&sARAPTYPE&"','"&sSavMN&"','"&sModUser&"','"&sDEFACCTYN&"' )}"
response.end
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbiTms_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('"&sCustCD&"' , '"&sBankcd&"' ,'"&sAcctNo&"', '"&sARAPTYPE&"','"&sSavMN&"','"&sModUser&"','"&sDEFACCTYN&"' )}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF 	returnValue =0 THEN  		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	IF 	returnValue <> 1 then response.end
	
	''if (session("ssBctID")="icommang") then response.end
%>
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->
<!-- #include virtual="/lib/db/dbOpen.asp" -->
<%
    prcName = "db_partner.[dbo].sp_Ten_TMS_BA_CUST_getAllData"
    IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('"&sCustCD&"','"&sBizNo&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
	Set objCmd = nothing
%>
<!-- #include virtual="/lib/db/dbClose.asp" -->
<script language="javascript">
	alert("계좌번호가 삭제되었습니다."); 
	parent.location.reload();
	</script>
<%
	response.end
  
		
CASE "R" '//erp 목록 재수신 // 전체 데이터수신.
	%>
	<!-- #include virtual="/lib/db/dbOpen.asp" -->
<%
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
	Set objCmd = nothing
%>
<!-- #include virtual="/lib/db/dbClose.asp" -->
<%	response.redirect "popGetCust.asp"
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
