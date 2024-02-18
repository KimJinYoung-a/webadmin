<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 관리 데이터처리
' History : 2008.04.07 정윤정 생성
'			2022.07.06 한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim sMode
Dim sCode, eCode,iGroupCode, ssName, dSDay, dEDay, isRate, isMargin, isStatus,isUsing
Dim strSql
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,strParm,ssStatus,sOpenDate,isMValue
Dim sSaleType, dSTime, dETime, dSTSec, dETSec
sMode     = requestCheckVar(Request("sM"),1)	
sCode     = requestCheckVar(Request("sC"),10)	
eCode     = requestCheckVar(Request("eC"),10)	
ssName			= html2db(requestCheckVar(Request.Form("sSN"),64))
dSDay 			= requestCheckVar(Request.Form("sSD"),10)  
dEDay			= requestCheckVar(Request.Form("sED"),10)  
isRate			= requestCheckVar(Request.Form("iSR"),10)  
isMargin		= requestCheckVar(Request.Form("salemargin"),10)  
isStatus		= requestCheckVar(Request.Form("salestatus"),10)  
iGroupCode		= requestCheckVar(Request.Form("selG"),10)  	
isUsing			= requestCheckVar(Request.Form("sSU"),1)  
sOpenDate		= requestCheckVar(Request.Form("sOD"),30)  	
isMValue		= requestCheckVar(Request.Form("isMV"),10)  	
sSaleType       = requestCheckVar(Request("rdoT"),1)	
dSTime          = requestCheckVar(Request("sSTi"),2)	
dETime          = requestCheckVar(Request("sETi"),2)	
dSTSec			= requestCheckVar(Request("sSTSec"),5)	
dETSec			= requestCheckVar(Request("sETSec"),2)	

if sSaleType = 2 then
  dSDay =  dSDay &" "& Format00(2,dSTime)&":"&dSTSec
  dEDay =  dEDay &" "& Format00(2,dETime)&":"&dETSec&":00"
end if
 
IF eCode ="" THEN eCode = 0 
IF iGroupCode ="" THEN iGroupCode = 0 
IF isRate = "" then	isRate = 0
IF isMValue = "" THEN isMValue =0
if isStatus = "" then isStatus = 0
IF isUsing = "" then isUsing = 1
Select Case sMode
	Case "I"	
	IF isStatus = "7" THEN
		if sOpenDate = "" then
			 sOpenDate = "getdate()"
		else
			sOpenDate = " convert(nvarchar(10),'"&sOpenDate&"',21)"&"+' "&formatdatetime(sOpenDate,4)&"'"
		end if	 
	END IF
	IF sOpenDate = "" THEN sOpenDate = "null"	
		if ssName <> "" and not(isnull(ssName)) then
			ssName = ReplaceBracket(ssName)
		end If

		strSql = "INSERT INTO [db_event].[dbo].[tbl_sale] ([sale_name], [sale_rate], [sale_margin], [evt_code], [evtgroup_code], [sale_startdate], [sale_enddate], [sale_status], [adminid], [opendate],[lastupdate],sale_marginvalue ,sale_type)"&_
				" Values ('"&ssName&"',"&isRate&","&isMargin&","&eCode&","&iGroupCode&",'"&dSDay&"','"&dEDay&"',"&isStatus&",'"&session("ssBctId")&"',"&sOpenDate&",getdate(),"&isMValue&","&sSaleType&")	"				
		dbget.execute strSql
	
	IF Err.Number <> 0 THEN
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
       dbget.close()	:	response.End	
	END IF	
	
		'strSql = "select SCOPE_IDENTITY() From [db_event].[dbo].[tbl_sale] "	'/사용금지.전체 라인 몽땅 뿌려짐. '/2016.06.02 한용민
		strSql = "select SCOPE_IDENTITY()"

		rsget.Open strSql, dbget, 0
		sCode = rsget(0)
		rsget.Close
		 
	IF eCode = 0 THEN eCode = ""
	Alert_move "저장되었습니다.","saleReg.asp?menupos="&menupos&"&eC="&eCode&"&sC="&sCode   	   
dbget.close()	:	response.End
	Case "U"
		if ssName <> "" and not(isnull(ssName)) then
			ssName = ReplaceBracket(ssName)
		end If

	Dim strAdd : strAdd = ""
	
	IF isStatus ="7" AND sOpenDate="" THEN
		strAdd = " , [opendate] = getdate()"	
	END IF

'	'검색어 체크--------------------------------------------------------------
'	 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
'	 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'검색어	
'	 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
'	 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
'	 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일	
'	 iCurrpage 		= requestCheckVar(Request("iC"),10)			'현재 페이지 번호
'	 ssStatus		= requestCheckVar(Request("sstatus"),10)	'검색 상태
' 	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&salestatus="&ssStatus
' 	'--------------------------------------------------------------
' 	
		strSql ="UPDATE  [db_event].[dbo].[tbl_sale]  SET sale_name='"&ssName&"', sale_rate="&isRate&", sale_margin= "&isMargin&",evt_code= "&eCode 
		strSql = strSql&	", evtgroup_code="&iGroupCode&",sale_startdate= '"&dSDay&"',sale_enddate='"&dEDay&"',sale_status="&isStatus&",sale_using='"&isUsing&"'" 
		strSql = strSql&	" , sale_marginvalue = "&isMValue&", adminid='"&session("ssBctId")&"' , lastupdate =getdate() , sale_type="&sSaleType&strAdd 
		strSql = strSql&	" WHERE sale_code = "&sCode		
	dbget.execute strSql
	
	IF Err.Number <> 0 THEN
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
       dbget.close()	:	response.End	
	END IF	
	
	IF eCode = 0 THEN eCode = ""
	Alert_move "저장되었습니다.","saleReg.asp?menupos="&menupos&"&eC="&eCode&"&sC="&sCode   	  
dbget.close()	:	response.End
	CASE Else
	Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요2")    	
       dbget.close()	:	response.End
End Select	

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

