<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ���� ������ó��
' History : 2008.04.07 ������ ����
'			2022.07.06 �ѿ�� ����(isms�������ġ)
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
		Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
       dbget.close()	:	response.End	
	END IF	
	
		'strSql = "select SCOPE_IDENTITY() From [db_event].[dbo].[tbl_sale] "	'/������.��ü ���� ���� �ѷ���. '/2016.06.02 �ѿ��
		strSql = "select SCOPE_IDENTITY()"

		rsget.Open strSql, dbget, 0
		sCode = rsget(0)
		rsget.Close
		 
	IF eCode = 0 THEN eCode = ""
	Alert_move "����Ǿ����ϴ�.","saleReg.asp?menupos="&menupos&"&eC="&eCode&"&sC="&sCode   	   
dbget.close()	:	response.End
	Case "U"
		if ssName <> "" and not(isnull(ssName)) then
			ssName = ReplaceBracket(ssName)
		end If

	Dim strAdd : strAdd = ""
	
	IF isStatus ="7" AND sOpenDate="" THEN
		strAdd = " , [opendate] = getdate()"	
	END IF

'	'�˻��� üũ--------------------------------------------------------------
'	 iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
'	 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'�˻���	
'	 sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
'	 sSdate     	= requestCheckVar(Request("iSD"),10)		'������
'	 sEdate     	= requestCheckVar(Request("iED"),10)		'������	
'	 iCurrpage 		= requestCheckVar(Request("iC"),10)			'���� ������ ��ȣ
'	 ssStatus		= requestCheckVar(Request("sstatus"),10)	'�˻� ����
' 	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&salestatus="&ssStatus
' 	'--------------------------------------------------------------
' 	
		strSql ="UPDATE  [db_event].[dbo].[tbl_sale]  SET sale_name='"&ssName&"', sale_rate="&isRate&", sale_margin= "&isMargin&",evt_code= "&eCode 
		strSql = strSql&	", evtgroup_code="&iGroupCode&",sale_startdate= '"&dSDay&"',sale_enddate='"&dEDay&"',sale_status="&isStatus&",sale_using='"&isUsing&"'" 
		strSql = strSql&	" , sale_marginvalue = "&isMValue&", adminid='"&session("ssBctId")&"' , lastupdate =getdate() , sale_type="&sSaleType&strAdd 
		strSql = strSql&	" WHERE sale_code = "&sCode		
	dbget.execute strSql
	
	IF Err.Number <> 0 THEN
		Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
       dbget.close()	:	response.End	
	END IF	
	
	IF eCode = 0 THEN eCode = ""
	Alert_move "����Ǿ����ϴ�.","saleReg.asp?menupos="&menupos&"&eC="&eCode&"&sC="&sCode   	  
dbget.close()	:	response.End
	CASE Else
	Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���2")    	
       dbget.close()	:	response.End
End Select	

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

