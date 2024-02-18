<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 관리 데이터처리
' History : 2010.09.28 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
Dim sMode ,strSql ,iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,strParm,ssStatus,sOpenDate,isMValue
Dim sCode, eCode,iGroupCode, ssName, dSDay, dEDay, isRate, isMargin, isStatus,isUsing, addSql
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
  	if ssName <> "" then
		if checkNotValidHTML(ssName) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	IF eCode ="" THEN eCode = 0 
	IF iGroupCode ="" THEN iGroupCode = 0 
	IF isRate = "" then	isRate = 0
	IF isMValue = "" THEN isMValue =0
	if isStatus = "" then isStatus = 0
	
Select Case sMode
	
	'//신규등록
	Case "I"	
	IF isStatus = "7" THEN
		if sOpenDate = "" then
			 sOpenDate = "getdate()"
		else
			sOpenDate = " convert(nvarchar(10),'"&sOpenDate&"',21)"&"+' "&formatdatetime(sOpenDate,4)&"'"
		end if	 
	END IF
		
	IF sOpenDate = "" THEN sOpenDate = "null"	
		
		strSql = "INSERT INTO [db_academy].[dbo].[tbl_sale] ([sale_name], [sale_rate], [sale_margin], [evt_code], [evtgroup_code], [sale_startdate], [sale_enddate], [sale_status], [adminid], [opendate],[lastupdate],sale_marginvalue )"&_
				" Values ('"&ssName&"',"&isRate&","&isMargin&","&eCode&","&iGroupCode&",'"&dSDay&"','"&dEDay&"',"&isStatus&",'"&session("ssBctId")&"',"&sOpenDate&",getdate(),"&isMValue&")	"				
		
		'response.write strSql &"<br>"
		dbacademyget.execute strSql
	
	strSql = "Select IDENT_CURRENT('db_academy.dbo.tbl_sale') as salecode "
	rsACADEMYget.Open strSql,dbACADEMYget,1
	sCode = rsACADEMYget("salecode")
	
	IF Err.Number <> 0 THEN
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
       dbacademyget.close()	:	response.End	
	END IF	
	
	IF eCode = 0 THEN eCode = ""
	response.redirect("saleReg.asp?menupos="&menupos&"&sC="&sCode)
	dbacademyget.close()	:	response.End
	
	'//수정모드
	Case "U"
	Dim strAdd : strAdd = ""
	
	IF isStatus ="7" AND sOpenDate="" THEN
		strAdd = " , [opendate] = getdate()"	
	END IF

	'검색어 체크--------------------------------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'검색어	
	 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일	
	 iCurrpage 		= requestCheckVar(Request("iC"),10)			'현재 페이지 번호
	 ssStatus		= requestCheckVar(Request("sstatus"),10)	'검색 상태
 	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&salestatus="&ssStatus
 	'--------------------------------------------------------------
 	
	strSql ="UPDATE  [db_academy].[dbo].[tbl_sale]  SET sale_name='"&ssName&"', sale_rate="&isRate&", sale_margin= "&isMargin&",evt_code= "&eCode&_
			", evtgroup_code="&iGroupCode&",sale_startdate= '"&dSDay&"',sale_enddate='"&dEDay&"',sale_status="&isStatus&",sale_using='"&isUsing&"'"&_
			" , sale_marginvalue = "&isMValue&", adminid='"&session("ssBctId")&"' , lastupdate =getdate() "&strAdd&_
			" WHERE sale_code = "&sCode		
	
	'response.write strSql &"<Br>"		
	dbacademyget.execute strSql

	addSql = ""
IF isMargin = 1 THEN		'동일마진
	addSql = addSql&"(i.sellcash-(i.sellcash*"&isRate&"/100))- convert(int,(i.sellcash-(i.sellcash*"&isRate&"/100))*(100-convert(float,convert(int,i.orgsuplycash/i.orgprice*10000)/100))/100)"
ELSEIF 	isMargin = 2 THEN	'업체부담
	addSql = addSql&"(i.sellcash-(i.sellcash*"&isRate&"/100)) - (i.orgprice- i.orgsuplycash)"
ELSEIF 	isMargin = 3 THEN	'반반부담
	addSql = addSql&"i.orgsuplycash - Convert(int, (i.orgprice-(i.sellcash-(i.sellcash*"&isRate&"/100)))/2)"
ELSEIF 	isMargin = 4 THEN	'10x10부담
	addSql = addSql&"i.orgsuplycash"
ELSEIF 	isMargin = 5 THEN	'직접설정
	addSql = addSql&"(i.sellcash-(i.sellcash*"&isRate&"/100)) - convert(int, (i.sellcash-(i.sellcash*"&isRate&"/100))*convert(float,"&isMValue&")/100)"
END IF

	strSql = "update [db_academy].[dbo].[tbl_saleItem]"
	strSql = strSql&" set [saleprice]=i.sellcash-(i.sellcash*"&isRate&"/100)"
	strSql = strSql&" , [salesupplycash]="&addSql
	strSql = strSql&" FROM [db_academy].dbo.tbl_diy_item i left join [db_academy].[dbo].[tbl_saleitem] s on s.itemid=i.itemid"
	strSql = strSql&" WHERE s.sale_code="&sCode
	'response.write strSql &"<Br>"
	'Response.end
	dbacademyget.execute strSql	

	IF Err.Number <> 0 THEN
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
       dbget.close()	:	response.End	
	END IF	
	
	IF eCode = 0 THEN eCode = ""
	response.redirect("saleList.asp?menupos="&menupos&"&"&strParm)
	dbget.close()	:	response.End
	
	CASE Else
	Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요2")    	
       dbget.close()	:	response.End
End Select	

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->