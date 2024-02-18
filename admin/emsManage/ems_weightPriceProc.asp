<%@ Language=VBScript %>
<%
'==========================================================================
'	Description: EMS 서비스지역 처리, 서동석
'	History: 2009.04.07
'==========================================================================
	Option Explicit
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #Include Virtual="/lib/classes/order/clsEms_serviceArea.asp" -->
<%
Dim conListURL	: conListURL = "ems_weightPrice.asp"
Dim conSaveURL	: conSaveURL = "ems_weightPriceSave.asp"
Dim conProcURL	: conProcURL = "ems_weightPriceProc.asp"

Dim page		: page			    = requestCheckVar(request("page"),10)
Dim scCompanyCode	: scCompanyCode	= requestCheckVar(request("scCompanyCode"),3)
Dim scEmsAreaCode	: scEmsAreaCode	= requestCheckVar(request("scEmsAreaCode"),2)
Dim scWeightLimit	: scWeightLimit	= requestCheckVar(request("scWeightLimit"),10)
Dim menupos     : menupos           = requestCheckVar(request("menupos"),10)


if scCompanyCode = "" then scCompanyCode = req("CompanyCode","")

Dim qString
qString = "scEmsAreaCode=" & scEmsAreaCode & "&scCompanyCode=" & scCompanyCode & "&scWeightLimit=" & scWeightLimit & "&menupos=" & menupos
conListURL = conListURL & "?" & qString & "&page=" & page

Dim mode		: mode		= req("mode","INS")
Dim i, PKID


' 테이블클래스
Dim obj	: Set obj = new CEms

obj.GetWeightPriceData

obj.FoneItem.FCompanyCode				= req("CompanyCode","")
obj.FoneItem.FemsAreaCode				= req("emsAreaCode","")
obj.FoneItem.FweightLimit				= req("weightLimit","")
obj.FoneItem.FemsPrice				    = req("emsPrice","")

'rw mode
'rw "FcountryCode="&obj.FoneItem.FcountryCode
'rw "FcountryNameKr="&obj.FoneItem.FcountryNameKr

If mode = "USE" Then	' 삭제, 사용
'	PKID = Split(req("countryCode",""),",")
'	For i = 0 To UBound(PKID)
'		obj.FoneItem.FcountryCode		= PKID(i)
'		obj.ProcData mode
'	Next
Else					' 등록,수정
	obj.ProcWeightPrice mode
End If


response.redirect conListURL

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
