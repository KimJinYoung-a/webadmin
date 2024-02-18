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
Dim conListURL	: conListURL = "ems_serviceareaList.asp"
Dim conSaveURL	: conSaveURL = "ems_serviceareaSave.asp"
Dim conProcURL	: conProcURL = "ems_serviceareaProc.asp"

Dim page		: page			    = requestCheckVar(request("page"),10)
Dim scCountryCode	: scCountryCode	= requestCheckVar(request("scCountryCode"),2)
Dim IsUsing		: IsUsing		    = requestCheckVar(request("isusing"),1)
Dim menupos     : menupos           = requestCheckVar(request("menupos"),10)
Dim retUrl		: retUrl			= requestCheckVar(request("retUrl"),320)

'referer로 대체
'Dim qString : qString = "scCountryCode=" & scCountryCode & "&IsUsing=" & IsUsing & "&menupos=" & menupos
'conListURL = conListURL & "?" & qString & "&page=" & page

Dim mode		: mode		= req("mode","INS")
Dim i, PKID


' 테이블클래스
Dim obj	: Set obj = new CEms

obj.GetServiceAreaData

obj.FoneItem.FcompanyCode				= req("companyCode","")
obj.FoneItem.FcountryCode				= req("countryCode","")
obj.FoneItem.FcountryNameKr				= req("countryNameKr","")
obj.FoneItem.FcountryNameEn				= req("countryNameEn","")
obj.FoneItem.FemsAreaCode				= req("emsAreaCode","0")
obj.FoneItem.FemsMaxWeight				= req("emsMaxWeight",0)
obj.FoneItem.FreceiverPay				= req("receiverPay","N")
obj.FoneItem.Fisusing					= req("isusing","Y")
obj.FoneItem.FetcContents				= req("etcContents","")

'rw "mode=" & mode
'rw "FcountryCode="&obj.FoneItem.FcountryCode
'rw "FcountryNameKr="&obj.FoneItem.FcountryNameKr
'rw "Fisusing="&obj.FoneItem.Fisusing
'response.End

If mode = "DEL" Or mode = "USE" Then	' 삭제, 사용
'	PKID = Split(req("countryCode",""),",")
'	For i = 0 To UBound(PKID)
'		obj.FoneItem.FcountryCode		= PKID(i)
'		obj.ProcData mode
'	Next
Else					' 등록,수정
	obj.ProcServiceArea mode
End If


response.redirect retUrl

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
