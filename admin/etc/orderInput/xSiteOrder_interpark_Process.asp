<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteOrderXMLCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
If application("Svr_Info")="Dev" Then
	'lotteAPIURL = "http://openapi.lotte.com"
	'lotteAuthNo = "fe4d27496a64aac568c87bede1a195b0da6fe79e60274cf6b6ef19155d6d25698651567ede47da4a38bb2733d655b82e96d65308d8528b5b08ee6127ce87a2a1"
End If

Dim refer
refer = request.ServerVariables("HTTP_REFERER")
Dim sqlStr, buf
Dim i, j, k
Dim mode
Dim sellsite
Dim reguserid
Dim AssignedRow
Dim ErrMsg
Dim LastCheckDate, isSuccess
Dim maxCheckCount : maxCheckCount = 10
Dim resultCount
Dim divcd, yyyymmdd, idx
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)
yyyymmdd    = requestCheckVar(html2db(request("yyyymmdd")),8)
Dim oCxSiteOrderXML
Set oCxSiteOrderXML = new CxSiteOrderXML

maxCheckCount=1
If (mode = "getxsiteorderlist") Then
	oCxSiteOrderXML.FRectSellSite = sellsite
    IF (sellsite="interpark") Then
    	ErrMsg = ""
		For i = 0 to maxCheckCount - 1
			'// ================================================================
			Call oCxSiteOrderXML.GetCheckStatus(LastCheckDate, isSuccess)
			oCxSiteOrderXML.FRectStartYYYYMMDD = LastCheckDate
			oCxSiteOrderXML.FRectEndYYYYMMDD = LastCheckDate
			oCxSiteOrderXML.FRectAPIURL = "https://joinapi.interpark.com"
'rw LastCheckDate 2015-02-06/ 2015-02-07 오류
'response.end
if (LastCheckDate="2015-02-06") then LastCheckDate="2015-02-08"

            if (LastCheckDate<"2015-06-27") then
                oCxSiteOrderXML.FRectGubun = "new"
            else
                response.write "T:"&LastCheckDate
                ''response.end
            end if
   ''LastCheckDate="2014-01-01"

            isSuccess = "N"
			If (isSuccess = "Y") then
				oCxSiteOrderXML.FRectGubun = "new"
				If Not IsAutoScript then
					response.write "<br>" & LastCheckDate & " : 주문(신규) 가져오기 "
				End if
			Else
				oCxSiteOrderXML.FRectGubun = "all"
				If Not IsAutoScript Then
					response.write "<br>" & LastCheckDate & " : 주문(전체) 가져오기 "
				End if
			End If

'           oCxSiteOrderXML.FRectGubun = "js"
'			If Not IsAutoScript Then
'				rw LastCheckDate & " : 정산 가져오기 "
'			End if

			Call oCxSiteOrderXML.SetCheckStatusStarting(LastCheckDate)
			'// XML 가져오기
			Call oCxSiteOrderXML.SavexSiteOrderListtoDB
			Call oCxSiteOrderXML.ResetXML()

			response.write oCxSiteOrderXML.ErrMsg
			Call oCxSiteOrderXML.SetCheckStatusEnded()

			if Not IsAutoScript then
				rw "OK"
				rw "<form name=frmR method=post action=''><input type='hidden' name='mode' value='getxsiteorderlist'><input type='hidden' name='sellsite' value='interpark'><input type='button' name='reloadBtn' value='reload' onClick='document.frmR.submit();'></form>"
			end if

			if (CStr(LastCheckDate) >= CStr(Left(now, 10))) then
				exit for
			end if

			LastCheckDate = Left(DateAdd("d", 1, CDate(LastCheckDate)), 10)

			Call oCxSiteOrderXML.SetCheckDate(LastCheckDate)
		Next
    Else
        rw "미지정 sellsite:"&sellsite
        dbget.Close : response.end
    End If
Else

End If
%>
<% If  (IsAutoScript) Then %>
<% rw "OK" %>
<% Else %>
<script>
//if (confirm('저장되었습니다.')){document.frmR.submit();}
//setTimeout(function(){document.frmR.submit();},2000);
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
