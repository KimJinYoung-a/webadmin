<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 XML 주문처리
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteOrderXMLCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->

<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp" -->

<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%

if application("Svr_Info")="Dev" then
	''lotteAPIURL = "http://openapi.lotte.com"
	''lotteAuthNo = "afc92a6024a23c9ae7c6e8fa3647c9fc0de8384e2b7798af0961e8a127d30516efd5a556fd6008b89630b3cf2b40b09b7e4a7a5f1ebd67a6d29446a381ed803c"
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, buf
dim i, j, k



dim mode
dim sellsite
dim reguserid
Dim AssignedRow
Dim ErrMsg
dim LastCheckDate, isSuccess
dim maxCheckCount : maxCheckCount = 10

dim resultCount

dim divcd, yyyymmdd, idx

mode = requestCheckVar(html2db(request("mode")),32)
sellsite = requestCheckVar(html2db(request("sellsite")),32)
idx = requestCheckVar(html2db(request("idx")),32)


dim oCxSiteOrderXML
Set oCxSiteOrderXML = new CxSiteOrderXML

if (mode = "getxsiteorderlist") then

	oCxSiteOrderXML.FRectSellSite = sellsite
''maxCheckCount=1
    IF (sellsite="lotteimall") then
    	ErrMsg = ""

		for i = 0 to maxCheckCount - 1
			'// ================================================================
			Call oCxSiteOrderXML.GetCheckStatus(LastCheckDate, isSuccess)
'LastCheckDate="2017-12-30"		'요청날의 -7일로 설정하기 by.진영
'isSuccess = "N"
			oCxSiteOrderXML.FRectStartYYYYMMDD = LastCheckDate
			oCxSiteOrderXML.FRectEndYYYYMMDD = LastCheckDate

			oCxSiteOrderXML.FRectAPIURL = "http://openapi.lotteimall.com"
			oCxSiteOrderXML.FRectAuthNo = ltiMallAuthNo

			'// aaaaaaaaaaaaaaaaa
			''oCxSiteOrderXML.FRectAuthNo = "192fe6a8de59b03e5370b7ba5ae348a2bcc182af4f7a9dc0650cc5ddfbd0438c30f001871cc315cdb6dfe61f78afd4690960488daa9b5f696dc61d33fb9aafac&"

            ''오늘 이전은 전체주문 가져오게끔
            ''response.write dateDiff("d",CDate(LastCheckDate),now())
'            if (dateDiff("d",CDate(LastCheckDate),now())<5) and (dateDiff("d",CDate(LastCheckDate),now())>0) then
'                isSuccess="N"
'            end if

			if (isSuccess = "Y") then
				oCxSiteOrderXML.FRectGubun = "new" ''"new"

				if Not IsAutoScript then
					response.write "<br>" & LastCheckDate & " : 주문(신규) 가져오기 "
				end if
			Elseif (isSuccess = "A") then		'2021-05-07 김진영 추가..누락건 발생이라그래서 seloption을 03으로..강제로 타임스탬프테이블의 isuccess를 A로 고치고 돌림
				oCxSiteOrderXML.FRectGubun = "fin" ''"new"

				if Not IsAutoScript then
					response.write "<br>" & LastCheckDate & " : 주문(약정) 가져오기 "
				end if
			else
				oCxSiteOrderXML.FRectGubun = "all"

				if Not IsAutoScript then
					response.write "<br>" & LastCheckDate & " : 주문(전체) 가져오기 "
				end if
			end if

''특정일 지정시 아래 주석처리
			Call oCxSiteOrderXML.SetCheckStatusStarting(LastCheckDate)

			'// XML 가져오기
			Call oCxSiteOrderXML.SavexSiteOrderListtoDB
			Call oCxSiteOrderXML.ResetXML()

			response.write oCxSiteOrderXML.ErrMsg

			'// aaaaaaaaaaaaaaaaaaaaaaa
			Call oCxSiteOrderXML.SetCheckStatusEnded()

			if Not IsAutoScript then
				response.write "OK"
			end if

			if (CStr(LastCheckDate) >= CStr(Left(now, 10))) then
				exit for
			end if

			LastCheckDate = Left(DateAdd("d", 1, CDate(LastCheckDate)), 10)

			Call oCxSiteOrderXML.SetCheckDate(LastCheckDate)
		next
    else
        rw "미지정 sellsite:"&sellsite
        dbget.Close : response.end
    end if
else

end if

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr


%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('저장되었습니다.');</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
