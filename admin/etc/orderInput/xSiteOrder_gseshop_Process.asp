<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� XML �ֹ�ó��
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

<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->

<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp" -->

<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%

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

dim divcd, yyyymmdd

mode = requestCheckVar(html2db(request("mode")),32)
sellsite = requestCheckVar(html2db(request("sellsite")),32)


dim oCxSiteOrderXML
Set oCxSiteOrderXML = new CxSiteOrderXML

'// aaaaaaaaaaaaaaa
maxCheckCount = 1

if (mode = "getxsiteorderlist") then

	oCxSiteOrderXML.FRectSellSite = sellsite

    IF (sellsite="gseshop") then
    	ErrMsg = ""

		for i = 0 to maxCheckCount - 1
			'// ================================================================
			Call oCxSiteOrderXML.GetCheckStatus(LastCheckDate, isSuccess)

			'// aaaaaaaaaaaa	������(-)���� ����..YYYYMMDD
			LastCheckDate = "20140807"

			oCxSiteOrderXML.FRectStartYYYYMMDD = LastCheckDate
			oCxSiteOrderXML.FRectEndYYYYMMDD = LastCheckDate

			'// tnsType : �ֹ�����(�ֹ�/��ǰ : S, ��� : C)
			'// ���� : test1 � : ecb2b
			oCxSiteOrderXML.FRectAPIURL = "http://test1.gsshop.com/SupSendOrderInfo.gs?supCd=" + CStr(COurCompanyCode) + "&sdDt=" + CStr(LastCheckDate) + "&tnsType=S"

			if (isSuccess = "Y") then
				oCxSiteOrderXML.FRectGubun = "new" ''"new"

				if Not IsAutoScript then
					response.write "<br>" & LastCheckDate & " : �ֹ�(�ű�) ��û "
				end if
			else
				oCxSiteOrderXML.FRectGubun = "all"

				if Not IsAutoScript then
					response.write "<br>" & LastCheckDate & " : �ֹ�(��ü) ��û "
				end if
			end if

			Call oCxSiteOrderXML.SetCheckStatusStarting(LastCheckDate)

			'// �ű��ֹ� ���ۿ�û�� �Ѵ�.(XML ����X)
			Call oCxSiteOrderXML.RequestxSiteOrderListOnly

			response.write oCxSiteOrderXML.ErrMsg

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
        rw "������ sellsite:"&sellsite
        dbget.Close : response.end
    end if
else

end if

%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('����Ǿ����ϴ�.');</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
