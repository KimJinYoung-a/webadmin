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

<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp" -->

<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%


'// ������



response.write "������"
response.end




































if application("Svr_Info")="Dev" then
	lotteAPIURL = "http://openapi.lotte.com"
	lotteAuthNo = "afc92a6024a23c9ae7c6e8fa3647c9fc0de8384e2b7798af0961e8a127d30516efd5a556fd6008b89630b3cf2b40b09b7e4a7a5f1ebd67a6d29446a381ed803c"
end if

'' response.write lotteAuthNo
response.write ltiMallAuthNo
'' response.end

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, buf
dim i, j, k

response.write "aa" & ltiMallAuthNo
response.end

'response.write lotteAuthNo
'response.end

'' -- ��ȯ/��ǰ
'' http://openapi.lotte.com/openapi/searchReturnList.lotte?subscriptionId=[subscriptionId]&start_date=20130415&end_date=20130416&ord_dtl_stat_cd=20

'' -- ���
'' http://openapi.lotte.com/openapi/searchCnclList.lotte?subscriptionId=[subscriptionId]&start_date=20130416&end_date=20130416&selCval=17

'' -- �ű��ֹ�
'' http://openapi.lotte.com/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=[subscriptionId]&start_date=20130416&end_date=20130416&SelOption=01

'' -- ����Ȯ���ֹ�
'' http://openapi.lotte.com/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=[subscriptionId]&start_date=20130416&end_date=20130416&SelOption=02


'// ============================================================================
'// [divcd]
'// ============================================================================
'A008			�ֹ����
'
'A004			��ǰ����(��ü���)
'A010			ȸ����û(�ٹ����ٹ��)
'
'A001			������߼�
'A002			���񽺹߼�
'
'A000			�±�ȯ���
'A100			��ǰ���� �±�ȯ���
'
'A009			��Ÿ����
'A006			�������ǻ���
'A700			��ü��Ÿ����
'
'A003			ȯ��
'A005			�ܺθ�ȯ�ҿ�û
'A007			ī��,��ü,�޴�����ҿ�û
'
'A011			�±�ȯȸ��(�ٹ����ٹ��)
'A012			�±�ȯ��ǰ(��ü���)

'A111			��ǰ���� �±�ȯȸ��(�ٹ����ٹ��)
'A112			��ǰ���� �±�ȯ��ǰ(��ü���)
'// ============================================================================

dim mode
dim sellsite
dim reguserid
Dim AssignedRow
Dim ErrMsg

dim resultCount

dim divcd, yyyymmdd, idx

mode = requestCheckVar(html2db(request("mode")),32)
sellsite = requestCheckVar(html2db(request("sellsite")),32)
idx = requestCheckVar(html2db(request("idx")),32)


dim oCxSiteCSOrderXML
Set oCxSiteCSOrderXML = new CxSiteCSOrderXML

if (mode = "getxsitecslist") then
    IF (sellsite="lotteCom") then
    	ErrMsg = ""

    	'// ========================================================================
    	'// ���
    	oCxSiteCSOrderXML.FRectSellSite = sellsite
    	oCxSiteCSOrderXML.FRectDivCD = "A008"
    	'oCxSiteCSOrderXML.FRectStartYYYYMMDD = "2013-04-24"
    	oCxSiteCSOrderXML.FRectStartYYYYMMDD = Left(DateAdd("d", -1, now), 10)				'// 2013-01-01
    	oCxSiteCSOrderXML.FRectEndYYYYMMDD = Left(now, 10)

    	Call oCxSiteCSOrderXML.SavexSiteCSOrderListtoDB

    	Call oCxSiteCSOrderXML.ResetXML()

    	'// ========================================================================
    	'// ��ǰ
    	oCxSiteCSOrderXML.FRectSellSite = sellsite
    	oCxSiteCSOrderXML.FRectDivCD = "A004"
    	'oCxSiteCSOrderXML.FRectStartYYYYMMDD = "2013-04-24"
    	oCxSiteCSOrderXML.FRectStartYYYYMMDD = Left(DateAdd("d", -1, now), 10)				'// 2013-01-01
    	oCxSiteCSOrderXML.FRectEndYYYYMMDD = Left(now, 10)

    	Call oCxSiteCSOrderXML.SavexSiteCSOrderListtoDB

    	Call oCxSiteCSOrderXML.ResetXML()


    else
        rw "������ sellsite:"&sellsite
        dbget.Close : response.end
    end if
elseif (mode = "setfinish") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B007' "
	sqlStr = sqlStr + " where idx = " + CStr(idx) + " and currstate = 'B001' "
	''rw strSql
	rsget.Open sqlStr, dbget, 1

elseif (mode = "delfinish") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B001' "
	sqlStr = sqlStr + " where idx = " + CStr(idx) + " and currstate = 'B007' "
	''rw strSql
	rsget.Open sqlStr, dbget, 1

else

end if

%>

<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('����Ǿ����ϴ�.');</script>
<script>location.replace('<%= refer %>');</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
