<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� CS �ֹ�ó��
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteCSOrderCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<% ' #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" '�ּ�ó��..2020/08/25 ������ %>
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%

'// 2014-08-27, skyer9
Server.ScriptTimeout = 60

if application("Svr_Info")="Dev" then
	'�ּ�ó��..2020/08/25 ������
	' lotteAPIURL = "http://openapi.lotte.com"
	' lotteAuthNo = "afc92a6024a23c9ae7c6e8fa3647c9fc0de8384e2b7798af0961e8a127d30516efd5a556fd6008b89630b3cf2b40b09b7e4a7a5f1ebd67a6d29446a381ed803c"
end if

'' response.write lotteAuthNo
'' response.end

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, buf
dim i, j, k

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

dim divcd, yyyymmdd, idx, finUserid, cksel

mode = requestCheckVar(html2db(request("mode")),32)
sellsite = requestCheckVar(html2db(request("sellsite")),32)
idx = requestCheckVar(html2db(request("idx")),32)
cksel = request("cksel")
finUserid = session("ssBctID")
If finUserid = "" Then
	finUserid = "system"
End If

dim oCxSiteCSOrderXML
Set oCxSiteCSOrderXML = new CxSiteCSOrderXML

dim getDivCD
if (mode = "getxsitecslist") then
    IF (sellsite="lotteCom") then
    	ErrMsg = ""

		getDivCD = Trim(application("xSiteGetLotteComCS_DIVCD"))
		if (getDivCD = "") then
			getDivCD = "A008"
		elseif (getDivCD = "A004") then
			getDivCD = "A008"
		else
			getDivCD = "A004"
		end if

		'// ========================================================================
		'// ��� or ��ǰ
		oCxSiteCSOrderXML.FRectSellSite = sellsite
		oCxSiteCSOrderXML.FRectDivCD = getDivCD
		if Hour(Now()) < 6 then
			'// �����ð��� ��������
			oCxSiteCSOrderXML.FRectStartYYYYMMDD = Left(DateAdd("d", -1, now), 10)				'// 2013-01-01
			oCxSiteCSOrderXML.FRectEndYYYYMMDD = Left(DateAdd("d", -1, now), 10)
		else
			oCxSiteCSOrderXML.FRectStartYYYYMMDD = Left(now, 10)
			oCxSiteCSOrderXML.FRectEndYYYYMMDD = Left(now, 10)
		end if

		Call oCxSiteCSOrderXML.SavexSiteCSOrderListtoDB

		Call oCxSiteCSOrderXML.ResetXML()

		if (getDivCD <> Trim(application("xSiteGetLotteComCS_DIVCD"))) then
			application("xSiteGetLotteComCS_DIVCD") = getDivCD
		end if
    else
        rw "������ sellsite:"&sellsite
        dbget.Close : response.end
    end if
elseif (mode = "setfinish") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B007' "
	sqlStr = sqlStr + " ,finUserId = '"&finUserId&"' "			'2015-07-06 ������ �߰�
	sqlStr = sqlStr + " where idx = " + CStr(idx) + " and currstate <> 'B007' "
	''rw strSql
	rsget.Open sqlStr, dbget, 1

elseif (mode = "setfinisharr") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B007' "
	sqlStr = sqlStr + " ,finUserId = '"&finUserId&"' "			'2015-07-06 ������ �߰�
	sqlStr = sqlStr + " where idx in (" & CStr(cksel) & ") and currstate <> 'B007' "
	''rw strSql
	rsget.Open sqlStr, dbget, 1

elseif (mode = "setjupsu") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B002' "
	sqlStr = sqlStr + " ,finUserId = '"&finUserId&"' "			'2015-07-06 ������ �߰�
	sqlStr = sqlStr + " where idx = " + CStr(idx) + " and currstate <> 'B002' "
	''rw strSql
	rsget.Open sqlStr, dbget, 1

elseif (mode = "setjupsuarr") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B002' "
	sqlStr = sqlStr + " ,finUserId = '"&finUserId&"' "			'2015-07-06 ������ �߰�
	sqlStr = sqlStr + " where idx in (" & CStr(cksel) & ") and currstate <> 'B002' "
	''rw strSql
	rsget.Open sqlStr, dbget, 1

elseif (mode = "delfinish") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B001' "
	sqlStr = sqlStr + " where idx = " + CStr(idx) + " and currstate <> 'B001' "
	''rw strSql
	rsget.Open sqlStr, dbget, 1

elseif (mode = "setdelfinisharr") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B001' "
	sqlStr = sqlStr + " ,finUserId = '"&finUserId&"' "			'2015-07-06 ������ �߰�
	sqlStr = sqlStr + " where idx in (" & CStr(cksel) & ") and currstate <> 'B001' "
	''rw strSql
	rsget.Open sqlStr, dbget, 1

else

end if

%>

<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('����Ǿ����ϴ�.(<%= application("xSiteGetLotteComCS_DIVCD") %>)');</script>
<script>location.replace('<%= refer %>');</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
