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

dim resultCount

dim divcd, yyyymmdd, idx

mode = requestCheckVar(html2db(request("mode")),32)
sellsite = requestCheckVar(html2db(request("sellsite")),32)
idx = requestCheckVar(html2db(request("idx")),32)

if (sellsite = "") then
	sellsite = "lotteimall"
end if

dim ord_no, ord_dtl_sn, sendQnt, sendDate, outmallGoodsID, hdc_cd, inv_no

ord_no = requestCheckVar(html2db(request("ord_no")),32)
ord_dtl_sn = requestCheckVar(html2db(request("ord_dtl_sn")),32)
sendQnt = requestCheckVar(html2db(request("sendQnt")),32)
sendDate = requestCheckVar(html2db(request("sendDate")),32)
outmallGoodsID = requestCheckVar(html2db(request("outmallGoodsID")),32)
hdc_cd = requestCheckVar(html2db(request("hdc_cd")),32)
inv_no = requestCheckVar(html2db(request("inv_no")),32)

dim oCxSiteCSOrderXML
Set oCxSiteCSOrderXML = new CxSiteCSOrderXML

if (mode = "getxsitecslist") then

    IF (sellsite="lotteimall") then
    	ErrMsg = ""

    	'// ========================================================================
    	'// ���
    	oCxSiteCSOrderXML.FRectSellSite = sellsite
    	oCxSiteCSOrderXML.FRectDivCD = "A008"
    	'' oCxSiteCSOrderXML.FRectStartYYYYMMDD = "2013-07-01"
		'' oCxSiteCSOrderXML.FRectEndYYYYMMDD = "2013-07-10"
    	oCxSiteCSOrderXML.FRectStartYYYYMMDD = Left(DateAdd("d", -10, now), 10)				'// 2013-01-01
    	oCxSiteCSOrderXML.FRectEndYYYYMMDD = Left(now, 10)

    	Call oCxSiteCSOrderXML.SavexSiteCSOrderListtoDB

    	Call oCxSiteCSOrderXML.ResetXML()

    	'// ========================================================================
    	'// ��ǰ
    	oCxSiteCSOrderXML.FRectSellSite = sellsite
    	oCxSiteCSOrderXML.FRectDivCD = "A004"
    	'' oCxSiteCSOrderXML.FRectStartYYYYMMDD = "2013-07-01"
		'' oCxSiteCSOrderXML.FRectEndYYYYMMDD = "2013-07-10"
    	oCxSiteCSOrderXML.FRectStartYYYYMMDD = Left(DateAdd("d", -10, now), 10)				'// 2013-01-01
    	oCxSiteCSOrderXML.FRectEndYYYYMMDD = Left(now, 10)

    	Call oCxSiteCSOrderXML.SavexSiteCSOrderListtoDB

    	Call oCxSiteCSOrderXML.ResetXML()


    else
        rw "������ sellsite:"&sellsite
        dbget.Close : response.end
    end if

elseif (mode = "sendsongjang") then
	'// ���Ϸ� ��������

    if (hdc_cd="99") and Len(replace(inv_no,"-",""))>15 then inv_no=Left(replace(inv_no,"-",""),15)
    if (inv_no="11�ù�ۿϷ�") then inv_no="11�ù��"
    if (inv_no="�ڵ����������ۿ���:)") then inv_no="��Ÿ"

	oCxSiteCSOrderXML.FRectSellSite = sellsite
	oCxSiteCSOrderXML.FRectDivCD = "sendsongjang"
	oCxSiteCSOrderXML.FRectStartYYYYMMDD = sendDate
	oCxSiteCSOrderXML.FRectEndYYYYMMDD = sendDate

	Call oCxSiteCSOrderXML.SendxSiteSongjangNo(ord_no, ord_dtl_sn, sendQnt, sendDate, outmallGoodsID, hdc_cd, inv_no)

	response.write oCxSiteCSOrderXML.ErrMsg

	Call oCxSiteCSOrderXML.ResetXML()

	dbget.close()
	response.end

elseif (mode = "updateSendState") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder"
	sqlStr = sqlStr & " set sendstate=" & requestCheckVar(request("updateSendState"), 10)
	sqlStr = sqlStr & " ,sendreqCnt=IsNULL(sendreqCnt,0)+1"
	sqlStr = sqlStr & " where outmallorderserial='"&ord_no&"'"
	sqlStr = sqlStr & " and orgdetailkey='"&ord_dtl_sn&"'"
	sqlStr = sqlStr & " and IsNULL(sendstate,0)=0"
	sqlStr = sqlStr & " and IsNULL(matchstate,'') <> 'D' and IsNULL(ordercsgbn, 0) = 0"
	dbget.Execute sqlStr

	response.write "OK"
	dbget.close()
	response.end

else

end if

%>

<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('����Ǿ����ϴ�.');</script>
<% if (mode = "getxsitecslist") then %>
<script>location.replace('<%= refer %>');</script>
<% end if %>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
