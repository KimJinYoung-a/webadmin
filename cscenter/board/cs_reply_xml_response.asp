<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : 1:1 상담
' History : 이상구 생성
'			2021.09.10 한용민 수정(이문재이사님요청 자사몰 필드추가, 보안강화)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_replycls.asp"-->
<?xml version="1.0"  encoding="euc-kr"?>
<response>
<%
dim sqlStr, i, mode, gubuncode, masteridx, defaultMasterStr, defaultDetailStr, defaultsitenameStr, oCReply, sitename
	mode = requestcheckvar(request("mode"),32)
	masteridx = requestcheckvar(getNumeric(request("masteridx")),10)
	gubuncode = requestcheckvar(request("gubuncode"),4)
	sitename = requestcheckvar(request("sitename"),32)

if (gubuncode = "") then
	gubuncode = "0001"
end if

if mode="replymaster" then
	defaultsitenameStr = "사이트구분"
	defaultMasterStr = "기본"
	defaultDetailStr = "상세"

	Set oCReply = new CReply
	oCReply.FPageSize = 30
	oCReply.FCurrPage = 1
	oCReply.FRectMasterUseYN = "Y"
	oCReply.FRectGubunCode = gubuncode
	oCReply.FRectsitename = sitename
	oCReply.GetReplyMasterList()

	response.write "<item><value1>XX</value1><value2><![CDATA[" + CStr(defaultMasterStr) + "]]></value2></item>" + VbCrlf
	for i = 0 to oCReply.FresultCount - 1
		response.write "<item><value1>" + CStr(oCReply.FItemList(i).Fidx) + "</value1><value2><![CDATA[" + CStr(oCReply.FItemList(i).Ftitle) + "]]></value2></item>" + VbCrlf
	next

	Set oCReply = Nothing

elseif mode="replydetail" then
	defaultsitenameStr = "사이트구분"
	defaultMasterStr = "기본"
	defaultDetailStr = "상세"

	Set oCReply = new CReply
	oCReply.FPageSize = 30
	oCReply.FCurrPage = 1
	oCReply.FRectMasterIDX = masterIdx
	oCReply.FRectMasterUseYN = "Y"
	oCReply.FRectDetailUseYN = "Y"
	oCReply.FRectGubunCode = gubunCode
	oCReply.FRectsitename = sitename
	oCReply.GetReplyDetailList()

	response.write "<item><value1>XX</value1><value2><![CDATA[" + CStr(defaultDetailStr) + "]]></value2><value3><![CDATA[XX]]></value3></item>" + VbCrlf
	for i = 0 to oCReply.FresultCount - 1
		response.write "<item><value1>" + CStr(oCReply.FItemList(i).Fidx) + "</value1><value2><![CDATA[" + CStr(oCReply.FItemList(i).Fsubtitle) + "]]></value2><value3><![CDATA[" + CStr(oCReply.FItemList(i).Fcontents) + "]]></value3></item>" + VbCrlf
	next

	Set oCReply = Nothing

elseif mode="replysitename" then
	defaultsitenameStr = "사이트구분"
	defaultMasterStr = "기본"
	defaultDetailStr = "상세"

	Set oCReply = new CReply
	oCReply.FPageSize = 500
	oCReply.FCurrPage = 1
	oCReply.FRectMasterUseYN = "Y"
	oCReply.FRectGubunCode = gubuncode
	oCReply.GetReplysitenameList()

	response.write "<item><value1>XX</value1><value2><![CDATA[" + CStr(defaultsitenameStr) + "]]></value2></item>" + VbCrlf
	for i = 0 to oCReply.FresultCount - 1
		response.write "<item><value1>" + CStr(oCReply.FItemList(i).fsitename) + "</value1><value2><![CDATA[" + CStr(replysitename(oCReply.FItemList(i).fsitename)) + "]]></value2></item>" + VbCrlf
	next
	Set oCReply = Nothing

end if
%>
</response>
<!-- #include virtual="/lib/db/dbclose.asp" -->
