<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<% response.Charset="euc-kr" %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_replycls.asp"-->
<?xml version="1.0"  encoding="euc-kr"?>
<response>
<%

dim sqlStr, i

dim mode
dim gubuncode, masteridx

mode = request("mode")
gubuncode = request("gubuncode")
masteridx = request("masteridx")
if (gubuncode = "") then
	gubuncode = "0001"
end if

dim defaultMasterStr, defaultDetailStr

dim oCReply

if mode="replymaster" then

	defaultMasterStr = "기본"
	defaultDetailStr = "상세"

	Set oCReply = new CReply
	oCReply.FPageSize = 30
	oCReply.FCurrPage = 1
	oCReply.FRectMasterUseYN = "Y"
	oCReply.FRectGubunCode = gubuncode

	oCReply.GetReplyMasterList()

	response.write "<item><value1>XX</value1><value2><![CDATA[" + CStr(defaultMasterStr) + "]]></value2></item>" + VbCrlf
	for i = 0 to oCReply.FresultCount - 1
		response.write "<item><value1>" + CStr(oCReply.FItemList(i).Fidx) + "</value1><value2><![CDATA[" + CStr(oCReply.FItemList(i).Ftitle) + "]]></value2></item>" + VbCrlf
	next

	Set oCReply = Nothing

elseif mode="replydetail" then

	defaultMasterStr = "기본"
	defaultDetailStr = "상세"

	Set oCReply = new CReply
	oCReply.FPageSize = 30
	oCReply.FCurrPage = 1
	oCReply.FRectMasterIDX = masterIdx
	oCReply.FRectMasterUseYN = "Y"
	oCReply.FRectDetailUseYN = "Y"
	oCReply.FRectGubunCode = gubunCode

	oCReply.GetReplyDetailList()

	response.write "<item><value1>XX</value1><value2><![CDATA[" + CStr(defaultDetailStr) + "]]></value2><value3><![CDATA[XX]]></value3></item>" + VbCrlf
	for i = 0 to oCReply.FresultCount - 1
		response.write "<item><value1>" + CStr(oCReply.FItemList(i).Fidx) + "</value1><value2><![CDATA[" + CStr(oCReply.FItemList(i).Fsubtitle) + "]]></value2><value3><![CDATA[" + CStr(oCReply.FItemList(i).Fcontents) + "]]></value3></item>" + VbCrlf
	next

	Set oCReply = Nothing

end if
%>
</response>
<!-- #include virtual="/lib/db/dbclose.asp" -->
