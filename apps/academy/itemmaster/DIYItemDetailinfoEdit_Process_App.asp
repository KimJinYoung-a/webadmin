<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<%
dim i,k

dim refer
''// 유효 접근 주소 검사 //
refer = request.ServerVariables("HTTP_REFERER")


dim sellyn, isusing, itemid, sqlStr, makerid

itemid = RequestCheckVar(Request("itemid"),10)
sellyn = RequestCheckVar(Request("sellyn"),2)
isusing = RequestCheckVar(Request("isusing"),2)
makerid = RequestCheckVar(Request("makerid"),32)

If (ItemCheckMyItemYN(makerid,itemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If
'###########################################################################
'상품 데이터 수정
'###########################################################################
sqlStr = "update db_academy.dbo.tbl_diy_item" + vbCrlf
sqlStr = sqlStr & " set sellyn='" & sellyn & "'" + vbCrlf
sqlStr = sqlStr & " ,isusing='" & isusing & "'" + vbCrlf
sqlStr = sqlStr & " where itemid=" + CStr(itemid) + vbCrlf
dbACADEMYget.Execute sqlStr
'###########################################################################
%>
<script>
<!--
	parent.fnDetailStateInfoEnd();
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->