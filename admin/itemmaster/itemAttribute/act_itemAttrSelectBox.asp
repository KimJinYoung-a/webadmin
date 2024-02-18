<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%
'###############################################
' Discription : 상품속성 선택상자 Ajax
' History : 2013.08.27 허진원 : 신규 생성
'###############################################
Response.CharSet = "euc-kr"

dim dispCate, attribDiv
dispCate = request("dispCate")
attribDiv = request("attribDiv")

Response.Write getAttribDivSelectbox("attribDiv",attribDiv,dispCate,"")
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->