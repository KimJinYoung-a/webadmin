<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Session.codepage="65001"
Response.codepage="65001"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
<%
dim i,k

dim refer
''// 유효 접근 주소 검사 //
refer = request.ServerVariables("HTTP_REFERER")

if InStr(refer,"webadmin.10x10.co.kr")<1 then
	Call Alert_Return("잘못된 접속입니다.")
	response.end
end if

dim sqlStr, DesignerID, delfilename, itemid
dim delmode, addimgnameidx, waititemid, makerid

delmode = Request.Form("delmode")
DesignerID = Request.Form("designerid")
itemid = Request.Form("itemid")
waititemid = Request.Form("waititemid")
delfilename = Request.Form("delfilename")
makerid = request.cookies("partner")("userid")

If waititemid <> "" Then
	If (WaitItemCheckMyItemYN(DesignerID,waititemid)<>true) Then
		Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
		Response.end
	End If
Else
	If (ItemCheckMyItemYN(DesignerID,itemid)<>true) Then
		Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
		Response.end
	End If
End If

If delmode="waitedit" Then
	'###########################################################################
	'이미지 데이터 넣기
	'###########################################################################
	sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_addimage "
	sqlStr = sqlStr & " where ADDIMAGE='" & delfilename & "'"
	sqlStr = sqlStr & " and itemid='" & Cstr(waititemid) & "'"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
ElseIf delmode="edit" Then
	'###########################################################################
	'이미지 데이터 넣기
	'###########################################################################
	sqlStr = "delete from db_academy.dbo.tbl_diy_item_addimage "
	sqlStr = sqlStr & " where ADDIMAGE='" & delfilename & "'"
	sqlStr = sqlStr & " and itemid='" & Cstr(itemid) & "'"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
End If

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->