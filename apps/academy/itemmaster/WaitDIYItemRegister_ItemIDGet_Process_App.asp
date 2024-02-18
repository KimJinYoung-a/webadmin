<%@ language=vbscript %>
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

dim sqlStr,DesignerID, waititemid
dim buycash, sellcash, imargin
dim target

DesignerID = Request.Form("designerid")
target = Request("target")
buycash=0
sellcash=0
'###########################################################################
'상품 데이터 입력
'###########################################################################
sqlStr = "insert into db_academy.dbo.tbl_diy_wait_item" + vbCrlf
sqlStr = sqlStr & " (itemdiv,makerid,itemname,regdate,buycash, sellcash, mileage, sellyn, deliverytype,limityn,currstate)" + vbCrlf
sqlStr = sqlStr & " values(" + vbCrlf
sqlStr = sqlStr & "'" & Cstr(Request.Form("itemdiv")) & "'" + vbCrlf
sqlStr = sqlStr & ",'" & DesignerID & "'" + vbCrlf
sqlStr = sqlStr & ",'tempitem'" + vbCrlf
sqlStr = sqlStr & ",getdate()" + vbCrlf
sqlStr = sqlStr & "," & buycash & "" + vbCrlf
sqlStr = sqlStr & "," & sellcash & "" + vbCrlf
sqlStr = sqlStr & "," & CLng(CLng(sellcash)*0.01) & "" + vbCrlf
sqlStr = sqlStr & ",'N'" + vbCrlf
sqlStr = sqlStr & ",'" & Request.Form("deliverytype") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & Request.Form("limityn") & "'" + vbCrlf
sqlStr = sqlStr & ",3)" + vbCrlf
'Response.write sqlStr
'Response.end
dbACADEMYget.Execute sqlStr
'###########################################################################
'상품 아이디 가져오기
'###########################################################################
sqlStr = "Select IDENT_CURRENT('db_academy.dbo.tbl_diy_wait_item') as maxitemid "
rsACADEMYget.Open sqlStr,dbACADEMYget,1
	waititemid = rsACADEMYget("maxitemid")
rsACADEMYget.close

%>
<script>
<!--
<% if target="1" then %>
	parent.fnImgTempSaveEnd1('<%=waititemid%>');
<% elseif target="2" then %>
	parent.fnImgTempSaveEnd2('<%=waititemid%>');
<% elseif target="3" then %>
	parent.fnImgTempSaveEnd3('<%=waititemid%>');
<% elseif target="4" then %>
	parent.fnImgTempSaveEnd4('<%=waititemid%>');
<% end if %>
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->