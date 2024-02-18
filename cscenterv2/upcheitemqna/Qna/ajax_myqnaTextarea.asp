<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'// UTF-8 변환
session.codePage = 65001
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
Dim groupcd, commcd, prfCont, cplCont, regid
Dim sqlStr, inputAnswerCont
groupcd = RequestCheckvar(request("groupcd"),4)
commcd = RequestCheckvar(request("commcd"),4)
regid = RequestCheckvar(request("regid"),32)

sqlStr = ""
sqlStr = sqlStr & " SELECT TOP 1 prfCont "
sqlStr = sqlStr & " FROM db_academy.dbo.tbl_preface "
sqlStr = sqlStr & " WHERE commCd = '"&groupcd&"' "
rsACADEMYget.Open sqlStr,dbACADEMYget,1
If Not(rsACADEMYget.EOF or rsACADEMYget.BOF) Then
	prfCont = rsACADEMYget("prfCont")
	prfCont = Replace(prfCont,"(아이디)", regid)
	prfCont = Replace(prfCont,"(이름)", session("ssBctCname"))
End If
rsACADEMYget.close

If commcd <> "" Then
	sqlStr = ""
	sqlStr = sqlStr & " SELECT TOP 1 cplCont "
	sqlStr = sqlStr & " FROM db_academy.dbo.tbl_compliment "
	sqlStr = sqlStr & " WHERE commCd = '"&commcd&"' "
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	If Not(rsACADEMYget.EOF or rsACADEMYget.BOF) Then
		cplCont = rsACADEMYget("cplCont")
		cplCont = Replace(cplCont,"(아이디)", regid)
		cplCont = Replace(cplCont,"(이름)", session("ssBctCname"))
	End If
	rsACADEMYget.close
End If
inputAnswerCont = prfCont & vbcrlf & vbcrlf & cplCont

Response.write "OK|" & inputAnswerCont
session.codePage = 949
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->