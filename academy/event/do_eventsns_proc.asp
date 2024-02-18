<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description :  핑거스 이벤트 SNS 내용 저장
' History : 2017-04-17 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, eCode, fbtitle, fbdesc, fbimage, twlink, twtag1, twtag2, katitle, kaimage, kalink, mode

idx		= RequestCheckVar(request("idx"),10)
eCode 	= RequestCheckVar(request("eCode"),10)
fbtitle	= RequestCheckVar(request("fbtitle"),100)
fbdesc	= RequestCheckVar(request("fbdesc"),400)
fbimage	= RequestCheckVar(request("fbimage"),100)
twlink	= requestCheckvar(request("twlink"),200)
twtag1	= RequestCheckVar(request("twtag1"),50)
twtag2	= requestCheckvar(request("twtag2"),50)
katitle	= RequestCheckVar(request("katitle"),400)
kaimage	= RequestCheckVar(request("kaimage"),100)
kalink	= RequestCheckVar(request("kalink"),200)

if eCode = "" or eCode = "0" or isnull(eCode) then
  	response.write "<script language='javascript'>alert('잘못된 접속 입니다.');</script>"
  	dbget.close(): response.End
end If

if idx = "" or idx = "0" or isnull(idx) then
	''신규 등록
	mode = "add"
else
	''수정
	mode = "edit"
end if

dim sqlStr

if (mode = "add") then
''신규 등록
    sqlStr = " insert into [db_academy].[dbo].[tbl_event_sns] " + VbCrlf
    sqlStr = sqlStr + " (evtcode, fbtitle, fbdesc, fbimage, twtitle, twlink, twtag1, twtag2, katitle, kaimage, kalink)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " " + eCode + "" + VbCrlf
    sqlStr = sqlStr + " ,'" + fbtitle + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + fbdesc + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + trim(fbimage) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + fbtitle + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + trim(twlink) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + twtag1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + twtag2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + katitle + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + trim(kaimage) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + trim(kalink) + "'" + VbCrlf
    sqlStr = sqlStr + " )"
	'response.write sqlStr
    dbACADEMYget.Execute sqlStr

elseif mode = "edit" Then
''수정
	if idx = "" or idx = "0" or isnull(idx) then
	  	response.write "<script language='javascript'>alert('잘못된 접속 입니다.');</script>"
	  	dbget.close(): response.End
	end If

   sqlStr = " update  [db_academy].[dbo].[tbl_event_sns] " + VbCrlf
   sqlStr = sqlStr + " set " + VbCrlf
   sqlStr = sqlStr + " fbtitle='" + fbtitle + "'" + VbCrlf
   sqlStr = sqlStr + " ,fbdesc='" + fbdesc + "'" + VbCrlf
   sqlStr = sqlStr + " ,fbimage='" + trim(fbimage) + "'" + VbCrlf
   sqlStr = sqlStr + " ,twlink='" + trim(twlink) + "'" + VbCrlf
   sqlStr = sqlStr + " ,twtag1='" + twtag1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,twtag2='" + twtag2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,katitle='" + katitle + "'" + VbCrlf
   sqlStr = sqlStr + " ,kaimage='" + trim(kaimage) + "'" + VbCrlf
   sqlStr = sqlStr + " ,kalink='" + trim(kalink) + "'" + VbCrlf
   sqlStr = sqlStr + " where idx=" + CStr(idx)
   dbACADEMYget.Execute sqlStr
end if

%>
<script language = "javascript">
	alert("저장되었습니다.");
	opener.location.reload();
	self.close();
</script>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->