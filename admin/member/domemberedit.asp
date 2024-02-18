<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
dim mode
dim userid,username,userpass
dim usermail,divcd,isusing
dim bigo,discountrate,commission

dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode = request.Form("mode")
userid = request.Form("userid")
username = request.Form("username")
userpass = request.Form("userpass")
usermail = request.Form("usermail")
divcd = request.Form("divcd")
isusing = request.Form("isusing")
bigo = request.Form("bigo")
discountrate = request.Form("discountrate")
commission  = request.Form("commission")

if bigo="" then bigo="0"
'response.write mode + "<br>"
'response.write userid + "<br>"
'response.write username + "<br>"
'response.write userpass + "<br>"
'response.write usermail + "<br>"
'response.write divcd + "<br>"
'response.write isusing + "<br>"'

dim onepartner
set onepartner = new CPartnerUser

'on Error resume Next

if instr(mode,"add")>0 then
	if onepartner.duplicateUserID(userid) then
		response.write "<script>alert('중복된 아이디 또는 사용중 아이디.');</script>"
		response.write "<script>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if
end if

if mode="employadd" then                    
	onepartner.addNewEmploy userid,userpass,username,usermail,divcd,bigo        ''사용안함
elseif mode="employedit" then
	onepartner.editEmploy userid,userpass,username,usermail,divcd,bigo,isusing  ''사용안함
elseif mode="designeradd" then
	onepartner.addNewEmploy userid,userpass,username,usermail,divcd,bigo        ''사용안함
elseif mode="designeredit" then
	onepartner.editEmploy userid,userpass,username,usermail,divcd,bigo,isusing  ''사용안함
elseif mode="partneradd" then
        onepartner.addNewPartner userid,userpass,username,usermail,divcd,discountrate,commission,bigo           ''임시사용
elseif mode="partneredit" then
        onepartner.editPartner userid,userpass,username,usermail,divcd,isusing,discountrate,commission,bigo     ''임시사용
end if

'if err then
'	err.Description
'end if

set onepartner = Nothing
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->