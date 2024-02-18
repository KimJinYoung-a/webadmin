<%@ language=vbscript %>
<% option explicit %>
 
 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->

<%
	Dim ix,i, page, pgsize
	Dim TotalPage, TotalCount
	Dim prepage, nextpage
	Dim mode
	Dim nIndent
	Dim userid,name,email,title,contents
    Dim nboard
	Dim fixnotics, fixSdate, fixEdate, fixSH, fixSM, fixEH, fixEM
    Dim sql,mailcontent
	Dim nboardmail
	dim mailcheck
	dim dispCate
	dim fileName
	dim isPopup, popSdate, popEdate, popSH, popSM, popEH, popEM
if Request.Form("writemode") = "write" then


userid = session("ssBctId")
name = Request("name")
email = Request("email")
title = html2db(Request("title"))
contents = html2db(Request("contents"))
title = replace(title,"'" , "&#8217;")
contents = replace(contents,"'","&#8217;")
fixnotics = Request("fixnotics")
mailcheck = Request("mailcheck")
dispCate	= requestCheckVar(Request("disp"),10) 
 fileName 	= ReplaceRequestSpecialChar(Request("sFileP")) 
 
 
 fixSdate =  requestCheckVar(Request("sSD"),10) 
 fixEdate =  requestCheckVar(Request("sED"),10) 
 fixSH =  requestCheckVar(Request("sSH"),2) 
 fixSM =  requestCheckVar(Request("sSM"),5) 
 fixEH =  requestCheckVar(Request("sEH"),2) 
 fixEM =  requestCheckVar(Request("sEM"),5) 
 
 
 fixSdate = fixSdate&" "&format00(2,fixSH)&":"&fixSH 
 fixEdate = fixEdate&" "&format00(2,fixEH)&":"&fixEM 


 isPopup =  requestCheckVar(Request("isPop"),1) 
 

 popSdate =  requestCheckVar(Request("sPSD"),10) 
 popEdate =  requestCheckVar(Request("sPED"),10) 
 popSH =  requestCheckVar(Request("sPSH"),2) 
 popSM =  requestCheckVar(Request("sPSM"),5) 
 popEH =  requestCheckVar(Request("sPEH"),2) 
 popEM =  requestCheckVar(Request("sPEM"),5) 
 
 popSdate = popSdate&" "&format00(2,popSH)&":"&popSM 
 popEdate = popEdate&" "&format00(2,popEH)&":"&popEM 

 

set nboard = new CBoard
nboard.FRectID = userid
nboard.FRectName = name
nboard.FRectEmail = email
nboard.FRectTitle = title
nboard.FRectContents = contents
nboard.FRectDispCate = dispCate
nboard.FRectfileName	= fileName
nboard.FFixNotics = fixnotics
nboard.FFixSDate = fixSdate
nboard.FFixEDate = fixEdate
nboard.FIsPopup = isPopup
nboard.FpopSDate = popSdate
nboard.FpopEDate = popEdate
 

nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
nboard.design_notice_write

set nboard = nothing

if mailcheck = "Y" then
'	set nboardmail = new CBoard
'	nboardmail.FRectContents = nl2br(db2html(contents))
'	nboardmail.FRectEmail = email
'	nboardmail.FRectTitle = "[텐바이텐 안내메일]" & db2html(title)
'	nboardmail.design_notice_mail_send

'	set nboardmail = nothing
end if

%>

<script language="javascript">
alert('저장되었습니다.');
location.replace('designer_notice.asp');
</script>

<%
end if
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
 