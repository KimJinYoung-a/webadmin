<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim mode
dim idx, mastergubun, gubun, gubunname, contents, disporder, isusing
dim menupos


mode 		= RequestCheckvar(request("mode"),16)
idx 		= RequestCheckvar(request("idx"),10)
mastergubun = RequestCheckvar(request("mastergubun"),2)
gubun 		= RequestCheckvar(request("gubun"),2)
gubunname 	= html2db(RequestCheckvar(request("gubunname"),32))
contents 	= html2db(request("contents"))
disporder 	= RequestCheckvar(request("disporder"),10)
isusing 	= RequestCheckvar(request("isusing"),1)
menupos 	= RequestCheckvar(request("menupos"),10)

if contents <> "" then
	if checkNotValidHTML(contents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

dim sqlStr

if (mode = "addgubun") then

	'// =======================================================================
	gubun = "01"

	sqlStr = " select IsNull(max(gubun), '00') as gubun "
	sqlStr = sqlStr + " from [db_academy].[dbo].[tbl_ACA_cs_template] "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and mastergubun = '" + CStr(mastergubun) + "' "

	rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if  not rsACADEMYget.EOF  then
			gubun = Right("0" + CStr(rsACADEMYget("gubun") + 1), 2)
		end if
	rsACADEMYget.Close


	'// =======================================================================
	sqlStr = " insert into [db_academy].[dbo].[tbl_ACA_cs_template](mastergubun, gubun, gubunname, contents, disporder, isusing) "
	sqlStr = sqlStr + " values('" + Cstr(mastergubun) + "','" + Cstr(gubun) + "','" + Cstr(gubunname) + "','" + Cstr(contents) + "'," + Cstr(disporder) + ",'" + Cstr(isusing) + "') "
	'response.write sqlStr
	'response.end
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

elseif (mode = "editgubun") then

	sqlStr = "update [db_academy].[dbo].[tbl_ACA_cs_template] "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " gubunname = '" + Cstr(gubunname) + "' "
	sqlStr = sqlStr + " , contents = '" + Cstr(contents) + "' "
	sqlStr = sqlStr + " , disporder = " + Cstr(disporder) + " "
	sqlStr = sqlStr + " , isusing = '" + Cstr(isusing) + "' "
	sqlStr = sqlStr + " , lastupdate = getdate() "
	sqlStr = sqlStr + " where idx = '" + Cstr(idx) + "' "
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('저장 되었습니다.');
<% if (mastergubun = "20") then %>
	location.replace('mail_template_gubun.asp?menupos=<%= menupos %>');
<% else %>
	location.replace('sms_template_gubun.asp?menupos=<%= menupos %>');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
