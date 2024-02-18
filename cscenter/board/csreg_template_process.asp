<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim mode
dim idx, mastergubun, gubun, gubunname, contents, disporder, isusing
dim menupos


mode 		= request("mode")
idx 		= request("idx")
mastergubun = request("mastergubun")
gubun 		= request("gubun")
gubunname 	= html2db(request("gubunname"))
contents 	= html2db(request("contents"))
disporder 	= request("disporder")
isusing 	= request("isusing")
menupos 	= request("menupos")

if (checkNotValidHTML(gubunname) = True) then
	Alert_return("구분명에는 HTML을 사용하실 수 없습니다.")
	dbget.Close
	response.end
end If

if (checkNotValidHTML(contents) = True) then
	Alert_return("내용에는 HTML을 사용하실 수 없습니다.")
	dbget.Close
	response.end
end if

if (checkNotValidHTML(disporder) = True) then
	Alert_return("표시순서에는 HTML을 사용하실 수 없습니다.")
	dbget.Close
	response.end
end if


dim sqlStr

if (mode = "addgubun") then

	'// =======================================================================
	gubun = "01"

	sqlStr = " select IsNull(max(gubun), '00') as gubun "
	sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_template "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and mastergubun = '" + CStr(mastergubun) + "' "

	rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
			gubun = Right("0" + CStr(rsget("gubun") + 1), 2)
		end if
	rsget.Close


	'// =======================================================================
	sqlStr = " insert into db_cs.dbo.tbl_cs_template(mastergubun, gubun, gubunname, contents, disporder, isusing) "
	sqlStr = sqlStr + " values('" + Cstr(mastergubun) + "','" + Cstr(gubun) + "','" + Cstr(gubunname) + "','" + Cstr(contents) + "'," + Cstr(disporder) + ",'" + Cstr(isusing) + "') "
	'response.write sqlStr
	'response.end
	rsget.Open sqlStr, dbget, 1

elseif (mode = "editgubun") then

	sqlStr = "update db_cs.dbo.tbl_cs_template "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " gubunname = '" + Cstr(gubunname) + "' "
	sqlStr = sqlStr + " , contents = '" + Cstr(contents) + "' "
	sqlStr = sqlStr + " , disporder = " + Cstr(disporder) + " "
	sqlStr = sqlStr + " , isusing = '" + Cstr(isusing) + "' "
	sqlStr = sqlStr + " , lastupdate = getdate() "
	sqlStr = sqlStr + " where idx = '" + Cstr(idx) + "' "
	rsget.Open sqlStr, dbget, 1

end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert("저장 되었습니다.");
location.replace("csreg_template_gubun.asp?menupos=<%= menupos %>&mastergubun=<%= mastergubun %>");
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
