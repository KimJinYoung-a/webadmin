<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sqlStr, mduserid, catecode, mode, menupos
mduserid	= requestCheckvar(request("mduserid"),34)
catecode	= requestCheckvar(request("catecode"),6)
mode		= requestCheckvar(request("mode"),1)
menupos		= requestCheckvar(request("menupos"),5)

If mode = "I" Then
	If Len(catecode) = 3 Then
		sqlStr = "delete from db_partner.dbo.tbl_mdmenu_category Where left(catecode,3) = '"&catecode&"';" & vbCrLf
		sqlStr = sqlStr & " INSERT INTO db_partner.dbo.tbl_mdmenu_category (userid, catecode) "
		sqlStr = sqlStr & " SELECT '"&mduserid&"', catecode "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_display_cate "
		sqlStr = sqlStr & " where depth <= 2 and useyn = 'Y' "
		sqlStr = sqlStr & " and left(catecode,3) = '"&catecode&"' "
		dbget.execute(sqlStr)
		response.write	"<script language='javascript'>" &_
						"	alert('저장되었습니다');" &_
						"	location.replace('/admin/mdMenu/category/?menupos="&menupos&"');" &_
						"</script>"	
	Else
		sqlStr = "IF Exists(select * from db_partner.dbo.tbl_mdmenu_category where catecode='"&catecode&"')"
		sqlStr = sqlStr & " BEGIN"& VbCRLF
		sqlStr = sqlStr & " update R" & VbCRLF
		sqlStr = sqlStr & "	Set userid='"&mduserid&"' "  & VbCRLF
		sqlStr = sqlStr & "	From db_partner.dbo.tbl_mdmenu_category R"& VbCRLF
		sqlStr = sqlStr & " Where R.catecode='" & catecode & "'"
		sqlStr = sqlStr & " END ELSE "
		sqlStr = sqlStr & " BEGIN"& VbCRLF
		sqlStr = sqlStr & " Insert into db_partner.dbo.tbl_mdmenu_category "
	    sqlStr = sqlStr & " (userid, catecode) "
	    sqlStr = sqlStr & " values ('"&mduserid&"', '" & catecode & "')"
		sqlStr = sqlStr & " END "
	    dbget.Execute sqlStr
		response.write	"<script language='javascript'>" &_
						"	alert('저장되었습니다');" &_
						"	location.replace('/admin/mdMenu/category/?menupos="&menupos&"');" &_
						"</script>"	
	End If
ElseIf mode = "D" Then
	If Len(catecode) = 3 Then
		sqlStr = "delete from db_partner.dbo.tbl_mdmenu_category Where left(catecode,3) = '"&catecode&"'"
		dbget.execute(sqlStr)
		response.write	"<script language='javascript'>" &_
						"	alert('삭제되었습니다');" &_
						"	location.replace('/admin/mdMenu/category/?menupos="&menupos&"');" &_
						"</script>"	
	Else
		sqlStr = "delete from db_partner.dbo.tbl_mdmenu_category Where catecode = '"&catecode&"'"
	    dbget.Execute sqlStr
		response.write	"<script language='javascript'>" &_
						"	alert('삭제되었습니다');" &_
						"	location.replace('/admin/mdMenu/category/?menupos="&menupos&"');" &_
						"</script>"	
	End If
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->