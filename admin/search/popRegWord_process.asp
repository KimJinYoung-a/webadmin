<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, sqlStr
dim aliasWord, mainWord

mode = requestCheckVar(request("mode"), 32)
aliasWord = requestCheckVar(Trim(request("aliasWord")), 32)
mainWord = requestCheckVar(Trim(request("mainWord")), 32)

select case mode
	case "ins"
		sqlStr = " if not exists(select top 1 mainWord from [db_analyze].[dbo].[tbl_word_mainWord] where mainWord = '" & html2db(mainWord) & "') "
		sqlStr = sqlStr & " begin "
		sqlStr = sqlStr & " 	insert into [db_analyze].[dbo].[tbl_word_mainWord](mainWord, wordType) "
		sqlStr = sqlStr & " 	values('" & html2db(mainWord) & "', 'TP') "
		sqlStr = sqlStr & " end "
		sqlStr = sqlStr & " else "
		sqlStr = sqlStr & " begin "
		sqlStr = sqlStr & " 	update [db_analyze].[dbo].[tbl_word_mainWord] "
		sqlStr = sqlStr & " 	set wordType = 'TP', lastupdate = getdate() "
		sqlStr = sqlStr & " 	where mainWord = '" & html2db(mainWord) & "' "
		sqlStr = sqlStr & " end "
		dbEVTget.Execute sqlStr

		sqlStr = " if not exists(select top 1 mainWord from [db_analyze].[dbo].[tbl_word_aliasWord] where aliasWord = '" & html2db(aliasWord) & "') "
		sqlStr = sqlStr & " begin "
		sqlStr = sqlStr & " 	insert into [db_analyze].[dbo].[tbl_word_aliasWord](aliasWord, mainWord) "
		sqlStr = sqlStr & " 	values('" & html2db(aliasWord) & "', '" & html2db(mainWord) & "') "
		sqlStr = sqlStr & " end "
		sqlStr = sqlStr & " else "
		sqlStr = sqlStr & " begin "
		sqlStr = sqlStr & " 	update [db_analyze].[dbo].[tbl_word_aliasWord] "
		sqlStr = sqlStr & " 	set mainWord = '" & html2db(mainWord) & "', lastupdate = getdate() "
		sqlStr = sqlStr & " 	where aliasWord = '" & html2db(aliasWord) & "' "
		sqlStr = sqlStr & " end "
		dbEVTget.Execute sqlStr

		sqlStr = " if not exists(select top 1 mainWord from [db_analyze].[dbo].[tbl_word_aliasWord] where aliasWord = '" & html2db(mainWord) & "') "
		sqlStr = sqlStr & " begin "
		sqlStr = sqlStr & " 	insert into [db_analyze].[dbo].[tbl_word_aliasWord](aliasWord, mainWord) "
		sqlStr = sqlStr & " 	values('" & html2db(mainWord) & "', '" & html2db(mainWord) & "') "
		sqlStr = sqlStr & " end "
		dbEVTget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
						"	alert('저장되었습니다.'); location.href = '" + CStr(refer) + "' " &_
						"</script>"
	case else
		response.end
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
