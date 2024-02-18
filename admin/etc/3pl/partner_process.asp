<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim mode
dim companyid, partnercompanyid, partnercompanyname, apiAvail, useyn
dim sqlStr

mode        		= requestCheckVar(request("mode"),32)
companyid    		= requestCheckVar(request("companyid"),32)
partnercompanyid    = requestCheckVar(request("partnercompanyid"),32)
partnercompanyname  = html2db(requestCheckVar(request("partnercompanyname"),32))
apiAvail       		= requestCheckVar(request("apiAvail"),32)
useyn        		= requestCheckVar(request("useyn"),32)

select case mode
	case "modi"
		sqlStr = ""
		sqlStr = sqlStr & " update [db_threepl].[dbo].[tbl_partnerinfo] "
		sqlStr = sqlStr & " set lastupdt = getdate(), "
		sqlStr = sqlStr & " partnercompanyname = '" & partnercompanyname & "', "
		sqlStr = sqlStr & " useyn = '" & useyn & "' "
		sqlStr = sqlStr & " where partnercompanyid = " & partnercompanyid
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('수정 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case "ins"
		sqlStr = ""
		sqlStr = sqlStr & " insert into [db_threepl].[dbo].[tbl_partnerinfo](partnercompanyname, useyn) "
		sqlStr = sqlStr & " values('" & partnercompanyname & "', '" & useyn & "')"
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('저장 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case "modiLink"
		sqlStr = ""
		sqlStr = sqlStr & " update c "
		sqlStr = sqlStr & " set c.partnercompanyname = p.partnercompanyname, c.apiAvail = '" & apiAvail & "', c.useyn = '" & useyn & "' "
		sqlStr = sqlStr & " from "
		sqlStr = sqlStr & " 	[db_threepl].[dbo].[tbl_partnercompany] c "
		sqlStr = sqlStr & " 	join [db_threepl].[dbo].[tbl_partnerinfo] p on c.partnercompanyid = p.partnercompanyid "
		sqlStr = sqlStr & " where "
		sqlStr = sqlStr & " 	1 = 1 "
		sqlStr = sqlStr & " 	and c.companyid = '" & companyid & "' "
		sqlStr = sqlStr & " 	and c.partnercompanyid = '" & partnercompanyid & "' "
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('수정 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case "insLink"
		sqlStr = ""
		sqlStr = sqlStr & " insert into [db_threepl].[dbo].[tbl_partnercompany](companyid, partnercompanyid, partnercompanyname, apiAvail, useyn) "
		sqlStr = sqlStr & " select top 1 c.companyid, p.partnercompanyid, p.partnercompanyname, '" & apiAvail & "', '" & useyn & "' "
		sqlStr = sqlStr & " from "
		sqlStr = sqlStr & " 	[db_threepl].[dbo].[tbl_company] c "
		sqlStr = sqlStr & " 	join [db_threepl].[dbo].[tbl_partnerinfo] p "
		sqlStr = sqlStr & " 	on "
		sqlStr = sqlStr & " 		1 = 1 "
		sqlStr = sqlStr & " 		and c.companyid = '" & companyid & "' "
		sqlStr = sqlStr & " 		and p.partnercompanyid = '" & partnercompanyid & "' "
		sqlStr = sqlStr & " 	left join [db_threepl].[dbo].[tbl_partnercompany] pc "
		sqlStr = sqlStr & " 	on "
		sqlStr = sqlStr & " 		1 = 1 "
		sqlStr = sqlStr & " 		and pc.companyid = c.companyid "
		sqlStr = sqlStr & " 		and pc.partnercompanyid = p.partnercompanyid "
		sqlStr = sqlStr & " where "
		sqlStr = sqlStr & " 	pc.companyid is NULL "
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('저장 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case else
		response.write "에러"
end select

%>
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
