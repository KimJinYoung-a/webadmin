<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%

dim ogroup, sql

dim socno
dim subsocno
dim socname
dim ceoname
dim socaddr
dim socstatus
dim socevent
dim managername
dim managerphone
dim managermail

socno = request("socno")



'==========================================================================
SQL = " select top 1 IsNull(busiSubNo, '') as busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel " & VbCRLF
SQL = SQL + " from db_order.[dbo].tbl_busiinfo " & VbCRLF
SQL = SQL + " where busiNo='" + socno + "' " & VbCRLF
SQL = SQL + " order by busiIdx desc "
rsget.Open SQL, dbget, 1
if  not rsget.EOF  then
	subsocno = rsget("busiSubNo")
	socname = rsget("busiName")
	ceoname = rsget("busiCEOName")
	socaddr = rsget("busiAddr")
	socevent = rsget("busiType")
	socstatus = rsget("busiItem")

	managername = rsget("repName")
	managerphone = rsget("repTel")
	managermail = rsget("repEmail")

	if IsNull(managername) then
		 managername = ""
		 managerphone = ""
		 managermail = ""
	end if

end if
rsget.close



if (socname = "") then
		response.write "<script>alert('등록되지 않은 업체입니다.')</script>"
		dbget.close()	:	response.End
end if


subsocno = Replace(subsocno, "'", "")
socname = Replace(socname, "'", "")
ceoname = Replace(ceoname, "'", "")
socaddr = Replace(socaddr, "'", "")
socevent = Replace(socevent, "'", "")
socstatus = Replace(socstatus, "'", "")

managername = Replace(managername, "'", "")
managerphone = Replace(managerphone, "'", "")
managermail = Replace(managermail, "'", "")

%>


<script language='javascript'>
parent.setCompanyInfo('<%= subsocno %>', '<%= socname %>', '<%= ceoname %>', '<%= socaddr %>', '<%= socevent %>', '<%= socstatus %>', '<%= managername %>', '<%= managerphone %>', '<%= managermail %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->