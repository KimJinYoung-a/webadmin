<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%

dim mode
dim sitename, outmallorderserial, outmallorderseq, shppDivDtlNm
dim asid, cnt

Dim sqlStr

mode = RequestCheckVar(Request("mode"),32)
sitename = RequestCheckVar(Request("sitename"),32)
outmallorderserial = RequestCheckVar(Request("outmallorderserial"),32)
outmallorderseq = RequestCheckVar(Request("outmallorderseq"),32)
shppDivDtlNm = RequestCheckVar(Request("shppDivDtlNm"),32)

select case mode
    case "matchcs"
        sqlStr = " select min(a.id) as asid, count(a.id) as cnt "
        sqlStr = sqlStr + " 	from "
        sqlStr = sqlStr + " 	[db_temp].[dbo].[tbl_xSite_TMPMiChulList] m "
        sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_list] a "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and m.Matchorderserial = a.orderserial "
        sqlStr = sqlStr + " 		and ( "
        sqlStr = sqlStr + " 		(m.shppDivDtlNm = '교환출고' and a.divcd = 'A000' and a.deleteyn = 'N') "
        sqlStr = sqlStr + " 		) "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and m.SellSite = '" & sitename & "' "
        sqlStr = sqlStr + " 	and m.OutMallOrderSerial = '" & outmallorderserial & "' "
        sqlStr = sqlStr + " 	and m.OrgDetailKey = '" & outmallorderseq & "' "
        sqlStr = sqlStr + " 	and m.shppDivDtlNm = '" & shppDivDtlNm & "' "
        sqlStr = sqlStr + " 	and m.asid is NULL "

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            asid = rsget("asid")
            cnt = rsget("cnt")
        end if
        rsget.Close

        if (cnt = 0) then
            rw "ERR : 매칭실패, 검색된 CS건이 없습니다."
            dbget.close() : response.end
        elseif (cnt > 1) then
            rw "ERR : 매칭실패, 두개 이상의 CS건이 있습니다."
            dbget.close() : response.end
        end if

        sqlStr = " update m "
        sqlStr = sqlStr + " set m.asid = " & asid
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_temp].[dbo].[tbl_xSite_TMPMiChulList] m "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and m.SellSite = '" & sitename & "' "
        sqlStr = sqlStr + " 	and m.OutMallOrderSerial = '" & outmallorderserial & "' "
        sqlStr = sqlStr + " 	and m.OrgDetailKey = '" & outmallorderseq & "' "
        sqlStr = sqlStr + " 	and m.shppDivDtlNm = '" & shppDivDtlNm & "' "
        sqlStr = sqlStr + " 	and m.asid is NULL "
        dbget.Execute sqlStr

        response.write "<script>alert('매칭되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
	case else
        '
end select

%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
