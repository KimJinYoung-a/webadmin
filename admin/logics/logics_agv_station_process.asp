<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : agv
' History : 이상구 생성
'           2020.05.12 정태훈 수정
'           2020.05.20 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode
dim stationCd, stationName, stationGubun, sortNo, regdate, updt, useYN
dim sqlStr, addStr, i, affectedRows
dim errMsg

mode = requestcheckvar(request("mode"),32)
stationCd = requestcheckvar(request("stationCd"),32)
stationName = html2db(requestcheckvar(request("stationName"),32))
stationGubun = requestcheckvar(request("stationGubun"),32)
sortNo = requestcheckvar(request("sortNo"),32)
useYN = requestcheckvar(request("useYN"),32)

select case mode
    case "addstation"
        sqlStr = " if exists (select top 1 stationCd from [db_aLogistics].[dbo].[tbl_agv_stationInfo] where stationCd = '" & stationCd & "' and useYN = 'N') "
        sqlStr = sqlStr + " begin "
        sqlStr = sqlStr + " 	update [db_aLogistics].[dbo].[tbl_agv_stationInfo] "
        sqlStr = sqlStr + " 	set useYN = 'Y', updt = getdate(), stationName = '" & stationName & "', stationGubun = '" & stationGubun & "', sortNo = " & sortNo & " "
        sqlStr = sqlStr + " 	where stationCd = '" & stationCd & "' "
        sqlStr = sqlStr + " end "
        sqlStr = sqlStr + " else if not exists (select top 1 stationCd from [db_aLogistics].[dbo].[tbl_agv_stationInfo] where stationCd = '" & stationCd & "' and useYN = 'Y') "
        sqlStr = sqlStr + " begin "
        sqlStr = sqlStr + " 	insert into [db_aLogistics].[dbo].[tbl_agv_stationInfo](stationCd, stationName, stationGubun, sortNo) "
        sqlStr = sqlStr + " 	values('" & stationCd & "', '" & stationName & "', '" & stationGubun & "', " & sortNo & ") "
        sqlStr = sqlStr + " end "
        dbget_Logistics.Execute sqlStr, affectedRows

        if (affectedRows > 0) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('저장 되었습니다.');"
            response.write "	opener.focus(); opener.location.reload(); window.close();"
            response.write "</script>"
        else
            response.write "<script type='text/javascript'>"
            response.write "	alert('이미 존재하는 스테이션코드입니다.');"
            response.write "	history.back();"
            response.write "</script>"
        end if
    case "editstation"
        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_stationInfo] "
        sqlStr = sqlStr + " 	set updt = getdate(), stationName = '" & stationName & "', stationGubun = '" & stationGubun & "', sortNo = " & sortNo & " "
        sqlStr = sqlStr + " 	where stationCd = '" & stationCd & "' "
        dbget_Logistics.Execute sqlStr, affectedRows

        if (affectedRows > 0) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('저장 되었습니다.');"
            response.write "	opener.focus(); opener.location.reload(); window.close();"
            response.write "</script>"
        else
            response.write "<script type='text/javascript'>"
            response.write "	alert('잘못된 접근입니다.');"
            response.write "	history.back();"
            response.write "</script>"
        end if
    case "delstation"
        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_stationInfo] "
        sqlStr = sqlStr + " 	set updt = getdate(), useYN = 'N' "
        sqlStr = sqlStr + " 	where stationCd = '" & stationCd & "' "
        dbget_Logistics.Execute sqlStr, affectedRows

        if (affectedRows > 0) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('삭제 되었습니다.');"
            response.write "	opener.focus(); opener.location.reload(); window.close();"
            response.write "</script>"
        else
            response.write "<script type='text/javascript'>"
            response.write "	alert('잘못된 접근입니다.');"
            response.write "	history.back();"
            response.write "</script>"
        end if
    case else
        response.write "잘못된 접근입니다."
end select

%>
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
