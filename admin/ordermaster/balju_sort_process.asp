<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 출고지시서 정렬순서관리
' History : 2020.12.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->
<%
dim i, menupos, check
dim midx, title, comment, isusing, didx, rackcode, sortno, strSql, mode, adminid, layer
	mode		=  requestCheckVar(Request("mode"),42)
    menupos = requestCheckVar(getNumeric(request("menupos")),10)
    midx = requestCheckVar(getNumeric(request("midx")),10)
    title = requestCheckVar(request("title"),128)
    comment = requestCheckVar(request("comment"),1000)
    isusing = requestCheckVar(request("isusing"),1)
    didx = request("didx")
    rackcode = request("rackcode")
    sortno = request("sortno")
    layer = request("layer")

adminid = session("ssBctId")

dim refer
    refer = request.ServerVariables("HTTP_REFERER")

if mode = "baljusortreg" then
	if title <> "" and not(isnull(title)) then
    	title = ReplaceBracket(title)

		'if checkNotValidHTML(title) then
        '    response.write "<script type='text/javascript'>"
        '    response.write "	alert('제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
        '    response.write "</script>"
        '    response.End
		'end if
	end if
	if comment <> "" and not(isnull(comment)) then
    	comment = ReplaceBracket(comment)

		'if checkNotValidHTML(comment) then
        '    response.write "<script type='text/javascript'>"
        '    response.write "	alert('코맨트에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
        '    response.write "</script>"
        '    response.End
		'end if
	end if

	'//신규저장	
	If midx = "" Then
		strSql = " INSERT INTO db_aLogistics.dbo.tbl_chulgo_sheet_sort_master (" & VbCRLF
		strSql = strSql & " title,comment,isusing,regdate,lastupdate,regadminid,lastadminid" & VbCRLF
		strSql = strSql & " ) VALUES (" & VbCRLF
		strSql = strSql & " '" & html2db(trim(title)) & "','" & html2db(trim(comment)) & "', '" & trim(isusing) & "', getdate()" & VbCRLF
		strSql = strSql & " , getdate(), '"& adminid &"', '"& adminid &"'" & VbCRLF
		strSql = strSql & " )"
		
		'response.write strSql &"<br>"
		dbget_Logistics.execute strSql
    	
		strSql = ""
		strSql = " SELECT IDENT_CURRENT('db_aLogistics.dbo.tbl_chulgo_sheet_sort_master')"

		'response.write strSql &"<br>"
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open strSql, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
		
		IF Not rsget_Logistics.EOF THEN
			midx = rsget_Logistics(0)
		ELSE	
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.end
		END IF
		rsget_Logistics.close

        if rackcode <> "" and layer <> "" and sortno <> "" then
            rackcode = split(rackcode,",")
            layer = split(layer,",")
            sortno = split(sortno,",")
            
            if isarray(rackcode) and isarray(layer) and isarray(sortno) then
                for i = 0 to ubound(rackcode)

                if trim(rackcode(i))<>"" and trim(getNumeric(layer(i)))<>"" and trim(getNumeric(sortno(i)))<>"" then
                    strSql = " insert into db_aLogistics.dbo.tbl_chulgo_sheet_sort_detail ("&VbCRLF
                    strSql = strSql & " midx, rackcode, layer, sortno) values"&VbCRLF
                    strSql = strSql & " ("& trim(midx) &",'"& trim(html2db(rackcode(i))) &"',"& trim(getNumeric(layer(i))) &","& trim(getNumeric(sortno(i))) &")"
                    
                    'response.write strSql &"<br>"
                    dbget_Logistics.execute strSql
                end if
                next
            end if
        end if

	'//수정	
	Else
		strSql = " UPDATE db_aLogistics.dbo.tbl_chulgo_sheet_sort_master" & VbCRLF
		strSql = strSql & " set title = '" & html2db(trim(title)) & "'" & VbCRLF
		strSql = strSql & " , comment = '" & html2db(trim(comment)) & "'" & VbCRLF
		strSql = strSql & " , isusing = '" & trim(isusing) & "'" & VbCRLF
		strSql = strSql & " , lastupdate = getdate()" & VbCRLF
		strSql = strSql & " , lastadminid = '" & adminid & "' WHERE" & VbCRLF
		strSql = strSql & " midx = " & midx & ""
		
		'response.write strSql &"<br>"
		dbget_Logistics.execute strSql

	    strSql = " delete from db_aLogistics.dbo.tbl_chulgo_sheet_sort_detail where midx="& midx &""&VbCRLF

        'response.write strSql &"<br>"
        dbget_Logistics.Execute strSql

        if rackcode <> "" and layer <> "" and sortno <> "" then
            rackcode = split(rackcode,",")
            layer = split(layer,",")
            sortno = split(sortno,",")
            
            if isarray(rackcode) and isarray(layer) and isarray(sortno) then
                for i = 0 to ubound(rackcode)

                if trim(rackcode(i))<>"" and trim(getNumeric(layer(i)))<>"" and trim(getNumeric(sortno(i)))<>"" then
                    strSql = " insert into db_aLogistics.dbo.tbl_chulgo_sheet_sort_detail ("&VbCRLF
                    strSql = strSql & " midx, rackcode, layer, sortno) values"&VbCRLF
                    strSql = strSql & " ("& trim(midx) &",'"& trim(html2db(rackcode(i))) &"',"& trim(getNumeric(layer(i))) &","& trim(getNumeric(sortno(i))) &")"
                    
                    'response.write strSql &"<br>"
                    dbget_Logistics.execute strSql
                end if
                next
            end if
        end if
    	        
	End If

    Response.Write "<script type='text/javascript'>"
    Response.Write "    alert('저장되었습니다.');"
    Response.Write "    opener.location.reload();"
    Response.Write "    self.location.replace('/admin/ordermaster/balju_sort_reg.asp?midx="& midx &"&menupos="& menupos &"')</script>"
    response.write "</script>"
    response.end

elseif mode = "baljumasterdel" then
    for i=1 to request("check").count
        midx = requestCheckVar(getNumeric(trim(request("check")(i))),10)

        if midx <> "" then
            strSql = "delete from db_aLogistics.dbo.tbl_chulgo_sheet_sort_master where midx="& midx &""

            response.write strSql &"<br>"
            dbget_Logistics.execute strSql
        end if
    next

    Response.Write "<script type='text/javascript'>"
    Response.Write "    alert('삭제되었습니다.');"
    Response.Write "    parent.location.reload();"
    response.write "</script>"
    response.end

elseif mode = "baljumasterreg" then
    for i=1 to request("check").count
        midx = requestCheckVar(getNumeric(trim(request("check")(i))),10)
        sortno = requestCheckVar(getNumeric(trim(request("sortno_"&midx))),10)

        if midx <> "" and sortno <> "" then
            strSql = "update db_aLogistics.dbo.tbl_chulgo_sheet_sort_master set sortno="& sortno &" where midx="& midx &""

            'response.write strSql &"<br>"
            dbget_Logistics.execute strSql
        end if
    next

    Response.Write "<script type='text/javascript'>"
    Response.Write "    alert('저장되었습니다.');"
    Response.Write "    parent.location.reload();"
    response.write "</script>"
    response.end
end if

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
