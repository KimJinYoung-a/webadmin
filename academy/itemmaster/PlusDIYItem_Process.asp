<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim itemid, PlusSalePro, PlusSaleMaginFlag, PlusSaleMargin
dim termsGubun, PlusSaleStartDate, PlusSaleEndDate
dim mode
dim itemidarr, PlusSaleItemID, PlusSaleLinkItemid, chkitem

itemid              = RequestCheckVar(request("itemid"),9)
PlusSalePro          = RequestCheckVar(request("PlusSalePro"),9)
PlusSaleMaginFlag    = RequestCheckVar(request("PlusSaleMaginFlag"),9)
PlusSaleMargin       = RequestCheckVar(request("PlusSaleMargin"),9)
termsGubun          = RequestCheckVar(request("termsGubun"),9)
PlusSaleStartDate    = RequestCheckVar(request("PlusSaleStartDate"),10)
PlusSaleEndDate      = RequestCheckVar(request("PlusSaleEndDate"),10)
mode                = RequestCheckVar(request("mode"),32)

PlusSaleItemID      = RequestCheckVar(request("PlusSaleItemID"),9)
PlusSaleLinkItemid  = RequestCheckVar(request("PlusSaleLinkItemid"),9)
itemidarr           = request("itemidarr")
chkitem             = request("chkitem")

  	if itemidarr <> "" then
		if checkNotValidHTML(itemidarr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end If
  	if chkitem <> "" then
		if checkNotValidHTML(chkitem) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if

PlusSaleEndDate      = PlusSaleEndDate + " 23:59:59"


dim i,cnt
dim sqlStr, AssignedRows
dim Tot_AssignedRows
if (mode="regPlusSale") or (mode="editPlusSale") then 
    sqlStr = "IF Exists(select * from db_academy.dbo.tbl_diy_PlusSaleRegedItem where PlusSaleItemID=" & itemid & ")" & VbCrlf
    sqlStr = sqlStr + "     update db_academy.dbo.tbl_diy_PlusSaleRegedItem" & VbCrlf
    sqlStr = sqlStr + "     set PlusSalePro=" & PlusSalePro & VbCrlf
    sqlStr = sqlStr + "     , PlusSaleMargin=" & PlusSaleMargin & "" & VbCrlf
    sqlStr = sqlStr + "     , PlusSaleMaginFlag='" & PlusSaleMaginFlag & "'" & VbCrlf
    sqlStr = sqlStr + "     , PlusSaleStartDate='" & PlusSaleStartDate & "'" & VbCrlf
    sqlStr = sqlStr + "     , PlusSaleEndDate='" & PlusSaleEndDate & "'" & VbCrlf
    sqlStr = sqlStr + "     where PlusSaleItemID=" & itemid & VbCrlf
    sqlStr = sqlStr + " ELSE " & VbCrlf
    sqlStr = sqlStr + "     insert into db_academy.dbo.tbl_diy_PlusSaleRegedItem" & VbCrlf
    sqlStr = sqlStr + "     (PlusSaleItemID,PlusSalePro,PlusSaleMargin, PlusSaleMaginFlag, PlusSaleStartDate, PlusSaleEndDate)" & VbCrlf
    sqlStr = sqlStr + "     values(" & VbCrlf
    sqlStr = sqlStr + "     " & itemid & VbCrlf
    sqlStr = sqlStr + "     ," & PlusSalePro & VbCrlf
    sqlStr = sqlStr + "     ," & PlusSaleMargin & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleMaginFlag & "'" & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleStartDate & "'" & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleEndDate & "'" & VbCrlf
    sqlStr = sqlStr + "     )"
    
    rsACADEMYget.Open sqlStr,dbACADEMYget,1

elseif (mode="delPlusSale") then
    ''링크 삭제
    sqlStr = "delete from db_academy.dbo.tbl_diy_PlusSaleLinkItem"
    sqlStr = sqlStr + " where plusSaleItemID=" & itemid
    dbACADEMYget.Execute sqlStr, AssignedRows
    
    response.write "<script>alert('메인 링크 " & AssignedRows & " 개 삭제');</script>"
    
    ''플러스세일상품 삭제
    sqlStr = "delete from db_academy.dbo.tbl_diy_PlusSaleRegedItem"
    sqlStr = sqlStr + " where plusSaleItemID=" & itemid
    dbACADEMYget.Execute sqlStr, AssignedRows
    
    response.write "<script>alert('플러스 세일 상품 " & AssignedRows & " 개 삭제');</script>"

elseif (mode="PlusMainAddArr") then
    ''PlusSaleItemID
    Tot_AssignedRows = 0
    itemidarr = split(itemidarr,",")
    
    if IsArray(itemidarr) then
        for i=LBound(itemidarr) to UBound(itemidarr)
            PlusSaleLinkItemid = Trim(itemidarr(i))
            if (PlusSaleLinkItemid<>"") then
                if (PlusSaleLinkItemid<>PlusSaleItemID) then
                    sqlStr = " insert into db_academy.dbo.tbl_diy_PlusSaleLinkItem"
                    sqlStr = sqlStr + " (plusSaleItemid, plusSaleLinkItemid)"
                    
                    sqlStr = sqlStr + " select r.plusSaleItemid, " & PlusSaleLinkItemid
                    sqlStr = sqlStr + "   from db_academy.dbo.tbl_diy_PlusSaleRegedItem r"
                    sqlStr = sqlStr + " 	left join  db_academy.dbo.tbl_diy_PlusSaleLinkItem s"
                    sqlStr = sqlStr + "	    on r.plusSaleItemid=s.plusSaleItemid"
                    sqlStr = sqlStr + "	    and s.plusSaleLinkItemID=" & plusSaleLinkItemid
                    sqlStr = sqlStr + " where r.plusSaleItemid=" & PlusSaleItemID
                    sqlStr = sqlStr + " and s.plusSaleLinkItemID is NULL"
                    sqlStr = sqlStr + " and s.plusSaleLinkItemID Not in ("
                    sqlStr = sqlStr + "     select plusSaleItemid from db_academy.dbo.tbl_diy_PlusSaleRegedItem where plusSaleItemid=" & plusSaleLinkItemid
                    sqlStr = sqlStr + " )"

                    dbACADEMYget.Execute sqlStr, AssignedRows
                    
                    Tot_AssignedRows = Tot_AssignedRows + AssignedRows
                end if
            end if
        next
    end if
elseif (mode="PlusMainDellArr") then
    Tot_AssignedRows = 0
    chkitem = split(chkitem,",")
    
    if IsArray(chkitem) then
        for i=LBound(chkitem) to UBound(chkitem)
            PlusSaleLinkItemid = Trim(chkitem(i))
            if (PlusSaleLinkItemid<>"") then
                sqlStr = " delete from db_academy.dbo.tbl_diy_PlusSaleLinkItem"
                sqlStr = sqlStr + " where plusSaleItemid=" & PlusSaleItemID
                sqlStr = sqlStr + " and plusSaleLinkItemID=" & PlusSaleLinkItemid

                dbACADEMYget.Execute sqlStr, AssignedRows
                
                Tot_AssignedRows = Tot_AssignedRows + AssignedRows
            end if
        next
    end if
elseif (mode="PlusSaleAddArr") then
    Tot_AssignedRows = 0
    chkitem = split(chkitem,",")
    
    if IsArray(chkitem) then
        for i=LBound(chkitem) to UBound(chkitem)
            PlusSaleItemid = Trim(chkitem(i))
            if (PlusSaleItemid<>"") then
                if (PlusSaleLinkItemid<>PlusSaleItemid) then
                    
                    sqlStr = " insert into db_academy.dbo.tbl_diy_PlusSaleLinkItem"
                    sqlStr = sqlStr + " (plusSaleItemid, plusSaleLinkItemid)"
                    
                    sqlStr = sqlStr + " select r.plusSaleItemid, " & PlusSaleLinkItemid
                    sqlStr = sqlStr + "   from db_academy.dbo.tbl_diy_PlusSaleRegedItem r"
                    sqlStr = sqlStr + " 	left join  db_academy.dbo.tbl_diy_PlusSaleLinkItem s"
                    sqlStr = sqlStr + "	    on r.plusSaleItemid=s.plusSaleItemid"
                    sqlStr = sqlStr + "	    and s.plusSaleLinkItemID=" & plusSaleLinkItemid
                    sqlStr = sqlStr + " where r.plusSaleItemid=" & PlusSaleItemID
                    sqlStr = sqlStr + " and s.plusSaleLinkItemID is NULL"

                    dbACADEMYget.Execute sqlStr, AssignedRows
                    
                    Tot_AssignedRows = Tot_AssignedRows + AssignedRows
                end if
            end if
        next
    end if
elseif (mode="PlusSaleDellArr") then
    Tot_AssignedRows = 0
    chkitem = split(chkitem,",")
    
    if IsArray(chkitem) then
        for i=LBound(chkitem) to UBound(chkitem)
            PlusSaleItemid = Trim(chkitem(i))
            if (PlusSaleItemid<>"") then
                sqlStr = " delete from db_academy.dbo.tbl_diy_PlusSaleLinkItem"
                sqlStr = sqlStr + " where plusSaleItemid=" & PlusSaleItemID
                sqlStr = sqlStr + " and plusSaleLinkItemID=" & PlusSaleLinkItemid

                dbACADEMYget.Execute sqlStr, AssignedRows
                
                Tot_AssignedRows = Tot_AssignedRows + AssignedRows
            end if
        next
    end if
end if

%>

<% if (mode="regPlusSale") then %>
    <script language='javascript'>
    alert('등록 되었습니다.');
    opener.location.reload();
	location.replace('<%= refer %>');
	</script>
<% elseif mode="editPlusSale" then %>
    <script language='javascript'>
    alert('수정 되었습니다.');
    opener.location.reload();
	location.replace('<%= refer %>');
	</script>
<% elseif mode="delPlusSale" then %>
    <script language='javascript'>
    alert('삭제 되었습니다.');
    opener.location.reload();
    window.close();
	//location.replace('<%= refer %>');
	</script>	
<% elseif (mode="PlusMainAddArr") then %>
    <script language='javascript'>
    alert('<%= Tot_AssignedRows %>건 추가 되었습니다.');
    //opener.location.reload();
    //window.close();
	</script>	
<% elseif (mode="PlusMainDellArr") then %>
    <script language='javascript'>
    alert('<%= Tot_AssignedRows %>건 삭제 되었습니다.');
    location.replace('<%= refer %>');
	</script>		
<% elseif (mode="PlusSaleDellArr") then %>
    <script language='javascript'>
    alert('<%= Tot_AssignedRows %>건 삭제 되었습니다.');
    location.replace('<%= refer %>');
	</script>		
<% elseif (mode="PlusSaleAddArr") then %>
    <script language='javascript'>
    alert('<%= Tot_AssignedRows %>건 추가 되었습니다.');
    opener.location.reload();
    location.replace('<%= refer %>');
	</script>	
<% else %>
    <script language='javascript'>
    alert('정의 되지 않았습니다. - <%= mode %>');
	location.replace('<%= refer %>');
	</script>
<% end if %>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->