<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
function AddPlusLinkLog(plusSaleitemid, plusSaleLinkitemid, isDelete)
    dim sqlStr, upflag
    if (plusSaleitemid<>0) and (plusSaleLinkitemid<>0) then
        if (isDelete) then
            sqlStr = "insert into db_log.dbo.tbl_PlusSaleLinkItemList_Log"
            sqlStr = sqlStr & " (plusSaleItemID, plusSaleLinkitemid, upflag)"
            sqlStr = sqlStr & " values(" & plusSaleitemid & "," & plusSaleLinkitemid & "," & "'D')"
            
            dbget.Execute sqlStr
        else
            sqlStr = "IF Not Exists(select * from db_item.dbo.tbl_PlusSaleLinkItemList where PlusSaleItemID=" & plusSaleitemid & " and plusSaleLinkitemid=" & plusSaleLinkitemid & ")" & VbCrlf
            sqlStr = sqlStr & " insert into db_log.dbo.tbl_PlusSaleLinkItemList_Log"
            sqlStr = sqlStr & " (plusSaleItemID, plusSaleLinkitemid, upflag)"
            sqlStr = sqlStr & " values(" & plusSaleitemid & "," & plusSaleLinkitemid & "," & "'I')"
            
            dbget.Execute sqlStr
        end if
    elseif (plusSaleitemid<>0) then
        if (isDelete) then
            sqlStr = "insert into db_log.dbo.tbl_PlusSaleLinkItemList_Log"
            sqlStr = sqlStr & " (plusSaleItemID, plusSaleLinkitemid, upflag)"
            sqlStr = sqlStr & " select plusSaleItemID, plusSaleLinkitemid,'D'"
            sqlStr = sqlStr & " from db_item.dbo.tbl_PlusSaleLinkItemList"
            sqlStr = sqlStr & " where plusSaleitemid=" & plusSaleitemid
            
            dbget.Execute sqlStr
        end if
    elseif (plusSaleLinkitemid<>0) then
        if (isDelete) then
            sqlStr = "insert into db_log.dbo.tbl_PlusSaleLinkItemList_Log"
            sqlStr = sqlStr & " (plusSaleItemID, plusSaleLinkitemid, upflag)"
            sqlStr = sqlStr & " select plusSaleItemID, plusSaleLinkitemid,'D'"
            sqlStr = sqlStr & " from db_item.dbo.tbl_PlusSaleLinkItemList"
            sqlStr = sqlStr & " where plusSaleLinkitemid=" & plusSaleLinkitemid
            
            dbget.Execute sqlStr
        end if
    end if

    
    
end function

function AddPlusItemLog(itemid, PlusSalePro, PlusSaleMargin, PlusSaleMaginFlag, PlusSaleStartDate, PlusSaleEndDate, isDelete)
    dim sqlStr, upflag
    dim AlreadyExists
    dim pPlusSalePro, pPlusSaleMargin, pPlusSaleMaginFlag, pPlusSaleStartDate, pPlusSaleEndDate
    
    sqlStr = "select * from db_item.dbo.tbl_PlusSaleRegedItem where PlusSaleItemID=" & itemid
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        AlreadyExists       = true
        pPlusSalePro        = rsget("PlusSalePro")
        pPlusSaleMargin     = rsget("PlusSaleMargin")
        pPlusSaleMaginFlag  = rsget("PlusSaleMaginFlag")
        pPlusSaleStartDate  = rsget("PlusSaleStartDate")
        pPlusSaleEndDate    = rsget("PlusSaleEndDate")
    end if
    rsget.Close
    
    if (isDelete) then
        PlusSalePro        = pPlusSalePro
        PlusSaleMargin     = pPlusSaleMargin
        PlusSaleMaginFlag  = pPlusSaleMaginFlag
        PlusSaleStartDate  = pPlusSaleStartDate
        PlusSaleEndDate    = pPlusSaleEndDate
        upflag = "D"
    elseif (AlreadyExists) then
        if (pPlusSalePro<>PlusSalePro) or (pPlusSaleMargin<>PlusSaleMargin) _
           or (pPlusSaleMaginFlag<>PlusSaleMaginFlag) or (pPlusSaleStartDate<>PlusSaleStartDate) _
           or (pPlusSaleStartDate<>PlusSaleStartDate) or (pPlusSaleEndDate<>PlusSaleEndDate) then
                upflag = "M"
        end if        
    else
        upflag = "I"
    end if
    
    
    sqlStr = sqlStr + "     insert into db_log.dbo.tbl_PlusSaleRegedItem_Log" & VbCrlf
    sqlStr = sqlStr + "     (PlusSaleItemID,PlusSalePro,PlusSaleMargin, PlusSaleMaginFlag, PlusSaleStartDate, PlusSaleEndDate, upflag)" & VbCrlf
    sqlStr = sqlStr + "     values(" & VbCrlf
    sqlStr = sqlStr + "     " & itemid & VbCrlf
    sqlStr = sqlStr + "     ," & PlusSalePro & VbCrlf
    sqlStr = sqlStr + "     ," & PlusSaleMargin & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleMaginFlag & "'" & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleStartDate & "'" & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleEndDate & "'" & VbCrlf
    sqlStr = sqlStr + "     ,'" & upflag & "'"
    sqlStr = sqlStr + "     )"
    
    dbget.Execute sqlStr
end function


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

PlusSaleEndDate      = PlusSaleEndDate + " 23:59:59"


dim i,cnt
dim sqlStr, AssignedRows
dim Tot_AssignedRows
if (mode="regPlusSale") or (mode="editPlusSale") then 
    Call AddPlusItemLog(itemid, PlusSalePro, PlusSaleMargin, PlusSaleMaginFlag, PlusSaleStartDate, PlusSaleEndDate,false)
    
    sqlStr = "IF Exists(select * from db_item.dbo.tbl_PlusSaleRegedItem where PlusSaleItemID=" & itemid & ")" & VbCrlf
    sqlStr = sqlStr + "     update db_item.dbo.tbl_PlusSaleRegedItem" & VbCrlf
    sqlStr = sqlStr + "     set PlusSalePro=" & PlusSalePro & VbCrlf
    sqlStr = sqlStr + "     , PlusSaleMargin=" & PlusSaleMargin & "" & VbCrlf
    sqlStr = sqlStr + "     , PlusSaleMaginFlag='" & PlusSaleMaginFlag & "'" & VbCrlf
    sqlStr = sqlStr + "     , PlusSaleStartDate='" & PlusSaleStartDate & "'" & VbCrlf
    sqlStr = sqlStr + "     , PlusSaleEndDate='" & PlusSaleEndDate & "'" & VbCrlf
    sqlStr = sqlStr + "     where PlusSaleItemID=" & itemid & VbCrlf
    sqlStr = sqlStr + " ELSE " & VbCrlf
    sqlStr = sqlStr + "     insert into db_item.dbo.tbl_PlusSaleRegedItem" & VbCrlf
    sqlStr = sqlStr + "     (PlusSaleItemID,PlusSalePro,PlusSaleMargin, PlusSaleMaginFlag, PlusSaleStartDate, PlusSaleEndDate)" & VbCrlf
    sqlStr = sqlStr + "     values(" & VbCrlf
    sqlStr = sqlStr + "     " & itemid & VbCrlf
    sqlStr = sqlStr + "     ," & PlusSalePro & VbCrlf
    sqlStr = sqlStr + "     ," & PlusSaleMargin & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleMaginFlag & "'" & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleStartDate & "'" & VbCrlf
    sqlStr = sqlStr + "     ,'" & PlusSaleEndDate & "'" & VbCrlf
    sqlStr = sqlStr + "     )"
    
    rsget.Open sqlStr,dbget,1

elseif (mode="delPlusSale") then
    ''��ũ Log
    Call AddPlusLinkLog(itemid,0,true)
    
    ''��ũ ����
    sqlStr = "delete from db_item.dbo.tbl_PlusSaleLinkItemList"
    sqlStr = sqlStr + " where plusSaleItemID=" & itemid
    dbget.Execute sqlStr, AssignedRows
    
    response.write "<script>alert('���� ��ũ " & AssignedRows & " �� ����');</script>"
    ''PlusSale Log              
    Call AddPlusItemLog(itemid, "", "", "", "", "",true)
    
    ''�÷������ϻ�ǰ ����
    sqlStr = "delete from db_item.dbo.tbl_PlusSaleRegedItem"
    sqlStr = sqlStr + " where plusSaleItemID=" & itemid
    dbget.Execute sqlStr, AssignedRows
    
    response.write "<script>alert('�÷��� ���� ��ǰ " & AssignedRows & " �� ����');</script>"

elseif (mode="PlusMainAddArr") then
    ''PlusSaleItemID
    Tot_AssignedRows = 0
    itemidarr = split(itemidarr,",")
    
    if IsArray(itemidarr) then
        for i=LBound(itemidarr) to UBound(itemidarr)
            PlusSaleLinkItemid = Trim(itemidarr(i))
            if (PlusSaleLinkItemid<>"") then
                if (PlusSaleLinkItemid<>PlusSaleItemID) then
                    CAll AddPlusLinkLog(PlusSaleItemID,PlusSaleLinkItemid,false)
                    
                    sqlStr = " insert into db_item.dbo.tbl_PlusSaleLinkItemList"
                    sqlStr = sqlStr + " (plusSaleItemid, plusSaleLinkItemid)"
                    
                    sqlStr = sqlStr + " select r.plusSaleItemid, " & PlusSaleLinkItemid
                    sqlStr = sqlStr + "   from db_item.dbo.tbl_PlusSaleRegedItem r"
                    sqlStr = sqlStr + " 	left join  db_item.dbo.tbl_PlusSaleLinkItemList s"
                    sqlStr = sqlStr + "	    on r.plusSaleItemid=s.plusSaleItemid"
                    sqlStr = sqlStr + "	    and s.plusSaleLinkItemID=" & plusSaleLinkItemid
                    sqlStr = sqlStr + " where r.plusSaleItemid=" & PlusSaleItemID
                    sqlStr = sqlStr + " and s.plusSaleLinkItemID is NULL"
                    sqlStr = sqlStr + " and s.plusSaleLinkItemID Not in ("
                    sqlStr = sqlStr + "     select plusSaleItemid from db_item.dbo.tbl_PlusSaleRegedItem where plusSaleItemid=" & plusSaleLinkItemid
                    sqlStr = sqlStr + " )"

                    dbget.Execute sqlStr, AssignedRows
                    
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
                CAll AddPlusLinkLog(PlusSaleItemID,PlusSaleLinkItemid,true)
                
                sqlStr = " delete from db_item.dbo.tbl_PlusSaleLinkItemList"
                sqlStr = sqlStr + " where plusSaleItemid=" & PlusSaleItemID
                sqlStr = sqlStr + " and plusSaleLinkItemID=" & PlusSaleLinkItemid

                dbget.Execute sqlStr, AssignedRows
                
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
                    
                    CAll AddPlusLinkLog(PlusSaleItemID,PlusSaleLinkItemid,false)
                    
                    sqlStr = " insert into db_item.dbo.tbl_PlusSaleLinkItemList"
                    sqlStr = sqlStr + " (plusSaleItemid, plusSaleLinkItemid)"
                    
                    sqlStr = sqlStr + " select r.plusSaleItemid, " & PlusSaleLinkItemid
                    sqlStr = sqlStr + "   from db_item.dbo.tbl_PlusSaleRegedItem r"
                    sqlStr = sqlStr + " 	left join  db_item.dbo.tbl_PlusSaleLinkItemList s"
                    sqlStr = sqlStr + "	    on r.plusSaleItemid=s.plusSaleItemid"
                    sqlStr = sqlStr + "	    and s.plusSaleLinkItemID=" & plusSaleLinkItemid
                    sqlStr = sqlStr + " where r.plusSaleItemid=" & PlusSaleItemID
                    sqlStr = sqlStr + " and s.plusSaleLinkItemID is NULL"

                    dbget.Execute sqlStr, AssignedRows
                    
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
                CAll AddPlusLinkLog(PlusSaleItemID,PlusSaleLinkItemid,true)
                
                sqlStr = " delete from db_item.dbo.tbl_PlusSaleLinkItemList"
                sqlStr = sqlStr + " where plusSaleItemid=" & PlusSaleItemID
                sqlStr = sqlStr + " and plusSaleLinkItemID=" & PlusSaleLinkItemid

                dbget.Execute sqlStr, AssignedRows
                
                Tot_AssignedRows = Tot_AssignedRows + AssignedRows
            end if
        next
    end if
elseif (mode="PlusSubDirectAddArr") then
    Tot_AssignedRows = 0
    itemidarr = split(itemidarr,",")            ''PlusSaleItemid

    if IsArray(itemidarr) then

        for i=LBound(itemidarr) to UBound(itemidarr)
            PlusSaleItemid = Trim(itemidarr(i))
            if (PlusSaleItemid<>"") then
                if (PlusSaleLinkItemid<>PlusSaleItemID) then
                    '// �߰����� ��ǰ ���(regPlusSale)
                    '�⺻�� ����
                    PlusSalePro  = 0        '������:0%
                    PlusSaleMargin = 0      '������:0%
                    PlusSaleMaginFlag = 4   '�������� : �ٹ����ٺδ�
                    PlusSaleStartDate = "1901-01-01"  '�Ⱓ:���
                    PlusSaleEndDate = "9999-12-31"    '�Ⱓ:���
                    cnt = 0

                    sqlStr = "select count(*) cnt from db_item.dbo.tbl_PlusSaleRegedItem where PlusSaleItemID=" & PlusSaleItemid
                    rsget.Open sqlStr,dbget,1
                        cnt = rsget(0)
                    rsget.Close

                    if cnt=0 then
                        Call AddPlusItemLog(PlusSaleItemid, PlusSalePro, PlusSaleMargin, PlusSaleMaginFlag, PlusSaleStartDate, PlusSaleEndDate,false)
                        
                        sqlStr = "insert into db_item.dbo.tbl_PlusSaleRegedItem" & VbCrlf
                        sqlStr = sqlStr + " (PlusSaleItemID,PlusSalePro,PlusSaleMargin, PlusSaleMaginFlag, PlusSaleStartDate, PlusSaleEndDate)" & VbCrlf
                        sqlStr = sqlStr + " values(" & VbCrlf
                        sqlStr = sqlStr + " " & PlusSaleItemid & VbCrlf
                        sqlStr = sqlStr + " ," & PlusSalePro & VbCrlf
                        sqlStr = sqlStr + " ," & PlusSaleMargin & VbCrlf
                        sqlStr = sqlStr + " ,'" & PlusSaleMaginFlag & "'" & VbCrlf
                        sqlStr = sqlStr + " ,'" & PlusSaleStartDate & "'" & VbCrlf
                        sqlStr = sqlStr + " ,'" & PlusSaleEndDate & "'" & VbCrlf
                        sqlStr = sqlStr + ")"
                        dbget.execute sqlStr
                    end if


                    '// �߰����� ��ǰ ����(PlusSaleAddArr)
                    CAll AddPlusLinkLog(PlusSaleItemID,PlusSaleLinkItemid,false)
                    sqlStr = " insert into db_item.dbo.tbl_PlusSaleLinkItemList"
                    sqlStr = sqlStr + " (plusSaleItemid, plusSaleLinkItemid)"
                    sqlStr = sqlStr + " select r.plusSaleItemid, " & PlusSaleLinkItemid
                    sqlStr = sqlStr + "   from db_item.dbo.tbl_PlusSaleRegedItem r"
                    sqlStr = sqlStr + " 	left join  db_item.dbo.tbl_PlusSaleLinkItemList s"
                    sqlStr = sqlStr + "	    on r.plusSaleItemid=s.plusSaleItemid"
                    sqlStr = sqlStr + "	    and s.plusSaleLinkItemID=" & plusSaleLinkItemid
                    sqlStr = sqlStr + " where r.plusSaleItemid=" & PlusSaleItemID
                    sqlStr = sqlStr + " and s.plusSaleLinkItemID is NULL"

                    dbget.Execute sqlStr, AssignedRows
                    
                    Tot_AssignedRows = Tot_AssignedRows + AssignedRows
                end if
            end if
        next
    end if
end if

%>

<% if (mode="regPlusSale") then %>
    <script language='javascript'>
    alert('��� �Ǿ����ϴ�.');
    opener.location.reload();
	location.replace('<%= refer %>');
	</script>
<% elseif mode="editPlusSale" then %>
    <script language='javascript'>
    alert('���� �Ǿ����ϴ�.');
    opener.location.reload();
	location.replace('<%= refer %>');
	</script>
<% elseif mode="delPlusSale" then %>
    <script language='javascript'>
    alert('���� �Ǿ����ϴ�.');
    opener.location.reload();
    window.close();
	//location.replace('<%= refer %>');
	</script>	
<% elseif (mode="PlusMainAddArr") then %>
    <script language='javascript'>
    alert('<%= Tot_AssignedRows %>�� �߰� �Ǿ����ϴ�.');
    //opener.location.reload();
    //window.close();
	</script>	
<% elseif (mode="PlusMainDellArr") then %>
    <script language='javascript'>
    alert('<%= Tot_AssignedRows %>�� ���� �Ǿ����ϴ�.');
    location.replace('<%= refer %>');
	</script>		
<% elseif (mode="PlusSaleDellArr") then %>
    <script language='javascript'>
    alert('<%= Tot_AssignedRows %>�� ���� �Ǿ����ϴ�.');
    location.replace('<%= refer %>');
	</script>		
<% elseif (mode="PlusSaleAddArr") or (mode="PlusSubDirectAddArr") then %>
    <script language='javascript'>
    alert('<%= Tot_AssignedRows %>�� �߰� �Ǿ����ϴ�.');
    opener.location.reload();
    location.replace('<%= refer %>');
	</script>	
<% else %>
    <script language='javascript'>
    alert('���� ���� �ʾҽ��ϴ�. - <%= mode %>');
	location.replace('<%= refer %>');
	</script>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->