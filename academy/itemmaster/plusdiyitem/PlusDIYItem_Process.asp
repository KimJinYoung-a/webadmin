<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  다이샾 플러스 상품
' History : 2010.11.09 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim itemid, PlusSalePro, PlusSaleMaginFlag, PlusSaleMargin , i,cnt ,sqlStr, AssignedRows ,Tot_AssignedRows
dim termsGubun, PlusSaleStartDate, PlusSaleEndDate ,itemidarr, PlusSaleItemID, PlusSaleLinkItemid, chkitem
dim mode , tmp
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

'/플러스 다이 상품 & 추가 구성 등록
if (mode="PlusMainAddArr") then
    Tot_AssignedRows = 0
    itemidarr = split(itemidarr,",")
	
	dbACADEMYget.beginTrans
	    
    if IsArray(itemidarr) then
        for i=LBound(itemidarr) to UBound(itemidarr)
            PlusSaleLinkItemid = Trim(itemidarr(i))
            if (PlusSaleLinkItemid<>"") then
                if (PlusSaleLinkItemid<>PlusSaleItemID) then

					'/플러스 다이 상품 저장                    
                    sqlStr = ""
                    sqlStr = "insert into db_academy.dbo.tbl_diy_PlusSaleLinkItem"
                    sqlStr = sqlStr + " (plusSaleItemid, plusSaleLinkItemid)"                    
					sqlStr = sqlStr + " 	select i.itemid , "& PlusSaleItemID &"" + vbcrlf
					sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_item i" + vbcrlf
					'sqlStr = sqlStr + " 	'left join db_academy.dbo.tbl_diy_PlusSaleRegedItem r" + vbcrlf
					'sqlStr = sqlStr + " 	on i.itemid = r.plusSaleItemid" + vbcrlf
					sqlStr = sqlStr + " 	left join (" + vbcrlf
					sqlStr = sqlStr + " 		select plusSaleLinkItemID" + vbcrlf
					sqlStr = sqlStr + " 		from db_academy.dbo.tbl_diy_PlusSaleLinkItem" + vbcrlf
					sqlStr = sqlStr + " 		group by plusSaleLinkItemID" + vbcrlf
					sqlStr = sqlStr + " 	) as T" + vbcrlf
					sqlStr = sqlStr + " 	on i.itemid = t.plusSaleLinkItemID" + vbcrlf
					sqlStr = sqlStr + " 	where itemid = "&plusSaleLinkItemid&"" + vbcrlf					
					sqlStr = sqlStr + " 	and t.plusSaleLinkItemID is null" + vbcrlf		'플러스 상품 제낌
					'sqlStr = sqlStr + " 	and r.plusSaleItemid is null" + vbcrlf		'추가구성제낌
					sqlStr = sqlStr + " 	and i.isusing='Y'" + vbcrlf		'사용하는 상품만 
					sqlStr = sqlStr + "		and saleYn='N'" + vbcrlf		'할인 상품은 제낀다

					'response.write sqlStr &"<br>"
                    dbACADEMYget.Execute sqlStr, AssignedRows

                    Tot_AssignedRows = Tot_AssignedRows + AssignedRows
					
					'/추가 구성 상품 저장			
					sqlStr = ""
					sqlStr = "insert into db_academy.dbo.tbl_diy_PlusSaleRegedItem"
					sqlStr = sqlStr & " (plusSaleItemID,plusSalePro,plusSaleMargin,plusSaleMaginFlag,plusSaleStartDate,plusSaleEndDate,regdate)"
					sqlStr = sqlStr & " 	select" 
					sqlStr = sqlStr & " 	"&PlusSaleLinkItemid&",0, 100-Floor(buycash*1/sellcash*1*100*100)/100 , 1,'1901-01-01 00:00:00','9999-12-31 23:59:59',getdate()"
					sqlStr = sqlStr & " 	from db_academy.dbo.tbl_diy_item i"
					sqlStr = sqlStr & " 	left join db_academy.dbo.tbl_diy_PlusSaleRegedItem r"
					sqlStr = sqlStr & " 	on i.itemid = r.plusSaleItemid"
			        sqlStr = sqlStr & " 	left join ("
			        sqlStr = sqlStr & " 		select plusSaleLinkItemID from db_academy.dbo.tbl_diy_PlusSaleLinkItem group by plusSaleLinkItemID"
			        sqlStr = sqlStr & "  	) as T"
			        sqlStr = sqlStr & "  	on i.itemid = t.plusSaleLinkItemID"
					sqlStr = sqlStr & " 	where itemid = "&PlusSaleLinkItemid&""
					sqlStr = sqlStr + " 	and itemdiv <> 20" + vbcrlf		'추가전용상품 제낌					
					sqlStr = sqlStr & " 	and t.plusSaleLinkItemID is null"	'플러스 등록 상품은 제낀다 
					sqlStr = sqlStr & " 	and r.plusSaleItemid is null"	'추가구성제낌				
					sqlStr = sqlStr + " 	and i.isusing='Y'" + vbcrlf		'사용하는 상품만
					sqlStr = sqlStr & " 	and saleYn='N'"		'할인 상품은 제낀다
										
					'response.write sqlStr &"<br>"
                    dbACADEMYget.Execute sqlStr, AssignedRows
                    
                end if
            end if
        next
    end if

    If Err.Number = 0 Then
        dbACADEMYget.CommitTrans
        response.write "<script language='javascript'>alert('"& Tot_AssignedRows &"건 추가 되었습니다');</script>"
        'response.write "<script language='javascript'>alert('OK');</script>"
        response.write "<script language='javascript'>opener.location.reload();</script>"
        response.write "<script language='javascript'>opener.opener.location.reload();</script>"        
        response.write "<script language='javascript'>window.close();</script>"
        dbACADEMYget.close()	: response.End
    Else
        dbACADEMYget.RollBackTrans
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다')</script>"
        response.write "<script>history.back()</script>"
        dbACADEMYget.close()	: response.End
    End If

'/추가 구성 삭제
elseif (mode="PlusSaleDellArr") then
    Tot_AssignedRows = 0
    chkitem = split(chkitem,",")
    
    dbACADEMYget.beginTrans
    
    if IsArray(chkitem) then
        for i=LBound(chkitem) to UBound(chkitem)
            PlusSaleItemid = Trim(chkitem(i))
            if (PlusSaleItemid<>"") then
            	
            	sqlStr = ""
                sqlStr = " delete from db_academy.dbo.tbl_diy_PlusSaleLinkItem"
                sqlStr = sqlStr + " where plusSaleItemid=" & PlusSaleItemID
                sqlStr = sqlStr + " and plusSaleLinkItemID=" & PlusSaleLinkItemid
				
				'response.write sqlStr &"<br>"			
                dbACADEMYget.Execute sqlStr, AssignedRows				

				Tot_AssignedRows = Tot_AssignedRows + AssignedRows
				
				tmp = ""
				sqlStr = ""
            	sqlStr = sqlStr + " select count(plusSaleItemid) as cnt"
            	sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_PlusSaleLinkItem"
            	sqlStr = sqlStr + " where plusSaleItemid=" & PlusSaleItemID				
				
				'response.write sqlStr &"<br>"
		        rsACADEMYget.Open sqlStr,dbACADEMYget,1
		            tmp = rsACADEMYget("cnt")
		        rsACADEMYget.Close
            	
            	if tmp = "0" then
	            	sqlStr = ""
	                sqlStr = "delete from db_academy.dbo.tbl_diy_PlusSaleRegedItem"
	                sqlStr = sqlStr + " where plusSaleItemid=" & PlusSaleItemID                
	 				
	 				'response.write sqlStr &"<br>"			
	                dbACADEMYget.Execute sqlStr, AssignedRows
                end if
                               
            end if
        next
    end if

    If Err.Number = 0 Then
        dbACADEMYget.CommitTrans
	    response.write "<script language='javascript'>"
	    response.write "	alert('"& Tot_AssignedRows &"건 삭제 되었습니다.');"
	    response.write "	location.replace('"& refer &"');"
	    response.write "	opener.location.reload();"
		response.write "</script>"
        dbACADEMYget.close()	: response.End
    Else
        dbACADEMYget.RollBackTrans
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다')</script>"
        response.write "<script>history.back()</script>"
        dbACADEMYget.close()	: response.End
    End If
end if
%>
	
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->