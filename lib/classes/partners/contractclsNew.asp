 <%
 Class CCtrNew
 
 public FTotCnt
 public FgroupCnt
 public FPSize
 public FCPage
 public FRectDispCateCode
 public FRectMakerid  
 public FRectCompanyName 
 public FRectGroupID  
 public FRectcontracttype
 
 public FRectregDefuserid
 public FRectreguserid     
 public FRectcontractNo    
 public FRectContractState 
 public FRectreqCtrSearch  
 public FRectreqCtr        
 public FRectnotboru       
 public FRectctrType       
 public FRectselSP  
              
 public FRectnreguserid    
 public FRectncontractNo   
 public FRectnContractState
 public FRectnreqCtrSearch 
 public FRectnreqCtr       
 public FRectnnotboru      
 public FRectnctrType      
 public FRectnselSP

 public Function fnGetCtrList
    dim strSql  , strSqlAdd, strSqlAdd2 , strSqlAdd3, strSqlAdd4
 	dim iSPageNo, iEPageNo
		iSPageNo = (FPSize*(FCPage-1)) + 1
		iEPageNo = FPSize*FCPage	
	 
 
	strSqlAdd =""
	strSqlAdd2 = ""
	strSqlAdd3 = ""
	strSqlAdd4 = ""
	 if FRectDispCateCode <> "" then
	 strSqlAdd = strSqlAdd & " and  catecode= "&FRectDispCateCode
	 end if
	
	 if FRectMakerid <> "" then
	 strSqlAdd = strSqlAdd & " and  brandid= '"&FRectMakerid&"'"
	 end if
	 
	 if FRectCompanyName <> "" then
	 strSqlAdd = strSqlAdd & " and  comname like '%"&FRectCompanyName&"%'  or replace(company_no,'-','')='"&replace(FRectCompanyName,"-","")&"'"
	 end if
	 
	 if FRectGroupID <> "" then
	 strSqlAdd = strSqlAdd & " and   groupid= '"&FRectGroupID&"'"
	 end if
	 
	 if FRectregDefuserid <> "" then
	      strSqlAdd = strSqlAdd & " and groupid in ( select groupid from db_partner.dbo.tmp_partner_ctr_new where (reguserid   = '"&FRectregDefuserid&"' or regusername  = '"&FRectregDefuserid&"' ) and contracttype = 8 ) "
	 end if
	 
	 if FRectreguserid <> "" then
	     strSqlAdd = strSqlAdd & " and (  reguserid   =  '"&FRectreguserid&"' or regusername  = '"&FRectreguserid&"' ) "
	 end if    
	
    if FRectcontractNo    <> "" then
         strSqlAdd = strSqlAdd & " and  ctrno='"&FRectcontractNo&"'"
    end if
    
    if FRectContractState <> "" then 
          if FRectContractState = "M" then 
             strSqlAdd = strSqlAdd & " and c.contracttype is null  "
        else
             strSqlAdd = strSqlAdd & " and  ctrstate =  '"&FRectContractState&"'"
        end if 
    end if
 
    if FRectctrType       <> "" then
     strSqlAdd = strSqlAdd & " and  contracttype = "&FRectctrType
    end if


    if FRectselSP       <> "" then
        if FRectselSP ="on" then
          strSqlAdd = strSqlAdd & " and  sellplace = 'on'"
        else
             strSqlAdd = strSqlAdd & " and  sellplace <> 'on' and sellplace <> '' "
        end if
    end if
    
	'    
'    if FRectreqCtrSearch  <> "" then
'         strSqlAdd = strSqlAdd & " and "
'    end if
    
'    if FRectreqCtr        <> "" then
'     strSqlAdd = strSqlAdd & " and "
'    end if
'    
'    if FRectnotboru       <> "" then
'         strSqlAdd = strSqlAdd & " and "
'    end if

    if strSqlAdd <> "" then
         strSqlAdd2 = strSqlAdd2 & "  and  c.groupid in ( select groupid from db_partner.dbo.tmp_partner_ctr_new where 1=1 " &strSqlAdd &" )"
    end if
    

    if FRectnreguserid <> "" then
	     strSqlAdd2 = strSqlAdd2 & " and  (t.reguserid  = '"&FRectnreguserid&"' or nregusername =  '"&FRectnreguserid&"') "
	 end if    
	
    if FRectncontractNo    <> "" then
         strSqlAdd2 = strSqlAdd2 & " and   t.ctrno='"&FRectncontractNo&"'"
    end if
    
    if FRectnContractState <> "" then
       if FRectnContractState = "D" then 
             strSqlAdd2 = strSqlAdd2 & " and c.isusing = 0   "
       else
          strSqlAdd2 = strSqlAdd2 & " and c.isusing = 1   " 
           
           if FRectnContractState = "M" then 
              strSqlAdd2 = strSqlAdd2 & " and t.groupid is null  " 
           else
              strSqlAdd3 = strSqlAdd3 & " and  ctrstate = '"&FRectnContractState&"'"
           end if 
       end if 
    else 
        strSqlAdd2 = strSqlAdd2 & " and c.isusing = 1   "    
    end if

    
    if FRectnctrType   <> "" then
     strSqlAdd2 = strSqlAdd2 & " and  t.contracttype = "&FRectnctrType 
    end if 
     
'     if FRectnselSP       <> "" then
'     strSqlAdd3 = strSqlAdd3 & " and sellplace = '"&FRectselSP&"'"
'    end if
'    
     if strSqlAdd3 <> "" then
          strSqlAdd2 = strSqlAdd2 & "  and  c.groupid in ( select groupid from db_partner.dbo.tbl_partner_ctr_master where contracttype not in (8,9,10,16,17,18) " &strSqlAdd3 &" )"
    end if
	 
	 strSql = " select count(distinct c.groupid) as gcount, count(c.ctridx) as ctrcount "	
	 strSql = strSql & " from db_partner.dbo.tmp_partner_ctr_New as c " 
	 strSql = strSql &"  left outer join ( select groupid, makerid, sellplace, mwdiv, defaultmargin, ctrno, ctrstate, m.regdate, reguserid, senduserid, finuserid,m.contracttype , contractname as nctrname, username as nregusername  "                                                                                                                                    
     strSql = strSql &"                      from  db_partner.dbo.tbl_partner_ctr_master as m left outer join  db_partner.dbo.tbl_partner_ctr_sub as s on m.ctrKey = s.ctrKey "                                                           
     strSql = strSql &"                         left outer join  db_partner.dbo.tbl_partner_contractType as nct on m.contracttype = nct.contracttype" 
     strSql= strSql & "                         left outer join db_partner.dbo.tbl_user_tenbyten as u on m.reguserid = u.userid "
     strSql = strSql &"                     where   m.contracttype not in (8,9,10,16,17,18)   " 
     strSql = strSql &" ) as t on c.groupid = t.groupid and isNull(c.brandid,'') = t.makerid and isNull(t.sellplace,'') = isNull(c.sellplace,'') and isNull(t.mwdiv,'') = isNull(c.mwdiv,'') and isNull(t.defaultmargin,0) = isNull(c.defaultmargin,0) " 
     strSql = strSql &" where 1=1 " &   strSqlAdd2  
   '  response.write strSql
	 rsget.Open strSql,dbget,0  
        if Not rsget.Eof then
            FTotCnt = rsget("ctrcount") 
            FgroupCnt = rsget("gcount") 
        end if
	 rsget.close
	 
	 IF FTotCnt >0 THEN
	    
         strSql = " select  groupid, comname,  brandid,  sellplace "
         strSql = strSql&" , mwdiv "
         strSql = strSql&", mcnt ,wcnt,ucnt,catecode ,  cateName,contractName, ctrno,ctrstate,   regdate  , regUserid, regusername "
         strSql = strSql & " ,senduserid,sendusername,finuserid, finusername "
         strSql = strSql & " , nctrtype, nctrno, nctrstate, nregdate, nreguserid, nsenduserid, nfinuserid ,ctridx ,nctrname, nregusername, nsendusername, nfinusername, contracttype "
         strSql = strSql & " from ( "
         strSql = strSql & " select ROW_NUMBER() OVER ( order by c.groupid, c.brandid, c.contracttype  ) as rowNum, c.groupid, c.comname,  c.brandid,  c.sellplace, c.mwdiv "
         strSql = strSql & "      , c.mcnt, c.wcnt, c.ucnt, c.catecode, c.catename, c.contracttype, c.ctrno, c.ctrstate, c.regdate, c.reguserid, c.regusername, c.senduserid, c.sendusername, c.finuserid, c.finusername " 
         strSql = strSql &"       , t.contracttype as nctrtype, t.ctrno as nctrno, t.ctrstate as nctrstate , t.regdate as nregdate, t.reguserid as nreguserid, t.senduserid as nsenduserid, t.finuserid as nfinuserid "                                                                                                                                                                                                                                                             
         strSql = strSql & "      , ct.contractName , c.ctridx,nctrname, nregusername, nsendusername, nfinusername "
         strSql = strSql &"  from db_partner.dbo.tmp_partner_ctr_New as c  "   
         strSql = strSql & " left outer join db_partner.dbo.tbl_partner_contractType as ct on c.contracttype = ct.contracttype "     
         strSql = strSql &"  left outer join ( select groupid, makerid, sellplace, mwdiv, defaultmargin, ctrno, ctrstate, m.regdate, reguserid, senduserid, finuserid,m.contracttype , contractname as nctrname, u.username as nregusername , u2.username as nsendusername, u3.username as nfinusername "                                                                                                                                    
         strSql = strSql &"                      from  db_partner.dbo.tbl_partner_ctr_master as m left outer join  db_partner.dbo.tbl_partner_ctr_sub as s on m.ctrKey = s.ctrKey "                                                           
         strSql = strSql &"                         left outer join  db_partner.dbo.tbl_partner_contractType as nct on m.contracttype = nct.contracttype" 
         strSql= strSql & "                         left outer join db_partner.dbo.tbl_user_tenbyten as u on m.reguserid = u.userid "
         strSql= strSql & "                         left outer join db_partner.dbo.tbl_user_tenbyten as u2 on m.senduserid = u2.userid "
         strSql= strSql & "                         left outer join db_partner.dbo.tbl_user_tenbyten as u3 on m.finuserid = u3.userid "
         strSql = strSql &"                     where   m.contracttype not in (8,9,10,16,17,18)  " 
         strSql = strSql &" ) as t on c.groupid = t.groupid and isNull(c.brandid,'') = t.makerid and isNull(t.sellplace,'') = isNull(c.sellplace,'') and isNull(t.mwdiv,'') = isNull(c.mwdiv,'') and isNull(t.defaultmargin,0) = isNull(c.defaultmargin,0) " 
         strSql = strSql &" where 1=1  " &   strSqlAdd2 
         strSql = strSql & ") as TB "
         strSql = strSql & " WHERE TB.RowNum  Between "&iSPageNo&"  AND  "&iEPageNo  
       ' response.write strSql
         rsget.Open strSql,dbget,0  
        if Not rsget.Eof then
            fnGetCtrList = rsget.getRows()
        end if
        rsget.close
    END IF
  END function
  
End Class 
 

    public function GetContractStateColor(FCtrState)
        Select Case FCtrState
            Case 0
                : GetContractStateColor = "#000000"
            Case 1
                : GetContractStateColor = "#44BB44"
            Case 3
                : GetContractStateColor = "#7777FF"
            Case 7
                : GetContractStateColor = "#FF7777"
            Case -1
                : GetContractStateColor = "#AAAAAA"
           Case else
                : GetContractStateColor = "#000000"
        end Select
    end function

    public function GetContractStateName(FCtrState)
        dim buf
        Select Case FCtrState
            Case 0
                : buf = "수정중"
            Case 1
                : buf = "계약오픈"
            Case 3
                : buf = "업체확인"
            Case 7
                : buf = "계약완료"
            Case -1
                : buf = "삭제"
            Case else
                : buf = FCtrState
        end Select

        GetContractStateName = "<font color='"&GetContractStateColor(FCtrState)&"'>"&buf&"</font>"
    end function
    
    
'' 매입구분
public function fnMaeipdivName(imaeipdiv)
    if isNULL(imaeipdiv) then Exit function

    select case imaeipdiv
        CASE "M" : fnMaeipdivName="매입"
        CASE "W" : fnMaeipdivName="위탁"
        CASE "U" : fnMaeipdivName="업체"

        CASE "B011" : fnMaeipdivName="위탁판매"
        CASE "B012" : fnMaeipdivName="업체위탁"
        CASE "B013" : fnMaeipdivName="출고위탁"
        CASE "B021" : fnMaeipdivName="오프매입"
        CASE "B022" : fnMaeipdivName="매장매입"
        CASE "B023" : fnMaeipdivName="가맹점매입"
        CASE "B031" : fnMaeipdivName="출고매입"
        CASE "B032" : fnMaeipdivName="센터매입"
        CASE ELSE : fnMaeipdivName=imaeipdiv
    end select
end function
 %>