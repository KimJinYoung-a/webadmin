<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   
'###########################################################
' Description : 브랜드 계약 관리
' Hieditor : 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim sMode
dim arrgroupid, intLoop, groupid, arrList, intX
dim strSql, ctrKey
dim sregUserid
dim contractContents 
dim  comname,company_no,ceoname,compaddr,jungsandate, nowdate,enddate ,nctrKey,ngroupid,ctrNo
dim ctrContents(3),ctrName(3),onoffgubun(3),subType(3)
dim sellplace , mwdivnm
dim mwdiv, defaultmargin, sellplacename,regDefUserid
 dim strParm
 dim makerid, regUserid,dispCate,arect,contractNo,uniqGroupID, reqCtrSearch,grpType,ctrType,crect,ContractState
 dim nregUserid, ncontractNo ,nreqCtrSearch, nctrType,nContractState,nreqCtr,nnotboru  ,iCurrpage
 dim arrC,intC,bufStr, i , mailfrom 
   
 sMode = requestCheckvar(request("hidM"),1)
 sregUserid = session("ssBctID")
 arrgroupid =  split(request("chkid"),",") 
 nowdate = date()
 enddate = "2016-06-30"
 
    makerid = requestCheckVar(request("makerid"),32) 
	dispCate = requestCheckvar(request("dispCate"),10)
	arect   = requestCheckVar(request("arect"),32)
 	crect   = requestCheckVar(request("crect"),32) 
 	
 	regDefUserid    = requestCheckVar(request("rDU"),32)
 	reguserid       = requestCheckVar(request("rU"),32)
 	contractNo      = requestCheckVar(request("contractNo"),20)
 	ContractState   = requestCheckVar(request("ContractState"),10) 
    ctrType         = requestCheckvar(request("ctrType"),10)
	
	nreguserid      = requestCheckVar(request("nrU"),32)
 	ncontractNo     = requestCheckVar(request("ncontractNo"),20)
 	nContractState  = requestCheckVar(request("nContractState"),10) 
    nctrType        = requestCheckvar(request("nctrType"),10)
    iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
    
    strParm = "makerid="&makerid&"&dispcate="&dispcate&"&arect="&arect&"&crect="&crect&"&rU="&reguserid&"&contractNo="&contractNo&"&ContractState="&ContractState&"&ctrType="&ctrType&"&nrU="&nreguserid&"&ncontractNo="&ncontractNo&"&nContractState="&nContractState&"&nctrType="&nctrType&"&iC="&iCurrpage&"&arrgid="&request("chkid")&"&rDU="&regDefUserid
    
SELECT case sMode
case  "I"
 '--거래기본계약서: 그룹코드별, 매입구분- 업체, 위탁 
 '--거래기본계약부속합의서 : 브랜드별, 판매처별, 매입구분별 
 '--직매입계약서 : 그룹코드별, 매입구분 - 매입
 '--직매입계약부속합의서: X
   
   '-- 계약서내용
   strSql = "select contractType, contractContents, contractName ,onoffgubun, subType" &vbcrlf
   strSql = strSql & " from db_partner.dbo.tbl_partner_contractType" &vbcrlf
   strSql = strSql & " where contractType not in (8,9,10,16,17,18) order by contractType "   
   rsget.Open strSql,dbget,1
   if Not rsget.Eof then
     arrC = rsget.getRows()
   end if
   rsget.close 
   
   if isArray(arrC) then 
     For intC = 0 To uBound(arrC,2)
     if arrC(0,intC) = 11 then
         ctrContents(0) = db2Html(arrC(1,intC))
         ctrName(0) = db2Html(arrC(2,intC))
         onoffgubun(0) = arrC(3,intC)
         subType(0)    = arrC(4,intC)
     elseif arrC(0,intC) = 12 then
         ctrContents(1) = db2Html(arrC(1,intC))
         ctrName(1) = db2Html(arrC(2,intC))
         onoffgubun(1) = arrC(3,intC)
         subType(1)    = arrC(4,intC)
     elseif arrC(0,intC) = 13 then
         ctrContents(2) = db2Html(arrC(1,intC))
         ctrName(2)= db2Html(arrC(2,intC))
         onoffgubun(2) = arrC(3,intC)
         subType(2)    = arrC(4,intC)
     end if
     Next
     
   end if
     
        
 For intLoop = 0 To UBound(arrgroupid)
 
      groupid =   trim(arrgroupid(intLoop))
      
      strSql = " select n.groupid, n.comname, n.company_no, n.ceoname, n.compaddr, n.jungsandate, m.ctrKey "
      strSql = strSql & " from db_partner.dbo.tmp_partner_ctr_new as n "
      strSql = strSql & "  left outer join  db_partner.dbo.tbl_partner_ctr_master as m on n.groupid = m.groupid and m.contracttype= " & DEFAULT_CONTRACTTYPE & " and m.ctrstate >=0 "
      strSql = strSql & " where n.isusing =1 and n.groupid = '"&groupid&"' and (mwdiv in ('U','W','B012','B013','B031' ) or wcnt >0 or ucnt > 0)" 
      strSql = strSql & " group by n.groupid, comname, company_no, ceoname, compaddr, jungsandate, m.ctrKey "    
      rsget.Open strSql, dbget 
     IF not (rsget.EOF or rsget.BOF) THEN
        ngroupid  = rsget("groupid")
        comname  = rsget("comname")
        company_no  = rsget("company_no")
        ceoname  = rsget("ceoname")
        compaddr  = rsget("compaddr")
        jungsandate = rsget("jungsandate")
        nctrKey = rsget("ctrKey") 
     end if
       rsget.close
       
      '--거래기본계약서
      if not isNull(ngroupid) and isNull(nctrKey)  then
        dbget.beginTrans
          strSql = " insert into db_partner.dbo.tbl_partner_ctr_master ([contractType] ,[groupid] , makerid, ctrno,[ctrState] , [regUserID] ) "
          strSql = strSql & " values (11,'"&groupid&"','','',0,'"&sregUserid&"')" 
          dbget.Execute  strSql
          
          'strSql = "select SCOPE_IDENTITY() From [db_partner].[dbo].[tbl_partner_ctr_master] "		'/사용금지.전체 라인 몽땅 뿌려짐. '/2016.06.02 한용민
          strSql = "select SCOPE_IDENTITY()"

    	  rsget.Open strSql, dbget, 0
    	  ctrKey = rsget(0)
    	  rsget.Close 

           strSql = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
           strSql =  strSql&" (ctrKey,detailKey,detailValue)"
           strSql =  strSql&" select "&ctrKey&",detailKey,"
           strSql =  strSql&" (CASE WHEN detailKey='$$A_CEONAME$$' THEN '최은희'"
           strSql =  strSql&" 	  WHEN detailKey='$$A_COMPANY_ADDR$$' THEN '서울 종로구 대학로 12길 31, 2층'"
           strSql =  strSql&" 	  WHEN detailKey='$$A_COMPANY_NO$$' THEN '211-87-00620'"
           strSql =  strSql&" 	  WHEN detailKey='$$A_UPCHENAME$$' THEN '(주)텐바이텐'"
           strSql =  strSql&" 	  WHEN detailKey='$$B_CEONAME$$' THEN '"&ceoname&"'"
           strSql =  strSql&"     WHEN detailKey='$$B_COMPANY_ADDR$$' THEN '"&compaddr&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$B_COMPANY_NO$$' THEN '"&company_no&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$B_UPCHENAME$$' THEN '"&html2db(comname)&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$CONTRACT_DATE$$' THEN '"&nowdate&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$DEFAULT_JUNGSANDATE$$' THEN '"&jungsandate&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$ENDDATE$$' THEN '"&enddate&"'"
           strSql =  strSql&" 	  ELSE '' END)"
           strSql =  strSql&" from db_partner.dbo.tbl_partner_contractDetailType"
           strSql =  strSql&" where contractType=" & DEFAULT_CONTRACTTYPE & " " 
          dbget.Execute  strSql
                
          contractContents =  ctrContents(0)
          contractContents = Replace(contractContents,"$$A_CEONAME$$","최은희") 
          contractContents = Replace(contractContents,"$$A_COMPANY_ADDR$$","서울 종로구 대학로 12길 31, 2층") 
          contractContents = Replace(contractContents,"$$A_COMPANY_NO$$","211-87-00620") 
          contractContents = Replace(contractContents,"$$A_UPCHENAME$$","(주)텐바이텐") 
          contractContents = Replace(contractContents,"$$B_CEONAME$$", ceoname ) 
          contractContents = Replace(contractContents,"$$B_COMPANY_ADDR$$",compaddr) 
          contractContents = Replace(contractContents,"$$B_COMPANY_NO$$",company_no) 
          contractContents = Replace(contractContents,"$$B_UPCHENAME$$",comname) 
          contractContents = Replace(contractContents,"$$CONTRACT_DATE$$",Left(nowdate,4) & "년 " & Mid(nowdate,6,2) & "월 " & Mid(nowdate,9,2) & "일") 
          contractContents = Replace(contractContents,"$$DEFAULT_JUNGSANDATE$$",jungsandate) 
          contractContents = Replace(contractContents,"$$ENDDATE$$",  Left(enddate,4) & "년 " & Mid(enddate,6,2) & "월 " & Mid(enddate,9,2) & "일")
         
            ctrNo = nowdate
            ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
            ctrNo = ctrNo & "-" & Format00(2,11) & "-" & ctrKey
    
             strSql = " update db_partner.dbo.tbl_partner_ctr_master"
             strSql =  strSql & " set contractContents='" & Newhtml2db(contractContents) & "'"
             strSql =  strSql & " ,ctrNo='" & ctrNo & "'"
             strSql =  strSql & ", enddate='"&enddate&"'"
             strSql =  strSql & " where ctrKey=" & ctrKey &" and ctrstate>=0"
            dbget.Execute  strSql
             
       
        '--거래기본계약부속합의서  
         strSql = " insert into db_partner.dbo.tbl_partner_ctr_master ([contractType] ,[groupid],makerid , ctrno, [ctrState] , [regUserID] ) "
         strSql = strSql & "  select 12, groupid,brandid,'',0 ,'"&sregUserid&"'"
         strSql = strSql & " from  db_partner.dbo.tmp_partner_ctr_new as n "
         strSql = strSql & " where n.isusing =1 and groupid = '"&groupid&"' and (mwdiv in ('U','W','B012','B013','B031')  or wcnt >0 or ucnt > 0) and brandid is not null " 
         strSql = strSql & " group by groupid, comname, company_no, ceoname, compaddr, jungsandate,mwdiv, sellplace, brandid " 
         dbget.Execute  strSql
         
         arrList= ""
         strSql = " select ctrKey,makerid from db_partner.dbo.tbl_partner_ctr_master where contracttype = " & ADD_CONTRACTTYPE & " and groupid ='"&groupid&"' and ctrstate>=0" 
       
         rsget.Open strSql, dbget 
           IF not (rsget.EOF or rsget.BOF) THEN 
             arrList =  rsget.getRows()
           end if 
         rsget.close 
         
           if isarray(arrList) then  
            for intX = 0 To UBound(arrList,2) 
                strSql = " select top 1 n.sellplace, n.mwdiv, isNull(n.defaultmargin,0) as defaultmargin "
                strSql = strSql &" ,(CASE WHEN sellplace='ON' then '온라인'"
	            strSql = strSql &" WHEN sellplace<>'ON' and isNULL(u.shopname,'')<>'' THEN u.shopname + ' 매장' ELSE sellplace END) as sellplacename"
	            strSql = strSql&" ,(CASE WHEN sellplace='ON' and mwdiv='M' THEN '텐바이텐배송' "
        	    strSql = strSql&"   WHEN sellplace='ON' and mwdiv='W' THEN '텐바이텐배송' "
        	    strSql = strSql&"   WHEN sellplace='ON' and mwdiv='U' THEN '협력사배송' "
        	    strSql = strSql&"   WHEN sellplace<>'ON' and mwdiv= 'B031' THEN '텐바이텐배송'"
        	    strSql = strSql&"   WHEN sellplace<>'ON' and mwdiv= 'B013' THEN '텐바이텐배송'"
        	    strSql = strSql&"   WHEN sellplace<>'ON' and mwdiv= 'B012' THEN '오프라인위탁'" 
        	    strSql = strSql&"   ELSE mwdiv END) as mwdivName"
                strSql = strSql & " from db_partner.dbo.tmp_partner_ctr_new as n "
                strSql = strSql &"    left join db_shop.dbo.tbl_shop_user u  on n.sellplace=u.userid"
                strSql = strSql &"   where brandid ='"&arrList(1,intX)&"' and n.mwdiv <> 'M' "
                strSql = strSql & "    and (n.isSubIn = 0 or n.isSubIn is null) "
                'strSql = strSql & "   and  ("
                'strSql = strSql & "     sellplace not in ( "
                'strSql = strSql & "            select s.sellplace from  db_partner.dbo.tbl_partner_ctr_master as m "
        	    'strSql = strSql & "                            inner join  db_partner.dbo.tbl_partner_ctr_sub as s on m.ctrkey = s.ctrkey "
        	    'strSql = strSql & "                            where m.makerid = '"&arrList(1,intX)&"'  and m.contracttype = 12 and m.ctrstate>=0
        	    'strSql = strSql & "                      ) or "
        	    'strSql = strSql & "    mwdiv not in ( "
        	    'strSql = strSql & "            select s.mwdiv from  db_partner.dbo.tbl_partner_ctr_master as m "
        	   ' strSql = strSql & "                        inner join  db_partner.dbo.tbl_partner_ctr_sub as s on m.ctrkey = s.ctrkey "
        	   ' strSql = strSql & "                        where n.isusing =1 and m.makerid = '"&arrList(1,intX)&"'  and m.contracttype = 12 and m.ctrstate>=0"
        	   ' strSql = strSql & "                )"
        	    'strSql = strSql & " ) "
        	    strSql = strSql &" order by n.regdate desc " 
        	   
                 rsget.Open strSql, dbget 
                 IF not (rsget.EOF or rsget.BOF) THEN
                    sellplace = rsget("sellplace")
                    mwdiv = rsget("mwdiv")
                    defaultmargin = rsget("defaultmargin")
                    sellplaceName = rsget("sellplacename")
                    mwdivnm      =rsget("mwdivName") 
                end if
                rsget.close
                ctrKey = arrList(0,intX)  
                
                strSql =  "update  db_partner.dbo.tmp_partner_ctr_new set isSubIn = 1 where groupid ='"&groupid&"' and brandid ='"&arrList(1,intX)&"' and sellplace ='"&sellplace&"' and mwdiv='"&mwdiv&"'"
                dbget.Execute  strSql
                
                strSql = " insert into db_partner.dbo.tbl_partner_ctr_Sub"
                strSql = strSql & " (ctrKey,sellplace,mwdiv,defaultmargin)" 
                strSql = strSql & " values ( '" &ctrKey & "', '"&sellplace&"','"&mwdiv&"','"&defaultmargin&"')"  
                dbget.Execute  strSql
             
                bufStr = ""
                bufStr="<thead><tr><th>브랜드ID</th><th>판매채널</th><th>배송방식</th><th>수수료율</th><th>비고</th></tr>" 
                bufStr = bufStr & "<tr>"
                bufStr = bufStr & "<td align='center'>"&arrList(1,intX)&"</td>"
                bufStr = bufStr & "<td align='center'>"&sellplaceName&"</td>"
                bufStr = bufStr & "<td align='center'>"&mwdivnm&"</td>" 
                bufStr = bufStr & "<td align='center'>"&CLNG(defaultmargin*100)/100&" %</td>"  
                bufStr = bufStr & "<td align='center' width='50'>&nbsp;</td>"  
                bufStr = bufStr & "</tr>"
                    
               strSql = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
               strSql =  strSql&" (ctrKey,detailKey,detailValue)"
               strSql =  strSql&" select "&arrList(0,intX)&",detailKey,"
               strSql =  strSql&" (CASE WHEN detailKey='$$A_CEONAME$$' THEN '최은희'"
               strSql =  strSql&" 	  WHEN detailKey='$$A_COMPANY_ADDR$$' THEN '서울 종로구 대학로 12길 31, 2층'"
               strSql =  strSql&" 	  WHEN detailKey='$$A_COMPANY_NO$$' THEN '211-87-00620'"
               strSql =  strSql&" 	  WHEN detailKey='$$A_UPCHENAME$$' THEN '(주)텐바이텐'"
               strSql =  strSql&" 	  WHEN detailKey='$$B_CEONAME$$' THEN '"&ceoname&"'"
               strSql =  strSql&"     WHEN detailKey='$$B_COMPANY_ADDR$$' THEN '"&compaddr&"'"
               strSql =  strSql&" 	  WHEN detailKey='$$B_COMPANY_NO$$' THEN '"&company_no&"'"
               strSql =  strSql&" 	  WHEN detailKey='$$B_UPCHENAME$$' THEN '"&html2db(comname)&"'"
               strSql =  strSql&" 	  WHEN detailKey='$$CONTRACT_DATE$$' THEN '"&nowdate&"'"
               strSql =  strSql&" 	  WHEN detailKey='$$CONTRACT_CONTS$$' THEN '"&Newhtml2db(bufStr)&"'" 
               strSql =  strSql&" 	  ELSE '' END)"
               strSql =  strSql&" from db_partner.dbo.tbl_partner_contractDetailType"
               strSql =  strSql&" where contractType= " & ADD_CONTRACTTYPE & ""  
              dbget.Execute  strSql
                    
              contractContents =  ctrContents(1)
              contractContents = Replace(contractContents,"$$A_CEONAME$$","최은희") 
              contractContents = Replace(contractContents,"$$A_COMPANY_ADDR$$","서울 종로구 대학로 12길 31, 2층") 
              contractContents = Replace(contractContents,"$$A_COMPANY_NO$$","211-87-00620") 
              contractContents = Replace(contractContents,"$$A_UPCHENAME$$","(주)텐바이텐") 
              contractContents = Replace(contractContents,"$$B_CEONAME$$", ceoname ) 
              contractContents = Replace(contractContents,"$$B_COMPANY_ADDR$$",compaddr) 
              contractContents = Replace(contractContents,"$$B_COMPANY_NO$$",company_no) 
              contractContents = Replace(contractContents,"$$B_UPCHENAME$$",comname) 
              contractContents = Replace(contractContents,"$$CONTRACT_DATE$$",Left(nowdate,4) & "년 " & Mid(nowdate,6,2) & "월 " & Mid(nowdate,9,2) & "일") 
              contractContents = Replace(contractContents,"$$CONTRACT_CONTS$$",bufStr) 
               
                ctrNo = nowdate
                ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
                ctrNo = ctrNo & "-" & Format00(2,11) & "-" & arrList(0,intX)
        
                 strSql = " update db_partner.dbo.tbl_partner_ctr_master"
                 strSql =  strSql & " set contractContents='" & Newhtml2db(contractContents) & "'"
                 strSql =  strSql & " ,ctrNo='" & ctrNo & "'"
                 strSql =  strSql & ", enddate='"&enddate&"'"
                 strSql =  strSql & " where ctrKey=" & arrList(0,intX) 
                dbget.Execute  strSql 
                
             Next
           end if 
           
            IF Err.Number = 0 THEN
		   dbget.CommitTrans
		 else
		    response.write Err.Description
		    dbget.RollBackTrans
		     Call Alert_move("처리중 에러가 발생했습니다. ","newctrList.asp")
            response.end
		 end if
	  else
	    response.write "er"	 
      end if  
      
      ngroupid = ""
        '--직매입계약서
      strSql = " select n.groupid, n.comname, n.company_no, n.ceoname, n.compaddr, n.jungsandate, m.ctrKey "
      strSql = strSql & " from db_partner.dbo.tmp_partner_ctr_new as n "
      strSql = strSql & "  left outer join  db_partner.dbo.tbl_partner_ctr_master as m on n.groupid = m.groupid and m.contracttype = " & DEFAULT_CONTRACTTYPE_M & " "
      strSql = strSql & " where n.isusing =1 and n.groupid = '"&groupid&"' and (mwdiv in ('M' ) or mcnt>0) " 
      strSql = strSql & " group by n.groupid, comname, company_no, ceoname, compaddr, jungsandate, m.ctrKey  " 
      rsget.Open strSql, dbget 
     IF not (rsget.EOF or rsget.BOF) THEN
        ngroupid  = rsget("groupid")
        comname  = rsget("comname")
        company_no  = rsget("company_no")
        ceoname  = rsget("ceoname")
        compaddr  = rsget("compaddr") 
        nctrKey = rsget("ctrKey") 
        jungsandate = rsget("jungsandate")
     end if
      rsget.close
 

      if not (isNull(ngroupid) or ngroupid = "") and isNull(nctrKey)  then
       
        dbget.beginTrans
          strSql = " insert into db_partner.dbo.tbl_partner_ctr_master ([contractType] ,[groupid] , makerid, ctrno,[ctrState] , [regUserID] ) "
          strSql = strSql & " values (13,'"&groupid&"','','',0,'"&sregUserid&"')"
          dbget.Execute  strSql
          
          'strSql = "select SCOPE_IDENTITY() From [db_partner].[dbo].[tbl_partner_ctr_master] "		'/사용금지.전체 라인 몽땅 뿌려짐. '/2016.06.02 한용민
          strSql = "select SCOPE_IDENTITY()"

    	  rsget.Open strSql, dbget, 0
    	  ctrKey = rsget(0)
    	  rsget.Close 
          
           strSql = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
           strSql =  strSql&" (ctrKey,detailKey,detailValue)"
           strSql =  strSql&" select "&ctrKey&",detailKey,"
           strSql =  strSql&" (CASE WHEN detailKey='$$A_CEONAME$$' THEN '최은희'"
           strSql =  strSql&" 	  WHEN detailKey='$$A_COMPANY_ADDR$$' THEN '서울 종로구 대학로 12길 31, 2층'"
           strSql =  strSql&" 	  WHEN detailKey='$$A_COMPANY_NO$$' THEN '211-87-00620'"
           strSql =  strSql&" 	  WHEN detailKey='$$A_UPCHENAME$$' THEN '(주)텐바이텐'"
           strSql =  strSql&" 	  WHEN detailKey='$$B_CEONAME$$' THEN '"&ceoname&"'"
           strSql =  strSql&"     WHEN detailKey='$$B_COMPANY_ADDR$$' THEN '"&compaddr&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$B_COMPANY_NO$$' THEN '"&company_no&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$B_UPCHENAME$$' THEN '"&html2db(comname)&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$DEFAULT_JUNGSANDATE$$' THEN '"&jungsandate&"'"
           strSql =  strSql&" 	  WHEN detailKey='$$CONTRACT_DATE$$' THEN '"&nowdate&"'" 
           strSql =  strSql&" 	  WHEN detailKey='$$ENDDATE$$' THEN '"&enddate&"'"
           strSql =  strSql&" 	  ELSE '' END)"
           strSql =  strSql&" from db_partner.dbo.tbl_partner_contractDetailType"
           strSql =  strSql&" where contractType=" & DEFAULT_CONTRACTTYPE_M & " " 
          dbget.Execute  strSql
                
          contractContents = ctrContents(2)
          contractContents = Replace(contractContents,"$$A_CEONAME$$","최은희") 
          contractContents = Replace(contractContents,"$$A_COMPANY_ADDR$$","서울 종로구 대학로 12길 31, 2층") 
          contractContents = Replace(contractContents,"$$A_COMPANY_NO$$","211-87-00620") 
          contractContents = Replace(contractContents,"$$A_UPCHENAME$$","(주)텐바이텐") 
          contractContents = Replace(contractContents,"$$B_CEONAME$$", ceoname ) 
          contractContents = Replace(contractContents,"$$B_COMPANY_ADDR$$",compaddr) 
          contractContents = Replace(contractContents,"$$B_COMPANY_NO$$",company_no) 
          contractContents = Replace(contractContents,"$$B_UPCHENAME$$",comname) 
          contractContents = Replace(contractContents,"$$DEFAULT_JUNGSANDATE$$",jungsandate) 
          contractContents = Replace(contractContents,"$$CONTRACT_DATE$$",Left(nowdate,4) & "년 " & Mid(nowdate,6,2) & "월 " & Mid(nowdate,9,2) & "일")  
          contractContents = Replace(contractContents,"$$ENDDATE$$",  Left(enddate,4) & "년 " & Mid(enddate,6,2) & "월 " & Mid(enddate,9,2) & "일")
         
            ctrNo = nowdate
            ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
            ctrNo = ctrNo & "-" & Format00(2,11) & "-" & ctrKey
    
             strSql = " update db_partner.dbo.tbl_partner_ctr_master"
             strSql =  strSql & " set contractContents='" & Newhtml2db(contractContents) & "'"
             strSql =  strSql & " ,ctrNo='" & ctrNo & "'"
             strSql =  strSql & ", enddate='"&enddate&"'"
             strSql =  strSql & " where ctrKey=" & ctrKey 
            dbget.Execute  strSql
            
         IF Err.Number = 0 THEN
		   dbget.CommitTrans
		 else
		    response.write Err.Description
		    dbget.RollBackTrans
		     Call Alert_move("처리중 에러가 발생했습니다. ","newctrList.asp")
            response.end
		 end if 
    else
	    response.write "er2"
        end if 
 Next  
 
  Call Alert_move("성공","newctrList.asp?"&strParm )
response.end

CASE "M" '오픈 & 메일발송
   dim cmakerid, cscmmwdiv, cscmmargin, csellplace, csellplacename, cctrmwdiv, cctrmargin, cmjmaeipdiv, cmjdefaultmargin, cuseitemcnt, cuseitemmargin, csellitemcnt, csellitemmargin
   dim isDisabledMWMargin,dsbleCnt,errgroupid
   dim mnghp, mngEmail
   dim   iCtrKeyArr
   dim intCK,arrCtrKey
   dim sqlStr,ocontract,oMdInfoList, mailtitle, mailcontent
   
    For intLoop = 0 To UBound(arrgroupid) 
      groupid =   trim(arrgroupid(intLoop))
      
    
          dsbleCnt = 0
          '-- 마진체크
           sqlStr = "db_partner.[dbo].[sp_Ten_partner_AddContract_CheckList] ('"&groupid&"')"
            rsget.CursorLocation = adUseClient
    		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
           IF Not (rsget.EOF OR rsget.BOF) THEN
    		do until rsget.eof
    		  cmakerid      = rsget("Makerid")             
              cscmmwdiv     = rsget("scmmwdiv")            
              cscmmargin    =  rsget("scmmargin")           
                                                 
              csellplace     = rsget("sellplace")           
              csellplacename = rsget("sellplaceName")       
                                                 
              cctrmwdiv      = rsget("ctrmwdiv")            
              cctrmargin     = rsget("ctrmargin")           
                        
              cmjmaeipdiv     = rsget("Mjmaeipdiv")          
              cmjdefaultmargin = rsget("Mjdefaultmargin")     
                                                 
              cuseitemcnt       =rsget("useitemCnt")          
              cuseitemmargin    =   rsget("useitemmargin")       
                                                 
              csellitemcnt     = rsget("sellitemCnt")         
              csellitemmargin   =    rsget("sellitemmargin")      
    '
    '        if (csellplace="ON") then
    '            if not isNULL(cscmmwdiv) then  
    '                if (cctrmargin<>cscmmargin) then
    '                    isreqCheckMargin = true
    '                end if
    '            end if
    '
    '            if (cscmmwdiv=cmjmaeipdiv) then
    '                if (cctrmargin<>cscmmargin) then
    '                    isreqCheckMargin = true
    '                end if
    '            end if
    '        else
    '            if isNULL(cscmmwdiv) then
    '                if (cctrmargin<>cscmmargin) then
    '                    isreqCheckMargin = true
    '                end if
    '            else
    '                if (cctrmargin<>cscmmargin) then
    '                    isreqCheckMargin = true
    '                end if
    '           end if
    '        end if
    
             if not (csellplace="ON") then
                 
                if isNULL(cscmmwdiv) and isNULL(cmjmaeipdiv) then
                    isDisabledMWMargin = true
                end if
            end if
    
            if (cctrmargin<=0) or (cctrmargin>=100) then
                isDisabledMWMargin = true
            end if
                 
                 
             if (isDisabledMWMargin) then '계약불가
                 dsbleCnt=dsbleCnt+1
                 errgroupid = errgroupid + ","+groupid
            end if   
            
    		  i=i+1
    			rsget.moveNext
    		loop
    		END IF
    		rsget.close
		
		if dsbleCnt < 1 then
		
        		strSql = "select manager_hp,manager_email  from [db_partner].[dbo].tbl_partner_group  where groupid ='"&groupid&"' "
        		  rsget.Open strSql, dbget, 0
        		 IF Not (rsget.EOF OR rsget.BOF) THEN
            	    mnghp = rsget(0)
            	    mngEmail = rsget(1)
            	end if
            	  rsget.Close     
          ''이메일 체크
         
            if (mngEmail<>"") then
                if (mngEmail="") or (InStr(mngEmail,"@")<0) or (Len(mngEmail)<8) then
                    response.write "<script>alert('업체 담당자 Email 주소가 유효하지 않습니다.');</script>"
                    response.write "<script>location.replace('" & refer & "');</script>"
                    dbget.close() : response.End
                end if
        
                sqlStr = "select IsNULL(p.usermail,'') as email from db_partner.dbo.tbl_user_tenbyten p"
                sqlStr = sqlStr & " where p.userid='"&session("ssBctID")&"'"
                sqlStr = sqlStr & " and p.userid<>''"
        
                rsget.Open sqlStr,dbget,1
                if Not rsget.Eof then
                    mailfrom = db2Html(rsget("email"))
                end if
                rsget.Close
        
                mailfrom = Trim(mailfrom)
        
                if (mailfrom="") or (InStr(mailfrom,"@")<0) or (Len(mailfrom)<8) then
                    response.write "<script>alert('발송자 Email  주소가 유효하지 않습니다.마이 정보에서 Email 수정 후 사용하시기 바랍니다.(등록된 이메일주소:"&mailfrom&")');</script>"
                    response.write "<script>location.replace('" & refer & "');</script>"
                    dbget.close()	:	response.End
                end if
            end if
         
        
                sqlstr = " update db_partner.dbo.tbl_partner_ctr_master"&VbCRLF
                sqlstr = sqlstr & " set ctrState=1"&VbCRLF                              ''업체 오픈
                sqlstr = sqlstr & " ,sendUserID='"&session("ssBctID")&"'"&VbCRLF
                sqlstr = sqlstr & " ,sendDate=getdate()"
                sqlstr = sqlstr & " where groupid='"&groupid&"'"&VbCRLF
                sqlstr = sqlstr & " and ctrState=0 and contracttype not in (8,9,10,16,17,18) "&VbCRLF ''수정중만 오픈가능 
                dbget.Execute  sqlstr 
        
           
        
         '   mngHp="010-6249-2706" ''임시 TEST
         '   mngEmail="@10x10.co.kr" ''임시 TEST
        
            if ( mngHp<>"") then
                '' SMS 발송 
                ''call SendNormalSMS(mngHp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 또는 SCM 업체계약관리 메뉴 참조")
                call SendNormalSMS_LINK(mngHp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 또는 SCM 업체계약관리 메뉴 참조")
            end if
        
            if (  mngEmail<>"") then
           
                strSql = " select ctrKey from db_partner.dbo.tbl_partner_ctr_master where groupid = '"&groupid&"'  and contracttype not in (8,9,10,16,17,18) and ctrstate = 1 " 
                rsget.Open strSql,dbget,1
                if not rsget.eof then
                 arrCtrKey = rsget.getRows() 
                end if
                rsget.close
                if isArray(arrCtrKey) then
                for intCK = 0 To UBound(arrCtrKey,2)
                if intCK = 0 then
                 iCtrKeyArr = arrCtrKey(0,intCK)
                else
                  iCtrKeyArr = iCtrKeyArr&","& arrCtrKey(0,intCK)
                end if
                Next
            end if
                '' 이메일 발송
                set ocontract = new CPartnerContract
                ocontract.FPageSize=50
                ocontract.FCurrPage = 1
                ocontract.FRectContractState = 1 ''오픈
                ocontract.FRectGroupID = groupid
                ocontract.FRectCtrKeyArr = iCtrKeyArr
                ocontract.GetNewContractList
        
                set oMdInfoList = new CPartnerContract
                oMdInfoList.FRectGroupID = groupid
                oMdInfoList.FRectContractState = 1 ''오픈
                 oMdInfoList.FRectCtrKeyArr = iCtrKeyArr
                oMdInfoList.getContractEmailMdList(FALSE)
        
                mailtitle       = "[텐바이텐] 신규 계약서가 발송 되었습니다."
                mailcontent   = makeCtrMailContents(ocontract,oMdInfoList,False)
        
                  Call SendMail(mailfrom, mngEmail, mailtitle, mailcontent)
        
                set ocontract=nothing
                set oMdInfoList=nothing
            end if
        end if
        
    NEXT
if (application("Svr_Info")	= "Dev") then
    response.write mailcontent
    
    response.end 
else 
     
     Call Alert_move("계약서가 발송되었습니다.","newctrList.asp?"&strParm )
    response.end 
end if
CASE "D"  '종료
dim gcheck,errG,strMsg
errG = ""
    For intLoop = 0 To UBound(arrgroupid) 
      groupid =   trim(arrgroupid(intLoop))
      
      strSql =" select groupid from db_partner.dbo.tbl_partner_ctr_master where groupid ='"&groupid&"' and contracttype not in (8,9,10,16,17,18) and ctrstate >=0 "
      rsget.Open strSql,dbget,1
      if Not rsget.Eof then
         gcheck = rsget(0)
      end if
      rsget.Close
                
      if  isNull(gcheck) or gcheck="" then         
      strSql = " update  db_partner.dbo.tmp_partner_Ctr_new set isusing = 0 , newfinuserid ='"&sregUserid&"' , isSubIn = 1 "
      strSql = strSql & " where groupid = '"&groupid&"'"  
      dbget.Execute  strSql 
     else
        if errG="" then
             errG = gcheck
        else
             errG = errG + ","+gcheck
        end if
      end if
    Next
    if errG <> "" then
        strMsg = "그룹번호 ["&errG&"]는 이미 계약중인 계약서가 존재하므로 계약종료가 불가능합니다\n"
    end if    
       Call Alert_move(strMsg&"계약서가 종료되었습니다.","newctrList.asp?"&strParm )
    response.end 
    
case "P" '개별종료
dim ctrIdx, chkIdx, cType
 ctrIdx   =  requestCheckvar(request("hidCI"),10)
 
     strSql =" select n.ctrIdx, n.contracttype  from db_partner.dbo.tmp_partner_ctr_New as n "
     strSql = strSql & "   left outer join db_partner.dbo.tbl_partner_ctr_master as m on m.groupid =n.groupid  "
     strSql = strSql & " and (isNull(n.brandid,'') = isNull(m.makerid,'')  ) and m.contracttype not in (8,9,10,16,17,18) "
     strSql = strSql & " left outer join db_partner.dbo.tbl_partner_ctr_sub as s on m.ctrKey = s.ctrKey and s.sellplace = n.sellplace and s.mwdiv = n.mwdiv "
     strSql = strSql & " where ctridx ="&ctrIdx&" and m.ctrKey is  Null " 
      rsget.Open strSql,dbget,1
      if Not rsget.Eof then
         chkIdx = rsget(0)
         cType = rsget(1)
      end if
      rsget.Close
      
      if isNull(chkIdx) or chkIdx ="" then 
           Call Alert_move("이미 진행중인 계약서가 존재합니다. 계약종료가 불가능합니다.","newctrList.asp?"&strParm )
    response.end 
     end if  
 
'     if cType = "8" then
'         Call Alert_move("기본계약서는 그룹단위 종료만 가능합니다. 그룹계약종료를 이용해주세요","newctrList.asp?"&strParm )
'    response.end 
'     end if
    
     strSql = " update  db_partner.dbo.tmp_partner_Ctr_new set isusing = 0 , newfinuserid ='"&sregUserid&"', isSubIn = 1 "
      strSql = strSql & " where ctrIdx = '"&ctrIdx&"'"  
      dbget.Execute  strSql 
    
     Call Alert_move("계약서가 종료되었습니다.","newctrList.asp?"&strParm )
    response.end 
end select    
%>