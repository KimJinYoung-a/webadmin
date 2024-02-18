<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 계약 관리
' Hieditor : 2016.06.30 정윤정 생성 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
1=0
dim sqlStr
dim contractType,contractName,onoffgubun,subType,contractContents11,contractContents13,contractContents12,contractContents14,contractContents19,contractContents
dim enddate, strenddate, contractdate
dim arrconts
dim arrKey,intLoop
  
'--거래기본
'--직매입
'--부속합의서
  enddate = "2016-09-30"
  strenddate= "2016년 09월 30일"
  contractdate = "2016-06-30"
  			 
        sqlStr = "select contractType, contractContents " &vbcrlf
        sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType " &vbcrlf
        sqlStr = sqlStr & " where contractType not in (8,9,10,16,17,18) "  

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            arrconts = rsget.getRows()
        end if
        rsget.Close
        
        if isArray(arrconts) then
        	for intLoop = 0 to uBound(arrconts,2)
        		if arrconts(0,intLoop) = 11 then
        			contractContents11 = arrconts(1,intLoop)
        		elseif arrconts(0,intLoop) = 12 then
        			contractContents12 = arrconts(1,intLoop)
        		elseif arrconts(0,intLoop) = 13 then
        			contractContents13 = arrconts(1,intLoop)
        		elseif arrconts(0,intLoop) = 14 then
        			contractContents14 = arrconts(1,intLoop)
        		elseif arrconts(0,intLoop) = 19 then
        			contractContents19 = arrconts(1,intLoop)
        	  end if			
          next
        end if
      
        sqlStr = " select top 1000 ctrkey ,contractType "
        sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master as m    "
	  		sqlStr = sqlStr & " left outer join db_partner.dbo.tbl_partner as p on m.makerid = p.id  "
  			sqlStr = sqlStr & " where m.contractType not in (8,9,10,16,17,18)  "
	      sqlStr = sqlStr & " and m.enddate = '2016-06-30' and m.ctrState = 7  "
	      sqlStr = sqlStr & " and ( p.groupid = m.groupid or p.groupid is null) " 
	      sqlStr = sqlStr & " order by m.ctrKey "
	      
	      rsget.Open sqlStr,dbget
	      if not rsget.eof then
	      	arrKey =  rsget.getRows()
	    	end if
        rsget.close()
        
        if isArray(arrKey) then
        	for intLoop = 0 to uBound(arrKey,2)
				if arrKey(1,intLoop) = 11 then
					contractContents = contractContents11
				elseif arrKey(1,intLoop) = 12 then
					contractContents = contractContents12
				elseif arrKey(1,intLoop) = 13 then
					contractContents = contractContents13
				elseif arrKey(1,intLoop) = 14 then
					contractContents = contractContents14
				elseif arrKey(1,intLoop) = 19 then
					contractContents = contractContents19
				end if
        	  	
        	 	if arrKey(1,intLoop) = 11 or arrKey(1,intLoop) = 13 or arrKey(1,intLoop) = 19 then 
	        		sqlStr = "		update  db_partner.dbo.tbl_partner_ctr_Detail "
	  					sqlStr = sqlStr & "		set detailvalue = '"&strenddate&"'"
	  					sqlStr = sqlStr & "		where detailKey ='$$ENDDATE$$'  "
	  					sqlStr = sqlStr & "				and   ctrkey = "&arrKey(0,intLoop)  
	  					dbget.Execute sqlStr		
  					end if	 
  						 
  					sqlStr = "		update  db_partner.dbo.tbl_partner_ctr_Detail "
  					sqlStr = sqlStr & "	set detailvalue = '"&contractdate &"'"
  					sqlStr = sqlStr & "	where detailKey ='$$CONTRACT_DATE$$'   "
  					sqlStr = sqlStr & "			and   ctrkey = "&arrKey(0,intLoop)		
  					 
  					dbget.Execute sqlStr		
  						
  					 '' 계약서 DB 내용으로 치환
				    if  (FillContractContentsByDB_Re(arrKey(0,intLoop)	, contractContents)) then  
				
				        sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
				        sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
				        sqlStr = sqlStr & " , renewdate = getdate() "
				        sqlStr = sqlStr & " , enddate = '"&enddate&"' " 
				        sqlStr = sqlStr & " where ctrKey=" & arrKey(0,intLoop)		  
				        dbget.Execute sqlStr
				        
				        sqlStr = "insert into db_partner.dbo.tbl_partner_ctr_stateLog "
				        sqlStr = sqlStr & " (ctrkey,logtype, logmsg) "
				        sqlStr = sqlStr & " values "
				        sqlStr = sqlStr & " ("&arrKey(0,intLoop)&",8, '계약서 자동연장') "
				          dbget.Execute sqlStr
				        	 response.write intLoop&"-"&arrKey(0,intLoop)&"_성공" &"<BR>"
				    else
				        response.write arrKey(0,intLoop)& "_계약서 작성실패"&"<BR>"
				    end if
    			Next
      	end if 
      
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->