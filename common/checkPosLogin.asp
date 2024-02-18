<%
'###########################################################
' Description : 포스 로그인
' Hieditor : 2011.02.15 서동석 생성
'			 2011.02.19 한용민 수정
'###########################################################

'//함부로 임의로 수정 하지 마세요. 오프라인 포스와 연동 되는 키 입니다.
function checkPosSSN(pssnkey, posuid, dummikey)
    dim ret : ret = false
   	dim tmpssnkey
    
		'/아이디  세번째 값을 넣고, 아이디가 없거나 길이가 세자리가 안되면 0 고정  
		tmpssnkey = tmpssnkey & mid(posuid,3,1)
		if posuid = "" or len(posuid) < 3 then tmpssnkey = tmpssnkey & "0"
			
		'/아이디 첫번째 값을 넣고, 아이디가 없거나 길이가 한자리가 안되면 3 고정
		tmpssnkey = tmpssnkey & left(posuid,1)
		if posuid = "" or len(posuid) < 1 then tmpssnkey = tmpssnkey & "3"
					
		'/더미키  세번째를 넣고, 더미키가 없거나 길이가 세자리가 안되면 2 고정
		tmpssnkey = tmpssnkey & mid(dummikey,3,1)
		if dummikey = "" or len(dummikey) < 3 then tmpssnkey = tmpssnkey & "2"
			
		'/더미키 두번째를 넣고, 더미키가 없거나 길이가 두자리가 안되면 1 고정
		tmpssnkey = tmpssnkey & mid(dummikey,2,1)
		if dummikey = "" or len(dummikey) < 2 then tmpssnkey = tmpssnkey & "1" 
		
		ret = (UCASE(md5(tmpssnkey))=UCASE(pssnkey))
    checkPosSSN = ret
end function

Dim IsPosLogin : IsPosLogin = (session("poslogin") = 1)

''포스 로그인 2010-06추가
IF ((session("ssBctId")="") or (request.Form("pssnkey")<>"")) then 
    if (LEft(Request.ServerVariables("HTTP_TPSDUMMI"),Len("tenPos Client v"))="tenPos Client v") and (request.Form("posuid")<>"") and (request.Form("pssnkey")<>"") and (request.Form("dummikey")<>"") then 
        if (checkPosSSN(request.Form("pssnkey"),request.Form("posuid"),request.Form("dummikey"))) then
            
            Dim Tsql, tid, tuserdiv , tssAdminLsn, tbigo
            
            Tsql = "select top 1 A.id, A.company_name, A.tel, A.fax, A.url, A.email, A.userdiv, A.password, A.groupid " + vbCrlf
            Tsql = Tsql + "	, B.part_sn, A.level_sn, B.job_sn, B.username,  B.direct070, B.usermail " + vbCrlf
			Tsql = Tsql + " ,(select top 1 shopid" + vbCrlf
			Tsql = Tsql + " 	from db_partner.dbo.tbl_partner_shopuser" + vbCrlf
			Tsql = Tsql + " 	where b.empno=empno and firstisusing='Y') as firstshopid" + vbCrlf          
            Tsql = Tsql + " from [db_partner].[dbo].tbl_partner as A " + vbCrlf
            Tsql = Tsql + " left join db_partner.dbo.tbl_user_tenbyten as B"
            Tsql = Tsql + " 	ON A.id = B.userid AND B.isUsing = 1" + vbCrlf

            ' 퇴사예정자 처리	' 2018.10.16 한용민
            Tsql = Tsql & " 		and (b.statediv ='Y' or (b.statediv ='N' and datediff(dd,b.retireday,getdate())<=0))" & vbcrlf
            Tsql = Tsql + " where A.id = '" + request.Form("posuid") + "'" + vbCrlf
            Tsql = Tsql + " and A.isusing='Y'"

            'response.write Tsql & "<br>"
            rsget.Open Tsql,dbget,1
            if not rsget.Eof then          
                tid = rsget("id")
                tuserdiv = rsget("userdiv")
                tssAdminLsn = rsget("part_sn")
                tbigo = rsget("firstshopid")
            end if
            rsget.close           
          
            session("ssBctId")=tid
            session("ssBctDiv")=tuserdiv
            session("poslogin") = 1
            session("ssAdminPsn") = tssAdminLsn
            session("ssBctBigo") = tbigo
        end if
    end if
end if    
%>