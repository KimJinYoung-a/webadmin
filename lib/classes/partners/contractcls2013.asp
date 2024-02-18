<%
'###########################################################
' Description : 브랜드 계약 클래스
' Hieditor : 2009.04.07 서동석 생성
'			 2010.05.25 한용민 수정
'            2022.01.25 원승현 수정(DocuSign 추가)
'###########################################################
'' DocuSign 도입전 사용하던 계약서
'' 기존 수기 계약과 UPLUS 계약은 아래 계약서로 진행
CONST DEFAULT_CONTRACTTYPE = 19         '거래기본계약서_202102
CONST DEFAULT_CONTRACTTYPE_M = 13       '직매입계약서
CONST ADD_CONTRACTTYPE = 12             '거래기본계약부속합의서
CONST ADD_CONTRACTTYPE_M = 14           '직매입계약부속합의서

'' DocuSign은 아래 계약서로 진행
DIM DEFAULT_NEWCONTRACTTYPE         'DocuSign도입과 같이 갱신된 거래기본계약서(2021.11)
DIM DEFAULT_NEWCONTRACTTYPE_M       'DocuSign도입과 같이 갱신된 물품공급계약서(직매입)(2021.11)
DIM ADD_NEWCONTRACTTYPE             'DocuSign도입과 같이 갱신된 거래기본계약 부속합의서(2021.11)
DIM ADD_NEWCONTRACTTYPE_M           'DocuSign도입과 같이 갱신된 물품공급계약(직매입) 부속합의서(2021.11)
DIM AUTHEX_NEWCONTRACTTYPE          'DocuSign도입과 같이 추가된 공인전자서명 면제 요청서(2021.11)
DIM SPECIALAPPOINTMENTCONTRACTTYPE  'DocuSign도입과 같이 추가된 특약 계약서(2022.02)
'' 개발서버에 있는 id값은 틀리므로 분기처리
If (application("Svr_Info")	= "Dev") then
    DEFAULT_NEWCONTRACTTYPE         = 20
    DEFAULT_NEWCONTRACTTYPE_M       = 22
    ADD_NEWCONTRACTTYPE             = 21
    ADD_NEWCONTRACTTYPE_M           = 23
    AUTHEX_NEWCONTRACTTYPE          = 24
    SPECIALAPPOINTMENTCONTRACTTYPE  = 25
Else
    DEFAULT_NEWCONTRACTTYPE         = 20
    DEFAULT_NEWCONTRACTTYPE_M       = 31
    ADD_NEWCONTRACTTYPE             = 22
    ADD_NEWCONTRACTTYPE_M           = 32
    AUTHEX_NEWCONTRACTTYPE          = 33
    SPECIALAPPOINTMENTCONTRACTTYPE  = 34
End If

Session.CodePage = 65001
'' DocuSign용 도장 이미지 Base64
CONST DocuSignStampBase64 = "data:image/jpeg;base64,/9j/4QAYRXhpZgAASUkqAAgAAAAAAAAAAAAAAP/sABFEdWNreQABAAQAAAA8AAD/4QMsaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLwA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/PiA8eDp4bXBtZXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEvIiB4OnhtcHRrPSJBZG9iZSBYTVAgQ29yZSA2LjAtYzAwNiA3OS4xNjQ3NTMsIDIwMjEvMDIvMTUtMTE6NTI6MTMgICAgICAgICI+IDxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyI+IDxyZGY6RGVzY3JpcHRpb24gcmRmOmFib3V0PSIiIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bXA6Q3JlYXRvclRvb2w9IkFkb2JlIFBob3Rvc2hvcCAyMi4zIChXaW5kb3dzKSIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDpBRjhBMzczRTdGMUMxMUVDQjAxRUU5NEFFRTAzNEM1MyIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDpBRjhBMzczRjdGMUMxMUVDQjAxRUU5NEFFRTAzNEM1MyI+IDx4bXBNTTpEZXJpdmVkRnJvbSBzdFJlZjppbnN0YW5jZUlEPSJ4bXAuaWlkOkFGOEEzNzNDN0YxQzExRUNCMDFFRTk0QUVFMDM0QzUzIiBzdFJlZjpkb2N1bWVudElEPSJ4bXAuZGlkOkFGOEEzNzNEN0YxQzExRUNCMDFFRTk0QUVFMDM0QzUzIi8+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+/+4ADkFkb2JlAGTAAAAAAf/bAIQABgQEBAUEBgUFBgkGBQYJCwgG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"
Session.CodePage = 949

CONST CPrvContract = TRUE
CONST ChashVal = "TBTCTR"


    public function GetContractEcStateName(ContractState)
        dim buf
        Select Case ContractState
            Case 0
                : buf = "수정중(미전송)"
            Case 1
                : buf = "계약오픈(검토대기)"
            Case 2
                : buf = "계약반려(검토반려)"
            Case 3
                : buf = "계약진행(결재완료)"    
            Case 6
                : buf = "서명진행"    
            Case 7
                : buf = "계약완료"
             Case 8
                : buf = "계약파기요청"   
            Case 9
                : buf = "계약종료"    
            Case -1
                : buf = "삭제"
            Case else
                : buf = ContractState
        end Select

        GetContractEcStateName = buf
    end function
    
    
    public function GetContractEcState(ContractStateName)
        dim buf
        Select Case ContractStateName
            Case "미전송"
                : buf = "0"
            Case "검토대기"
                : buf =  "1"
             Case "검토반려"
                : buf =  "2"    
            Case "결재완료"
                : buf = "3"
            Case "전자서명진행"  
                : buf =   "6"
            Case "계약완료"
                : buf = "7"
            Case "계약파기요청"
                : buf = "8"
            Case "계약파기"
                : buf = "9"   
            Case "계약종료" 
                : buf =    "9"
            Case "삭제"
                : buf = -1
            Case "-1"
                : buf = -1
            Case else
                : buf = "-2"
        end Select

        GetContractEcState = buf
    end function
    
function getLastPrecontractID(imakerid) ''기존 계약서
    dim sqlStr
    getLastPrecontractID = -1
    sqlStr = "select top 1 c.contractID "
	sqlStr = sqlStr & " , c.makerid, c.contractType, c.contractName"
	sqlStr = sqlStr & " , c.contractNo, c.contractState, c.reguserid, c.regdate, c.confirmdate, c.finishdate"
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract c"
    sqlStr = sqlStr & " where makerid='" & imakerid & "'"
    sqlStr = sqlStr & " and contractState>0"
    sqlStr = sqlStr & " and contractState<3"
    sqlStr = sqlStr & " order by c.contractID desc"
    
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        getLastPrecontractID = rsget("contractID")
    end if
    rsget.Close
end function

function isPreTypeContractExists(imakerid) ''기존타입 계약서 ''브랜드별
    dim sqlStr
    sqlStr = "select count(*) as CNT"
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract c"
    sqlStr = sqlStr & " where makerid='" & imakerid & "'"
    sqlStr = sqlStr & " and contractState>0"
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        isPreTypeContractExists = rsget("CNT")>0
    end if
    rsget.Close
end function


function isNotFinishNewContractExists(imakerid, igroupid, byRef isNewContractTypeExists)
    ''istate 3 :업체확인 , 7 :계약완료
    if (imakerid="") and (igroupid="") then Exit function

    if (igroupid="") then igroupid = getPartnerId2GroupID(imakerid)
    dim sqlStr, retCNT : retCNT=0
    isNewContractTypeExists=false

    sqlStr = " select count(*) as ttlCNT"
    sqlStr = sqlStr & " ,isNULL(sum(CASE WHEN ctrState<3 then 1 else 0 end),0) notConfirmdCNT"
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master M"&VbCRLF
    sqlStr = sqlStr & " where groupid='"&igroupid&"'"&VbCRLF
    sqlStr = sqlStr & " and ctrState>0 and contracttype not in (8,9,10,16,17,18) "&VbCRLF   ''수정중 이상.
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        retCNT         = rsget("notConfirmdCNT")
        isNewContractTypeExists = rsget("ttlCNT")>0
    end if
    rsget.Close

    isNotFinishNewContractExists = (retCNT>0)
end function

function fnCheckHoldContract(imakerid, byRef isOnHoldContract,byRef isOFHoldContract)
    dim sqlStr

    isOnHoldContract = false
    isOFHoldContract = false
    sqlStr = " select "
    sqlStr = sqlStr & " isNULL(sum(CASE WHEN onoffgbn='ON' THEN 1 else 0 end),0) as OnHoldExosts"
    sqlStr = sqlStr & " ,isNULL(sum(CASE WHEN onoffgbn='OF' THEN 1 else 0 end),0) as OfHoldExosts"
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_Hold"
    sqlStr = sqlStr & " where makerid='"&imakerid&"'"
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        isOnHoldContract = rsget("OnHoldExosts")>0
        isOFHoldContract = rsget("OfHoldExosts")>0
    end if
    rsget.Close

end function

function fnCgeckIsOldBrand(imakerid,imonth)
    dim sqlStr, diffmonth : diffmonth=0

    sqlStr = "select datediff(m,regdate,getdate()) as diffmonth from db_partner.dbo.tbl_partner where id='"&imakerid&"'"
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        diffmonth         = rsget("diffmonth")
    end if
    rsget.Close

    fnCgeckIsOldBrand = (diffmonth>=imonth)

end function

function FillContractContentsByDB(ctrKey, byref contractContents)
    dim sqlStr, bufStr
    dim subContractExists : subContractExists=false
    dim isMeaipContract

    ''잘림.
    ''sqlStr = "exec db_partner.[dbo].[sp_Ten_partner_AddContract_FillContents] "&ctrKey
    ''dbget.Execute sqlStr
    dim detailKey,detailValue,ctrState,onoffgubun,subtype, makerid
    dim sellplace,mwdiv,defaultmargin,defaultdeliveryType,defaultFreebeasongLimit,defaultdeliverpay,mwdivName, sellplaceName, signType

    sqlStr = "select M.ctrState, M.makerid, T.contractContents, T.subtype, T.onoffgubun"
    sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_master M"
	sqlStr = sqlStr&" 	Join db_partner.dbo.tbl_partner_contractType T"
	sqlStr = sqlStr&" 	on M.contractType=T.contractType"
    sqlStr = sqlStr&" where M.ctrKey="&ctrKey
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ctrState         = rsget("ctrState")
        subtype          = rsget("subtype")
        onoffgubun       = rsget("onoffgubun")
        makerid          = rsget("makerid")
        contractContents = db2Html(rsget("contractContents"))
    end if
    rsget.Close

    if (ctrState>=3) then exit function

    isMeaipContract = (subtype=7)

    ''-- 부속 합의서인경우
	if ((subtype<>0)) then
	    bufStr=""

	    sqlStr = "select top 1110 S.sellplace, S.mwdiv,S.defaultmargin,S.defaultdeliveryType,S.defaultFreebeasongLimit,S.defaultdeliverpay "
	    sqlStr = sqlStr&" ,(CASE WHEN sellplace='ON' and mwdiv='M' THEN '매입' "
	    sqlStr = sqlStr&"   WHEN sellplace='ON' and mwdiv='W' THEN '위탁' "
	    sqlStr = sqlStr&"   WHEN sellplace='ON' and mwdiv='U' THEN '업체' "
	    sqlStr = sqlStr&"   WHEN sellplace<>'ON' and isNULL(c.comm_name,'')<>'' THEN c.comm_name ELSE mwdiv END) as mwdivName"
	    sqlStr = sqlStr&" ,(CASE WHEN sellplace='ON' then '온라인'"
	    sqlStr = sqlStr&" WHEN sellplace<>'ON' and isNULL(u.shopname,'')<>'' THEN u.shopname + ' 매장' ELSE sellplace END) as sellplacename"
        sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_Sub S"
        sqlStr = sqlStr&" left join db_jungsan.dbo.tbl_jungsan_comm_code C"
        sqlStr = sqlStr&" on S.mwdiv=C.comm_cd"
        sqlStr = sqlStr&" left join db_shop.dbo.tbl_shop_user u"
        sqlStr = sqlStr&" on s.sellplace=u.userid"
        sqlStr = sqlStr&" where S.ctrKey="&ctrKey
        sqlStr = sqlStr&" order by S.ctrSubKey"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            do until rsget.eof
                sellplace           = rsget("sellplace")
                mwdiv               = rsget("mwdiv")
                mwdivName           = rsget("mwdivName")
                sellplaceName       = rsget("sellplaceName")
                defaultmargin       = rsget("defaultmargin")

                defaultdeliveryType = rsget("defaultdeliveryType")
                defaultFreebeasongLimit = rsget("defaultFreebeasongLimit")
                defaultdeliverpay = rsget("defaultdeliverpay")

                if (onoffgubun="ON") then
                    bufStr = bufStr & "<tr>"
                    bufStr = bufStr & "<td align='center'>"&makerid&"</td>"
                    bufStr = bufStr & "<td align='center'>"&sellplaceName&"</td>"
                    bufStr = bufStr & "<td align='center'>"&mwdivName&"</td>"
                    if (isMeaipContract) then ''매입인경우 공급율표시
                        bufStr = bufStr & "<td align='center'>"&(100-CLNG(defaultmargin*100)/100)&" %</td>"
                    else
                        bufStr = bufStr & "<td align='center'>"&CLNG(defaultmargin*100)/100&" %</td>"
                    end if
                    IF (mwdiv="U") and (defaultdeliveryType="7") then
                        bufStr = bufStr & "<td align='center'>업체착불배송</td>"
                    elseIF (mwdiv="U") and (defaultdeliveryType="9") then
                        bufStr = bufStr & "<td align='center'>업체조건배송"
                        bufStr = bufStr & "<br>"&FormatNumber(defaultFreebeasongLimit,0)&"원미만"
                        bufStr = bufStr & "<br>배송료"&FormatNumber(defaultdeliverpay,0)&"원"
                        bufStr = bufStr & "</td>"
                    else
                        bufStr = bufStr & "<td align='center'>&nbsp;</td>"
                    end if

                    bufStr = bufStr & "</tr>"
                else

                end if
            rsget.movenext
		    loop
        end if
        rsget.Close

        if (bufStr<>"") then
            if (isMeaipContract) then
                bufStr="<thead><tr><th>브랜드ID</th><th>판매채널</th><th>계약형태</th><th>기본공급율</th><th>비고</th></tr>"&bufStr
            else
                bufStr="<thead><tr><th>브랜드ID</th><th>판매채널</th><th>계약형태</th><th>기본수수료</th><th>비고</th></tr>"&bufStr
            end if
            bufStr=bufStr&"</thead><tbody>"

            sqlStr = "update D"
            sqlStr = sqlStr&" SET detailvalue='"&Newhtml2db(bufStr)&"'"
            sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_Detail D"
            sqlStr = sqlStr&" where D.ctrKey="&ctrKey
            sqlStr = sqlStr&" and D.detailKey='$$CONTRACT_CONTS$$'"

            dbget.Execute sqlStr
        end if

    end if

    '' SignType값을 가져온다.
    sqlStr = "SELECT  signType FROM db_partner.dbo.tbl_partner_ctr_master where ctrKey="&ctrKey
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        signType         = rsget("signType")
    end if
    rsget.Close

    sqlStr = "SELECT  detailKey,detailValue FROM db_partner.dbo.tbl_partner_ctr_Detail where ctrKey="&ctrKey
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        do until rsget.eof
            detailKey         = rsget("detailKey")
            detailValue       = rsget("detailValue")

            if (detailKey="$$CONTRACT_DATE$$") then
                bufStr  = detailValue
                bufStr  = Left(bufStr,4) & "년" & Mid(bufStr,6,2) & "월" & Mid(bufStr,9,2) & "일"
                If signType <> "D" Then
                    contractContents = Replace(contractContents,detailKey,bufStr)
                End If
            else
                contractContents = Replace(contractContents,detailKey,detailValue)
            end if


        rsget.movenext
		loop
    end if
    rsget.Close

    FillContractContentsByDB = true
end function


function FillContractContentsByDB_Re(ctrKey, byref contractContents)
    dim sqlStr, bufStr
    dim subContractExists : subContractExists=false
    dim isMeaipContract

    ''잘림.
    ''sqlStr = "exec db_partner.[dbo].[sp_Ten_partner_AddContract_FillContents] "&ctrKey
    ''dbget.Execute sqlStr
    dim detailKey,detailValue,ctrState,onoffgubun,subtype, makerid
    dim sellplace,mwdiv,defaultmargin,defaultdeliveryType,defaultFreebeasongLimit,defaultdeliverpay,mwdivName, sellplaceName

    sqlStr = "select M.ctrState, M.makerid, T.contractContents, T.subtype, T.onoffgubun"
    sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_master M"
	sqlStr = sqlStr&" 	Join db_partner.dbo.tbl_partner_contractType T"
	sqlStr = sqlStr&" 	on M.contractType=T.contractType"
    sqlStr = sqlStr&" where M.ctrKey="&ctrKey
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ctrState         = rsget("ctrState")
        subtype          = rsget("subtype")
        onoffgubun       = rsget("onoffgubun")
        makerid          = rsget("makerid")
        contractContents = db2Html(rsget("contractContents"))
    end if
    rsget.Close

    'if (ctrState>=7) then exit function

    isMeaipContract = (subtype=7)

    ''-- 부속 합의서인경우
	if ((subtype<>0)) then
	    bufStr=""

	    sqlStr = "select top 1110 S.sellplace, S.mwdiv,S.defaultmargin,S.defaultdeliveryType,S.defaultFreebeasongLimit,S.defaultdeliverpay "
	    sqlStr = sqlStr&" ,(CASE WHEN sellplace='ON' and mwdiv='M' THEN '매입' "
	    sqlStr = sqlStr&"   WHEN sellplace='ON' and mwdiv='W' THEN '위탁' "
	    sqlStr = sqlStr&"   WHEN sellplace='ON' and mwdiv='U' THEN '업체' "
	    sqlStr = sqlStr&"   WHEN sellplace<>'ON' and isNULL(c.comm_name,'')<>'' THEN c.comm_name ELSE mwdiv END) as mwdivName"
	    sqlStr = sqlStr&" ,(CASE WHEN sellplace='ON' then '온라인'"
	    sqlStr = sqlStr&" WHEN sellplace<>'ON' and isNULL(u.shopname,'')<>'' THEN u.shopname + ' 매장' ELSE sellplace END) as sellplacename"
        sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_Sub S"
        sqlStr = sqlStr&" left join db_jungsan.dbo.tbl_jungsan_comm_code C"
        sqlStr = sqlStr&" on S.mwdiv=C.comm_cd"
        sqlStr = sqlStr&" left join db_shop.dbo.tbl_shop_user u"
        sqlStr = sqlStr&" on s.sellplace=u.userid"
        sqlStr = sqlStr&" where S.ctrKey="&ctrKey
        sqlStr = sqlStr&" order by S.ctrSubKey"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            do until rsget.eof
                sellplace           = rsget("sellplace")
                mwdiv               = rsget("mwdiv")
                mwdivName           = rsget("mwdivName")
                sellplaceName       = rsget("sellplaceName")
                defaultmargin       = rsget("defaultmargin")

                defaultdeliveryType = rsget("defaultdeliveryType")
                defaultFreebeasongLimit = rsget("defaultFreebeasongLimit")
                defaultdeliverpay = rsget("defaultdeliverpay")

                if (onoffgubun="ON") then
                    bufStr = bufStr & "<tr>"
                    bufStr = bufStr & "<td align='center'>"&makerid&"</td>"
                    bufStr = bufStr & "<td align='center'>"&sellplaceName&"</td>"
                    bufStr = bufStr & "<td align='center'>"&mwdivName&"</td>"
                    if (isMeaipContract) then ''매입인경우 공급율표시
                        bufStr = bufStr & "<td align='center'>"&(100-CLNG(defaultmargin*100)/100)&" %</td>"
                    else
                        bufStr = bufStr & "<td align='center'>"&CLNG(defaultmargin*100)/100&" %</td>"
                    end if
                    IF (mwdiv="U") and (defaultdeliveryType="7") then
                        bufStr = bufStr & "<td align='center'>업체착불배송</td>"
                    elseIF (mwdiv="U") and (defaultdeliveryType="9") then
                        bufStr = bufStr & "<td align='center'>업체조건배송"
                        bufStr = bufStr & "<br>"&FormatNumber(defaultFreebeasongLimit,0)&"원미만"
                        bufStr = bufStr & "<br>배송료"&FormatNumber(defaultdeliverpay,0)&"원"
                        bufStr = bufStr & "</td>"
                    else
                        bufStr = bufStr & "<td align='center'>&nbsp;</td>"
                    end if

                    bufStr = bufStr & "</tr>"
                else

                end if
            rsget.movenext
		    loop
        end if
        rsget.Close

        if (bufStr<>"") then
            if (isMeaipContract) then
                bufStr="<thead><tr><th>브랜드ID</th><th>판매채널</th><th>계약형태</th><th>기본공급율</th><th>비고</th></tr>"&bufStr
            else
                bufStr="<thead><tr><th>브랜드ID</th><th>판매채널</th><th>계약형태</th><th>기본수수료</th><th>비고</th></tr>"&bufStr
            end if
            bufStr=bufStr&"</thead><tbody>"

            sqlStr = "update D"
            sqlStr = sqlStr&" SET detailvalue='"&Newhtml2db(bufStr)&"'"
            sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_Detail D"
            sqlStr = sqlStr&" where D.ctrKey="&ctrKey
            sqlStr = sqlStr&" and D.detailKey='$$CONTRACT_CONTS$$'"

            dbget.Execute sqlStr
        end if

    end if

    '' SignType값을 가져온다.
    sqlStr = "SELECT  signType FROM db_partner.dbo.tbl_partner_ctr_master where ctrKey="&ctrKey
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        signType         = rsget("signType")
    end if
    rsget.Close    

    sqlStr = "SELECT  detailKey,detailValue FROM db_partner.dbo.tbl_partner_ctr_Detail where ctrKey="&ctrKey
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        do until rsget.eof
            detailKey         = rsget("detailKey")
            detailValue       = rsget("detailValue")

            if (detailKey="$$CONTRACT_DATE$$") then
                bufStr  = detailValue
                bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
                If signType<>"D" Then
                    contractContents = Replace(contractContents,detailKey,bufStr)
                End If
            else
                contractContents = Replace(contractContents,detailKey,detailValue)
            end if


        rsget.movenext
		loop
    end if
    rsget.Close

    FillContractContentsByDB_Re = true
end function
function getDefaultContractValue(aKey, ogroupinfo)
    dim mdusername
    dim Buf
    select case aKey
        ''case "$$CONTRACT_NO$$"          ''계약서 번호. -> 자동생성
        ''    : getDefaultContractValue = mdusername
        case "$$A_CHARGE$$"             ''계약담당자.
            mdusername = getMdUserName(ogroupinfo.FOneItem.Fmduserid)
            : getDefaultContractValue = mdusername
        case "$$A_UPCHENAME$$"
            : getDefaultContractValue = "(주)텐바이텐"
        case "$$A_CEONAME$$"
            : getDefaultContractValue = "최은희"
        case "$$A_COMPANY_NO$$"
            : getDefaultContractValue = "211-87-00620"
        case "$$A_COMPANY_ADDR$$"
            : getDefaultContractValue = "서울 종로구 대학로 57, 교육동 14층"

        case "$$B_CHARGE$$"
            : getDefaultContractValue = ogroupinfo.FOneItem.FManager_Name
        case "$$B_UPCHENAME$$"
            : getDefaultContractValue = ogroupinfo.FOneItem.Fcompany_name
        case "$$B_CEONAME$$"
            : getDefaultContractValue = ogroupinfo.FOneItem.Fceoname
        case "$$B_COMPANY_NO$$"
            : getDefaultContractValue = ogroupinfo.FOneItem.Fcompany_no
        case "$$B_COMPANY_ADDR$$"
            : getDefaultContractValue = ogroupinfo.FOneItem.Fcompany_address & " " & ogroupinfo.FOneItem.Fcompany_address2
        case "$$B_BRANDNAME$$"
            : getDefaultContractValue = ogroupinfo.FOneItem.Fsocname_kor
        ''case "$$B_DELIVER_MANAGER$$"
        ''    : getDefaultContractValue = ogroupinfo.FOneItem.Fdeliver_name


        case "$$DEFAULT_ITEM_MARGIN$$"
            : getDefaultContractValue = ogroupinfo.FOneItem.Fdefaultmargine & "%"
        case "$$DEFAULT_SERVICE_MARGIN$$"
            : getDefaultContractValue = ogroupinfo.FOneItem.Fdefaultmargine & "%"

        ''온/오프 통합 할것. :: 온라인기준으로
        case "$$DEFAULT_JUNGSANDATE$$"
            :
            if (ogroupinfo.FOneItem.Fjungsan_date="15일") then
                getDefaultContractValue = "판매(제공)월의 " & "익월 15일"
            elseif (ogroupinfo.FOneItem.Fjungsan_date="말일") then
                getDefaultContractValue = "판매(제공)월의 " & "익월 말일"
            elseif (ogroupinfo.FOneItem.Fjungsan_date="수시") then
                getDefaultContractValue = "판매(제공)월의 " & "익월 5일"
            elseif (ogroupinfo.FOneItem.Fjungsan_date="5일") then
                getDefaultContractValue = "판매(제공)월의 " & "익월 5일"
            end if

            if Not isNULL(ogroupinfo.FOneItem.Fjungsan_date_off) and (ogroupinfo.FOneItem.Fjungsan_date_off<>"") then
                if ogroupinfo.FOneItem.Fjungsan_date<>ogroupinfo.FOneItem.Fjungsan_date_off then
                        if (ogroupinfo.FOneItem.Fjungsan_date="말일") and  (ogroupinfo.FOneItem.Fjungsan_date_off="15일") then
                            getDefaultContractValue = "판매(제공)월의 " & "익월 15일"
                        end if

                        if (ogroupinfo.FOneItem.Fjungsan_date="말일") and  ((ogroupinfo.FOneItem.Fjungsan_date_off="수시") or (ogroupinfo.FOneItem.Fjungsan_date_off="5일")) then
                            getDefaultContractValue = "판매(제공)월의 " & "익월 5일"
                        end if

                        if (ogroupinfo.FOneItem.Fjungsan_date="15일") and  ((ogroupinfo.FOneItem.Fjungsan_date_off="수시") or (ogroupinfo.FOneItem.Fjungsan_date_off="5일")) then
                            getDefaultContractValue = "판매(제공)월의 " & "익월 5일"
                        end if
                        
                        if (ogroupinfo.FOneItem.Fjungsan_date="15일") and   (ogroupinfo.FOneItem.Fjungsan_date_off="말일")   then
                            getDefaultContractValue = "판매(제공)월의 " & "익월 말일"
                        end if
                end if
            end if
        case "$$CONTRACT_DATE$$"  ''2013년 00월 00일
            :
            if (Now()<"2014-01-01") then
                getDefaultContractValue = "2014-01-01"
            else
                getDefaultContractValue = Left(Now(),10)  ''Left(Buf,4)+"년 "+Mid(Buf,6,2)+"월 "+Mid(Buf,9,2)+"일" //계약서 내용만 치환
            end if
        case "$$INSURANCE_FEE$$"                        '' 보증보험
            : getDefaultContractValue = "0 만원"
        
        case "$$ENDDATE$$" 
            :
            dim nMonth : nMonth = month(date())
            if (nMonth<=3) then   
                getDefaultContractValue = year(date())&"-06-30"
            elseif (nMonth>3 and nMonth<=6) then   
                getDefaultContractValue = year(date())&"-09-30"
            elseif (nMonth>6 and nMonth<=9) then   
                getDefaultContractValue = year(date())&"-12-31"
            elseif (nMonth>9 and nMonth<=12) then   
                getDefaultContractValue = year(dateadd("yyyy",1,date())) &"-03-31" 
            end if    
        case Else
            : getDefaultContractValue = ""
    end select
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

public function fnContractStateColor(iCtrState)
    Select Case iCtrState
        Case 0
            : fnContractStateColor = "#000000"
        Case 1
            : fnContractStateColor = "#44BB44"
        Case 3
            : fnContractStateColor = "#7777FF"
        Case 7
            : fnContractStateColor = "#FF7777"
        Case -1
            : fnContractStateColor = "#AAAAAA"
       Case else
            : fnContractStateColor = "#000000"
    end Select
end function

public function fnContractStateName(iCtrState)
    dim buf
    Select Case iCtrState
        Case 0
            : buf = "수정중(미전송)"
        Case 1
            : buf = "계약오픈(검토대기)"
        Case 2
            : buf = "계약반려(검토반려)"
         Case 3
            : buf = "계약확인(결재완료)"    
        Case 6
            : buf = "서명진행"
        Case 7
            : buf = "계약완료"    
        Case 8
            : buf = "계약파기요청"   
        Case 9
            : buf = "계약종료"        
        Case -1
            : buf = "삭제"
        Case else
            : buf = iCtrState
    end Select

                
    fnContractStateName = "<font color='"&fnContractStateColor(iCtrState)&"'>"&buf&"</font>"
end function


function drawOfJungsanGbn(selname,selval,addcont)
    dim ret
    ret = "<select name='"&selname&"' "&addcont&">"
    ret = ret &"<option value='B012' "&CHKIIF(selval="B012","selected","")&">업체위탁"
    ret = ret &"<option value='B031' "&CHKIIF(selval="B031","selected","")&">출고매입"
    ret = ret &"<option value='B013' "&CHKIIF(selval="B013","selected","")&">출고위탁"
    ret = ret &"</select>"
    drawOfJungsanGbn=ret
end function

function getPartnerId2GroupID(ipartnerid)
    dim sqlStr
	sqlStr = "select groupid from db_partner.dbo.tbl_partner where id='"&ipartnerid&"'"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
	    getPartnerId2GroupID = rsget("groupid")
    end if
    rsget.Close
end function

''개인정보 수집 관련 계약서
function getPriContractContents(bUpchename)
    dim ret
    dim fs,f
    set fs=Server.CreateObject("Scripting.FileSystemObject")
    set f=fs.OpenTextFile(Server.MapPath("/designer/company/contract/viewContractWeb_Pri.html"),1)
    ret = f.ReadAll
    f.Close
    set f=Nothing
    set fs=Nothing

    getPriContractContents = replace(ret,"$$B_UPCHENAME$$",bUpchename)
end function

''개인정보 수집 관련 계약서
function getPriContractContentsDocuSign()
    dim ret
    dim fs,f
    set fs=Server.CreateObject("Scripting.FileSystemObject")
    set f=fs.OpenTextFile(Server.MapPath("/designer/company/contract/viewContractWeb_Pri_DocuSign.html"),1)
    ret = f.ReadAll
    f.Close
    set f=Nothing
    set fs=Nothing

    getPriContractContentsDocuSign = ret
end function

function makeCtrMailContents(ocontract,oMdInfoList, isPreView)
    dim ret, buf
    dim fs,f, i, offmdCnt : offmdCnt=0
    set fs=Server.CreateObject("Scripting.FileSystemObject")
    set f=fs.OpenTextFile(Server.MapPath("/lib/email/mailtemplate/mail_contract2016.htm"),1)
    ret = f.ReadAll
    f.Close
    set f=Nothing
    set fs=Nothing

    buf=""
    for i=0 to ocontract.FResultCount - 1
    	buf=buf&"<tr>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).FContractName &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).FctrNo &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).FMakerid &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).getMajorSellplaceName &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).FcontractDate &"</td>"
    	if (isPreView) then
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;""><a target=""_blank"" href="""& ocontract.FITemList(i).getPdfDownLinkUrlAdm &"""><img src=""http://scm.10x10.co.kr/images/pdficon.gif"" style=""border:0;"" /></a></td>"
    	else
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;""><a target=""_blank"" href="""& ocontract.FITemList(i).getPdfDownLinkUrl &"""><img src=""http://scm.10x10.co.kr/images/pdficon.gif"" style=""border:0;"" /></a></td>"
        end if
    	buf=buf&"</tr>"&VBCRLF
    next
    ret = replace(ret,"$$CTRLIST$$",buf)

    buf=""
    if oMdInfoList.FResultCount>0 then
	buf=buf&"<tr>"
	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6"">"
	buf=buf&"<strong>* 담당엠디</strong><br />"
	    for i=0 to oMdInfoList.FResultCount-1
	    if (oMdInfoList.FItemList(i).isMaybeOffMD) then offmdCnt=offmdCnt+1

		buf=buf&"&nbsp;&nbsp;&nbsp;- "& oMdInfoList.FItemList(i).Fusername&"&nbsp;"&CHKIIF(oMdInfoList.FItemList(i).isMaybeOffMD,"&nbsp;(오프라인 담당)","")
		buf=buf&"<br />&nbsp;&nbsp;&nbsp;- tel : 02-554-2033 "& CHKIIF(oMdInfoList.FItemList(i).Fextension="","","(내선 "&oMdInfoList.FItemList(i).Fextension&")")&" "& CHKIIF(oMdInfoList.FItemList(i).Fdirect070="",""," / 직통 :"&oMdInfoList.FItemList(i).Fdirect070)
			if (oMdInfoList.FItemList(i).Fusermail<>"") then
		buf=buf&"<br />&nbsp;&nbsp;&nbsp;- 이메일 : <a href=""mailto:"""& oMdInfoList.FItemList(i).Fusermail &" style=""color:#333;"">"& oMdInfoList.FItemList(i).Fusermail &"</a>"
			end if
		buf=buf&"<br /><br />"
			next
		buf=buf&"</td>"
	buf=buf&"</tr>"&VBCRLF
	end if
	ret = replace(ret,"$$MDLIST$$",buf)


    buf=""
   
    buf="서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 협력사 계약서 담당자 앞"
   
    ret = replace(ret,"$$RECEIVEADDR$$",buf)


    makeCtrMailContents = ret
end function


function makeEcCtrMailContents(ocontract,oMdInfoList, isPreView,manageUrl)
    dim ret, buf
    dim fs,f, i, offmdCnt : offmdCnt=0
    set fs=Server.CreateObject("Scripting.FileSystemObject")
    set f=fs.OpenTextFile(Server.MapPath("/lib/email/mailtemplate/mail_contractEc2016.htm"),1)
    ret = f.ReadAll
    f.Close
    set f=Nothing
    set fs=Nothing

    buf=""
    for i=0 to ocontract.FResultCount - 1
    	buf=buf&"<tr>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).FContractName &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).FctrNo &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).FMakerid &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).getMajorSellplaceName &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"& ocontract.FITemList(i).FcontractDate &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;""><a href="""&partnerScmUrl&"/partner/company/contract/ctrContsBrand.asp?ctrKey="&ocontract.FITemList(i).FCtrKey&""" target=""_blank"">바로가기></td>"
       
    	buf=buf&"</tr>"&VBCRLF
    next
    ret = replace(ret,"$$CTRLIST$$",buf)

    buf=""
    if oMdInfoList.FResultCount>0 then
    	   for i=0 to oMdInfoList.FResultCount-1
	    if (oMdInfoList.FItemList(i).isMaybeOffMD) then offmdCnt=offmdCnt+1
	    	
		buf=buf&"<tr>"
	  	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">"&  oMdInfoList.FItemList(i).Fusername&"&nbsp;"&CHKIIF(oMdInfoList.FItemList(i).isMaybeOffMD,"&nbsp;(오프라인 담당)","")&"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;""> tel : 02-554-2033"& CHKIIF(oMdInfoList.FItemList(i).Fextension="","","(내선 "&oMdInfoList.FItemList(i).Fextension&")")&" "& CHKIIF(oMdInfoList.FItemList(i).Fdirect070="",""," / 직통 :"&oMdInfoList.FItemList(i).Fdirect070) &"</td>"
    	buf=buf&"<td style=""font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"">" 
    	 	if (oMdInfoList.FItemList(i).Fusermail<>"") then  
    	buf=buf&"<a href=""mailto:"""& oMdInfoList.FItemList(i).Fusermail &" style=""color:#333;"">"& oMdInfoList.FItemList(i).Fusermail &"</a>"&"</td>"    	
    		end if
		buf=buf&"</tr>"&VBCRLF
			next
		
	end if
	ret = replace(ret,"$$MDLIST$$",buf)


    buf=""
   
    buf="서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 협력사 계약서 담당자 앞"
   
    ret = replace(ret,"$$RECEIVEADDR$$",buf)


    makeEcCtrMailContents = ret
end function

''관련(동일업체) 브랜드 콤보Box
sub DrawSameGroupBrand(igroupid,imakerid,selboxname, addStr)
    dim sqlStr, id, socname, ret, i
    sqlStr ="select p.id, c.socname"
    sqlStr = sqlStr& " from db_partner.dbo.tbl_partner p"
    sqlStr = sqlStr& "  Join db_user.dbo.tbl_user_c c"
    sqlStr = sqlStr& "  on p.id=c.userid"
    sqlStr = sqlStr& "  and c.userdiv='02'"
    sqlStr = sqlStr& "  and p.userdiv='9999'"
    sqlStr = sqlStr& " where p.groupid='"&igroupid&"'"
    sqlStr = sqlStr& " and p.id<>'"&imakerid&"'"
    sqlStr = sqlStr& " and c.isusing='Y'"
    sqlStr = sqlStr& " order by p.id"

    rsget.Open sqlStr,dbget,1

    i=0
	if Not rsget.Eof then
	    do until rsget.Eof
	        id = rsget("id")
	        socname = db2html(rsget("socname"))
	        ret = ret&"<option value='"&id&"' >"&socname&" ["&id&"]"
	        rsget.moveNext
	        i=i+1
	    loop
    end if
    rsget.Close

    if (ret<>"") then
        ret = "<select name='"&selboxname&"' "&addStr&"><option value=''>관련브랜드 선택 ("&i&")"&ret&"</select>"
    end if

    response.write ret
end Sub

''관련(동일업체) 브랜드 콤보Box
sub DrawSameGroupBrandUpche(igroupid,imakerid,selboxname, addStr)
    dim sqlStr, id, socname, ret, i
    sqlStr ="select p.id, c.socname"
    sqlStr = sqlStr& " from db_partner.dbo.tbl_partner p"
    sqlStr = sqlStr& "  Join db_user.dbo.tbl_user_c c"
    sqlStr = sqlStr& "  on p.id=c.userid"
    sqlStr = sqlStr& "  and c.userdiv='02'"
    sqlStr = sqlStr& "  and p.userdiv='9999'"
    sqlStr = sqlStr& " where p.groupid='"&igroupid&"'"
    sqlStr = sqlStr& " and c.isusing='Y'"
    sqlStr = sqlStr& " order by p.id"

    rsget.Open sqlStr,dbget,1

    i=0
	if Not rsget.Eof then
	    do until rsget.Eof
	        id = rsget("id")
	        socname = db2html(rsget("socname"))
	        ret = ret&"<option value='"&id&"' "&CHKIIF(LCASE(imakerid)=LCASE(id),"selected","")&">"&socname&" ["&id&"]"
	        rsget.moveNext
	        i=i+1
	    loop
    end if
    rsget.Close

    if (ret<>"") then
        ret = "<select name='"&selboxname&"' "&addStr&"><option value=''>전체"&ret&"</select>"
    end if

    response.write ret
end Sub

Class CContractMDItem
    public Fusername
    public Fusermail
    public Finterphoneno
    public Fextension
    public Fdirect070
    public Fpart_name
    public Fposit_name

    public function isMaybeOffMD()
        isMaybeOffMD = InStr(Fpart_name,"오프")>0 or InStr(Fpart_name,"매장")>0
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class
Class CPartnerContractDetailTypeItem
    public FcontractType
    public FdetailKey
    public FdetailDesc
    public FDefaultValue

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPartnerContractTypeItem
    public FContractType
    public FContractName
    public FContractContents
    public Fisusing
    public Fregdate
    public Fsubtype
    public fonoffgubun

    public function getSubTypeName()
        if (Fsubtype=0) then
            getSubTypeName = "기본계약서"
        elseif (Fsubtype=5) then
            getSubTypeName = "부속합의서"
        elseif (Fsubtype=7) then
            getSubTypeName = "물품공급계약서"
        elseif (Fsubtype=9) then
            getSubTypeName = "기타"
        else
            getSubTypeName = Fsubtype
        end if
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPartnerContractDetailItem
    public FctrKey
    public FdetailKey
    public FdetailValue
    public FdetailDesc

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

''계약필요업체
Class CPartnerContractReqItem
    public Fgroupid
    public Fmakerid
    public Fcompany_name

    public FMsellcnt
    public FWsellcnt
    public FUsellcnt
    public FTTLsellcnt

    public FMjungsanSum
    public FWjungsanSum
    public FUjungsanSum
    public FTTLjungsanSum
    public FBrandRegdate
    public FHolddate
    public Fholdregid

    public function getBrandActiveMonth()
        getBrandActiveMonth = dateDiff("m",FBrandRegdate,now())
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

'' 부속합의서 item
Class CPartnerAddContractSubItem
    public FSeq
    public FctrSubKey
    public Fmaeipdiv
    public Fscmdefaultmargine
	public FctrKey
	public FctrState
	public Fsellplace
	public FsellplaceName
	public Fcontractmwdiv
	public Fcontractmargin

	public FctrNo
	public FcontractName
    public FcontractDate
		public Fenddate
		
    public FuseitemCnt
    public Fuseitemmargin
    public FsellitemCnt
    public Fsellitemmargin
    public FjungsanCnt
    public FjungsanSum

    public FMjshopid
    public FMjshopname
    public FMjmaeipdiv
    public FMjdefaultmargin

    public FMakerid

    public Fcontractdeliverytype
    public FcontractFreebeasongLimit
    public FcontractdeliverPay
    public Fdefaultdeliverytype
    public FdefaultFreebeasongLimit
    public FdefaultdeliverPay

	public FecCtrSeq
	public FecAUser
	public FecBUser
    public FsignType
    public FdocuSignId
    public FdocuSignUri
    public FdocuSignSenddate
	 
	
    public function isDisabledMWMargin()
        if (Fsellplace="ON") then
            if isNULL(Fmaeipdiv) then

            else
                if (Fcontractmwdiv<>Fmaeipdiv) then

                end if
            end if
        else
            if isNULL(Fmaeipdiv) and isNULL(FMjmaeipdiv) then
                isDisabledMWMargin = true
            end if
        end if

        if (Fcontractmargin<=0) or (Fcontractmargin>=100) then
            isDisabledMWMargin = true
        end if
    end function

    public function isreqCheckMW()
        if (Fsellplace="ON") then
            if isNULL(Fmaeipdiv) then

            else
                if (Fcontractmwdiv<>Fmaeipdiv) then
                    isreqCheckMW = true
                end if
            end if
        else
            if isNULL(Fmaeipdiv) then
                if (Fcontractmwdiv<>FMjmaeipdiv) then
                    isreqCheckMW = true
                end if
            else
                if (Fcontractmwdiv<>Fmaeipdiv) then
                    isreqCheckMW = true
                end if
            end if
        end if
    end function

    public function isreqCheckMargin()
        if (Fsellplace="ON") then
            if isNULL(Fmaeipdiv) then

            else
                if (Fcontractmargin<>Fscmdefaultmargine) then
                    isreqCheckMargin = true
                end if
            end if

            if (Fmaeipdiv=FMjmaeipdiv) then
                if (Fcontractmargin<>FMjdefaultmargin) then
                    isreqCheckMargin = true
                end if
            end if
        else
            if isNULL(Fmaeipdiv) then
                if (Fcontractmargin<>FMjdefaultmargin) then
                    isreqCheckMargin = true
                end if
            else
                if (Fcontractmargin<>Fscmdefaultmargine) then
                    isreqCheckMargin = true
                end if
           end if
        end if
    end function

    '' 판매처
    public function getSellplaceName()
        getSellplaceName = FsellplaceName

    end function

    '' 기본마진SCM
    public function getSCMDefaultmargineStr()
        if isNULL(Fscmdefaultmargine) then exit function
        if (Fscmdefaultmargine=0) then exit function
        getSCMDefaultmargineStr = Fscmdefaultmargine&"%"
    end function

    ''계약마진
    public function getContractMarginStr()
        if isNULL(Fcontractmargin) then exit function
        getContractMarginStr = Fcontractmargin&"%"
    end function

    ''계약형태(계약서)
    public function getContractMwDivStr()
        if isNULL(Fcontractmwdiv) then exit function
        getContractMwDivStr = fnMaeipdivName(Fcontractmwdiv)
    end function

    ''대표계약마진
    public function getMjContractMarginStr()
        if isNULL(FMjdefaultmargin) then exit function
        getMjContractMarginStr = FMjdefaultmargin&"%"
    end function

    ''배송비정책SCM
    public function getSCMDefaultDlvStr()
        if (FSeq=0) then Exit function
        if isNULL(Fmaeipdiv) then Exit function
        if (Fmaeipdiv<>"U") then Exit function

        dim buf
        if (FdefaultDeliveryType="9") then
            buf="업체<font color='red'>조건</font>"
        elseif (FdefaultDeliveryType="7") then
            buf="업체<font color='blue'>착불</font>"
        end if

        if (FdefaultFreeBeasongLimit<>0) then
            buf = buf & CHKIIF(buf="","","<br>")&ForMatNumber(FdefaultFreeBeasongLimit,0) & "미만"
        end if

        if (FdefaultDeliverPay<>0) then
            buf = buf & CHKIIF(buf="","(반품시) ","<br>")&ForMatNumber(FdefaultDeliverPay,0) & "원"
        end if

        getSCMDefaultDlvStr = buf

    end function

    ''계약배송비정책
    public function getContractDefaultDlvStr()

    end function

    ''신규추가시 기본 마진
    public function getAddDefaultMargin()
        if (Fscmdefaultmargine<=0) then
            getAddDefaultMargin = CLNG(Fuseitemmargin)
        else
            getAddDefaultMargin = Fscmdefaultmargine
        end if
    end function

'    public function getBigoStr()
'        dim bufStr
'        IF (Fmwdiv="U") and (FdefaultdeliveryType="7") then
'            bufStr = bufStr & "업체착불배송"
'        elseIF (Fmwdiv="U") and (FdefaultdeliveryType="9") then
'            bufStr = bufStr & "업체조건배송"
'            bufStr = bufStr & "<br>"&FormatNumber(FdefaultFreebeasongLimit,0)&"원미만"
'            bufStr = bufStr & "<br>배송료"&FormatNumber(Fdefaultdeliverpay,0)&"원"
'        else
'            bufStr = bufStr & ""
'        end if
'
'        getBigoStr = bufStr
'    end function

    public function GetSignType()
        Select Case FsignType
            Case "H" : GetSignType="수기계약"
            Case "U" : GetSignType="U+전자계약"
            Case "D" : GetSignType="DocuSign"
            Case Else : GetSignType=""
        end Select
    end Function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPartnerContractItem2013

    public FctrKey
    public FcontractType
    public Fgroupid
    public Fmakerid
    public FcompanyNo
    public FctrState
    public FctrNo
    public FregUserID
    public FsendUserid
    public FfinUserID
    public Fregdate
    public Fsenddate
    public Fconfirmdate
    public Ffinishdate
    public Fdeletedate
    public FcontractContents

    public FcontractName
    public FcontractSubCnt
    public FcontractDate
    public FcontractJungsanDate
    public FsubType
    public FsignType
    public FdocuSignId
    public FdocuSignUri
    public FdocuSignSenddate    

    public FcompanyName
    public FRegUserName
    public FSendUserName
    public FfinUserName

    public FMajorSellplace
    public FMajorSellplaceName
    public FMajorMwdiv
    public FMajorDefaultmargin

    public FB_UPCHENAME
    public FonplaceCnt
    public FoFplaceCnt

    public FcurrGroupid
    
    public Fenddate				
    public Fonbrandusing	
    public Foffbrandusing	
    
    public FecCtrSeq
    public FecAUser
    public FecBUser
    public FAcompany_no
    public FBcompany_no



    ''업체가 다운로드 하는 CASE
    public function getPdfDownLinkUrl()
        getPdfDownLinkUrl = getPdfDownLinkUrlAdm&"&chkcf=1"  ''업체가 다운로드시 업체확인체크 위한 플래그
    end function

    public function getPdfDownLinkUrlAdm()
        dim addparam
        addparam = "?ctrKey="&FctrKey
        addparam = addparam&"&gkey="&Fgroupid
        addparam = addparam&"&ctrNo="&FctrNo
        addparam = addparam&"&sTp="&FsubType
        addparam = addparam&"&cTp="&Fcontracttype 
        addparam = addparam&"&vTp="&"d"                                   ''뷰인지 다운로드인지. (d 다운로드,else 뷰)
        addparam = addparam&"&pTp="&CHKIIF(CPrvContract,"1","")           ''개인정보 수집존재
        addparam = addparam&"&ekey="&MD5(ChashVal&FctrKey&Fgroupid)
        addparam = addparam&"&dTp="&FsignType                               ''문서타입(D:docuSign,기타)
        if (application("Svr_Info")	= "Dev") then
            getPdfDownLinkUrlAdm = "https://testwebadmin.10x10.co.kr/admin/member/contract/dnContractPdf.asp"&addparam
        else
            getPdfDownLinkUrlAdm = "https://apps.10x10.co.kr:442/pdf/dnContractPdf.asp"&addparam
        end if
    end function

    public function IsCtrOpenValidState()
        IsCtrOpenValidState = (FctrState=0)
    end function

    public function getMajorSellplaceName()
        if isNULL(FMajorSellplaceName) then Exit function
        dim ret : ret=FMajorSellplaceName
        if (FMajorSellplaceName="ON") then
            ret="온라인"
            if (FoFplaceCnt>0) and (FcontractSubCnt>0) then
                ret = ret & " (외 "&CHKIIF(FcontractSubCnt>1,FcontractSubCnt-1,"")&")"
            end if
        else
            if (FcontractSubCnt>1) then
                ret = ret & " (외)"
            end if
        end if


        getMajorSellplaceName = ret
    end function

    public function getMajorMarginStr()
        if isNULL(FMajorMwdiv) then Exit function
        dim ret : ret=FMajorMwdiv


        ret = fnMaeipdivName(FMajorMwdiv)

        if Not isNULL(FMajorDefaultmargin) then
            ret = ret &" "& FMajorDefaultmargin & "%"
        end if

        if (FcontractSubCnt>1) then
            ret = ret &" (외"&FcontractSubCnt-1&"건)"
        end if

        getMajorMarginStr=ret
    end function

    public function IsDefaultContract() ''기본계약서 여부
        IsDefaultContract = (FsubType=0)
    end function

    public function GetStateActiondate()
        if isNULL(FCtrState) then Exit function
        Select Case FCtrState
            Case 0
                : GetStateActiondate = Fregdate
            Case 1
                : GetStateActiondate = FSendDate
            Case 3
                : GetStateActiondate = Fconfirmdate
            Case 7
                : GetStateActiondate = Ffinishdate
            Case -1
                : GetStateActiondate = Fdeletedate
           Case else
                : GetStateActiondate = ""
        end Select

    end function

    public function GetContractStateColor()
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

    public function GetContractStateName()
        dim buf
        Select Case FCtrState
            Case 0
                : buf = "수정중(미전송)"
            Case 1
                : buf = "계약오픈(검토대기)"
            Case 2
                : buf = "계약반려(검토반려)"    
            Case 3
                : buf = "계약확인(결재완료)"
            Case 6
                : buf = "서명진행"    
            Case 7
                : buf = "계약완료"
            Case 9
                : buf = "계약종료"    
            Case -1
                : buf = "삭제"
            Case else
                : buf = FContractState
        end Select

        GetContractStateName = "<font color='"&GetContractStateColor&"'>"&buf&"</font>"
    end function

    public function GetSignType()
        Select Case FsignType
            Case "H" : GetSignType="수기계약"
            Case "U" : GetSignType="U+전자계약"
            Case "D" : GetSignType="DocuSign"
            Case Else : GetSignType=""
        end Select
    end Function
    
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CPartnerContract
    public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FRectCateCode
	public FRectDispCateCode
	public FRectMakerid
	public FRectCompanyName
	public FRectManagerName
	public FRectContractType
	public FRectdetailKey
	public FRectContractID
	public FRectContractno
	public FRectContractState
	public FRectContractTypeGbn
	public FRectOnOffGubun
	
    public FRectCtrKey
    public FRectGroupID

    public FRectMdID
    public FRectCtrKeyArr
    public FRectRegScope
    public FRectGrpType
    public FRectNotIncboru

	public FRectSendUserID
	
	public FCtrState
	
	public FRectonusing
	public FRectoffusing
	public FRectctrenddate
	public FRectchkend
	
	public FRectctryyyy
	public FRectctrnum
	public FRectctrdtype
	
	public Fcontractname    
	public FctrNo 				     
	public Fregdate 			     
	public Fsenddate 		     
	public Fconfirmdate 	     
	public Ffinishdate 	     
	public Fenddate 			  
	public FRegUserName   
	public FSendUserName      
	public FcontractContents
    public Fgroupid
 	public Fcontracttype
    public FecCtrSeq
    public FcompanyNo
    public FctrKey
    public FsubType
    public FecBUser
       
	public FRectCState  
	public FRectCState1 
	public FRectCState2
	public FRectCState3 
	public FRectCState6 
	public FRectCState7 
	public FRectCState8 
	public FRectCState9 
	
    public sub GetCurrAddContractListONBrand()
        dim sqlStr, sqlStrAdd, i

        sqlStr = "db_partner.[dbo].[sp_Ten_partner_AddContract_CurrentList] ('"&FRectGroupid&"','"&FRectMakerid&"')"

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc


		FResultCount = rsget.RecordCount
        IF (FResultCount<1) then FResultCount=0
        FTotalCount = FResultCount
		redim preserve FItemList(FResultCount)
		i=0

		IF Not (rsget.EOF OR rsget.BOF) THEN
		do until rsget.eof
		    set FItemList(i)          = new CPartnerAddContractSubItem
		    FItemList(i).FSeq               = rsget("Seq")
            FItemList(i).Fmaeipdiv          = rsget("maeipdiv")
            FItemList(i).Fscmdefaultmargine = rsget("scmdefaultmargine")
            FItemList(i).FctrKey            = rsget("ctrKey")
            FItemList(i).FctrState          = rsget("ctrState")
            FItemList(i).Fsellplace         = rsget("sellplace")
            FItemList(i).FsellplaceName     = rsget("sellplaceName")
            FItemList(i).Fcontractmwdiv     = rsget("contractmwdiv")
            FItemList(i).Fcontractmargin    = rsget("contractmargin")
            FItemList(i).FctrNo             = rsget("ctrNo")
            FItemList(i).FcontractName      = rsget("contractName")
            FItemList(i).FcontractDate      = rsget("contractDate")
            FItemList(i).FendDate      			= rsget("endDate")

            FItemList(i).FuseitemCnt        = rsget("useitemCnt")
            FItemList(i).Fuseitemmargin     = rsget("useitemmargin")

            FItemList(i).FsellitemCnt       = rsget("sellitemCnt")
            FItemList(i).Fsellitemmargin    = rsget("sellitemmargin")

            FItemList(i).FjungsanCnt        = rsget("jungsanCnt")
            FItemList(i).FjungsanSum        = rsget("jungsanSum")

            FItemList(i).Fdefaultdeliverytype     = rsget("defaultdeliverytype")
            FItemList(i).FdefaultFreebeasongLimit = rsget("defaultFreebeasongLimit")
            FItemList(i).FdefaultdeliverPay       = rsget("defaultdeliverPay")

            FItemList(i).Fcontractdeliverytype     = rsget("contractdeliverytype")
            FItemList(i).FcontractFreebeasongLimit = rsget("contractFreebeasongLimit")
            FItemList(i).FcontractdeliverPay       = rsget("contractdeliverPay")

			 FItemList(i).FecCtrSeq               = rsget("ecCtrSeq")  
			 FItemList(i).FecAUser                = rsget("ecAUser")  
			 FItemList(i).FecBUser                = rsget("ecBUser")
             FItemList(i).FsignType               = rsget("signType") 
		
            if (i>0) and (FItemList(0).Fmaeipdiv<>FItemList(i).Fmaeipdiv) then
                if ((FItemList(0).FuseitemCnt+FItemList(i).FuseitemCnt)<>0) then
                    FItemList(0).Fuseitemmargin     = (FItemList(0).Fuseitemmargin*FItemList(0).FuseitemCnt+FItemList(i).Fuseitemmargin*FItemList(i).FuseitemCnt)/(FItemList(0).FuseitemCnt+FItemList(i).FuseitemCnt)
                end if
                FItemList(0).FuseitemCnt        = FItemList(0).FuseitemCnt + FItemList(i).FuseitemCnt

                if ((FItemList(0).FsellitemCnt+FItemList(i).FsellitemCnt)<>0) then
                    FItemList(0).Fsellitemmargin    = (FItemList(0).Fsellitemmargin*FItemList(0).FsellitemCnt+FItemList(i).Fsellitemmargin*FItemList(i).FsellitemCnt)/(FItemList(0).FsellitemCnt+FItemList(i).FsellitemCnt)
                end if
                FItemList(0).FsellitemCnt       = FItemList(0).FsellitemCnt + FItemList(i).FsellitemCnt

                FItemList(0).FjungsanCnt        = FItemList(0).FjungsanCnt + FItemList(i).FjungsanCnt
                FItemList(0).FjungsanSum        = FItemList(0).FjungsanSum + FItemList(i).FjungsanSum
            end if
		    i=i+1
			rsget.moveNext
		loop
		END IF
		rsget.close
    end sub

    public sub GetCurrAddContractListOFBrand()
        dim sqlStr
        sqlStr = "db_partner.[dbo].[sp_Ten_partner_AddContract_CurrentListOF] ('"&FRectGroupid&"','"&FRectMakerid&"')"

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc


		FResultCount = rsget.RecordCount
        IF (FResultCount<1) then FResultCount=0
        FTotalCount = FResultCount
		redim preserve FItemList(FResultCount)
		i=0

		IF Not (rsget.EOF OR rsget.BOF) THEN
		do until rsget.eof
		    set FItemList(i)          = new CPartnerAddContractSubItem
		    FItemList(i).FSeq               = rsget("Seq")
            FItemList(i).Fmaeipdiv          = rsget("maeipdiv")
            FItemList(i).Fscmdefaultmargine = rsget("scmdefaultmargine")

            FItemList(i).FctrKey            = rsget("ctrKey")
            FItemList(i).FctrState          = rsget("ctrState")
            FItemList(i).Fsellplace         = rsget("sellplace")
            FItemList(i).FsellplaceName     = rsget("sellplaceName")

            FItemList(i).Fcontractmwdiv     = rsget("contractmwdiv")
            FItemList(i).Fcontractmargin    = rsget("contractmargin")
            FItemList(i).FctrNo             = rsget("ctrNo")
            FItemList(i).FcontractName      = rsget("contractName")
            FItemList(i).FcontractDate      = rsget("contractDate")
						FItemList(i).FendDate      			= rsget("endDate")

            FItemList(i).FjungsanCnt        = rsget("jungsanCnt")
            FItemList(i).FjungsanSum        = rsget("jungsanSum")

            ''대표매장 마진
            FItemList(i).FMjshopid          = rsget("Mjshopid")
            FItemList(i).FMjshopname        = rsget("Mjshopname")
            FItemList(i).FMjmaeipdiv        = rsget("Mjmaeipdiv")
            FItemList(i).FMjdefaultmargin   = rsget("Mjdefaultmargin")

	 		 FItemList(i).FecCtrSeq               = rsget("ecCtrSeq")  
			 FItemList(i).FecAUser                = rsget("ecAUser")  
			 FItemList(i).FecBUser                = rsget("ecBUser")
             FItemList(i).FsignType               = rsget("signType")
			 
'            FItemList(i).FuseitemCnt        = rsget("useitemCnt")
'            FItemList(i).Fuseitemmargin     = rsget("useitemmargin")

'            FItemList(i).FsellitemCnt       = rsget("sellitemCnt")
'            FItemList(i).Fsellitemmargin    = rsget("sellitemmargin")

'            FItemList(i).Fdefaultdeliverytype     = rsget("defaultdeliverytype")
'            FItemList(i).FdefaultFreebeasongLimit = rsget("defaultFreebeasongLimit")
'            FItemList(i).FdefaultdeliverPay       = rsget("defaultdeliverPay")
'
'            FItemList(i).Fcontractdeliverytype     = rsget("contractdeliverytype")
'            FItemList(i).FcontractFreebeasongLimit = rsget("contractFreebeasongLimit")
'            FItemList(i).FcontractdeliverPay       = rsget("contractdeliverPay")

'            if ((FItemList(0).FuseitemCnt+FItemList(i).FuseitemCnt)<>0) then
'                FItemList(0).Fuseitemmargin     = (FItemList(0).Fuseitemmargin*FItemList(0).FuseitemCnt+FItemList(i).Fuseitemmargin*FItemList(i).FuseitemCnt)/(FItemList(0).FuseitemCnt+FItemList(i).FuseitemCnt)
'            end if
'            FItemList(0).FuseitemCnt        = FItemList(0).FuseitemCnt + FItemList(i).FuseitemCnt
'
'            if ((FItemList(0).FsellitemCnt+FItemList(i).FsellitemCnt)<>0) then
'                FItemList(0).Fsellitemmargin    = (FItemList(0).Fsellitemmargin*FItemList(0).FsellitemCnt+FItemList(i).Fsellitemmargin*FItemList(i).FsellitemCnt)/(FItemList(0).FsellitemCnt+FItemList(i).FsellitemCnt)
'            end if
'            FItemList(0).FsellitemCnt       = FItemList(0).FsellitemCnt + FItemList(i).FsellitemCnt
'
'            FItemList(0).FjungsanCnt        = FItemList(0).FjungsanCnt + FItemList(i).FjungsanCnt
'            FItemList(0).FjungsanSum        = FItemList(0).FjungsanSum + FItemList(i).FjungsanSum

		    i=i+1
			rsget.moveNext
		loop
		END IF
		rsget.close
    end sub

    public sub GetCurrAddContractListCheckMargin()
        dim sqlStr
        sqlStr = "db_partner.[dbo].[sp_Ten_partner_AddContract_CheckList] ('"&FRectGroupid&"')"

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc


		FResultCount = rsget.RecordCount
        IF (FResultCount<1) then FResultCount=0
        FTotalCount = FResultCount
		redim preserve FItemList(FResultCount)
		i=0

		IF Not (rsget.EOF OR rsget.BOF) THEN
		do until rsget.eof
		    set FItemList(i)          = new CPartnerAddContractSubItem
		    FItemList(i).FMakerid           = rsget("Makerid")
            FItemList(i).Fmaeipdiv          = rsget("scmmwdiv")
            FItemList(i).Fscmdefaultmargine = rsget("scmmargin")

            FItemList(i).Fsellplace         = rsget("sellplace")
            FItemList(i).FsellplaceName     = rsget("sellplaceName")

            FItemList(i).Fcontractmwdiv     = rsget("ctrmwdiv")
            FItemList(i).Fcontractmargin    = rsget("ctrmargin")



            ''대표매장 마진
            'FItemList(i).FMjshopid          = rsget("Mjshopid")
            'FItemList(i).FMjshopname        = rsget("Mjshopname")
            FItemList(i).FMjmaeipdiv        = rsget("Mjmaeipdiv")
            FItemList(i).FMjdefaultmargin   = rsget("Mjdefaultmargin")

            FItemList(i).FuseitemCnt        = rsget("useitemCnt")
            FItemList(i).Fuseitemmargin     = rsget("useitemmargin")

            FItemList(i).FsellitemCnt       = rsget("sellitemCnt")
            FItemList(i).Fsellitemmargin    = rsget("sellitemmargin")


            ''FItemList(i).FjungsanCnt        = rsget("jungsanCnt")
            ''FItemList(i).FjungsanSum        = rsget("jungsanSum")



'            FItemList(i).Fdefaultdeliverytype     = rsget("defaultdeliverytype")
'            FItemList(i).FdefaultFreebeasongLimit = rsget("defaultFreebeasongLimit")
'            FItemList(i).FdefaultdeliverPay       = rsget("defaultdeliverPay")
'
'            FItemList(i).Fcontractdeliverytype     = rsget("contractdeliverytype")
'            FItemList(i).FcontractFreebeasongLimit = rsget("contractFreebeasongLimit")
'            FItemList(i).FcontractdeliverPay       = rsget("contractdeliverPay")

'            if ((FItemList(0).FuseitemCnt+FItemList(i).FuseitemCnt)<>0) then
'                FItemList(0).Fuseitemmargin     = (FItemList(0).Fuseitemmargin*FItemList(0).FuseitemCnt+FItemList(i).Fuseitemmargin*FItemList(i).FuseitemCnt)/(FItemList(0).FuseitemCnt+FItemList(i).FuseitemCnt)
'            end if
'            FItemList(0).FuseitemCnt        = FItemList(0).FuseitemCnt + FItemList(i).FuseitemCnt
'
'            if ((FItemList(0).FsellitemCnt+FItemList(i).FsellitemCnt)<>0) then
'                FItemList(0).Fsellitemmargin    = (FItemList(0).Fsellitemmargin*FItemList(0).FsellitemCnt+FItemList(i).Fsellitemmargin*FItemList(i).FsellitemCnt)/(FItemList(0).FsellitemCnt+FItemList(i).FsellitemCnt)
'            end if
'            FItemList(0).FsellitemCnt       = FItemList(0).FsellitemCnt + FItemList(i).FsellitemCnt
'
'            FItemList(0).FjungsanCnt        = FItemList(0).FjungsanCnt + FItemList(i).FjungsanCnt
'            FItemList(0).FjungsanSum        = FItemList(0).FjungsanSum + FItemList(i).FjungsanSum

		    i=i+1
			rsget.moveNext
		loop
		END IF
		rsget.close
    end sub

    public function GetNewContractListReq(ireqCtr,jmonth)
        dim sqlStr
        if (reqCtr="OJN") or (reqCtr="OJNN") then
            sqlStr = "db_partner.[dbo].[sp_Ten_partner_AddContract_Require_NoJungsan]('"&reqCtr&"',"&FPagesize&","&jmonth&",'"&FRectMakerid&"','"&FRectGroupID&"','"&FRectCateCode&"','"&FRectDispCateCode&"',"&CHKIIF(FRectNotIncboru<>"","1","0")&")"
        else
            sqlStr = "db_partner.[dbo].[sp_Ten_partner_AddContract_Require]('"&reqCtr&"',"&FPagesize&","&jmonth&",'"&FRectMakerid&"','"&FRectGroupID&"','"&FRectCateCode&"','"&FRectDispCateCode&"',"&CHKIIF(FRectNotIncboru<>"","1","0")&")"
        end if
''rw sqlStr
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc


		FResultCount = rsget.RecordCount
        IF (FResultCount<1) then FResultCount=0
        FTotalCount = FResultCount
		redim preserve FItemList(FResultCount)
		i=0

		IF Not (rsget.EOF OR rsget.BOF) THEN
		do until rsget.eof
            set FItemList(i)          = new CPartnerContractReqItem

            FItemList(i).FGroupid           = rsget("Groupid")
            FItemList(i).FMakerid           = rsget("Makerid")

            FItemList(i).Fcompany_name      = db2html(rsget("company_name"))

            FItemList(i).FMsellcnt          = rsget("sellMCnt")
            FItemList(i).FWsellcnt          = rsget("sellWCnt")

            FItemList(i).FTTLsellcnt        = rsget("sellitemCnt")

            FItemList(i).FMjungsanSum       = rsget("me_totalsuplycash")
            FItemList(i).FWjungsanSum       = rsget("wi_totalsuplycash")


'            if (reqCtr<>"OJ") and (reqCtr<>"OT") then
'                FItemList(i).FUsellcnt          = 0
'                FItemList(i).FUjungsanSum       = 0
'            else
                FItemList(i).FUsellcnt          = rsget("sellUCnt")
                FItemList(i).FUjungsanSum       = rsget("ub_totalsuplycash")
'            end if

            FItemList(i).FTTLjungsanSum     = FItemList(i).FMjungsanSum+FItemList(i).FWjungsanSum+FItemList(i).FUjungsanSum
            FItemList(i).FBrandRegdate      = rsget("BrandRegdate")
            FItemList(i).FHolddate      = rsget("Holddate")
            FItemList(i).Fholdregid     = rsget("holdregid")
            i=i+1
			rsget.moveNext
		loop
		END IF
		rsget.close
    end function

    public Sub GetNewContractListUpcheView()
        dim sqlStr, sqlStrAdd, i

        sqlStrAdd = sqlStrAdd & " and M.ctrState>=1"
        sqlStrAdd = sqlStrAdd & " and M.groupid='"&FRectGroupID&"'"

'        if (FRectContractState<>"") then
'	        if (FRectContractState="M") then ''미완료
'	            sqlStrAdd = sqlStrAdd & " and (M.ctrState>=0 or M.ctrState<7)"
'	        else
'    	        sqlStrAdd = sqlStrAdd & " and M.ctrState=" & FRectContractState
'    	    end if
'	    end if
dim sqlStrAddsb : sqlStrAddsb = ""
      if FRectCState1 ="1" or FRectCState2="1" or FRectCState3 ="1"  or  FRectCState6 ="1" or  FRectCState7 ="1" or  FRectCState8 ="1" or  FRectCState9 ="1" then
     	
     	  	if  FRectCState1 ="1" then
     	  		if sqlStrAddsb <>"" then 
     	  			sqlStrAddsb = sqlStrAddsb & " or "
     		  	end if
		     		sqlStrAddsb = sqlStrAddsb & " ctrstate =1 "
     		end if 
     		if  FRectCState2 ="1" then
     	  		if sqlStrAddsb <>"" then 
     	  			sqlStrAddsb = sqlStrAddsb & " or "
     		  	end if
		     		sqlStrAddsb = sqlStrAddsb & " ctrstate =2 "
     		end if 
     		
     		if  FRectCState3 ="1" then
     			if sqlStrAddsb <>"" then 
     	  			sqlStrAddsb = sqlStrAddsb & " or "
     	   		end if
     				sqlStrAddsb = sqlStrAddsb & "   ctrstate =3 "
     		end if 
     		if  FRectCState6 ="1" then
     			if sqlStrAddsb <>"" then 
     	  			sqlStrAddsb = sqlStrAddsb & " or "
     	   		end if
     				sqlStrAddsb = sqlStrAddsb & "  ctrstate =6 "
     		end if 
     		if  FRectCState7 ="1" then
     			if sqlStrAddsb <>"" then 
     	  			sqlStrAddsb = sqlStrAddsb & " or "
     	   end if
     				sqlStrAddsb = sqlStrAddsb & "   ctrstate =7 "
     		end if 
     		if  FRectCState8 ="1" then
     			if sqlStrAddsb <>"" then 
     	  			sqlStrAddsb = sqlStrAddsb & " or "
     	   		end if
     				sqlStrAddsb = sqlStrAddsb & "   ctrstate =8 "
     		end if 
     		if  FRectCState9 ="1" then
     			if sqlStrAddsb <>"" then 
     	  			sqlStrAddsb = sqlStrAddsb & " or "
     	   		end if
     			sqlStrAddsb = sqlStrAddsb & "  ctrstate =9 "
     		end if 
     		  sqlStrAdd = sqlStrAdd & " and ( "
     		  sqlStrAdd = sqlStrAdd&sqlStrAddsb
     	sqlStrAdd = sqlStrAdd & " )"	
    end if 

	    if FRectMakerid<>"" then
	        ''sqlStrAdd = sqlStrAdd & " and (M.makerid='' or makerid='"&FRectMakerid&"')"
	        sqlStrAdd = sqlStrAdd & " and makerid='"&FRectMakerid&"'"
	    end if

	    if (FRectCtrKeyArr<>"") then
            sqlStrAdd = sqlStrAdd & " and M.ctrKey in ("&FRectCtrKeyArr&")"
        end if

        sqlStr = "select count(M.ctrKey) as cnt "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master M"
	    sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_contractType T"
	    sqlStr = sqlStr & "     on T.contractType=M.contractType"
	    if (FRectGrpType="M") then
	        sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner_ctr_sub Sb"
	        sqlStr = sqlStr & " on M.ctrKey=sb.ctrKey"
	    end if
	    sqlStr = sqlStr & "     left join db_user.dbo.tbl_user_c c"
	    sqlStr = sqlStr & "     on M.makerid=c.userid"
	    sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & sqlStrAdd

	    rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

        if (FTotalCount<1) then
            Exit Sub
        end if

		sqlStr = "select top " & CStr(FPageSize*FCurrpage)
		sqlStr = sqlStr & " m.ctrKey,m.groupid,m.makerid,m.ctrState,m.ctrNo,m.regUserID,m.sendUserid,m.finUserID,m.regdate,m.senddate,m.confirmdate,m.finishdate,m.deletedate"
		sqlStr = sqlStr & " ,g.company_no,g.company_name"
		sqlStr = sqlStr & " ,T.contractName, T.subType "
		if (FRectGrpType="M") then
		    sqlStr = sqlStr & " , sb.sellplace as MajorSellplace"
		    sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey and S.sellplace='ON') as onplaceCnt"
		    sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey and S.sellplace<>'ON') as oFplaceCnt"
		    sqlStr = sqlStr & " ,isNULL(su.shopname,Sb.sellplace) as MajorSellplaceName"
		    sqlStr = sqlStr & " ,Sb.mwdiv as MajorMwdiv"
		    sqlStr = sqlStr & " ,Sb.defaultmargin as MajorDefaultmargin"
		    sqlStr = sqlStr & " ,0 as contractSubCnt"
		else
    		sqlStr = sqlStr & " ,(select top 1 sellplace from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey order by ctrSubKey) as MajorSellplace"
    		sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey and S.sellplace='ON') as onplaceCnt"
    		sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey and S.sellplace<>'ON') as oFplaceCnt"
    		sqlStr = sqlStr & " ,(select top 1 isNULL(u.shopname,s.sellplace) from db_partner.dbo.tbl_partner_ctr_sub S left join db_shop.dbo.tbl_shop_user U on S.sellplace=U.userid where S.ctrKey=M.ctrKey order by ctrSubKey) as MajorSellplaceName"
    		sqlStr = sqlStr & " ,(select top 1 mwdiv from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey order by ctrSubKey) as MajorMwdiv"
    		sqlStr = sqlStr & " ,(select top 1 defaultmargin from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey order by ctrSubKey) as MajorDefaultmargin"
    		sqlStr = sqlStr & " ,(select count(*) as CNT from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey) as contractSubCnt"
	    end if
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$CONTRACT_DATE$$') as contractDate"
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$DEFAULT_JUNGSANDATE$$') as contractJungsanDate"
		sqlStr = sqlStr & " ,U.userName as regUserName, U2.userName as SendUserName, U3.userName as finUserName"
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master M"
	    sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_contractType T"
	    sqlStr = sqlStr & "     on T.contractType=M.contractType"
	    if (FRectGrpType="M") then
	        sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner_ctr_sub Sb"
	        sqlStr = sqlStr & " on M.ctrKey=sb.ctrKey"
	        sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user sU "
	        sqlStr = sqlStr & " on Sb.sellplace=sU.userid "
	    end if
	    sqlStr = sqlStr & "     left join db_user.dbo.tbl_user_c c"
	    sqlStr = sqlStr & "     on M.makerid=c.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner_group G"
	    sqlStr = sqlStr & "     on M.groupid=G.groupid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U"
	    sqlStr = sqlStr & "     on M.regUserID=U.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U2"
	    sqlStr = sqlStr & "     on M.sendUserid=U2.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U3"
	    sqlStr = sqlStr & "     on M.finUserid=U3.userid"

	    sqlStr = sqlStr & " where 1=1"
	    sqlStr = sqlStr & sqlStrAdd
		sqlStr = sqlStr & " order by convert(varchar(7),m.regdate,121) desc, m.contractType asc, contractDate desc, m.ctrKey desc "
 
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerContractItem2013

				FItemList(i).FctrKey  = rsget("ctrKey")
                FItemList(i).Fgroupid     = rsget("groupid")
                FItemList(i).Fmakerid     = rsget("makerid")
                FItemList(i).FctrState    = rsget("ctrState")
                FItemList(i).FctrNo       = rsget("ctrNo")
                FItemList(i).FregUserID   = rsget("regUserID")
                FItemList(i).FsendUserid  = rsget("sendUserid")
                FItemList(i).FfinUserID   = rsget("finUserID")
                FItemList(i).Fregdate     = rsget("regdate")
                FItemList(i).Fsenddate    = rsget("senddate")
                FItemList(i).Fconfirmdate = rsget("confirmdate")
                FItemList(i).Ffinishdate  = rsget("finishdate")
                FItemList(i).Fdeletedate  = rsget("deletedate")

                FItemList(i).FcompanyNo   = rsget("company_no")
                FItemList(i).FcompanyName = db2html(rsget("company_name"))

                FItemList(i).FcontractName = rsget("contractName")
                FItemList(i).FcontractSubCnt = rsget("contractSubCnt")
                FItemList(i).FcontractDate = rsget("contractDate")
                FItemList(i).FcontractJungsanDate = rsget("contractJungsanDate")

                FItemList(i).FRegUserName = rsget("regUserName")
                FItemList(i).FSendUserName = rsget("SendUserName")
                FItemList(i).FfinUserName  = rsget("finUserName")
                FItemList(i).FMajorSellplace        = rsget("MajorSellplace")
                FItemList(i).FMajorSellplaceName    = rsget("MajorSellplaceName")
                FItemList(i).FMajorMwdiv            = rsget("MajorMwdiv")
                FItemList(i).FMajorDefaultmargin    = rsget("MajorDefaultmargin")

                FItemList(i).FonplaceCnt    = rsget("onplaceCnt")
                FItemList(i).FoFplaceCnt    = rsget("oFplaceCnt")
                FItemList(i).FsubType       = rsget("subType")
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub
 
	public Function fnGetContractInfoUpcheview()
		dim strSql
		strSql = " select m.ctrKey, m.groupid, m.contracttype, t.contractName, m.ctrNo, m.ctrstate, m.regdate ,m.senddate, m.confirmdate, m.finishdate"
		strSql =strSql & ", m.enddate , u1.username as regusername, u2.username as sendusername , m.contractContents, g.company_no , m.ecCtrSeq, t.subtype,ecBUser "
		strSql = strSql & " from db_partner.dbo.tbl_partner_ctr_master as m "
		strSql = strSql & " inner join db_partner.dbo.tbl_partner_contractType as t on m.contractType = t.contractType "
	   strSql = strSql & "  inner join db_partner.dbo.tbl_partner_group G"
	   strSql = strSql & "     on M.groupid=G.groupid"
		strSql = strSql & "  inner join db_partner.dbo.tbl_user_tenbyten as u1 on reguserid = u1.userid  "
		strSql = strSql & "  left outer  join db_partner.dbo.tbl_user_tenbyten as u2 on senduserid = u2.userid  "
		strSql = strSql & " where m.ctrkey = " &FRectCtrKey
	
		rsget.Open strSql,dbget,1
		IF not rsget.eof then	
			FctrKey				= rsget("ctrKey")
			  Fgroupid			= rsget("groupid")		
			  Fcontracttype  = rsget("contracttype")
				Fcontractname = rsget("contractname") 
				FctrNo 				= rsget("ctrNo")
				Fctrstate         = rsget("ctrstate")
				Fregdate 			= rsget("regdate")
				Fsenddate 		= rsget("senddate")
				Fconfirmdate 	= rsget("confirmdate")
				Ffinishdate   	= rsget("finishdate")
				Fenddate 			= rsget("enddate") 
				Fregusername	   = rsget("regusername")
				FSendUserName = rsget("sendusername")
				FcontractContents = rsget("contractContents")
			   FecCtrSeq		 = rsget("ecCtrSeq")
			   FcompanyNo	 = rsget("company_no")
			   FsubType        =rsget("subtype")
			   FecBUser        =rsget("ecBUser")
		end IF
		rsget.close
	End Function

	public Sub GetNewContractList()
	    dim sqlStr, sqlStrAdd, i
	    dim ctrsdate, ctredate

        if (FRectContractState<>"-1") then      ''//2016/05/04추가
        sqlStrAdd = sqlStrAdd & " and M.ctrState>=0"
        end if
        
        if (FRectContractTypeGbn<>"") then  ''기본계약서 구분 D:기본,A:추가
            if (FRectContractTypeGbn="D") then
                sqlStrAdd = sqlStrAdd & " and T.subtype=0"
            elseif  (FRectContractTypeGbn="A") then
                sqlStrAdd = sqlStrAdd & " and T.subtype<>0"
            end if
        end if

        if (FRectContractType<>"") then
            sqlStrAdd = sqlStrAdd & " and M.ContractType="&FRectContractType&""
        end if

        if (FsubType<>"") then
            sqlStrAdd = sqlStrAdd & " and T.subtype="&FsubType&""
        end if

        if (FRectContractState<>"") then
	        if (FRectContractState="M") then ''미완료
	            sqlStrAdd = sqlStrAdd & " and (M.ctrState>=0 and M.ctrState<7)"
	        else
    	        sqlStrAdd = sqlStrAdd & " and M.ctrState=" & FRectContractState
    	    end if
	    end if

	    if FRectMakerid<>"" then
	        sqlStrAdd = sqlStrAdd & " and (M.groupid in (select groupid from db_partner.dbo.tbl_partner where id='" & FRectMakerid & "')"
	        sqlStrAdd = sqlStrAdd & "  or (M.makerid='"&FRectMakerid&"'))"
	    end if

	    if FRectContractno<>"" then
	        sqlStrAdd = sqlStrAdd & " and M.ctrNo='" & FRectContractno & "'"
	    end if

        if (FRectDispCateCode<>"") then
	        sqlStrAdd = sqlStrAdd & " and c.standardcatecode='" & FRectDispCateCode & "'"
        end if

	    if FRectCateCode<>"" then
	        sqlStrAdd = sqlStrAdd & " and c.catecode='" & FRectCateCode & "'"
	    end if

	    if FRectCompanyName<>"" then
	        sqlStrAdd = sqlStrAdd & " and M.groupid in ("
	        sqlStrAdd = sqlStrAdd & "     select groupid from db_partner.[dbo].tbl_partner_group"
	        sqlStrAdd = sqlStrAdd & "     where company_name like '%" & FRectCompanyName & "%'"
	        sqlStrAdd = sqlStrAdd & "     or replace(company_no,'-','')='"&replace(FRectCompanyName,"-","")&"'"
            sqlStrAdd = sqlStrAdd & " )"
	    end if

        if (FRectGroupID<>"") then
            sqlStrAdd = sqlStrAdd & " and M.groupid='"&FRectGroupID&"'"
        end if

        if (FRectCtrKeyArr<>"") then
            sqlStrAdd = sqlStrAdd & " and M.ctrKey in ("&FRectCtrKeyArr&")"
        end if

        if (FRectRegScope<>"") then
            if (FRectRegScope="R") then
                sqlStrAdd = sqlStrAdd & " and M.regUserID='"&session("ssBctID")&"'"
            elseif (FRectRegScope="S") then
                sqlStrAdd = sqlStrAdd & " and M.sendUserID='"&session("ssBctID")&"'"
            elseif (FRectRegScope="F") then
                sqlStrAdd = sqlStrAdd & " and M.finUserID='"&session("ssBctID")&"'"
            end if
        end if

		if (FRectSendUserID <> "") then
			if (FRectSendUserID = "xxxxxx") then
				sqlStrAdd = sqlStrAdd & " and M.sendUserID not in ("
				sqlStrAdd = sqlStrAdd & " 	select userid "
				sqlStrAdd = sqlStrAdd & " 	from db_partner.dbo.tbl_user_tenbyten "
				sqlStrAdd = sqlStrAdd & " 	where isusing = 1 "

                ' 퇴사예정자 처리	' 2018.10.16 한용민
                sqlStrAdd = sqlStrAdd & "	and (IsNull(ut.statediv, 'Y') ='Y' or (IsNull(ut.statediv, 'Y') ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
				sqlStrAdd = sqlStrAdd & " )"
			else
				sqlStrAdd = sqlStrAdd & " and M.sendUserID='" & FRectSendUserID & "'"
			end if
		end if
		
		if (FRectonUsing <> "") then
			sqlStrAdd = sqlStrAdd & " and c.isusing ='"&FRectonUsing&"'"
		end if
		
		if (FRectoffUsing <> "") then
			sqlStrAdd = sqlStrAdd & " and c.isoffusing ='"&FRectoffUsing&"'"
		end if

		if FRectchkend = "1" then
				sqlstrAdd = sqlStrAdd & " and m.enddate < '"&FRectctrenddate&"' "
	end if
'    if(FRectctrdtype = "1") then
'    	if FRectctrnum=1 then
'    		ctrsdate = FRectctryyyy&"-01-01" 
'    		ctredate = FRectctryyyy&"-04-01" 
'    	elseif FRectctrnum=2 then
'    		ctrsdate = FRectctryyyy&"-04-01" 
'    		ctredate = FRectctryyyy&"-07-01" 
'    	elseif FRectctrnum=3 then
'    		ctrsdate = FRectctryyyy&"-07-01" 
'    		ctredate = FRectctryyyy&"-10-01" 
'    	elseif FRectctrnum=4 then
'    		ctrsdate = FRectctryyyy&"-10-01" 
'    		ctredate = dateadd("d",1,FRectctryyyy&"-12-31") 
'    	end if
'    	
'    	sqlstrAdd = sqlStrAdd & " and m.regdate >='"&ctrsdate&"' and m.regdate < '"&ctredate&"'"
'    end if


  	
	    sqlStr = "select count(M.ctrKey) as cnt "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master M"
	    sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_contractType T"
	    sqlStr = sqlStr & "     on T.contractType=M.contractType"
	    if (FRectGrpType="M") then
	        sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner_ctr_sub Sb"
	        sqlStr = sqlStr & " on M.ctrKey=sb.ctrKey"
	    end if
	    sqlStr = sqlStr & "     left join db_user.dbo.tbl_user_c c"
	    sqlStr = sqlStr & "     on M.makerid=c.userid"
	    sqlStr = sqlStr & " where  m.contracttype not in (8,9,10,16,17,18)  "
        sqlStr = sqlStr & sqlStrAdd

	    rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

        if (FTotalCount<1) then
            Exit Sub
        end if

		sqlStr = "select top " & CStr(FPageSize*FCurrpage)
		sqlStr = sqlStr & " m.ctrKey,m.groupid,m.makerid,m.ctrState,m.ctrNo,m.regUserID,m.sendUserid,m.finUserID,m.regdate,m.senddate,m.confirmdate,m.finishdate,m.deletedate"
		sqlStr = sqlStr & " ,g.company_no,g.company_name"
		sqlStr = sqlStr & " ,isNULL(p.groupid,m.groupid) as currGroupid"
		sqlStr = sqlStr & " ,T.contractName, T.subType, m.signType "
		if (FRectGrpType="M") then
		    sqlStr = sqlStr & " , sb.sellplace as MajorSellplace"
		    sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey and S.sellplace='ON') as onplaceCnt"
		    sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey and S.sellplace<>'ON') as oFplaceCnt"
		    sqlStr = sqlStr & " ,isNULL(su.shopname,Sb.sellplace) as MajorSellplaceName"
		    sqlStr = sqlStr & " ,Sb.mwdiv as MajorMwdiv"
		    sqlStr = sqlStr & " ,Sb.defaultmargin as MajorDefaultmargin"
		    sqlStr = sqlStr & " ,0 as contractSubCnt"
		else
    		sqlStr = sqlStr & " ,(select top 1 sellplace from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey order by ctrSubKey) as MajorSellplace"
    		sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey and S.sellplace='ON') as onplaceCnt"
    		sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey and S.sellplace<>'ON') as oFplaceCnt"
    		sqlStr = sqlStr & " ,(select top 1 isNULL(u.shopname,s.sellplace) from db_partner.dbo.tbl_partner_ctr_sub S left join db_shop.dbo.tbl_shop_user U on S.sellplace=U.userid where S.ctrKey=M.ctrKey order by ctrSubKey) as MajorSellplaceName"
    		sqlStr = sqlStr & " ,(select top 1 mwdiv from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey order by ctrSubKey) as MajorMwdiv"
    		sqlStr = sqlStr & " ,(select top 1 defaultmargin from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey order by ctrSubKey) as MajorDefaultmargin"
    		sqlStr = sqlStr & " ,(select count(*) as CNT from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey) as contractSubCnt"
	    end if
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$CONTRACT_DATE$$') as contractDate"
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$DEFAULT_JUNGSANDATE$$') as contractJungsanDate"
		sqlStr = sqlStr & " ,U.userName as regUserName, U2.userName as SendUserName, U3.userName as finUserName, M.enddate, c.isusing as onbrandusing, c.isoffusing as offbrandusing "
		sqlStr = sqlStr & " ,M.ecCtrSeq, M.ecAUser, M.ecBUser "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master M"
	    sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_contractType T"
	    sqlStr = sqlStr & "     on T.contractType=M.contractType"
	    if (FRectGrpType="M") then
	        sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner_ctr_sub Sb"
	        sqlStr = sqlStr & " on M.ctrKey=sb.ctrKey"
	        sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user sU "
	        sqlStr = sqlStr & " on Sb.sellplace=sU.userid "
	    end if
	    sqlStr = sqlStr & "     left join db_user.dbo.tbl_user_c c"
	    sqlStr = sqlStr & "     on M.makerid=c.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner_group G"
	    sqlStr = sqlStr & "     on M.groupid=G.groupid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U"
	    sqlStr = sqlStr & "     on M.regUserID=U.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U2"
	    sqlStr = sqlStr & "     on M.sendUserid=U2.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U3"
	    sqlStr = sqlStr & "     on M.finUserid=U3.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "     on M.makerid=p.id"

	    sqlStr = sqlStr & " where m.contracttype not in (8,9,10,16,17,18) "
	    sqlStr = sqlStr & sqlStrAdd

	    if (FRectGroupID<>"") or (FRectMakerid<>"") then
		    ''sqlStr = sqlStr & " order by m.contractType asc, m.contractDate desc, m.ctrKey desc "
		    sqlStr = sqlStr & " order by m.contractType asc, m.ctrKey desc "  ''느려서 바꿈 2014/12/10
        else
            sqlStr = sqlStr & " order by m.ctrKey desc "
        end if
''rw FRectGrpType
' rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerContractItem2013

				FItemList(i).FctrKey  = rsget("ctrKey")
                FItemList(i).Fgroupid     = rsget("groupid")
                FItemList(i).Fmakerid     = rsget("makerid")
                FItemList(i).FctrState    = rsget("ctrState")
                FItemList(i).FctrNo       = rsget("ctrNo")
                FItemList(i).FregUserID   = rsget("regUserID")
                FItemList(i).FsendUserid  = rsget("sendUserid")
                FItemList(i).FfinUserID   = rsget("finUserID")
                FItemList(i).Fregdate     = rsget("regdate")
                FItemList(i).Fsenddate    = rsget("senddate")
                FItemList(i).Fconfirmdate = rsget("confirmdate")
                FItemList(i).Ffinishdate  = rsget("finishdate")
                FItemList(i).Fdeletedate  = rsget("deletedate")

                FItemList(i).FcompanyNo   = rsget("company_no")
                FItemList(i).FcompanyName = db2html(rsget("company_name"))

                FItemList(i).FcontractName = rsget("contractName")
                FItemList(i).FcontractSubCnt = rsget("contractSubCnt")
                FItemList(i).FcontractDate = rsget("contractDate")
                FItemList(i).FcontractJungsanDate = rsget("contractJungsanDate")

                FItemList(i).FRegUserName = rsget("regUserName")
                FItemList(i).FSendUserName = rsget("SendUserName")
                FItemList(i).FfinUserName  = rsget("finUserName")
                FItemList(i).FMajorSellplace        = rsget("MajorSellplace")
                FItemList(i).FMajorSellplaceName    = rsget("MajorSellplaceName")
                FItemList(i).FMajorMwdiv            = rsget("MajorMwdiv")
                FItemList(i).FMajorDefaultmargin    = rsget("MajorDefaultmargin")

                FItemList(i).FonplaceCnt    = rsget("onplaceCnt")
                FItemList(i).FoFplaceCnt    = rsget("oFplaceCnt")
                FItemList(i).FsubType       = rsget("subType")
                FItemList(i).FsignType       = rsget("signType")
                FItemList(i).FcurrGroupid   = rsget("currGroupid")
                FItemList(i).Fenddate				= rsget("enddate")
                FItemList(i).Fonbrandusing	= rsget("onbrandusing")
                FItemList(i).Foffbrandusing	= rsget("offbrandusing")
                
                FItemList(i).FecCtrSeq			= rsget("ecCtrSeq")
                FItemList(i).FecAUser 			= rsget("ecAUser")
                FItemList(i).FecBUser 			= rsget("ecBUser")
                
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub

    public function getContractEmailMdList(isFreView)
        dim sqlStr
        dim isOffContractExists : isOffContractExists = false
        dim offMajorMd : offMajorMd = "'john6136'" ''일단 이요한 대리로 // 오픈시에만 필요할듯.
		 dim i
        ''오픈시에만 체크 / 차후에는 없어도?
        sqlStr = " select count(*) as CNT "
        sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master M"
        sqlStr = sqlStr & " 	Join db_partner.dbo.tbl_partner_ctr_Sub S"
        sqlStr = sqlStr & " 	on M.ctrKey=S.ctrKey"
        sqlStr = sqlStr & " where ctrState="&FRectContractState
        sqlStr = sqlStr & " and groupid='"&FRectGroupID&"'"
        if (FRectCtrKeyArr<>"") then
            sqlStr = sqlStr & " and M.ctrKey in ("&FRectCtrKeyArr&")"
        end if
        sqlStr = sqlStr & " and sellplace<>'ON'"
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            if (rsget("CNT")<1) then
                offMajorMd="''"
            end if
        end if
        rsget.Close

        if (isFreView) then
            sqlStr = " select t.username, isNULL(t.usermail,'') as usermail, isNULL(t.interphoneno,'') as interphoneno"
            sqlStr = sqlStr & " , isNULL(t.extension,'') as extension, isNULL(t.direct070,'') as direct070"
            sqlStr = sqlStr & " , isNULL(pt.part_name,'') as part_name, isNULL(po.posit_name,'') as posit_name"
            sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
            sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partinfo pt"
        	sqlStr = sqlStr & "     on t.part_sn=pt.part_sn"
        	sqlStr = sqlStr & "     left join db_partner.dbo.tbl_positinfo po"
        	sqlStr = sqlStr & "     on t.posit_sn=po.posit_sn"
            sqlStr = sqlStr & " where t.userid in ('"&FRectMdId&"',"&offMajorMd&")"

            ' 퇴사예정자 처리	' 2018.10.16 한용민
            sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
            sqlStr = sqlStr & " and t.userid<>''"
        else
            sqlStr = " select t.username, isNULL(t.usermail,'') as usermail, isNULL(t.interphoneno,'') as interphoneno"
            sqlStr = sqlStr & " , isNULL(t.extension,'') as extension, isNULL(t.direct070,'') as direct070"
            sqlStr = sqlStr & " , isNULL(pt.part_name,'') as part_name, isNULL(po.posit_name,'') as posit_name, count(*) as CNT"
            sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
            sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_ctr_master M"
            sqlStr = sqlStr & "     on M.ctrState="&FRectContractState
            sqlStr = sqlStr & "     and M.groupid='"&FRectGroupID&"'"
            if (FRectCtrKeyArr<>"") then
                sqlStr = sqlStr & " and M.ctrKey in ("&FRectCtrKeyArr&")"
            end if
            if (offMajorMd<>"") then
                sqlStr = sqlStr & "     and ((t.userid=M.sendUserID) or (t.userid in ("&offMajorMd&")))"
            else
                sqlStr = sqlStr & "     and t.userid=M.sendUserID"
            end if
            sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partinfo pt"
        	sqlStr = sqlStr & "     on t.part_sn=pt.part_sn"
        	sqlStr = sqlStr & "     left join db_partner.dbo.tbl_positinfo po"
        	sqlStr = sqlStr & "     on t.posit_sn=po.posit_sn"
            sqlStr = sqlStr & " where t.userid<>''"

            ' 퇴사예정자 처리	' 2018.10.16 한용민
            sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf            
            sqlStr = sqlStr & " group by t.username, t.usermail, t.interphoneno, t.extension,t.direct070 , pt.part_name, po.posit_name"
            sqlStr = sqlStr & " order by CNT desc"

        end if

        rsget.Open sqlStr,dbget,1
        FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			i=0
			do until rsget.eof
                set FItemList(i) = new CContractMDItem
                FItemList(i).Fusername            = rsget("username")
                FItemList(i).Fusermail            = rsget("usermail")
                FItemList(i).Finterphoneno        = rsget("interphoneno")
                FItemList(i).Fextension           = rsget("extension")
                FItemList(i).Fdirect070           = rsget("direct070")

                FItemList(i).Fpart_name     = rsget("part_name")
                FItemList(i).Fposit_name    = rsget("posit_name")

                i=i+1
			    rsget.movenext
		    loop
	    end if
	    rsget.close

    end function

    ''// 2013  ajaxContract.asp
    public function getDefaultValueByKey(ikey)
        dim i
        for i=0 to FResultCount -1
			if Not (FItemList(i) is Nothing) then
			    if (ikey=FItemList(i).FdetailKey) then
    				getDefaultValueByKey = FItemList(i).FDefaultValue
    				Exit function
    			end if
			end if
		next
    end function

    public function getValueByKey(ikey)
        dim i
        for i=0 to FResultCount -1
			if Not (FItemList(i) is Nothing) then
			    if (ikey=FItemList(i).FdetailKey) then
    				getValueByKey = FItemList(i).FdetailValue
    				Exit function
    			end if
			end if
		next
    end function

    ''// 2013  ajaxContract.asp
    public sub getContractDetailProtoTypeWithGroupInfo()
        dim sqlStr, i, ogroupInfo

        sqlStr = " select * from db_partner.dbo.tbl_partner_contractDetailType"
	    sqlStr = sqlStr & " where contractType=" & FRectContractType & ""

        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			i=0
			do until rsget.eof
    			set FItemList(i) = new CPartnerContractDetailTypeItem

    			FItemList(i).FcontractType          = rsget("contractType")
                FItemList(i).FdetailKey             = rsget("detailKey")
                FItemList(i).FdetailDesc            = db2html(rsget("detailDesc"))
                i=i+1
				rsget.movenext
			loop
		end if
		rsget.close

		if (i>0) then
            SET ogroupInfo = new CPartnerGroup
            ogroupInfo.FRectGroupid = FRectGroupid
            if (FRectGroupid<>"") then
                ogroupInfo.GetOneGroupInfo
            end if

            if (ogroupInfo.FResultCount<1) then
                SET ogroupInfo = Nothing
                exit sub
            end if

		    for i=0 to FResultCount -1
    			if Not (FItemList(i) is Nothing) then
    				FItemList(i).FDefaultValue = getDefaultContractValue(FItemList(i).FdetailKey,ogroupInfo)
    			end if
    		next
    		SET ogroupInfo = Nothing
		end if
    end sub

    public Sub GetOneContractMaster()
	    dim sqlStr

	    sqlStr = " select M.*,U.username as regUserNAme, U.usermail, U.interphoneno, U.extension,U.direct070, U2.username as sendUsername,U3.username as finUserName "
	    sqlStr = sqlStr & " ,T.contractName, T.subType  "
	    sqlStr = sqlStr & " ,(select count(*) as CNT from db_partner.dbo.tbl_partner_ctr_sub S where S.ctrKey=M.ctrKey) as contractSubCnt"
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$CONTRACT_DATE$$') as contractDate"
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$DEFAULT_JUNGSANDATE$$') as contractJungsanDate"
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$B_UPCHENAME$$') as B_UPCHENAME"
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$A_COMPANY_NO$$') as A_COMPANY_NO"
		sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$B_COMPANY_NO$$') as B_COMPANY_NO"
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master M"
	    sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_contractType T"
	    sqlStr = sqlStr & "     on T.contractType=M.contractType"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U"
	    sqlStr = sqlStr & "     on M.reguserid=U.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U2"
	    sqlStr = sqlStr & "     on M.senduserid=U2.userid"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten U3"
	    sqlStr = sqlStr & "     on M.finUserid=U3.userid"
	       sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner_group G"
	    sqlStr = sqlStr & "     on M.groupid=G.groupid"
	    sqlStr = sqlStr & " where M.ctrState>=0"
	    sqlStr = sqlStr & " and M.ctrKey=" & FRectCtrKey & ""

	    if FRectGroupid<>"" then
	        sqlStr = sqlStr & " and M.groupid='" & FRectGroupid & "'"
	    end if

	    if FRectMakerid<>"" then
	        sqlStr = sqlStr & " and M.makerid='" & FRectMakerid & "'"
	    end if

	    'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractItem2013
			FOneItem.FctrKey              = rsget("ctrKey")
			FOneItem.Fgroupid             = rsget("groupid")
            FOneItem.Fmakerid             = rsget("makerid")
            FOneItem.FContractType        = rsget("ContractType")
            FOneItem.FctrState            = rsget("ctrState")
            FOneItem.FctrNo               = rsget("ctrNo")

            FOneItem.FcontractContents    = db2html(rsget("contractContents"))

            FOneItem.FregUserID           = rsget("regUserID")
            FOneItem.FsendUserid          = rsget("sendUserid")
            FOneItem.FfinUserID           = rsget("finUserID")

            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Fsenddate            = rsget("senddate")
            FOneItem.Fconfirmdate         = rsget("confirmdate")
            FOneItem.Ffinishdate          = rsget("finishdate")
            FOneItem.Fdeletedate          = rsget("deletedate")

            FOneItem.FcontractName          = rsget("contractName")
            FOneItem.FcontractDate          = rsget("contractDate")
            FOneItem.FcontractJungsanDate   = rsget("contractJungsanDate")
            FOneItem.FB_UPCHENAME           = rsget("B_UPCHENAME")

            FOneItem.FsubType               = rsget("subType")
            FOneItem.FsignType              = rsget("signType")
            FOneItem.FdocuSignId            = rsget("docuSignId")
            FOneItem.FdocuSignUri           = rsget("docuSignUri")
            FOneItem.FdocuSignSenddate      = rsget("docuSignSenddate")

            FOneItem.FregUsername          = rsget("regusername")
            FOneItem.FsendUsername          = rsget("sendUsername")
            FOneItem.FfinUserName            = rsget("finUserName")
            

            FOneItem.FecCtrSeq  = rsget("ecCtrSeq")
            FOneItem.FecAUser  = rsget("ecAUser")
            FOneItem.FecBUser  = rsget("ecBUser")
            FOneItem.FAcompany_no = replace(rsget("A_COMPANY_NO"),"-","")
            FOneItem.FBcompany_no = replace(rsget("B_COMPANY_NO"),"-","")
            
            'FOneItem.FContractEtcContetns = db2html(rsget("ContractEtcContetns"))
'            FOneItem.Fusermail            = rsget("usermail")
'            FOneItem.Finterphoneno        = rsget("interphoneno")
'            FOneItem.Fextension           = rsget("extension")
'            FOneItem.Fdirect070           = rsget("direct070")

		end if
		rsget.close

    end Sub
    
public Facctoken
public Freftoken

	public Function fnGetContractToken()
		dim sqlStr
		sqlStr = "select top 1 access_token, refresh_token from db_partner.dbo.tbl_partner_ctrLg_token order by tidx desc "
		rsget.Open sqlStr,dbget,1
		if not rsget.eof then
			Facctoken = rsget("access_token")
			Freftoken = rsget("refresh_token")
		end if
		rsget.close
	End Function

  


    public Sub GetContractDetailList()
	    dim sqlStr, i

	    sqlStr = " select A.*, t.detailDesc ,t.orderno from "
	    sqlStr = sqlStr & " ("
	    sqlStr = sqlStr & "     select c.ContractType, d.* from db_partner.dbo.tbl_partner_ctr_master c,"
	    sqlStr = sqlStr & "     db_partner.dbo.tbl_partner_ctr_Detail d"
	    sqlStr = sqlStr & "     where d.ctrKey=" & FRectCtrKey & ""
	    sqlStr = sqlStr & "     and d.ctrKey=c.ctrKey"
	    sqlStr = sqlStr & " ) A"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner_contractDetailType t"
	    sqlStr = sqlStr & "     on A.ContractType=t.ContractType"
	    sqlStr = sqlStr & "     and A.detailKey=t.detailKey"
	    sqlStr = sqlStr & "     order by t.orderno asc"

	    'response.write sqlStr &"<br>"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			i=0
			do until rsget.eof
    			set FItemList(i) = new CPartnerContractDetailItem
    			FItemList(i).FctrKey                = rsget("ctrKey")
                FItemList(i).FdetailKey             = rsget("detailKey")
                FItemList(i).FdetailValue           = db2html(rsget("detailValue"))
                FItemList(i).FdetailDesc            = db2html(rsget("detailDesc"))
                i=i+1
				rsget.movenext
			loop
		end if
		rsget.close

    end Sub

    public Sub GetContractSubList()
        dim sqlStr, i

	    sqlStr = "select top 1110 S.ctrSubKey, S.sellplace, S.mwdiv,S.defaultmargin,S.defaultdeliveryType,S.defaultFreebeasongLimit,S.defaultdeliverpay "
	    sqlStr = sqlStr&" ,(CASE WHEN S.sellplace='ON' THEN '온라인' "
	    sqlStr = sqlStr&"   ELSE isNULL(u.shopname,S.sellplace) END) as sellplaceName"
        sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_Sub S"
        sqlStr = sqlStr&"       left join db_shop.dbo.tbl_shop_user U"
        sqlStr = sqlStr&"       on S.sellplace=U.userid"
        sqlStr = sqlStr&" where S.ctrKey="&ctrKey
        sqlStr = sqlStr&" order by S.ctrSubKey"

        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			i=0
			do until rsget.eof
    			set FItemList(i) = new CPartnerAddContractSubItem ''CPartnerContractSubItem
    			FItemList(i).FctrSubKey           = rsget("ctrSubKey")
    			FItemList(i).Fsellplace           = rsget("sellplace")
                FItemList(i).Fcontractmwdiv       = rsget("mwdiv")
                'FItemList(i).FmwdivName           = rsget("mwdivName")
                FItemList(i).Fcontractmargin      = rsget("defaultmargin")

                FItemList(i).FsellplaceName           = rsget("sellplaceName")

                FItemList(i).FdefaultdeliveryType = rsget("defaultdeliveryType")
                FItemList(i).FdefaultFreebeasongLimit = rsget("defaultFreebeasongLimit")
                FItemList(i).Fdefaultdeliverpay = rsget("defaultdeliverpay")

                i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub


    public Sub getRecentDefaultContract()
        dim sqlStr, i
        sqlStr = " select top 1 M.ctrKey,M.groupid, M.ContractType,M.ctrState,M.ctrNo"
        sqlStr = sqlStr&" ,M.regUserID,M.sendUserid,M.finUserID,M.regdate,M.senddate,M.confirmdate,M.finishdate,T.contractName,M.ctrNo "
        sqlStr = sqlStr & " ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$CONTRACT_DATE$$') as contractDate"
        sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_master M"
        sqlStr = sqlStr&" 	Join db_partner.dbo.tbl_partner_contractType T"
        sqlStr = sqlStr&" 	on M.contractType=T.contractType"
        sqlStr = sqlStr&" where M.groupid='"&FRectGroupId&"'"
        sqlStr = sqlStr&" and T.subType=0"
        sqlStr = sqlStr&" and M.ctrState>=0"
        sqlStr = sqlStr&" order by M.ctrKey desc"

        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractItem2013
			FOneItem.FctrKey              = rsget("ctrKey")
			FOneItem.Fgroupid             = rsget("groupid")
            FOneItem.FContractType        = rsget("ContractType")
            FOneItem.FctrState            = rsget("ctrState")
            FOneItem.FctrNo               = rsget("ctrNo")

            FOneItem.FregUserID           = rsget("regUserID")
            FOneItem.FsendUserid          = rsget("sendUserid")
            FOneItem.FfinUserID           = rsget("finUserID")

            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Fsenddate            = rsget("senddate")
            FOneItem.Fconfirmdate         = rsget("confirmdate")
            FOneItem.Ffinishdate          = rsget("finishdate")

            FOneItem.FcontractName          = rsget("contractName")
            FOneItem.FcontractDate          = rsget("contractDate")

		end if
		rsget.close
    end Sub

    public Sub getRecentAddContract(isOff)
        dim sqlStr, i

        sqlStr = " select top 1 T.contractName,M.ctrNo ,M.CtrState,S.sellplace, (CASE WHEN S.sellplace='ON' THEN '온라인' ELSE U.shopname END) as sellplaceName, s.mwdiv, s.defaultMargin "
        sqlStr = sqlStr&" ,(select count(*) from db_partner.dbo.tbl_partner_ctr_Sub S1 where M.ctrKey=S1.ctrKey and S.sellplace<>'ON') as addCtrCNT "
        sqlStr = sqlStr&" ,(select D.detailValue from db_partner.dbo.tbl_partner_ctr_Detail D where D.ctrKey=M.ctrKey and D.detailKey='$$CONTRACT_DATE$$') as contractDate"
        sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_ctr_master M "
        sqlStr = sqlStr&" 	Join db_partner.dbo.tbl_partner_contractType T "
        sqlStr = sqlStr&" 	on M.contractType=T.contractType "
        sqlStr = sqlStr&" 	Join db_partner.dbo.tbl_partner_ctr_Sub S "
        sqlStr = sqlStr&" 	on M.ctrKey=S.ctrKey "
        sqlStr = sqlStr&" 	left join db_shop.dbo.tbl_shop_user U "
        sqlStr = sqlStr&" 	on S.sellplace=U.UserID "
        sqlStr = sqlStr&" where M.groupid='"&FRectGroupId&"' "
        sqlStr = sqlStr&" and M.makerid='"&FRectMakerId&"' "
        sqlStr = sqlStr&" and T.subType>0 "
        sqlStr = sqlStr&" and M.ctrState>=0 "
        if (isOff) then
            sqlStr = sqlStr&" and S.sellplace<>'ON' "
        else
            sqlStr = sqlStr&" and S.sellplace='ON' "
        end if
				sqlStr = sqlStr&" order by m.ctrKey desc"
				
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then
            set FOneItem = new CPartnerAddContractSubItem
            FOneItem.FcontractName    = rsget("contractName")
    		FOneItem.FctrNo           = rsget("ctrNo")
    		FOneItem.Fsellplace       = rsget("sellplace")
    		FOneItem.FsellplaceName   = rsget("sellplaceName")
            FOneItem.Fcontractmwdiv       = rsget("mwdiv")
            FOneItem.Fcontractmargin      = rsget("defaultmargin")
            FOneItem.FCtrState          = rsget("CtrState")
            FOneItem.FcontractDate           = rsget("contractDate")

        end if
		rsget.close
    end Sub
    
    '계약서 상태값 가져오기 2014-12-12
    '-- 계약서 여러개중 상태값이 가장 작은 값 하나 가져와서 7(계약완료)인가 아닌가 확인
    public Function fnGetCtrState
      dim sqlStr 
      	sqlStr = " select top 1 ctrstate "
    	sqlStr = sqlStr&" from db_partner.dbo.tbl_partner as p "
    	sqlStr = sqlStr&" 		inner join  db_partner.dbo.tbl_partner_ctr_master as m on p.groupid = m.groupid "
    	sqlStr = sqlStr&"		left outer join db_partner.dbo.tbl_partner_ctr_sub as s on m.ctrKey = s.ctrKey and s.sellplace ='ON'"
    	sqlStr = sqlStr&" where p.id='"&FRectMakerId&"' " 
        sqlStr = sqlStr&" and m.ctrState>=0  order by ctrstate " 
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1
        if Not ( rsget.Eof or rsget.Bof ) then
        	FCtrState   = rsget("CtrState")	
        else 	
        	FCtrState = 0
        end if
        rsget.close
	End Function

''-----------------------------------------------------------------------------------------------------------------------
'Class CPartnerContractSubItem
'    public FctrSubKey
'    public Fsellplace
'    public Fmwdiv
'    public FmwdivName
'    public Fdefaultmargin
'
'    public FdefaultdeliveryType
'    public FdefaultFreebeasongLimit
'    public Fdefaultdeliverpay
'
'    '' 판매처
'    public function getSellplaceName()
'        if isNULL(Fsellplace) then Exit function
'        select case Fsellplace
'            CASE "DF" :
'                getSellplaceName="대표마진"
'            CASE "ON" :
'                getSellplaceName="온라인"
'            CASE ELSE : getSellplaceName=Fsellplace
'        end select
'    end function
'
'    public function getBigoStr()
'        dim bufStr
'        IF (Fmwdiv="U") and (FdefaultdeliveryType="7") then
'            bufStr = bufStr & "업체착불배송"
'        elseIF (Fmwdiv="U") and (FdefaultdeliveryType="9") then
'            bufStr = bufStr & "업체조건배송"
'            bufStr = bufStr & "<br>"&FormatNumber(FdefaultFreebeasongLimit,0)&"원미만"
'            bufStr = bufStr & "<br>배송료"&FormatNumber(Fdefaultdeliverpay,0)&"원"
'        else
'            bufStr = bufStr & ""
'        end if
'
'        getBigoStr = bufStr
'    end function
'
'    Private Sub Class_Initialize()
'	End Sub
'	Private Sub Class_Terminate()
'	End Sub
'end Class

	'//designer/company/popContract.asp



    public Sub GetRecentContractbyOnOff()
         dim sqlStr
	    sqlStr = " select C.* from db_partner.dbo.tbl_partner_contract C"
	    sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_contractType T"
	    sqlStr = sqlStr & "     on C.contractType=T.contractType"
	    sqlStr = sqlStr & "     and T.onoffgubun='"&FRectOnOffGubun&"'"
	    sqlStr = sqlStr & " where C.makerid='" & FRectMakerid & "'"
	    sqlStr = sqlStr & " and C.contractState>=0"
        sqlStr = sqlStr & " order by C.contractID desc"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractItem2013

			FOneItem.FcontractID          = rsget("contractID")
            FOneItem.Fmakerid             = rsget("makerid")
            FOneItem.FContractType        = rsget("ContractType")
            FOneItem.FContractState       = rsget("ContractState")
            FOneItem.FcontractName        = db2html(rsget("contractName"))
            FOneItem.FcontractNo          = rsget("contractNo")
            FOneItem.FContractContents    = db2html(rsget("ContractContents"))
            FOneItem.FContractEtcContetns = db2html(rsget("ContractEtcContetns"))
            FOneItem.Freguserid           = rsget("reguserid")
            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Fconfirmdate         = rsget("confirmdate")
            FOneItem.Ffinishdate          = rsget("finishdate")

		end if
		rsget.close
    end Sub

	public Sub GetLastOneContract()
	    dim sqlStr
	    sqlStr = " select * from db_partner.dbo.tbl_partner_contract"
	    sqlStr = sqlStr & " where makerid='" & FRectMakerid & "'"
	    sqlStr = sqlStr & " and contractState>=0"
	    sqlStr = sqlStr & " and contractState<7"
        sqlStr = sqlStr & " order by contractID desc"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractItem2013

			FOneItem.FcontractID          = rsget("contractID")
            FOneItem.Fmakerid             = rsget("makerid")
            FOneItem.FContractType        = rsget("ContractType")
            FOneItem.FContractState       = rsget("ContractState")
            FOneItem.FcontractName        = db2html(rsget("contractName"))
            FOneItem.FcontractNo          = rsget("contractNo")
            FOneItem.FContractContents    = db2html(rsget("ContractContents"))
            FOneItem.FContractEtcContetns = db2html(rsget("ContractEtcContetns"))
            FOneItem.Freguserid           = rsget("reguserid")
            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Fconfirmdate         = rsget("confirmdate")
            FOneItem.Ffinishdate          = rsget("finishdate")

		end if
		rsget.close


    end Sub



	'/admin/member/contractPrototypeReg.asp
	public sub getContractDetailProtoType()
	    dim sqlStr, i

	    sqlStr = " select * from db_partner.dbo.tbl_partner_contractDetailType"
	    if FRectContractType<>"" then
	        sqlStr = sqlStr & " where contractType=" & FRectContractType & ""
	    else
	        sqlStr = sqlStr & " where 1=0"
        end if

        sqlStr = sqlStr & " order by orderno asc"

        'response.write sqlStr &"<br>"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			i=0
			do until rsget.eof
    			set FItemList(i) = new CPartnerContractDetailTypeItem

    			FItemList(i).FcontractType          = rsget("contractType")
                FItemList(i).FdetailKey             = rsget("detailKey")
                FItemList(i).FdetailDesc            = db2html(rsget("detailDesc"))

                i=i+1
				rsget.movenext
			loop
		end if
		rsget.close

    end Sub


    public sub getOneContractDetailProtoType()
	    dim sqlStr, i

	    sqlStr = " select * from db_partner.dbo.tbl_partner_contractDetailType"
	    sqlStr = sqlStr & " where contractType=" & FRectContractType & ""
        sqlStr = sqlStr & " and detailKey='" & FRectdetailKey & "'"

        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractDetailTypeItem
			FOneItem.FcontractType          = rsget("contractType")
            FOneItem.FdetailKey             = rsget("detailKey")
            FOneItem.FdetailDesc            = db2html(rsget("detailDesc"))
		end if
		rsget.close

    end Sub

    '//admin/member/contractPrototypeReg.asp
	public sub getOneContractProtoType()
	    dim sqlStr, i

	    sqlStr = " select * "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType"
	    if FRectContractType<>"" then
	        sqlStr = sqlStr & " where contractType=" & FRectContractType & ""
	    else
	        sqlStr = sqlStr & " where 1=0"
        end if

	    'response.write sqlStr &"<br>"
	    rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractTypeItem

			FOneItem.FContractType           = rsget("contractType")
            FOneItem.FContractName           = db2html(rsget("contractName"))
            FOneItem.FContractContents       = db2html(rsget("ContractContents"))
            FOneItem.Fregdate                = rsget("regdate")
            FOneItem.fonoffgubun             = rsget("onoffgubun")
            FOneItem.fsubtype                = rsget("subtype")
		end if
		rsget.close

    end sub

	'//admin/member/contractPrototypeReg.asp
	public Sub getValidContractProtoTypeList()
	    dim sqlStr, i
	    sqlStr = "select count(contractType) as cnt "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType"
	    sqlStr = sqlStr & " where isusing='Y'"
        sqlStr = sqlStr & " and subtype>=0"

	    rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " & CStr(FPageSize*FCurrPage) & " contractType,contractName,regdate ,onoffgubun , subtype"
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType"
	    sqlStr = sqlStr & " where isusing='Y'"
        sqlStr = sqlStr & " and subtype>=0"

		'response.write sqlStr &"<Br>"
	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof

				set FItemList(i) = new CPartnerContractTypeItem

				FItemList(i).fonoffgubun           = rsget("onoffgubun")
				FItemList(i).FcontractType         = rsget("contractType")
                FItemList(i).FcontractName         = db2html(rsget("contractName"))
                FItemList(i).Fregdate              = rsget("regdate")
                FItemList(i).Fsubtype              = rsget("subtype")
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close

    end Sub

    public Sub GetMakerNotConfirmContractList()
        dim sqlStr, i
        sqlStr = "select top " & CStr(FPageSize*FCurrpage) & " c.contractID "
		sqlStr = sqlStr & " , c.makerid, c.contractType, c.contractName"
		sqlStr = sqlStr & " , c.contractNo, c.contractState, c.reguserid, c.regdate, c.confirmdate, c.finishdate"
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract c"
	    sqlStr = sqlStr & " where makerid='" & FRectMakerid & "'"
	    sqlStr = sqlStr & " and contractState>0"
	    sqlStr = sqlStr & " and contractState<3"
	    sqlStr = sqlStr & " order by c.contractID desc"

	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerContractItem2013

				FItemList(i).FcontractID           = rsget("contractID")
                FItemList(i).Fmakerid              = rsget("makerid")
                FItemList(i).FcontractType         = rsget("contractType")
                FItemList(i).FcontractNo           = rsget("contractNo")
                FItemList(i).FcontractName         = db2html(rsget("contractName"))
                FItemList(i).FcontractState        = rsget("contractState")
                FItemList(i).Freguserid            = rsget("reguserid")
                FItemList(i).Fregdate              = rsget("regdate")
                FItemList(i).Fconfirmdate          = rsget("confirmdate")
                FItemList(i).Ffinishdate           = rsget("finishdate")
                'FItemList(i).FcontractContents     = rsget("contractContents")
                'FItemList(i).FcontractEtcContetns  = rsget("contractEtcContetns")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub

    '//designer/company/popContract.asp
    public Sub GetMakerValidContractList()
        dim sqlStr, i
        sqlStr = "select top " & CStr(FPageSize*FCurrpage) & " c.contractID "
		sqlStr = sqlStr & " , c.makerid, c.contractType, c.contractName"
		sqlStr = sqlStr & " , c.contractNo, c.contractState, c.reguserid, c.regdate, c.confirmdate, c.finishdate"
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract c"
	    sqlStr = sqlStr & " where makerid='" & FRectMakerid & "'"
	    sqlStr = sqlStr & " and contractState>0   "
	    sqlStr = sqlStr & " order by c.contractID desc"

	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerContractItem2013
				FItemList(i).FcontractID           = rsget("contractID")
                FItemList(i).Fmakerid              = rsget("makerid")
                FItemList(i).FcontractType         = rsget("contractType")
                FItemList(i).FcontractNo           = rsget("contractNo")
                FItemList(i).FcontractName         = db2html(rsget("contractName"))
                FItemList(i).FcontractState        = rsget("contractState")
                FItemList(i).Freguserid            = rsget("reguserid")
                FItemList(i).Fregdate              = rsget("regdate")
                FItemList(i).Fconfirmdate          = rsget("confirmdate")
                FItemList(i).Ffinishdate           = rsget("finishdate")

                'FItemList(i).FcontractContents     = rsget("contractContents")
                'FItemList(i).FcontractEtcContetns  = rsget("contractEtcContetns")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub



	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage       = 1
		FPageSize       = 12
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     = 0
	End Sub

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class

Sub drawSelectBoxContractTypeWithChangeEvent(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select name="<%= selectBoxName %>" onchange="ChangeContractType(this)" class="select">
     <option value="" <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select contractType,contractName from db_partner.dbo.tbl_partner_contractType"
   query1 = query1 & " where isusing='Y' and subtype>=0"
   query1 = query1 & " order by contractType"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("contractType")) then
               tmp_str = " selected"
           end if
           response.write("<option value="""&rsget("contractType")&""" "&tmp_str&">"& db2html(rsget("contractName"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

function getMdUserName(mduserid)
    dim sqlStr
    sqlStr = "select company_name from [db_partner].[dbo].tbl_partner"
    sqlStr = sqlStr & " where id='" & mduserid & "'"
    rsget.Open sqlStr,dbget,1
    if  not rsget.EOF  then
        getMdUserName = db2html(rsget("company_name"))
    end if
    rsget.close
end function


function drawSubTypeGubun(boxname ,stats)
%>
	<select name='<%=boxname%>' class="select">
		<option value='' <% if stats = "" then response.write " selected" %>>선택</option>
		<option value='0' <% if stats = "0" then response.write " selected" %>>기본계약서(0)</option>
		<option value='5' <% if stats = "5" then response.write " selected" %>>부속합의서(5)</option>
		<option value='7' <% if stats = "7" then response.write " selected" %>>물품공급계약서(7)</option>
		<!-- <option value='9' <% if stats = "9" then response.write " selected" %>>기타(9)</option> -->
	</select>

<%
end function

'//해당 브랜드에 대한 샵의 마진을 반환한다
Sub drawSelectOffShopmargin(selectBoxName,selectedId)
dim tmp_str,query1
%>
   <select class="select" name="<%=selectBoxName%>" onchange="SelectOffShopmargin(this.value);">
   <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where isusing='Y' "
   'query1 = query1 & " and userid<>'streetshop000'"	'직영점(대표)
   'query1 = query1 & " and userid<>'streetshop800'"		'가맹점(대표)
   'query1 = query1 & " and userid<>'streetshop870'"		'도매(대표)
   'query1 = query1 & " and userid<>'streetshop700'"		'해외(대표)

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
   'response.write query1 &"<br>"
%>
	<script language='javascript'>
		function SelectOffShopmargin(tmp){
			frmReg.shopid.value = tmp;
			frmReg.target = 'view';
			frmReg.action = 'contractReg_selectshopmargin.asp';
			frmReg.submit();
		}
	</script>
	<input type="hidden" name="shopid">
	<iframe id="view" name="view" frameborder="0" width=0 height=0></iframe>
<%
end sub

function drawSelectshopuser(selectBoxName,selectedId,btcid,changeflg)
dim tmp_str,query1
%>
   <select class="select" name="<%=selectBoxName%>" <%=changeflg%>>
   <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
	query1 = "select top 30" + VbCrlf
	query1 = query1 & " ut.userid , ps.shopid ,su.shopname, ps.firstisusing" + VbCrlf
	query1 = query1 & " ,(case when ps.firstisusing='Y' then '[대표]' end) as firstname" + VbCrlf
	query1 = query1 + " from db_partner.dbo.tbl_user_tenbyten ut" + vbcrlf
	query1 = query1 + " join db_partner.dbo.tbl_partner_shopuser ps" + vbcrlf
	query1 = query1 + " on ps.empno = ut.empno" + vbcrlf
	query1 = query1 + " join db_shop.dbo.tbl_shop_user su" + vbcrlf
	query1 = query1 + " on ps.shopid = su.userid" + vbcrlf
	query1 = query1 + " where ut.isusing=1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	query1 = query1 & "	and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
    query1 = query1 & "	and ut.userid = '"&btcid&"'"

	'response.write query1 &"<br>"
	rsget.Open query1,dbget,1

	if  not rsget.EOF  then
	   rsget.Movefirst

	   do until rsget.EOF
	       if Lcase(selectedId) = Lcase(rsget("shopid")) then
	           tmp_str = " selected"
	       end if
	       response.write("<option value='"&rsget("shopid")&"' "&tmp_str&">"&rsget("shopid")&"/"&rsget("shopname")&rsget("firstname")&"</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   loop
	end if
	rsget.close
	response.write("</select>")

end function

function fnGetEcBUser(byVal groupid)
dim strSql, ecBUser
 strSql = "select top 1 ecBUser  FROM db_partner.dbo.tbl_partner_ctr_master where groupid = '"&groupid&"' and ecCtrSeq is not null order by isnull(lastupdate, regdate) desc"
 rsget.Open strSql,dbget,1
 if not rsget.eof then
 	ecBUser = rsget("ecBUser")
end if
rsget.close
 fnGetEcBUser =ecBUser
End Function

Function getLgEcErrMessage(byVal cd)
    dim strErrMsg
    Select Case cd
        Case "001"
            strErrMsg= "venderno 값 없음"
        Case "002"
            strErrMsg= "type_seq 값 없음"
        Case "003"
            strErrMsg= "title 값 없음"
        Case "004"
            strErrMsg= "contract_dt 값 없음"
        Case "005"
            strErrMsg= "rcontract_money 값 없음"
        Case "011"
            strErrMsg= "membList(계약자 정보) 값 없음"
        Case "012"
            strErrMsg= "membList(계약자 정보)가 10이상"
        Case "013"
            strErrMsg= "계약자 구분 A(작성자) 정보와 계약서 본문의 사업자번호 다름"
        Case "014"
            strErrMsg= "계약자 구분 값이 순차적이지 않음 (A,B,C,D...)"
        Case "015"
            strErrMsg= "membList.venderno 값없음"
        Case "016"
            strErrMsg=" membList.company 값없음"
        Case "020"
            strErrMsg="venderno 에 사용자 존재하지 않음"
        Case "021"
            strErrMsg=" 해당정보에 대한 문서가 존재하지 않음"
        Case "030"
            strErrMsg="membList 에서 venderno 에 대한 사용자가 존재하지않음."
        Case Else
            strErrMsg=""
    end Select

    getLgEcErrMessage = strErrMsg
End Function

Function getSignTypeCode(signtype)
    Select Case Trim(SignType)
        Case "1" ''수기
            getSignTypeCode = "H"
        Case "2" ''유플러스
            getSignTypeCode = "U"
        Case "3" ''DocuSign
            getSignTypeCode = "D"
        Case Else
            getSignTypeCode = "H"
    End Select
End Function
%>
