<%
'###########################################################
' Description : 핑거스 입점 동의 관련
' Hieditor : 2016.08.10 서동석 생성
'###########################################################
'' db_partner.[dbo].[tbl_partner_contractType]
CONST YAKGAN_FINGERS        = 17
CONST CONTRACTTYPE_FINGERS  = 16
CONST PREMIUM_CONTRACTTYPE_FINGERS  = 18

CONST CPRM_MONTHPAY = 30000
CONST CPRM_COMMISION = 5
CONST CPG_COMMISION = 3



CONST ChashVal = "TBTCTR"

Function chang_money(imoney)
    dim num1 ' 한글 숫자 배열
    dim num2 ' 한글 숫자 단위 배열
    dim posNoLevel ' 한글 숫자 단위 출력 위치
    dim tempNo ' 한글 숫자 현재 단위 조립용
    dim strNo ' 한글 숫자 전체 조립용
    dim cntNo ' 변환할 숫자의 길이
    dim posNo ' 변환할 숫자의 현재 변환 위치
    dim mo, no
    
    num1 = Array("", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구")
    num2 = Array("", "십", "백", "천", "만", "십", "백", "천", "억", "십", "백", "천", "조", "십", "백", "천", "경")
    
    cntNo = Len(imoney)
    
    ' 숫자가 0 일 경우
    if imoney = 0 then 
            strNo = "영"
    else
            strNo = ""
            posNoLevel = 0
            posNo = cntNo
            do
                    mo = Cint( Mid(imoney, posNo, 1) )

                    ' 나머지 값이 0 이 아닐 경우
                    if 0 < mo then
                            tempNo = num1(mo)
                            tempNo = tempNo & num2(posNoLevel)
                            strNo = tempNo & strNo
                    else
                            ' 나머지 값이 0 이면서 10000 단위일때(만, 억, ..)
                            if (posNoLevel Mod 4) = 0 then
                                    strNo = num2(posNoLevel) & strNo
                            end if
                    end if
                    
                    posNoLevel = posNoLevel + 1
                    posNo = posNo - 1
            loop while 0 < posNo
    end if
    
    chang_money = strNo
End Function

Function chang_money2(imoney)
    dim strNo, i, mo, cntNo, imoneyStr
    dim num1 : num1 = Array("", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구")
    imoneyStr = CStr(imoney)
    cntNo = Len(imoneyStr)
    
    if imoney = 0 then 
       strNo = "영"
    else
        for i=0 to cntNo-1
            mo = Cint( Mid(imoneyStr, i+1, 1) )
            strNo = strNo&num1(mo)
        next
    end if
    chang_money2 = strNo
end function

function getNum2KORFormat(iorg)
    dim buf : buf= CSTR(iorg)
   
    if (FIX(iorg)<>iorg) then
        getNum2KORFormat ="("&chang_money(FIX(iorg))&"."&chang_money2(CLNG((iorg-FIX(iorg))*100))&")"
    else
        getNum2KORFormat = "("&chang_money(iorg)&")"
    end if
end function

Sub DrawfingerAgreeMasterCombo(selectBoxName,selectedId)
   dim tmp_str,sqlStr
   %><select name="<%= selectBoxName %>" onchange="ChangeContractType(this)">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   sqlStr = " select contractType,contractName,isusing from db_partner.dbo.tbl_partner_contractType"
   sqlStr = sqlStr & " where 1=1"
   sqlStr = sqlStr & " and contractType in ("&YAKGAN_FINGERS&","&CONTRACTTYPE_FINGERS&")"
   sqlStr = sqlStr & " order by subType"
   rsget.CursorLocation = adUseClient
   rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("contractType")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("contractType")&"' "&tmp_str&">"&rsget("contractType")&" ["&db2html(rsget("contractName"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub DrawAgreeStateCombo(selectBoxName,selectedId)
    dim tmp_str
    %>
    <select name="<%= selectBoxName %>">
    	<option value="">전체
    	<option value="1" <% if selectedId="1" then response.write "selected" %> >미완료
    	<option value="3" <% if selectedId="3" then response.write "selected" %> >동의완료
    </select>
    <%
End Sub    	

''관련(동일업체) 브랜드 콤보Box
sub DrawSameGroupBrandUpche(igroupid,imakerid,selboxname, addStr)
    dim sqlStr, id, socname, ret, i
    sqlStr ="select p.id, c.socname"
    sqlStr = sqlStr& " from db_partner.dbo.tbl_partner p"
    sqlStr = sqlStr& "  Join db_user.dbo.tbl_user_c c"
    sqlStr = sqlStr& "  on p.id=c.userid"
    sqlStr = sqlStr& "  and c.userdiv='14'"
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

''아카데미 DIY 업체인지
function IsAcademyDiyUpche(imakerid)
    dim sqlStr
    IsAcademyDiyUpche = FALSE
    if (session("ssUserCDiv")<>"14") then Exit function
    
    sqlStr ="select lecturer_id , lec_yn, diy_yn, lec_margin, mat_margin, diy_margin,diy_dlv_gubun,DefaultFreeBeasongLimit,DefaultDeliveryPay"&VBCRLF
    sqlStr = sqlStr & " from [academydb].db_academy.dbo.tbl_lec_user"&VBCRLF
    sqlStr = sqlStr & " where lecturer_id='"&imakerid&"'"&VBCRLF
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        IsAcademyDiyUpche = (rsget("diy_yn")="Y")
    end if
    rsget.close
end function

function IsFingersUpcheAgreeNotiRequire(igroupid,imakerid)
    dim sqlStr
    dim agreeidx : agreeidx =0
    
    IsFingersUpcheAgreeNotiRequire = false
    if (session("ssUserCDiv")<>"14") then Exit function
    
    ''DIY 업체인지 체크 // DIY 업체만 계약서, 약관 승인
    if (NOT IsAcademyDiyUpche(imakerid)) then
        IsFingersUpcheAgreeNotiRequire = false
        session("isAgreeReq")="N"
        
    end if
    
    ''이용약관 승인 체크
    sqlStr = " select top 1 agreeidx "
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_fingers_agreeHist"
    sqlStr = sqlStr & " where 1=1"
    sqlStr = sqlStr & " and groupid='"&igroupid&"'"
    sqlStr = sqlStr & " and contractType="&YAKGAN_FINGERS&""
    sqlStr = sqlStr & " and agreedate is Not NULL"
    sqlStr = sqlStr & " and deldate is NULL"
    sqlStr = sqlStr & " order by agreeidx desc"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        agreeidx = rsget("agreeidx")
    end if
    rsget.close   
    
    if (agreeidx<1) then
        IsFingersUpcheAgreeNotiRequire = true
        session("isAgreeReq")="Y"
        Exit function
    end if
    
    agreeidx =0
    
    ''계약서 승인 체크
    sqlStr = " select top 1 agreeidx "
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_fingers_agreeHist"
    sqlStr = sqlStr & " where groupid='"&igroupid&"'"
    sqlStr = sqlStr & " and makerid='"&imakerid&"'"
    sqlStr = sqlStr & " and contractType="&CONTRACTTYPE_FINGERS&""
    sqlStr = sqlStr & " and agreedate is Not NULL"
    sqlStr = sqlStr & " and deldate is NULL"
    sqlStr = sqlStr & " order by agreeidx desc"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        agreeidx = rsget("agreeidx")
    end if
    rsget.close   
    
    if (agreeidx<1) then
        IsFingersUpcheAgreeNotiRequire = true
        session("isAgreeReq")="Y"
        Exit function
    end if
    
    IsFingersUpcheAgreeNotiRequire = false
    session("isAgreeReq")="N"
end function

''개인정보 수집 관련 계약서 
function getPriContractContents(bUpchename)
    dim ret
    dim fs,f
    set fs=Server.CreateObject("Scripting.FileSystemObject")
    set f=fs.OpenTextFile(Server.MapPath("/lectureadmin/contract/viewContractWeb_Pri.html"),1)
    ret = f.ReadAll
    f.Close
    set f=Nothing
    set fs=Nothing

    getPriContractContents = replace(ret,"$$B_UPCHENAME$$",bUpchename)
end function

function getContractNoFormat(isubno,iagreeidx)
    Dim ret
    ret = TRim(replace(LEFT(Date(),10),"-",""))
    ret = ret & "-" & Format00(2,isubno) & "-" & Format00(6,iagreeidx)
    
    getContractNoFormat=ret
end function
            
''업체 이용약관 체크 후 생성
function checkUpcheYakganAgreeMake(igroupid,imakerid,byref agreeidx)
    dim sqlStr
    dim viewdate, agreedate, ContractNo, ContractContents
    agreeidx = 0
    
    if (groupid="")  then
        checkUpcheYakganAgreeMake = FALSE
        exit function
    end if
    
    sqlStr = " IF NOT EXISTS(select * from db_partner.dbo.tbl_partner_fingers_agreeHist where groupid='"&igroupid&"' and contractType="&YAKGAN_FINGERS&" and deldate is NULL)"
    sqlStr = sqlStr & " BEGIN"
    sqlStr = sqlStr & " insert into db_partner.dbo.tbl_partner_fingers_agreeHist"
    sqlStr = sqlStr & " (ContractType,groupid,makerid,ContractContents)"
    sqlStr = sqlStr & " select top 1 "&YAKGAN_FINGERS&""
    sqlStr = sqlStr & " ,'"&groupid&"'"
    sqlStr = sqlStr & " ,''" ''이용약관은 그룹코드별로.
    sqlStr = sqlStr & " ,NULL" ''',ContractContents from db_partner.[dbo].[tbl_partner_contractType] where ContractType="&YAKGAN_FINGERS&""
    sqlStr = sqlStr & " END"
    dbget.Execute sqlStr
        
    sqlStr = " select top 1 agreeidx,viewdate,agreedate,ContractContents, ContractNo from db_partner.dbo.tbl_partner_fingers_agreeHist"
    sqlStr = sqlStr & " where 1=1"
    sqlStr = sqlStr & " and groupid='"&igroupid&"'"
    sqlStr = sqlStr & " and contractType="&YAKGAN_FINGERS&""
    sqlStr = sqlStr & " and deldate is NULL"
    sqlStr = sqlStr & " order by agreeidx desc"
 
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        agreeidx = rsget("agreeidx")
        viewdate = rsget("viewdate")
        agreedate = rsget("agreedate")
        ContractContents = rsget("ContractContents")
        ContractNo = rsget("ContractNo")
    end if
    rsget.close

'    ''계약서 내용 생성. 2016/08/26
    dim retContractContents
    if (agreeidx>0) then
        if (IsNULL(ContractContents)) then
            retContractContents = MakeConTractContents(agreeidx)
            sqlStr = " update db_partner.dbo.tbl_partner_fingers_agreeHist"
            sqlStr = sqlStr&" set ContractContents='"&HTML2DB(retContractContents)&"'"
            sqlStr = sqlStr&" where agreeidx="&agreeidx
            sqlStr = sqlStr&" and ContractType="&YAKGAN_FINGERS
            sqlStr = sqlStr&" and groupid='"&groupid&"'"
 
            dbget.Execute sqlStr
        end if
    end if
    
    
    dim ctrNo
    if (agreeidx>0) then
        if (isNULL(ContractNo)) then
            ctrNo = getContractNoFormat(YAKGAN_FINGERS,agreeidx)
            sqlStr = " update db_partner.dbo.tbl_partner_fingers_agreeHist"
            sqlStr = sqlStr&" set ContractNo='"&ctrNo&"'"
            sqlStr = sqlStr&" where agreeidx="&agreeidx
            sqlStr = sqlStr&" and ContractType="&YAKGAN_FINGERS
            sqlStr = sqlStr&" and groupid='"&groupid&"'"
            
            dbget.Execute sqlStr
        end if
    end if
    
    checkUpcheYakganAgreeMake = (agreeidx>0)
   
    
end function


''업체 계약서 체크 후 생성
function checkUpcheContractMake(igroupid,imakerid,byref agreeidx)
    dim sqlStr
    dim viewdate, agreedate, ContractContents, ContractNo
    agreeidx = 0
    
    if (groupid="") or (imakerid="") then
        checkUpcheContractMake = FALSE
        exit function
    end if
    
    sqlStr = " IF NOT EXISTS(select * from db_partner.dbo.tbl_partner_fingers_agreeHist where groupid='"&igroupid&"' and makerid='"&imakerid&"' and contractType="&CONTRACTTYPE_FINGERS&" and deldate is NULL)"
    sqlStr = sqlStr & " BEGIN"
    sqlStr = sqlStr & " insert into db_partner.dbo.tbl_partner_fingers_agreeHist"
    sqlStr = sqlStr & " (ContractType,groupid,makerid,ContractContents)"
    sqlStr = sqlStr & " select top 1 "&CONTRACTTYPE_FINGERS&""
    sqlStr = sqlStr & " ,'"&groupid&"'"
    sqlStr = sqlStr & " ,'"&imakerid&"'"
    sqlStr = sqlStr & " ,NULL" '''ContractContents from db_partner.[dbo].[tbl_partner_contractType] where ContractType="&CONTRACTTYPE_FINGERS&""
    sqlStr = sqlStr & " END"
    dbget.Execute sqlStr
    
    sqlStr = " select top 1 agreeidx,viewdate,agreedate, ContractContents, ContractNo from db_partner.dbo.tbl_partner_fingers_agreeHist"
    sqlStr = sqlStr & " where 1=1"
    sqlStr = sqlStr & " and groupid='"&igroupid&"'"
    sqlStr = sqlStr & " and makerid='"&imakerid&"'"
    sqlStr = sqlStr & " and contractType="&CONTRACTTYPE_FINGERS&""
    sqlStr = sqlStr & " order by agreeidx desc"

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        agreeidx = rsget("agreeidx")
        viewdate = rsget("viewdate")
        agreedate = rsget("agreedate")
        ContractContents = rsget("ContractContents")
        ContractNo = rsget("ContractNo")
    end if
    rsget.close

    ''계약서 내용 생성.
    dim retContractContents
    if (agreeidx>0) then
        if (IsNULL(ContractContents)) then
            retContractContents = MakeConTractContents(agreeidx)
            sqlStr = " update db_partner.dbo.tbl_partner_fingers_agreeHist"
            sqlStr = sqlStr&" set ContractContents='"&HTML2DB(retContractContents)&"'"
            sqlStr = sqlStr&" where agreeidx="&agreeidx
            sqlStr = sqlStr&" and ContractType="&CONTRACTTYPE_FINGERS
            sqlStr = sqlStr&" and groupid='"&groupid&"'"
 
             dbget.Execute sqlStr
        end if
    end if
    
    ''계약서 번호 생성.
    dim ctrNo
    if (agreeidx>0) then
        if (isNULL(ContractNo)) then
            ctrNo = getContractNoFormat(CONTRACTTYPE_FINGERS,agreeidx)
            sqlStr = " update db_partner.dbo.tbl_partner_fingers_agreeHist"
            sqlStr = sqlStr&" set ContractNo='"&ctrNo&"'"
            sqlStr = sqlStr&" where agreeidx="&agreeidx
            sqlStr = sqlStr&" and ContractType="&CONTRACTTYPE_FINGERS
            sqlStr = sqlStr&" and groupid='"&groupid&"'"
            dbget.Execute sqlStr
        end if
    end if
    
'rw "<textarea cols=90 rows=20>"&retContractContents&"</textarea>"
    checkUpcheContractMake = (agreeidx>0)
    
end function


function MakeConTractContents(agreeidx)
    dim sqlStr
    dim oagree  , ogroupInfo
    SET oagree = New CFingersUpcheAgree
    oagree.FRectAgreeIdx = agreeidx
    oagree.getOneFingersUpcheAgree
    
    if (oagree.FresultCount<1) then
        SET oagree = Nothing
        Exit function    
    end if
    
    SET ogroupInfo = new CPartnerGroup
    ogroupInfo.FRectGroupid = oagree.FOneItem.FGroupid

    if (ogroupInfo.FRectGroupid<>"") then
        ogroupInfo.GetOneGroupInfo
    end if
    
    if (ogroupInfo.FResultCount<1) then
        SET ogroupInfo = Nothing
        SET oagree = Nothing
        exit function
    end if
    
    dim originContractContents
    sqlStr ="select top 1 ContractContents from db_partner.[dbo].[tbl_partner_contractType] where contractType="&oagree.FOneItem.FcontractType
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        originContractContents = rsget("ContractContents")
    end if
    rsget.close
    
    if (originContractContents="") then
        SET ogroupInfo = Nothing
        SET oagree = Nothing
        exit function
    end if
    
    dim detailKey,detailValue, dtlArray, intLoop
    sqlStr = " select detailKey,detailDesc from db_partner.[dbo].[tbl_partner_contractDetailType]"
    sqlStr = sqlStr&" where contractType="&oagree.FOneItem.FContractType
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        dtlArray = rsget.getRows()
    end if
    rsget.close
         
    For intLoop = 0 To UBound(dtlArray,2) 
        detailKey       = dtlArray(0,intLoop)
        detailValue     = dtlArray(1,intLoop)
        
        if (detailKey="$$A_UPCHENAME$$") then
            detailValue = "(주)텐바이텐"
        end if
        
        if (detailKey="$$A_CEONAME$$") then
            detailValue = "최은희"
        end if
        
        if (detailKey="$$A_COMPANY_NO$$") then
            detailValue = "211-87-00620"
        end if
        
        if (detailKey="$$A_COMPANY_ADDR$$") then
            detailValue = "서울시 종로구 대학로12길 31 , 2층"
        end if
        
        ''
        
        if (detailKey="$$B_UPCHENAME$$") then
            detailValue = ogroupInfo.FOneItem.Fcompany_name
        end if
        
        if (detailKey="$$B_CEONAME$$") then
            detailValue = ogroupInfo.FOneItem.Fceoname
        end if
        
        if (detailKey="$$B_COMPANY_NO$$") then
            detailValue = ogroupInfo.FOneItem.Fcompany_no  '' or getDecCompNo
        end if
        
        if (detailKey="$$B_COMPANY_ADDR$$") then
            detailValue = ogroupinfo.FOneItem.Fcompany_address & " " & ogroupinfo.FOneItem.Fcompany_address2
        end if
        
        if (detailKey="$$DEFAULT_JUNGSANDATE$$") then
            if (ogroupinfo.FOneItem.Fjungsan_date="15일") then
                detailValue = "판매(제공)월의 " & "익월 15일"
            elseif (ogroupinfo.FOneItem.Fjungsan_date="말일") then
                detailValue = "판매(제공)월의 " & "익월 말일"
            elseif (ogroupinfo.FOneItem.Fjungsan_date="수시") then
                detailValue = "판매(제공)월의 " & "익월 5일"
            elseif (ogroupinfo.FOneItem.Fjungsan_date="5일") then
                detailValue = "판매(제공)월의 " & "익월 5일"
            end if
        end if
        
        if (detailKey="$$CONTRACT_DATE$$") then
            detailValue = Left(Now(),10) 
            detailValue = Left(detailValue,4) & "년 " & Mid(detailValue,6,2) & "월 " & Mid(detailValue,9,2) & "일"
        end if
        
        if (detailKey="$$ENDDATE$$") then
          ''  detailValue = "6개월 후"
        end if
        
        if (detailKey="$$CONTRACT_CONTS$$") then
            detailValue=getFingerBrandMarginConts(oagree.FOneItem.Fmakerid, oagree.FOneItem.FcontractType)
        end if
        
        originContractContents=replace(originContractContents,detailKey,detailValue)
        ''case "$$B_BRANDNAME$$"
        ''    : getDefaultContractValue = ogroupinfo.FOneItem.Fsocname_kor
    Next            
    
    set ogroupInfo = Nothing
    set oagree = Nothing
    
    MakeConTractContents = originContractContents
end function


function getFingerBrandMarginConts(imakerid,icontractType)
    Dim sqlStr, bufStr
    Dim isMeaipContract : isMeaipContract=FALSE
    
    dim retval
    dim lec_yn, diy_yn
    dim mwdiv, mwdivName, sellplaceName, defaultmargin, defaultdeliveryType, defaultFreebeasongLimit, defaultdeliverpay
    sqlStr ="select lecturer_id , lec_yn, diy_yn, lec_margin, mat_margin, diy_margin,diy_dlv_gubun,DefaultFreeBeasongLimit,DefaultDeliveryPay"&VBCRLF
    sqlStr = sqlStr & " from [academydb].db_academy.dbo.tbl_lec_user"&VBCRLF
    sqlStr = sqlStr & " where lecturer_id='"&imakerid&"'"&VBCRLF
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        lec_yn = rsget("lec_yn")
        diy_yn = rsget("diy_yn")
        mwdiv = "U"
        mwdivName = "업체배송"
        sellplaceName ="핑거스"
        defaultmargin       = rsget("diy_margin")
        defaultdeliveryType         = rsget("diy_dlv_gubun")
        defaultFreebeasongLimit     = rsget("DefaultFreeBeasongLimit")
        defaultdeliverpay           = rsget("DefaultDeliveryPay")
    end if
    rsget.Close
    
    bufStr = ""
    if (diy_yn="Y") then
        if (icontractType=PREMIUM_CONTRACTTYPE_FINGERS) then ''프리미엄
            bufStr = bufStr & "<tr>"
            bufStr = bufStr & "<td align='center'>기본 수수료</td>"
            bufStr = bufStr & "<td align='center'>월 "&FormatNumber(CPRM_MONTHPAY,0)&getNum2KORFormat(CPRM_MONTHPAY)&" 원</td>"
            bufStr = bufStr & "<td align='center'>월정액</td>"
            bufStr = bufStr & "</tr>"
            
            bufStr = bufStr & "<tr>"
            bufStr = bufStr & "<td align='center'>상품 판매 수수료</td>"
            bufStr = bufStr & "<td align='center'>상품 판매 금액의<br>"&CLNG(CPRM_COMMISION*100)/100&getNum2KORFormat(CPRM_COMMISION)&" %</td>"
            bufStr = bufStr & "<td align='center'>배송비 제외</td>"
            bufStr = bufStr & "</tr>"
            
            bufStr = bufStr & "<tr>"
            bufStr = bufStr & "<td align='center'>결제 수수료</td>"
            bufStr = bufStr & "<td align='center'>구매 고객 결제 금액의 <br>"&CLNG(CPG_COMMISION*100)/100&getNum2KORFormat(CPG_COMMISION)&" %</td>"
            bufStr = bufStr & "<td align='center'>배송비를 포함한 구매 고객의<br>총 결제 금액을 기준으로 산정</td>"
            bufStr = bufStr & "</tr>"
        else
            bufStr = bufStr & "<tr>"
            bufStr = bufStr & "<td align='center'>상품 판매 수수료</td>"
            bufStr = bufStr & "<td align='center'>상품 판매 금액의<br>"&CLNG(defaultmargin/1.1*100)/100&getNum2KORFormat(CLNG(defaultmargin/1.1*100)/100)&" %</td>"
            bufStr = bufStr & "<td align='center'>배송비 제외</td>"
            bufStr = bufStr & "</tr>"
            
            bufStr = bufStr & "<tr>"
            bufStr = bufStr & "<td align='center'>결제 수수료</td>"
            bufStr = bufStr & "<td align='center'>구매 고객 결제 금액의 <br>"&CLNG(CPG_COMMISION*100)/100&getNum2KORFormat(CPG_COMMISION)&" %</td>"
            bufStr = bufStr & "<td align='center'>배송비를 포함한 구매 고객의<br>총 결제 금액을 기준으로 산정</td>"
            bufStr = bufStr & "</tr>"
        end if
    end if

    if (bufStr<>"") then
        if (icontractType=PREMIUM_CONTRACTTYPE_FINGERS) then
            bufStr="<thead><tr><th>항목</th><th>수수료</th><th>비고</th></tr></thead><tbody>"&bufStr
        else
            bufStr="<thead><tr><th>항목</th><th>수수료</th><th>비고</th></tr></thead><tbody>"&bufStr
        end if
        bufStr=bufStr&"</tbody>"

        
        bufStr="<table class='tMar10'><colgroup><col width='30%' /><col width='40%' /><col width='40%' /></colgroup>"&bufStr&"</table><p align='right'>(각 수수료에 대한 부가가치세는 별도)</p"

    end if
    
    getFingerBrandMarginConts = bufStr
end function

Class CFingersUpcheAgreeMasterItem

    public FcontractType
    public FctrContents
    public Fisusing
    public FsubType
    public Fregdate
    
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CFingersUpcheAgreeHistItem

    public FagreeIdx
    public FcontractType
    public Fgroupid
    public Fmakerid
    public Fregdate
    public Fviewdate
    public Fagreedate
    public FagreeRefIP
    
    public FContractNo
    public FcontractName
    public FMasterisusing
    public FsubType
    public Fdeldate
    
    public Fcompanyname
    public FContractContents
    
    public function getContractContents()
        dim retVal : retVal = FContractContents
        dim chgVal, bufVal
        
        if (FcontractType=YAKGAN_FINGERS) then
            bufVal = "약관동의일 : "
        else
            bufVal = "계약승인일 : "
        end if
        
        ''약관승인이 안되었을경우
        if isNULL(Fagreedate) then
            chgVal = Left(Now(),10) 
            chgVal = bufVal & Left(chgVal,4) & "년 " & Mid(chgVal,6,2) & "월 " & Mid(chgVal,9,2) & "일"
            
            retVal = replace(retVal,"$$AGREE_CONTRACT_DATE$$",chgVal)
        else
            chgVal = LEFT(Fagreedate,10)
            chgVal = bufVal & Left(chgVal,4) & "년 " & Mid(chgVal,6,2) & "월 " & Mid(chgVal,9,2) & "일"
            
            retVal = replace(retVal,"$$AGREE_CONTRACT_DATE$$",chgVal)
        end if
        
        
        getContractContents = retVal
    end function

    public function IsPrivContractAddItem
        IsPrivContractAddItem = FALSE
        Exit function
        IsPrivContractAddItem = (FcontractType=YAKGAN_FINGERS)
    end function
    
    public function getEkey()
        getEkey = MD5(ChashVal&FagreeIdx&Fgroupid)
    end function
    
    public function getPdfDownLinkUrl()
        getPdfDownLinkUrl = getPdfDownLinkUrlAdm&"&chkcf=1"  ''업체가 다운로드시 업체확인체크 위한 플래그
    end function
    
    public function getPdfDownLinkUrlAdm()
        dim addparam
        addparam = "?agreeIdx="&FagreeIdx
        addparam = addparam&"&gkey="&Fgroupid
        addparam = addparam&"&ctrNo="&FContractNo
        addparam = addparam&"&cTp="&FcontractType 
        addparam = addparam&"&vTp="&"d"                                    ''뷰인지 다운로드인지. (d 다운로드,else 뷰)
        addparam = addparam&"&pTp="&CHKIIF(IsPrivContractAddItem,"1","")   ''개인정보 수집존재
        addparam = addparam&"&ekey="&getEkey

        if (application("Svr_Info")	= "Dev") then
            getPdfDownLinkUrlAdm = "http://testwebadmin.10x10.co.kr/admin/member/contract/fingers/dnContractPdf_Fingers.asp"&addparam
        else
            getPdfDownLinkUrlAdm = "http://apps.10x10.co.kr/pdf/dnContractPdf_Fingers.asp"&addparam
        end if
    end function

    public function isDeletedContract()
        isDeletedContract = (NOT isNULL(Fdeldate))
    end function
    
    public function IsAgreeFinished()
        IsAgreeFinished = NOT isNULL(Fagreedate)
    end function
    
    public function getAgreeText()
        getAgreeText = Fagreedate & " 에 동의 하셨습니다."
    end function
    
    public function getContractTypeAgreeName()
        getContractTypeAgreeName = ""
        if (FcontractType=YAKGAN_FINGERS) then
            getContractTypeAgreeName = "약관"
        elseif (FcontractType=CONTRACTTYPE_FINGERS) then
            getContractTypeAgreeName = "계약서"
        end if
    end function

    public function getAgreeStateName
        if NOT isNULL(Fagreedate) then
            getAgreeStateName = "동의완료"
        end if
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CFingersUpcheAgree
    public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	
	public FRectAgreeIdx
	public FRectContractType
    public FRectGroupID
    public FRectMakerid
    public FRectagreeState
    public FRectDelInclude
    
    public Sub getOneFingersUpcheAgree()
        dim sqlStr, i
        dim addSQL
        
        addSQL = addSQL &" and H.agreeIdx="&FRectAgreeIdx&""
        
        if (FRectContractType<>"") then
            addSQL = addSQL &" and H.contractType="&FRectContractType&""
        end if
        
        if (FRectMakerid<>"") then
            addSQL = addSQL &" and (H.contractType="&YAKGAN_FINGERS&" or H.makerid='"&FRectMakerid&"')"
        end if
        
        if (FRectGroupID<>"") then
            addSQL = addSQL &" and H.groupid='"&FRectGroupID&"'"
        end if
    
        if (FRectagreeState="0") then
            addSQL = addSQL &" and H.viewdate is Not NULL"
        elseif (FRectagreeState="1") then
            addSQL = addSQL &" and H.agreedate is NULL"
        elseif (FRectagreeState="3") then
            addSQL = addSQL &" and H.agreedate is Not NULL"
        end if
        
        if (FRectDelInclude="on") then
                
        else
            addSQL = addSQL &" and H.deldate is NULL"
        end if
        
        sqlStr = "select H.agreeIdx, H.contractType, H.groupid, H.makerid, H.viewdate, H.agreedate" &vbCRLF
        sqlStr = sqlStr & ", H.agreeRefIP,H.ContractNo,H.deldate,M.contractName,H.ContractNo,H.ContractContents,H.regdate,M.isusing as Masterisusing" &vbCRLF
        sqlStr = sqlStr & ", g.company_name as companyname"
        sqlStr = sqlStr & " from db_partner.[dbo].[tbl_partner_fingers_agreeHist] H" &vbCRLF
        sqlStr = sqlStr & "     Join db_partner.[dbo].[tbl_partner_contractType] M" &vbCRLF
        sqlStr = sqlStr & "     on H.contractType=M.contractType"&vbCRLF
        sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner_group g"&vbCRLF
        sqlStr = sqlStr & "     on H.groupid=g.groupid"
        sqlStr = sqlStr & " where 1=1"&vbCRLF
        sqlStr = sqlStr & addSQL
        sqlStr = sqlStr & " order by H.agreeIdx desc"

		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0


		if Not rsget.Eof then
			set FOneItem = new CFingersUpcheAgreeHistItem

			FOneItem.FagreeIdx  = rsget("agreeIdx")
            FOneItem.FcontractType    = rsget("contractType")
            FOneItem.Fgroupid   = rsget("groupid")
            FOneItem.Fmakerid   = rsget("makerid")
            FOneItem.Fregdate   = rsget("regdate")
            FOneItem.Fviewdate  = rsget("viewdate")
            FOneItem.Fagreedate = rsget("agreedate")
            FOneItem.FagreeRefIP= rsget("agreeRefIP")
            FOneItem.FcontractName  = rsget("contractName")
            FOneItem.FMasterisusing   = rsget("Masterisusing")
            
            FOneItem.Fcompanyname   = rsget("companyname")
            FOneItem.FContractContents = rsget("ContractContents")
            FOneItem.FContractNo = rsget("ContractNo")
            
            FOneItem.Fdeldate = rsget("deldate")
		end if
		rsget.close
		
    end Sub

    public Sub GetFingersUpcheAgreeHistList_UpcheView()
        dim sqlStr, i
        dim addSQL
        
        if (FRectAgreeIdx<>"") then
            addSQL = addSQL &" and H.agreeIdx="&FRectAgreeIdx&""
        end if 
        
        if (FRectContractType<>"") then
            addSQL = addSQL &" and H.contractType="&FRectContractType&""
        end if
        
        
        addSQL = addSQL &" and (H.contractType="&YAKGAN_FINGERS&" or H.makerid='"&FRectMakerid&"')"
        
        
        ''필수
        addSQL = addSQL &" and H.groupid='"&FRectGroupID&"'"
                
        if (FRectagreeState="0") then
            addSQL = addSQL &" and H.viewdate is Not NULL"
        elseif (FRectagreeState="1") then
            addSQL = addSQL &" and H.agreedate is NULL"
        elseif (FRectagreeState="3") then
            addSQL = addSQL &" and H.agreedate is Not NULL"
        end if
        
        if (FRectDelInclude="on") then
                
        else
            addSQL = addSQL &" and H.deldate is NULL"
        end if
        
        sqlStr = "select  H.agreeIdx, H.contractType, H.groupid, H.makerid, H.viewdate, H.agreedate" &vbCRLF
        sqlStr = sqlStr & ", H.agreeRefIP,H.ContractNo,H.deldate,M.contractName,M.isusing as Masterisusing, H.regdate" &vbCRLF
        sqlStr = sqlStr & ", g.company_name as companyname"
        sqlStr = sqlStr & " from db_partner.[dbo].[tbl_partner_fingers_agreeHist] H" &vbCRLF
        sqlStr = sqlStr & "     Join db_partner.[dbo].[tbl_partner_contractType] M" &vbCRLF
        sqlStr = sqlStr & "     on H.contractType=M.contractType"&vbCRLF
        sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner_group g"&vbCRLF
        sqlStr = sqlStr & "     on H.groupid=g.groupid"
        sqlStr = sqlStr & " where 1=1"&vbCRLF
        sqlStr = sqlStr & addSQL
        sqlStr = sqlStr & " order by H.agreeIdx desc"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new CFingersUpcheAgreeHistItem

				FItemList(i).FagreeIdx  = rsget("agreeIdx")
                FItemList(i).FcontractType    = rsget("contractType")
                FItemList(i).Fgroupid   = rsget("groupid")
                FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).Fregdate   = rsget("regdate")
                FItemList(i).Fviewdate  = rsget("viewdate")
                FItemList(i).Fagreedate = rsget("agreedate")
                FItemList(i).FagreeRefIP= rsget("agreeRefIP")
                FItemList(i).FcontractName  = rsget("contractName")
                FItemList(i).FMasterisusing   = rsget("Masterisusing")
                
                FItemList(i).Fcompanyname   = rsget("companyname")
                FItemList(i).FContractNo = rsget("ContractNo")
                FItemList(i).Fdeldate = rsget("deldate")
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub
    
    
    public Sub GetFingersUpcheAgreeHistList()
        dim sqlStr, i
        dim addSQL
        
        if (FRectAgreeIdx<>"") then
            addSQL = addSQL &" and H.agreeIdx="&FRectAgreeIdx&""
        end if 
        
        if (FRectContractType<>"") then
            addSQL = addSQL &" and H.contractType="&FRectContractType&""
        end if
        
        if (FRectMakerid<>"") then
            addSQL = addSQL &" and H.makerid='"&FRectMakerid&"'"
        end if
        
        if (FRectGroupID<>"") then
            addSQL = addSQL &" and H.groupid='"&FRectGroupID&"'"
        end if
        
        if (FRectagreeState="0") then
            addSQL = addSQL &" and H.viewdate is Not NULL"
        elseif (FRectagreeState="1") then
            addSQL = addSQL &" and H.agreedate is NULL"
        elseif (FRectagreeState="3") then
            addSQL = addSQL &" and H.agreedate is Not NULL"
        end if
        
        ''삭제 검색
        if (FRectDelInclude="on") then
                
        else
            addSQL = addSQL &" and H.deldate is NULL"
        end if
        
        sqlStr = "select top "&FPageSize&" H.agreeIdx, H.contractType, H.groupid, H.makerid, H.viewdate, H.agreedate" &vbCRLF
        sqlStr = sqlStr & ", H.agreeRefIP,H.ContractNo,H.deldate,M.contractName,M.isusing as Masterisusing,H.regdate" &vbCRLF
        sqlStr = sqlStr & ", g.company_name as companyname"
        sqlStr = sqlStr & " from db_partner.[dbo].[tbl_partner_fingers_agreeHist] H" &vbCRLF
        sqlStr = sqlStr & "     Join db_partner.[dbo].[tbl_partner_contractType] M" &vbCRLF
        sqlStr = sqlStr & "     on H.contractType=M.contractType"&vbCRLF
        sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner_group g"&vbCRLF
        sqlStr = sqlStr & "     on H.groupid=g.groupid"
        sqlStr = sqlStr & " where 1=1"&vbCRLF
        sqlStr = sqlStr & addSQL
        sqlStr = sqlStr & " order by H.agreeIdx desc"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new CFingersUpcheAgreeHistItem

				FItemList(i).FagreeIdx  = rsget("agreeIdx")
                FItemList(i).FcontractType    = rsget("contractType")
                FItemList(i).Fgroupid   = rsget("groupid")
                FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).Fregdate  = rsget("regdate")
                FItemList(i).Fviewdate  = rsget("viewdate")
                FItemList(i).Fagreedate = rsget("agreedate")
                FItemList(i).FagreeRefIP= rsget("agreeRefIP")
                FItemList(i).FcontractName  = rsget("contractName")
                FItemList(i).FMasterisusing   = rsget("Masterisusing")
                
                FItemList(i).Fcompanyname   = rsget("companyname")
                FItemList(i).FContractNo = rsget("ContractNo")
                FItemList(i).Fdeldate = rsget("deldate")
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub
    
    public Sub GetFingersUpcheAgreeMasterList()
    
    end Sub
    
	Private Sub Class_Terminate()
	End Sub
	
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
	
end Class	
%>