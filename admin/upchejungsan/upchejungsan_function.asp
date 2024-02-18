<%
public function getJGubunName(ijgubun)
    if isNULL(ijgubun) then Exit function
    if (ijgubun="MM") then
        getJGubunName = "매입"
    elseif (ijgubun="CC") then
        getJGubunName = "<font color=blue>수수료</font>"
    else
        getJGubunName = ijgubun
    end if

end function

public function GetHoldingJungSanSum(ototalsum)
    dim TreePercentTax
    TreePercentTax = Fix(Fix(ototalsum*0.03)/10)*10

    GetHoldingJungSanSum = ototalsum

    ''3%세금이 1000원 이하이면 세금없음. =>미만(2018/10/01)  	artbookjs case
    ''if ABS(TreePercentTax)<=1000 then Exit function
    if (TreePercentTax<1000) and (TreePercentTax>-1000) then Exit function

    GetHoldingJungSanSum = ototalsum - TreePercentTax - Fix(Fix(TreePercentTax*0.1)/10)*10

end function

function fnGetJFixIpkumListSum(ipFileNo)
    '' ** 합계 쿼리에주의 할것..! == 입금업로드 파일.
    '' 지급처(거래처코드), 입금은행, 입금계좌, 이체금액, 출금통장인쇄내용(거래처명 블랭크제외)==예금주명?,입금통장인쇄내용((주)텐바이텐)
    '' 1. 엑셀 파일 작성 / 대량 예금주 조회 실행 후/ 지출예정 or 대량이체 실행

    Dim sqlStr
    sqlStr = " Select T.groupid"
    sqlStr = sqlStr & ", (CASE WHEN T.ipkumbank='조흥' THEN '신한' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='외환' THEN 'KEB하나' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='홍콩샹하이' THEN 'HSBC' " ''2017/01/25
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='시티' THEN '씨티' "
	''sqlStr = sqlStr & "	 WHEN T.ipkumbank='새마을금고' THEN '새마을' "
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='제일' THEN '스탠다드차타드' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='케이뱅크' THEN 'K뱅크' "
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='단위농협' THEN '농협' ELSE T.ipkumbank END)"   ''제일, SC제일 스탠다드차타드
    sqlStr = sqlStr & " , T.ipkumacctno, Sum(JSum) as jsum, IsNULL(refipFileDetailIdx,ipFileDetailIdx) as ipFileDetailIdx"
    sqlStr = sqlStr & " ,Replace(G.company_name,' ','') as company_name , Replace(G.jungsan_acctname,' ','') as jungsan_acctname"
    sqlStr = sqlStr & " From ("
    sqlStr = sqlStr & " 	select (CASE WHEN F.targetGbn='ON' THEN m.groupid"
    sqlStr = sqlStr & " 			WHEN F.targetGbn='OF' THEN j.groupid"
    sqlStr = sqlStr & " 			ELSE '' END) as groupid"
    sqlStr = sqlStr & " 	,(CASE WHEN F.targetGbn='ON' THEN m.ipkum_bank"
    sqlStr = sqlStr & " 			WHEN F.targetGbn='OF' THEN j.ipkum_bank"
    sqlStr = sqlStr & " 			ELSE '' END) as ipkumbank"
    sqlStr = sqlStr & " 	,(CASE WHEN F.targetGbn='ON' THEN m.ipkum_acctno"
    sqlStr = sqlStr & " 			WHEN F.targetGbn='OF' THEN j.ipkum_acctno"
    sqlStr = sqlStr & " 			ELSE '' END) as ipkumacctno"
    sqlStr = sqlStr & " 	,(CASE WHEN F.targetGbn='ON' THEN m.ub_totalsuplycash+m.me_totalsuplycash+m.wi_totalsuplycash+m.et_totalsuplycash+m.sh_totalsuplycash+m.dlv_totalsuplycash  "
    sqlStr = sqlStr & " 			WHEN F.targetGbn='OF' THEN J.tot_jungsanprice"
    sqlStr = sqlStr & " 			ELSE 0 END) as JSum"
    sqlStr = sqlStr & " 	,F.ipFileDetailIDx"
    sqlStr = sqlStr & " 	,F.refipFileDetailIDx"
    sqlStr = sqlStr & " 	from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail  F"
    sqlStr = sqlStr & " 		Left Join [db_jungsan].[dbo].tbl_designer_jungsan_master m"
    sqlStr = sqlStr & " 		On F.targetGbn='ON'"
    sqlStr = sqlStr & " 		and F.targetIdx=m.id"
    sqlStr = sqlStr & " 		Left Join [db_jungsan].[dbo].tbl_off_jungsan_master J"
    sqlStr = sqlStr & " 		On F.targetGbn='OF'"
    sqlStr = sqlStr & " 		and F.targetIdx=J.idx"
    sqlStr = sqlStr & " 	where F.ipFileNo="&ipFileNo
    sqlStr = sqlStr & " ) T"
    sqlStr = sqlStr & " 	Left Join [db_partner].dbo.tbl_partner_group G"
    sqlStr = sqlStr & " 	On T.groupid=G.groupid"
    sqlStr = sqlStr & " group by T.groupid"
    sqlStr = sqlStr & ", (CASE WHEN T.ipkumbank='조흥' THEN '신한' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='외환' THEN 'KEB하나' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='홍콩샹하이' THEN 'HSBC' " ''2017/01/25
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='시티' THEN '씨티' "
	''sqlStr = sqlStr & "	 WHEN T.ipkumbank='새마을금고' THEN '새마을' "
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='제일' THEN 'SC제일' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='케이뱅크' THEN 'K뱅크' "
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='단위농협' THEN '농협' ELSE T.ipkumbank END)"
	sqlStr = sqlStr & "	 , T.ipkumbank, T.ipkumacctno,IsNULL(refipFileDetailIdx,ipFileDetailIdx) ,G.company_name, G.jungsan_acctname"
    sqlStr = sqlStr & " order by T.groupid,ipFileDetailIdx"

    'rsget.Open sqlStr,dbget,1
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    
    IF Not rsget.Eof THEN
        fnGetJFixIpkumListSum = rsget.getRows
    ENd IF
    rsget.Close

end function

function fnGetJFixIpkumList(ipFileNo)
    ''--리스트
    Dim sqlStr
    sqlStr = " select F.ipFileDetailIdx, F.targetGbn,F.targetIdx,F.ipfileDetailState"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN m.ub_totalsuplycash+m.me_totalsuplycash+m.wi_totalsuplycash+m.et_totalsuplycash+m.sh_totalsuplycash+m.dlv_totalsuplycash  "
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN J.tot_jungsanprice"
    sqlStr = sqlStr & " 		ELSE 0 END) as JSum"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN convert(varchar(10),m.taxregdate,21)"
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN convert(varchar(10),J.taxregdate,21)"
    sqlStr = sqlStr & " 		ELSE '' END) as TaxDate"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN m.yyyymm"
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN j.yyyymm"
    sqlStr = sqlStr & " 		ELSE '' END) as yyyymm"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN m.groupid"
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN j.groupid"
    sqlStr = sqlStr & " 		ELSE '' END) as groupid"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN m.designerid"
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN j.makerid"
    sqlStr = sqlStr & " 		ELSE '' END) as makerid"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN m.ipkum_bank"
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN j.ipkum_bank"
    sqlStr = sqlStr & " 		ELSE '' END) as ipkumbank"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN m.ipkum_acctno"
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN j.ipkum_acctno"
    sqlStr = sqlStr & " 		ELSE '' END) as ipkumacctno"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN G.jungsan_date"
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN G.jungsan_date_OFF"
    sqlStr = sqlStr & " 		ELSE '' END) as jungsan_date"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN m.finishflag"
	sqlStr = sqlStr & "	        WHEN F.targetGbn='OF' THEN j.finishflag"
	sqlStr = sqlStr & " 	    ELSE 0 END) as finishflag"
    sqlStr = sqlStr & " ,(select count(*) from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail O with (nolock) where O.refipFileDetailIdx=F.ipFileDetailIdx) as smrCnt"
    sqlStr = sqlStr & " ,F.refipFileDetailidx"
    sqlStr = sqlStr & " ,G.jungsan_acctname,G.company_name"
    sqlStr = sqlStr & " ,G.erpCust_cd,G.erpUsing"
    sqlStr = sqlStr & " ,B.CUST_CD as ERPCUSTCD"
    sqlStr = sqlStr & " ,(CASE WHEN F.targetGbn='ON' THEN m.jgubun"
	sqlStr = sqlStr & "	        WHEN F.targetGbn='OF' THEN j.jgubun"
	sqlStr = sqlStr & " 	    ELSE '' END) as jgubun"
    sqlStr = sqlStr & " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail  F with (nolock)"
    sqlStr = sqlStr & " 	Left Join [db_jungsan].[dbo].tbl_designer_jungsan_master m with (nolock)"
    sqlStr = sqlStr & " 	On F.targetGbn='ON'"
    sqlStr = sqlStr & " 	and F.targetIdx=m.id"
    sqlStr = sqlStr & " 	Left Join [db_jungsan].[dbo].tbl_off_jungsan_master J with (nolock)"
    sqlStr = sqlStr & " 	On F.targetGbn='OF'"
    sqlStr = sqlStr & " 	and F.targetIdx=J.idx"
    sqlStr = sqlStr & " 	Left Join [db_partner].dbo.tbl_partner_group G with (nolock)"
    sqlStr = sqlStr & " 	On (CASE WHEN F.targetGbn='ON' THEN m.groupid"
    sqlStr = sqlStr & " 		WHEN F.targetGbn='OF' THEN j.groupid"
    sqlStr = sqlStr & " 		ELSE '' END)=G.groupid"
    sqlStr = sqlStr & " 	Left Join db_partner.dbo.tbl_TMS_BA_CUST B with (nolock)"
    sqlStr = sqlStr & " 	On isNULL(G.erpCust_cd,G.groupid)=B.CUST_CD"
    sqlStr = sqlStr & " 	and B.USE_YN='Y' and B.DEL_YN='N'"
    sqlStr = sqlStr & " where F.ipFileNo="&ipFileNo
    sqlStr = sqlStr & " order by groupid, IsNULL(F.refipFileDetailIDx,ipFileDetailIDx), smrCnt desc, yyyymm, taxdate"

	'response.write sqlStr & "<br>"
    'rsget.Open sqlStr,dbget,1
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    
    IF Not rsget.Eof THEN
        fnGetJFixIpkumList = rsget.getRows
    ENd IF
    rsget.Close
end Function

function MakePayReq(payrequestdate, payRequestPrice, eapppartIdx)
    Dim objCmd
    Dim payrequesttype : payrequesttype = 9
    Dim reportidx : reportidx = 0
    Dim payRequestTitle : payRequestTitle ="상품매입 정기결제-"

    IF (eapppartIdx="0000000101") THEN payRequestTitle=payRequestTitle&"온라인사업부"
    IF (eapppartIdx="0000000201") THEN payRequestTitle=payRequestTitle&"오프라인사업부"

    Dim arap_cd : arap_cd="106" ''상품매입.

    Dim payrequestState : payrequestState =1
    Dim adminId : adminId= session("ssBctId")

    Dim cust_cd : cust_cd=""
    Dim inBank  : inBank=""
    Dim accountNo : accountNo=""
    Dim accountHolder : accountHolder=""
    Dim divMoney : divMoney = 2000000
    Dim Comment : Comment =""

    Dim authId1 : authId1=""
    Dim authId2 : authId2=""
	Dim authposition : authposition=0
    Dim authState : authState=0
    Dim isLast : isLast=1
    Dim returnValue   ''' 결제요청서 IDX
    Dim payrequestidx : payrequestidx=0
    Dim partMoney   : partMoney = payRequestPrice

    IF payrequestPrice >= divMoney THEN
		authstate = 0
	ELSE
		authstate = 1
	END IF

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_Insert]( "&payrequesttype&","&reportidx&" ,'"&payRequestTitle&"',"&arap_cd&",'"&payrequestdate&"', '"&payrequestPrice&"'"&_
						+",'"&cust_cd&"','"&InBank&"','"&accountNo&"','"&accountHolder&"',"&payrequestState&",'"&adminId&"','"&authId1&"','"&authId2&"',"&authstate&",'"&Comment&"','"&divMoney&"')}"

'			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_Insert]( "&payrequesttype&","&reportidx&" ,'"&payRequestTitle&"',"&arap_cd&",'"&payrequestdate&"', '"&payrequestPrice&"'"&_
'						+",'"&cust_cd&"','"&InBank&"','"&accountNo&"','"&accountHolder&"',"&payrequestState&",'"&adminId&"','"&authId&"',"&authposition&","&authstate&","&isLast&_
'						+",'"&Comment&"','"&divMoney&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue > 0 THEN
		payrequestidx = returnValue

'		'증빙서류 등록
'		Set objCmd = Server.CreateObject("ADODB.COMMAND")
'		With objCmd
'			.ActiveConnection = dbget
'			.CommandType = adCmdText
'			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayDoc_Insert]("&payrequestidx&",'"&iDockind&"','"&sVatKind&"','"&dIssuedate&"','"&sItemName&"','"&mTotPrice&"','"&mSupplyPrice&"','"&mVatPrice&"','"&setaxkey&"','"&sDocbigo&"','"&sfile2&"','"&adminid&"')}"
'			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
'			.Execute, , adExecuteNoRecords
'			End With
'		    returnValue = objCmd(0).Value
'		Set objCmd = nothing

		MakePayReq = payrequestidx
	ELSE
	    MakePayReq =payrequestidx
	END IF
end function

function AddEappPartMoney(eapppartIdx,partMoney,payrequestidx)
    Dim reportIdx : reportIdx =0
    Dim returnValue
    Dim objCmd

    '부서별 자금구분 등록
	IF eapppartIdx <> "" THEN
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,"&payrequestidx&",'"&eapppartIdx&"','"&partMoney&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
	END IF
	AddEappPartMoney = returnValue
end function


function RecalcuPayRequestPrice(payRequestIdx)
    Dim sqlstr
    Dim onSum, offSum
    OnSum = 0
    offSum = 0
    sqlstr = " select "
    sqlstr = sqlstr + " IsNULL(sum(M.ub_totalsuplycash+ M.me_totalsuplycash+M.wi_totalsuplycash+M.et_totalsuplycash+M.sh_totalsuplycash+M.dlv_totalsuplycash),0) as OnSum"
    sqlstr = sqlstr + " , IsNULL(sum(M2.tot_jungsanprice),0)  as offSum"
    sqlstr = sqlstr + " From db_partner.dbo.tbl_eAppPayRequest R"
    sqlstr = sqlstr + " 	 Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Master F"
    sqlstr = sqlstr + " 	 On R.payRequestIdx=F.payreqIdx"
    sqlstr = sqlstr + " 	 Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
    sqlstr = sqlstr + " 	 on F.ipfileno=S.ipfileno"
    sqlstr = sqlstr + " 	 left Join  [db_jungsan].[dbo].tbl_designer_jungsan_master M"
    sqlstr = sqlstr + " 	 on S.targetIdx=M.id and S.targetGbn='ON' "
    sqlstr = sqlstr + " 	 left Join  [db_jungsan].[dbo].tbl_off_jungsan_master M2"
    sqlstr = sqlstr + " 	 on S.targetIdx=M2.idx and S.targetGbn='OF' "
    sqlstr = sqlstr + " where R.payRequestIdx="&payRequestIdx

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        OnSum = rsget("OnSum")
        offSum = rsget("offSum")
    end if
    rsget.Close


    sqlstr = " Update db_partner.dbo.tbl_eAppPayRequest"
    sqlstr = sqlstr + " set payRequestPrice="& (OnSum+offSum)
    sqlstr = sqlstr + " where payRequestIdx="&payRequestIdx
    dbget.Execute sqlStr

    '''부서별 자금구분 수정. // (1) 0000000101 / (2) 0000000201
    IF (OnSum<>0) THEN
        sqlstr = " IF Exists(select * from db_partner.dbo.tbl_eAppPartMoney P where P.payRequestIdx="&payRequestIdx&" and P.BIZSection_CD='0000000101')"
        sqlstr = sqlstr & " BEGIN"
        sqlstr = sqlstr & "     update db_partner.dbo.tbl_eAppPartMoney SET partmoney="&OnSum&" where payRequestIdx="&payRequestIdx&" AND BIZSection_CD='0000000101'"
        sqlstr = sqlstr & " END"
        sqlstr = sqlstr & " ELSE"
        sqlstr = sqlstr & " BEGIN"
        sqlstr = sqlstr & "     Insert into db_partner.dbo.tbl_eAppPartMoney"
        sqlstr = sqlstr & "     (reportIdx, payRequestIdx, BizSection_CD, partMoney, isUsing)"
        sqlstr = sqlstr & "     Values(0,"&payRequestIdx&",'0000000101',"&OnSum&",1)"
        sqlstr = sqlstr & " END"
        dbget.Execute sqlStr
    ENd IF

    IF (offSum<>0) THEN
        sqlstr = " IF Exists(select * from db_partner.dbo.tbl_eAppPartMoney P where P.payRequestIdx="&payRequestIdx&" and P.BIZSection_CD='0000000201')"
        sqlstr = sqlstr & " BEGIN"
        sqlstr = sqlstr & "     update db_partner.dbo.tbl_eAppPartMoney SET partmoney="&offSum&" where payRequestIdx="&payRequestIdx&" AND BIZSection_CD='0000000201'"
        sqlstr = sqlstr & " END"
        sqlstr = sqlstr & " ELSE"
        sqlstr = sqlstr & " BEGIN"
        sqlstr = sqlstr & "     Insert into db_partner.dbo.tbl_eAppPartMoney"
        sqlstr = sqlstr & "     (reportIdx, payRequestIdx, BizSection_CD, partMoney, isUsing)"
        sqlstr = sqlstr & "     Values(0,"&payRequestIdx&",'0000000201',"&offSum&",1)"
        sqlstr = sqlstr & " END"
        dbget.Execute sqlStr
    ENd IF
end function

%>