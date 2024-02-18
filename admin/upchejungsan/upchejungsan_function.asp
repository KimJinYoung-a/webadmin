<%
public function getJGubunName(ijgubun)
    if isNULL(ijgubun) then Exit function
    if (ijgubun="MM") then
        getJGubunName = "����"
    elseif (ijgubun="CC") then
        getJGubunName = "<font color=blue>������</font>"
    else
        getJGubunName = ijgubun
    end if

end function

public function GetHoldingJungSanSum(ototalsum)
    dim TreePercentTax
    TreePercentTax = Fix(Fix(ototalsum*0.03)/10)*10

    GetHoldingJungSanSum = ototalsum

    ''3%������ 1000�� �����̸� ���ݾ���. =>�̸�(2018/10/01)  	artbookjs case
    ''if ABS(TreePercentTax)<=1000 then Exit function
    if (TreePercentTax<1000) and (TreePercentTax>-1000) then Exit function

    GetHoldingJungSanSum = ototalsum - TreePercentTax - Fix(Fix(TreePercentTax*0.1)/10)*10

end function

function fnGetJFixIpkumListSum(ipFileNo)
    '' ** �հ� ���������� �Ұ�..! == �Աݾ��ε� ����.
    '' ����ó(�ŷ�ó�ڵ�), �Ա�����, �Աݰ���, ��ü�ݾ�, ��������μ⳻��(�ŷ�ó�� ��ũ����)==�����ָ�?,�Ա������μ⳻��((��)�ٹ�����)
    '' 1. ���� ���� �ۼ� / �뷮 ������ ��ȸ ���� ��/ ���⿹�� or �뷮��ü ����

    Dim sqlStr
    sqlStr = " Select T.groupid"
    sqlStr = sqlStr & ", (CASE WHEN T.ipkumbank='����' THEN '����' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='��ȯ' THEN 'KEB�ϳ�' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='ȫ�ἧ����' THEN 'HSBC' " ''2017/01/25
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='��Ƽ' THEN '��Ƽ' "
	''sqlStr = sqlStr & "	 WHEN T.ipkumbank='�������ݰ�' THEN '������' "
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='����' THEN '���Ĵٵ���Ÿ��' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='���̹�ũ' THEN 'K��ũ' "
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='��������' THEN '����' ELSE T.ipkumbank END)"   ''����, SC���� ���Ĵٵ���Ÿ��
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
    sqlStr = sqlStr & ", (CASE WHEN T.ipkumbank='����' THEN '����' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='��ȯ' THEN 'KEB�ϳ�' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='ȫ�ἧ����' THEN 'HSBC' " ''2017/01/25
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='��Ƽ' THEN '��Ƽ' "
	''sqlStr = sqlStr & "	 WHEN T.ipkumbank='�������ݰ�' THEN '������' "
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='����' THEN 'SC����' "
    sqlStr = sqlStr & "	 WHEN T.ipkumbank='���̹�ũ' THEN 'K��ũ' "
	sqlStr = sqlStr & "	 WHEN T.ipkumbank='��������' THEN '����' ELSE T.ipkumbank END)"
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
    ''--����Ʈ
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
    Dim payRequestTitle : payRequestTitle ="��ǰ���� �������-"

    IF (eapppartIdx="0000000101") THEN payRequestTitle=payRequestTitle&"�¶��λ����"
    IF (eapppartIdx="0000000201") THEN payRequestTitle=payRequestTitle&"�������λ����"

    Dim arap_cd : arap_cd="106" ''��ǰ����.

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
    Dim returnValue   ''' ������û�� IDX
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

'		'�������� ���
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

    '�μ��� �ڱݱ��� ���
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

    '''�μ��� �ڱݱ��� ����. // (1) 0000000101 / (2) 0000000201
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