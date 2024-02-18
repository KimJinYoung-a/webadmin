<%

function getLGD_FINANCECODE2Name(fCode)
    select Case fCode
        CASE "11" : getLGD_FINANCECODE2Name = "����"
        CASE "06" : getLGD_FINANCECODE2Name = "����"
        CASE "20" : getLGD_FINANCECODE2Name = "�츮"
        CASE "26" : getLGD_FINANCECODE2Name = "����"
        CASE "81" : getLGD_FINANCECODE2Name = "�ϳ�"
        CASE "03" : getLGD_FINANCECODE2Name = "���"
        CASE "05" : getLGD_FINANCECODE2Name = "��ȯ"
        CASE "39" : getLGD_FINANCECODE2Name = "�泲"
        CASE "32" : getLGD_FINANCECODE2Name = "�λ�"
        CASE "71" : getLGD_FINANCECODE2Name = "��ü��"
        CASE "07" : getLGD_FINANCECODE2Name = "����"
        CASE "31" : getLGD_FINANCECODE2Name = "�뱸"
        CASE ELSE : getLGD_FINANCECODE2Name = ""
    end Select
end function

function CheckNChangeCyberAcct(iorderserial)
    dim sqlStr
    dim ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, OLDsubtotalPrice, OLDCancelyn, sumPaymentEtc
    ipkumdiv = 0
    OLDsubtotalPrice = 0
    OLDCancelyn      = ""

    CheckNChangeCyberAcct = false

    sqlStr = " select orderserial, ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, sumPaymentEtc"
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_master"
    sqlStr = sqlStr & " where orderserial='" & iorderserial & "'"

    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        ipkumdiv    = rsget("ipkumdiv")
		accountdiv  = rsget("accountdiv")
		accountNo   = rsget("accountNo")
		cancelyn    = rsget("cancelyn")
		subtotalPrice = rsget("subtotalPrice")
        sumPaymentEtc = rsget("sumPaymentEtc")
    end if
	rsget.close

	if (ipkumdiv<>2) then Exit function
	if (accountdiv<>"7") then Exit function

	if (accountNo="���� 470301-01-014754") _
        or (accountNo="���� 100-016-523130") _
        or (accountNo="�츮 092-275495-13-001") _
        or (accountNo="�ϳ� 146-910009-28804") _
        or (accountNo="��� 277-028182-01-046") _
        or (accountNo="���� 029-01-246118") then
            Exit function
    end if

    dim CLOSEDATE
    if (cancelyn<>"N") then
        CLOSEDATE = Replace(Left(CStr(now()),10),"-","") & "000000"
    else
        CLOSEDATE = Replace(Left(CStr(DateAdd("d",10,now())),10),"-","") & "235959"
    end if

    sqlStr = " select top 1 subtotalPrice, convert(varchar(19),CLOSEDATE,20) as CLOSEDATE "
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_CyberAccountLog"
    sqlStr = sqlStr & " where orderserial='" & iorderserial & "'"
    sqlStr = sqlStr & " order by differencekey desc"
    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        OLDsubtotalPrice = rsget("subtotalPrice")
        OLDCancelyn      = rsget("CLOSEDATE")

        if (RIGHT(OLDCancelyn,8)="00:00:00") then
            OLDCancelyn="Y"
        else
            OLDCancelyn="N"
        end if
    end if
    rsget.close

    if (OLDsubtotalPrice<>subtotalPrice) or (OLDCancelyn<>Cancelyn) then
        '// ���÷��� ���۽ÿ��� �������� ���ݾ� ����
        CheckNChangeCyberAcct = ChangeCyberAcct(iorderserial, subtotalPrice-sumPaymentEtc, CLOSEDATE)
    end if
end function

function CheckNAssignCyberAcct(asid, iorderserial, CyberAcctCode)
	'// CyberAcctCode = �Աݰ��������ڵ�
    dim sqlStr
    dim ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, goodname, buyname, accountname, buyhp, buyemail, userid
    ipkumdiv = 0

    CheckNAssignCyberAcct = false

    sqlStr = " select orderserial, ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, buyname, accountname, buyhp, buyemail, userid "
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_master"
    sqlStr = sqlStr & " where orderserial='" & iorderserial & "'"

    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        ipkumdiv    = rsget("ipkumdiv")
		accountdiv  = rsget("accountdiv")
		accountNo   = rsget("accountNo")
		cancelyn    = rsget("cancelyn")
		subtotalPrice = rsget("subtotalPrice")
		buyname    	= rsget("buyname")
		accountname	= rsget("accountname")
		buyhp    	= rsget("buyhp")
		buyemail    = rsget("buyemail")
		userid    	= rsget("userid")
    end if
	rsget.close

	if (ipkumdiv<>0) then Exit function
	if (accountdiv<>"7") then Exit function

    sqlStr = " select max(itemname) as itemname, count(*) as cnt "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " [db_order].[dbo].[tbl_order_detail] "
    sqlStr = sqlStr & " where orderserial = '" & iorderserial & "' and itemid <> 0 and cancelyn <> 'Y' "
    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
		if rsget("cnt") > 0 then
			goodname = rsget("itemname")
			if rsget("cnt") > 1 then
				goodname = goodname & " �� " & rsget("cnt")
			end if
		else
			goodname = "��ۺ�"
		end if
    end if
	rsget.close

	CheckNAssignCyberAcct = AssignCyberAcct(asid, iorderserial, subtotalPrice, goodname, buyname, accountname, buyhp, buyemail, userid, CyberAcctCode)
end function

function ChangeCyberAcct(LGD_OID, LGD_AMOUNT, LGD_CLOSEDATE)
    '/*
    ' * [������� �߱�/�����û ������]
    ' *
    ' * ������� �߱� ����(CHANGE)�� �ݾװ� �����ϸ� ���� �Ҽ� �ֽ��ϴ�.
    ' */
    dim CST_PLATFORM : CST_PLATFORM         = ""         ' LG�ڷ��� �������� ����(test:�׽�Ʈ, service:����)
    IF application("Svr_Info")="Dev" THEN CST_PLATFORM = "test"
''CST_PLATFORM = ""

    dim CST_MID : CST_MID = "tenbyten01"                 ' LG�ڷ������� ���� �߱޹����� �������̵� �Է��ϼ���.

    dim LGD_MID                                                  ' �׽�Ʈ ���̵�� 't'�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                                ' �������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if

    dim LGD_METHOD : LGD_METHOD          = "CHANGE"                              ' ASSIGN:�Ҵ�, CHANGE:����

    'LGD_PRODUCTINFO   	 = trim(request("LGD_PRODUCTINFO"))  	 ' ��ǰ����
    'LGD_BUYER          	 = trim(request("LGD_BUYER"))         	 ' �����ڸ�
	'LGD_ACCOUNTOWNER     = trim(request("LGD_ACCOUNTOWNER"))  	 ' �Ա��ڸ�
	'LGD_ACCOUNTPID       = trim(request("LGD_ACCOUNTPID"))       ' �Ա����ֹι�ȣ(�ɼ�)
	'LGD_BUYERPHONE       = trim(request("LGD_BUYERPHONE"))       ' �������޴�����ȣ
	'LGD_BUYEREMAIL       = trim(request("LGD_BUYEREMAIL"))       ' �������̸���(�ɼ�)
	'LGD_BANKCODE         = trim(request("LGD_BANKCODE"))         ' �Աݰ��������ڵ�
	'LGD_CASHRECEIPTUSE   = trim(request("LGD_CASHRECEIPTUSE"))   ' ���ݿ����� ���౸��('1':�ҵ����, '2':��������)
	'LGD_CASHCARDNUM      = trim(request("LGD_CASHCARDNUM"))      ' ���ݿ����� ī���ȣ
	'LGD_TAXFREEAMOUNT    = trim(request("LGD_TAXFREEAMOUNT"))    ' �鼼�ݾ�
	'LGD_CASNOTEURL       = "http://61.252.133.2:8888/admin/apps/DC_CA_noteurl.asp" ''"http://����URL/cas_noteurl.asp"       ' �Աݰ�� ó���� ���� ������������ �ݵ�� ������ �ּ���


    'configPath           = "C:/lgdacom"         				 ' LG�ڷ��޿��� ������ ȯ������("/conf/lgdacom.conf") ��ġ ����.
    dim configPath : configPath				   = "C:/lgdacom" '''"C:/lgdacom/conf/" & CST_MID  ''conf ���� ���� 2013/02/15

    dim xpay
    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "CyberAccount"
    xpay.Set "LGD_METHOD", LGD_METHOD
    xpay.Set "LGD_OID", LGD_OID
    xpay.Set "LGD_AMOUNT", LGD_AMOUNT
    xpay.Set "LGD_CLOSEDATE", LGD_CLOSEDATE
    'xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
    'xpay.Set "LGD_BUYER", LGD_BUYER
    'xpay.Set "LGD_ACCOUNTOWNER", LGD_ACCOUNTOWNER
    'xpay.Set "LGD_ACCOUNTPID", LGD_ACCOUNTPID
    'xpay.Set "LGD_BUYERPHONE", LGD_BUYERPHONE
    'xpay.Set "LGD_BUYEREMAIL", LGD_BUYEREMAIL
    'xpay.Set "LGD_BANKCODE", LGD_BANKCODE
    'xpay.Set "LGD_CASHRECEIPTUSE", LGD_CASHRECEIPTUSE
    'xpay.Set "LGD_CASHCARDNUM", LGD_CASHCARDNUM

    'xpay.Set "LGD_TAXFREEAMOUNT", LGD_TAXFREEAMOUNT
    'xpay.Set "LGD_CASNOTEURL", LGD_CASNOTEURL


    '/*
    ' * 1. ������� �߱�/���� ��û ���ó��
    ' *
    ' * ��� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
    ' */
    Dim itemCount, itemName, resCount, i, j
    Dim sqlStr

    ChangeCyberAcct = false

    if (xpay.TX()) then
        if LGD_METHOD = "ASSIGN" then      '������� �߱��� ���

'        	'1)������� �߱ް�� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
'        	Response.Write("������� �߱� ��ûó���� �Ϸ�Ǿ����ϴ�. <br>")
'        	Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'        	Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
'
'			Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
'	    	Response.Write("�ŷ���ȣ : " & xpay.Response("LGD_TID", 0) & "<p>")
'
'        	'�Ʒ��� ������û ��� �Ķ���͸� ��� ��� �ݴϴ�.
'
'        	itemCount = xpay.resNameCount
'        	resCount = xpay.resCount
'
'        	For i = 0 To itemCount - 1
'            	itemName = xpay.ResponseName(i)
'            	Response.Write(itemName & "&nbsp:&nbsp")
'            	For j = 0 To resCount - 1
'                	Response.Write(xpay.Response(itemName, j) & "<br>")
'            	Next
'        	Next

        else		'������� ������ ���
        	'1)������� ������ ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)


        	ChangeCyberAcct = (Trim(xpay.resCode)="0000")

        	if (Trim(xpay.resCode)="0000") then
        	    sqlStr = " IF EXISTS (select orderserial from db_order.dbo.tbl_order_CyberAccountLog where orderserial='" & LGD_OID & "')" & VbCrlf
                sqlStr = sqlStr & " BEGIN" & VbCrlf
                sqlStr = sqlStr & "	Insert Into db_order.dbo.tbl_order_CyberAccountLog" & VbCrlf
                sqlStr = sqlStr & "	(orderserial, differencekey, userid, FINANCECODE,ACCOUNTNUM" & VbCrlf
                sqlStr = sqlStr & "	, subtotalPrice, CLOSEDATE"& VbCrlf
                sqlStr = sqlStr & "	,RefIP)" & VbCrlf
                sqlStr = sqlStr & "	select top 1 orderserial, (differencekey+1) as differencekey" & VbCrlf
                sqlStr = sqlStr & "	,userid, FINANCECODE, ACCOUNTNUM" & VbCrlf
                sqlStr = sqlStr & "	, " & LGD_AMOUNT & " as subtotalprice" & VbCrlf
                sqlStr = sqlStr & "	, '" & Left(LGD_CLOSEDATE,4) + "-" + Mid(LGD_CLOSEDATE,5,2) + "-" + Mid(LGD_CLOSEDATE,7,2) + " " + Mid(LGD_CLOSEDATE,9,2) + ":" + Mid(LGD_CLOSEDATE,11,2) + ":" + Mid(LGD_CLOSEDATE,13,2) & "' as CLOSEDATE" & VbCrlf
                sqlStr = sqlStr & "	, '" & Left(request.ServerVariables("REMOTE_ADDR"),32) & "' as refip" & VbCrlf
                sqlStr = sqlStr & "	from db_order.dbo.tbl_order_CyberAccountLog" & VbCrlf
                sqlStr = sqlStr & "	where orderserial='" & LGD_OID & "'" & VbCrlf
                sqlStr = sqlStr & "	order by differencekey desc" & VbCrlf
                sqlStr = sqlStr & " END"

                dbget.Execute sqlStr
            ELSE
            	Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
                Response.Write("�ֹ���ȣ : " & LGD_OID & "<br>")
                Response.Write("�Աݾ� : " & LGD_AMOUNT & "<br>")
            	Response.Write("�Աݸ����� : " & LGD_CLOSEDATE & "<p>")


            	itemCount = xpay.resNameCount
            	resCount = xpay.resCount

            	For i = 0 To itemCount - 1
                	itemName = xpay.ResponseName(i)
                	Response.Write(itemName & "&nbsp:&nbsp")
                	For j = 0 To resCount - 1
                    	Response.Write(xpay.Response(itemName, j) & "<br>")
                	Next
            	Next
        	end if
        end if
    else
        '2)API ��û ���� ȭ��ó��
        ''Response.Write("������� �߱�/���� ��ûó���� ���еǾ����ϴ�. <br>")
        ''Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        ''Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if

end function

function AssignCyberAcct(asid, iorderserial, subtotalPrice, goodname, buyname, accountname, buyhp, buyemail, userid, CyberAcctCode)
    '/*
    ' * [������� �߱޿�û ������]
    ' */
	dim LGD_FINANCECODE, LGD_ACCOUNTNUM, Tid, accountno
	dim FINANCECODE, ACCOUNTNUM, CLOSEDATE, IsSuccess
	dim sqlStr, iresultmsg

    dim CST_PLATFORM : CST_PLATFORM         = ""         		' LG�ڷ��� �������� ����(test:�׽�Ʈ, service:����)
    IF application("Svr_Info")="Dev" THEN CST_PLATFORM = "test"

    dim CST_MID : CST_MID = "tenbyten01"                 		' LG�ڷ������� ���� �߱޹����� �������̵� �Է��ϼ���.

    dim LGD_MID                                                 ' �׽�Ʈ ���̵�� 't'�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                               ' �������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if

    dim LGD_METHOD       : LGD_METHOD        = "ASSIGN"             				' ASSIGN:�Ҵ�, CHANGE:����
    dim LGD_OID          : LGD_OID     		 = iorderserial    						' �ֹ���ȣ(�������� ����ũ�� �ֹ���ȣ�� �Է��ϼ���)
    dim LGD_AMOUNT       : LGD_AMOUNT      	 = subtotalprice      					' �ݾ�("," �� ������ �ݾ��� �Է��ϼ���)
    dim LGD_PRODUCTINFO  : LGD_PRODUCTINFO   = trim(goodname)  	 					' ��ǰ����
    dim LGD_BUYER        : LGD_BUYER         = trim(buyname)         				' �����ڸ�
	dim LGD_ACCOUNTOWNER : LGD_ACCOUNTOWNER  = trim(accountname)  					' �Ա��ڸ�
	dim LGD_ACCOUNTPID
	    LGD_ACCOUNTPID = Left(asid, 13)         									' �Ա����ֹι�ȣ(�ɼ�)/���̵� MAX 13 ,�ݾ�üũ

	dim LGD_BUYERPHONE   : LGD_BUYERPHONE       = trim(Replace(buyhp,"-",""))       ' �������޴�����ȣ
	dim LGD_BUYEREMAIL   : LGD_BUYEREMAIL       = trim(buyemail)       				' �������̸���(�ɼ�)
	dim LGD_BANKCODE     : LGD_BANKCODE         = trim(CyberAcctCode)         		' �Աݰ��������ڵ�

	dim LGD_CASHRECEIPTUSE, LGD_CASHCARDNUM
''�̴Ͻý� ���ݿ��������� ���
''	if (request.Form("cashreceiptreq")="Y") then
''	    LGD_CASHRECEIPTUSE   = trim(useopt+1)   ' ���ݿ����� ���౸��('1':�ҵ����, '2':��������)
''	    LGD_CASHCARDNUM      = trim(request.Form("cashReceipt_ssn")) ''trim(request("LGD_CASHCARDNUM"))      ' ���ݿ����� ī���ȣ
''	else
''	    LGD_CASHRECEIPTUSE  =""
''	    LGD_CASHCARDNUM     =""
''    end if

	dim LGD_CLOSEDATE
		LGD_CLOSEDATE       = trim(Replace(Left(dateadd("d",10,now()),10),"-","") + "235959")        ' �Ա� ������ 20100331 000000
	dim LGD_TAXFREEAMOUNT : LGD_TAXFREEAMOUNT   = "0 "    ' �鼼�ݾ�
	dim LGD_CASNOTEURL    : LGD_CASNOTEURL      = "http://scm.10x10.co.kr/admin/apps/DC_CA_noteurl.asp"       ' �Աݰ�� ó���� ���� ������������ �ݵ�� ������ �ּ���
IF application("Svr_Info")="Dev" THEN LGD_CASNOTEURL = "http://61.252.133.2:8888/admin/apps/DC_CA_noteurl.asp"

    dim configPath : configPath				   = "C:/lgdacom" '''/conf/" & CST_MID
    dim xpay

    On Error Resume Next
    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    IF (ERR) then
        response.write Err.Description
        response.write "<script language='javascript'>alert('������ �̷�� ���� �ʾҽ��ϴ�. \n\n: �˼��մϴ�. ������� �߱޿� ������ �ֽ��ϴ�. \n\n����� �ٽ� �õ��� �ֽñ� �ٶ��ϴ�.');</script>"
        response.end
    End IF
    On Error Goto 0

    xpay.Set "LGD_TXNAME", "CyberAccount"
    xpay.Set "LGD_METHOD", LGD_METHOD
    xpay.Set "LGD_OID", LGD_OID
    xpay.Set "LGD_AMOUNT", LGD_AMOUNT
    xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
    xpay.Set "LGD_BUYER", LGD_BUYER
    xpay.Set "LGD_ACCOUNTOWNER", LGD_ACCOUNTOWNER
    xpay.Set "LGD_ACCOUNTPID", LGD_ACCOUNTPID
    xpay.Set "LGD_BUYERPHONE", LGD_BUYERPHONE
    xpay.Set "LGD_BUYEREMAIL", LGD_BUYEREMAIL
    xpay.Set "LGD_BANKCODE", LGD_BANKCODE
    xpay.Set "LGD_CASHRECEIPTUSE", LGD_CASHRECEIPTUSE
    xpay.Set "LGD_CASHCARDNUM", LGD_CASHCARDNUM
    xpay.Set "LGD_CLOSEDATE", LGD_CLOSEDATE
    xpay.Set "LGD_TAXFREEAMOUNT", LGD_TAXFREEAMOUNT
    xpay.Set "LGD_CASNOTEURL", LGD_CASNOTEURL

    xpay.Set "LGD_CUSTOM_CASSMSMSG", "[�ٹ�����] [LGD_FINANCENAME] [LGD_SA] [LGD_COMPANYNAME] [LGD_AMOUNT]�� �ֹ���ȣ:"&iorderserial&" �����մϴ�"  ''2015/07/22

    if xpay.TX() then
        if LGD_METHOD = "ASSIGN" then      '������� �߱��� ���
            LGD_FINANCECODE = xpay.Response("LGD_FINANCECODE", 0)   ''����
            LGD_ACCOUNTNUM = xpay.Response("LGD_ACCOUNTNUM", 0)   ''�������
            Tid = xpay.Response("LGD_TID", 0)
        end if
    else
        response.write " [" + xpay.resCode + "] " & Replace(Left(xpay.resMsg,60),"'","")
		response.end
    end if

    IsSuccess = (xpay.resCode="0000")

    iresultmsg  = Left(xpay.resMsg,90)
    paygatetid = Tid

    if IsSuccess then
        FINANCECODE = LGD_FINANCECODE
        ACCOUNTNUM  = LGD_ACCOUNTNUM
        CLOSEDATE   = LGD_CLOSEDATE
        accountno = getLGD_FINANCECODE2Name(LGD_FINANCECODE) & " " & LGD_ACCOUNTNUM
        if (iresultmsg="") then
            iresultmsg =  "[�������] " & accountno
        end if
    else
        iresultmsg = "[" & xpay.resCode & "]" & iresultmsg
    end if

    if Not IsSuccess then
        ''������µ� ���а� ���� �� �ֵ��� ������.
		'// �ֹ����� �ֹ����� ó��
		sqlStr = " update [db_order].[dbo].tbl_order_master" + vbCrlf
		sqlStr = sqlStr + " set ipkumdiv='1' " + vbCrlf
		if (iresultmsg<>"") then
		    sqlStr = sqlStr + " ,resultmsg=convert(varchar(100),'" + iresultmsg + "')" + vbCrlf
		end if
		sqlStr = sqlStr + " where orderserial='" + CStr(iorderserial) + "'" + vbCrlf

		''response.write sqlStr & "<br>"
		dbget.Execute(sqlStr)
	else
		''' �ֹ� ����Ÿ ���Ӹ� ������
		sqlStr = " update [db_order].[dbo].tbl_order_master" + vbCrlf
		sqlStr = sqlStr + " set accountno='" + accountno + "' " + vbCrlf
		sqlStr = sqlStr + " ,ipkumdiv='2'" + vbCrlf

		if (paygatetid<>"") then
		    sqlStr = sqlStr + " ,paygatetid='" + paygatetid + "'" + vbCrlf
		end if

		if (iresultmsg<>"") then
		    sqlStr = sqlStr + " ,resultmsg=convert(varchar(100),'" + iresultmsg + "')" + vbCrlf
		end if

		sqlStr = sqlStr + " where orderserial='" + CStr(iorderserial) + "'" + vbCrlf

		''response.write sqlStr & "<br>"
		dbget.Execute(sqlStr)

        sqlStr = " insert into db_order.dbo.tbl_order_CyberAccountLog"
        sqlStr = sqlStr & " (orderserial, differencekey, userid, FINANCECODE, ACCOUNTNUM, subtotalPrice, CLOSEDATE, RefIP)"
        sqlStr = sqlStr & " values('" & iorderserial & "'"
        sqlStr = sqlStr & " ,0"
        sqlStr = sqlStr & " ,'" & userid & "'"
        sqlStr = sqlStr & " ,'" & FINANCECODE & "'"
        sqlStr = sqlStr & " ,'" & ACCOUNTNUM & "'"
        sqlStr = sqlStr & " ,'" & subtotalprice & "'"
        sqlStr = sqlStr & " ,'" & Left(CLOSEDATE,4) + "-" + Mid(CLOSEDATE,5,2) + "-" + Mid(CLOSEDATE,7,2) + " " + Mid(CLOSEDATE,9,2) + ":" + Mid(CLOSEDATE,11,2) + ":" + Mid(CLOSEDATE,13,2) & "'"
        sqlStr = sqlStr & " ,'" & Left(request.ServerVariables("REMOTE_ADDR"),32) & "'"
        sqlStr = sqlStr & " )"

        dbget.Execute sqlStr
    end if
    SET xpay = Nothing

	AssignCyberAcct = IsSuccess
end function

%>
