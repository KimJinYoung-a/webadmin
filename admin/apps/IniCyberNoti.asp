<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/rndSerial.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<% 

'*******************************************************************************
' FILE NAME : vacctinput.asp
' DATE : 2006.09
' �̴Ͻý� ������� �Աݳ��� ó��demon���� �Ѿ���� �Ķ���͸� control �ϴ� �κ� �Դϴ�.
'*******************************************************************************

'**********************************************************************************
'�̴Ͻý��� �����ϴ� ���������ü�� ����� �����Ͽ� DB ó�� �ϴ� �κ� �Դϴ�.	
'�ʿ��� �Ķ���Ϳ� ���� DB �۾��� �����Ͻʽÿ�.
' [��������] �ڼ��� ������ �޴��� ����
'**********************************************************************************	


Dim TEMP_IP : TEMP_IP = Request.ServerVariables("REMOTE_ADDR")
Dim PG_IP   : PG_IP	= Left(TEMP_IP,10)

''203.238.37.3, 203.238.37.15, 203.238.37.16, 203.238.37.25 
''39.115.212.9

IF (PG_IP <> "203.238.37") and (PG_IP <> "39.115.212") THEN  'PG���� ���´��� IP�� üũ 
    response.write "ERR"
    response.end
END IF
	
if (request("NO_OID")="G1100900031") or (request("NO_OID")="G1100900030")  then  ''10041273170 ''10051389839 ''iisue00
    response.write "OK"
    response.end
end if

	Dim NO_TID : NO_TID = Request("NO_TID")		            '�ŷ���ȣ
	Dim NO_OID : NO_OID = Request("NO_OID") 		        '���� �ֹ���ȣ
	Dim ID_MERCHANT : ID_MERCHANT = Request("ID_MERCHANT")	'���� ���̵�
	Dim CD_BANK : CD_BANK = Request("CD_BANK")		        '�ŷ� �߻� ��� �ڵ�
	Dim CD_DEAL : CD_DEAL = Request("CD_DEAL")		        '��� ��� �ڵ�	
	Dim DT_TRANS : DT_TRANS = Request("DT_TRANS")		    '�ŷ� ���� 
	Dim TM_TRANS : TM_TRANS = Request("TM_TRANS")		    '�ŷ� �ð�
	Dim NO_MSGSEQ : NO_MSGSEQ = Request("NO_MSGSEQ")	    '���� �Ϸ� ��ȣ
	Dim CD_JOINORG : CD_JOINORG = Request("CD_JOINORG")	    '���� ��� �ڵ�
	
	Dim DT_TRANSBASE : DT_TRANSBASE = Request("DT_TRANSBASE")	'�ŷ� ���� ����
	Dim NO_TRANSEQ : NO_TRANSEQ = Request("NO_TRANSEQ")	        '�ŷ� �Ϸ� ��ȣ
	Dim TYPE_MSG : TYPE_MSG = Request("TYPE_MSG")		        '�ŷ� ���� �ڵ� 
	Dim CL_CLOSE : CL_CLOSE = Request("CL_CLOSE")		        '���� �����ڵ�
	Dim CL_KOR : CL_KOR = Request("CL_KOR")		                '�ѱ� ���� �ڵ�
	Dim NO_MSGMANAGE : NO_MSGMANAGE = Request("NO_MSGMANAGE")	'���� ���� ��ȣ
	Dim NO_VACCT : NO_VACCT = Request("NO_VACCT")		        '������¹�ȣ
	Dim AMT_INPUT : AMT_INPUT = Request("AMT_INPUT")	        '�Աݱݾ�
	Dim AMT_CHECK : AMT_CHECK = Request("AMT_CHECK")	        '�̰��� Ÿ���� �ݾ�
	Dim NM_INPUTBANK : NM_INPUTBANK = Request("NM_INPUTBANK")	'�Ա� ���������
	Dim NM_INPUT : NM_INPUT = Request("NM_INPUT")		        '�Ա� �Ƿ���
	Dim DT_INPUTSTD : DT_INPUTSTD = Request("DT_INPUTSTD")	    '�Ա� ���� ����
	Dim DT_CALCULSTD : DT_CALCULSTD = Request("DT_CALCULSTD")	'���� ���� ����
	Dim FLG_CLOSE : FLG_CLOSE = Request("FLG_CLOSE")	        '���� ��ȭ 

' �������ä���� ���ݿ����� �ڵ��߱޽�û�ÿ��� ����
  Dim DT_CSHR : DT_CSHR      = Request("DT_CSHR")               '���ݿ����� �߱�����
  Dim TM_CSHR : TM_CSHR      = Request("TM_CSHR")               '���ݿ����� �߱޽ð�
  Dim NO_CSHR_APPL : NO_CSHR_APPL = Request("NO_CSHR_APPL")     '���ݿ����� �߱޹�ȣ
  Dim NO_CSHR_TID : NO_CSHR_TID  = Request("NO_CSHR_TID")       '���ݿ����� �߱�TID
	
if (NO_OID="G6011919280") then NO_OID="G6011923014" '' ������..
    
  Dim sqlStr, paramInfo, retParamInfo	

'// ǥ�������� �ֹ��� ��� ���ֹ���ȣ ����(2016.02.15; ������)
if left(NO_OID,1)<>"G" and len(NO_OID)>11 then
	sqlStr = "Select Top 1 giftOrderSerial from db_order.dbo.tbl_giftcard_order_temp Where no_OID='" & trim(NO_OID) & "'"
	rsget.Open sqlStr,dbget,1
		NO_OID = rsget("giftOrderSerial")
	rsget.Close
end if

'  rw NO_TID             'teenxteen820111009154648640402
'  rw NO_OID             'G1100900024
'  rw ID_MERCHANT        'teenxteen8
'  rw CD_BANK            '00000011
'  rw CD_DEAL            '00000011
'  rw DT_TRANS           '20111009
'  rw TM_TRANS           '154648
'  rw NO_MSGSEQ          '9000011158
'  rw CD_JOINORG         '01306001
'  rw DT_TRANSBASE       '20111009
'  rw NO_TRANSEQ         
'  rw TYPE_MSG           '0200
'  rw CL_CLOSE           '0
'  rw CL_KOR             '2
'  rw NO_MSGMANAGE       '15464858
'  rw NO_VACCT           '01444464225683
'  rw AMT_INPUT          '<br>0000000010000<br>  
'  rw AMT_CHECK          '0000000010000
'  rw NM_INPUTBANK       '__�׽�Ʈ__
'  rw NM_INPUT           'ȫ�浿
'  rw DT_INPUTSTD        '20111009
'  rw DT_CALCULSTD       '20111009       
'  rw FLG_CLOSE          '0
'  rw DT_CSHR            ' 
'  rw TM_CSHR            '
'  rw NO_CSHR_APPL       '
'  rw NO_CSHR_TID        '
    
'    NO_TID             ="teenxteen820111009154648640402"
'    NO_OID             ="G1100900024"                  
'    ID_MERCHANT        ="teenxteen8"                     
'    CD_BANK            ="00000011"                       
'    CD_DEAL            ="00000011"                       
'    DT_TRANS           ="20111009"                       
'    TM_TRANS           ="154648"                         
'    NO_MSGSEQ          ="9000011158"                     
'    CD_JOINORG         ="01306001"            
'    DT_TRANSBASE       ="20111009"                       
'    NO_TRANSEQ         =""                                
'    TYPE_MSG           ="0200"                           
'    CL_CLOSE           ="0"                            
'    CL_KOR             ="2"                              
'    NO_MSGMANAGE       ="15464858"                       
'    NO_VACCT           ="01444464225683"                 
'    AMT_INPUT          ="0000000010000"      
'    AMT_CHECK          ="0000000010000"                  
'    NM_INPUTBANK       ="__�׽�Ʈ__"                     
'    NM_INPUT           ="ȫ�浿"                         
'    DT_INPUTSTD        ="20111009"                       
'    DT_CALCULSTD       ="20111009"                       
'    FLG_CLOSE          ="0"                              
'    DT_CSHR            =""                              
'    TM_CSHR            =""                               
'    NO_CSHR_APPL       =""                               
'    NO_CSHR_TID        =""              
''    
   
	paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
	        ,Array("@NO_TID"  , adVarchar	, adParamInput, 40 , NO_TID)	_
	        ,Array("@NO_OID"  , adVarchar	, adParamInput, 40 , NO_OID)	_
	        ,Array("@ID_MERCHANT"  , adVarchar	, adParamInput, 10 , ID_MERCHANT)	_
	        ,Array("@CD_BANK"  , adVarchar	, adParamInput, 8 , CD_BANK)	_
	        ,Array("@CD_DEAL"  , adVarchar	, adParamInput, 8, CD_DEAL)	_
	        ,Array("@DT_TRANS"  , adVarchar	, adParamInput, 8 , DT_TRANS)	_
	        ,Array("@TM_TRANS"  , adVarchar	, adParamInput, 6 , TM_TRANS)	_
	        ,Array("@NO_MSGSEQ"  , adVarchar	, adParamInput, 20, NO_MSGSEQ)	_
	        ,Array("@CD_JOINORG"  , adVarchar	, adParamInput, 10 , CD_JOINORG)	_
	        ,Array("@DT_TRANSBASE"  , adVarchar	, adParamInput, 10 , DT_TRANSBASE)	_
	        ,Array("@NO_TRANSEQ"  , adVarchar	, adParamInput, 10, NO_TRANSEQ)	_
	        ,Array("@TYPE_MSG"  , adChar	, adParamInput, 4 , TYPE_MSG)	_
	        ,Array("@CL_CLOSE"  , adChar	, adParamInput, 1 , CL_CLOSE)	_
	        ,Array("@CL_KOR"  , adVarchar	, adParamInput, 10, CL_KOR)	_
	        ,Array("@NO_MSGMANAGE"  , adVarchar	, adParamInput,10, NO_MSGMANAGE)	_
	        ,Array("@NO_VACCT"  , adVarchar	, adParamInput, 20, NO_VACCT)	_
	        ,Array("@AMT_INPUT"  , adCurrency	, adParamInput,  , CLNG(AMT_INPUT))	_
	        ,Array("@AMT_CHECK"  , adCurrency	, adParamInput,  , AMT_CHECK)	_
	        ,Array("@NM_INPUTBANK"  , adVarchar	, adParamInput, 10 , TRIM(NM_INPUTBANK))	_
	        ,Array("@NM_INPUT"  , adVarchar	, adParamInput, 20 , NM_INPUT)	_ 
	        ,Array("@DT_INPUTSTD"  , adVarchar	, adParamInput, 10 , DT_INPUTSTD)	_
	        ,Array("@DT_CALCULSTD"  , adVarchar	, adParamInput, 10 , DT_CALCULSTD)	_
	        ,Array("@FLG_CLOSE"  , adChar	, adParamInput, 1 , FLG_CLOSE)	_
	        ,Array("@DT_CSHR"  , adVarchar	, adParamInput, 10 , DT_CSHR)	_
	        ,Array("@TM_CSHR"  , adVarchar	, adParamInput, 10 , TM_CSHR)	_
	        ,Array("@NO_CSHR_APPL"  , adVarchar	, adParamInput, 20 , NO_CSHR_APPL)	_
	        ,Array("@NO_CSHR_TID"  , adVarchar	, adParamInput, 40 , NO_CSHR_TID)	_  
	)
	
	sqlStr = "db_order.dbo.sp_Ten_CyberAcct_LogSaveINI"  '''//usp_Back_CyberAcct_LogSave
    retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
    
    dim LogIdx : LogIdx      = GetValue(retParamInfo, "@RETURN_VALUE")   ' ��������

    if (LogIdx<0) then
        call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[������� GiftCard]["&Left(now(),10)&"]�α� ������ ����", request("NO_OID") & "|" & request("TYPE_MSG"))
    
        response.write "ERR:1"
        response.end
    end if 
    
    Dim RetErr,retval, RetMsg, orderserial, clsOrder
    Dim resultMSG
    Dim iOrderHpNo, iBuyName, iOrderEmail, iReqEmail
    Dim retSMSok
    
    IF (TYPE_MSG="0200") THEN
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                    ,Array("@Orderserial"  , adVarchar	, adParamInput, 13 , NO_OID)	_
                    ,Array("@BackUserID", adVarchar	, adParamInput, 32, "system")	_
        			,Array("@NO_TID", adVarchar	, adParamInput, 40, NO_TID)	_
        			,Array("@AMT_INPUT", adCurrency	, adParamInput,, AMT_INPUT)	_
        			,Array("@AMT_CHECK", adCurrency	, adParamInput,, AMT_CHECK)	_
        			,Array("@CD_BANK", adVarchar	, adParamInput,8, CD_BANK)	_
        			,Array("@NO_VACCT", adVarchar	, adParamInput,20, NO_VACCT)	_
        			,Array("@NO_CSHR_APPL", adVarchar	, adParamInput,20, NO_CSHR_APPL)	_
        			,Array("@NM_INPUT", adVarchar	, adParamInput,20, NM_INPUT)	_
        			,Array("@LogIdx", adInteger	, adParamInput,, LogIdx)	_
        			,Array("@RetVal"	, adInteger  , adParamOutput,, 0) _
        			,Array("@RetMsg", adVarchar	, adParamOutput,100,"") _
        			,Array("@MatchOrderSerial", adVarchar	, adParamOutput,13,"") _
		)

        sqlStr = "db_order.dbo.sp_Ten_IpkumConfirm_Cyber_ProcINI" ''//usp_Back_IpkumConfirm_Cyber_Proc
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
        RetErr      = GetValue(retParamInfo, "@RETURN_VALUE")   ' ��������
        retval      = GetValue(retParamInfo, "@RetVal")         ' 
        RetMsg      = GetValue(retParamInfo, "@RetMsg")    
        orderserial  =  GetValue(retParamInfo, "@MatchOrderSerial") ' ��Ī�� �ֹ���ȣ
        
'rw "orderserial="&orderserial
'rw "RetErr="&RetErr
'rw "retval="&retval
'rw "RetMsg="&RetMsg
        IF (RetErr=0) then
            if (retval=1)  then
                if(orderserial<>"") then
                   
                    Set clsOrder = new cGiftCardOrder
                    clsOrder.FRectGiftOrderSerial = orderserial
                    clsOrder.getCSGiftcardOrderDetail
                    iOrderHpNo = clsOrder.FOneItem.Fbuyhp
                    iOrderEmail= clsOrder.FOneItem.Fbuyemail
                    iBuyName   = clsOrder.FOneItem.Fbuyname
                    iReqEmail  = clsOrder.FOneItem.Freqemail
                    Set clsOrder = Nothing
            
            
                    'On Error Resume Next
                    ''�����Ϸ� �̸��� 
                    call sendmailbankok_GIFTCard(iOrderEmail,iBuyName,orderserial)
                    
                    '''' ���� �Ϸ� SMS ����
                    Dim osms
                    set osms = new CSMSClass
                    CALL osms.SendAcctIpkumOkMsg(iOrderHpNo,orderserial)
                    
                    
                    ''' �����ڵ� ����.
                    ''retSMSok = osms.sendGiftCardLMSMsg(orderserial)
                    retSMSok = osms.sendGiftCardLMSMsg2016(orderserial)
                    set osms = Nothing
                    
                    
                    if (retSMSok) then
                        sqlStr = "update db_order.dbo.tbl_giftcard_order"
                        sqlStr = sqlStr & " set jumundiv=5"
                        sqlStr = sqlStr & " ,senddate=getdate()"
                        sqlStr = sqlStr & " ,ipkumdiv=8"
                        sqlStr = sqlStr & " where giftorderserial='"&orderserial&"'"
                        sqlStr = sqlStr & " and ipkumdiv>3"
                        sqlStr = sqlStr & " and jumundiv<5"
                        sqlStr = sqlStr & " and cancelyn='N'"
                        
                        dbget.Execute sqlStr
                    end if
                    
                    ''' �����ڵ� �̸��� ����
                    '// Giftī�� MMS �߼�::�����ο���
    	            Call sendGiftCardEmail_SMTP(orderserial)
                
                    
                end if
                
                resultMSG = "OK"
                if (TRUE) then
                    'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[������� GiftCard]["&Left(now(),10)&"]�Ա�Ȯ�� �Ϸ�" & orderserial, resultMSG)
                end if
            else
                
                resultMSG = "ERR : orderserial=" & orderserial & " : retval=" & retval & " : RetMsg=" & RetMsg
                call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[������� GiftCard]["&Left(now(),10)&"]�Ա�Ȯ�� ���� �߻�" & orderserial, resultMSG)
                
                if (retval=-1) then
                    ''�� ��ҵ� ������ǿ� ���� �Ա�ó�� ��û�� ���ð��.
                    
                end if
            end if
            
            
        ELSE
            resultMSG = "ERR"
            call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[������� GiftCard]["&Left(now(),10)&"]�Ա�Ȯ�� ERR", trim(request("NO_OID")) & "|"&trim(request("AMT_INPUT"))&"|"&resultMSG&"|"&RetErr)

        END IF

    ELSEIF (TYPE_MSG="0400") THEN
        '' �Ա����.
        Dim TmpOrderserial, TmpUserSeq, Tmpbuyhp
        sqlStr = " select giftorderserial, userid, buyhp from db_order.dbo.tbl_giftcard_order"
        sqlStr = sqlStr & " where giftorderserial='" + orderserial + "'" & VbCRLF
        sqlStr = sqlStr & " and ipkumdiv=4"
        sqlStr = sqlStr & " and cancelyn='N'"
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            TmpOrderserial = rsget("giftorderserial")
            TmpUserSeq     = rsget("userid")
            Tmpbuyhp       = rsget("buyhp")
        end if
        rsget.Close
        
        if (TmpOrderserial<>"") then
            sqlStr = " update db_order.dbo.tbl_giftcard_order"
            sqlStr = sqlStr & " set ipkumdiv=3"
            sqlStr = sqlStr & " , ipkumdate=NULL"
            sqlStr = sqlStr & " where giftorderserial='" + orderserial + "'" & VbCRLF
            sqlStr = sqlStr + " and ipkumdiv=4"
            sqlStr = sqlStr + " and cancelyn='N'"
            
            dbget.Execute sqlStr
            
            ''�α� �ٽ� �� ��Ī���� ����..
            ''sqlStr = " update  db_agirlOrder.dbo.tbl_IniCyberAcctNotiLog"
            ''sqlStr = sqlStr + " set isMatched='N'"
            ''sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
            ''dbget.Execute sqlStr
            
            
            ''�޸𳲱�.
'            sqlStr = " insert into db_Agirlcs.dbo.tbl_CsMemo"
'            sqlStr = sqlStr + " (orderserial, commCd, UserSeq, callCd"
'            sqlStr = sqlStr + " , qnaCd, phoneNo, writeUserSeq, finishUserSeq"
'            sqlStr = sqlStr + " , contents, isfinish,finishdate,regdate) "
'            sqlStr = sqlStr + " values('" + CStr(TmpOrderserial) + "','1','" + CStr(TmpUserSeq) + "','0','99','','0','0','�Ա���� - ������� " & trim(request("NO_VACCT")) & "," & trim(request("AMT_INPUT")) & " ','Y',getdate(),getdate()) "
'            dbget.Execute sqlStr
            
            resultMSG = "OK"
            ''On Error Resume Next
            ''SendAcctIpkumCancelMsg Tmpbuyhp,orderserial
            Dim iMsg : iMsg = "[10x10]�Ա� �� ����� ������ ��� �Ǿ����ϴ�. ����Ȯ���� �� �Ա� �� �ּ���"
            call NomalSendSMS("",Tmpbuyhp,iMsg)
            
            ''On Error Goto 0
                    
            'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[������� GiftCard]["&Left(now(),10)&"]�Ա���� �Ϸ�", trim(request("NO_OID")) & "|"&trim(request("AMT_INPUT"))&"|"&resultMSG&"|"&RetErr)

        else
            resultMSG = "ERR"
            
            'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[������� GiftCard]["&Left(now(),10)&"]�Ա�Ȯ�� ���� - ��Һ�", trim(request("NO_OID")) & "|"&trim(request("AMT_INPUT"))&"|"&resultMSG&"|"&RetErr)

        end if     
    ELSE
        '' unknown
        response.write "ERR"
    END IF

    response.write resultMSG
    
'''**********************************************************************************
'''   �̺κп� �α����� ��θ� �������ּ���.	
'' Dim objFSO,f
''	Set objFSO = CreateObject("Scripting.FileSystemObject")
''    Set f = objFSO.CreateTextFile("c:\inipay41\log\result.log",True)
''
'''**********************************************************************************	
''
''    f.WriteLine("************************************************")
''    f.WriteLine("ID_MERCHANT : " + ID_MERCHANT)
''    f.WriteLine("NO_TID : " + NO_TID)
''    f.WriteLine("NO_OID : " + NO_OID)
''    f.WriteLine("NO_VACCT : " + NO_VACCT)
''    f.WriteLine("AMT_INPUT : " + AMT_INPUT)
''    f.WriteLine("NM_INPUTBANK : " + NM_INPUTBANK)
''    f.WriteLine("NM_INPUT : " + NM_INPUT)
''    f.WriteLine("************************************************")
''    f.WriteLine("")
''
''    
''	f.WriteLine("��ü �����")
''	''f.WriteLine(msg_id)
''	f.WriteLine(NO_TID)
''	f.WriteLine(NO_OID)
''	f.WriteLine(ID_MERCHANT)
''	f.WriteLine(CD_BANK)
''	f.WriteLine(DT_TRANS)
''	f.WriteLine(TM_TRANS)
''	f.WriteLine(NO_MSGSEQ)
''	f.WriteLine(TYPE_MSG)
''	f.WriteLine(CL_CLOSE)
''	f.WriteLine(CL_KOR)
''	f.WriteLine(NO_MSGMANAGE)
''	f.WriteLine(NO_VACCT)
''	f.WriteLine(AMT_INPUT)
''	f.WriteLine(AMT_CHECK)
''	f.WriteLine(NM_INPUTBANK)
''	f.WriteLine(NM_INPUT)
''	f.WriteLine(DT_INPUTSTD)
''	f.WriteLine(DT_CALCULSTD)
''	f.WriteLine(FLG_CLOSE)
''	f.Close
''	

	
'************************************************************************************

	'������ ���� �����ͺ��̽��� ��� ���������� ���� �����ÿ��� "OK"�� �̴Ͻý���
	'�����ϼž��մϴ�. �Ʒ� ���ǿ� �����ͺ��̽� ������ �޴� FLAG ������ ��������
	'(����) OK�� �������� �����ø� �̴Ͻý� ���� ������ "OK"�� �����Ҷ����� ��� �������� �õ��մϴ�
	'��Ÿ �ٸ� ������ PRINT(response.write)�� ���� �����ñ� �ٶ��ϴ�
	
'	IF (�����ͺ��̽� ��� ���� ���� ���Ǻ��� = true) THEN

		'''response.write "OK" 			  ' ����� ������������ 
	
'*************************************************************************************	


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->