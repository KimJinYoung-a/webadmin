<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/db/dbiTMSHelper.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tax/sERP_EseroTaxCls.asp"-->

<%
DIM isTESTMODE : isTESTMODE = FALSE


Dim mode        : mode =  requestCheckvar(request("mode"),32)
Dim taxKey      : taxKey =  requestCheckvar(request("taxKey"),24)
Dim appDate     : appDate =  requestCheckvar(request("appDate"),10)
Dim taxSellType : taxSellType =  requestCheckvar(request("taxSellType"),10)
Dim hidcustcd   : hidcustcd =  requestCheckvar(request("hidcustcd"),16)

Dim sellCorpNo  : sellCorpNo =  requestCheckvar(request("sellCorpNo"),13)
Dim sellJongNo  : sellJongNo =  requestCheckvar(request("sellJongNo"),4)
Dim sellCorpName : sellCorpName =  requestCheckvar(request("sellCorpName"),60)
Dim sellCeoName  : sellCeoName =  requestCheckvar(request("sellCeoName"),32)

Dim buyCorpNo    : buyCorpNo =  requestCheckvar(request("buyCorpNo"),13)
Dim buyJongNo    : buyJongNo =  requestCheckvar(request("buyJongNo"),4)
Dim buyCorpName  : buyCorpName =  requestCheckvar(request("buyCorpName"),60)
Dim buyCeoName   : buyCeoName =  requestCheckvar(request("buyCeoName"),32)

Dim taxType      : taxType =  requestCheckvar(request("taxType"),10)
Dim suplySum    : suplySum =  requestCheckvar(request("suplySum"),10)
Dim taxSum      : taxSum =  requestCheckvar(request("taxSum"),10)
Dim totSum      : totSum =  requestCheckvar(request("totSum"),10)
Dim bigo        : bigo =  requestCheckvar(request("bigo"),100)
Dim DtlName     : DtlName =  requestCheckvar(request("DtlName"),100)
Dim DtlBigo     : DtlBigo =  requestCheckvar(request("DtlBigo"),100)
Dim evalTypeNm  : evalTypeNm =  requestCheckvar(request("evalTypeNm"),20)
Dim recreqGubunNm : recreqGubunNm =  requestCheckvar(request("recreqGubunNm"),10)

Dim sellEmail   : sellEmail =  requestCheckvar(request("sellEmail"),100)
Dim buyEmail   : buyEmail =  requestCheckvar(request("buyEmail"),100)
Dim stDt : stDt =  requestCheckvar(request("stDt"),10)
Dim edDt : edDt =  requestCheckvar(request("edDt"),10)

Dim matchSeq : matchSeq =  requestCheckvar(request("matchSeq"),10)
Dim bizSecCd : bizSecCd =  requestCheckvar(request("bizSecCd"),10)
Dim arap_cd  : arap_cd =  requestCheckvar(request("arap_cd"),13)
Dim chkPLANDATE : chkPLANDATE =  requestCheckvar(request("chkPLANDATE"),10)
Dim chkTaxKey : chkTaxKey = request.Form("chkTaxKey")
Dim ipFileNo  : ipFileNo = request.Form("ipFileNo")
Dim duppConfirm : duppConfirm = request.Form("duppConfirm")

Dim autoIcheIdx  : autoIcheIdx =  requestCheckvar(request.Form("autoIcheIdx"),10)
Dim autoIcheTitle: autoIcheTitle =  requestCheckvar(request.Form("autoIcheTitle"),50)
Dim corpNo       : corpNo =  requestCheckvar(request.Form("corpNo"),16)
Dim mayPrice     : mayPrice =  requestCheckvar(request.Form("mayPrice"),10)
Dim mayAcctDate  : mayAcctDate =  requestCheckvar(request.Form("mayAcctDate"),2)
Dim mayIcheDate  : mayIcheDate =  requestCheckvar(request.Form("mayIcheDate"),2)
Dim mayPumok     : mayPumok =  requestCheckvar(request.Form("mayPumok"),100)
Dim mayAcctJukyo : mayAcctJukyo =  requestCheckvar(request.Form("mayAcctJukyo"),30)
Dim matchType : matchType =  requestCheckvar(request.Form("matchType"),10)
Dim CUST_CD   : CUST_CD =  requestCheckvar(request.Form("CUST_CD"),13)
Dim taxkeyArr : taxkeyArr =  request.Form("taxkeyArr")

Dim ref : ref = request.serverVariables("HTTP_REFERER")

Dim sqlStr, pCNT, AssignedRow
Dim paramInfo, retParamInfo, RetErr, retErrStr,retErpLinkType, ret_SLTRKEY
Dim clsEsero, ArrVal
Dim PROD_CD,BIZSECTION_CD,SLDATE,RMK,PLAN_DATE, VAT_KIND, TotAMT, CURR_AMT, VAT_AMT, PUMMOK
Dim matchKey, payrealdate, orderOrChulgoSerial, iCorpNo, retVal
Dim arap_nm, erpDocLinkType,erpDocLinkKey, erpLinkType, erpLinkKey
Dim eseroKey, targetArr, preMapExists, targetCnt, targetGb
Dim SuccRow
Dim i

Dim chk2 : chk2=request.Form("chk2")
Dim opExpDailyCardidx
IF (mode="handTaxInput") then
    IF (taxKey="") then
        '''���� ��꼭 ���� �ϴ��� check
        IF (duppConfirm="") then
            sqlStr = "select taxKey from db_partner.dbo.tbl_esero_Tax"
            sqlStr = sqlStr & " where appDate='"&appDate&"'"
            sqlStr = sqlStr & " and sellCorpNo='"&sellCorpNo&"'"
            sqlStr = sqlStr & " and buyCorpNo='"&buyCorpNo&"'"
            sqlStr = sqlStr & " and totSum="&totSum

            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        	IF Not (rsget.EOF OR rsget.BOF) THEN
        		taxKey = rsget("taxKey")
        	END IF
        	rsget.close

        	if (taxKey<>"") then
        	    response.write "<script>parent.confirmedSubmit();</script>"
        	    dbget.Close()
        	    response.end
        	end if
        End IF

        pCNT = 1

''        sqlStr = "select Count(*)+1 as pCNT"
''        sqlStr = sqlStr & " from db_partner.dbo.tbl_esero_Tax"
''        ''sqlStr = sqlStr & " where appDate='"&appDate&"'"
''        sqlStr = sqlStr & " where Left(taxKey,8)='"&Replace(appDate,"-","")&"'"         '''����..
''        sqlStr = sqlStr & " and taxModiType=9"
''        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
''    	IF Not (rsget.EOF OR rsget.BOF) THEN
''    		pCNT = rsget("pCNT")
''    	END IF
''    	rsget.close
''
''    	'''�հ� �ȸ���.
''    	IF (appDate="2012-03-27") or (appDate="2012-01-31") THEN
''    	    pCNT=pCNT+10
''    	ENd If

    	sqlStr = "select IsNULL(Max(Right(taxkey,9)),'0') as pCNT from db_partner.dbo.tbl_esero_tax"
        ''sqlStr = sqlStr & " where appDate='"&appDate&"'"
        sqlStr = sqlStr & " where Left(taxkey,8)='"&Replace(appDate,"-","")&"'"
        sqlStr = sqlStr & " and taxModiType=9"
        sqlStr = sqlStr & " and Left(SubString(taxkey,9,16),7)='9999999'"
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        IF Not (rsget.EOF OR rsget.BOF) THEN
            pCNT = rsget("pCNT")
            pCNT = CLNG(pCNT)+1
        END IF
        rsget.close

    	taxKey = Replace(appDate,"-","")
    	taxKey = taxKey + "9999999" + Format00(9,pCNT)

        evalTypeNm = "�����Է�"

        sqlStr = "Insert into db_partner.dbo.tbl_esero_Tax"
        sqlStr = sqlStr & "(taxKey,appDate,sellCorpNo,sellJongNo,sellCorpName,sellCeoName,sellEmail"
        sqlStr = sqlStr & ",buyCorpNo,buyJongNo,BuyCorpName,BuyCeoName,buyEmail,totSum,suplySum,taxSum"
        sqlStr = sqlStr & ",taxSellType,taxModiType,taxType,evalTypeNm,Bigo,recreqGubunNm"
        sqlStr = sqlStr & ",DtlDate,DtlName,DtlSuplysum,DtltaxSum,DtlBigo,reqDate,sendDate,regdate,tax_cust_CD)"
        sqlStr = sqlStr & " values('"&taxKey&"'"
        sqlStr = sqlStr & " ,'"&appDate&"'"
        sqlStr = sqlStr & " ,'"&sellCorpNo&"'"
        sqlStr = sqlStr & " ,'"&sellJongNo&"'"
        sqlStr = sqlStr & " ,'"&html2DB(sellCorpName)&"'"
        sqlStr = sqlStr & " ,'"&html2DB(sellCeoName)&"'"
        sqlStr = sqlStr & " ,'"&html2DB(sellEmail)&"'"
        sqlStr = sqlStr & " ,'"&buyCorpNo&"'"
        sqlStr = sqlStr & " ,'"&buyJongNo&"'"
        sqlStr = sqlStr & " ,'"&html2DB(BuyCorpName)&"'"
        sqlStr = sqlStr & " ,'"&html2DB(BuyCeoName)&"'"
        sqlStr = sqlStr & " ,'"&html2DB(buyEmail)&"'"
        sqlStr = sqlStr & " ,"&totSum
        sqlStr = sqlStr & " ,"&suplySum
        sqlStr = sqlStr & " ,"&taxSum
        sqlStr = sqlStr & " ,'"&taxSellType&"'"
        sqlStr = sqlStr & " ,9"                                     '''�����꼭.
        sqlStr = sqlStr & " ,'"&taxType&"'"
        sqlStr = sqlStr & " ,'"&evalTypeNm&"'"
        sqlStr = sqlStr & " ,'"&html2DB(Bigo)&"'"
        sqlStr = sqlStr & " ,'"&html2DB(recreqGubunNm)&"'"
        sqlStr = sqlStr & " ,'"&appDate&"'"
        sqlStr = sqlStr & " ,'"&html2DB(DtlName)&"'"
        sqlStr = sqlStr & " ,"&suplySum
        sqlStr = sqlStr & " ,"&taxSum
        sqlStr = sqlStr & " ,'"&html2DB(DtlBigo)&"'"
        sqlStr = sqlStr & " ,NULL"
        sqlStr = sqlStr & " ,NULL"
        sqlStr = sqlStr & " ,getdate()"
        sqlStr = sqlStr & " ,'"&hidcustcd&"'"
        sqlStr = sqlStr & " )"

        dbget.Execute sqlStr,AssignedRow

        IF (duppConfirm="") then
            response.write "<script>alert('"&AssignedRow&" �� �ݿ��Ǿ����ϴ�.\n\n���� ��û �Ͻô°�� ���ݰ�꼭 �˻��� ���� ����Ͻñ� �ٶ��ϴ�.');parent.window.close();</script>"
            ''response.write "<script>parent.opener.location.reload()</script>"
            ''response.write "<script>parent.location.href='/admin/tax/popRegfileHand.asp?taxKey="+taxKey+"'</script>"
            dbget.close
            response.end
        End IF
    ELSE
        sqlStr = "Update  db_partner.dbo.tbl_esero_Tax"
        sqlStr = sqlStr & " set appDate='"&appDate&"'"
        sqlStr = sqlStr & " ,sellCorpNo='"&sellCorpNo&"'"
        sqlStr = sqlStr & " ,sellJongNo='"&sellJongNo&"'"
        sqlStr = sqlStr & " ,sellCorpName='"&html2DB(sellCorpName)&"'"
        sqlStr = sqlStr & " ,sellCeoName='"&html2DB(sellCeoName)&"'"
        sqlStr = sqlStr & " ,sellEmail='"&html2DB(sellEmail)&"'"
        sqlStr = sqlStr & " ,buyCorpNo='"&buyCorpNo&"'"
        sqlStr = sqlStr & " ,buyJongNo='"&buyJongNo&"'"
        sqlStr = sqlStr & " ,BuyCorpName='"&html2DB(BuyCorpName)&"'"
        sqlStr = sqlStr & " ,BuyCeoName='"&html2DB(BuyCeoName)&"'"
        sqlStr = sqlStr & " ,buyEmail='"&html2DB(buyEmail)&"'"
        sqlStr = sqlStr & " ,totSum="&totSum
        sqlStr = sqlStr & " ,suplySum="&suplySum
        sqlStr = sqlStr & " ,taxSum="&taxSum
        sqlStr = sqlStr & " ,taxSellType='"&taxSellType&"'"
        sqlStr = sqlStr & " ,taxType='"&taxType&"'"
        sqlStr = sqlStr & " ,evalTypeNm='"&evalTypeNm&"'"
        sqlStr = sqlStr & " ,Bigo='"&html2DB(Bigo)&"'"
        sqlStr = sqlStr & " ,recreqGubunNm='"&html2DB(recreqGubunNm)&"'"
        sqlStr = sqlStr & " ,DtlDate='"&appDate&"'"
        sqlStr = sqlStr & " ,DtlName='"&html2DB(DtlName)&"'"
        sqlStr = sqlStr & " ,DtlSuplysum="&suplySum
        sqlStr = sqlStr & " ,DtltaxSum="&taxSum
        sqlStr = sqlStr & " ,DtlBigo='"&html2DB(DtlBigo)&"'"
        sqlStr = sqlStr & " ,tax_cust_CD='"&hidcustcd&"'"
        sqlStr = sqlStr & " where taxKey='"&taxKey&"'"
 'rw sqlStr
        dbget.Execute sqlStr,AssignedRow
    ENd IF

    response.write "<script>alert('"&AssignedRow&" �� �ݿ��Ǿ����ϴ�.')</script>"
    response.write "<script>opener.location.reload()</script>"
    response.write "<script>location.href='/admin/tax/popRegfileHand.asp?taxKey="+taxKey+"'</script>"
    dbget.close
    response.end
ELSEIF (mode="delHandTax") then
    sqlStr = "delete E"
    sqlStr = sqlStr & "  from db_partner.dbo.tbl_esero_Tax E"
	sqlStr = sqlStr & "     Left Join db_partner.dbo.tbl_esero_TaxMatch M"
	sqlStr = sqlStr & "     on E.taxKey=M.taxKey"
	sqlStr = sqlStr & "     and M.matchSeq=0"
    sqlStr = sqlStr & " where E.taxModiType=9"          ''���� ��꼭��.
    sqlStr = sqlStr & " and E.taxKey='"&taxKey&"'"
    sqlStr = sqlStr & " and M.erpLinkType is NULL"      ''���� ������.

    dbget.Execute sqlStr,AssignedRow

    response.write "<script>alert('"&AssignedRow&" �� ���� �Ǿ����ϴ�.')</script>"
    if (AssignedRow=1) then
        response.write "<script>opener.location.reload();window.close();</script>"
    else
        response.write "<script>alert('���� ���°� ���� �Ϸ��ΰ�� ���� �Ұ�.')</script>"
        response.write "<script>location.href='/admin/tax/popRegfileHand.asp?taxKey="+taxKey+"'</script>"
    end if
    dbget.close
    response.end
ELSEIF (mode="autoMapp") then
'    IF (stDt="") and (edDt="") then
'        stDt = LEFT(CStr(dateADD("m",-1,now())),7)+"-01"
'        edDt = LEFT(CStr(dateADD("m",0,now())),7)+"-01"
'        sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_Match] '"&stDt&"','"&edDt&"'"
'        dbget.Execute sqlStr
'
'        stDt = LEFT(CStr(dateADD("m",0,now())),7)+"-01"
'        edDt = LEFT(CStr(dateADD("m",1,now())),7)+"-01"
'        sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_Match] '"&stDt&"','"&edDt&"'"
'
'        dbget.Execute sqlStr
'    ELSE
'        sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_Match] '"&stDt&"','"&edDt&"'"
'        dbget.Execute sqlStr
'    END IF

    sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_Match] '"&stDt&"','"&edDt&"'"
    dbget.Execute sqlStr

    ''���̶�ҿ�
    sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_Match_ITS] '"&stDt&"','"&edDt&"'"  '''�� ó��.
    dbget.Execute sqlStr
''rw sqlStr
''response.end
    response.write "<script>location.href='"&ref&"'</script>"
ELSEIF (mode="modiDtlName") then
    sqlStr = "update db_Partner.dbo.tbl_Esero_Tax"
    sqlStr = sqlStr & " set dtlnameorg=isNULL(dtlnameorg,dtlname)" & vbCRLF
    sqlStr = sqlStr & " ,dtlname='"&DtlName&"'"& vbCRLF
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF

    dbget.Execute sqlStr, AssignedRow
    response.write "<script>alert('���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="modiBizSec") then
    sqlStr = "update db_Partner.dbo.tbl_Esero_TaxMatch"
    sqlStr = sqlStr & " set bizSecCd='"&bizSecCd&"'" & vbCRLF
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF
    sqlStr = sqlStr & " and matchSeq="&matchSeq & vbCRLF

    dbget.Execute sqlStr, AssignedRow

    IF (AssignedRow<1) then
        sqlStr = "insert into db_Partner.dbo.tbl_Esero_TaxMatch"
        sqlStr = sqlStr & " (taxKey,matchSeq,matchType,matchKey,matchState,bizSecCD)"
        sqlStr = sqlStr & " values('"&taxKey&"'"
        sqlStr = sqlStr & " ,0"
        sqlStr = sqlStr & " ,0"                ''-- matchType 0 ����
        sqlStr = sqlStr & " ,0"
        sqlStr = sqlStr & " ,1"
        sqlStr = sqlStr & " ,'"&bizSecCd&"'"
        sqlStr = sqlStr & " )"
        dbget.Execute sqlStr, AssignedRow
    end if

    response.write "<script>alert('���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="modiArapCD") then
    sqlStr = "update db_Partner.dbo.tbl_Esero_Tax"
    sqlStr = sqlStr & " set tax_arap_cd="&arap_cd&"" & vbCRLF
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF

    dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="taxmapcomplex") then
    eseroKey = requestCheckVar(request("eseroKey"),30)

    sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOneComplex] '"&eseroKey&"',3"  ''' 3:: �ν�Ź�� CASE ��꼭1 : �¶���+��������
    dbget.Execute sqlStr

ELSEIF (mode="handTaxMapping") then
    eseroKey = requestCheckVar(request("eseroKey"),30)
    targetArr = request("targetArr")
    targetCnt = requestCheckVar(request("targetCnt"),10)
    targetGb  = requestCheckVar(request("targetGb"),10)

    IF Right(targetArr,1)="," then targetArr = Left(targetArr,Len(targetArr)-1)
    IF Right(taxkeyArr,1)="," then taxkeyArr = Left(taxkeyArr,Len(taxkeyArr)-1)

    rw "eseroKey="&eseroKey
    rw "taxkeyArr="&taxkeyArr
    rw "targetArr="&targetArr
    rw "targetCnt="&targetCnt
    rw "targetGb="&targetGb

    IF (targetCnt>"1")  then ''and (targetGb="9")

    ELSE
        rw "������"
        response.end
    END If

    IF (targetCnt="1") then
        IF (targetGb="1") then
            sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne] '"&eseroKey&"',1,"&targetArr&""
            dbget.Execute sqlStr
        ELSEIF (targetGb="2") then
            sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne] '"&eseroKey&"',2,"&targetArr&""
            dbget.Execute sqlStr
        ELSEIF (targetGb="9") then
            IF (eseroKey="") and (taxkeyArr<>"") then  ''�̼���N:������û1
                sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOneMulti_etcBuy] "&targetArr&",'"&taxkeyArr&"',1"
                dbget.Execute sqlStr
            ELSE                                                            ''�̼���1:������û1
                sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne_etcBuy] "&targetArr&",'"&eseroKey&"'"
                dbget.Execute sqlStr
            ENd IF
        ELSE
            rw "������"
        END IF
    ELSEIF (targetCnt>"1") then
        IF (targetGb="1") then
'            DECLARE @TaxKey varchar(24)
'            SET @TaxKey='2012043041000061a4cbc635'
'
            sqlStr = "insert into db_partner.dbo.tbl_esero_taxMatch" &vbCRLF
            sqlStr = sqlStr&" (TaxKey,matchseq,matchType,matchkey,matchstate,bizsecCD,erpLinkType,erpLinkKey)"&vbCRLF
            sqlStr = sqlStr&" select '"&eseroKey&"',"&vbCRLF
            sqlStr = sqlStr&" row_number() over (order by ub_totalsuplycash+me_totalsuplycash+wi_totalsuplycash+et_totalsuplycash+sh_totalsuplycash+dlv_totalsuplycash desc) -1"&vbCRLF
            sqlStr = sqlStr&" ,1,m.id ,1,'0000000101',NULL,NULL"&vbCRLF
            sqlStr = sqlStr&" from db_jungsan.dbo.tbl_designer_jungsan_master m"&vbCRLF
            sqlStr = sqlStr&" where eseroEvalseq='"&eseroKey&"'"&vbCRLF
            rw replace(sqlStr,VbCRLF,"<br>")

            dbget.Execute sqlStr
        ELSEIF (targetGb="2") then
'            DECLARE @TaxKey varchar(24)
'            SET @TaxKey='201204304100009677449481'

            sqlStr = "insert into db_partner.dbo.tbl_esero_taxMatch" &vbCRLF
            sqlStr = sqlStr&" (TaxKey,matchseq,matchType,matchkey,matchstate,bizsecCD,erpLinkType,erpLinkKey)"&vbCRLF
            sqlStr = sqlStr&" select '"&eseroKey&"',"&vbCRLF
            sqlStr = sqlStr&" row_number() over (order by tot_jungsanprice desc) -1"&vbCRLF
            sqlStr = sqlStr&" ,2,m.idx ,1,'0000000201',NULL,NULL"&vbCRLF
            sqlStr = sqlStr&" from db_jungsan.dbo.tbl_off_jungsan_master m"&vbCRLF
            sqlStr = sqlStr&" where eseroEvalseq='"&eseroKey&"'"&vbCRLF
            rw replace(sqlStr,VbCRLF,"<br>")

            dbget.Execute sqlStr
'            insert into db_partner.dbo.tbl_esero_taxMatch
'            (TaxKey,matchseq,matchType,matchkey,matchstate,bizsecCD,erpLinkType,erpLinkKey)
'            select @TaxKey,
'            row_number() over (order by tot_jungsanprice desc) -1
'            ,2,m.idx ,1,'0000000201',NULL,NULL
'            from db_jungsan.dbo.tbl_off_jungsan_master m
'            where eseroEvalseq=@TaxKey

        ELSEIF (targetGb="9") then
            IF (eseroKey<>"") and (targetArr<>"") then ''�̼���1:������ûN

                sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOneMulti_etcBuy] '"&targetArr&"','"&eseroKey&"',2"
                rw sqlStr
                dbget.Execute sqlStr
            ELSE
                rw "������"
            ENd IF
        ELSE
            rw "������"
        END If
'        targetArr = split(targetArr,",")
'        For i=LBound(targetArr) to UBound(targetArr)
'
'        Next
    ENd IF
    'rw sqlStr
    response.end

ELSEIF (mode="modiTaxMapping") then '''������꼭 ó�� / �ϰ� ����
    eseroKey = request("eseroKey")
    preMapExists =0

    eseroKey = Trim(eseroKey)
    if Right(eseroKey,1)="," then eseroKey=Left(eseroKey,Len(eseroKey)-1)
    eseroKey = Replace(eseroKey,",","','")
    eseroKey = "'"&eseroKey&"'"

    ''�̹� ���ε� ������ ������ �Ұ�.
    sqlStr = " select count(*) CNT from db_Partner.dbo.tbl_Esero_TaxMatch"
    sqlStr = sqlStr & " where taxKey in ("&eseroKey&")" & vbCRLF

    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    IF Not (rsget.EOF OR rsget.BOF) THEN
    	preMapExists = (rsget("CNT")>0)
    END IF
    rsget.close


    if (preMapExists) then
        response.write "<script>alert('�̹� ��Ī�� ������ ó�� �Ұ�.');location.href='"&ref&"'</script>"
        response.End
    end if

    if (matchType="") then matchType="0"

    sqlStr = "insert into db_Partner.dbo.tbl_Esero_TaxMatch"
    sqlStr = sqlStr & " (taxKey,matchSeq,matchType,matchKey,matchState"
    if (bizSecCd<>"") then
        sqlStr = sqlStr & ",bizSecCd"
    end if
    sqlStr = sqlStr & " )"
    sqlStr = sqlStr & " select T.taxKey,0,"&matchType&",0,1"
    if (bizSecCd<>"") then
        sqlStr = sqlStr & ",'"&bizSecCd&"'"
    end if
    sqlStr = sqlStr & " from  db_Partner.dbo.tbl_Esero_Tax T"
    sqlStr = sqlStr & " where t.taxKey in ("&eseroKey&")" & vbCRLF

    dbget.Execute sqlStr, AssignedRow

    if (arap_cd<>"") then
        sqlStr = " update T"
        sqlStr = sqlStr & " set tax_arap_cd="&arap_cd & vbCRLF
        sqlStr = sqlStr & " from db_Partner.dbo.tbl_Esero_Tax T" & vbCRLF
        sqlStr = sqlStr & " where T.taxKey in ("&eseroKey&")" & vbCRLF

        dbget.Execute sqlStr, AssignedRow
    end if

    if (cust_cd<>"") then
        sqlStr = " update T"
        sqlStr = sqlStr & " set tax_cust_cd='"&cust_cd &"'"& vbCRLF
        sqlStr = sqlStr & " from db_Partner.dbo.tbl_Esero_Tax T" & vbCRLF
        sqlStr = sqlStr & " where T.taxKey in ("&eseroKey&")" & vbCRLF

        dbget.Execute sqlStr, AssignedRow
    end if


    response.write "<script>alert('"&AssignedRow&" �� �ݿ� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="regCardUp") then
    chk2 = split(chk2,",")
    
    If IsArray(chk2) THEN
    For i=LBound(chk2) to UBound(chk2)
        opExpDailyCardidx = Trim(chk2(i))
        IF (opExpDailyCardidx<>"") THEN
            paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                ,Array("@opExpDailyCardidx"	,adInteger, adParamInput,, opExpDailyCardidx) _
                ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
        	)

            sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_CARD_Acc_cd_update"

        	IF application("Svr_Info")="Dev" THEN
        	    sqlStr = sqlStr&"_TEST"
            END IF

            retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

            RetErr       = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ� or IDX
            retErrStr    = GetValue(retParamInfo, "@retErrStr")  
            
            if (isTESTMODE) or (RetErr<0) then
                rw "key:"&opExpDailyCardidx&":ERR:["&RetErr&"]"&retErrStr&":ret_SLTRKEY:"&ret_SLTRKEY
            else
                sqlStr = "update db_Partner.dbo.tbl_OpExpDailyCard"
                sqlStr = sqlStr & " set erpLinkSeq='"&RetErr&"'"
                sqlStr = sqlStr & " where opExpDailyCardidx='"&opExpDailyCardidx&"'" & vbCRLF

                dbget.Execute sqlStr, AssignedRow
                
                rw "key:"&opExpDailyCardidx&":OK:["&RetErr&"]"
            end if
        ENd IF
    Next   
    ENd IF 
    
ELSEIF (mode="regCardMeaip") then
    rw "chk2="&chk2    
    chk2 = split(chk2,",")
    
    'rw "������"
    'response.end

    If IsArray(chk2) THEN
    For i=LBound(chk2) to UBound(chk2)
        opExpDailyCardidx = Trim(chk2(i))
        IF (opExpDailyCardidx<>"") THEN
            paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                ,Array("@opExpDailyCardidx"	,adInteger, adParamInput,, opExpDailyCardidx) _
                ,Array("@RET_SLTRKEY"	,adVarchar, adParamOutput,12, "") _
                ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
        	)

            sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_CARD_SALE_INPUT_sERP"

        	IF application("Svr_Info")="Dev" THEN
        	    sqlStr = sqlStr&"_TEST"
            END IF

            retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

            RetErr       = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ� or IDX
            ret_SLTRKEY = GetValue(retParamInfo, "@RET_SLTRKEY")   'ret_SLTRKEY
            retErrStr    = GetValue(retParamInfo, "@retErrStr")  
            
            if (isTESTMODE) or (RetErr<0) then
                rw "key:"&opExpDailyCardidx&":ERR:["&RetErr&"]"&retErrStr&":ret_SLTRKEY:"&ret_SLTRKEY
            else
                sqlStr = "update db_Partner.dbo.tbl_OpExpDailyCard"
                sqlStr = sqlStr & " set erpLinkSeq='"&ret_SLTRKEY&"'"
                sqlStr = sqlStr & " where opExpDailyCardidx='"&opExpDailyCardidx&"'" & vbCRLF

                dbget.Execute sqlStr, AssignedRow
                
                rw "key:"&opExpDailyCardidx&":OK:["&RetErr&"]"
            end if
        ENd IF
    Next   
    ENd IF 
ELSEIF (mode="sendDocErp") then
    Dim IsICheByArrayProc
    Dim IsEtaxArray : IsEtaxArray = false  ''2016/05/13
    
    rw "taxKey="&taxKey
    rw "chkTaxKey="&chkTaxKey
    rw "taxKeyArr="&taxKeyArr

    if (chkTaxKey<>"") and (taxKey="") then
        IsICheByArrayProc = true                            '''' ��ü���Ͽ��� ����.
    else
        IsICheByArrayProc = false
        chkTaxKey = taxKey
    end if
    
    if (IsICheByArrayProc) then 
        rw "���Ұ�(��ü���Ͽ��� ����)"
        response.end    
    end if

'rw chkTaxKey
'rw taxKey
'rw "��� ������"
'response.end
   
    if (chkTaxKey="") and (taxKeyArr<>"") then 
        chkTaxKey = taxKeyArr
        IsEtaxArray = true
    end if
    
    chkTaxKey = split(chkTaxKey,",")

    ''���� ���� Ÿ�Խ� ������ Tax�� �ִ��� üũ �ؾ���.. 2012/03/09
    ''201202209999999000000001 ==> ����/���� ��꼭 ������ ������ TaxKey�� �־����. (������û��)
    ''popHandMapping.asp ���� �����ϵ���..

    If IsArray(chkTaxKey) THEN
    For i=LBound(chkTaxKey) to UBound(chkTaxKey)
        taxKey = Trim(chkTaxKey(i))
        IF (taxKey<>"") THEN
            set clsEsero = new CEsero
            clsEsero.FtaxKey = taxKey
            ArrVal = clsEsero.fnGetEseroOneTax
            set clsEsero = Nothing

            IF IsArray(ArrVal) then
                taxSellType = ArrVal(15,0)                              '' 0���� 1 ����
                cust_cd     = ArrVal(36,0)
            	arap_cd     = ArrVal(38,0)
            	prod_cd     = ArrVal(40,0)
            	BIZSECTION_CD    = ArrVal(32,0)
            	SLDATE = replace(replace(ArrVal(1,0),"-",""),"/","")
            	matchType   = ArrVal(29,0)
            	matchKey    = ArrVal(30,0)
            	payrealdate = ArrVal(42,0)
            	IF IsNULL(payrealdate) then payrealdate=""
            	payrealdate = replace(replace(payrealdate,"-",""),"/","")

            	'''****************************************************** �߿�
            	                                                          ''���� ���޿����� :: ���� ���ڰ� �־�� ���� ��꼭 ���� ����.
                If (taxSellType="0") then
                    iCorpNo = ArrVal(2,0)  ''buyCorpNo
                    IF (matchType=9) or (IsICheByArrayProc) then           ''��Ÿ �����ΰ�� ==> �����Ϸ� Setting
                        PLAN_DATE = payrealdate                           ''���� ���ޱ� ó���� ���̽��� ���(��꼭 ���� �Է�) ������ �Է�.
                        IF (IsNULL(PLAN_DATE)) or (PLAN_DATE="") then PLAN_DATE=SLDATE
                    else
                        PLAN_DATE = SLDATE
                    end if
                else
                    iCorpNo = ArrVal(7,0)  ''sellCorpNo
                    PLAN_DATE   = SLDATE                                  ''������ ��� ��꼭 �����Ϸ�.
                End IF

            	IF (chkPLANDATE="") and (Not IsICheByArrayProc) then
            	    PLAN_DATE = ""
                END IF
            	''IF (taxSellType="0") then PLAN_DATE=""                  ''�̹� ������ ���̽��� ���� �ʿ� ����../���� �ʿ� ���°��
            	'''******************************************************

            	''���� �׸� ���ޱ����� ���� �Ұ� üũ
            	RMK         = LEFT(ArrVal(19,0),60) ''LEFT/60�߰� 2015/04/06 (BIGO)
            	
            	VAT_KIND    = ArrVal(17,0)
            	IF (VAT_KIND=2) then ''0����,2�鼼,3����.
            	    VAT_KIND = "3"
            	ELSEIF (VAT_KIND=3) then
            	    VAT_KIND = "2"
            	ELSE
            	    VAT_KIND = "0"
                END IF

            	TotAMT      = ArrVal(12,0)
            	CURR_AMT    = ArrVal(13,0)
            	VAT_AMT      = ArrVal(14,0)

            	PUMMOK     = ArrVal(22,0)               ''(DtlName)
            	RMK = LEFT(PUMMOK,60) ''2016/06/21 ȭ�� ��û
            	
            	DTLBIGO     = ArrVal(25,0)
            	orderOrChulgoSerial = ArrVal(46,0)

            	erpLinkType = ArrVal(33,0)
	            erpLinkKey  = ArrVal(34,0)

            	IF (TRIM(orderOrChulgoSerial)<>"") and (RMK="") then
            	    RMK = "�ֹ�/����ڵ�"&orderOrChulgoSerial
            	END IF
            end if
            ''VAT_KIND--0����,2�鼼,3����.
            '' PLAN_DATE ���� ��� ������ ���� ���� ����..

            ''����. 1. �����׸�� ���� ���� ������ �´���.
            ''      2. �ŷ�ó �ڵ尡 ������� �ű��Է�.
            Dim arap_gb

            IF ((isNULL(arap_cd)) or (arap_cd="")) Then
        	    response.write "<script>alert('�����׸� �� ���ε��� �ʾҽ��ϴ�.');location.href='"&ref&"'</script>"
        	    dbget.Close()
        	    response.end
            ENd IF
            
            sqlStr = "select arap_gb, arap_nm from db_partner.dbo.tbl_TMS_BA_ARAP_CD"& CHKIIF(isTESTMODE,"_sERP","")
            sqlStr = sqlStr + " where arap_cd="&arap_cd
            sqlStr = sqlStr + " and del_yn='N'"
            sqlStr = sqlStr + " and use_yn='Y'"
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        	IF Not (rsget.EOF OR rsget.BOF) THEN
        		arap_gb = rsget("arap_gb")                      ''1���� 2����
        		arap_nm = rsget("arap_nm")
        	END IF
        	rsget.close

        	IF ((taxSellType=0) and (arap_gb<>"2")) or ((taxSellType=1) and (arap_gb<>"1")) Then
        	    response.write "<script>alert('�����׸�� ���� ���� ������ ��ġ ���� �ʽ��ϴ�.');location.href='"&ref&"'</script>"
        	    dbget.Close()
        	    response.end
            ENd IF

            IF (arap_cd="390") or (arap_cd="20") or (arap_cd="174") THEN     '''���ޱ�(390) �����(20) �����ݴ�ü(174) ���� ��ϵȰ��� �߸��Ȱ��ε�.. ?? ���� 2016/05/11 ���� , �ϴ� �����
                response.write "<script>alert('���� �׸� Ȯ�� ��� - ["&arap_cd&"] "&arap_nm&"');location.href='"&ref&"'</script>"
        	    dbget.Close()
        	    response.end
            ENd IF

            ''2016/05/13 �߰� �̺�Ʈ����ǰ, �����
            if (IsEtaxArray) and ((arap_cd="855") or (arap_cd="940") or (arap_cd="813") or (arap_cd="912")) then
                if (taxKey<>"201611304100009612560116") then ''imsi
                    response.write "<script>alert('�Ұ� ���� �����׸��� �������۸� ���� - ["&arap_cd&"] "&arap_nm&"');location.href='"&ref&"'</script>"
            	    dbget.Close()
            	    response.end
        	    end if
            end if
            
            IF (Not isTESTMODE) and (Not (IsNULL(erpLinkType) or (erpLinkType=""))) then
                 response.write "<script>alert('�� ���� ���� ["&erpLinkType&"] "&erpLinkKey&"');location.href='"&ref&"'</script>"
        	     dbget.Close()
        	     response.end
            ENd IF
            'sqlStr = "select taxKey,matchSeq "
            'sqlStr = sqlStr + " from db_partner.dbo.tbl_esero_taxMatch order by matchSeq"
            'sqlStr = sqlStr + " where taxKey='"&taxKey&"'"

            Dim PayrequestState:PayrequestState=0
            '''������ �������� üũ
            rw matchType
            IF (matchType="9") THEN ''�Ϲ� ���԰�
                sqlStr = " select erpDocLinkType,erpDocLinkKey,IsNULL(p.PayrequestState,0) as PayrequestState "
                sqlStr = sqlStr & " from db_partner.dbo.tbl_eappPayDoc D"
            	sqlStr = sqlStr & " Join db_partner.dbo.tbl_eappPayrequest P"
            	sqlStr = sqlStr & " on D.payrequestIdx=P.payrequestIdx"
                sqlStr = sqlStr & " where D.payrequestIdx="&matchKey
                sqlStr = sqlStr & " and D.erpDocLinkType is NULL"
                sqlStr = sqlStr & " and P.isusing=1 "
         ''rw  sqlStr
                rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            	IF Not (rsget.EOF OR rsget.BOF) THEN
            		erpDocLinkType = rsget("erpDocLinkType")
            		erpDocLinkKey = rsget("erpDocLinkKey")
            		PayrequestState = rsget("PayrequestState")
            	END IF
            	rsget.close


                IF (erpDocLinkType<>"") THEN
                    response.write "<script>alert('�� ���� ���� ["&erpDocLinkType&"] "&erpDocLinkKey&"');location.href='"&ref&"'</script>"
        	        dbget.Close()
        	        response.end
                END IF

                ''2012/05/18 �ּ� ó�� / ������û�� ��������.
'                IF (PayrequestState<8) then
'                    response.write "<script>alert('���� ��û�� ERP ���� �� ��밡��.');location.href='"&ref&"'</script>"
'        	        dbget.Close()
'        	        response.end
'                END IF

            END IF

            if (cust_cd="") or (IsNULL(cust_cd)) then
                if (Not IsICheByArrayProc) then
                    '' ����� ��ȣ�� �ŷ�ó ������. �ߺ��ϰ�� return / ���°�� ���� ��� �Ľ���.
                    retVal = fnGetOrMakeCUST_sERP(iCorpNo,taxKey,cust_cd)

                    if (retVal=-1) then
                        response.write "<script>alert('����ڹ�ȣ : "&iCorpNo&" �ߺ��� �ŷ�ó�� ���� �մϴ�.-1');location.href='"&ref&"'</script>"
            	        dbget.Close()
            	        response.end
                    elseif (retVal=-9) then
                        response.write "<script>alert('�ŷ�ó ��� �� �����. ����ڹ�ȣ :"&iCorpNo&"');location.href='"&ref&"'</script>"
            	        dbget.Close()
            	        response.end
                    end if
                end if
            end if
  '' rw "üũ��.."
  '' response.end         
           ' rw "taxKey="&taxKey
           ' rw "arap_cd="&arap_cd
           ' rw "prod_cd="&prod_cd
           ' rw "cust_cd="&cust_cd
           ' rw "BIZSECTION_CD="&BIZSECTION_CD
           ' rw "PLAN_DATE="&PLAN_DATE
           ' rw "RMK="&RMK
           ' rw "VAT_KIND="&VAT_KIND
           ' rw "TotAMT="&TotAMT
           ' rw "CURR_AMT="&CURR_AMT
           ' rw "VAT_AMT="&VAT_AMT
           ' rw "PUMMOK="&PUMMOK
           ' rw "DTLBIGO="&DTLBIGO
            
           ' rw "chkPLANDATE="&chkPLANDATE
'response.end

'            IF (matchType="1") or (matchType="2") then
'                rw "�¿��� ���԰ŷ� ��� ���� �Ұ�."
'                dbget.Close()
'            	response.end
'            end if

            paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                ,Array("@taxKey"	,adVarchar, adParamInput,24, taxKey) _
                ,Array("@ARAP_CD"	,adInteger, adParamInput,, ARAP_CD) _
                ,Array("@PROD_CD"	,adVarchar, adParamInput,10, PROD_CD) _
                ,Array("@CUST_CD"	,adVarchar, adParamInput,13, CUST_CD) _
                ,Array("@BIZSECTION_CD"	,adVarchar, adParamInput,10, BIZSECTION_CD) _
                ,Array("@SLDATE"	,adVarchar, adParamInput,8, SLDATE) _
                ,Array("@PLAN_DATE"	,adVarchar, adParamInput,8, PLAN_DATE) _
                ,Array("@RMK"	,adVarchar, adParamInput,200, RMK) _
                ,Array("@VAT_KIND"	,adVarchar, adParamInput,10, VAT_KIND) _
                ,Array("@TotAMT"	,adCurrency, adParamInput,, TotAMT) _
                ,Array("@CURR_AMT"	,adCurrency, adParamInput,, CURR_AMT) _
                ,Array("@VAT_AMT"	,adCurrency, adParamInput,, VAT_AMT) _
                ,Array("@PUMMOK"	,adVarchar, adParamInput,100, PUMMOK) _
                ,Array("@DTLBIGO"	,adVarchar, adParamInput,200, DTLBIGO) _


                ,Array("@RET_SLTRKEY"	,adVarchar, adParamOutput,12, "") _
                ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
        	)

            sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_DocInputByTaxKey_sERP"

        	IF application("Svr_Info")="Dev" THEN
        	    sqlStr = sqlStr&"_TEST"
            END IF

            retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

            RetErr       = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ� or IDX
            retErrStr    = GetValue(retParamInfo, "@retErrStr")   ' ������ �����ȣ
            ''retErpLinkType = GetValue(retParamInfo, "@erpLinkType")   'S:����,F:�ڱݼ���,C:ȸ��
            ret_SLTRKEY = GetValue(retParamInfo, "@RET_SLTRKEY")   'ret_SLTRKEY
             

            if (isTESTMODE) or (RetErr<0) then ''then  '' or (RetErr<0)  ���� TRUE
                rw "RetErr="&RetErr
                rw "retErrStr="&retErrStr
                rw "ret_SLTRKEY="&ret_SLTRKEY
            ELSE
                retErpLinkType = "S"
                
                sqlStr = "update db_Partner.dbo.tbl_Esero_TaxMatch"
                sqlStr = sqlStr & " set erpLinkType='"&retErpLinkType&"'"
                sqlStr = sqlStr & " ,erpLinkKey='"&ret_SLTRKEY&"'"
                sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF
''rw sqlStr
                dbget.Execute sqlStr, AssignedRow
                '''����ó update
                IF (matchType="9") then
                    sqlStr = "update db_partner.dbo.tbl_eAppPayDoc"
                    sqlStr = sqlStr & " set erpDocSendDate=(CASE WHEN erpDocSendDate is NULL THEN getdate() ELSE erpDocSendDate END)"
                    sqlStr = sqlStr & " ,erpDocLinkType=(CASE WHEN erpDocLinkType is NULL THEN '"&retErpLinkType&"' ELSE erpDocLinkType END)"
                    sqlStr = sqlStr & " ,erpDocLinkKey=(CASE WHEN erpDocLinkKey is NULL THEN '"&ret_SLTRKEY&"' ELSE erpDocLinkKey END)"
                    sqlStr = sqlStr & " where payrequestIdx="&matchKey
        ''rw sqlStr
                    dbget.Execute sqlStr, AssignedRow
                ELSEIF (matchType="1") or (matchType="2") then
                    IF (ipFileNo<>"") THEN
                        sqlStr = "update D"
                        sqlStr = sqlStr & " set erpLinkType=M.erpLinkType"
                        sqlStr = sqlStr & " ,erpLinkKey=M.erpLinkKey"
                        sqlStr = sqlStr & " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D"
                        sqlStr = sqlStr & " left Join db_partner.dbo.tbl_Esero_TaxMatch M"
                        sqlStr = sqlStr & " 	on (CASE WHEN D.targetGbn='ON' then 1 "
                        sqlStr = sqlStr & " 	 WHEN D.targetGbn='OF' then 2"
                        sqlStr = sqlStr & " 	ELSE -1 END)=M.matchType"
                        sqlStr = sqlStr & "     and D.targetIdx=M.matchKey"
                        sqlStr = sqlStr & " Left Join db_partner.dbo.tbl_Esero_Tax T"
                        sqlStr = sqlStr & " on M.taxKey=T.TaxKey"
                        sqlStr = sqlStr & " where D.ipFileNo="&ipFileNo
                        sqlStr = sqlStr & " and T.taxKey='"&taxKey&"'"
                        sqlStr = sqlStr & " and M.erpLinkType is Not NULL"
                        sqlStr = sqlStr & " and D.erpLinkType is NULL"

                        dbget.Execute sqlStr, AssignedRow
                    END IF
                END IF

                SuccRow = SuccRow + AssignedRow
                rw "RetErr="&RetErr
                rw "retErpLinkType="&retErpLinkType
            ENd IF
        ENd IF
    Next

    '' �ʿ����.
'    if (NOT isTESTMODE) and (IsICheByArrayProc) then '' ���� FALSE
'        '' �ٸ������� ���۽� ������ ������ �����Ƿ� Flag update
'        IF (ipFileNo<>"") THEN
'            sqlStr = " exec [db_jungsan].[dbo].[sp_Ten_ipFileErpFlagUpdate] "&ipFileNo&""
'            rw sqlStr
'            dbget.Execute sqlStr
'        ENd IF
'    end if

    if (SuccRow<1) then 
        response.write "<script>alert('"&SuccRow&"�� �Է� �Ǿ����ϴ�.');/*location.href='"&ref&"'*/</script>"
    else
        response.write "<script>alert('"&SuccRow&"�� �Է� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
    end if
    End IF
ELSEIF (mode="ErpInOutMapping") then
    sqlStr = " exec dbo.[sp_TEN_ICheMapping] '"&Replace(request("ichedate"),"-","")&"','"&request("BIZSECTION_CD")&"'"
    rw sqlStr
    dbiTms_dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('"&AssignedRow&" �� �ݿ� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="finishDocHand") then
    sqlStr = "update db_Partner.dbo.tbl_Esero_TaxMatch"
    sqlStr = sqlStr & " set erpLinkType='H'"
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF

    dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="delErpLinkKey") then
    sqlStr = "update db_Partner.dbo.tbl_Esero_TaxMatch"
    sqlStr = sqlStr & " set erpLinkType=NULL"
    sqlStr = sqlStr & " ,erpLinkKey=NULL"
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF

    dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="delMapDTL") then
    ''������ ��cHECK �� �����ڷᰡ �ִ���..

    sqlStr = "delete from db_Partner.dbo.tbl_Esero_TaxMatch"
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF
    sqlStr = sqlStr & " and matchSeq="&matchSeq

    dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('"&AssignedRow&" �� ���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="chgHandMap") then
    sqlStr = "update  db_Partner.dbo.tbl_Esero_TaxMatch"
    sqlStr = sqlStr & " set matchType=0"
    sqlStr = sqlStr & " ,matchKey=0"
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF
    sqlStr = sqlStr & " and matchSeq="&matchSeq

    dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('"&AssignedRow&" �� ���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="delPeriod") then
    IF (autoIcheIdx<>"") then
        sqlStr = "delete from db_partner.dbo.tbl_Esero_AutoIcheMapData "
        sqlStr = sqlStr&" where autoIcheIdx="&autoIcheIdx
        dbget.Execute sqlStr, AssignedRow

        response.write "<script>alert('"&AssignedRow&" �� ���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
    ENd IF
ELSEIF (mode="mayErrStat") then
    sqlStr = "update  db_Partner.dbo.tbl_Esero_Tax"
    sqlStr = sqlStr & " set mayErrType=1"
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF
    sqlStr = sqlStr & " and mayErrType is NULL"

    dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('"&AssignedRow&" �� ���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
ELSEIF (mode="mayErrStatDel") then
    sqlStr = "update  db_Partner.dbo.tbl_Esero_Tax"
    sqlStr = sqlStr & " set mayErrType=NULL"
    sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF
    sqlStr = sqlStr & " and mayErrType=1"

    dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('"&AssignedRow&" �� ���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"


ELSEIF (mode="regPeriod") then

    mayPrice=replace(mayPrice,",","")
    mayAcctDate=Format00(2,mayAcctDate)
    mayIcheDate=Format00(2,mayIcheDate)

    if (mayPrice="") then mayPrice="NULL"


    IF (autoIcheIdx<>"") then
        sqlStr = "Update db_partner.dbo.tbl_Esero_AutoIcheMapData "
        sqlStr = sqlStr&" set matchType="&matchType
        sqlStr = sqlStr&" , TaxSellType="&TaxSellType
        sqlStr = sqlStr&" , corpNo='"&corpNo&"'"
        sqlStr = sqlStr&" , cust_cd='"&cust_cd&"'"
        sqlStr = sqlStr&" , autoIcheTitle='"&autoIcheTitle&"'"
        sqlStr = sqlStr&" , mayPrice="&mayPrice&""
        IF (mayAcctDate="") then
            sqlStr = sqlStr&" , mayAcctDate=NULL"
        ELSE
            sqlStr = sqlStr&" , mayAcctDate='"&mayAcctDate&"'"
        ENd IF
        sqlStr = sqlStr&" , mayPumok='"&mayPumok&"'"
        IF (mayIcheDate="") then
            sqlStr = sqlStr&" , mayIcheDate=NULL"
        ELSE
            sqlStr = sqlStr&" , mayIcheDate='"&mayIcheDate&"'"
        ENd IF
        sqlStr = sqlStr&" , mayAcctJukyo='"&mayAcctJukyo&"'"
        sqlStr = sqlStr&" , AssignBizSec='"&bizSecCd&"'"
        sqlStr = sqlStr&" , AssignArap_cd='"&Arap_cd&"'"
        sqlStr = sqlStr&" where autoIcheIdx="&autoIcheIdx
        dbget.Execute sqlStr, AssignedRow

        response.write "<script>alert('"&AssignedRow&" �� ���� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
    ELSE
        sqlStr = "Insert into db_partner.dbo.tbl_Esero_AutoIcheMapData "
        sqlStr = sqlStr&" (matchType,TaxSellType,corpNo,cust_cd,autoIcheTitle,mayPrice"
        sqlStr = sqlStr&" ,mayAcctDate,mayPumok,mayIcheDate,mayAcctJukyo,AssignBizSec,AssignArap_cd )"
        sqlStr = sqlStr&" values("
        sqlStr = sqlStr&" "&matchType
        sqlStr = sqlStr&" ,"&TaxSellType
        sqlStr = sqlStr&" ,'"&corpNo&"'"
        sqlStr = sqlStr&" ,'"&cust_cd&"'"
        sqlStr = sqlStr&" ,'"&autoIcheTitle&"'"
        sqlStr = sqlStr&" ,"&mayPrice&""
        IF (mayAcctDate="") then
            sqlStr = sqlStr&" ,NULL"
        ELSE
            sqlStr = sqlStr&" ,'"&mayAcctDate&"'"
        END IF
        sqlStr = sqlStr&" ,'"&mayPumok&"'"
        IF (mayIcheDate="") then
            sqlStr = sqlStr&" ,NULL"
        ELSE
            sqlStr = sqlStr&" ,'"&mayIcheDate&"'"
        ENd IF
        sqlStr = sqlStr&" ,'"&mayAcctJukyo&"'"
        sqlStr = sqlStr&" ,'"&bizSecCd&"'"
        sqlStr = sqlStr&" ,'"&Arap_cd&"'"
        sqlStr = sqlStr&" )"

''rw sqlStr
        dbget.Execute sqlStr, AssignedRow

        response.write "<script>alert('"&AssignedRow&" �� ��� �Ǿ����ϴ�.');location.href='"&ref&"'</script>"
    ENd IF
ELSE
    response.write "mode=["&mode&"] ������"
END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->