<%

'' ���� ���ν��� ����.
function fnGetmappingTargetInfo(byval targetGb,pDate,isocno,imatchKey)
    Dim sqlStr
    IF (targetGb="1") then
        sqlStr = "select top 30 m.id,m.yyyymm,m.designerid,m.groupid "
        sqlStr = sqlStr & " , (m.ub_totalsuplycash+m.me_totalsuplycash+wi_totalsuplycash+sh_totalsuplycash+et_totalsuplycash+dlv_totalsuplycash) as jungsansum"
        sqlStr = sqlStr & " , m.finishflag, convert(varchar(10),m.taxregdate,21), m.eseroevalseq"
        sqlStr = sqlStr & " , m.billsitecode, Replace(g.company_no,'-','') as company_no, g.company_name, m.taxtype"
        sqlStr = sqlStr & " from db_jungsan.dbo.tbl_designer_jungsan_master m"
        sqlStr = sqlStr & " 	Join db_partner.dbo.tbl_partner_group g"
        sqlStr = sqlStr & " 	on m.groupid=g.groupid and Replace(g.company_no,'-','')='"&Replace(isocno,"-","")&"'"
        sqlStr = sqlStr & " where m.yyyymm>='"&Left(pDate,7)&"'"
        IF (imatchKey<>"") then
            sqlStr = sqlStr & " and m.id="&imatchKey
        end if
        sqlStr = sqlStr & " and m.cancelyn='N'"
        sqlStr = sqlStr & " order by m.id desc"

        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        IF Not (rsget.EOF OR rsget.BOF) THEN
        	fnGetmappingTargetInfo = rsget.getRows()
        END IF
        rsget.close
    ELSEIF targetGb="2" then
        sqlStr = "select top 30 m.idx,m.yyyymm,m.makerid,m.groupid "
        sqlStr = sqlStr & " , m.tot_jungsanprice"
        sqlStr = sqlStr & " , m.finishflag, convert(varchar(10),m.taxregdate,21), m.eseroevalseq"
        sqlStr = sqlStr & " , m.billsitecode, Replace(g.company_no,'-','') as company_no, g.company_name, m.taxtype"
        sqlStr = sqlStr & " from db_jungsan.dbo.tbl_off_jungsan_master m"
        sqlStr = sqlStr & " 	Join db_partner.dbo.tbl_partner_group g"
        sqlStr = sqlStr & " 	on m.groupid=g.groupid and Replace(g.company_no,'-','')='"&Replace(isocno,"-","")&"'"
        sqlStr = sqlStr & " where m.yyyymm>='"&Left(pDate,7)&"'"
        IF (imatchKey<>"") then
            sqlStr = sqlStr & " and m.idx="&imatchKey
        end if
        sqlStr = sqlStr & " order by m.idx desc"

        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        IF Not (rsget.EOF OR rsget.BOF) THEN
        	fnGetmappingTargetInfo = rsget.getRows()
        END IF
        rsget.close
    ELSEIF targetGb="9" then        '' ��Ÿ����
        sqlStr = "select top 50 p.payrequestIdx, '' as yyyymm ,'' as makerid, p.Cust_cd as groupid "
    	sqlStr = sqlStr & " , D.totprice as totalPrice"
    	sqlStr = sqlStr & " , P.payrequestState, convert(varchar(10),D.issuedate,21) as issuedate, D.etaxkey"
    	sqlStr = sqlStr & " , '' as billSiteCode, B.BIZ_NO as company_no, B.CUST_NM as company_name, D.vatKind as taxtype"
    	sqlStr = sqlStr & " , D.supplyprice,D.vatPrice,D.itemname, P.erpLinkType, P.ErpLinkKey, P.reportIdx "
    	sqlStr = sqlStr & " , P.payrequestPrice, D.payDocKind"
    	sqlStr = sqlStr & " , D.erpDocLinkType, D.erpDocLinkKey"
    	sqlStr = sqlStr & " , P.payrealDate"                                  ''������ �߰�
    	sqlStr = sqlStr & " from db_partner.dbo.tbl_eAppPayRequest P"
    	sqlStr = sqlStr & "     Join db_partner.dbo.tbl_TMS_BA_CUST B"
    	sqlStr = sqlStr & "     on P.Cust_cd=B.Cust_cd"
    	sqlStr = sqlStr & " Join db_partner.dbo.tbl_eAppPayDoc D"
    	sqlStr = sqlStr & "     On P.payrequestIdx=D.payrequestIdx"
        sqlStr = sqlStr & " where P.payRequestType in (1,2)"
        sqlStr = sqlStr & " and P.isusing=1"
        IF (imatchKey<>"") then
            sqlStr = sqlStr & " and P.payrequestIdx="&imatchKey
        end if
        sqlStr = sqlStr & " and P.payrequestDate>='"&pDate&"'"
        sqlStr = sqlStr & " and B.BIZ_NO='"&Replace(isocno,"-","")&"'"
        sqlStr = sqlStr & " order by p.payrequestIdx desc,D.issuedate desc"

        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        IF Not (rsget.EOF OR rsget.BOF) THEN
        	fnGetmappingTargetInfo = rsget.getRows()
        END IF
        rsget.close
    ELSEIF targetGb="11" then               '''����

        sqlStr = "select top 30 isNULL(T.taxIdx,0),'' as yyyymm,'' as makerid,'' as groupid "
        sqlStr = sqlStr & " , T.TotalPrice"
        sqlStr = sqlStr & " , T.isueyn as finishflag, convert(varchar(10),T.isueDate,21), T.no_iss"
        sqlStr = sqlStr & " , 'B' as billsitecode, Replace(B.busiNo,'-','') as company_no, B.businame, (CASE WHEN T.taxtype='0' THEN '2' WHEN IsNULL(T.taxtype,'Y')='N' THEN '3' ELSE '0' END) as taxtype"
        sqlStr = sqlStr & " , (T.TotalPrice-T.totalTax), T.totalTax, T.itemname"
        sqlStr = sqlStr & " from db_order.dbo.tbl_taxSheet T"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_busiInfo B"
        sqlStr = sqlStr & " 	on T.busiIdx=B.busiIdx and Replace(b.busiNo,'-','')='"&Replace(isocno,"-","")&"'"
        sqlStr = sqlStr & " where T.regdate>='"&pDate&"'"
        IF (imatchKey<>"") then
            sqlStr = sqlStr & " and T.taxIdx="&imatchKey
        end if
        sqlStr = sqlStr & " order by T.taxIdx desc"

        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        IF Not (rsget.EOF OR rsget.BOF) THEN
        	fnGetmappingTargetInfo = rsget.getRows()
        END IF
        rsget.close
    End IF
end Function

function fnGetOrMakeCUST(iCorpNo,iTaxKey,byRef icustcd)
    dim retVal : retVal=-9

    retval = fnGetCustCDByCorpNo(iCorpNo,icustcd)

    if (retval=1) or (retval=-1) then
        fnGetOrMakeCUST = retVal
        exit function
    end if

    retVal = fnMakeCustByTaxKey(iTaxKey, icustcd)
    if (retVal=1) then
        retVal = fnGetCustCDByCorpNo(iCorpNo,icustcd)
    end if
    fnGetOrMakeCUST = retVal
end function

function fnGetCustCDByCorpNo(iCorpNo,byRef icustcd)
    Dim sqlStr, retArr
    sqlStr = "select top 100 CUST_CD,CUST_NM "
    sqlStr = sqlStr & " from db_partner.dbo.tbl_TMS_BA_CUST"
    sqlStr = sqlStr & " where BIZ_NO='"&replace(iCorpNo,"-","")&"'"
    sqlStr = sqlStr & " and USE_YN='Y'"
    sqlStr = sqlStr & " and DEL_YN='N'"

    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    IF Not (rsget.EOF OR rsget.BOF) THEN
    	retArr = rsget.getRows()
    END IF
    rsget.close

    IF isArray(retArr) then
        if (UBound(retArr,2)+1)>1 then
            fnGetCustCDByCorpNo = -1        '''�̹� 2�� �̻� ���� �� ���.
        else
            icustcd = retArr(0,0)
            fnGetCustCDByCorpNo=1
        end if
    ELSE
        fnGetCustCDByCorpNo = -9        '' ����ڷ� �� �ŷ�ó ����.
    end if
end function

function fnMakeCustByTaxKey(iTaxKey,byRef icustcd)
    fnMakeCustByTaxKey = -9

    Dim clsEsero, ArrVal
    Dim sBRNTYPE : sBRNTYPE="0"               '' 0����/1����/2����(���µ�) /4��Ÿ(����) /5�¶��ΰŷ�ó/ 7���� .. [9�Һ��� ����ŷ�ó]
    Dim sCoYN                                 ''���λ���� ���λ����
    Dim sARYN : sARYN="Y"                     ''����
    Dim sAPYN : sAPYN="N"                     ''����
    Dim sBSCD : sBSCD=""                      ''����
    Dim sINTP : sINTP=""                      ''����

    Dim sPostCd
    Dim sAddr
    Dim sTelNo
    Dim sFaxNo
    Dim sTaxType
    Dim sDispSeq : sDispSeq = "9999"
    Dim sModUser : sModUser = session("ssBctId")
    Dim sBIGO

    Dim sEmpNm
    Dim sPos
    Dim sDeptNM
    Dim sEmpTel
    Dim sEmpHP
    Dim sEmpEmail

    Dim sBankcd
    Dim sAcctNo
    Dim sARAPTYPE
    Dim sSavMN
    Dim sDEFACCTYN
    Dim sPSGB  : sPSGB="1"             '' 1����ڹ�ȣ 2�ֹι�ȣ

    set clsEsero = new CEsero
    clsEsero.FtaxKey = taxKey
    ArrVal = clsEsero.fnGetEseroOneTax
    set clsEsero = Nothing

    Dim taxSellType, matchType
    Dim iCorpNo, iJongNo, iCorpName, iCeoName, iEmail

    IF IsArray(ArrVal) then
        taxSellType = ArrVal(15,0)
        matchType   = ArrVal(29,0)

        IF (taxSellType=1) and (matchType=11) THEN sBRNTYPE="9"

        IF (taxSellType="0") then
            iCorpNo = ArrVal(2,0)
        	iJongNo = ArrVal(3,0)
        	iCorpName = ArrVal(4,0)
        	iCeoName  = ArrVal(5,0)
        	iEmail    = ArrVal(6,0)
        ELSE
        	iCorpNo  = ArrVal(7,0)
        	iJongNo  = ArrVal(8,0)
        	iCorpName  = ArrVal(9,0)
        	iCeoName  = ArrVal(10,0)
        	iEmail    = ArrVal(11,0)
        ENd IF

    	IF Mid(iCorpNo,4,1)="8" Then
    	    sCoYN="Y"
    	ELSE
    	    sCoYN="N"
        END IF

        sAddr = ArrVal(43,0)
    	sBSCD = ArrVal(44,0)
    	sINTP = ArrVal(45,0)

    	Dim objCmd, prcName, returnValue
    	prcName = "db_SCM_LINK.[dbo].sp_BA_CUST_ContsInsert"

    	IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"
    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbiTms_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('"&sBRNTYPE&"', '"&sCoYN&"' ,'"&sARYN&"', '"&sAPYN&"'"&_
						+",'"&Replace(iCorpName,"'","")&"','"&iCorpNo&"','"&iCeoName&"','"&sBSCD&"','"&sINTP&"','"&sPostCd&"','"&sAddr&"','"&iEmail&"','"&sTelNo&"'"&_
						+",'"&sFaxNo&"','"&sTaxType&"','"&sDispSeq&"','"&sModUser&"','"&sBIGO&"'"&_
						+", '"&sEmpNm&"' ,'"&sPos&"', '"&sDeptNM&"','"&sEmpTel&"','"&sEmpHP&"','"&sEmpEmail&"'"&_
						+", '"&sBankcd&"' ,'"&sAcctNo&"', '"&sARAPTYPE&"','"&sSavMN&"','"&sDEFACCTYN&"',"&sPSGB&_
						+")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	    Set objCmd = nothing
rw "�ŷ�ó���� �۾�"	& returnValue
    	IF 	returnValue <>1 THEN Exit function
rw "��ϼ���"
    	''��� ����.
    	prcName = "db_partner.[dbo].sp_Ten_TMS_BA_CUST_getAllData"
        IF (application("Svr_Info")="Dev") THEN prcName = prcName & "_TEST"

    	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call "&prcName&"('','"&iCorpNo&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
	    Set objCmd = nothing
	    fnMakeCustByTaxKey = returnValue
    end if
end function

function IsERPSendAvail(imatchState, imatchType, ierpLinkType, ierpLinkKey, bizSecCD, iTargetState, iarap_cd, byREF inValidStr)
    inValidStr =""
    IsERPSendAvail = False

    if (IsNULL(imatchState)) or (imatchState=0) then
        inValidStr = "������ �������� �Ұ�."
        exit function
    end if

    if (imatchType="1")  then
        inValidStr = "�¶��� ���԰� �������� �Ұ�."
        exit function
    end if

    if (imatchType="2")  then
        inValidStr = "�������� ���԰� �������� �Ұ�."
        exit function
    end if

    if (IsNULL(bizSecCD)) or (bizSecCD="") then
        inValidStr = "����ι� ������ ���� �Ұ�."
        exit function
    end if

    if (IsNULL(iarap_cd)) or (iarap_cd="") then
        inValidStr = "�����׸� �̸�Ī ���� �Ұ�."
        exit function
    end if

    if (ierpLinkType="S") and (Not IsNULL(ierpLinkKey)) then
        inValidStr = "�� ���� ���� ���� �Ұ�."
        exit function
    end if

    IF (imatchType="9") then
        ''rw "=iTargetState="&iTargetState
        if (iTargetState<8) then
            inValidStr = "���� ���� ERP ���� �� ���� ���� �Ұ�."
            exit function
        end if
    end if

    if (ierpLinkType="H") then
        inValidStr = "�����Է� �Ϸ�� "
        exit function
    end if
    ''���ڰ��� ���� �������� ������ ���� �Ұ� proc���� check
    IsERPSendAvail = true
end function

function IsERPHandInpuAvail(imatchState, imatchType, ierpLinkType, ierpLinkKey, bizSecCD, iTargetState,iarap_cd, byREF inValidStr)
    inValidStr =""
    IsERPHandInpuAvail = False

    if (imatchType="1")  then
        inValidStr = "�¶��� ���԰� �����Է� ó�� �Ұ�."
        exit function
    end if

    if (imatchType="2")  then
        inValidStr = "�������� ���԰� �����Է� ó�� �Ұ�."
        exit function
    end if

    if (IsNULL(bizSecCD)) or (bizSecCD="") then
        inValidStr = "����ι� ������ �����Է� ó�� �Ұ�."
        exit function
    end if

    if (IsNULL(iarap_cd)) or (iarap_cd="") then
        inValidStr = "�����׸� �̸�Ī ���� �Ұ�."
        exit function
    end if

    if (ierpLinkType="S") and (Not IsNULL(ierpLinkKey)) then
        inValidStr = "�� ���� ���� �����Է� ó�� �Ұ�."
        exit function
    end if

    IF (imatchType="9") then
        if (iTargetState<8) then
            inValidStr = "���� ���� ERP ���� �� ���� �����Է� ó�� �Ұ�."
            exit function
        end if
    end if

    if (ierpLinkType="H") then
        inValidStr = "�����Է� �Ϸ�� "
        exit function
    end if
    IsERPHandInpuAvail = true
end function

function getCommonTargetStatus(imatchType,istatus)
    If IsNULL(imatchType) then Exit function

    if (imatchType=1) or (imatchType=2) or (imatchType=3) then
        getCommonTargetStatus = "<font color="&GetJungsanStateColor(istatus)&">"&GetJungsanStateName(istatus)&"</font>"
    elseif (imatchType=9) then
        getCommonTargetStatus = fnGetPayRequestState(istatus)
    end if
end function

function isPLAN_DATEDefaultSend(imatchType, itaxSellType, iarap_cd)
    isPLAN_DATEDefaultSend = true
'    if (itaxSellType="1" and iarap_cd="118") then
'        isPLAN_DATEDefaultSend = false
'        exit function
'    end if
'
    if (imatchType=999) then
        isPLAN_DATEDefaultSend = false
        exit function
    end if
end function

function getSellTypeName(itaxsellType)
    SELECT CASE itaxsellType
        CASE 0 : getSellTypeName = "����"
        CASE 1 : getSellTypeName = "����"
        CASE ELSE : getSellTypeName = itaxsellType
    ENd SELECT
end function

function gettaxModiTypeName(itaxModiType)
    SELECT CASE itaxModiType
        CASE 0 : gettaxModiTypeName = "����"
        CASE 1 : gettaxModiTypeName = "<font color=blue>����</font>"
        CASE 9 : gettaxModiTypeName = "<font color=red>����</font>"
        CASE ELSE : gettaxModiTypeName = itaxModiType
    ENd SELECT
end function

function gettaxTypeName(itaxType)
    SELECT CASE itaxType
        CASE 1 : gettaxTypeName = "����"
        CASE 2 : gettaxTypeName = "<font color=blue>����</font>"
        CASE 3 : gettaxTypeName = "<font color=red>�鼼</font>"
        CASE ELSE : gettaxTypeName = itaxType
    ENd SELECT
end function

function getMatchStateName(imatState)
    if IsNULL(imatState) then Exit function

    SELECT CASE imatState
        CASE 1 : getMatchStateName = "��Ī"
        CASE ELSE : getMatchStateName = imatState
    ENd SELECT
end function

function getMatchTypeName(imatchType)
    if IsNULL(imatchType) then Exit function

    SELECT CASE imatchType
        CASE 999 : getMatchTypeName = "<font color=blue>������꼭</font>"
        CASE 900 : getMatchTypeName = "<font color=blue>�ڵ���ü</font>"
        CASE 910 : getMatchTypeName = "<font color=blue>��Ÿ��� ��Ī</font>"
        CASE 0 : getMatchTypeName = "���� ��Ī"
        CASE 1 : getMatchTypeName = "�¶��θ���"
        CASE 2 : getMatchTypeName = "�������θ���"
        CASE 3 : getMatchTypeName = "���̶�Ҹ���"
        CASE 9 : getMatchTypeName = "��Ÿ����"
        CASE 11 : getMatchTypeName = "����"
        CASE 19 : getMatchTypeName = "��Ÿ����"
        CASE 21 : getMatchTypeName = "���������"
        CASE 22 : getMatchTypeName = "���������"
        CASE ELSE : getMatchTypeName = imatchType
    ENd SELECT
end function

function getbizSecCDName(iBizCd)
    if IsNULL(iBizCd) then Exit function

    SELECT CASE iBizCd
        CASE "0000000101" : getbizSecCDName = "�¶���(����)"

        CASE "0000000201" : getbizSecCDName = "��������"
        CASE "0000000202" : getbizSecCDName = "���з�1������"
        CASE "0000000203" : getbizSecCDName = "���з�2������"
        CASE "0000000204" : getbizSecCDName = "��õCGV��"
        CASE "0000000205" : getbizSecCDName = "��Ÿ��"
        CASE "0000000206" : getbizSecCDName = "��õ������"
        CASE "0000000207" : getbizSecCDName = "�����Ե���"
        CASE "0000000208" : getbizSecCDName = "�����ö���"

        CASE "0000000301" : getbizSecCDName = "���̶��"
        CASE "0000000302" : getbizSecCDName = "���̶��ȫ�����"
        CASE "0000000303" : getbizSecCDName = "���̶�����ο�ǳ"

        CASE "0000000401" : getbizSecCDName = "��ī����"
        CASE "0000000402" : getbizSecCDName = "ī��1010"

        CASE "0000000501" : getbizSecCDName = "����"
        CASE "0000000502" : getbizSecCDName = "CS"
        CASE "0000000503" : getbizSecCDName = "�濵"
        CASE "0000000504" : getbizSecCDName = "SYS"
        CASE "0000000505" : getbizSecCDName = "����"

        CASE "0000000990" : getbizSecCDName = "��Ÿ"
        CASE "0000009010" : getbizSecCDName = "����Ⱥ�"
        CASE ELSE : getbizSecCDName = iBizCd
    ENd SELECT
end function

public function GetJungsanStateName(ifinishflag)
    if (IsNULL(ifinishflag)) then Exit function

    if ifinishflag="0" then
    	GetJungsanStateName = "������"
    elseif ifinishflag="1" then
        GetJungsanStateName = "��üȮ�δ��"
    elseif ifinishflag="2" then
        GetJungsanStateName = "��üȮ�οϷ�"
    elseif ifinishflag="3" then
    	GetJungsanStateName = "����Ȯ��"
    elseif ifinishflag="7" then
    	GetJungsanStateName = "�ԱݿϷ�"
    else

    end if
end function

public function GetJungsanStateColor(ifinishflag)
    if (IsNULL(ifinishflag)) then Exit function

    if ifinishflag="0" then
    	GetJungsanStateColor = "#000000"
    elseif ifinishflag="1" then
        GetJungsanStateColor = "#448888"
    elseif ifinishflag="2" then
        GetJungsanStateColor = "#0000FF"
    elseif ifinishflag="3" then
    	GetJungsanStateColor = "#0000FF"
    elseif ifinishflag="7" then
    	GetJungsanStateColor = "#FF0000"
    else

    end if
end function

public function GetJungsanTaxtypeName(itaxtype)
	if itaxtype="01" then
		GetJungsanTaxtypeName = "����"
	elseif itaxtype="02" then
		GetJungsanTaxtypeName = "�鼼"
	elseif itaxtype="03" then
		GetJungsanTaxtypeName = "��õ" '''"����" '''����?
	end if
end function

public function GetEAppTaxtypeName(itaxtype)
	if itaxtype="0" then
		GetEAppTaxtypeName = "����"
	elseif itaxtype="2" then
		GetEAppTaxtypeName = "�鼼"
	elseif itaxtype="3" then
		GetEAppTaxtypeName = "��õ" '''"����" '''����?
	end if
end function

Class CAutoIcheMapData
    public FautoIcheIdx
    public FmatchType
    public FTaxSellType
    public FcorpNo
    public Fcust_cd
    public FcorpName
    public FautoIcheTitle
    public FmayPrice
    public FmayAcctDate
    public FmayPumok
    public FmayIcheDate
    public FmayAcctJukyo
    public FAssignBizSec
    public FAssignArap_cd
    public FAssignBizSecName
    public FAssignArapNm
    public FtaxKey
    public FappDate
    public FsellCorpNo
    public FsellJongNo
    public FsellCorpName
    public FsellCeoName
    public FsellEmail
    public FbuyCorpNo
    public FbuyJongNo
    public FBuyCorpName
    public FBuyCeoName
    public FbuyEmail
    public FtotSum
    public FsuplySum
    public FtaxSum
    public FtaxModiType
    public FtaxType
    public FevalTypeNm
    public FBigo
    public FDtlName
    public Fmatchstate
    public Fbizseccd
    public FerplinkType
    public FerplinkKey
End Class

Class CEsero
    public FOneItem
    public FItemList()
    public FTotalCount
    public FSDate
    public FEDate
    public FsearchText
    public FtaxsellType
    public FtaxModiType
    public FtaxType

    public FTotCnt
    public FPageSize
    public FCurrPage
    public FSPageNo
    public FEPageNo
    public FResultCount

    public FMappingTypeYn
    public FMappingType
    public FTaxKey
    public FRectCorpNo
    public FErpSendType
    public FTotSum

    public FRectautoIcheIdx
    public FRectMatchType
    public FRectTaxSellType
    public FRectautoIcheTitle
    public FRectmayPumok
    public FmayAcctJukyo
    public FRectmayPrice
    public FRectBizSecCd
    public FExpectType

	public FRectArapCD

	'//admin/tax/popsendnottax.asp
	public sub getsendnottax()
		dim sqlStr,i

		'������ ����Ʈ
		sqlStr = "select top 7000"
		sqlStr = sqlStr & " t.taxKey, t.appDate, t.sellCorpNo, t.sellJongNo, t.sellCorpName, t.sellCeoName"
		sqlStr = sqlStr & " , t.sellEmail, t.buyCorpNo, t.BuyCorpName, t.BuyCeoName, t.buyEmail"
		sqlStr = sqlStr & " , t.totSum, t.suplySum, t.taxSum"
		sqlStr = sqlStr & " ,(case when t.taxSellType = 0 then '����' else '����' end) as taxSellType"
		sqlStr = sqlStr & " ,(case when t.taxModiType = 0 then '����' else '����' end) as taxModiType"
		sqlStr = sqlStr & " ,(case"
		sqlStr = sqlStr & "  	when t.taxType = 1 then '����'"
		sqlStr = sqlStr & "  	when t.taxType = 2 then '����'"
		sqlStr = sqlStr & " else '�鼼' end) as taxType"
		sqlStr = sqlStr & " ,t.Bigo, t.DtlName"
		sqlStr = sqlStr & " ,m.bizseccd"
		sqlStr = sqlStr & " from db_partner.dbo.tbl_esero_tax T"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_esero_taxMatch M"
		sqlStr = sqlStr & " 	on T.taxkey=M.taxkey"
		sqlStr = sqlStr & " 	and M.matchseq=0"
		sqlStr = sqlStr & " where M.erpLinkType is NULL"
		sqlStr = sqlStr & " order by bizseccd,appdate"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new CAutoIcheMapData

				FItemList(i).ftaxKey = rsget("taxKey")
				FItemList(i).fappDate = rsget("appDate")
				FItemList(i).fsellCorpNo = rsget("sellCorpNo")
				FItemList(i).fsellJongNo = rsget("sellJongNo")
				FItemList(i).fsellCorpName = rsget("sellCorpName")
				FItemList(i).fsellCeoName = rsget("sellCeoName")
				FItemList(i).fsellEmail = rsget("sellEmail")
				FItemList(i).fbuyCorpNo = rsget("buyCorpNo")
				FItemList(i).fBuyCorpName = rsget("BuyCorpName")
				FItemList(i).fBuyCeoName = rsget("BuyCeoName")
				FItemList(i).fbuyEmail = rsget("buyEmail")
				FItemList(i).ftotSum = rsget("totSum")
				FItemList(i).fsuplySum = rsget("suplySum")
				FItemList(i).ftaxSum = rsget("taxSum")
				FItemList(i).ftaxSellType = rsget("taxSellType")
				FItemList(i).ftaxModiType = rsget("taxModiType")
				FItemList(i).ftaxType = rsget("taxType")
				FItemList(i).fBigo = rsget("Bigo")
				FItemList(i).fDtlName = rsget("DtlName")
				FItemList(i).fbizseccd = rsget("bizseccd")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    public function getMonthTaxList()
		dim sqlStr,i

		'������ ����Ʈ
		sqlStr = "select top 10000"
		sqlStr = sqlStr & " t.taxKey, t.appDate, t.sellCorpNo, t.sellJongNo, t.sellCorpName, t.sellCeoName"
		sqlStr = sqlStr & " , t.sellEmail, t.buyCorpNo, t.BuyCorpName, t.BuyCeoName, t.buyEmail"
		sqlStr = sqlStr & " , t.totSum, t.suplySum, t.taxSum"
		sqlStr = sqlStr & " ,(case when t.taxSellType = 0 then '����' else '����' end) as taxSellType"
		sqlStr = sqlStr & " ,(case when t.taxModiType = 0 then '����' else '����' end) as taxModiType"
		sqlStr = sqlStr & " ,(case"
		sqlStr = sqlStr & "  	when t.taxType = 1 then '����'"
		sqlStr = sqlStr & "  	when t.taxType = 2 then '����'"
		sqlStr = sqlStr & " else '�鼼' end) as taxType"
		sqlStr = sqlStr & " ,t.Bigo, t.DtlName"
		sqlStr = sqlStr & " ,m.bizseccd"
		sqlStr = sqlStr & " from db_partner.dbo.tbl_esero_tax T"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_esero_taxMatch M"
		sqlStr = sqlStr & " 	on T.taxkey=M.taxkey"
		sqlStr = sqlStr & " 	and M.matchseq=0"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & " and T.appDate>='"&FSDate&"'"
		sqlStr = sqlStr & " and T.appDate<'"&FEDate&"'"
		sqlStr = sqlStr & " order by bizseccd,appdate"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount
        if  not rsget.EOF  then
            getMonthTaxList = rsget.getRows()
        end if
        rsget.Close

	end function

    public Function fnGetEseroTaxMatchExpectList
        Dim strSql

    	IF FtaxsellType="" THEN FtaxsellType = -1
    	IF FtaxModiType="" THEN FtaxModiType = -1
    	IF FtaxType="" THEN FtaxType = -1
        IF FMappingType="" THEN FMappingType=-1
        IF (FTotSum="") THEN FTotSum="NULL"

		strSql ="[db_partner].[dbo].[sp_Ten_Esero_Tax_getMatchExpectListCnt]('"&FExpectType&"','"&FSDate&"','"&FEDate&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		IF FTotCnt > 0 THEN
    		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
    		FEPageNo = FPageSize*FCurrPage

    		strSql ="[db_partner].[dbo].sp_Ten_Esero_Tax_getMatchExpectList('"&FExpectType&"','"&FSDate&"','"&FEDate&"',"&FsPageNO&","&FePageNO&")"
''rw strSql
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			fnGetEseroTaxMatchExpectList = rsget.getRows()
    		END IF
    		rsget.close
		END IF
	End Function

	public Function fnGetEseroTaxList
		Dim strSql

    	IF FtaxsellType="" THEN FtaxsellType = -1
    	IF FtaxModiType="" THEN FtaxModiType = -1
    	IF FtaxType="" THEN FtaxType = -1
        IF FMappingType="" THEN FMappingType=-1
        IF (FTotSum="") THEN FTotSum="NULL"
		IF FRectArapCD="" THEN FRectArapCD = -1

		strSql ="[db_partner].[dbo].[sp_Ten_Esero_Tax_getListCnt]('"&FSDate&"','"&FEDate&"','"&FsearchText&"',"&FtaxsellType&","&FtaxModiType&","&FtaxType&",'"&FMappingTypeYn&"',"&FMappingType&",'"&FRectCorpNo&"','"&FErpSendType&"',"&FTotSum&",'"&FRectBizSecCd&"','"&FRectArapCD&"')"
		''rw strSql

		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
    		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
    		FEPageNo = FPageSize*FCurrPage

    		strSql ="[db_partner].[dbo].sp_Ten_Esero_Tax_getList('"&FSDate&"','"&FEDate&"','"&FsearchText&"',"&FtaxsellType&","&FtaxModiType&","&FtaxType&","&FsPageNO&","&FePageNO&",'"&FMappingTypeYn&"',"&FMappingType&",'"&FRectCorpNo&"','"&FErpSendType&"',"&FTotSum&",'"&FRectBizSecCd&"','"&FRectArapCD&"')"
''rw strSql
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			fnGetEseroTaxList = rsget.getRows()
    		END IF
    		rsget.close
		END IF
	End Function

	Function fnGetEseroOneTax()
	    Dim strSql
	    strSql ="[db_partner].[dbo].sp_Ten_Esero_getOneTax('"&FTaxKey&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEseroOneTax = rsget.getRows()
		END IF
		rsget.close
	End Function

	function fnGetMappingList()
        Dim strSql
	    strSql ="[db_partner].[dbo].sp_Ten_Esero_getMappingList('"&FTaxKey&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetMappingList = rsget.getRows()
		END IF
		rsget.close
    end function

    Function fnGetAutoIcheMapOne()
        Dim strSql, ArrList
        strSql ="[db_partner].[dbo].sp_Ten_Esero_getgetOneIcheMapData("&FRectautoIcheIdx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		FResultCount = 0
		FTotCnt = FResultCount
		IF Not (rsget.EOF OR rsget.BOF) THEN
		    FResultCount = 1
		    FTotCnt = FResultCount
		    ArrList = rsget.getRows()
			set FOneItem = new CAutoIcheMapData
	        FOneItem.FautoIcheIdx   = ArrList(0,0)
	        FOneItem.FmatchType     = ArrList(1,0)
            FOneItem.FTaxSellType   = ArrList(2,0)
            FOneItem.FcorpNo        = ArrList(3,0)

            FOneItem.FautoIcheTitle = ArrList(4,0)
            FOneItem.FmayPrice      = ArrList(5,0)
            FOneItem.FmayAcctDate   = ArrList(6,0)
            FOneItem.FmayPumok      = ArrList(7,0)
            FOneItem.FmayIcheDate   = ArrList(8,0)
            FOneItem.FmayAcctJukyo  = ArrList(9,0)
            FOneItem.FAssignBizSec  = ArrList(10,0)
            FOneItem.FAssignArap_cd = ArrList(11,0)
            FOneItem.Fcust_cd       = ArrList(12,0)
            FOneItem.FcorpName      = ArrList(13,0)
            FOneItem.FAssignBizSecName   = ArrList(14,0)
            FOneItem.FAssignArapNm       = ArrList(15,0)
		END IF
		rsget.close
	end function

    Function fnGetAutoIcheMapDataList()
        Dim strSql, ArrList, i
        IF FRectautoIcheIdx="" then FRectautoIcheIdx="NULL"
        IF FRectMatchType="" then FRectMatchType="NULL"
        IF FRectTaxSellType="" then FRectTaxSellType="NULL"
        IF FRectmayPrice="" then FRectmayPrice="NULL"

        strSql ="[db_partner].[dbo].[sp_Ten_Esero_getAutoIcheMapDataCnt]("&FRectautoIcheIdx&",'"&FRectcorpNo&"',"&FRectMatchType&","&FRectTaxSellType&",'"&FRectautoIcheTitle&"','"&FRectmayPumok&"','"&FmayAcctJukyo&"',"&FRectmayPrice&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
    		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
    		FEPageNo = FPageSize*FCurrPage

    	    strSql ="[db_partner].[dbo].sp_Ten_Esero_getAutoIcheMapDataList("&FRectautoIcheIdx&",'"&FRectcorpNo&"',"&FRectMatchType&","&FRectTaxSellType&",'"&FRectautoIcheTitle&"','"&FRectmayPumok&"','"&FmayAcctJukyo&"',"&FRectmayPrice&","&FsPageNO&","&FePageNO&")"
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			ArrList = rsget.getRows()
    		END IF
    		rsget.close

    		If IsArray(ArrList) then
    		    FResultCount = UBound(ArrList,2)+1
    		    redim preserve FItemList(FResultCount)
    		    For i=0 to FResultCount-1
    		        set FItemList(i) = new CAutoIcheMapData
    		        FItemList(i).FautoIcheIdx   = ArrList(0,i)
    		        FItemList(i).FmatchType     = ArrList(1,i)
                    FItemList(i).FTaxSellType   = ArrList(2,i)
                    FItemList(i).FcorpNo        = ArrList(3,i)

                    FItemList(i).FautoIcheTitle = ArrList(4,i)
                    FItemList(i).FmayPrice      = ArrList(5,i)
                    FItemList(i).FmayAcctDate   = ArrList(6,i)
                    FItemList(i).FmayPumok      = ArrList(7,i)
                    FItemList(i).FmayIcheDate   = ArrList(8,i)
                    FItemList(i).FmayAcctJukyo  = ArrList(9,i)
                    FItemList(i).FAssignBizSec  = ArrList(10,i)
                    FItemList(i).FAssignArap_cd = ArrList(11,i)
                    FItemList(i).Fcust_cd       = ArrList(12,i)
                    FItemList(i).FcorpName      = ArrList(13,i)
                    FItemList(i).FAssignBizSecName   = ArrList(14,i)
                    FItemList(i).FAssignArapNm       = ArrList(15,i)
    		    Next
    		end if
        END IF
    end function
End Class
%>
