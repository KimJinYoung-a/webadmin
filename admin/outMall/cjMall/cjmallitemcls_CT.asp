<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "cjmall"
CONST CMAXLIMITSELL = 5        '' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CCJMALLMARGIN = 12       ''���� 12%...// �� 12? // 2013-11-05 ������..12->15�� ���� =>12�� ���� ������.(2013/11/21)
CONST CitemGbnKey ="K1099999" ''��ǰ����Ű ''�ϳ��� ����
CONST CUPJODLVVALID = True   ''��ü ���ǹ�� ��� ���ɿ���

CONST CVENDORID = 411378					'���¾�ü�ڵ�
CONST CVENDORCERTKEY = "CJ03074113780"		'����Ű
CONST CUNIQBRANDCD = 24049000				'�귣���ڵ�
CONST MD_CODE = "5103"						'MD_Code

Class cjmallItem
	Public Fitemid
	Public Fitemname
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public Fsellcash
	Public Fbuycash
	Public FsellYn
	Public Fsaleyn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fdeliverytype
	Public FoptionCnt
	Public FcjmallRegdate
	Public FcjmallLastUpdate
	Public FcjmallPrdNo
	Public FcjmallPrice
	Public FcjmallSellYn
	Public FregUserid
	Public FcjmallStatCd
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT
	Public FlastErrStr
	Public FCateMapCnt
	Public FcdmKey
	Public FdefaultfreeBeasongLimit
	'ī�װ�
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FDispNo
	Public FDispNm
	Public FDispLrgNm
	Public FDispMidNm
	Public FDispSmlNm
	Public FDispThnNm
	Public FCateIsUsing
	Public Fdisptpcd
	Public FisUsing

	'��ǰ�з�
	Public Finfodiv
	Public Ficnt
	Public FCddKey

	Public Fcdd_Name
	Public Fcdl_Name
	Public Fcdm_Name
	Public Fcds_Name
	Public FPrdDivIsUsing

	Public FRectMode
	Public FRectItemID

	Public FCdm
	Public FCdd

	'��ǰ��� ��Ī
	Public FitemDiv
	Public ForgSuplyCash
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FitemGbnKey
	Public Fdeliverfixday

	Public FItemOption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public Fsocname_kor

	Public MustPrice

    public function getItemNameFormat()
        dim buf
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","����")
        buf = replace(buf,"[������]","")
        buf = replace(buf,"[���� ���]","")
        getItemNameFormat = buf
    end function

	'// ǰ������
	Public Function IsSoldOut()
		ISsoldOut = (FSellyn <> "Y") or ((FLimitYn = "Y") and (FLimitNo - FLimitSold < 1))
	End Function

    public Function IsCjFreeBeasong()
        IsCjFreeBeasong = False
        ''IsCjFreeBeasong = (FSellcash>=30000)  ''�������̾�� (�����ۿ��� Y ���� 3���� �̻��̸� ���������� ǥ�õ�)
    end Function

'	Public Function IsFreeBeasong()
'		IsFreeBeasong = False
'		If (FdeliveryType = 2) or (FdeliveryType = 4) or (FdeliveryType = 5) Then
'			IsFreeBeasong = True
'		End If
'	End Function

	function getLimitEa()
		Dim ret : ret = (FLimitno - FLimitSold)
		If (ret < 1) Then ret = 0
		getLimitEa = ret
	end function

	Function getLimitHtmlStr()
		If IsNULL(FLimityn) Then Exit Function

		If (FLimityn="Y") Then
			getLimitHtmlStr = "<font color=blue>����:"&getLimitEa&"</font>"
		End If
	End Function

	Public Function getcjmallStatName
	    If IsNULL(FcjmallStatCd) then FcjmallStatCd=-1
		Select Case FcjmallStatCd
			CASE -9 : getcjmallStatName = "�̵��"
			CASE -2 : getcjmallStatName = "<font color=red>�ݷ�</font>"
			CASE -1 : getcjmallStatName = "��Ͻ���"
			CASE 0 : getcjmallStatName = "<font color=blue>��Ͽ���</font>"
			CASE 1 : getcjmallStatName = "���۽õ�"
			CASE 3 : getcjmallStatName = "���δ��"
			CASE 7 : getcjmallStatName = getLimitHtmlStr ''"" ''��ϿϷ�
			CASE ELSE : getcjmallStatName = FcjmallStatCd
		End Select
	End Function

	Function getDispGubunNm()
		getDispGubunNm = getDisptpcdName
	End Function

	Public Function getDisptpcdName
		If (Fdisptpcd="B") Then
			getDisptpcdName = "<font color='blue'>����</font>"
		Elseif (Fdisptpcd = "D") Then
			getDisptpcdName = "�Ϲ�"
		Else
			getDisptpcdName = Fdisptpcd
		End if
	End Function

	'ȭ����� ����
	Public Function getdeliverfixday()
		If (Fdeliverfixday = "C") or (Fdeliverfixday = "X") Then
			getdeliverfixday = 20
		Else
			getdeliverfixday = 10
		End If
	End Function

    ''�ֹ����� ����
    Public Function getzCostomMadeInd()
		dim ret, CMadeInd
        ret = (Fitemdiv="06" or Fitemdiv="16")
        ret = ret or (FtenCateLarge="010" and FtenCateMid="070" and FtenCateSmall="070")	'�����ι���	������	�ֹ�����
		ret = ret or (FtenCateLarge="035" and FtenCateMid="016" and FtenCateSmall="010")	'����/���	����̺�	������
		ret = ret or (FtenCateLarge="040")													'����
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="001")	'����/��Ȱ	����/������ǰ	������
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="002")	'����/��Ȱ	����/������ǰ	ƴ��������
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="005")	'����/��Ȱ	����/������ǰ	��������
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	'����/��Ȱ	����/������ǰ	�����̼�����
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	'����/��Ȱ	����/������ǰ	�����̼�����
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="019")	'����/��Ȱ	����/������ǰ	�̵��ļ�����
		ret = ret or (FtenCateLarge="045" and FtenCateMid="003")							'����/��Ȱ	����ũ����
		ret = ret or (FtenCateLarge="045" and FtenCateMid="006")							'����/��Ȱ	���ڼ���
		ret = ret or (FtenCateLarge="045" and FtenCateMid="007" and FtenCateSmall="008")	'����/��Ȱ	Ű�����	Ű�� ������
		ret = ret or (FtenCateLarge="050" and FtenCateMid="010" and FtenCateSmall="050")	'Ȩ/����	����	�̴ϼ�/�޼�������
		ret = ret or (FtenCateLarge="050" and FtenCateMid="030" and FtenCateSmall="010")	'Ȩ/����	��ļ�ǰ	�̴ϼ����
		ret = ret or (FtenCateLarge="050" and FtenCateMid="045" and FtenCateSmall="120")	'Ȩ/����	Ȩ������	���۾� �ֹ�����
		ret = ret or (FtenCateLarge="055" and FtenCateMid="070")							'�к긯 > ħ����Ʈ
		ret = ret or (FtenCateLarge="055" and FtenCateMid="080")							'�к긯 > Ŀư
		ret = ret or (FtenCateLarge="055" and FtenCateMid="090")							'�к긯 > ���/�漮
		ret = ret or (FtenCateLarge="055" and FtenCateMid="100")							'�к긯 > ��Ʈ/����
		ret = ret or (FtenCateLarge="055" and FtenCateMid="110")							'�к긯 > �к긯��ǰ
		ret = ret or (FtenCateLarge="055" and FtenCateMid="120")							'�к긯 > ħ����ǰ
		ret = ret or (FtenCateLarge="060" and FtenCateMid="130")							'Űģ > �۰� ��Ȱ�ڱ�
		ret = ret or (FtenCateLarge="070" and FtenCateMid="160")							'����/����/��� > ���
		ret = ret or (FtenCateLarge="090" and FtenCateMid="070" and FtenCateSmall="010")	'Men > ���/��ȭ > �ð�/���
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="020")	'���̺� > ����/ħ��/���� > ���ڽ�ƼĿ/����
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="040")	'���̺� > ����/ħ��/���� > ������/å����
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="050")	'���̺� > ����/ħ��/���� > ����
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="060")	'���̺� > ����/ħ��/���� > ����/����
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="066")	'���̺� > ����/ħ��/���� > ���̺�/å��
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="070")	'���̺� > ����/ħ��/���� > ������ǰ
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="100")	'���̺� > ����/ħ��/���� > �Ʊ�ħ��
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="110")	'���̺� > ����/ħ��/���� > �÷��̸�Ʈ
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="120")	'���̺� > ����/ħ��/���� > �����/�Ʊ���
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="130")	'���̺� > ����/ħ��/���� > ���
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="140")	'���̺� > ����/ħ��/���� > ���/ħ��/Ŀư
		If ret Then
			CMadeInd = "Y"
		Else
			CMadeInd = "N"
		End If
        getzCostomMadeInd = CMadeInd
    End Function

    ''����Ÿ�� ���
    Public Function getzLeadTime()
		If (FtenCateLarge = "010" and FtenCateMid = "070" and FtenCateSmall = "070") OR (FtenCateLarge="035" and FtenCateMid="016" and FtenCateSmall="010") OR (FtenCateLarge="050" and FtenCateMid="010" and FtenCateSmall="050") OR (FtenCateLarge="050" and FtenCateMid="030" and FtenCateSmall="010") OR (FtenCateLarge="050" and FtenCateMid="045" and FtenCateSmall="120") OR (FtenCateLarge="060" and FtenCateMid="130") OR (FtenCateLarge="070" and FtenCateMid="160") OR (FtenCateLarge="090" and FtenCateMid="070" and FtenCateSmall="010") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="020") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="060") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="070") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="110") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="130") Then
			getzLeadTime = "08"
		ElseIf (FtenCateLarge="040") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="001") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="002")	or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="005") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="019") or (FtenCateLarge="045" and FtenCateMid="003")	or (FtenCateLarge="045" and FtenCateMid="006") or (FtenCateLarge="045" and FtenCateMid="007" and FtenCateSmall="008")	or (FtenCateLarge="055" and FtenCateMid="070") or (FtenCateLarge="055" and FtenCateMid="080")	or (FtenCateLarge="055" and FtenCateMid="090") or (FtenCateLarge="055" and FtenCateMid="100")	or (FtenCateLarge="055" and FtenCateMid="110") or (FtenCateLarge="055" and FtenCateMid="120")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="040") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="050")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="066") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="100")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="120") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="140") Then
			getzLeadTime = "08"
		End If
	End Function

	'//cjmall ��ϻ��� ��ȯ
	Public Function getcjItemStatCd()
	    getcjItemStatCd = getcjmallStatName
	End Function

    Function getCJmallSuplyPrice(optaddprice)
'        getCJmallSuplyPrice = CLNG(FSellCash * (100-CCJMALLMARGIN) / 100)
		'�ϴ��� CJ�޴��� ���� ����
		'* ������ Ȯ�ο���
		'1. ������ǰ : ���Կ���(VAT����) = Round(�ǸŰ�/1.1 - 0.1 * (�ǸŰ�/1.1)), 0)
		'2. �鼼��ǰ : ���Կ���(VAT����) = Round(�ǸŰ� - 0.1 * �ǸŰ�, 0)
		If FVatInclude = "Y" Then		'����
			getCJmallSuplyPrice = Round((MustPrice+optaddprice) /1.1 - (CCJMALLMARGIN/100) * ((MustPrice+optaddprice)/1.1))
		Else							'�鼼
			getCJmallSuplyPrice = Round((MustPrice+optaddprice) - (CCJMALLMARGIN/100) * (MustPrice+optaddprice))
		End If
    End Function

    Function getCJmallSuplyPrice2()
'        getCJmallSuplyPrice2 = CLNG(FSellCash * (100-CCJMALLMARGIN) / 100)
		'�ϴ��� CJ�޴��� ���� ����
		'* ������ Ȯ�ο���
		'1. ������ǰ : ���Կ���(VAT����) = Round(�ǸŰ�/1.1 - 0.1 * (�ǸŰ�/1.1)), 0)
		'2. �鼼��ǰ : ���Կ���(VAT����) = Round(�ǸŰ� - 0.1 * �ǸŰ�, 0)
		If FVatInclude = "Y" Then		'����
			getCJmallSuplyPrice2 = Round((MustPrice) /1.1 - (CCJMALLMARGIN/100) * ((MustPrice)/1.1))
		Else							'�鼼
			getCJmallSuplyPrice2 = Round((MustPrice) - (CCJMALLMARGIN/100) * (MustPrice))
		End If
    End Function

    public function getDeliverytypeName
        if (Fdeliverytype="9") then
            getDeliverytypeName = "<font color='blue'>[���� "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
        elseif (Fdeliverytype="7") then
            getDeliverytypeName = "<font color='red'>[��ü����]</font>"
        elseif (Fdeliverytype="2") then
            getDeliverytypeName = "<font color='blue'>[��ü]</font>"
        else
            getDeliverytypeName = ""
        end if
    end function

	'// ��ǰ���: MD��ǰ�� �� ���� ī�װ� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getCjCateParamToReg()
		Dim strSql, strRst, i
		strSql = ""
		strSql = strSql & " SELECT top 100 c.CateKey "
		strSql = strSql & " FROM db_outMall.dbo.tbl_cjmall_cate_mapping as m "
		strSql = strSql & " JOIN db_outMall.dbo.tbl_cjMall_Category as c on m.CateKey = c.CateKey "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
		strSql = strSql & " ORDER BY c.cateGbn ASC " ''B : �귣�� / D : �Ϲ�
		rsCTget.Open strSql,dbCTget,1
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			strRst = ""
			i = 0
			Do until rsCTget.EOF
				If i = 0 Then
					strRst = strRst &"		<tns:mallCtg>"
					strRst = strRst &"			<tns:mainInd>Y</tns:mainInd>"
					strRst = strRst &"			<tns:ctgName>" & rsCTget("CateKey") & "</tns:ctgName>"
					strRst = strRst &"		</tns:mallCtg>"
				Else
					strRst = strRst &"		<tns:mallCtg>"
					strRst = strRst &"			<tns:ctgName>" & rsCTget("CateKey") & "</tns:ctgName>"
					strRst = strRst &"		</tns:mallCtg>"
				End If
				rsCTget.MoveNext
				i = i + 1
			Loop
		End If
		rsCTget.Close
		getCjCateParamToReg = strRst
	End Function

	'��ǰǰ������
    public function getCjmallItemInfoCdToReg()
		Dim strSql, buf, addSql
		Dim mallinfoCd,infoContent,infotype, infocd, mallinfodiv
		Dim chkInfodiv, chkCdmKey

		strSql = ""
		strSql = strSql & " SELECT top 1 PD.infodiv, PD.cdmKey " & vbcrlf
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i  " & vbcrlf
		strSql = strSql & " INNER JOIN db_AppWish.dbo.tbl_item_contents as c on i.itemid = c.itemid  " & vbcrlf
		strSql = strSql & " LEFT JOIN db_outMall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and c.infodiv = PD.infodiv " & vbcrlf
		strSql = strSql & " WHERE i.itemid ='"&FItemID&"' "
		rsCTget.Open strSql,dbCTget,1
		If Not(rsCTget.EOF or rsCTget.BOF) then
			chkInfodiv	= rsCTget("infodiv")
			chkCdmKey	= rsCTget("cdmKey")
		End If
		rsCTget.Close

		If chkInfodiv = "01" and chkCdmKey = "1006" Then
			addSql = " and M.infocd <> '00000'  "
		End If



		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , " & vbcrlf
		strSql = strSql & "		CASE " & vbcrlf
		'strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) = 'KCC' AND (IC.infoDiv not in ('06','23')) THEN 'Y' " & vbcrlf
		'strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) <> 'KCC' AND (IC.infoDiv not in ('06','23')) THEN 'N' " & vbcrlf
		'strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (IC.infoDiv in ('06','23')) THEN 'Y' " & vbcrlf
		'strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN 'Y' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='Y' THEN 'Y' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='P' THEN 'I' " & vbcrlf
		strSql = strSql & "		ELSE 'I' " & vbcrlf
		strSql = strSql & " END AS infoType, " & vbcrlf
		strSql = strSql & "		CASE " & vbcrlf
'		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) = 'KCC' AND (IC.infoDiv not in ('06','23')) THEN IC.safetyNum " & vbcrlf
'		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) <> 'KCC' AND (IC.infoDiv not in ('06','23')) THEN '�ش����' " & vbcrlf
'		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (IC.infoDiv in ('06','23')) THEN IC.safetyNum " & vbcrlf
'		strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '�ش����' " & vbcrlf
        strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '�ش����' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00001') THEN '�ش����' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00002') AND (M.mallinfoCd='25044') THEN '�ش����' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00003') THEN '�󼼳�������' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN '�ش����' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='P' AND c.infoCd <> '22009' THEN '�ٹ����� ���ູ���� 1644-6035' " & vbcrlf
		'strSql = strSql & "		WHEN c.infotype='P' THEN replace(F.infocontent,'1644-6030','1644-6035') " & vbcrlf
		strSql = strSql & "		ELSE convert(varchar(500),F.infocontent) " & vbcrlf
		strSql = strSql & " END AS infocontent " & vbcrlf
		strSql = strSql & " FROM db_outMall.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " INNER JOIN db_AppWish.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " LEFT JOIN db_AppWish.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " LEFT JOIN db_AppWish.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' " & vbcrlf
		strSql = strSql & " WHERE M.mallid = 'cjmall' and IC.itemid='"&FItemID&"' " & addSql
		rsCTget.Open strSql,dbCTget,1

		If Not(rsCTget.EOF or rsCTget.BOF) then
			Do until rsCTget.EOF
			    mallinfoCd  = rsCTget("mallinfoCd")
			    infotype	= rsCTget("infotype")
			    infoContent = rsCTget("infoContent")
				infocd		= rsCTget("infocd")
				mallinfodiv = rsCTget("mallinfodiv")

                if (mallinfodiv="02") and (mallinfoCd="25012") and (infoContent="") then  '' ����/������
                    infoContent="�ش����"
                end if

                If (FItemID = "674455" OR FItemID = "881879")  AND (mallinfoCd = "25008" OR mallinfoCd = "25013") Then	'2013-06-25 ������ ����(��ü�� ������ -(������)���� ����� ���..�̷� ��찡 ���涧���� �б���� �� �� ����
                	infoContent = "������"
                End If

				buf = buf &"	<tns:goodsReport>"
				buf = buf &"		<tns:pedfId>"&mallinfoCd&"</tns:pedfId>"
				buf = buf &"		<tns:html><![CDATA["&infoContent&"]]></tns:html>"
				buf = buf &"	</tns:goodsReport>"
				rsCTget.MoveNext
			Loop
		End If
		rsCTget.Close

'2014-06-09 ������ �ϴ� �ּ� ���� / db_outMall.dbo.tbl_OutMall_infoCodeMap�� �ϴ� �ڵ�(25066) ���ԿϷ�
'		if chkInfodiv = "19" and chkCdmKey = "8504" Then  ''����/��ű� : �ð� �ΰ�� 25066 �ʿ�
'            buf = buf &"	<tns:goodsReport>"
'			buf = buf &"		<tns:pedfId>25066</tns:pedfId>"
'			buf = buf &"		<tns:html><![CDATA[������]]></tns:html>"
'			buf = buf &"	</tns:goodsReport>"
'        end if

		getCjmallItemInfoCdToReg = buf
	End Function

	public function GetCJLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetCJLmtQty = 0
			Else
				GetCJLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetCJLmtQty = 999
		End If
	End Function

	Public Function getOptionLimitNo()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (IsOptionSoldOut) Then
			getOptionLimitNo = 0
		Else
			If (Foptlimityn = "Y") Then
				If (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO) Then
					getOptionLimitNo = 0
				Else
					getOptionLimitNo = Foptlimitno - Foptlimitsold - CLIMIT_SOLDOUT_NO
				End If
			Else
				getOptionLimitNo = 999
			End if
		End If
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		If (FItemOption = "0000") Then Exit Function
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO))
	End Function

	'// ��ǰ���: �ɼ� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getCJOptionParamToReg()
		Dim strSql, strRst, itemSu, itemoption, validSellno, optionname, fixday, optaddprice
		Dim GetTenTenMargin
		'2013-07-24 ������//���ٸ����� CJMALL�� �������� ���� �� orgprice�� ���� ����
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-24 ������//���ٸ����� CJMALL�� �������� ���� �� orgprice�� ���� ��

		optaddprice		= 0
		strSql = ""
		strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, i.deliverfixday, o.optaddprice " & VBCRLF
		strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " LEFT JOIN db_AppWish.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
		rsCTget.Open strSql, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			For i = 1 to rsCTget.RecordCount
				If rsCTget.RecordCount = 1 AND IsNull(rsCTget("itemoption")) Then  ''���ϻ�ǰ
					FItemOption = "0000"
					optionname = DdotFormat(chrbyte(getItemNameFormat,40,""),20)
					itemSu = GetCJLmtQty
					optaddprice		= 0
				Else
					FItemOption 	= rsCTget("itemoption")
					optionname 		= rsCTget("optionname")
					Foptsellyn 		= rsCTget("optsellyn")
					Foptlimityn 	= rsCTget("optlimityn")
					Foptlimitno 	= rsCTget("optlimitno")
					Foptlimitsold 	= rsCTget("optlimitsold")
					optaddprice		= rsCTget("optaddprice")
					itemSu = getOptionLimitNo

					if rsCTget("optnmLen")>40 then
					    optionname=DdotFormat(optionname,20)
					end if
				End If

				If rsCTget("deliverfixday") = "C" OR rsCTget("deliverfixday") = "X" Then
					fixday = "60"
				Else
					fixday = "20"
				End If


				strRst = strRst &"	<tns:unit>"
				''strRst = strRst &"		<tns:unitNm><![CDATA["&DDotFormat(optionname, 16)&"]]></tns:unitNm>"	'��ǰ���� - ��ǰ��(�ɼǸ��� �ؽ�Ʈ�� �ѱ�� ��)
				strRst = strRst &"		<tns:unitNm><![CDATA["&optionname&"]]></tns:unitNm>"
				strRst = strRst &"		<tns:unitRetail>"&FSellCash+optaddprice&"</tns:unitRetail>"				'��ǰ���� - �ǸŰ�
				strRst = strRst &"		<tns:unitCost>"&getCJmallSuplyPrice(optaddprice)&"</tns:unitCost>"					'��ǰ���� - ���Կ���
				strRst = strRst &"		<tns:availableQty>"&itemSu&"</tns:availableQty>"						'��ǰ���� - ���ް��ɼ��� (��ǰ ��� �ľ��� �ȵǴ°��� 999���� ���ڸ� �ֽ��ϴ�.)
			If getzCostomMadeInd = "Y" Then
				strRst = strRst &"		<tns:leadTime>"&getzLeadTime()&"</tns:leadTime>"						'��ǰ���� - ����Ÿ�� (* �����Ÿ�� ���ܵ�� �ؾ� ��, ����Ÿ�� ���� ���ǿ��� 00 : ���ù��, 01 : �������, 02 : �������, 03 : 2�������, 04 : 4��, 05 : 5��, 06 : 6��.....)
'			ElseIf Left(FCddkey,2) = "35" OR Left(FCddkey,2) = "37" Then											'��ǰ��Ͻ� ��з���(35 ��������/37 �������)�ϰ�� ����Ÿ���� ���� '02' ��ϸ� �����ϵ��� ó���Ǿ��ֽ��ϴ�.
'				strRst = strRst &"		<tns:leadTime>02</tns:leadTime>"										'��ǰ���� - ����Ÿ�� (* �����Ÿ�� ���ܵ�� �ؾ� ��, ����Ÿ�� ���� ���ǿ��� 00 : ���ù��, 01 : �������, 02 : �������, 03 : 2�������, 04 : 4��, 05 : 5��, 06 : 6��.....)
			Else
				strRst = strRst &"		<tns:leadTime>03</tns:leadTime>"										'��ǰ���� - ����Ÿ�� (* �����Ÿ�� ���ܵ�� �ؾ� ��, ����Ÿ�� ���� ���ǿ��� 00 : ���ù��, 01 : �������, 02 : �������, 03 : 2�������, 04 : 4��, 05 : 5��, 06 : 6��.....)
			End If
				strRst = strRst &"		<tns:unitApplyRsn>"&fixday&"</tns:unitApplyRsn>"						'��ǰ���� - ������� (10 : �������, 20 : ��ǰ����, 30 : ��ǰ����, 40 : �԰�˻�, 50 : ���˻�, 60 : ��ġ��ǰ)
				strRst = strRst &"		<tns:startSaleDt>"&FormatDate(now(), "0000-00-00")&"</tns:startSaleDt>"	'��ǰ���� - �ǸŽ�������
				strRst = strRst &"		<tns:endSaleDt>9999-12-30</tns:endSaleDt>"								'��ǰ���� - �Ǹ��������� (�ǸŻ��¼�������..)
			If Fitemid = "899506" Then
				strRst = strRst &"		<tns:vpn>"&rsCTget("itemid")&"_Q"&FItemOption&"</tns:vpn>"				'��ǰ���� - ���»��ǰ�ڵ�(899506�� Q��� ���ڻ���)
			Else
				strRst = strRst &"		<tns:vpn>"&rsCTget("itemid")&"_"&FItemOption&"</tns:vpn>"					'��ǰ���� - ���»��ǰ�ڵ�
			End If
				strRst = strRst &"	</tns:unit>"
				rsCTget.MoveNext
			Next
		End If
		rsCTget.Close
		getCJOptionParamToReg = strRst
	End Function

	'// ��ǰ���¼����� �ɼ��� �߰��� ���
	Public Function getCJOptionParamToEdit()
		Dim strSql, strRst, itemSu, itemoption, validSellno, optionname, fixday, optaddprice
		Dim GetTenTenMargin
		'2013-07-24 ������//���ٸ����� CJMALL�� �������� ���� �� orgprice�� ���� ����
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-24 ������//���ٸ����� CJMALL�� �������� ���� �� orgprice�� ���� ��

		optaddprice = 0
		strSql = ""
		strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, isnull(R.outmallOptCode, '') as outmallOptCode, i.deliverfixday, isnull(o.optaddprice,'') as optaddprice " & VBCRLF
		strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " JOIN db_AppWish.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF ''LEFT Join => Join
		strSql = strSql & " LEFT JOIN [db_AppWish].[dbo].tbl_OutMall_regedoption as R on i.itemid = R.itemid and R.itemoption = o.itemoption and R.mallid='cjmall' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
		rsCTget.Open strSql, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			For i = 1 to rsCTget.RecordCount
				If rsCTget("outmallOptCode") = "" Then
					itemSu = getOptionLimitNo
					FItemOption 	= rsCTget("itemoption")
					optionname 		= rsCTget("optionname")
					Foptsellyn 		= rsCTget("optsellyn")
					Foptlimityn 	= rsCTget("optlimityn")
					Foptlimitno 	= rsCTget("optlimitno")
					Foptlimitsold 	= rsCTget("optlimitsold")
					optaddprice		= rsCTget("optaddprice")
					If rsCTget("deliverfixday") = "C" OR rsCTget("deliverfixday") = "X" Then
						fixday = "60"
					Else
						fixday = "20"
					End If

                    if rsCTget("optnmLen")>40 then
					    optionname=DdotFormat(optionname,20)
					end if

					If itemSu <> 0 Then
						strRst = strRst &"	<tns:unit>"
						strRst = strRst &"		<tns:unitNm><![CDATA["&optionname&"]]></tns:unitNm>"					'��ǰ���� - ��ǰ��(�ɼǸ��� �ؽ�Ʈ�� �ѱ�� ��)
						strRst = strRst &"		<tns:unitRetail>"&FSellCash+optaddprice&"</tns:unitRetail>"							'��ǰ���� - �ǸŰ�
						strRst = strRst &"		<tns:unitCost>"&getCJmallSuplyPrice(optaddprice)&"</tns:unitCost>"					'��ǰ���� - ���Կ���
						strRst = strRst &"		<tns:availableQty>"&itemSu&"</tns:availableQty>"						'��ǰ���� - ���ް��ɼ��� (��ǰ ��� �ľ��� �ȵǴ°��� 999���� ���ڸ� �ֽ��ϴ�.)
						If getzCostomMadeInd = "Y" Then
							strRst = strRst &"		<tns:leadTime>"&getzLeadTime()&"</tns:leadTime>"						'��ǰ���� - ����Ÿ�� (* �����Ÿ�� ���ܵ�� �ؾ� ��, ����Ÿ�� ���� ���ǿ��� 00 : ���ù��, 01 : �������, 02 : �������, 03 : 2�������, 04 : 4��, 05 : 5��, 06 : 6��.....)
	        '			ElseIf Left(FCddkey,2) = "35" OR Left(FCddkey,2) = "37" Then											'��ǰ��Ͻ� ��з���(35 ��������/37 �������)�ϰ�� ����Ÿ���� ���� '02' ��ϸ� �����ϵ��� ó���Ǿ��ֽ��ϴ�.
	        '				strRst = strRst &"		<tns:leadTime>02</tns:leadTime>"										'��ǰ���� - ����Ÿ�� (* �����Ÿ�� ���ܵ�� �ؾ� ��, ����Ÿ�� ���� ���ǿ��� 00 : ���ù��, 01 : �������, 02 : �������, 03 : 2�������, 04 : 4��, 05 : 5��, 06 : 6��.....)
	        			Else
	        				strRst = strRst &"		<tns:leadTime>03</tns:leadTime>"										'��ǰ���� - ����Ÿ�� (* �����Ÿ�� ���ܵ�� �ؾ� ��, ����Ÿ�� ���� ���ǿ��� 00 : ���ù��, 01 : �������, 02 : �������, 03 : 2�������, 04 : 4��, 05 : 5��, 06 : 6��.....)
	        			End If
						strRst = strRst &"		<tns:unitApplyRsn>"&fixday&"</tns:unitApplyRsn>"						'��ǰ���� - ������� (10 : �������, 20 : ��ǰ����, 30 : ��ǰ����, 40 : �԰�˻�, 50 : ���˻�, 60 : ��ġ��ǰ)
						strRst = strRst &"		<tns:startSaleDt>"&FormatDate(now(), "0000-00-00")&"</tns:startSaleDt>"	'��ǰ���� - �ǸŽ�������
						strRst = strRst &"		<tns:endSaleDt>9999-12-30</tns:endSaleDt>"								'��ǰ���� - �Ǹ��������� (�ǸŻ��¼�������..)
						strRst = strRst &"		<tns:vpn>"&rsCTget("itemid")&"_"&FItemOption&"</tns:vpn>"					'��ǰ���� - ���»��ǰ�ڵ�
						strRst = strRst &"	</tns:unit>"
					End If
				End If
				rsCTget.MoveNext
			Next
		End If
		rsCTget.Close
		getCJOptionParamToEdit = strRst
	End Function


	'// ��ǰ���: ��ǰ�߰��̹��� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getCJAddImageParamToReg()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
		End If

		strRst = strRst &"	<tns:image>"
		strRst = strRst &"		<tns:imageMain>"&FbasicImage&"</tns:imageMain>"
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_AppWish].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsCTget.CursorLocation = adUseClient
		rsCTget.CursorType=adOpenStatic
		rsCTget.Locktype=adLockReadOnly
		rsCTget.Open strSQL, dbCTget

		If Not(rsCTget.EOF or rsCTget.BOF) Then
			For i=1 to rsCTget.RecordCount
				If rsCTget("imgType")="0" Then
					strRst = strRst &"		<tns:imageSub"&i&">http://webimage.10x10.co.kr/image/add" & rsCTget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsCTget("addimage_400") &"</tns:imageSub"&i&">"
				End If
				rsCTget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsCTget.Close
		strRst = strRst &"	</tns:image>"
		getCJAddImageParamToReg = strRst
	End Function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getCJItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		'2014-01-17 10:00 ������ ž �̹��� �߰�
		strRst = strRst & ("<p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>")
		'#�⺻ ��ǰ����
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & (nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
		End Select
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_AppWish].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsCTget.CursorLocation = adUseClient
		rsCTget.CursorType=adOpenStatic
		rsCTget.Locktype=adLockReadOnly
		rsCTget.Open strSQL, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			Do Until rsCTget.EOF
				If rsCTget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsCTget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
				End If
				rsCTget.MoveNext
			Loop
		End If
		rsCTget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#��� ���ǻ���
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg"">")

		strRst = strRst & ("</div>")
		getCJItemContParamToReg = strRst
		''2013-06-10 ������ �߰�(�Ե�����ó�� ��ǰ�̹����� ��� ���ڳ����� ����)
		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_outMall.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','cjmall') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF  '' mallid='cjmall' => mallid in ('','cjmall')
		rsCTget.Open strSQL, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			strRst = rsCTget("textVal")
			strRst = "<div align=""center""><p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>" & strRst & "<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg""></div>"
			getCJItemContParamToReg = strRst
		End If
		rsCTget.Close
	End Function

	Public Function getCjmallXMLTEST()
		Dim tXML
		tXML = ""
		tXML = tXML &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		tXML = tXML &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_01' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_01.xsd'>"
		tXML = tXML &"<tns:vendorId>411378</tns:vendorId>"
		tXML = tXML &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
		tXML = tXML &"<tns:good>"
		tXML = tXML &"	<tns:chnCls>30</tns:chnCls>"
		tXML = tXML &"	<tns:tGrpCd>50010104</tns:tGrpCd>"
		tXML = tXML &"	<tns:uniqBrandCd>24049000</tns:uniqBrandCd>"
		tXML = tXML &"	<tns:giftInd>Y</tns:giftInd>"
		tXML = tXML &"	<tns:uniqMkrNatCd>901</tns:uniqMkrNatCd>"
		tXML = tXML &"	<tns:uniqMkrCompCd>54498</tns:uniqMkrCompCd>"
		tXML = tXML &"	<tns:itemDesc>Bonnie Stripe Pony  </tns:itemDesc>"
		tXML = tXML &"	<tns:zLocalBolDesc>Bonnie Stripe Pony  </tns:zLocalBolDesc>"
		tXML = tXML &"	<tns:zlocalCcDesc>Bonnie Str</tns:zlocalCcDesc>"
		tXML = tXML &"	<tns:vatCode>S</tns:vatCode>"
		tXML = tXML &"	<tns:zDeliveryType>20</tns:zDeliveryType>"
		tXML = tXML &"	<tns:zShippingMethod>10</tns:zShippingMethod>"
		tXML = tXML &"	<tns:courier>11</tns:courier>"
		tXML = tXML &"	<tns:deliveryHomeCost>2500</tns:deliveryHomeCost>"
		tXML = tXML &"	<tns:zreturnNotReqInd>10</tns:zreturnNotReqInd>"
		tXML = tXML &"	<tns:zCostomMadeInd>N</tns:zCostomMadeInd>"
		tXML = tXML &"	<tns:stockMgntLevel>2</tns:stockMgntLevel>"
		tXML = tXML &"	<tns:lowpriceInd>N</tns:lowpriceInd>"
		tXML = tXML &"	<tns:delayShipRewardIind>N</tns:delayShipRewardIind>"
		tXML = tXML &"	<tns:reserveDayInd>Y</tns:reserveDayInd>"
		tXML = tXML &"	<tns:zContactSeqNo>10003</tns:zContactSeqNo>"
		tXML = tXML &"	<tns:zSupShipSeqNo>10002</tns:zSupShipSeqNo>"
		tXML = tXML &"	<tns:zReturnSeqNo>10002</tns:zReturnSeqNo>"
		tXML = tXML &"	<tns:zAsSupShipSeqNo>10002</tns:zAsSupShipSeqNo>"
		tXML = tXML &"	<tns:zAsReturnSeqNo>10002</tns:zAsReturnSeqNo>"
		tXML = tXML &"	<tns:unit>"
		tXML = tXML &"		<tns:unitNm>MULTICOLOR - ONESIZE</tns:unitNm>"
		tXML = tXML &"		<tns:unitRetail>92500</tns:unitRetail>"
		tXML = tXML &"		<tns:unitCost>70775</tns:unitCost>"
		tXML = tXML &"		<tns:availableQty>3000</tns:availableQty>"
		tXML = tXML &"		<tns:leadTime>03</tns:leadTime>"
		tXML = tXML &"		<tns:unitApplyRsn>20</tns:unitApplyRsn>"
		tXML = tXML &"		<tns:startSaleDt>2013-03-22</tns:startSaleDt>"
		tXML = tXML &"		<tns:endSaleDt>9999-12-30</tns:endSaleDt>"
		tXML = tXML &"		<tns:vpn>11111111</tns:vpn>"
		tXML = tXML &"	</tns:unit>"
		tXML = tXML &"	<tns:mallitem>"
		tXML = tXML &"		<tns:mallItemDesc>[Coach]Bonnie Stripe Pony  </tns:mallItemDesc>"
		tXML = tXML &"		<tns:keyword>Coach</tns:keyword>"
		tXML = tXML &"		<tns:mallCtg>"
		tXML = tXML &"			<tns:mainInd>Y</tns:mainInd>"
		tXML = tXML &"			<tns:ctgName>155376</tns:ctgName>"
		tXML = tXML &"		</tns:mallCtg>"
		tXML = tXML &"		<tns:mallCtg>"
		tXML = tXML &"			<tns:ctgName>106252</tns:ctgName>"
		tXML = tXML &"		</tns:mallCtg>"
		tXML = tXML &"	</tns:mallitem>"
		tXML = tXML &"  <tns:goodsReport>"
		tXML = tXML &"		<tns:pedfId>25004</tns:pedfId>"
		tXML = tXML &"		<tns:html>"
		tXML = tXML &"			��???1"
		tXML = tXML &"		</tns:html>"
		tXML = tXML &"	</tns:goodsReport>"
		tXML = tXML &"  <tns:goodsReport>"
		tXML = tXML &"		<tns:pedfId>25005</tns:pedfId>"
		tXML = tXML &"		<tns:html>"
		tXML = tXML &"			��???1"
		tXML = tXML &"		</tns:html>"
		tXML = tXML &"	</tns:goodsReport>"
		tXML = tXML &"  <tns:goodsReport>"
		tXML = tXML &"		<tns:pedfId>25024</tns:pedfId>"
		tXML = tXML &"		<tns:html>"
		tXML = tXML &"			��???1"
		tXML = tXML &"		</tns:html>"
		tXML = tXML &"	</tns:goodsReport>"
		tXML = tXML &"  <tns:goodsReport>"
		tXML = tXML &"		<tns:pedfId>25154</tns:pedfId>"
		tXML = tXML &"		<tns:html>"
		tXML = tXML &"			��???1"
		tXML = tXML &"		</tns:html>"
		tXML = tXML &"	</tns:goodsReport>"
		tXML = tXML &"  <tns:goodsReport>"
		tXML = tXML &"		<tns:pedfId>25155</tns:pedfId>"
		tXML = tXML &"		<tns:html>"
		tXML = tXML &"			��???1"
		tXML = tXML &"		</tns:html>"
		tXML = tXML &"	</tns:goodsReport>"
		tXML = tXML &"	<tns:goodsReport>"
		tXML = tXML &"		<tns:pedfId>91059</tns:pedfId>"
		tXML = tXML &"		<tns:html>"
		tXML = tXML &"			<![CDATA[ <script language='javascript' src='http://image.cjmall.com/common/jsCommon.js'></script>"
		tXML = tXML &"			<div style='text-align:center;width:738px'>"
		tXML = tXML &"			<img src='http://image.cjmall.com/prd/new2008/njoyny_notice_west.jpg' border='0' align='absmiddle'  usemap='#NjoyNY_Map1'>"
		tXML = tXML &"			</div>"
		tXML = tXML &"			<map name='NjoyNY_Map1'>"
		tXML = tXML &"			<area shape='rect' coords='10,280,155,330' href='javascript:win_pop('http://www.cjmall.com/prd/popup/NjoyNY_pop1.jsp','','690','570','auto');' onFocus='this.blur()'>"
		tXML = tXML &"			</map>"
		tXML = tXML &"			<br/><br/>"
		tXML = tXML &"			[Coach] Bonnie Stripe Pony [98589]<br><br><li> Signature Coach Op Art and legacy stripe prints on imported silk<li> 2.1/2 (W) x 35 (L)<br><table width='600'><tr><td><img src='http://img.buynjoy.com/images_2009_1/automd/coach/200908/1250146500_main_300.jpg'></td><td></td></tr></table>]]>"
		tXML = tXML &"		</tns:html>"
		tXML = tXML &"	</tns:goodsReport>"
		tXML = tXML &"	<tns:image>"
		tXML = tXML &"		<tns:imageMain>http://img.buynjoy.com/images_2009_1/automd/coach/200908/1250146500_main_550.jpg</tns:imageMain>"
		tXML = tXML &"		<tns:imageSub1>http://img.buynjoy.com/images_2009_1/automd/coach/200908/1250146500_main_550.jpg</tns:imageSub1>"
		tXML = tXML &"	</tns:image>"
		tXML = tXML &"</tns:good>"
		tXML = tXML &"</tns:ifRequest>"
		getCjmallXMLTEST = tXML
	End Function

	Public Function getCjmallItemRegXML
		Dim strRst
		Dim ioriginCode, ioriginname
		Dim makercompCode, makercompName
		ioriginCode 	= getOriginName2Code(Fsourcearea, ioriginname) 		'�������ڵ�
		makercompCode	= getmakerName2Code(Fsocname_kor, makercompName)	'�������ڵ�
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_01' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_01.xsd'>"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"									'!!!���¾�ü�ڵ�
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"					'!!!����Ű
		strRst = strRst &"<tns:good>"
		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"												'!!!��ǰ�з�ü�� - �����ä�α���(30:���ͳ�, 40:īŻ�α�)
		strRst = strRst &"	<tns:tGrpCd>"&FCddKey&"</tns:tGrpCd>"										'!!!��ǰ�з�ü�� - ��ǰ�з�
		strRst = strRst &"	<tns:uniqBrandCd>"&CUNIQBRANDCD&"</tns:uniqBrandCd>"						'!!!��ǰ�з�ü�� - �귣��(�ٹ�����:24049000)
		strRst = strRst &"	<tns:giftInd>Y</tns:giftInd>"											    '!!!��ǰ�з�ü�� - ��ǰ���� (Y=�Ϲ��ǸŻ�ǰ, N=����ǰ)
		strRst = strRst &"	<tns:uniqMkrNatCd>"&ioriginCode&"</tns:uniqMkrNatCd>"						'!!!��ǰ�з�ü�� - ������
		strRst = strRst &"	<tns:uniqMkrCompCd>"&makercompCode&"</tns:uniqMkrCompCd>"					'!!!��ǰ�з�ü�� - ������
'		strRst = strRst &"	<tns:ingredient></tns:ingredient>"											'��ǰ�з�ü�� - �ֿ����	(���� ���������� ����)
'		strRst = strRst &"	<tns:zingredientOrigin></tns:zingredientOrigin>"							'��ǰ�з�ü�� - ���������	(���� ���������� ����) // ��ǰ�з�(��з�)�� ��ǰ�϶��� ������ �ʼ�(�����..;;)
	If Fitemid = "899506" Then
		strRst = strRst &"	<tns:mdCode>5066</tns:mdCode>"
	Else
		strRst = strRst &"	<tns:mdCode>"&MD_CODE&"</tns:mdCode>"										'!!!MD�ڵ�						(�ִ� ���õ� �ְ�, ������ ���õ� ����) ���ƾ� ���� (�ٹ����� ���� ����)
	End If
		strRst = strRst &"	<tns:itemDesc><![CDATA["&DDotFormat(getItemNameFormat, 100)&"]]></tns:itemDesc>"			'!!!�⺻���� - ��ǰ��(120�� ����) (���ÿ� CDATA������ �߰�)
		strRst = strRst &"	<tns:zLocalBolDesc><![CDATA["&DDotFormat(getItemNameFormat, 10)&"]]></tns:zLocalBolDesc>"	'!!!�⺻���� - ������(40�� ����)
		strRst = strRst &"	<tns:zlocalCcDesc><![CDATA["&DDotFormat(getItemNameFormat, 5)&"]]></tns:zlocalCcDesc>"		'!!!�⺻���� - SMS��ǰ��(20�� ����)
		strRst = strRst &"	<tns:vatCode>"&CHKIIF(FVatInclude="N","E","S")&"</tns:vatCode>"			 	'!!!�⺻���� - �������� (S:����, E:�鼼, N:�����, Z:����)
		strRst = strRst &"	<tns:zDeliveryType>20</tns:zDeliveryType>"									'!!!�⺻���� - ��۱��� (10:���͹��, 20:���»���, 30:���ù�, 35:���ù襱, 40:����, 99:��۾���)
		strRst = strRst &"	<tns:zShippingMethod>"&getdeliverfixday&"</tns:zShippingMethod>"			'!!!�⺻���� - ������� (10:�ù���, 20:��ġ��ǰ, 30:��޼���, 40:����/�����) ''ȭ����� Ȯ��
		strRst = strRst &"	<tns:courier>22</tns:courier>"												'!!!�⺻���� - �ù�� (�����ù�� �ϳ� ���� �� ������ ���)(11:�����ù�, 12:�������, 15:�����ù�, 22:CJGLS, 29:CJHTH, 87:�����ͽ�������) CJ�ù� �ڵ�� ���
		strRst = strRst &"	<tns:deliveryHomeCost>2500</tns:deliveryHomeCost>"							'�⺻���� - ��ۺ� (��۱����� ���»���, ������ ��� �ʼ� �Է�)
		strRst = strRst &"	<tns:zreturnNotReqInd>10</tns:zreturnNotReqInd>"							'�⺻���� - ȸ������ (��۱��п� ���� �ʼ�/�ɼ�)
'		strRst = strRst &"	<tns:zJointPackingQty></tns:zJointPackingQty>"								'�⺻���� - ��������� (��۱��п� ���� �ʼ�/�ɼ�) (�������������� ����)
		strRst = strRst &"	<tns:zCostomMadeInd>"&getzCostomMadeInd()&"</tns:zCostomMadeInd>"			'!!!�⺻���� - �ֹ����ۿ��� (Y=�ֹ�����, N=�ֹ����۾���)) ''' �ֹ����ۻ�ǰ, �ֹ������ۻ�ǰ =>Y
		strRst = strRst &"	<tns:stockMgntLevel>2</tns:stockMgntLevel>"									'�⺻���� - ���������� (1=�Ǹ��ڵ�,2=��ǰ�ڵ�)
'		strRst = strRst &"	<tns:leadtime></tns:leadtime>"												'�⺻���� - ����Ÿ�� (1. �����ڴ� NULL���� 2.������������ "�Ǹ��ڵ�"�϶� �ʼ�) (�������������� ����)
'		strRst = strRst &"	<tns:leadtimeChgRsn></tns:leadtimeChgRsn>"									'�⺻���� - ������� (1. �����ڴ� NULL���� 2.������������ "�Ǹ��ڵ�"�϶� �ʼ�) (�������������� ����)
		strRst = strRst &"	<tns:lowpriceInd>"&CHKIIF(IsCjFreeBeasong=False,"Y","N")&"</tns:lowpriceInd>"	'!!!�⺻���� - �����ۿ��� (Y=������,N=������)        '' Ȯ��.
		strRst = strRst &"	<tns:delayShipRewardIind>N</tns:delayShipRewardIind>"						'�⺻���� - �������󿩺� (Y=��������,N=�����������)
'		strRst = strRst &"	<tns:packingMethod></tns:packingMethod>"									'�⺻���� - �԰����� (���͹���� ��츸 �Է�)
'		strRst = strRst &"	<tns:zOrderMaxQty></tns:zOrderMaxQty>"										'�⺻���� - 1ȸ�ִ��ֹ����� (���� 1ȸ �ִ� �ֹ����� ����. ���Է½� ���Ѿ���
'		strRst = strRst &"	<tns:zDayOrderMaxQty></tns:zDayOrderMaxQty>"								'�⺻���� - 1���ִ��ֹ����� (���� ���� �ִ� �ֹ����� ����. ���Է½� ���Ѿ���)
		strRst = strRst &"	<tns:reserveDayInd>Y</tns:reserveDayInd>"									'�⺻���� - �����۹�� (* ����Ʈ: YN-�ֹ���� �������� Y-���ʰ��ް����� ��������_Default)
		strRst = strRst &"	<tns:zContactSeqNo>"&chkiif(application("Svr_Info")="Dev","10003","10002")&"</tns:zContactSeqNo>"		'�⺻���� - ���»�����
		strRst = strRst &"	<tns:zSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","23125")&"</tns:zSupShipSeqNo>"		'�⺻���� - ������
		strRst = strRst &"	<tns:zReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","23125")&"</tns:zReturnSeqNo>"			'�⺻���� - ȸ����
		strRst = strRst &"	<tns:zAsSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","23125")&"</tns:zAsSupShipSeqNo>"	'�⺻���� - AS������
		strRst = strRst &"	<tns:zAsReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","23125")&"</tns:zAsReturnSeqNo>"		'�⺻���� - ASȸ����
		strRst = strRst & getCJOptionParamToReg															'��ǰ����
		strRst = strRst &"	<tns:mallitem>"
		strRst = strRst &"		<tns:mallItemDesc><![CDATA["&"�ٹ����� " & Fsocname_kor & " "&DDotFormat(getItemNameFormat, 186)&"]]></tns:mallItemDesc>"	'!!!CJmall��ǰ���� - CJmall��ǰ�� , �ٹ����� �귣��� �߰�
		strRst = strRst &"		<tns:keyword><![CDATA["&"�ٹ�����;"&replace(Fkeywords,",",";")&"]]></tns:keyword>"						'!!!CJmall��ǰ���� - �˻�Ű����
		strRst = strRst & getCjCateParamToReg															'!!!����ī�װ�����(Y=ī�װ�,N=ī�װ��ƴ�) // CJmallī�װ�(��)
		strRst = strRst &"	</tns:mallitem>"
'		strRst = strRst &"	<tns:cert>"																			'QC���� �ذ��Ϸ��� �Ʒ������� �ʿ��ѵ�..(2013-06-04 ������)
'		strRst = strRst &"		<tns:certCode>350504</tns:certCode>"											'ǰ���������� - �׸��ڵ�
'		strRst = strRst &"		<tns:certNo>YU11100-12001</tns:certNo>"											'ǰ���������� - ������ȣ - ��������(50)
'		strRst = strRst &"		<tns:issueDate>2012-06-04</tns:issueDate>"										'ǰ���������� - �߱�����
'		strRst = strRst &"		<tns:certDate>2012-06-05</tns:certDate>"         								'ǰ���������� - ��������
'		strRst = strRst &"		<tns:avlStartDate>2012-06-04</tns:avlStartDate>"								'ǰ���������� - ��ȿ�Ⱓ(FROM)
'		strRst = strRst &"		<tns:avlEndDate>2013-06-04</tns:avlEndDate>"      								'ǰ���������� - ��ȿ�Ⱓ(TO)
'		strRst = strRst &"		<tns:itemModel>item</tns:itemModel>"        									'ǰ���������� - ��ǰ�� �� �𵨸�	-��������(200)
'		strRst = strRst &"		<tns:orgCode>��������</tns:orgCode>"            								'ǰ���������� - �����˻�����		-��������(200)
'		strRst = strRst &"		<tns:certField>������ǰ</tns:certField>"        								'ǰ���������� - �����о�			-��������(200)
'		strRst = strRst &"		<tns:originCode>������</tns:originCode>"     									'ǰ���������� - ������(������)
'		strRst = strRst &"		<tns:certSpec>����</tns:certSpec>"          									'ǰ���������� - ���λ���			-��������(2000)
'		strRst = strRst &"	</tns:cert>"
		strRst = strRst & getCjmallItemInfoCdToReg()													'��ǰ�����
		strRst = strRst &"	<tns:goodsReport>"
		strRst = strRst &"		<tns:pedfId>91059</tns:pedfId>"
		strRst = strRst &"		<tns:html>"
		strRst = strRst &"			<![CDATA["&getCJItemContParamToReg&"]]>"
		strRst = strRst &"		</tns:html>"
		strRst = strRst &"	</tns:goodsReport>"
														'daebeak	����ǰ�߰����� ��������
		strRst = strRst & getCJAddImageParamToReg		'!!!�̹�������
		strRst = strRst &"</tns:good>"
		strRst = strRst &"</tns:ifRequest>"
		getCjmallItemRegXML = strRst
	End Function

	'// ��ǰ ���� ���� �Ķ���� ����
	Public Function getcjmallItemModXML(unit)
		Dim strRst
		Dim ioriginCode, ioriginname
		ioriginCode = getOriginName2Code(Fsourcearea, ioriginname)
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_02"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_02.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"												'!!!���¾�ü�ڵ�
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"								'!!!����Ű
		strRst = strRst &"<tns:good>"
		strRst = strRst &"	<tns:sItem>"&FcjmallPrdNo&"</tns:sItem>"												'!!!�ǸŻ�ǰ�ڵ�(Ȩ����)
	If Fitemid = "899506" Then
		strRst = strRst &"	<tns:loc>110</tns:loc>"																	'!!!��ǰ�з�ü�� - ���ä�α���(��������)
	Else
		strRst = strRst &"	<tns:loc>30</tns:loc>"																	'!!!��ǰ�з�ü�� - ���ä�α���(store����)
	End If
		strRst = strRst &"	<tns:zLocalBolDesc><![CDATA["&DDotFormat(getItemNameFormat, 10)&"]]></tns:zLocalBolDesc>"		'!!!�⺻���� - ������
		strRst = strRst &"	<tns:zlocalCcDesc><![CDATA["&DDotFormat(getItemNameFormat, 5)&"]]></tns:zlocalCcDesc>"			'!!!�⺻���� - SMS��ǰ��
		strRst = strRst &"	<tns:zContactSeqNo>"&chkiif(application("Svr_Info")="Dev","10003","10002")&"</tns:zContactSeqNo>"		'!!!�⺻���� - ���»�����
		strRst = strRst &"	<tns:zSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","23125")&"</tns:zSupShipSeqNo>"		'!!!�⺻���� - ������
		strRst = strRst &"	<tns:zReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","23125")&"</tns:zReturnSeqNo>"			'!!!�⺻���� - ȸ����
		strRst = strRst &"	<tns:zAsSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","23125")&"</tns:zAsSupShipSeqNo>"	'!!!�⺻���� - AS������
		strRst = strRst &"	<tns:zAsReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","23125")&"</tns:zAsReturnSeqNo>"		'!!!�⺻���� - ASȸ����
'		strRst = strRst &"	<tns:zJointPackingQty>10008</tns:zJointPackingQty>"										'�⺻���� - ���������(���ÿ� ����)
'		strRst = strRst &"	<tns:lowpriceInd>10008</tns:lowpriceInd>"												'�⺻���� - �����ۿ���(���ÿ� ����)
        strRst = strRst &"	<tns:lowpriceInd>"&CHKIIF(IsCjFreeBeasong=False,"Y","N")&"</tns:lowpriceInd>"	'!!!�⺻���� - �����ۿ��� (Y=������,N=������)        '' Ȯ��.

		strRst = strRst & getCJOptionParamToEdit                                                                      '' Ȯ���� ���� ''864806

		strRst = strRst &"	<tns:mallitem>"
		strRst = strRst &"		<tns:mallItemDesc><![CDATA["&"�ٹ����� " & Fsocname_kor & " "&DDotFormat(getItemNameFormat, 186)&"]]></tns:mallItemDesc>"	'!!!CJmall��ǰ���� - CJmall��ǰ��
		strRst = strRst &"	</tns:mallitem>"
		strRst = strRst & getCjmallItemInfoCdToReg()													'��ǰ�����
		strRst = strRst &"	<tns:goodsReport>"
		strRst = strRst &"		<tns:pedfId>91059</tns:pedfId>"
		strRst = strRst &"		<tns:html>"
		strRst = strRst &"			<![CDATA["&getCJItemContParamToReg&"]]>"
		strRst = strRst &"		</tns:html>"
		strRst = strRst &"	</tns:goodsReport>"
		strRst = strRst & getCJAddImageParamToReg		'!!!�̹�������
		strRst = strRst &"</tns:good>"
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemModXML = strRst
	End Function

	'// ��ǰ ���� ���� �Ķ���� ����
	Public Function getcjmallItemPriceModXML(isPrdCode)
		Dim strRst, sqlStr, arrrows, chkOption, i, optAddPRcExists, GetTenTenMargin
		optAddPRcExists = false

		'2013-07-24 ������//���ٸ����� CJMALL�� �������� ���� �� orgprice�� ���� ����
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-24 ������//���ٸ����� CJMALL�� �������� ���� �� orgprice�� ���� ��

		sqlStr = ""
		sqlStr = sqlStr & " select distinct o.itemid, o.optAddPrice,  ro.outmallOptCode, o.itemoption"
		sqlStr = sqlStr & " from db_AppWish.dbo.tbl_item_option o "
		sqlStr = sqlStr & " Join [db_AppWish].[dbo].tbl_OutMall_regedoption ro on o.itemid=ro.itemid and ro.mallid ='cjmall' and ro.itemoption = o.itemoption "
		sqlStr = sqlStr & " where o.itemid = '"&Fitemid&"' "
		sqlStr = sqlStr & " group by o.itemid, o.optAddPrice, ro.outmallOptCode, o.itemoption"
		sqlStr = sqlStr & " order by o.optAddPrice, o.itemoption"
		rsCTget.Open sqlStr, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			arrrows = rsCTget.getRows
			chkOption = True
		Else
			chkOption = False
		End If
		rsCTget.close

        if (chkOption) then
            For i = 0 To UBound(ArrRows,2)
                optAddPRcExists = optAddPRcExists or (arrRows(1,i)>0)
    		Next
    	end if

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_04"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_04.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"						'!!!���¾�ü�ڵ�
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"		'!!!����Ű

        if (isPrdCode) or (Not optAddPRcExists) then  '' �հ� �� �̻��ϰ� ������ (�Ǹ��ڵ�/��ǰ�ڵ� �Ѵ� ������).. �ѹ������ǰ� �ѹ� �̻��ϰ� ����?
            strRst = strRst &"<tns:itemPrices>"
    		strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"									'01�̸� �Ǹ��ڵ� / 02�� ��ǰ�ڵ�
    		strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
    		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
    		strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
    		strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
    		strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
    		strRst = strRst &"</tns:itemPrices>"
    	ELSE
'            if (FItemID=813141) then  ''�ӽ� 2013/07/01
'                strRst = strRst &"<tns:itemPrices>"
'        		strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"									'01�̸� �Ǹ��ڵ� / 02�� ��ǰ�ڵ�
'        		strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
'        		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
'        		strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
'        		strRst = strRst &"	<tns:newUnitRetail>"&FSellCash&"</tns:newUnitRetail>"
'        		strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
'        		strRst = strRst &"</tns:itemPrices>"
'            end if
    		If chkOption = True Then
    			For i = 0 To UBound(ArrRows,2)
    				strRst = strRst &"<tns:itemPrices>"
    				strRst = strRst &"	<tns:typeCD>02</tns:typeCD>"									'01�̸� �Ǹ��ڵ� / 02�� ��ǰ�ڵ�
    				strRst = strRst &"	<tns:itemCD_ZIP>"&arrRows(2,i)&"</tns:itemCD_ZIP>"
    				strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
    				strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
    				strRst = strRst &"	<tns:newUnitRetail>"&MustPrice+arrRows(1,i)&"</tns:newUnitRetail>"
    				strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice(arrRows(1,i))&"</tns:newUnitCost>"
    				strRst = strRst &"</tns:itemPrices>"
    				optAddPRcExists = optAddPRcExists or (arrRows(1,i)>0)
    			Next
    		End If
    	ENd IF
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemPriceModXML = strRst
	End Function

	Function getcjmallItemSellPriceModXML()
		Dim strRst, sqlStr, i, GetTenTenMargin
		'2013-07-24 ������//���ٸ����� CJMALL�� �������� ���� �� orgprice�� ���� ����
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-24 ������//���ٸ����� CJMALL�� �������� ���� �� orgprice�� ���� ��

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_04"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_04.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"						'!!!���¾�ü�ڵ�
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"		'!!!����Ű
        strRst = strRst &"<tns:itemPrices>"
		strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"									'01�̸� �Ǹ��ڵ� / 02�� ��ǰ�ڵ�
		strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
		strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
		strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
		strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
		strRst = strRst &"</tns:itemPrices>"
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemSellPriceModXML = strRst
	End Function

	'// ��ǰ ���� ���� �Ķ���� ����
	Public Function getcjmallItemQTYXML
		Dim sqlStr, oneOpt, j
		Dim arrRows, i, strRst, validSellno
		sqlStr = ""
		sqlStr = sqlStr & " select isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName "
		sqlStr = sqlStr & " from [db_AppWish].[dbo].tbl_OutMall_regedoption as r "
		sqlStr = sqlStr & " left join [db_AppWish].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
		sqlStr = sqlStr & " where r.mallid = 'cjmall' and r.itemid="&Fitemid
		rsCTget.Open sqlStr, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			oneOpt = rsCTget.getRows
		End If
		rsCTget.close

		If (UBound(oneOpt,2) = "0") and (oneOpt(2,0) = "���ϻ�ǰ") Then
			strRst = ""
			strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
			strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_05"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_05.xsd"">"
			strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"						'!!!���¾�ü�ڵ�
			strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"		'!!!����Ű
			strRst = strRst &"<tns:ltSupplyPlans>"
			strRst = strRst &"	<tns:unitCd>"&oneOpt(1,0)&"</tns:unitCd>"
			strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"	<tns:strDt>"&FormatDate(now(), "0000-00-00")&"</tns:strDt>"
			If GetCJLmtQty = 0 Then
				strRst = strRst &"	<tns:endDt>"&FormatDate(now(), "0000-00-00")&"</tns:endDt>"
			Else
				strRst = strRst &"	<tns:endDt>9999-12-30</tns:endDt>"
			End If
			strRst = strRst &"	<tns:availSupQty>"&chkiif(GetCJLmtQty=0,"1",GetCJLmtQty)&"</tns:availSupQty>"
			strRst = strRst &"</tns:ltSupplyPlans>"
			strRst = strRst &"</tns:ifRequest>"
		Else
			sqlStr = ""
			sqlStr = sqlStr & " SELECT o.itemoption, o.optionTypeName, o.optionname, isnull(R.outmallOptCode, '') as outmallOptCode, (o.optlimitno-o.optlimitsold) as optLimit, o.optlimityn, o.isUsing, o.optsellyn " & VBCRLF
			sqlStr = sqlStr & " FROM [db_AppWish].[dbo].tbl_item_option o " & VBCRLF
			sqlStr = sqlStr & " left join [db_AppWish].[dbo].tbl_OutMall_regedoption R on o.itemid=R.itemid and o.itemoption=R.itemoption and R.mallid='cjmall' " & VBCRLF
			sqlStr = sqlStr & " where R.outmallOptCode <> '' and o.itemid="&Fitemid
			rsCTget.Open sqlStr, dbCTget
			If Not(rsCTget.EOF or rsCTget.BOF) Then
				arrRows = rsCTget.getRows
			End If
			rsCTget.close

			If isArray(arrRows) Then
				strRst = ""
				strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
				strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_05"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_05.xsd"">"
				strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"						'!!!���¾�ü�ڵ�
				strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"		'!!!����Ű
				For i = 0 To UBound(ArrRows,2)
					validSellno = 50				'�ִ� 50���� ��������
					If (FSellyn <> "Y") or ((arrRows(5,i) = "Y") and (arrRows(4,i) < 1)) or (arrRows(6,i) <> "Y") or (arrRows(7,i) <> "Y") Then
						validSellno = 0
					End If

					If (arrRows(5,i) = "Y") Then
						validSellno = arrRows(4,i)
					End If

					If (validSellno < CMAXLIMITSELL) Then validSellno = 0
					If (arrRows(5,i) = "Y") and (validSellno > 0) Then
						validSellno = validSellno - CMAXLIMITSELL
					End If
					If (validSellno < 1) then validSellno = 0
					If IsSoldOut Then validSellno = 0

					strRst = strRst &"<tns:ltSupplyPlans>"
					strRst = strRst &"	<tns:unitCd>"&arrRows(3,i)&"</tns:unitCd>"
					strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
					strRst = strRst &"	<tns:strDt>"&FormatDate(now(), "0000-00-00")&"</tns:strDt>"
					If validSellno = 0 Then
						strRst = strRst &"	<tns:endDt>"&FormatDate(now(), "0000-00-00")&"</tns:endDt>"
					Else
						strRst = strRst &"	<tns:endDt>9999-12-30</tns:endDt>"
					End If
					strRst = strRst &"	<tns:availSupQty>"&chkiif(validSellno=0,"1",validSellno)&"</tns:availSupQty>"
					strRst = strRst &"</tns:ltSupplyPlans>"
				Next
				strRst = strRst &"</tns:ifRequest>"
			End If
		End If
		getcjmallItemQTYXML = strRst
	End Function

    '// ��ǰ ������ ���� �Ķ���� ����
	Public Function getcjmallItemDateXML
		Dim sqlStr
		Dim arrRows, i, strRst, validSellno
		sqlStr = ""
		sqlStr = sqlStr & " SELECT o.itemoption, o.optionTypeName, o.optionname, isnull(R.outmallOptCode, '') as outmallOptCode, (o.optlimitno-o.optlimitsold) as optLimit, o.optlimityn, o.isUsing, o.optsellyn " & VBCRLF
		sqlStr = sqlStr & " FROM [db_AppWish].[dbo].tbl_item_option o " & VBCRLF
		sqlStr = sqlStr & " left join [db_AppWish].[dbo].tbl_OutMall_regedoption R on o.itemid=R.itemid and o.itemoption=R.itemoption and R.mallid='cjmall' " & VBCRLF
		sqlStr = sqlStr & " where R.outmallOptCode <> '' and o.itemid="&Fitemid
		rsCTget.Open sqlStr, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			arrRows = rsCTget.getRows
		End If
		rsCTget.close

		validSellno = 50				'�ִ� 50���� ��������
		If isArray(arrRows) Then

			strRst = ""
			strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
			strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_06"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_06.xsd"">"
			strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"						'!!!���¾�ü�ڵ�
			strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"		'!!!����Ű
			For i = 0 To UBound(ArrRows,2)
				If (FSellyn <> "Y") or ((arrRows(5,i) = "Y") and (arrRows(4,i) < 1)) or (arrRows(6,i) <> "Y") or (arrRows(7,i) <> "Y") Then
					validSellno = 0
				End If

				If (arrRows(5,i) = "Y") Then
					validSellno = arrRows(4,i)
				End If

				If (validSellno < CMAXLIMITSELL) Then validSellno = 0
				If (arrRows(5,i) = "Y") and (validSellno > 0) Then
					validSellno = validSellno - CMAXLIMITSELL
				End If
				If (validSellno < 1) then validSellno = 0
				If IsSoldOut Then validSellno = 0

				strRst = strRst &"<tns:contents>"
		        strRst = strRst &"	<tns:chnCd>30</tns:chnCd>"
		        strRst = strRst &"	<tns:unitCd>"&arrRows(3,i)&"</tns:unitCd>"
		        strRst = strRst &"	<tns:insRsn>10</tns:insRsn>"
		        strRst = strRst &"	<tns:applyStrDtm>"&FormatDate(now(), "0000-00-00")&"</tns:applyStrDtm>"
		        strRst = strRst &"	<tns:applyEndDtm>2013-06-10</tns:applyEndDtm>"
		        strRst = strRst &"	<tns:availSupDt>2013-06-11</tns:availSupDt>"
		        strRst = strRst &"	<tns:availOrdQty>"&validSellno&"</tns:availOrdQty>"
				strRst = strRst &"</tns:contents>"
			Next
			strRst = strRst &"</tns:ifRequest>"
		End If
		getcjmallItemDateXML = strRst
	End Function

	'// ��ǰ ����- �Ͻ��ߴ� �Ķ���� ����
    Public Function getcjmallItemSellStatusDTXML
		Dim stopYN, strRst

		If FSellYN = "N" Then
			stopYN = "I"
		ElseIf FSellYn = "Y" Then
			stopYN = "A"
		End If

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"					'!!!���¾�ü�ڵ�
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"	'!!!����Ű
		strRst = strRst &"<tns:itemStates>"
		strRst = strRst &"	<tns:typeCd>01</tns:typeCd>"								'!!!01=�Ǹ��ڵ�,02=��ǰ�ڵ�)
		strRst = strRst &"	<tns:itemCd_zip>"&FcjmallPrdNo&"</tns:itemCd_zip>"
		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
		strRst = strRst &"	<tns:packInd>"&stopYN&"</tns:packInd>"						'!!!A-����, I-�Ͻ��ߴ�
		strRst = strRst &"</tns:itemStates>"
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemSellStatusDTXML = strRst
	End Function

End Class

Class CCjmall
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectMakerid
	Public FRectItemName
	Public FRectCJMallPrdNo
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectItemID
	Public FRectEventid
	Public FRectExtNotReg
	Public FRectOrdType
	Public FRectMatchCate
	Public FRectPrdDivMatch
	''Public FRectMatchCateNotCheck
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectMinusMargin
	Public FRectFailCntExists
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptExists
	Public FRectoptnotExists
    Public FRectCjSell10x10Soldout
    Public FRectCjshowminusmagin
    Public FRectexpensive10x10
	Public FRectExtSellYn
	public FRectOnlyNotUsingCheck
	public FRectdiffPrc

	'ī�װ�
	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FRectDspNo

	Public FRectMode

	Public Finfodiv
	Public FCateName
	Public FsearchName
	Public FRectdisptpcd

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

	''Cjmall ��ϸ���Ʈ
	Public Sub GetCjmallRegedItemList()
		Dim i, sqlStr, addSql
		Dim ssnKey, dumitype
		If FRectEventid <> "" Then
		    ssnKey = session("ssBctID")&"["&NOW()&"]"
		    dumitype = "evt"&FRectEventid
		    
		    
		    sqlStr = "insert into db_outMall.dbo.tbl_OutMall_Q_dumi"
		    sqlStr = sqlStr &" (ssnKey,dumitype,itemid)"
		    sqlStr = sqlStr &" Select '"&ssnKey&"','"&dumitype&"',itemid From [TENDB].[db_event].[dbo].tbl_eventitem Where evt_code='" & FRectEventid & "'"
		    dbCTget.Execute sqlStr
		    
		    sqlStr = ""
	    end if
		
		'�귣��˻�
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'Cjmall��ǰ��ȣ �˻�
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'��ǰ�� �˻�
		If FRectCJMallPrdNo <> "" Then
			addSql = addSql & " and J.cjmallPrdNo='" & FRectCJMallPrdNo & "'"
		End If

		'ī�װ� �˻�
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'��ǰ��ȣ �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'�̺�Ʈ��ȣ �˻� // �����ʿ�..
		If FRectEventid <> "" Then
			addSql = addSql & " and i.itemid in (Select itemid From [db_outMall].[dbo].tbl_OutMall_Q_dumi Where ssnKey='"&ssnKey&"' and dumitype='"&dumitype&"')" & VbCrlf
		End If

		'��Ͽ��� �˻�
		Select Case FRectExtNotReg
			Case "M"	'�̵��
			    addSql = addSql & " and J.itemid is NULL "
			Case "Q"	''��Ͻ���
				addSql = addSql & " and J.cjmallStatCd = -1"
			Case "J"	'��Ͽ����̻�
				addSql = addSql & " and J.cjmallStatCd >= 0"
			Case "W"	'��Ͽ���
				addSql = addSql & " and J.cjmallStatCd = 0"
		    Case "A"	'���۽õ�
				addSql = addSql & " and J.cjmallStatCd = 1"
			Case "F"	'��ϿϷ�(�ӽ�)
			    addSql = addSql & " and J.cjmallStatCd = 3"
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.cjmallStatCd = 7"
				addSql = addSql & " and J.cjmallPrdNo is Not Null"
			Case "R"	'�������
			    addSql = addSql & " and J.cjmallStatCd = 7"
			    addSql = addSql & " and J.cjmallLastUpdate < i.lastupdate"
		End Select

		'ī�׸�Ī �˻�
		Select Case FRectMatchCate
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and isnull(c.mapCnt, '') <> ''"
			Case "N"	'�̸�Ī
				addSql = addSql & " and isnull(c.mapCnt, '') = ''"
		End Select

		'�з���Ī �˻�
		Select Case FRectPrdDivMatch
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and IsNull(pd.cddkey, '') <> '' "
			Case "N"	'�̸�Ī
				addSql = addSql & " and IsNull(pd.cddkey, '') = '' "
		End Select

		'�Ǹſ��� �˻�
		Select Case FRectSellYn
			Case "Y"	'�Ǹ�
				addSql = addSql & " and i.sellYn='Y'"
			Case "N"	'ǰ��
				addSql = addSql & " and i.sellYn in ('S','N')"
		End Select

        if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        addSql = addSql + " and J.cjmallSellYn<>'X'"
		    else
		        addSql = addSql + " and J.cjmallSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

        '�������� �˻�
		If FRectLimitYn<>"" then
			addSql = addSql & " and i.limitYn='" & FRectLimitYn & "'" & VbCrlf
		End if

		If FRectSailYn<>"" then
			addSql = addSql & " and i.sailYn='" & FRectSailYn & "'" & VbCrlf
		End if

		'������ �� ���� CMAXMARGIN �̻� �˻�  (CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END) ����//2014/07/21
		If (FRectCjshowminusmagin<>"") then
		   addSql = addSql & " and i.sellcash<>0"
		   addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<"&CMAXMARGIN & VbCrlf   
		   addSql = addSql & " and J.cjmallSellYn= 'Y' " '''  ���� �߰�.
		Else
		   IF (FRectonlyValidMargin<>"") then
		        addSql = addSql & " and i.sellcash<>0"
		        addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&CMAXMARGIN & VbCrlf
		   END IF
		End If


		If (FRectFailCntExists <> "") Then
			addSql = addSql & " and J.accFailCNT > 0"
		End If

		If Finfodiv <> "" then
			addSql = addSql & " and PD.infodiv = '" & Finfodiv & "'"
		End if

		''�ɼ��߰��ݾ� �����ǰ.
		If (FRectoptAddprcExists <> "") and (FRectExtNotReg <> "M") Then
		    addSql = addSql & " and J.optAddPrcCnt>0"
'			addSql = addSql & " and i.itemid in ("
'			addSql = addSql & "     select distinct ii.itemid "
'			addSql = addSql & "     from db_AppWish.dbo.tbl_item ii "
'			addSql = addSql & "     Join db_AppWish.dbo.tbl_item_option o "
'			addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
'			addSql = addSql & " )"
		End If

		''�ɼ��߰��ݾ� �����ǰ ����
		If (FRectoptAddprcExistsExcept <> "") Then
		    addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
'			addSql = addSql & " and i.itemid Not in ("
'			addSql = addSql & "     select distinct ii.itemid "
'			addSql = addSql & "     from db_AppWish.dbo.tbl_item ii "
'			addSql = addSql & "     Join db_AppWish.dbo.tbl_item_option o "
'			addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
'			addSql = addSql & " )"
		End If

		if (FRectoptExists<>"") then
            addSql = addSql & " and i.optioncnt>0"
        end if

		if (FRectoptnotExists<>"") then
            addSql = addSql & " and i.optioncnt=0"
        end if

        ''cj�Ǹ� 10x10 ǰ��
        IF (FRectCjSell10x10Soldout<>"") then
            addSql = addSql & " and i.sellyn<>'Y'"
            addSql = addSql & " and J.cjmallSellyn='Y'"
		Else
    		addSql = addSql & " and i.isExtUsing='Y'"
    		addSql = addSql & " and i.deliverytype not in ('7')"
            addSql = addSql + " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
        end if

        ''cj���� <10x10 ����
        IF (FRectexpensive10x10<>"") then
            addSql = addSql & " and J.cjmallPrice<i.sellcash"
        end if

        if FRectdiffPrc <> "" then
		   addSql = addSql & " and J.cjmallPrice is Not Null and i.sellcash <> J.cjmallPrice "
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i " & VBCRLF
		''sqlStr = sqlStr & " INNER JOIN db_AppWish.dbo.tbl_item_contents as t on i.itemid = t.itemid " & VBCRLF
	    sqlStr = sqlStr & " LEFT JOIN db_outMall.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid " & VBCRLF
		sqlStr = sqlStr & "	LEFT JOIN db_outMall.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_AppWish.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " LEFT JOIN db_outMall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and i.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and i.isusing='Y' " & VBCRLF
'		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') " & VBCRLF
		sqlStr = sqlStr & " and i.basicimage is not null " & VBCRLF
		sqlStr = sqlStr & " and i.itemdiv<50 " & VBCRLF  '''and i.itemdiv<>'08'
		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
		sqlStr = sqlStr & " and i.cate_large<>'' " & VBCRLF
		''sqlStr = sqlStr & " and (i.cate_large <> '999')" & VBCRLF     ''2013/07/19 ftroupe ����ó��
		sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
		sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
		sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
		sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000)"	'���ǹ���̸� 10000�� �̸� ����
		sqlStr = sqlStr & addSql
		'20130514 ä���� ���� ��û ī�װ� ����
'		sqlStr = sqlStr & "	and i.cate_large <> '080'"
'		sqlStr = sqlStr & "	and i.cate_large <> '090'"
'		sqlStr = sqlStr & "	and i.cate_large <> '070'"
'		sqlStr = sqlStr & "	and i.cate_large <> '100'"
'		sqlStr = sqlStr & "	and i.cate_large <> '075'"
		'2013-12-31 ä���� ���� ��û ��Ƽī�װ� �� �Ϻ�ī�װ��� ���� ==> ��Ƽ-���̾�Ʈ-��ⱸ, ��Ƽ-���̾�Ʈ-ü�߰�/������, ��Ƽ-�׷�� ��Ƽ-��Ÿ, ��Ƽ-��Ƽ���-����� ����
		sqlStr = sqlStr & " and (i.cate_large + i.cate_mid not in ('075001', '075002', '075003', '075004', '075005', '075006', '075009', '075010', '075012', '075013', '075016', '075018') ) "
		sqlStr = sqlStr & " and (i.cate_large + i.cate_mid + i.cate_small not in ('075020001', '075020005', '075020006', '075020007', '075020008', '075021001', '075021002', '075021003', '075014001', '075014002') ) "
		'2013-12-31 ä���� ���� ��û ��Ƽī�װ� �� �Ϻ�ī�װ��� ���� ==> ��Ƽ-���̾�Ʈ-��ⱸ, ��Ƽ-���̾�Ʈ-ü�߰�/������, ��Ƽ-�׷�� ��Ƽ-��Ÿ, ��Ƽ-��Ƽ���-����� ��
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110010')"
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110030')"
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110040')"
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110060')"
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110050')"
		'20130514 ä���� ���� ��û ī�װ� ���� ��
'rw sqlStr
'response.end

		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close
		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
''rw sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage " & VBCRLF
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash " & VBCRLF
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt, c.mapCnt " & VBCRLF
		sqlStr = sqlStr & "	, J.cjmallRegdate, J.cjmallLastUpdate, J.cjmallPrdNo, J.cjmallPrice, J.cjmallSellYn, J.regUserid, IsNULL(J.cjmallStatCd,-9) as cjmallStatCd  " & VBCRLF
		sqlStr = sqlStr & "	, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr, PD.infodiv, PD.cdmKey, PD.cddkey, UC.defaultfreeBeasongLimit " & VBCRLF
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i " & VBCRLF
'		sqlStr = sqlStr & " INNER JOIN db_AppWish.dbo.tbl_item_contents as t on i.itemid = t.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_outMall.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid " & VBCRLF
		sqlStr = sqlStr & "	LEFT JOIN db_outMall.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr + " LEFT JOIN db_AppWish.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " LEFT JOIN db_outMall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and i.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and i.isusing='Y' " & VBCRLF
'		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') " & VBCRLF
		sqlStr = sqlStr & " and i.basicimage is not null " & VBCRLF
		sqlStr = sqlStr & " and i.itemdiv < 50 " & VBCRLF  '''and i.itemdiv<>'08'
		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
		sqlStr = sqlStr & " and i.cate_large <> '' " & VBCRLF
		''sqlStr = sqlStr & " and (i.cate_large <> '999')" & VBCRLF     ''2013/07/19 ftroupe ����ó��
		sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
		sqlStr = sqlStr & " and i.sellcash >= 1000 "  & VBCRLF
		sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
		sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000)"		'���ǹ���̸� 10000�� �̸� ����
		sqlStr = sqlStr & addSql
		'20130514 ä���� ���� ��û ī�װ� ����
'		sqlStr = sqlStr & "	and i.cate_large <> '080'"
'		sqlStr = sqlStr & "	and i.cate_large <> '090'"
'		sqlStr = sqlStr & "	and i.cate_large <> '070'"
'		sqlStr = sqlStr & "	and i.cate_large <> '100'"
'		sqlStr = sqlStr & "	and i.cate_large <> '075'"
		'2013-12-31 ä���� ���� ��û ��Ƽī�װ� �� �Ϻ�ī�װ��� ���� ==> ��Ƽ-���̾�Ʈ-��ⱸ, ��Ƽ-���̾�Ʈ-ü�߰�/������, ��Ƽ-�׷�� ��Ƽ-��Ÿ, ��Ƽ-��Ƽ���-����� ����
		sqlStr = sqlStr & " and (i.cate_large + i.cate_mid not in ('075001', '075002', '075003', '075004', '075005', '075006', '075009', '075010', '075012', '075013', '075016', '075018') ) "
		sqlStr = sqlStr & " and (i.cate_large + i.cate_mid + i.cate_small not in ('075020001', '075020005', '075020006', '075020007', '075020008', '075021001', '075021002', '075021003', '075014001', '075014002') ) "
		'2013-12-31 ä���� ���� ��û ��Ƽī�װ� �� �Ϻ�ī�װ��� ���� ==> ��Ƽ-���̾�Ʈ-��ⱸ, ��Ƽ-���̾�Ʈ-ü�߰�/������, ��Ƽ-�׷�� ��Ƽ-��Ÿ, ��Ƽ-��Ƽ���-����� ��
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110010')"
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110030')"
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110040')"
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110060')"
		sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110050')"
		'20130514 ä���� ���� ��û ī�װ� ���� ��

		If FRectExtNotReg = "M" Then
			sqlStr = sqlStr & " ORDER BY i.itemid DESC"
		Else
			IF (FRectOrdType = "B") Then
				sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC"
			ElseIf (FRectOrdType = "BM") Then
				sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
			ElseIf (FRectOrdType = "PM") Then
				sqlStr = sqlStr & " ORDER BY J.lastPriceCheckDate"
			Else
				sqlStr = sqlStr & " ORDER BY J.itemid DESC"
		    End If
	    End If
''rw sqlStr

		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do Until rsCTget.EOF
				Set FItemList(i) = new cjmallItem
					FItemList(i).Fitemid			= rsCTget("itemid")
					FItemList(i).Fitemname			= rsCTget("itemname")
					FItemList(i).FsmallImage		= rsCTget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsCTget("itemid")) & "/" & rsCTget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).Fsellcash			= rsCTget("sellcash")
					FItemList(i).Fbuycash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).Fsaleyn			= rsCTget("sailyn")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")
					FItemList(i).Fdeliverytype		= rsCTget("deliverytype")
					FItemList(i).FoptionCnt			= rsCTget("optionCnt")
					FItemList(i).FcjmallRegdate		= rsCTget("cjmallRegdate")
					FItemList(i).FcjmallLastUpdate	= rsCTget("cjmallLastUpdate")
					FItemList(i).FcjmallPrdNo		= rsCTget("cjmallPrdNo")
					FItemList(i).FcjmallPrice		= rsCTget("cjmallPrice")
					FItemList(i).FcjmallSellYn		= rsCTget("cjmallSellYn")
					FItemList(i).FregUserid			= rsCTget("regUserid")
					FItemList(i).FcjmallStatCd		= rsCTget("cjmallStatCd")
					FItemList(i).FregedOptCnt		= rsCTget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsCTget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsCTget("accFailCNT")
					FItemList(i).FlastErrStr		= rsCTget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsCTget("mapCnt")
					FItemList(i).Finfodiv			= rsCTget("infodiv")
					FItemList(i).FcdmKey			= rsCTget("cdmKey")
					FItemList(i).FcddKey			= rsCTget("cddKey")
					FItemList(i).FdefaultfreeBeasongLimit = rsCTget("defaultfreeBeasongLimit")

				i = i + 1
				rsCTget.MoveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	''' ��ϵ��� ���ƾ� �� ��ǰ..
	Public Sub getCjmallreqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_outMall.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid and J.cjmallprdNo is not null " & VBCRLF
'		sqlStr = sqlStr & " INNER JOIN db_outMall.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid" & VBCRLF
'		sqlStr = sqlStr & " INNER JOIN db_AppWish.dbo.tbl_item_contents as t on i.itemid = t.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_AppWish.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & "	LEFT JOIN db_outMall.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_outMall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and i.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF

        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe ����ó��

		if (FRectOnlyNotUsingCheck="on") then
		    sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From db_outMall.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From db_outMall.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
    		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
            sqlStr = sqlStr + "	)"
            ''//���� ���ܻ�ǰ
            sqlStr = sqlStr & " and i.itemid not in ("
            sqlStr = sqlStr & "     select itemid from db_outMall.dbo.tbl_OutMall_etcLink"
 			sqlStr = sqlStr & "     where getdate() between stdt and eddt"
            sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
            sqlStr = sqlStr & "     and linkgbn='donotEdit'"
            sqlStr = sqlStr & " )"
		else
    		sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		sqlStr = sqlStr + "     or i.deliverytype in ('7') "
    		'//���ǹ�� 10000�� �̻�
            sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<10000))"
    		sqlStr = sqlStr + "     or i.deliverfixday in ('C','X') "
    		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From db_outMall.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From db_outMall.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
    		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
            sqlStr = sqlStr + "	)"

            ''//���� ���ܻ�ǰ
            sqlStr = sqlStr & " and i.itemid not in ("
            sqlStr = sqlStr & "     select itemid from db_outMall.dbo.tbl_OutMall_etcLink"
 			sqlStr = sqlStr & "     where getdate() between stdt and eddt"
            sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
            sqlStr = sqlStr & "     and linkgbn='donotEdit'"
            sqlStr = sqlStr & " )"
        end if

        if FRectMakerid<>"" then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		end if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        sqlStr = sqlStr + " and J.cjmallSellYn<>'X'"
		    else
		        sqlStr = sqlStr + " and J.cjmallSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

		Select Case FRectSellYn
			Case "Y"	'�Ǹ�
				sqlStr = sqlStr & " and i.sellYn='Y'"
			Case "N"	'ǰ��
				sqlStr = sqlStr & " and i.sellYn in ('S','N')"
		End Select

		if (Finfodiv<>"") then
		    if Finfodiv="Y" then
		        sqlStr = sqlStr + " and isNULL(i.infoDiv,'')<>''"
		    elseif Finfodiv="N" then
    		    sqlStr = sqlStr + " and isNULL(i.infoDiv,'')=''"
    		else
    		    sqlStr = sqlStr + " and i.infoDiv='"&Finfodiv&"'"
    		end if
		end if

		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage " & VBCRLF
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash " & VBCRLF
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt, c.mapCnt " & VBCRLF
		sqlStr = sqlStr & "	, J.cjmallRegdate, J.cjmallLastUpdate, J.cjmallPrdNo, J.cjmallPrice, J.cjmallSellYn, J.regUserid, IsNULL(J.cjmallStatCd,-9) as cjmallStatCd  " & VBCRLF
		sqlStr = sqlStr & "	, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr, PD.infodiv, PD.cdmKey, PD.cddkey, UC.defaultfreeBeasongLimit " & VBCRLF
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_outMall.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid and J.cjmallprdNo is not null " & VBCRLF
'		sqlStr = sqlStr & " INNER JOIN db_outMall.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid" & VBCRLF
'		sqlStr = sqlStr & " INNER JOIN db_AppWish.dbo.tbl_item_contents as t on i.itemid = t.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_AppWish.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & "	LEFT JOIN db_outMall.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_outMall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and i.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF

		if (FRectOnlyNotUsingCheck="on") then
		    sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From db_outMall.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From db_outMall.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
    		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
            sqlStr = sqlStr + "	)"
            ''//���� ���ܻ�ǰ
            sqlStr = sqlStr & " and i.itemid not in ("
            sqlStr = sqlStr & "     select itemid from db_outMall.dbo.tbl_OutMall_etcLink"
            sqlStr = sqlStr & "     where getdate() between stdt and eddt"
            sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
            sqlStr = sqlStr & "     and linkgbn='donotEdit'"
            sqlStr = sqlStr & " )"
		else
    		sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		sqlStr = sqlStr + "     or i.deliverytype in ('7') "
    		'//���ǹ�� 10000�� �̻�
            sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<10000))"
    		sqlStr = sqlStr + "     or i.deliverfixday in ('C','X') "
    		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From db_outMall.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From db_outMall.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
    		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
            sqlStr = sqlStr + "	)"

            ''//���� ���ܻ�ǰ //���� ������ �ҵ�.
            sqlStr = sqlStr & " and i.itemid not in ("
            sqlStr = sqlStr & "     select itemid from db_outMall.dbo.tbl_OutMall_etcLink"
 			sqlStr = sqlStr & "     where getdate() between stdt and eddt"
            sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
            sqlStr = sqlStr & "     and linkgbn='donotEdit'"
            sqlStr = sqlStr & " )"
        end if

        if FRectMakerid<>"" then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		end if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        sqlStr = sqlStr + " and J.cjmallSellYn<>'X'"
		    else
		        sqlStr = sqlStr + " and J.cjmallSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

		Select Case FRectSellYn
			Case "Y"	'�Ǹ�
				sqlStr = sqlStr & " and i.sellYn='Y'"
			Case "N"	'ǰ��
				sqlStr = sqlStr & " and i.sellYn in ('S','N')"
		End Select

		if (Finfodiv<>"") then
		    if Finfodiv="Y" then
		        sqlStr = sqlStr + " and isNULL(i.infoDiv,'')<>''"
		    elseif Finfodiv="N" then
    		    sqlStr = sqlStr + " and isNULL(i.infoDiv,'')=''"
    		else
    		    sqlStr = sqlStr + " and i.infoDiv='"&Finfodiv&"'"
    		end if
		end if

		sqlStr = sqlStr + " order by J.regdate desc, i.itemid desc "
''rw sqlStr

		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1

		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsCTget.EOF  then
			rsCTget.absolutepage = FCurrPage
			do until rsCTget.eof
				Set FItemList(i) = new cjmallItem

					FItemList(i).Fitemid			= rsCTget("itemid")
					FItemList(i).Fitemname			= rsCTget("itemname")
					FItemList(i).FsmallImage		= rsCTget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsCTget("itemid")) & "/" & rsCTget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).Fsellcash			= rsCTget("sellcash")
					FItemList(i).Fbuycash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).Fsaleyn			= rsCTget("sailyn")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")
					FItemList(i).Fdeliverytype		= rsCTget("deliverytype")
					FItemList(i).FoptionCnt			= rsCTget("optionCnt")
					FItemList(i).FcjmallRegdate		= rsCTget("cjmallRegdate")
					FItemList(i).FcjmallLastUpdate	= rsCTget("cjmallLastUpdate")
					FItemList(i).FcjmallPrdNo		= rsCTget("cjmallPrdNo")
					FItemList(i).FcjmallPrice		= rsCTget("cjmallPrice")
					FItemList(i).FcjmallSellYn		= rsCTget("cjmallSellYn")
					FItemList(i).FregUserid			= rsCTget("regUserid")
					FItemList(i).FcjmallStatCd		= rsCTget("cjmallStatCd")
					FItemList(i).FregedOptCnt		= rsCTget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsCTget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsCTget("accFailCNT")
					FItemList(i).FlastErrStr		= rsCTget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsCTget("mapCnt")
					FItemList(i).Finfodiv			= rsCTget("infodiv")
					FItemList(i).FcdmKey			= rsCTget("cdmKey")
					FItemList(i).FcddKey			= rsCTget("cddKey")
					FItemList(i).FdefaultfreeBeasongLimit = rsCTget("defaultfreeBeasongLimit")
				i=i+1
				rsCTget.moveNext
			loop
		end if
		rsCTget.Close
	End Sub

	'// �ٹ�����-cjmall ī�װ�
	Public Sub getTencjmallCateList
		Dim sqlStr, addSql, i

		If FRectCDL<>"" Then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		End if

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and T.CateKey is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.CateKey is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'cjmall �����ڵ� �˻�
					addSql = addSql & " and T.CateKey='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�ٹ����� �Һз���)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.D_Name,cc.L_Name,cc.M_Name,cc.S_Name, cc.isusing, cc.CateGbn "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_outMall.dbo.tbl_cjmall_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_outMall.dbo.tbl_cjMall_Category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and cc.CateGbn='"&FRectdisptpcd&"'"
        End If
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.CateKey as DispNo ,T.D_Name as DispNm, T.L_Name as DispLrgNm, T.M_Name as DispMidNm, T.S_Name as DispSmlNm ,T.IsUsing as CateIsUsing,T.cateGbn as disptpcd "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.D_Name,cc.L_Name,cc.M_Name,cc.S_Name, cc.isusing, cc.CateGbn  "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_outMall.dbo.tbl_cjmall_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_outMall.dbo.tbl_cjMall_Category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and cc.CateGbn='"&FRectdisptpcd&"'"
        End If
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small, T.CateGbn  ASC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cjmallItem
					FItemList(i).FtenCateLarge		= rsCTget("code_large")
					FItemList(i).FtenCateMid		= rsCTget("code_mid")
					FItemList(i).FtenCateSmall		= rsCTget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsCTget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsCTget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsCTget("small_nm"))
					FItemList(i).FDispNo			= rsCTget("DispNo")
					FItemList(i).FDispNm			= db2html(rsCTget("DispNm"))
					FItemList(i).FDispLrgNm			= db2html(rsCTget("DispLrgNm"))
					FItemList(i).FDispMidNm			= db2html(rsCTget("DispMidNm"))
					FItemList(i).FDispSmlNm			= db2html(rsCTget("DispSmlNm"))
					FItemList(i).Fdisptpcd			= rsCTget("disptpcd")
	                FItemList(i).FCateIsUsing		= rsCTget("CateIsUsing")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	'// �ٹ�����-cjmall ��ǰ�з�
	Public Sub getTencjmallprdDivList
		Dim sqlStr, addSql, i

		If FRectCDL<>"" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql = addSql & " and c.infodiv='" & Finfodiv & "'"
		End if

		If FRectIsMapping <> "" Then
			If FRectIsMapping = "Y" Then
				addSql = addSql & " and isnull(P.CddKey, '') <> '' "
			ElseIf FRectIsMapping = "N" Then
				addSql = addSql & " and isnull(P.CddKey, '') = '' "
			End If
		End if

		If FCateName <> "" AND FsearchName <> "" Then
			Select Case FCateName
				Case "cdlnm"
					addSql = addSql & " and v.nmlarge like '%" & FsearchName & "%'"
				Case "cdmnm"
					addSql = addSql & " and v.nmmid like '%" & FsearchName & "%'"
				Case "cdsnm"
					addSql = addSql & " and v.nmsmall like '%" & FsearchName & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM  ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " 	, v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & "		,P.CddKey ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name,P.IsUsing as PrdDivIsUsing, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_AppWish.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " 	INNER JOIN db_AppWish.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " 	LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "		LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 		SELECT dm.CddKey, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.cdd_Name, pv.cdl_Name, pv.cdm_Name, pv.cds_Name, pv.isusing, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 		FROM db_outMall.dbo.tbl_cjMall_prdDiv_mapping as dm  "  & VBCRLF
		sqlStr = sqlStr & " 		JOIN db_outMall.dbo.tbl_cjMall_prdDiv as pv on dm.CddKey = pv.cdd  "  & VBCRLF
		sqlStr = sqlStr & " 	) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv   "  & VBCRLF
		sqlStr = sqlStr & " 	WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall,P.CddKey ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.IsUsing, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " ) as T " & VBCRLF
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " , v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & " ,P.CddKey ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.IsUsing as PrdDivIsUsing, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_AppWish.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "	LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT dm.CddKey, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.cdd_Name, pv.cdl_Name, pv.cdm_Name, pv.cds_Name, pv.isusing, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_outMall.dbo.tbl_cjMall_prdDiv_mapping as dm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_outMall.dbo.tbl_cjMall_prdDiv as pv on dm.CddKey = pv.cdd  "  & VBCRLF
		sqlStr = sqlStr & " ) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall,P.CddKey ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.IsUsing, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " ORDER BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cjmallItem
					FItemList(i).Finfodiv		= rsCTget("infodiv")
					FItemList(i).FtenCateLarge	= rsCTget("cate_large")
					FItemList(i).FtenCateMid	= rsCTget("cate_mid")
					FItemList(i).FtenCateSmall	= rsCTget("cate_small")
					FItemList(i).FtenCDLName	= rsCTget("nmlarge")
					FItemList(i).FtenCDMName	= rsCTget("nmmid")
					FItemList(i).FtenCDSName	= rsCTget("nmsmall")
					FItemList(i).Ficnt			= rsCTget("icnt")
					FItemList(i).FCddKey		= rsCTget("CddKey")
					FItemList(i).Fcdd_Name		= rsCTget("cdd_Name")
					FItemList(i).Fcdl_Name		= rsCTget("cdl_Name")
					FItemList(i).Fcdm_Name		= rsCTget("cdm_Name")
					FItemList(i).Fcds_Name		= rsCTget("cds_Name")
					FItemList(i).FPrdDivIsUsing	= rsCTget("PrdDivIsUsing")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Function getTencjmallOneprdDiv
		Dim sqlStr, addSql, addsql2

		If FRectCDL<>"" Then
			addSql = addSql & " and v.cdlarge='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and v.cdmid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and v.cdsmall='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql2 = addSql2 & " and p.infodiv='" & Finfodiv & "'"
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 p.CddKey, p.infodiv, p.CdmKey, p.tenCateLarge, p.tenCateMid, p.tenCateSmall, v.nmlarge, v.nmmid, v.nmsmall, T.cdd_NAME " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.vw_category as v " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_outMall.dbo.tbl_cjMall_prdDiv_mapping p on p.tenCateLarge = v.cdlarge and p.tenCateMid = v.cdmid and p.tenCateSmall = v.cdsmall " & addsql2
		sqlStr = sqlStr & " LEFT JOIN db_outMall.dbo.tbl_cjMall_prdDiv as T on p.cddKey = T.cdd " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount

		If not rsCTget.EOF Then
			Set FItemList(0) = new cjmallItem
				FItemList(0).Finfodiv		= rsCTget("infodiv")
				FItemList(0).FtenCateLarge	= rsCTget("tenCateLarge")
				FItemList(0).FtenCateMid	= rsCTget("tenCateMid")
				FItemList(0).FtenCateSmall	= rsCTget("tenCateSmall")
				FItemList(0).FtenCDLName	= rsCTget("nmlarge")
				FItemList(0).FtenCDMName	= rsCTget("nmmid")
				FItemList(0).FtenCDSName	= rsCTget("nmsmall")
				FItemList(0).FCddKey		= rsCTget("CddKey")
				FItemList(0).Fcdd_Name		= rsCTget("cdd_NAME")
		End If
		rsCTget.Close
	End Function

	'// cjmall ī�װ�
	Public Sub getcjmallCategoryList
		Dim sqlStr, addSql, i

		If FRectDspNo <> "" Then
			addSql = addSql & " and c.cateKey = " & FRectDspNo
		End If

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "CCD"	'cjmall �����ڵ� �˻�
					addSql = addSql & " and c.cateKey='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���
					addSql = addSql & " and (c.D_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.S_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.M_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.L_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " )"
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(c.cateKey) as cnt, CEILING(CAST(Count(c.cateKey) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_outMall.dbo.tbl_cjMall_Category as c " & VBCRLF
		sqlStr = sqlStr & " WHERE c.M_Name like '%�ٹ�����%' " & addSql
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " c.* " & VBCRLF
		sqlStr = sqlStr & " FROM db_outMall.dbo.tbl_cjMall_Category as c " & VBCRLF
		sqlStr = sqlStr & " WHERE c.M_Name like '%�ٹ�����%' " & addSql
		sqlStr = sqlStr & " ORDER BY c.cateKey ASC"
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.eof
				Set FItemList(i) = new cjmallItem
					FItemList(i).FDispNo		= rsCTget("cateKey")
					FItemList(i).FDispNm		= db2html(rsCTget("D_Name"))
					FItemList(i).FDispLrgNm		= db2html(rsCTget("L_Name"))
					FItemList(i).FDispMidNm		= db2html(rsCTget("M_Name"))
					FItemList(i).FDispSmlNm		= db2html(rsCTget("S_Name"))
					FItemList(i).FDispThnNm		= db2html(rsCTget("D_Name"))
					FItemList(i).FisUsing		= rsCTget("isUsing")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	'// cjmall ��ǰ�з�
	Public Sub getcjmallPrdDivList
		Dim sqlStr, addSql, i

		If Finfodiv <> "" Then
			addSql = addSql & " and m.infodiv = '" & Finfodiv & "'"
		End If

		If FsearchName <> "" Then
			addSql = addSql & " and p.cdd_NAME like '%" & FsearchName & "%'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " from db_outMall.dbo.tbl_cjmall_PrddivMid_map as m " & VBCRLF
		sqlStr = sqlStr & " inner join db_outMall.dbo.tbl_cjMall_prdDiv as p on m.cjcdm = p.cdm " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " m.infodiv, p.cdm, p.cdd, p.cdl_NAME, p.cdm_NAME, p.cds_NAME, p.cdd_NAME, p.isusing  " & VBCRLF
		sqlStr = sqlStr & " FROM db_outMall.dbo.tbl_cjmall_PrddivMid_map as m " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_outMall.dbo.tbl_cjMall_prdDiv as p on m.cjcdm = p.cdm " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY m.infodiv, p.cdm, p.cdd"
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.eof
				Set FItemList(i) = new cjmallItem
					FItemList(i).Finfodiv		= rsCTget("infodiv")
					FItemList(i).FCdm			= rsCTget("cdm")
					FItemList(i).FCdd			= rsCTget("cdd")
					FItemList(i).FDispNm		= db2html(rsCTget("cdd_NAME"))
					FItemList(i).FDispLrgNm		= db2html(rsCTget("cdl_NAME"))
					FItemList(i).FDispMidNm		= db2html(rsCTget("cdm_NAME"))
					FItemList(i).FDispSmlNm		= db2html(rsCTget("cds_NAME"))
					FItemList(i).FDispThnNm		= db2html(rsCTget("cdd_NAME"))
					FItemList(i).Fisusing		= rsCTget("isusing")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	'// �̵�� ��ǰ ���(��Ͽ�)
	Public Sub getCJMallNotRegItemList
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' �ɼ� �߰��ݾ� �ִ°�� ��� �Ұ�. //�ɼ� ��ü ǰ���� ��� ��� �Ұ�.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "	SELECT itemid FROM ("
            addSql = addSql & "     SELECT itemid"
            addSql = addSql & " 	,count(*) as optCNT"
			addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	FROM db_AppWish.dbo.tbl_item_option"
            addSql = addSql & " 	WHERE itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	GROUP BY itemid"
            addSql = addSql & " ) T"
            'addSql = addSql & " WHERE optAddCNT>0 or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey "
		strSql = strSql & "	, isNULL(R.cjmallStatCD,-9) as cjmallStatCD "
		strSql = strSql & "	, UC.socname_kor, PD.cdmkey, PD.cddkey "
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_AppWish.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_outMall.dbo.tbl_cjmall_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as PmapCnt "
		strSql = strSql & "		FROM db_outMall.dbo.tbl_cjMall_prdDiv_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as Pm on Pm.tenCateLarge = i.cate_large and Pm.tenCateMid = i.cate_mid and Pm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_outMall.dbo.tbl_cjmall_regitem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_AppWish.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_outMall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and c.infodiv = PD.infodiv "
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'�ö��/ȭ�����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
'		strSql = strSql & " and i.sellcash=i.orgprice"              '''��а� ���� ���ϴ°͸�.. // ���ݼ��� ��� ����..?
'		strSql = strSql & " and (i.orgprice<>0 and ((i.orgprice-i.orgSuplyCash)/i.orgprice)*100>=" & CMAXMARGIN & ")"							'������ ��ǰ ����
		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM db_outMall.dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_outMall.dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_outMall.dbo.tbl_cjmall_regitem WHERE cjmallStatCD >= 3) "	''��ϿϷ��̻��� ��Ͼȵ�.										'�Ե���ϻ�ǰ ����
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & " and Pm.PmapCnt is Not Null "
		strSql = strSql & "		"	& addSql											'ī�װ� ��Ī ��ǰ��
		rsCTget.Open strSql,dbCTget,1
		FResultCount = rsCTget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsCTget.EOF Then
			Do until rsCTget.EOF
				Set FItemList(i) = new cjmallItem
					FItemList(i).Fitemid			= rsCTget("itemid")
					FItemList(i).FtenCateLarge		= rsCTget("cate_large")
					FItemList(i).FtenCateMid		= rsCTget("cate_mid")
					FItemList(i).FtenCateSmall		= rsCTget("cate_small")
					FItemList(i).Fitemname			= db2html(rsCTget("itemname"))
					FItemList(i).FitemDiv			= rsCTget("itemdiv")
					FItemList(i).FsmallImage		= rsCTget("smallImage")
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsCTget("orgSuplyCash")
					FItemList(i).FSellCash			= rsCTget("sellcash")
					FItemList(i).FBuyCash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).FsaleYn			= rsCTget("sailyn")
					FItemList(i).FisUsing			= rsCTget("isusing")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")
					FItemList(i).Fkeywords			= rsCTget("keywords")
					FItemList(i).Fvatinclude        = rsCTget("vatinclude")
					FItemList(i).ForderComment		= db2html(rsCTget("ordercomment"))
					FItemList(i).FoptionCnt			= rsCTget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsCTget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsCTget("makername"))
					FItemList(i).FUsingHTML			= rsCTget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsCTget("itemcontent"))
					FItemList(i).FitemGbnKey        = rsCTget("itemGbnKey")
					FItemList(i).FcjmallStatCD		= rsCTget("cjmallStatCD")
					FItemList(i).FRectMode			= FRectMode
					FItemList(i).Fdeliverfixday		= rsCTget("deliverfixday")
					FItemList(i).Fdeliverytype		= rsCTget("deliverytype")
					FItemList(i).Fsocname_kor		= rsCTget("socname_kor")
					FItemList(i).Fcdmkey			= rsCTget("cdmkey")
					FItemList(i).Fcddkey			= rsCTget("cddkey")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getCjmallEditedItemList
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		ElseIf FRectNotJehyu="Y" Then
			'���޸� ��ǰ�� �ƴѰ�
			addSql = " and i.isExtUsing='N' "
		Else
			'������ ��ǰ��
			''addSql = " and m.LTiMallLastUpdate<i.lastupdate"
		End If

        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & " 	SELECT itemid FROM db_outMall.dbo.tbl_jaehyumall_not_edit_itemid"
        addSql = addSql & " 	WHERE stDt < getdate()"
        addSql = addSql & " 	and edDt > getdate()"
        addSql = addSql & " 	and mallgubun = '"&CMALLNAME&"'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, m.cjmallPrdNo, m.cjmallSellYn, UC.socname_kor "
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_AppWish.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN db_outMall.dbo.tbl_cjmall_regitem as m on i.itemid = m.itemid "
		''If (FRectMatchCateNotCheck<>"on") Then
		IF (FRectMatchCate="Y") THEN '' eastone ���� 2013/09/01
			strSql = strSql & " INNER JOIN  (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt FROM db_outMall.dbo.tbl_cjmall_cate_mapping GROUP BY tenCateLarge, tenCateMid, tenCateSmall ) as cm "
			strSql = strSql & " 	on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
			strSql = strSql & " INNER JOIN  (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as PmapCnt FROM db_outMall.dbo.tbl_cjMall_prdDiv_mapping  GROUP BY tenCateLarge, tenCateMid, tenCateSmall ) as Pm "
			strSql = strSql & " 	on Pm.tenCateLarge = i.cate_large and Pm.tenCateMid = i.cate_mid and Pm.tenCateSmall = i.cate_small "
    	End If
		''strSql = strSql & " 	Join db_item.dbo.tbl_LTiMall_cateGbn_mapping G"
		''strSql = strSql & " 		on G.tenCateLarge=i.cate_large and G.tenCateMid=i.cate_mid and G.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_AppWish.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " WHERE 1 = 1 " & addSql
		strSql = strSql & " and m.cjmallPrdNo is Not Null "									'#��� ��ǰ��

'rw strSql
		rsCTget.Open strSql,dbCTget,1
		FResultCount = rsCTget.RecordCount
		Redim preserve FItemList(FResultCount)
		i=0
		If not rsCTget.EOF Then
			Do until rsCTget.EOF
				Set FItemList(i) = new CjmallItem
					FItemList(i).Fitemid			= rsCTget("itemid")
					FItemList(i).FtenCateLarge		= rsCTget("cate_large")
					FItemList(i).FtenCateMid		= rsCTget("cate_mid")
					FItemList(i).FtenCateSmall		= rsCTget("cate_small")
					FItemList(i).Fitemname			= db2html(rsCTget("itemname"))
					FItemList(i).FitemDiv			= rsCTget("itemdiv")
					FItemList(i).FsmallImage		= rsCTget("smallImage")
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsCTget("orgSuplyCash")
					FItemList(i).FSellCash			= rsCTget("sellcash")
					FItemList(i).FBuyCash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).FsaleYn			= rsCTget("sailyn")
					FItemList(i).FisUsing			= rsCTget("isusing")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")
					FItemList(i).Fkeywords			= rsCTget("keywords")
					FItemList(i).ForderComment		= db2html(rsCTget("ordercomment"))
					FItemList(i).FoptionCnt			= rsCTget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsCTget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsCTget("makername"))
					FItemList(i).FUsingHTML			= rsCTget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsCTget("itemcontent"))
					FItemList(i).FcjmallPrdNo		= rsCTget("cjmallPrdNo")
					FItemList(i).FcjmallSellYn		= rsCTget("cjmallSellYn")
	                FItemList(i).Fvatinclude        = rsCTget("vatinclude")
	                FItemList(i).Fsocname_kor		= rsCTget("socname_kor")
					i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub
End Class

'// ��ǰ�̹��� ���翩�� �˻�
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function
%>