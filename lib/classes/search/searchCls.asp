<%
''CS_EUCKR <=> CS_UTF8
''group by �ʵ�� ���� ���� ����.
'''--------------------------------------------------------------------------------------
DIM G_KSCOLORCD : G_KSCOLORCD = Array("023","001","002","010","021","003","004","024","019","005","016","006","007","020","008","018","017","009","011","012","022","013","014","015","025","026","027","028","029","030","031")
DIM G_KSCOLORNM : G_KSCOLORNM = Array("����","����","��Ȳ","����","ī��","���","������","���̺���","īŰ","�ʷ�","��Ʈ","���Ķ�","�Ķ�","���̺�","����","������","���̺���ũ","��ũ","���","����ȸ��","£��ȸ��","����","����","�ݻ�","üũ","��Ʈ������","��Ʈ","�ö��","�����","�ִϸ�","������")

Dim G_KSSTYLECD : G_KSSTYLECD = Array("010","020","030","040","050","060","070","080","090")
Dim G_KSSTYLENM : G_KSSTYLENM = Array("Ŭ����","ťƼ","���","���","���߷�","������Ż","��","�θ�ƽ","��Ƽ��")

DIM G_ORGSCH_ADDR
DIM G_1STSCH_ADDR
DIM G_2NDSCH_ADDR
DIM G_3RDSCH_ADDR
Dim G_4THSCH_ADDR

DIM G_SCH_TIME : G_SCH_TIME=formatdatetime(now(),4)

IF (application("Svr_Info") = "Dev") THEN
     G_1STSCH_ADDR = "192.168.50.10"  ''"110.93.128.109" ''
     G_2NDSCH_ADDR = "192.168.50.10"
     G_3RDSCH_ADDR = "192.168.50.10"
     G_4THSCH_ADDR = "192.168.50.10"
     G_ORGSCH_ADDR = "192.168.50.10"
ELSE
     G_1STSCH_ADDR = "192.168.0.210"        ''192.168.0.210  :: �˻�������(search.asp)   '
     G_2NDSCH_ADDR = "192.168.0.207"        ''192.168.0.207  :: ī�װ�, ��ǰ, �귣��   ''���� �߻��� 208 ��.
     G_3RDSCH_ADDR = "192.168.0.209"        ''192.168.0.209  :: GiftPlus , scn_dt_itemDispColor :: Ȯ��.
     G_4THSCH_ADDR = "192.168.0.208"        ''192.168.0.208  :: mobile 6:10�п� �ε��� ���� ī��
     G_ORGSCH_ADDR = "192.168.0.206"        ''192.168.0.206
END IF

''sample in doc
function escapeQuery( istr )
	dim ret, c, i
	ret = ""
	For i=1 To Len(istr)
		c = Mid(istr,i,1)
		select case c
		case "\"
			ret = ret & "\\"
		case "'"
			ret = ret & "\'"
		case chr(34)
			ret = ret & "\" & chr(34)
		case "*"
			ret = ret & "\*"
		case "("
			ret = ret & "\("
		case ")"
			ret = ret & "\)"
		case else
			ret = ret & c
		end select
	Next
	escapeQuery = ret
end function

function getTimeChkAddr(defaultAddr)
    '''6��10�� 1���� �ε��� �� 2�������� Copy
    '''6��50��~ 1��=>3�������� Copy
    getTimeChkAddr = defaultAddr

    IF (defaultAddr=G_4THSCH_ADDR) THEN
        IF (G_SCH_TIME>"06:00:00") and (G_SCH_TIME<"06:40:00") then
            getTimeChkAddr = G_2NDSCH_ADDR
        END IF
    ELSE
        IF (G_SCH_TIME>"06:40:00") and (G_SCH_TIME<"07:00:00") then
            getTimeChkAddr = G_4THSCH_ADDR
        END IF
    END IF
end function

function debugQuery(iDocruzer,Scn,iSearchQuery,iSortQuery,iFTotalCount,iFResultcount)
  exit function
    IF Not (application("Svr_Info")="Dev") THEN
        exit function
    ENd IF

    dim itime
    Call iDocruzer.GetResult_SearchTime(itime) '�ҿ�ð�
    rw "-------------------------------"
    rw Scn
    rw iSearchQuery
    rw iSortQuery
    rw "FTotalCount:"&iFTotalCount
    rw "FResultcount:"&iFResultcount
    rw "GetResult_SearchTime:"&itime
end function
'''--------------------------------------------------------------------------------------
''2015
'' StringList �˻� �����  �÷��˻� ����
function getDocArrMatchCodeVal(iRectArr,iResultCdArr,iResultValArr,byref retMatchCd,byref retMatchVal)
    dim findvalArr : findvalArr = split(trim(iRectArr),",")
    dim rsltCd : rsltCd = split(iResultCdArr," ")
    dim rsltVal : rsltVal = split(iResultValArr," ")
    dim findval, i,j

    for i=LBound(findvalArr) to Ubound(findvalArr)
        findval = findvalArr(i)
        for j=LBound(rsltCd) to Ubound(rsltCd)
            if (rsltCd(j)=findval) then
                retMatchCd  = rsltCd(j)
                retMatchVal = rsltVal(j)
                exit for
            end if
        next
    next
end function

function getCdPosVal(iCD,icdArr,iValArr)
    dim i,iPos
    getCdPosVal = ""
    if Not isArray(icdArr) then Exit function
    if Not isArray(iValArr) then Exit function
    if UBound(icdArr)<>UBound(iValArr) then Exit function

    iPos = -1
    for i=LBound(icdArr) to UBound(icdArr)
        if (iCD=icdArr(i)) then
            iPos = i
            Exit For
        end if
    next
    if iPos<0 then Exit function
    getCdPosVal=iValArr(iPos)
end function


Class SearchGroupByItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC FImageSmall
	PUBLIC FSubTotal

	PUBLIC FCateCode
	PUBLIC FCateName
	PUBLIC FCateCd1
	PUBLIC FCateCd2
	PUBLIC FCateCd3
	PUBLIC FCateDepth

	PUBLIC FcolorCode
	PUBLIC FcolorName
	PUBLIC FcolorIcon

	PUBLIC FStyleCd
	PUBLIC FStyleName

	PUBLIC FAttribCd
	PUBLIC FAttribName

	PUBLIC FminPrice
	PUBLIC FmaxPrice

End Class

Class SearchItemCls

	Private SUB Class_initialize()
        ''�⺻ 1�� ����.------------------------
		SvrAddr = getTimeChkAddr(G_1STSCH_ADDR)
		''--------------------------------------

		SvrPort = "6167"'DocSvrPort

		AuthCode = "" '������

		Logs = "" '�αװ�

		FResultCount = 0
		FTotalCount = 0
		FPageSize = 10
		FCurrPage = 1
		FPageSize = 30
		FRectColsSize =5
		FLogsAccept = false

	End SUB

	Private SUB Class_Terminate()

	End SUB

	dim FItemList
	dim FPageSize
	dim FCurrPage
	dim FScrollCount
	dim FResultCount
	dim FTotalCount
	dim FTotalPage

	dim FRectItemid			'��ǰ�ڵ�	'����2017-03-28
	dim FRectSearchTxt		'�˻���
	dim FRectSearchItemDiv	'ī�װ� �˻� ���� (y:�⺻ ī�װ���, n:�߰� ī�װ� ����)
	dim FRectSearchCateDep	'ī�װ� �˻� ���� (X:�ش� ī�װ���, T:���� ī�װ� ����)
	dim FRectPrevSearchTxt	'���� �˻���
	dim FRectExceptText		'���ܾ�
	dim FRectSortMethod		'���Ĺ�� (ne:�Ż�ǰ, be:�α��ǰ, lp:��������, hp:��������, hs:���η�, br:��ǰ�ı�, ws:���ü�)
	dim FRectSearchFlag 	'�˻����� (sc:��������, ea:������ü, ep:�������, ne:�Ż�ǰ, fv:���û�ǰ, pk:���弭��)

	dim FRectMakerid		'��ü ���̵�
	dim FRectCateCode		'ī�װ��ڵ�
	dim FListDiv			'ī�װ�/�˻� ���п�
	dim FSellScope			'�ǸŰ��� ��ǰ�˻� ����
	dim FGroupScope			'�˻��� �׷��� ���� (1:1depth, 2:2depth, 3:3depth)
	dim FdeliType			'��۹�� (FD:������, TN:�ٹ����� ���, FT:����+�ٹ����� ���, WD:�ؿܹ��)

	dim FcolorCode			'��ǰ�÷�Ĩ
	dim FstyleCd			'��ǰ��Ÿ��
	dim FattribCd			'��ǰ�Ӽ�

	dim FminPrice			'�����ּҰ�
	dim FmaxPrice			'�����ִ밪
	dim FSalePercentHigh	'������ �ִ밪
	dim FSalePercentLow		'������ �ּҰ�

	dim FCheckResearch 		'����� ��˻� üũ��
	dim FRectColsSize		'��� ����Ʈ ����
	dim FLogsAccept			'�߰� �α� ���� ����

	dim FarrCate			'���� ī�װ�
	dim FisTenOnly			'�ٹ����� �����ǰ
	dim FisLimit			'�����ǸŻ�ǰ
	dim FisFreeBeasong

	Private SvrAddr
	Private SvrPort
	Private AuthCode
	Private Logs
	Private Scn
	private strQuery
	Private Order
	Private StartNum

	Private SearchQuery
	Private SortQuery

    public FRectSynonymAssign
	public FRectExpandSearch
	public FRectIdxrect

	public function getRetSearchQuery
		getRetSearchQuery = SearchQuery
	end function

	public function getRetSortQuery
		getRetSortQuery = SortQuery
	end function

    public function InitDocruzer(iDocruzer)
        InitDocruzer = FALSE
        IF ( iDocruzer.BeginSession() < 0 ) THEN
			EXIT function
		End If

        IF NOT DocSetOption(iDocruzer) THEN
			EXIT function
		End If
		InitDocruzer = TRUE
    End function

    public function DocSetOption(iDocruzer)
        dim ret
        ret = iDocruzer.SetOption(iDocruzer.OPTION_REQUEST_CHARSET_UTF8,1)
        DocSetOption = (ret>=0)
    end function

	public function GetLevelUpCount()

		if (FCurrRank<FLastRank) then
			GetLevelUpCount = CStr(FLastRank-FCurrRank)
		elseif (FCurrRank=FLastRank) and (FLastRank=0) then
			GetLevelUpCount = ""
		elseif (FCurrRank=FLastRank) then
			GetLevelUpCount = ""
		elseif (FCurrRank>FLastRank) and (FLastRank=0) then
			GetLevelUpCount = ""
		else
			GetLevelUpCount = CStr(FCurrRank-FLastRank)
			if FCurrRank-FLastRank>=FCurrPos then
				GetLevelUpCount = ""
			end if
		end if
	end function

	public function GetLevelUpArrow()
		if (FCurrRank<FLastRank) then
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2013/award/ico_rank_up.gif' alt='���� ���' /> " & GetLevelUpCount()
		elseif (FCurrRank=FLastRank) and (FLastRank=0) then
			GetLevelUpArrow = ""
		elseif (FCurrRank=FLastRank) then
			'GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2010/bestaward/ico_none.gif' align='absmiddle' style='display:inline;'> <font class='eng11px00'><b>0</b></font>"
			GetLevelUpArrow = ""
		elseif (FCurrRank>FLastRank) and (FLastRank=0) then
			GetLevelUpArrow = ""
		else
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2013/award/ico_rank_down.gif' alt='���� �϶�' /> " & GetLevelUpCount()
			if FCurrRank-FLastRank>=FCurrPos then
				'GetLevelUpArrow = "<font class='eng11px00'><b>0</b></font>"
				GetLevelUpArrow = ""
			end if
		end if
	end function

	''/�˻� ���� ����
	FUNCTION getSearchQuery(byref query)
		dim strQue, arrCCD, arrSCD, arrACD, lp

		'### �˻����п� ���� �⺻�� Ȯ�� �� ���� ###
		Select Case FListDiv
			Case "search"
				'�˻� ������ ���
				IF (FRectSearchTxt="" or isNull(FRectSearchTxt)) Then EXIT FUNCTION
			Case "list"
				'ī�װ� ���
				IF (FRectCateCode="" or isNull(FRectCateCode)) Then EXIT FUNCTION
			Case "colorlist"
				'�÷� �˻� ���
				if (FcolorCode="" and FcolorCode="0") Then EXIT FUNCTION
			Case "brand"
				'�귣�� ��ǰ ���
				IF (FRectMakerid="" or isNull(FRectMakerid)) Then EXIT FUNCTION
				FRectSearchItemDiv = "y"
			Case "salelist"
				'���λ�ǰ ���
				FRectSearchFlag = "sc"
			Case "newlist"
				'�Ż�ǰ ���
				FRectSearchFlag = "ne"
				IF FRectCateCode="" Then FRectSearchItemDiv = "y"
			Case "bestlist"
				'����Ʈ��ǰ ���
				FRectSearchItemDiv = "y"
			Case "aboard"
				'�ؿ��Ǹ� ��ǰ ���
				FdeliType = "WD"
			Case "fulllist"
				'ī�װ����� ��ü.
			Case Else
				EXIT FUNCTION
		End Select

		'### �˻����� ���� ###

		if (FRectIdxrect="") then FRectIdxrect="idx_itemname" ''�űԹ��

		'@ �˻���(Ű����)
		If FRectSearchTxt<>"" Then
			FRectSearchTxt = chgCoinedKeyword(FRectSearchTxt)
			FRectSearchTxt = escapeQuery(FRectSearchTxt)  ''2015 �߰�
			IF FRectExceptText<>"" Then
			    FRectExceptText = escapeQuery(FRectExceptText)  ''2015 �߰�
				strQue = getQrCon(strQue) & "("&FRectIdxrect&"='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'���ܾ�
			else
				if (FRectExpandSearch="allwordOradjacent") then
					strQue = " aliasing "&FRectIdxrect&" as name1, "&FRectIdxrect&" as name2 (name1='"&FRectSearchTxt&"' alladjacent "&CHKIIF(FRectSynonymAssign<>"","synonym","")&" or name2='"&FRectSearchTxt&"' allword "&CHKIIF(FRectSynonymAssign<>"","synonym","")&") "
				else
					strQue = getQrCon(strQue) & ""&FRectIdxrect&"='" & FRectSearchTxt & "'  "	
					if (FRectExpandSearch<>"") then
						strQue = strQue & " "&FRectExpandSearch&" "
					end if

					IF (FRectExpandSearch<>"") and (FRectSynonymAssign<>"") then
						strQue = strQue & " synonym "	'Ű����˻�(���Ǿ� ���� => synonym )
					end if
					'strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  natural "		'�ڿ��� �˻�(���Ǿ� ����) synonym
				end if
			End if
		End If
'		If FRectSearch <> "" AND FRectSearchTxt<>"" Then
'			If (FRectSearch = "keyword") OR (FRectSearch = "itemname") Then
'				FRectSearchTxt = chgCoinedKeyword(FRectSearchTxt)
'				FRectSearchTxt = escapeQuery(FRectSearchTxt)  ''2015 �߰�
'				IF FRectExceptText<>"" Then
'				    FRectExceptText = escapeQuery(FRectExceptText)  ''2015 �߰�
'					strQue = getQrCon(strQue) & "(idx_itemname='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'���ܾ�
'				else
'					strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  allword "	'Ű����˻�(���Ǿ� ����) synonym
'					'strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  natural "		'�ڿ��� �˻�(���Ǿ� ����) synonym
'				End if
'			ElseIf FRectSearch = "itemid" Then
'				strQue = getQrCon(strQue) & "idx_itemid=" & FRectSearchTxt
'				'strQue = getQrCon(strQue) & "idx_itemid=279397 or idx_itemid=145983 "
'			End If
'		End If

		'@ ī�װ� �˻� ���� idx_isDefault ����
		''IF FRectSearchItemDiv="y" Then
		''	''strQue = strQue & getQrCon(strQue) & "idx_isDefault='y' "
		''END IF

		'@ ī�װ�
		IF FRectCateCode<>"" Then
			if FRectSearchCateDep="X" then
				strQue = strQue & getQrCon(strQue) & "idx_catecode='" & FRectCateCode & "'"
			else
			    IF FRectSearchItemDiv="y" Then ''�⺻ī�װ�
			        strQue = strQue & getQrCon(strQue) & "idx_catecode like '" & FRectCateCode & "*'"
			    else                           ''�߰�ī�װ˻�
			        strQue = strQue & getQrCon(strQue) & "idx_catecodeExt like '" & FRectCateCode & "*'"
			    end if
			end if
		END IF

		'@ ���� ī�װ�
		IF FarrCate<>"" THEN
			dim arrCt, lpCt
			if right(FarrCate,1)="," then FarrCate=left(FarrCate,len(FarrCate)-1)
			arrCt = split(FarrCate,",")
			strQue = strQue & getQrCon(strQue) & "("
			for lpCt=0 to ubound(arrCt)
				if FRectSearchCateDep="X" then
					strQue = strQue & " idx_catecode='" & RequestCheckVar(LCase(trim(arrCt(lpCt))),18) & "' "
				else
					strQue = strQue & " idx_catecode like '" & RequestCheckVar(LCase(trim(arrCt(lpCt))),18) & "*' "
				end if
				if lpCt<ubound(arrCt) then strQue = strQue & " or "
			next
			strQue = strQue & " )"
		END IF

		'@ �˻�����
		IF FRectSearchFlag<>"" THEN
			Select Case FRectSearchFlag
				Case "sc"	'��������
					strQue= strQue & getQrCon(strQue) & "(idx_saleyn='Y' or idx_itemcouponyn='Y') "
				Case "ea"	'��ü����
					strQue= strQue & getQrCon(strQue) & "(idx_evalcnt>0) "
				Case "ep"	'�������
					strQue= strQue & getQrCon(strQue) & "(idx_evalcntPhoto>0) "
				Case "ne"	'�Ż�ǰ
					strQue = strQue & getQrCon(strQue) & "idx_newyn='Y' "
				Case "fv"	'���û�ǰ
					strQue = strQue & getQrCon(strQue) & "(idx_favcount>0) "
				Case "pk"	'���弭��
					strQue = strQue & getQrCon(strQue) & "idx_pojangok='Y' "
			End Select
		END IF

		'@ �귣��
		IF FRectMakerid<>"" THEN
			dim arrMkr, lpMkr
			if right(FRectMakerid,1)="," then FRectMakerid=left(FRectMakerid,len(FRectMakerid)-1)
			arrMkr = split(FRectMakerid,",")
			strQue = strQue & getQrCon(strQue) & "("
			for lpMkr=0 to ubound(arrMkr)
				strQue = strQue & " idx_makerid='" & RequestCheckVar(LCase(trim(arrMkr(lpMkr))),32) & "'  "
				if lpMkr<ubound(arrMkr) then strQue = strQue & " or "
			next
			strQue = strQue & " )"
		END IF

		'@ ���ݹ���
		if FminPrice<>"" then
			strQue = strQue & getQrCon(strQue) & "idx_sellcash>='" & FminPrice & "'"
		end if
		if FmaxPrice<>"" then
			strQue = strQue & getQrCon(strQue) & "idx_sellcash<='" & FmaxPrice & "'"
		end if

		'@ ���ι���
		IF FSalePercentHigh<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_salepercent >=" & (1-FSalePercentHigh)*100 & " "
		End IF
		IF FSalePercentLow<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_salepercent <" & (1-FSalePercentLow)*100 & " "
		End IF

		'@ ��۹��
		Select Case FdeliType
			Case "FD"	'������
				strQue = strQue & getQrCon(strQue) & "isFreeBeasong='Y'"
			Case "TN"	'�ٹ����ٹ��
				strQue = strQue & getQrCon(strQue) & "(deliverytype='1' or deliverytype='4')"
			Case "FT"	'���� + �ٹ����ٹ��
				strQue = strQue & getQrCon(strQue) & "(deliverytype='1' or deliverytype='4') and isFreeBeasong='Y'"
			Case "WD"	'�ؿܹ��
				strQue = strQue & getQrCon(strQue) & "isAboard='Y'"
		end Select

		'@ �ٹ����� �����ǰ
		IF FisTenOnly="only" Then
			strQue = strQue & getQrCon(strQue) & "idx_tenOnlyYn='Y' "
		End IF

		'@ ������ǰ
		IF FisLimit="limit" Then
			strQue = strQue & getQrCon(strQue) & "idx_limityn='Y' "
		End IF

		'@ �����ۻ�ǰ
		IF FisFreeBeasong="free" Then
			strQue = strQue & getQrCon(strQue) & "idx_isFreeBeasong='Y' "
		End If

		'@ �÷��˻�
		if Not(FcolorCode="" or isNull(FcolorCode) or FcolorCode="0") then
			arrCCD = split(FcolorCode,",")
			if ubound(arrCCD)>0 then
				'���� �÷��ڵ� ����
				strQue = strQue & getQrCon(strQue) & "(idx_colorCd='"&replace(FcolorCode,","," ")&"' anystring)"  ''2015 ����
'				strQue = strQue & getQrCon(strQue) & "("
'				for lp=0 to ubound(arrCCD)
'					if getNumeric(arrCCD(lp))<>"" then
'						if lp>0 then strQue = strQue & " or "
'						strQue = strQue & "idx_colorCd='" & getNumeric(arrCCD(lp)) & "'"
'					end if
'				next
'				strQue = strQue & ")"
			else
				'���� �÷��ڵ� ����
				strQue = strQue & getQrCon(strQue) & "idx_colorCd='" & getNumeric(arrCCD(0)) & "'"
			end if
		end if

		'@ ��Ÿ�� ���� ����
		if Not(FstyleCd="" or isNull(FstyleCd)) then
			arrSCD = split(FstyleCd,",")
			if ubound(arrSCD)>0 then
				'���� ��Ÿ���ڵ� ����
				strQue = strQue & getQrCon(strQue) & "(idx_styleCd='"&replace(FstyleCd,","," ")&"' anystring)"  ''2015 ����
'				strQue = strQue & getQrCon(strQue) & "("
'				for lp=0 to ubound(arrSCD)
'					if getNumeric(arrSCD(lp))<>"" then
'						if lp>0 then strQue = strQue & " or "
'						strQue = strQue & "idx_styleCd='" & getNumeric(arrSCD(lp)) & "'"
'					end if
'				next
'				strQue = strQue & ")"
			else
				'���� ��Ÿ���ڵ� ����
				strQue = strQue & getQrCon(strQue) & "idx_styleCd='" & getNumeric(arrSCD(0)) & "'"
			end if
		end if

		'@ ��ǰ�Ӽ� ���� ����
		if Not(FattribCd="" or isNull(FattribCd)) then
			arrACD = split(FattribCd,",")
			if ubound(arrACD)>0 then
				'���� �Ӽ��ڵ� ����
				strQue = strQue & getQrCon(strQue) & "(idx_attribCd='"&replace(FattribCd,","," ")&"' anystring)"  ''2015 ����
'				strQue = strQue & getQrCon(strQue) & "("
'				for lp=0 to ubound(arrACD)
'					if getNumeric(arrACD(lp))<>"" then
'						if lp>0 then strQue = strQue & " or "
'						strQue = strQue & "idx_attribCd='" & getNumeric(arrACD(lp)) & "'"
'					end if
'				next
'				strQue = strQue & ")"
			else
				'���� �Ӽ��ڵ� ����
				strQue = strQue & getQrCon(strQue) & "idx_attribCd='" & getNumeric(arrACD(0)) & "'"
			end if
		end if

		'@ ��ǰ �Ǹ� ����
'		IF FSellScope="Y" Then
'			strQue = strQue & getQrCon(strQue) & "idx_sellyn='Y' "
'		ELSE
'			strQue = strQue & getQrCon(strQue) & "(idx_sellyn='Y' or idx_sellyn='S') "
'		End IF

		If FSellScope <> "" Then
			strQue = strQue & getQrCon(strQue) & "idx_sellyn='"& FSellScope &"' "
		Else
			strQue = strQue & getQrCon(strQue) & "(idx_sellyn='Y' or idx_sellyn='S' or idx_sellyn='N')  "
		End If

        ''2015 �߰� string list group by ���� �ش��ʵ忡 �ΰ��� �ִ°�� �� ��� 000 ����.
        IF scn="scn_dt_itemDispStyleGroup" then
            strQue = strQue & getQrCon(strQue) & "idx_styleCd!='000' "
        ELSEIF  scn="scn_dt_itemDispAttribGroup" then
            strQue = strQue & getQrCon(strQue) & "idx_attribgrp!='000' "
        ELSEIF  scn="scn_dt_itemDispColorGroup" then
            strQue = strQue & getQrCon(strQue) & "idx_colorgrp!='000' "
        ELSEIF  scn="scn_dt_itemDispCateGroup" then
            if (FGroupScope="2") then
                strQue = strQue & getQrCon(strQue) & "idx_cd2grp!='000' "
            elseif (FGroupScope="3") then
                strQue = strQue & getQrCon(strQue) & "idx_cd3grp!='000' "
            end if
        END IF

		' ���ܰ˻��� �߰� : ������ ���� �ڿ� ��ġ���־���Ѵ�. �� �ڷ� �� �߰����� ���� ������ �߰��Ұ�
		IF FRectSearchTxt<>"" and (scn="scn_dt_itemDisp" or scn="scn_dt_itemDispCateGroup"  or scn="scn_dt_itemDispBrandGroup"  or scn="scn_dt_itemDispColorGroup"  or scn="scn_dt_itemDispStyleGroup" or scn="scn_dt_itemDispAttribGroup"  or scn="") then 
			strQue = strQue & "exclude by keylist(0) "
		end if

		query = strQue
	
'rw strQue
	End FUNCTION

	Sub getSortQuery(byref query)
		dim strQue

		'// �ߺ� ��ǰ ����(�ߺ� ��� ī�װ��ϰ��) 2015 ����
		'IF (FRectCateCode<>"" and FRectSearchItemDiv<>"y") Then '' �߰� ī�װ� �˻���
    	'	'strQue = " GROUP BY itemid"
    	'END IF

		'// ����
		IF FRectSortMethod="ne" THEN '�Ż�ǰ
			strQue = strQue & " ORDER BY itemid desc"
		ELSEIF FRectSortMethod="be" THEN '�α��ǰ
			if FRectSearchFlag="fv" then
				'���û�ǰ����� ������ ���ü�����
				strQue = strQue & " ORDER BY favcount desc,itemscore desc,itemid desc"
			else
				strQue = strQue & " ORDER BY $MATCHFIELD(cateboostkeylist,bestkeylist) desc,itemscore desc,itemid desc"
			end if
		ELSEIF FRectSortMethod="lp" THEN '��������
			strQue = strQue & " ORDER BY sellcash "
		ELSEIF FRectSortMethod="hp" THEN'��������
			strQue = strQue & " ORDER BY sellcash desc"
		ELSEIF FRectSortMethod="hs" THEN '�ּ��� (�������� ������)
			strQue = strQue & " ORDER BY salepercent desc, saleprice desc"
		ELSEIF FRectSortMethod="br" THEN '����Ʈ�ı��
			strQue = strQue & " ORDER BY evalcnt desc,itemid desc"
		ELSEIF FRectSortMethod="ws" THEN '���ü�
			strQue = strQue & " ORDER BY favcount desc,itemid desc"
		ELSEIF FRectSortMethod="pj" THEN '�α������
			strQue = strQue & " ORDER BY pojangcnt desc,itemid desc"
		ELSEIF FRectSortMethod="bs0" THEN '�Ǹż���(����)
		    strQue = strQue & " ORDER BY sellCnt desc,sellcash desc,itemid desc"
		ELSEIF FRectSortMethod="bs1" THEN '�Ǹż���(�˻���)
			strQue = strQue & " ORDER BY sellCnt desc,$MATCHFIELD(cateboostkeylist,bestkeylist) desc, itemid desc"
		ELSEIF FRectSortMethod="bs2" THEN '�Ǹż���(�˻���)
			strQue = strQue & " ORDER BY sellCnt desc,$MATCHFIELD(cateboostkeylist) desc, itemid desc"
		ELSEIF FRectSortMethod="bs3" THEN 'TEST
			strQue = strQue & " ORDER BY sellCnt desc,$MATCHFIELD(cateboostkeylist,bestkeylist) desc,sellcash desc, itemid desc"
		ELSEIF FRectSortMethod="bs4" THEN 'TEST
			strQue = strQue & " ORDER BY $MATCHFIELD(cateboostkeylist,bestkeylist) desc,sellCnt desc,itemscore desc, itemid desc"
		ELSEIF FRectSortMethod="bs5" THEN 'TEST
			strQue = strQue & " ORDER BY $MATCHFIELD(cateboostkeylist) desc,sellCnt desc,sellcash desc, itemid desc"
		ELSEIF FRectSortMethod="bs6" THEN 'TEST
			strQue = strQue & " ORDER BY $CATEGORYFIELD( recomkeyword(1) seasonboost_groupid(3) categorynamelist(0) bestkeylist(2) makerid(4), '"&replace(TRIM(FRectSearchTxt)," ","")&"') desc, sellCnt desc, itemscore desc, itemid desc"
		ELSEIF FRectSortMethod="be6" THEN 'TEST
			strQue = strQue & " ORDER BY $CATEGORYFIELD( recomkeyword(1) seasonboost_groupid(3) categorynamelist(0) bestkeylist(2) makerid(4) , '"&replace(TRIM(FRectSearchTxt)," ","")&"') desc, itemscore desc, itemid desc"
		ELSEIF FRectSortMethod="bs7" THEN 'TEST
			strQue = strQue & " ORDER BY $CATEGORYFIELD( recomkeyword(1) seasonboost_groupid(3) categorynamelist(0) bestkeylist(2) makerid(4) , '"&replace(TRIM(FRectSearchTxt)," ","")&"') desc,$MATCHFIELD(name1,name2) desc, sellCnt desc, itemscore desc, itemid desc"
		ELSEIF FRectSortMethod="be7" THEN 'TEST
			strQue = strQue & " ORDER BY $CATEGORYFIELD( recomkeyword(1) seasonboost_groupid(3) categorynamelist(0) bestkeylist(2) makerid(4) , '"&replace(TRIM(FRectSearchTxt)," ","")&"') desc,$MATCHFIELD(name1,name2) desc, itemscore desc, itemid desc"
		ELSE
			strQue = strQue & " ORDER BY itemid desc"
		END IF
		query = strQue
		
	End Sub

	Function getQrCon(query)
		if Not(query="" or isNull(query)) then
			getQrCon = " and "
		end if
	End Function

	'// ��ǰ �̹��� ���� ��ȯ(�÷��ڵ� ������ ���� �Ϲ�/�÷�Ĩ ����)
	Function getItemImageUrl()
		IF application("Svr_Info")	= "Dev" THEN
			if FcolorCode="" or FcolorCode="0" then
				getItemImageUrl = "http://webimage.10x10.co.kr/image"
			else
				getItemImageUrl = "http://webimage.10x10.co.kr/color"
			end if
		Else
			if FcolorCode="" or FcolorCode="0" then
				getItemImageUrl = "http://webimage.10x10.co.kr/image"
			else
				getItemImageUrl = "http://webimage.10x10.co.kr/color"
			end if
		End If
	end function

	'####### ��ǰ �˻� - �˻� ���� ######
	PUBLIC SUB getSearchList()

		DIM Scn
		DIM Docruzer,ret

		DIM Logs ,iRows
		DIM arrData,arrSize, retMatchCd, retMatchVal

        if (FPageSize>300) then FPageSize=300  ''2016/11/16�߰�

		'// �˻� ��� ��� �ó�������
		if FcolorCode="" or FcolorCode="0" then
			Scn= "scn_dt_itemDisp"		'�Ϲ� ��ǰ �˻�
		else
			'Scn= "scn_dt_itemColor"		'��ǰ �÷��� �˻�
			'Scn= "scn_dt_itemDispColor"	'��ǰ �÷��� �˻�(����ī�װ�)
			Scn= "scn_dt_itemDisp"		    '�Ϲ� ��ǰ �˻� ���� 2015
		end if

		StartNum = (FCurrPage -1)*FPageSize '// �˻����� Row

		CALL getSearchQuery(SearchQuery)	'// �˻� ��������
		CALL getSortQuery(SortQuery)		'// ���� ���� ����
		''Response.Write SearchQuery &"<Br>"
		IF SearchQuery="" THEN
			EXIT SUB
		END IF

		IF (FLogsAccept) and (FRectSearchTxt<>"") and (FCurrPage="1") THEN

            'Logs = "��ǰ+^" & FRectSearchTxt & "]##" & FRectSearchTxt & "||" & FRectPrevSearchTxt  	'// �αװ�

            ''2015 search4
            '�⺻:[����Ʈ@ī�װ�+�����$�����ڵ�|����|�˻���Ÿ��(����)|ù�˻�|��������ȣ|���ļ�^�����˻���##�˻���] ''�⺻
            Dim iLOG_SITE : iLOG_SITE = "WEB"
            Dim iLOG_CATE : iLOG_CATE = "RECT"
            Dim iLOG_USER : iLOG_USER = GetUserLevelStr(GetLoginUserLevel) '' ȸ������� ���
            Dim iLOG_SEX  : iLOG_SEX  = "" '' 0��α���,1����,2����
            Dim iLOG_AGE  : iLOG_AGE  = "" '' 0��α���,1:10��,2:20��,3:30��,4:40��,5:50��
            Dim iLOG_STYPE : iLOG_STYPE = "" '' ���� ������ X
            Dim iLOG_FIRST : iLOG_FIRST = "" '' ù�˻�/��˻� ������ X  FCheckResearch

            Logs = iLOG_SITE&"@"                ''[ @
            Logs = Logs&iLOG_CATE&"+"           ''@ +
            Logs = Logs&iLOG_USER&"$"           ''+ $
            Logs = Logs&iLOG_SEX&"|"            ''$ |
            Logs = Logs&iLOG_AGE&"|"            ''| |
            Logs = Logs&iLOG_STYPE&"|"          ''| |
            Logs = Logs&iLOG_FIRST&"|"          ''| |
            Logs = Logs&FCurrPage&"|"           ''| |
            Logs = Logs&FRectSortMethod&"^"     ''| ^
            Logs = Logs&FRectPrevSearchTxt&"##" ''^ ##
            Logs = Logs&FRectSearchTxt          ''## ]


		END IF

        ''��ǰ�˻�/�귣��˻��� �ƴѰ�� 2��������.
        ''---------------------------------------------------------------------------------------------------------
        if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
            'response.write "2������<br>"
             SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR) ''G_4THSCH_ADDR ''G_2NDSCH_ADDR
        end if
        ''Į��Ĩ�˻�/�귣��˻� 3��
        if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then
        	'response.write "3������<br>"
        	SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
        end if
        ''---------------------------------------------------------------------------------------------------------

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		IF( ret < 0 ) THEN
		    rw "err1"
			dbget.execute "EXECUTE db_log.dbo.sp_Ten_DocLog @ErrMsg ='"& html2db(SearchQuery) & "[" & html2db(Docruzer.GetErrorMessage()) &"]'"

			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING

			IF FListDiv<>"search" THEN
				'// 1�� ���� ������ 2������ ����(2���� ������ Skip)
				if (SvrAddr = G_1STSCH_ADDR) then
					SvrAddr = G_2NDSCH_ADDR  ''"192.168.0.108"
					if (G_1STSCH_ADDR<>G_2NDSCH_ADDR) then  ''�߰� 2013/09
					    call getSearchList()
				    end if
				end if
			END IF

			EXIT SUB
		END IF

		Call Docruzer.GetResult_TotalCount(FTotalCount) '�˻���� �� ��
		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		'Response.write "�˻������ : " & FTotalCount & "<br>"
		IF( FResultCount <= 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB 'Response.write "GetResult_RowSize: " & Docruzer.GetErrorMessage()
		END IF

		FTotalPage =  Cdbl(FTotalCount\FPageSize)
		IF  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) THEN
			FtotalPage = FtotalPage +1
		END IF

		REDIM FItemList(FResultCount)

		FOR iRows=0 to FResultCount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.GetErrorMessage()
				EXIT FOR
			END IF

			SET FItemList(iRows) = NEW CCategoryPrdItem
				FItemList(iRows).FCateCode = arrData(0)
				FItemList(iRows).FarrCateCd = arrData(2)
				FItemList(iRows).FItemDiv	= arrData(3)
				FItemList(iRows).FItemid = arrData(4)
				FItemList(iRows).FItemName = db2html(arrData(5))
				FItemList(iRows).FKeyWords = db2html(arrData(6))
				FItemList(iRows).FSellCash = arrData(7)
				FItemList(iRows).FOrgPrice = arrData(8)
				FItemList(iRows).FMakerId = arrData(9)
				FItemList(iRows).FBrandName = db2html(arrData(10))
				if (FcolorCode="" or FcolorCode="0") then
				    FItemList(iRows).FImageBasic 	= getItemImageUrl & "/basic/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(11))
    				FItemList(iRows).FImageMask 	= getItemImageUrl & "/mask/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(12))
    				FItemList(iRows).FImageList 	= getItemImageUrl & "/list/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(13))
    				FItemList(iRows).FImageList120 	= getItemImageUrl & "/list120/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(14))
    				FItemList(iRows).FImageIcon1 	= getItemImageUrl & "/icon1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(15))
    				FItemList(iRows).FImageIcon2 	= getItemImageUrl & "/icon2/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(16))
    				FItemList(iRows).FImageSmall	= getItemImageUrl & "/small/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(17))
				else
				    retMatchCd  = ""
				    retMatchVal = ""

				    Call getDocArrMatchCodeVal(FcolorCode,arrData(35),arrData(44),retMatchCd,retMatchVal)

				    FItemList(iRows).FImageBasic 	= getItemImageUrl & "/basic/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &retMatchVal
    				FItemList(iRows).FImageMask 	= FItemList(iRows).FImageBasic
    				FItemList(iRows).FImageList 	= getThumbImgFromURL(FItemList(iRows).FImageBasic,100,100,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,100,100,85)
    				FItemList(iRows).FImageList120 	= getThumbImgFromURL(FItemList(iRows).FImageBasic,120,120,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,120,120,85)
    				FItemList(iRows).FImageIcon1 	= getThumbImgFromURL(FItemList(iRows).FImageBasic,200,200,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,200,200,85)
    				FItemList(iRows).FImageIcon2 	= getThumbImgFromURL(FItemList(iRows).FImageBasic,150,150,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,150,150,85)
    				FItemList(iRows).FImageSmall	= getThumbImgFromURL(FItemList(iRows).FImageBasic,50,50,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,50,50,85)
    			end if
				if (arrData(47)<>"") then ''2015 �߰� (�߰� �̹���)
				    if (FcolorCode="" or FcolorCode="0") then
				    	FItemList(iRows).FAddimage      = getItemImageUrl & "/add1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" & db2html(arrData(47))
				    else
				    	FItemList(iRows).FAddimage      = replace(getItemImageUrl,"/color","/image") & "/add1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" & db2html(arrData(47))
				    end if
			    elseif (arrData(12)<>"") then	'�߰��̹����� ���� ����ŷ�� �ִ� ��� ����ŷ�� �߰��̹����� ó��(2015.09.07; ������)
			    	FItemList(iRows).FAddimage      = FItemList(iRows).FImageMask
			    end if
				FItemList(iRows).FSellyn = arrData(18)
				FItemList(iRows).FSaleyn = arrData(19)
				FItemList(iRows).FLimityn = arrData(20)
				FItemList(iRows).FRegdate = dateserial(left(arrData(21),4),mid(arrData(21),5,2),mid(arrData(21),7,2))
				IF arrData(22)<>"" Then
					FItemList(iRows).FReipgodate= dateserial(left(arrData(22),4),mid(arrData(22),5,2),mid(arrData(22),7,2))
				End IF
				FItemList(iRows).FItemcouponyn = arrData(23)
				FItemList(iRows).FItemCouponValue = arrData(24)
				FItemList(iRows).FItemCouponType = arrData(25)
				FItemList(iRows).FEvalCnt = arrData(26)
				FItemList(iRows).FEvalcnt_Photo = arrData(27)
				FItemList(iRows).FfavCount = arrData(28)
				FItemList(iRows).FItemScore = arrData(29)
				FItemList(iRows).FtenOnlyYn = arrData(33)

                FItemList(iRows).Frecentsellcount = arrData(48) ''//2015 �߰�
                FItemList(iRows).FPojangOk = arrData(49)		''//2015.10.07
                FItemList(iRows).FAllCateName = arrData(41)		''//2015.10.07

                'FItemList(iRows).FcolorCd = arrData(35)
			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'####### ��ǰ �˻� ī�װ��� ī����  ######
	PUBLIC SUB getGroupbyCategoryList()

		'// �˻� ��� ��� �ó�������
		Scn= "scn_dt_itemDispCateGroup"		'�Ϲ� ��ǰ �˻�

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// �˻����� Row
		call getSearchQuery(SearchQuery)	'// �˻� ��������

        IF SearchQuery="" Then
			EXIT SUB
		End If

		'//�׷� ������ ����(���� ���� ����)
		Select Case FGroupScope
			Case "1"
				SortQuery = " GROUP BY cd1grp order by cd1grp "
			Case "2"
				SortQuery = " GROUP BY cd2grp order by cd2grp "
			Case "3"
				SortQuery = " GROUP BY cd3grp order by cd3grp "
			Case Else
				SortQuery = " GROUP BY cd1grp order by cd1grp "
		end Select

		dim Rowids,Scores
		FTotalCount = 0

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

 'response.write "group : " & SearchQuery & SortQuery & "<br>"
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
		    response.write "ERR:"&Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchGroupByItems
				FItemList(iRows).FCateCode	= arrData(0)
				FItemList(iRows).FCateName	= arrData(1)
				FItemList(iRows).FCateCd1	= arrData(2)
				FItemList(iRows).FCateCd2	= arrData(3)
				FItemList(iRows).FCateCd3	= arrData(4)
				''2015 ����

				if Len(FItemList(iRows).FCateCd1)>=4 then
				    FItemList(iRows).FCateCd1	= Mid(FItemList(iRows).FCateCd1,5,255)              '' sort4+[code3]
				end if

				if Len(FItemList(iRows).FCateCd2)>=11 then
				    FItemList(iRows).FCateCd2	= Mid(FItemList(iRows).FCateCd2,9+3,255)            '' sort4+4+code3+[code3]
				end if

				if Len(FItemList(iRows).FCateCd3)>=18 then
				    FItemList(iRows).FCateCd3	= Mid(FItemList(iRows).FCateCd3,13+3+3,255)         '' sort4+4+4+code3+[code3]
				end if

				FItemList(iRows).FCateDepth	= arrData(5)

                ''rw FItemList(iRows).FCateCd1&"|"&FItemList(iRows).FCateCd2&"|"&FItemList(iRows).FCateCd3&"|"&FItemList(iRows).FCateDepth

				FItemList(iRows).FImageSmall = getItemImageUrl & "/small/" & GetImageSubFolderByItemid(arrData(6)) & "/" &db2html( arrData(7))
				FItemList(iRows).FSubTotal 	= Scores(iRows)
				FTotalCount = FTotalCount + Scores(iRows)
			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'####### ��ǰ �˻� �귣�庰 ī����  ######
	PUBLIC SUB getGroupbyBrandList()

		'// �˻� ��� ��� �ó�������
		''Scn= "scn_dt_itemDisp"		'�Ϲ� ��ǰ �˻�

		Scn= "scn_dt_itemDispBrandGroup"

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// �˻����� Row
		call getSearchQuery(SearchQuery)	'// �˻� ��������

		'//�׷� ������ ����(���� ���� ����)
		'SortQuery = " GROUP BY makerid order by brandname "
		SortQuery = " GROUP BY makerid order by count(*) desc" ''desc $RELEVANCE

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores
		FTotalCount = 0

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

	'response.write "group : " & SearchQuery & SortQuery & "<br>"
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '�˻���� �� ��

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new CCategoryPrdItem
				FItemList(iRows).FMakerID		= arrData(0)
				FItemList(iRows).FBrandName		= arrData(1)
				FItemList(iRows).FImageSmall	= getItemImageUrl & "/small/" & GetImageSubFolderByItemid(arrData(2)) & "/" &db2html( arrData(3))
				FItemList(iRows).FisBestBrand	= arrData(4)
				FItemList(iRows).FItemScore 	= Scores(iRows)
				FItemList(iRows).FCurrRank		= arrData(5)

			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	'####### ��ǰ ��Ÿ�Ϻ� ī����  ######
	PUBLIC SUB getGroupbyStyleList()
		'// �˻� ��� ��� �ó�������
		Scn= "scn_dt_itemDispStyleGroup"		'��ǰ��Ÿ�� �˻�

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim groupKeyCnt,groupKeyVal,groupsize

		dim FseekTime

		StartNum = 0 						'// �˻����� Row
		call getSearchQuery(SearchQuery)	'// �˻� ��������

		'//�׷� ���� ����(���� ���� ����)
		SortQuery = " GROUP BY styleCd order by styleCd "

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores
		FTotalCount = 0

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

		ret = Docruzer.Search(SvrAddr&":"&SvrPort, _
						Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,"",StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
            'response.write Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		ret = Docruzer.GetResult_GroupBy(FResultcount,groupKeyCnt,groupKeyVal,groupsize,100)

		IF( ret < 0 ) THEN
			'Response.write "GetResult_Row: " & Docruzer.msg
			SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
			EXIT Sub
		END IF

		if (groupKeyCnt<1) then
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
		    Exit Sub
		end if

CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)
        dim bufStyle

		FOR iRows = 0 to FResultcount -1
		    bufStyle = groupKeyVal(iRows,0)
			SET FItemList(iRows) = new SearchGroupByItems
			FItemList(iRows).FStyleCd		= bufStyle
	        'FItemList(iRows).FStyleName		= buf

			FItemList(iRows).FSubTotal 	= groupsize(iRows)
			FTotalCount = FTotalCount + groupsize(iRows)
		NEXT

		SET groupKeyVal = NOTHING
        SET groupsize   = NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB



    '####### ��ǰ �Ӽ��� ī����  ######
	PUBLIC SUB getGroupbyAttribList()

		'// �˻� ��� ��� �ó�������
		Scn= "scn_dt_itemDispAttribGroup"		'��ǰ�Ӽ� �˻�

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim groupKeyCnt,groupKeyVal,groupsize

		dim FseekTime

		StartNum = 0 						'// �˻����� Row
		call getSearchQuery(SearchQuery)	'// �˻� ��������

		'//�׷� ���� ����(���� ���� ����)
		SortQuery = " GROUP BY attribgrp order by attribgrp "

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores
		FTotalCount = 0

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

		''response.write "group : " & SearchQuery & SortQuery & "<br>"
		ret = Docruzer.Search(SvrAddr&":"&SvrPort, _
						Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,"",StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
		    response.write "ERR:"&Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

        ret = Docruzer.GetResult_GroupBy(FResultcount,groupKeyCnt,groupKeyVal,groupsize,100)

        IF( ret < 0 ) THEN
			'Response.write "GetResult_Row: " & Docruzer.msg
			SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
			EXIT Sub
		END IF

		if (groupKeyCnt<1) then
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
		    Exit Sub
		end if
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)
        Dim bufAttr

		FOR iRows = 0 to FResultcount -1
            bufAttr = groupKeyVal(iRows,0)
            SET FItemList(iRows) = new SearchGroupByItems
			FItemList(iRows).FAttribCd		= LEFT(bufAttr,6)
			FItemList(iRows).FAttribName	= replace(Mid(bufAttr,7,500),"_"," ")

            'rw "FAttribCd:"&FItemList(iRows).FAttribCd
            'rw "FAttribName:"&FItemList(iRows).FAttribName

			FItemList(iRows).FSubTotal 	= groupsize(iRows)
			FTotalCount = FTotalCount + groupsize(iRows)
		NEXT

		SET groupKeyVal = NOTHING
        SET groupsize   = NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	'####### ��ǰ �˻� �� ī����  ######
	PUBLIC SUB getTotalCount()

		'// �˻� ��� ��� �ó�������
		if FcolorCode="" or FcolorCode="0" then
			Scn= "scn_dt_itemDisp"		'�Ϲ� ��ǰ �˻�
		else
			'Scn= "scn_dt_itemDispColor"		'��ǰ �÷��� �˻�(����ī�װ�)
			Scn= "scn_dt_itemDisp"		'�Ϲ� ��ǰ �˻�
		end if
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// �˻����� Row
		call getSearchQuery(SearchQuery)	'// �˻� ��������

		'IF (FRectCateCode<>"" and FRectSearchItemDiv<>"y") Then
		'    ' SortQuery = " GROUP BY itemid" ''2013 �����߰� '' �ʿ���� 2015
		'else
    	'	' SortQuery = " "	'// ���� ���� ����
    	'end if

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores

        '// �⺻���� �˻�2�� ��, �˻�� �ִٸ� 1���� ���
        ''---------------------------------------------------------------------------------------------------------
        if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
            SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR) ''G_4THSCH_ADDR
        end if
        ''Į��Ĩ�˻�/�귣��˻� 3��
        if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
        ''---------------------------------------------------------------------------------------------------------

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

	'rw "group : " & SearchQuery & SortQuery & "<br>" ''
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)


		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_TotalCount(FTotalCount) '�˻���� �� ��
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)


		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

    '####### ��ǰ���� ī����  ######
	PUBLIC SUB getTotalItemColorCount()

        '' �÷�Ĩ �ڽ�.
		'// �˻� ��� ��� �ó�������
		Scn= "scn_dt_itemDispColorGroup"		'��ǰ �÷��� �˻�(����ī�װ�)

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim groupKeyCnt,groupKeyVal,groupsize

		dim FseekTime

		StartNum = 0 						'// �˻����� Row
		call getSearchQuery(SearchQuery)	'// �˻� ��������
		SortQuery = "Group by colorgrp Order by colorgrp "	'// ���� ���� ����

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores
		FTotalCount = 0
        ''---------------------------------------------------------------------------------------------------------
        if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
            SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR)
        end if
        ''Į��Ĩ�˻�/�귣��˻� 3��
        if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
        ''---------------------------------------------------------------------------------------------------------

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

	    ret = Docruzer.Search(SvrAddr&":"&SvrPort, _
						Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,"",StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		ret = Docruzer.GetResult_GroupBy(FResultcount,groupKeyCnt,groupKeyVal,groupsize,100)

		IF( ret < 0 ) THEN
			'Response.write "GetResult_Row: " & Docruzer.msg
			SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
			EXIT Sub
		END IF

		if (groupKeyCnt<1) then            ''�׷� ������ 1��;
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
		    Exit Sub
		end if

CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)
        dim bufColor, pos1

		FOR iRows = 0 to FResultcount -1
            bufColor = groupKeyVal(iRows,0)

            ''rw "groupKeyVal:"&bufColor
            if (Len(bufColor)>=6) then
    			SET FItemList(iRows) = new SearchGroupByItems
				FItemList(iRows).FcolorCode		= mid(bufColor,4,3)  ''colorgrp
				FItemList(iRows).FcolorName		= getCdPosVal(FItemList(iRows).FcolorCode,G_KSCOLORCD,G_KSCOLORNM)                                  ''mid(bufColor,8,pos1)
				FItemList(iRows).FcolorIcon		= "http://webimage.10x10.co.kr/color/colorchip/" & "chip"&CLNG(FItemList(iRows).FcolorCode)&".gif"  ''mid(bufColor,pos1+1,255)

				FItemList(iRows).FSubTotal 	= groupsize(iRows)
				FTotalCount = FTotalCount + groupsize(iRows)
            end if

		NEXT

        SET groupKeyVal = NOTHING
        SET groupsize   = NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	'####### ��ǰ���� ���� (�ּҰ���,�ִ밡��)  ######
	PUBLIC SUB getItemPriceRange()

		'// �˻� ��� ��� �ó�������
		Scn= "scn_dt_itemDisp"		'��ǰ�Ϲ� �˻�

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// �˻����� Row
		FPageSize = 2						'// ��� ��
		call getSearchQuery(SearchQuery)	'// �˻� ��������

		IF SearchQuery="" Then EXIT SUB

		dim Rowids,Scores

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

		'//���� ���� ���� : �ּҰ���
		''SortQuery = " order by sellcash asc "
		SortQuery = " extract by minmax(sellcash)"  '2015 search4 style

		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, AuthCode, Logs, Scn, SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(0)

		if (FResultcount=1) or (FResultcount=2) then
            SET FItemList(0) = new SearchGroupByItems

			ret = Docruzer.GetResult_Row( arrData, arrSize, 0 )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				FResultcount = 0
				EXIT Sub
			END IF


			FItemList(0).FminPrice		= arrData(7)
			FItemList(0).FmaxPrice		= arrData(7)

		    if (FResultcount=2) then '' ����� 1���� ��찡 ����
    		    ret = Docruzer.GetResult_Row( arrData, arrSize, 1 )

    			IF( ret < 0 ) THEN
    				'Response.write "GetResult_Row: " & Docruzer.msg
    				FResultcount = 0
    				EXIT Sub
    			END IF

    			FItemList(0).FmaxPrice		= arrData(7)
            end if

			SET arrData = NOTHING
			SET arrSize = NOTHING

	    end if
		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

'		SET FItemList(0) = new SearchGroupByItems
'
'		if FResultcount>0 then
'			ret = Docruzer.GetResult_Row( arrData, arrSize, 0 )
'			IF ret>=0 THEN
'				FItemList(0).FminPrice		= arrData(7)
'			end if
'			SET arrData = NOTHING
'			SET arrSize = NOTHING
'		end if
'

exit sub

		'//���� ���� ���� : �ִ밡��
		SortQuery = " order by sellcash desc "
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, AuthCode, Logs, Scn, SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		if FResultcount>0 then
			ret = Docruzer.GetResult_Row( arrData, arrSize, 0 )
			IF ret>=0 THEN
				FItemList(0).FmaxPrice		= arrData(7)
			end if
			SET arrData = NOTHING
			SET arrSize = NOTHING
		end if

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'' ## ���¼� �м�
	'' iextractLevel (2 or 5) 2:�⺻��
	public Function ExtractKeyword(byval iextractLevel)
		Dim Docruzer,ret
		Dim iRows
		Dim out_count,out_kwd
		Dim MaxCnt : MaxCnt =10
        dim ret_val, i, j

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�

		ret = Docruzer.ExtractKeyword _
						(SvrAddr&":"&SvrPort,_
						out_count,out_kwd,MaxCnt,"h1" _
						,FRectSearchTxt _
						,0,Docruzer.LC_KOREAN, Docruzer.CS_UTF8,CLNG(iextractLevel))

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			rw "ERR"
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF
		ret_val = ""
        for i=0 to out_count-1
            ret_val=ret_val&out_kwd(i)&vbCRLF
        next

        ret_val=TRim(ret_val)
        if (Right(ret_val,1)=vbCRLF) then ret_val=Left(ret_val,Len(ret_val)-1)

		ExtractKeyword = ret_val
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING
	end Function

	public Function OLD_ExtractKeyword()
		Dim Docruzer,ret
		Dim iRows
		Dim out_count,out_kwd
		Dim MaxCnt : MaxCnt =10
        dim ret_val, i, j

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�

		ret = Docruzer.ExtractKeyword _
						(SvrAddr&":"&SvrPort,_
						out_count,out_kwd,MaxCnt,"h1" _
						,replace(FRectSearchTxt," ","") _
						,0,Docruzer.LC_KOREAN, Docruzer.CS_UTF8,2)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			rw "ERR"
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF
		ret_val = ""
        for i=0 to out_count-1
            ret_val=ret_val&out_kwd(i)&vbCRLF
        next

        ret_val=TRim(ret_val)
        if (Right(ret_val,1)=vbCRLF) then ret_val=Left(ret_val,Len(ret_val)-1)

		OLD_ExtractKeyword = ret_val
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING
	end Function

    '####### ���Ǿ� 2018/03/30  ######
	PUBLIC FUNCTION GetSynonymList()

		Dim Docruzer,ret
		Dim iRows
		Dim term_count,synonym_count,synonym_list
		Dim MaxCnt : MaxCnt =10
        dim ret_val, i, j

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�

		ret = Docruzer.GetSynonymList _
						(SvrAddr&":"&SvrPort,_
						term_count,synonym_count,synonym_list,_
						MaxCnt,replace(FRectSearchTxt," ","") _
						,1,1,1,4,2,2)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			rw "ERR"
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF
		ret_val = ""
        for i=0 to term_count-1
			ret_val = ret_val&"["
            for j=0 to synonym_count(i)-1
                ret_val=ret_val&synonym_list(i,j)&","
            next
			ret_val=ret_val&"]"
            if (i<term_count-1) then ret_val=ret_val&vbCRLF
        next

        ret_val=TRim(ret_val)
		ret_val=replace(ret_val,",]","]")
        if (Right(ret_val,1)=",") then ret_val=Left(ret_val,Len(ret_val)-1)

		GetSynonymList = ret_val
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

	'####### ��õ�˻���  ######
	PUBLIC FUNCTION getRecommendKeyWords()

		Dim Docruzer,ret
		Dim iRows
		Dim arrData,arrSize
		Dim MaxCnt : MaxCnt =5

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�

		ret = Docruzer.RecommendKeyWord _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,_
						MaxCnt,replace(FRectSearchTxt," ",""),0)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF

		getRecommendKeyWords = arrData
		SET arrData = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

	'####### �α�˻���  ######
	PUBLIC FUNCTION getPopularKeyWords()

		DIM Docruzer,ret
		DIM iRows
		DIM arrData,arrSize
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�

		ret = Docruzer.getPopularKeyword _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,_
						MaxCnt,4)
		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION
		END IF
		getPopularKeyWords = arrData
		SET arrData = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

	'####### �α�˻��� (�߰�����) ######
	PUBLIC FUNCTION getPopularKeyWords2(byRef arDt, byRef arTg)

		DIM Docruzer,ret
		DIM iRows
		DIM arrSize, arrData, arrTag
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT FUNCTION
		END IF

        ''SvrAddr = G_3RDSCH_ADDR ''G_ORGSCH_ADDR  '' 106������ �ϴ�
        SvrAddr = G_ORGSCH_ADDR

		'�α� �˻��� (�߰�����)
		ret = Docruzer.getPopularKeyword2 _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,arrTag,_
						MaxCnt,4)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION
		END IF

		arDt = arrData
		arTg = arrTag

		SET arrData = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

    '####### �ǽð��α�˻���:K-search  ######
	PUBLIC FUNCTION getRealtimePopularKeyWords(byRef arDt, byRef arTg)

		DIM Docruzer,ret
		DIM iRows
		DIM arrData,arrSize, arrTags
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�

		ret = Docruzer.getRealTimePopularKeyword _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,arrTags,_
						MaxCnt,1,0)                     ''' 0 file / 1 memory
		IF( ret < 0 ) THEN
		    'rw "TTT:"&Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION
		END IF
		arDt = arrData
		arTg = arrTags
		SET arrData = NOTHING
		SET arrTags = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION

End Class

'###############################
'### �� ��ǰ��ȸ �̷�      ###
'###############################
Class ViewHistoryCls
public FItemList()

	public FResultCount
	public FPageSize
	public FCurrpage
	public FTotalCount
	public FTotalPage
	public FScrollCount

	public FRectUserID

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FPageSize = 100
		FCurrpage = 1
	End Sub

	Private Sub Class_Terminate()
		'//
	End Sub

	public Function GetRandomUserID()
		dim sql, i

		sql = " select top 1 userid "
		sql = sql & " from ( "
		sql = sql & " 	select top 100 userid "
		sql = sql & " 	from "
		sql = sql & " 	[db_EVT].[dbo].[tbl_itemevent_userLogData_BACK] "
		sql = sql & " 	where type = 'item' "
		sql = sql & " 	order by idx desc "
		sql = sql & " ) T "
		sql = sql & " order by newid() "
		rsEVTget.Open SQL,dbEVTget,1

		GetRandomUserID = ""
		if  not rsEVTget.EOF  then
			GetRandomUserID = rsEVTget("userid")
		end if
		rsEVTget.close

	End Function

    public Sub getMyTodayViewListNew()
        dim sqlStr, i

        sqlStr = " Select top " & CStr(FPageSize*FCurrPage) & " * From "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + " 	select max(L.idx) as idx, i.cate_large, i.cate_mid, i.cate_small "
        sqlStr = sqlStr + " 		,i.itemid, i.itemname, i.sellcash, i.sellyn, i.limityn, i.limitno, i.limitsold "
        sqlStr = sqlStr + " 		,i.itemgubun, i.deliverytype, i.evalcnt, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue,i.curritemcouponidx "
        sqlStr = sqlStr + " 		,i.basicimage, i.smallimage, i.listimage, i.listimage120, i.icon1image, i.icon2image, (Case When isNull(i.frontMakerid,'')='' then i.makerid else i.frontMakerid end) as makerid, i.brandname, max(L.regdate) as regdate, i.sailyn, i.sailprice "
		sqlStr = sqlStr + " 		,i.orgprice, i.specialuseritem, i.optioncnt, i.itemdiv, i.dispcate1, i.sellSTDate, [db_analyze_data_raw].[dbo].[getDisplayCateName](convert(bigint, left(d.catecode, 9))) as code_nm "
		''sqlStr = sqlStr + " 		, [db_analyze_data_raw].[dbo].[getKeywordsFromItemName](i.itemname) as kwds "
        sqlStr = sqlStr + " 	From db_evt.[dbo].[tbl_itemevent_userLogData_FrontRecent] L "
        sqlStr = sqlStr + " 	inner join db_analyze_data_raw.dbo.tbl_item i on L.itemid = i.itemid "
		sqlStr = sqlStr + " 	left join [db_analyze_data_raw].[dbo].[tbl_display_cate_item] d on i.itemid = d.itemid and d.isDefault = 'Y' "
        sqlStr = sqlStr + " 	Where L.type in ('item', 'itemrect') And L.userid='"&FRectUserID&"' "
        sqlStr = sqlStr + " 	group by i.cate_large, i.cate_mid, i.cate_small "
        sqlStr = sqlStr + " 		,i.itemid, i.itemname, i.sellcash, i.sellyn, i.limityn, i.limitno, i.limitsold "
        sqlStr = sqlStr + " 		,i.itemgubun, i.deliverytype, i.evalcnt, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue,i.curritemcouponidx "
        sqlStr = sqlStr + " 		,i.basicimage, i.smallimage, i.listimage, i.listimage120, i.icon1image, i.icon2image, (Case When isNull(i.frontMakerid,'')='' then i.makerid else i.frontMakerid end), i.brandname, i.sailyn, i.sailprice "
		sqlStr = sqlStr + " 		,i.orgprice, i.specialuseritem, i.evalcnt, i.optioncnt, i.itemdiv, i.dispcate1, i.sellSTDate, [db_analyze_data_raw].[dbo].[getDisplayCateName](convert(bigint, left(d.catecode, 9))) "
		''sqlStr = sqlStr + " 		,[db_analyze_data_raw].[dbo].[getKeywordsFromItemName](i.itemname) "
        sqlStr = sqlStr + " )AA Where itemid is not null "
		sqlStr = sqlStr + " order by idx "

        rsEVTget.pagesize = FPageSize
        rsEVTget.Open sqlStr, dbEVTget, 1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsEVTget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)
		i=0
		if Not rsEVTget.Eof then
			rsEVTget.absolutepage = FCurrPage
			do until rsEVTget.eof
				set FItemList(i) = new CCategoryPrdItem

				FItemList(i).FItemID       = rsEVTget("itemid")
				FItemList(i).FItemName     = db2html(rsEVTget("itemname"))
				FItemList(i).FItemDiv 	= rsEVTget("itemdiv")		'��ǰ �Ӽ�

				FItemList(i).FSellcash     = rsEVTget("sellcash")
				FItemList(i).FSellYn       = rsEVTget("sellyn")
				FItemList(i).FLimitYn      = rsEVTget("limityn")
				FItemList(i).FLimitNo      = rsEVTget("limitno")
				FItemList(i).FLimitSold    = rsEVTget("limitsold")
				FItemList(i).Fitemgubun    = rsEVTget("itemgubun")
				FItemList(i).FDeliverytype = rsEVTget("deliverytype")

				FItemList(i).Fevalcnt       = rsEVTget("evalcnt")
				FItemList(i).Fitemcouponyn  = rsEVTget("itemcouponyn")
				FItemList(i).Fitemcoupontype 	= rsEVTget("itemcoupontype")
				FItemList(i).Fitemcouponvalue 	= rsEVTget("itemcouponvalue")
				FItemList(i).Fcurritemcouponidx = rsEVTget("curritemcouponidx")

				FItemList(i).FImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsEVTget("basicimage")
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsEVTget("smallimage")
				FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsEVTget("listimage")
				FItemList(i).FImageList120 = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsEVTget("listimage120")

				If FItemList(i).FItemDiv="21" Then
				FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" + rsEVTget("icon2image")
				Else
				FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsEVTget("icon2image")
				End If
				FItemList(i).FMakerID   = rsEVTget("makerid")
				FItemList(i).FBrandName = db2html(rsEVTget("brandname"))
				FItemList(i).FRegdate   = rsEVTget("regdate")

				FItemList(i).FSaleYn    = rsEVTget("sailyn")
				FItemList(i).FSailPrice = rsEVTget("sailprice")
				FItemList(i).FOrgPrice   = rsEVTget("orgprice")
				FItemList(i).FSpecialuseritem = rsEVTget("specialuseritem")
				FItemList(i).Fevalcnt = rsEVTget("evalcnt")
				FItemList(i).FOptioncnt	= rsEVTget("optioncnt")
				FItemList(i).FCateName = rsEVTget("code_nm")
				FItemList(i).Fsdate	= rsEVTget("sellSTDate")
		''		FItemList(i).Fkeywords	= rsEVTget("kwds")

				rsEVTget.movenext
				i=i+1
			loop
		end if
		rsEVTget.close
    end Sub

	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION

End Class

'###############################
'### ��ǰ�ı� �˻�           ###
'###############################
Class SearchItemEvaluate
	public FItemList()

	public FResultCount
	public FPageSize
	public FCurrpage
	public FTotalCount
	public FTotalPage
	public FScrollCount

	public FRectDispCate
	public FRectSort
	public FRectMode
	public FRectArrItemid

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'// ����Ʈ ���� ���� ���(�˻���) ���� : ��ǰ�� �ֱ� 1�� ���� ��� //
	public Sub GetBestReviewArrayList()
		dim sql, i

		if FRectArrItemid="" then
			Exit Sub
		end if

		'// ��� ���� //
		sql = " SELECT e.idx, e.userid, e.regdate, e.itemid " + vbcrlf
		sql = sql & " , e.contents, (Case When isNull(i.frontMakerid,'')='' then i.makerid else i.frontMakerid end) as makerid, i.brandname, e.TotalPoint as Tpoint " + vbcrlf
		sql = sql + " , isnull(e.Point_function,0) as Point_function "
		sql = sql + " , isnull(e.Point_Design,0) as Point_Design "
		sql = sql + " , isnull(e.Point_Price,0) as Point_Price "
		sql = sql + " , isnull(e.Point_satisfy,0) as Point_satisfy "
		sql = sql & " , i.itemname, i.sellcash, i.orgprice, i.sellyn, i.sailyn, i.limityn, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue " + vbcrlf
		sql = sql & " , i.listimage120, i.evalcnt, i.itemscore, e.File1, e.File2, i.icon1image, i.evalcnt, ic.favcount, i.tenOnlyYn  " + vbcrlf
		sql = sql & " , (case when isnull(ee.itemid,'') <> '' then 'Y' else 'N' end) as Eval_excludeyn"
		sql = sql & " FROM db_board.[dbo].tbl_item_evaluate e " + vbcrlf
		sql = sql & " JOIN  [db_item].[dbo].tbl_item i  " + vbcrlf
		sql = sql & " 	on e.itemid = i.itemid " + vbcrlf
		sql = sql & " JOIN  [db_item].[dbo].tbl_item_contents ic  " + vbcrlf
		sql = sql & " 	on i.itemid = ic.itemid " + vbcrlf
		sql = sql + " left join db_board.dbo.tbl_Item_Evaluate_exclude ee"
		sql = sql + " 	on e.itemid=ee.itemid"
		sql = sql & " WHERE e.idx in (" & vbcrlf
		sql = sql & "	 Select idx from (" & vbcrlf
		sql = sql & "		select itemid, max(idx) as idx " & vbcrlf
		sql = sql & "		from db_board.[dbo].tbl_item_evaluate" & vbcrlf
		sql = sql & "		where itemid in (" & FRectArrItemid & ")" & vbcrlf
		sql = sql & "			and isusing='Y'" & vbcrlf

		if FRectMode="photo" then
			sql = sql & " and (File1 is Not Null or File2 is Not Null) " + vbcrlf
		end if

		sql = sql & "		group by itemid) as T )"
		Select Case FRectSort
			Case "ne"
				'�Ż��
				sql = sql & " ORDER BY e.itemid DESC  " + vbcrlf
			Case "be"
				'�α��ǰ ��
				sql = sql & " ORDER BY i.itemscore DESC  " + vbcrlf
			Case "lp"
				'���� ���ݼ�
				sql = sql & " ORDER BY i.sellcash asc  " + vbcrlf
			Case "hp"
				'���� ���ݼ�
				sql = sql & " ORDER BY i.sellcash desc  " + vbcrlf
			Case "hs"
				'���� ���μ�
				sql = sql & " ORDER BY (i.sellcash/i.orgprice) desc  " + vbcrlf
			Case else
				sql = sql & " ORDER BY e.itemid DESC  " + vbcrlf
		End Select

		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount
		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			Do Until rsget.EOF or rsget.BOF
				set FItemList(i) = new CCategoryPrdItem

				FItemList(i).fEval_excludeyn 	= rsget("Eval_excludeyn")
				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fuserid			= rsget("userid")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
				FItemList(i).Fmakerid			= db2html(rsget("makerid"))
				FItemList(i).Fbrandname			= db2html(rsget("brandname"))
				FItemList(i).Fevalcnt			= rsget("evalcnt")
				FItemList(i).Fcontents			= db2html(rsget("contents"))
				FItemList(i).FOrgprice			= rsget("orgprice")
				FItemList(i).FSellYn			= rsget("sellyn")
				FItemList(i).FSaleYn			= rsget("sailyn")
				FItemList(i).FLimitYn			= rsget("limityn")
				FItemList(i).FSellCash			= rsget("sellcash")
				FItemList(i).FPoints			= rsget("TPoint")
				FItemList(i).FPoint_fun			= rsget("Point_Function")
				FItemList(i).FPoint_dgn			= rsget("Point_Design")
				FItemList(i).FPoint_prc			= rsget("Point_Price")
				FItemList(i).FPoint_stf			= rsget("Point_Satisfy")
				FItemList(i).Fitemcouponyn		= rsget("itemcouponyn")
				FItemList(i).Fitemcoupontype	= rsget("itemcoupontype")
				FItemList(i).FItemCouponValue	= rsget("itemcouponvalue")
				FItemList(i).FTenOnlyYn			= rsget("tenOnlyYn")

				if Not(rsget("File1")="" or isNull(rsget("File1"))) then
					FItemList(i).FImageIcon1		= "http://imgstatic.10x10.co.kr/goodsimage/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("File1")
				end if
				if Not(rsget("File2")="" or isNull(rsget("File2"))) then
					FItemList(i).FImageIcon2		= "http://imgstatic.10x10.co.kr/goodsimage/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("File2")
				end if
				FItemList(i).FImageList120    = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage120")
				FItemList(i).FIcon1Image	  = "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon1image")
				FItemList(i).FEvalcnt		  = rsget("evalcnt")
				FItemList(i).Ffavcount		  = rsget("favcount")

				rsget.moveNext
				i = i + 1
			Loop
		end if
		rsget.close
	end Sub

	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION
End Class



'###############################
'### �̺�Ʈ �˻�             ###
'###############################

Class SearchEventItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fevt_code
	PUBLIC Fevt_bannerimg
	PUBLIC Fevt_startdate
	PUBLIC Fevt_enddate
	PUBLIC Fevt_kind
	PUBLIC Fbrand
	PUBLIC Fevt_LinkType
	PUBLIC Fevt_bannerlink
	PUBLIC Fetc_itemid
	PUBLIC Fetc_itemimg
	PUBLIC Ficon1image
	PUBLIC Fevt_name
	PUBLIC Fevt_tag
	PUBLIC Fevt_subcopyK
	PUBLIC Fissale
	PUBLIC Fisgift
	PUBLIC Fisitemps
	PUBLIC Fiscoupon
	PUBLIC FisOnlyTen
	PUBLIC Fisoneplusone
	PUBLIC Fisfreedelivery
	PUBLIC Fisbookingsell
	PUBLIC Fiscomment
	PUBLIC Fevt_state

End Class

Class SearchEventCls
    ''�˻������� �̺�Ʈ.
	Private SUB Class_initialize()

        '' �⺻ 1������--------------------------------------------------------------------------------------------
		SvrAddr = getTimeChkAddr(G_1STSCH_ADDR)
		''---------------------------------------------------------------------------------------------------------

		SvrPort = "6167"'DocSvrPort
		AuthCode = "" '������
		Logs = "" '�αװ�

		FResultCount = 0
		FTotalCount = 0
		FPageSize = 10
		FCurrPage = 1
		FPageSize = 30

	End SUB

	Private SUB Class_Terminate()

	End SUB

	dim FItemList
	dim FPageSize
	dim FCurrPage
	dim FScrollCount
	dim FResultCount
	dim FTotalCount
	dim FTotalPage
	dim FRectSearchTxt		'�˻���
	dim FRectExceptText		'���ܾ�
	dim FRectChannel		'�˻� ä�� (W:isWeb, M:isMobile, A:isApp)

	Private SvrAddr
	Private SvrPort
	Private AuthCode
	Private Logs
	Private Scn
	private strQuery
	Private Order
	Private StartNum
	Private SearchQuery
	Private SortQuery

    public function InitDocruzer(iDocruzer)
        InitDocruzer = FALSE
        IF ( iDocruzer.BeginSession() < 0 ) THEN
			EXIT function
		End If

        IF NOT DocSetOption(iDocruzer) THEN
			EXIT function
		End If
		InitDocruzer = TRUE
    End function

    public function DocSetOption(iDocruzer)
        dim ret
        ret = iDocruzer.SetOption(iDocruzer.OPTION_REQUEST_CHARSET_UTF8,1)
        DocSetOption = (ret>=0)
    end function

	'####### �̺�Ʈ �˻� ######
	PUBLIC SUB getEventList()

		dim Scn : Scn= "scn_dt_event2015" 		'// �˻� ��� ��� �ó�������
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = (FCurrPage -1)*FPageSize	'// �˻����� Row

		'// �˻� �������� (�̺�Ʈ�� ���ܾ� �˻� �������� �������� ����)
		'IF FRectExceptText<>"" Then
		'	SearchQuery = " (idx_eventKeyword='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'���ܾ�
		'else
			SearchQuery = " idx_eventKeyword='" & FRectSearchTxt & "'  allword "	'Ű����˻�
		'End if

		SearchQuery = SearchQuery &_
					" and idx_evt_state='7' " &_
					" and idx_evt_startdate<='" & Replace(date(),"-","") & "000000' " &_
					" and idx_evt_enddate>='" & Replace(date(),"-","") & "000000' "

		Select Case FRectChannel
			Case "W"
				SearchQuery = SearchQuery & " and idx_evt_isWeb=1 "
			Case "M"
				SearchQuery = SearchQuery & " and idx_evt_isMobile=1 "
			Case "A"
				SearchQuery = SearchQuery & " and idx_evt_isApp=1 "
			Case Else
				SearchQuery = SearchQuery & " and idx_evt_isWeb=1 "
		End Select

		'//�׷� ������ ����(���� ���� ����)
		SortQuery = "Order by evt_startdate desc "

		dim Rowids,Scores

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '�˻���� �� ��

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchEventItems
				FItemList(iRows).Fevt_code			= arrData(0)
				FItemList(iRows).Fevt_bannerimg		= arrData(1)
				FItemList(iRows).Fevt_startdate		= dateSerial(left(arrData(2),4),mid(arrData(2),5,2),mid(arrData(2),7,2))	'### ASP ��¥���·� ��ȯ
				FItemList(iRows).Fevt_enddate		= dateSerial(left(arrData(3),4),mid(arrData(3),5,2),mid(arrData(3),7,2))	'### ASP ��¥���·� ��ȯ
				FItemList(iRows).Fevt_kind			= arrData(4)
				FItemList(iRows).Fbrand				= arrData(5)
				FItemList(iRows).Fevt_LinkType		= arrData(6)
				FItemList(iRows).Fevt_bannerlink	= arrData(7)
				FItemList(iRows).Fetc_itemid		= arrData(8)
				FItemList(iRows).Fetc_itemimg		= arrData(9)
				FItemList(iRows).Ficon1image		= arrData(10)
				FItemList(iRows).Fevt_name			= arrData(11)
				FItemList(iRows).Fevt_tag			= arrData(12)
				FItemList(iRows).Fevt_subcopyK		= arrData(13)
				FItemList(iRows).Fissale			= arrData(14)
				FItemList(iRows).Fisgift			= arrData(15)
				FItemList(iRows).Fisitemps			= arrData(16)
				FItemList(iRows).Fiscoupon			= arrData(17)
				FItemList(iRows).FisOnlyTen			= arrData(18)
				FItemList(iRows).Fisoneplusone		= arrData(19)
				FItemList(iRows).Fisfreedelivery	= arrData(20)
				FItemList(iRows).Fisbookingsell		= arrData(21)
				FItemList(iRows).Fiscomment			= arrData(22)
				FItemList(iRows).Fevt_state			= arrData(23)

			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION
end Class




'###############################
'### PLAY �˻�             ###
'###############################

Class SearchPlayItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fplaycate
	PUBLIC Fidx
	PUBLIC FlistImage
	PUBLIC Fplayname
	PUBLIC Fopendate
	PUBLIC FsortNo
	PUBLIC Fcont
	PUBLIC Ftag

	Function getPlayCateNm()
		Select Case Fplaycate
			Case "1"
				getPlayCateNm = "Ground"
			Case "2"
				getPlayCateNm = "Style+"
			Case "3"
				getPlayCateNm = "Color trend"
			Case "4"
				getPlayCateNm = "Design fingers"
			Case "5"
				getPlayCateNm = "�׸��ϱ�"
			Case "6"
				getPlayCateNm = "Video clip"
			Case "7"
				getPlayCateNm = "T-episode"
		End Select
	End function

	Function getPlayCateLink()
		Select Case Fplaycate
			Case "1"
				getPlayCateLink = "/play/playGround?gidx="
			Case "2"
				getPlayCateLink = "/play/playStylePlusView.asp?idx="
			Case "3"
				getPlayCateLink = "/play/playColorTrendView.asp?ctcode="
			Case "4"
				getPlayCateLink = "/play/playdesignfingers.asp?fingerid="
			Case "5"
				getPlayCateLink = "/play/playPicDiary.asp?idx="
			Case "6"
				getPlayCateLink = "/play/playVideoClip.asp?idx="
			Case "7"
				getPlayCateLink = "/play/playtEpisodePhotopick.asp?"
		End Select
	End function

End Class

Class SearchPlayCls
    ''�˻������� Play.
	Private SUB Class_initialize()

        '' �⺻ 1������--------------------------------------------------------------------------------------------
		SvrAddr = getTimeChkAddr(G_1STSCH_ADDR)
		''---------------------------------------------------------------------------------------------------------

		SvrPort = "6167"'DocSvrPort
		AuthCode = "" '������
		Logs = "" '�αװ�

		FResultCount = 0
		FTotalCount = 0
		FPageSize = 10
		FCurrPage = 1
		FPageSize = 30

	End SUB

	Private SUB Class_Terminate()

	End SUB

	dim FItemList
	dim FPageSize
	dim FCurrPage
	dim FScrollCount
	dim FResultCount
	dim FTotalCount
	dim FTotalPage
	dim FRectSearchTxt		'�˻���
	dim FRectExceptText		'���ܾ�

	Private SvrAddr
	Private SvrPort
	Private AuthCode
	Private Logs
	Private Scn
	private strQuery
	Private Order
	Private StartNum
	Private SearchQuery
	Private SortQuery

    public function InitDocruzer(iDocruzer)
        InitDocruzer = FALSE
        IF ( iDocruzer.BeginSession() < 0 ) THEN
			EXIT function
		End If

        IF NOT DocSetOption(iDocruzer) THEN
			EXIT function
		End If
		InitDocruzer = TRUE
    End function

    public function DocSetOption(iDocruzer)
        dim ret
        ret = iDocruzer.SetOption(iDocruzer.OPTION_REQUEST_CHARSET_UTF8,1)
        DocSetOption = (ret>=0)
    end function

	'####### PLAY �˻� ######
	PUBLIC SUB getPlayList()
        'rw "getPlayList - SKIP"
        'Exit SUB

		dim Scn : Scn= "scn_dt_play2013" 		'// �˻� ��� ��� �ó�������
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = (FCurrPage -1)*FPageSize	'// �˻����� Row

		'// �˻� �������� (Play�� ���ܾ� �˻� �������� �������� ����)
		SearchQuery = " idx_playKeyword='" & FRectSearchTxt & "'  allword "	'Ű����˻�

		'//�׷� ������ ����(���� ���� ����)
		SortQuery = "Order by opendate desc, sortNo, idx desc "

		dim Rowids,Scores

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '�˻���� �� ��

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchPlayItems
				FItemList(iRows).Fplaycate	= arrData(0)
				FItemList(iRows).Fidx		= arrData(1)
				FItemList(iRows).FlistImage	= arrData(2)
				FItemList(iRows).Fplayname	= arrData(3)
				FItemList(iRows).Fopendate	= dateSerial(left(arrData(4),4),mid(arrData(4),5,2),mid(arrData(4),7,2))	'### ASP ��¥���·� ��ȯ
				FItemList(iRows).FsortNo	= arrData(5)
				FItemList(iRows).Fcont		= arrData(6)
				FItemList(iRows).Ftag		= arrData(7)

			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION
end Class



'// ���˻��� ����ó�� (��Ű, �ֱ� 5�� ����漼������ ����;2014.07.29:������)
Sub procMySearchKeyword(kwd)
	Dim arrCKwd, rstKwd, i, excKwd
	'''arrCKwd = request.Cookies("search")("keyword")
	arrCKwd = session("myKeyword")
	arrCKwd = split(arrCKwd,",")
	''excKwd = "update,select,insert,and,union,from,alter,shutdown,kill,declare,exec,having,;,--"		'��Ű���� ���� �ܾ� (��Ű ������)

	rstKwd = trim(kwd)
	if ubound(arrCKwd)>-1 then
		for i=0 to ubound(arrCKwd)
			if not(chkArrValue(excKwd,lcase(arrCKwd(i)))) then
				if arrCKwd(i)<>trim(kwd) then rstKwd = rstKwd & "," & arrCKwd(i)
			end if
			if i>4 then Exit For
		next
	end if

	'��Ű����
	''response.Cookies("search").domain = "10x10.co.kr"
	''''response.cookies("search").Expires = Date + 3	'3�ϰ� ��Ű ���� => ������
	''response.Cookies("search")("keyword") = rstKwd
	session("myKeyword") = rstKwd
end Sub

'// ������/���Ǿ� ��ȯ ó�� (����� �� ���Ǿ �ȵǴ� ���� ������ ���)
Function chgCoinedKeyword(kwd)
	dim arrChgTxt, arrItm
	arrChgTxt = split("��8||ban8",",")

	for each arrItm in arrChgTxt
		arrItm = split(arrItm,"||")
		if ubound(arrItm)>0 then
			kwd = Replace(kwd,arrItm(0),arrItm(1))
		end if
	next

	chgCoinedKeyword = kwd
end Function


'// �߰� ī�װ� ��ȣ ���� (�߰�ī�װ����� �ش� ī�װ� ��ȣ�� ����)
Function getArrayDispCate(vDisp,vArr)
	Dim vRst, i

	if vArr="" or isNull(vArr) or vDisp="" or isNull(vDisp) then Exit Function

	vArr = replace(trim(vArr)," ",",")
	vRst = split(vArr,",")

	if Not(isArray(vRst)) then Exit Function

	for i=0 to ubound(vRst)
		if inStr(vRst(i),vDisp)>0 then
			getArrayDispCate = vRst(i)
			Exit function
		end if
	next
end Function
%>
