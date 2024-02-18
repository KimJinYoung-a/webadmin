<%
CONST CMAXMARGIN = 10
CONST CMALLGUBUN = "naverep"
CONST CMALLNAME = "nvstoremoonbangu"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST cspDlvrId	= "10040413"							'���ó�ڵ�

Class CNvstoremoonbanguItem
	Public FItemid
	Public FItemname
	Public FSmallImage
	Public FMakerid
	Public FRegdate
	Public FLastUpdate
	Public FOrgPrice
	Public FSellCash
	Public FBuyCash
	Public FSellYn
	Public FSaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FNvstoremoonbanguRegdate
	Public FNvstoremoonbanguLastUpdate
	Public FNvstoremoonbanguGoodNo
	Public FNvstoremoonbanguPrice
	Public FNvstoremoonbanguSellYn
	Public FRegUserid
	Public FNvstoremoonbanguStatCd
	Public FCateMapCnt
	Public FDeliverytype
	Public FDefaultdeliverytype
	Public FDefaultfreeBeasongLimit
	Public FOptionCnt
	Public FRegedOptCnt
	Public FRctSellCNT
	Public FAccFailCNT
	Public FLastErrStr
	Public FInfoDiv
	Public FOptAddPrcCnt
	Public FOptAddPrcRegType
	Public FAPIaddImg
	Public FItemDiv
	Public FOrgSuplyCash
	Public FIsusing
	Public FKeywords
	Public FVatinclude
	Public FOrderComment
	Public FBasicImage
	Public FBasicimageNm
	Public FMainImage
	Public FMainImage2
	Public FSourcearea
	Public FMakername
	Public FUsingHTML
	Public FItemcontent

	Public FTenCateLarge
	Public FTenCateMid
	Public FTenCateSmall
	Public FTenCDLName
	Public FTenCDMName
	Public FTenCDSName
	Public FCateKey
	Public FDepth1Nm
	Public FDepth2Nm
	Public FDepth3Nm
	Public FDepth4Nm
	Public FNeedCert

	Public FRequireMakeDay
	Public FSafetyyn
	Public FSafetyDiv
	Public FSafetyNum
	Public FMaySoldOut
	Public FRegitemname
	Public FRegImageName
	Public FSpecialPrice
	Public FStartDate
	Public FEndDate

	Public FIdx
	Public FImgtype
	Public FGubun
	Public FImagename
	Public FNotSchIdx

	Public Function getNvstoremoonbanguStatName
	    If IsNULL(FNvstoremoonbanguStatCd) then FNvstoremoonbanguStatCd=-1
		Select Case FNvstoremoonbanguStatCd
			CASE -9 : getNvstoremoonbanguStatName = "�̵��"
			CASE -1 : getNvstoremoonbanguStatName = "��Ͻ���"
			CASE 0 : getNvstoremoonbanguStatName = "<font color=blue>��Ͽ���</font>"
			CASE 1 : getNvstoremoonbanguStatName = "���۽õ�"
			CASE 7 : getNvstoremoonbanguStatName = ""
			CASE ELSE : getNvstoremoonbanguStatName = FNvstoremoonbanguStatCd
		End Select
	End Function

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[���� "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[��ü����]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[��ü]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function

	'// ǰ������
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<br><font color=blue>����:"&getLimitEa&"</font>"
	    End if
	End Function

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	Function getItemNameFormat()
		Dim buf
		buf = "[�ٹ�����]"&replace(FItemName,"'","")		'���� ��ǰ�� �տ� [�ٹ�����] �̶�� ����
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemNameFormat = buf
	End Function

	Public Function getTotalSuryang()
		If Flimityn = "Y" Then
			If FLimitno - FLimitSold - 5 < 1 Then
				getTotalSuryang = 0
			Else
				getTotalSuryang = FLimitno-FLimitSold-5
			End If
		Else
			getTotalSuryang = "999"
		End If
	End Function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

    public function isImageChanged()
        Dim ibuf : ibuf = getBasicImage
        if InStr(ibuf,"-")<1 then
            isImageChanged = FALSE
            Exit function
        end if
        isImageChanged = ibuf <> FregImageName
    end function

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(�ٹ�), 4,5(����)
			IsFreeBeasong = True
		End If
'		If (FSellcash>=30000) then IsFreeBeasong=True
		If (FdeliveryType=9) Then														'��ü����
'			If (Clng(FSellcash) >= Clng(FdefaultfreeBeasongLimit)) then
'				IsFreeBeasong=True
'			End If
			IsFreeBeasong = False
		End If
    End Function

	Public Function fngetMustPrice
		Dim strRst, GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			fngetMustPrice = Forgprice
		Else
			fngetMustPrice = FSellCash
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CNvstoremoonbangu
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectItemID
	Public FRectItemName
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectMakerid
	Public FRectNvstoremoonbanguGoodNo
	Public FRectMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectNvstoremoonbanguNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectNvstoremoonbanguYes10x10No
	Public FRectNvstoremoonbanguNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectregedOptNull
	Public FRectFailCntExists
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType
	Public FRectIsSpecialPrice
	Public FRectScheduleNotInItemid

	'// ���̹� ������� ���汸 ��ǰ ��� // ������ ������ �޶�� ��..
	Public Sub getNvstoremoonbanguRegedItemList
		Dim i, sqlStr, addSql
		'�귣��˻�
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'��ǰ�ڵ� �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'��ǰ�� �˻�
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'������� ��ǰ��ȣ �˻�
        If (FRectNvstoremoonbanguGoodNo <> "") then
            If Right(Trim(FRectNvstoremoonbanguGoodNo) ,1) = "," Then
            	FRectNvstoremoonbanguGoodNo = Replace(FRectNvstoremoonbanguGoodNo,",,",",")
            	FRectNvstoremoonbanguGoodNo = Replace(FRectNvstoremoonbanguGoodNo,"''","'")
            	addSql = addSql & " and J.nvstoremoonbanguGoodNo in (" & Left(FRectNvstoremoonbanguGoodNo, Len(FRectNvstoremoonbanguGoodNo)-1) & ")"
            Else
				FRectNvstoremoonbanguGoodNo = Replace(FRectNvstoremoonbanguGoodNo,",,",",")
				FRectNvstoremoonbanguGoodNo = Replace(FRectNvstoremoonbanguGoodNo,"''","'")
            	addSql = addSql & " and J.nvstoremoonbanguGoodNo in (" & FRectNvstoremoonbanguGoodNo & ")"
            End If
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

		'��Ͽ��� �˻�
		Select Case FRectExtNotReg
			Case "Q"	''��Ͻ���
				addSql = addSql & " and J.nvstoremoonbanguStatCd = -1"
			Case "J"	'��Ͽ����̻�
				addSql = addSql & " and J.nvstoremoonbanguStatCd >= 0"
		    Case "A"	'���۽õ��߿���
				addSql = addSql & " and J.nvstoremoonbanguStatCd = 1"
		    Case "I"	'�̹����� �Ϸ�
				addSql = addSql & " and isnull(J.nvstoremoonbanguGoodNo, '') = '' "
				addSql = addSql & " and J.APIaddImg = 'Y'"
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.nvstoremoonbanguStatCd = 7"
				addSql = addSql & " and J.nvstoremoonbanguGoodNo is Not Null"
		End Select

		'�̵�� ������ư Ŭ�� ��
		Select Case FRectIsReged
			Case "N"	'��Ͽ����̻�
			    addSql = addSql & " and J.itemid is NULL "
				If (FRectExcTrans <> "N") Then
					addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				end if
				if (FRectItemID = "") and (FRectMakerid = "") then
					'// �ֱ� 3������ ��ϵ� ��ǰ��
					addSql = addSql & " and i.regdate >= DateAdd(m, -3, getdate()) "
				end if
		End Select

		'�Ǹſ��� �˻�
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and i.sellYn='Y'"			'�Ǹ�
			Case "N"	addSql = addSql & " and i.sellYn in ('S','N')"	'ǰ��
		End Select

		'�ٹ����� �������� �˻�
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'�ٹ����� ���Ͽ��� �˻�
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
		End If

		'������ �� ���� CMAXMARGIN �̻� �˻�
		If (FRectonlyValidMargin <> "") Then
			IF (FRectonlyValidMargin = "Y") Then
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) >= " & CMAXMARGIN & VbCrlf
			Else
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) < " & CMAXMARGIN & VbCrlf
			End If
		End If

		If (FRectStartMargin <> "") OR (FRectEndMargin <> "") Then
			If (FRectStartMargin <> "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and ("
				addSql = addSql & " 	convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
				addSql = addSql & " 	and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
				addSql = addSql & " ) "
			ElseIf (FRectStartMargin <> "") And (FRectEndMargin = "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
			ElseIf (FRectStartMargin = "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
			End If
		End If

		'�ֹ����� ���� �˻�
		If FRectisMadeHand <> "" Then
			If (FRectisMadeHand = "Y") Then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

		'�ɼ� ���� �˻�
		If FRectIsOption <> "" Then
			If FRectIsOption = "optAll" Then			'�ɼ���ü
				addSql = addSql & " and i.optioncnt > 0"
			ElseIf FRectIsOption = "optaddpricey" Then	'�߰��ݾ�Y
				addSql = addSql & " and i.optioncnt > 0"
 				addSql = addSql & " and J.optAddPrcCnt > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'�߰��ݾ�N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
			ElseIf FRectIsOption = "optN" Then			'��ǰ
				addSql = addSql & " and i.optioncnt = 0"
			End If
		End If

		'�ٹ����� ǰ������ �˻�
		If (FRectInfoDiv <> "") then
			If (FRectInfoDiv = "YY") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			ElseIf (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End If
		End If

		'�ٹ����� ������� �귣�� ���� �˻�
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
			End If
		End If

		'�ٹ����� ������� ��ǰ ���� �˻�
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
			End If
		End If

		'�ٹ����� �������� ��ǰ ���� �˻�
		If (FRectScheduleNotInItemid <> "") then
			If (FRectScheduleNotInItemid = "Y") Then
				addSql = addSql & " and sc.idx is not null "
				'addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_schedule_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'WMP') "
			ElseIf (FRectScheduleNotInItemid = "N") Then
				addSql = addSql & " and sc.idx is null "
				'addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_schedule_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'WMP') "			End If
			End If
		End If

		'���޸� �������� ��ǰ �˻�
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				'// �ǸŰ� 1���� �̸��� ����
				'// �ּ� ������ 10%
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				''addSql = addSql & " or exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				''addSql = addSql & " or exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
				addSql = addSql & " or exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " or exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " or i.isExtUsing='N' "
'				addSql = addSql & " or uc.isExtUsing='N' "		''2018-12-03 ������ ���� // �����Ǹž����̶� ������� �Ǹ� �����̸� �Ǹ�
				addSql = addSql & " or i.deliveryType = 7 "
				''addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X') "
				addSql = addSql & " or i.itemdiv >= 50 "
				addSql = addSql & " or i.itemdiv = '08' "
				addSql = addSql & " or i.itemdiv = '09' "
				addSql = addSql & " or i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or not (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " or not ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				'// �ǸŰ� 1���� �̸��� ����
				'// �ּ� ������ 10%
				''addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				''addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
'				addSql = addSql & " and uc.isExtUsing='Y' "											'// 2018-12-03 ������ ���� // �����Ǹž����̶� ������� �Ǹ� �����̸� �Ǹ�
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X') "							'// �ɹ��, ȭ�����
				''addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
				addSql = addSql & " and i.itemdiv <> '08' "											'// Ƽ��(����) ��ǰ
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present��ǰ
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				'// �ǸŰ� 1���� �̸��� ����
				'// �ּ� ������ 10%
				''addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				''addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
'				addSql = addSql & " and uc.isExtUsing='Y' "											'// 2018-12-03 ������ ���� // �����Ǹž����̶� ������� �Ǹ� �����̸� �Ǹ�
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X') "							'// �ɹ��, ȭ�����
				''addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
				addSql = addSql & " and i.itemdiv <> '08' "											'// Ƽ��(����) ��ǰ
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present��ǰ
				addSql = addSql & " and i.cate_large <> '999' "										'// ī�װ� ������
				addSql = addSql & " and i.cate_large <> '' "										'// ī�װ� ������
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and i.itemdiv not in ('06', '16') "								'// �ֹ����ۻ�ǰ ����
				addSql = addSql & " and isNULL(ct.infodiv,'') not in ('','18','20','21','22') "		'// �Ϻ� ǰ��(ȭ��ǰ, ��ǰ(����깰), ������ǰ, �ǰ���ɽ�ǰ) ��ǰ
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 1000)) "		'// �ǸŰ�(���ΰ�) 1õ�� �̸�
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and o.optaddprice = 0 and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
			End If
		End If

        'Ư�� ��ǰ ����
        If (FRectIsSpecialPrice <> "") then
            If (FRectIsSpecialPrice = "Y") Then
				addSql = addSql & " and (GETDATE() > mi.startDate and GETDATE() <= mi.endDate) "
            End If
        End If

		'�ɼ��߰��ݾ�New
		If (FRectPriceOption <> "") then
			If (FRectPriceOption = "Y") Then
				addSql = addSql & " and i.itemid in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			ElseIf (FRectPriceOption = "N") Then
				addSql = addSql & " and i.itemid not in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			End If
		End If

		'������� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.nvstoremoonbanguSellYn <> 'X'"
			Else
				addSql = addSql & " and J.nvstoremoonbanguSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'��ϼ���������ǰ
		Select Case FRectFailCntExists
			Case "Y"	'����1ȸ�̻�
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'����0ȸ
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'������� ī�װ� ��Ī ����
		Select Case FRectMatchCate
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and isnull(c.CateKey, 0) <> 0"
			Case "N"	'�̸�Ī
				addSql = addSql & " and isnull(c.CateKey, 0) = 0"
		End Select

        '������� ���� < 10x10 ����
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.nvstoremoonbanguPrice is Not Null and J.nvstoremoonbanguPrice < i.sellcash"
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.nvstoremoonbanguPrice is Not Null and i.sellcash <> J.nvstoremoonbanguPrice "
		End If

		'������� �Ǹ� 10x10 ǰ��
		If (FRectNvstoremoonbanguYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.nvstoremoonbanguSellYn='Y'"
		End If

		'������� ǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectNvstoremoonbanguNo10x10Yes <> "" Then
			addSql = addSql & " and (J.nvstoremoonbanguSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.nvstoremoonbanguLastUpdate < i.lastupdate "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.nvstoremoonbanguLastUpdate, 10) "
		End If

		'��۱���
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'�ŷ�����
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		ElseIf FRectMWDiv<>"" Then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End If

		'���� ��� ����
		If (FRectIsextusing <> "") Then
			addSql = addSql & " and i.isextusing='" & FRectIsextusing & "'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as J with (nolock) "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as J with (nolock) "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.[dbo].[tbl_nvstorefarm_cate_mapping] as c with (nolock) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = 'nvstoremoonbangu' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then							'�����ٸ������� ���
				sqlStr = sqlStr & " and J.nvstoremoonbanguGoodNo is Not Null "
				If (FRectExcTrans <> "N") Then
					sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
					sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
					'sqlStr = sqlStr & " or i.isExtUsing='N' "
					'sqlStr = sqlStr & " or uc.isExtUsing='N' "
					sqlStr = sqlStr & " or i.deliveryType = 7 "
					sqlStr = sqlStr & " or i.sellyn<>'Y' "
					sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
					sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
					sqlStr = sqlStr & "	or i.itemdiv = '06' or i.itemdiv = '16' "
					sqlStr = sqlStr & " or exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
					sqlStr = sqlStr & " or exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
					sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				end if
			End If
		Else
			If (FRectExcTrans <> "N") Then
    			sqlStr = sqlStr & " and i.isusing='Y' "
    			sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    			sqlStr = sqlStr & " and i.basicimage is not null "
    			sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    			sqlStr = sqlStr & " and i.cate_large<>'' "
				sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
    			sqlStr = sqlStr & " and i.sellcash >= 1000 "
    			sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''�ֹ����� ��ǰ ���� 2013/01/15
    			'sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 �귣�� ���޻�뿩�� Y��.
			end if
		End If
		sqlStr = sqlStr & addSql
		''response.write sqlStr & "<br />"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplycash, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.nvstoremoonbanguRegdate, J.nvstoremoonbanguLastUpdate, J.nvstoremoonbanguGoodNo, J.nvstoremoonbanguPrice, J.nvstoremoonbanguSellYn, J.regUserid, IsNULL(J.nvstoremoonbanguStatCd,-9) as nvstoremoonbanguStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.CateKey, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr, J.APIaddImg "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, sc.idx as notSchIdx "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as J with (nolock) "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as J with (nolock) "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.[dbo].[tbl_nvstorefarm_cate_mapping] as c with (nolock) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = 'nvstoremoonbangu' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.nvstoremoonbanguGoodNo is Not Null "
				If (FRectExcTrans <> "N") Then
					sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
					sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
					'sqlStr = sqlStr & " or i.isExtUsing='N' "
					'sqlStr = sqlStr & " or uc.isExtUsing='N' "
					sqlStr = sqlStr & " or i.deliveryType = 7 "
					sqlStr = sqlStr & " or i.sellyn<>'Y' "
					sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
					sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
					sqlStr = sqlStr & "	or i.itemdiv = '06' or i.itemdiv = '16' "
					sqlStr = sqlStr & " or exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
					sqlStr = sqlStr & " or exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
					sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				end if
			End If
		Else
			If (FRectExcTrans <> "N") Then
    			sqlStr = sqlStr & " and i.isusing='Y' "
    			sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    			sqlStr = sqlStr & " and i.basicimage is not null "
    			sqlStr = sqlStr & " and i.itemdiv<50 and i.itemdiv<>'08' "
    			sqlStr = sqlStr & " and i.cate_large<>'' "
				sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
    			sqlStr = sqlStr & " and i.sellcash >= 1000 "
    			sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''�ֹ����� ��ǰ ���� 2013/01/15
    			'sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 �귣�� ���޻�뿩�� Y��.			'��������� isExtUsing �̰� üũ �� ��
			end if
		End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.regdate DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC"
	    End If
		''response.write sqlStr & "<br />"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CNvstoremoonbanguItem
					FItemList(i).FItemid						= rsget("itemid")
					FItemList(i).FItemname						= db2html(rsget("itemname"))
					FItemList(i).FSmallImage					= rsget("smallImage")
					FItemList(i).FMakerid						= rsget("makerid")
					FItemList(i).FRegdate						= rsget("regdate")
					FItemList(i).FLastUpdate					= rsget("lastUpdate")
					FItemList(i).FOrgPrice						= rsget("orgPrice")
					FItemList(i).ForgSuplycash					= rsget("orgSuplycash")
					FItemList(i).FSellCash						= rsget("sellcash")
					FItemList(i).FBuyCash						= rsget("buycash")
					FItemList(i).FSellYn						= rsget("sellYn")
					FItemList(i).FSaleYn						= rsget("sailyn")
					FItemList(i).FLimitYn						= rsget("LimitYn")
					FItemList(i).FLimitNo						= rsget("LimitNo")
					FItemList(i).FLimitSold						= rsget("LimitSold")
					FItemList(i).FNvstoremoonbanguRegdate		= rsget("nvstoremoonbanguRegdate")
					FItemList(i).FNvstoremoonbanguLastUpdate	= rsget("nvstoremoonbanguLastUpdate")
					FItemList(i).FNvstoremoonbanguGoodNo		= rsget("nvstoremoonbanguGoodNo")
					FItemList(i).FNvstoremoonbanguPrice			= rsget("nvstoremoonbanguPrice")
					FItemList(i).FNvstoremoonbanguSellYn		= rsget("nvstoremoonbanguSellYn")
					FItemList(i).FRegUserid						= rsget("regUserid")
					FItemList(i).FNvstoremoonbanguStatCd		= rsget("nvstoremoonbanguStatCd")
					FItemList(i).FCateMapCnt					= rsget("mapCnt")
	                FItemList(i).FDeliverytype					= rsget("deliverytype")
	                FItemList(i).FDefaultdeliverytype 			= rsget("defaultdeliverytype")
	                FItemList(i).FDefaultfreeBeasongLimit 		= rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FSmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FSmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FOptionCnt						= rsget("optionCnt")
	                FItemList(i).FRegedOptCnt					= rsget("regedOptCnt")
	                FItemList(i).FRctSellCNT					= rsget("rctSellCNT")
	                FItemList(i).FAccFailCNT					= rsget("accFailCNT")
	                FItemList(i).FLastErrStr					= rsget("lastErrStr")
	                FItemList(i).FInfoDiv						= rsget("infoDiv")
	                FItemList(i).FOptAddPrcCnt					= rsget("optAddPrcCnt")
	                FItemList(i).FOptAddPrcRegType				= rsget("optAddPrcRegType")
	                FItemList(i).FItemdiv						= rsget("itemdiv")
	                FItemList(i).FAPIaddImg						= rsget("APIaddImg")
                    FItemList(i).FSpecialPrice					= rsget("specialPrice")
					FItemList(i).FStartDate	    		  		= rsget("startDate")
					FItemList(i).FEndDate		    			= rsget("endDate")
					FItemList(i).FNotSchIdx						= rsget("notSchIdx")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' ��ϵ��� ���ƾ� �� ��ǰ..
    Public Sub getNvstoremoonbangureqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as m on i.itemid=m.itemid and m.nvstoremoonbanguGoodNo is Not Null and m.nvstoremoonbanguSellYn = 'Y' "     ''' nvstoremoonbangu �Ǹ����ΰŸ�.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_nvstorefarm_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.deliverytype in ('7') "
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From db_temp.dbo.tbl_EpShop_not_in_makerid WHERE mallgubun='"&CMALLGUBUN&"' and isusing = 'N') "	'������� �귣��
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&CMALLGUBUN&"' and isusing = 'Y') "		'������� ��ǰ
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''ȭ��ǰ, ��ǰ�� ����
        sqlStr = sqlStr & " )"

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		'�ٹ����� ��ǰ��ȣ �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		''2013/05/29 �߰�
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'')<>''"
			Elseif (FRectInfoDiv = "NN") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'')=''"
			Else
				sqlStr = sqlStr & " and ct.infodiv='"&FRectInfoDiv&"'"
			End if
		End If

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr & "	, m.nvstoremoonbanguRegdate, m.nvstoremoonbanguLastUpdate, m.nvstoremoonbanguGoodNo, m.nvstoremoonbanguPrice, m.nvstoremoonbanguSellYn, m.regUserid, m.nvstoremoonbanguStatCd "
		sqlStr = sqlStr & "	, cm.mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoremoonbangu_regItem as m on i.itemid=m.itemid and m.nvstoremoonbanguGoodNo is Not Null and m.nvstoremoonbanguSellYn = 'Y' "     ''' nvstoremoonbangu �Ǹ����ΰŸ�.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_nvstorefarm_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.deliverytype in ('7') "
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From db_temp.dbo.tbl_EpShop_not_in_makerid WHERE mallgubun='"&CMALLGUBUN&"' and isusing = 'N') "	'������� �귣��
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&CMALLGUBUN&"' and isusing = 'Y') "		'������� ��ǰ
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''ȭ��ǰ, ��ǰ�� ����
        sqlStr = sqlStr & " )"

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		'�ٹ����� ��ǰ��ȣ �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		''2013/05/29 �߰�
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'') <> ''"
			Elseif (FRectInfoDiv = "NN") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'') = ''"
			Else
				sqlStr = sqlStr & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End if
		End If
		sqlStr = sqlStr & " ORDER BY m.regdate DESC, i.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CNvstoremoonbanguItem
					FItemList(i).Fitemid						= rsget("itemid")
					FItemList(i).Fitemname						= db2html(rsget("itemname"))
					FItemList(i).FsmallImage					= rsget("smallImage")
					FItemList(i).Fmakerid						= rsget("makerid")
					FItemList(i).Fregdate						= rsget("regdate")
					FItemList(i).FlastUpdate					= rsget("lastUpdate")
					FItemList(i).ForgPrice						= rsget("orgPrice")
					FItemList(i).FSellCash						= rsget("sellcash")
					FItemList(i).FBuyCash						= rsget("buycash")
					FItemList(i).FsellYn						= rsget("sellYn")
					FItemList(i).FsaleYn						= rsget("sailyn")
					FItemList(i).FLimitYn						= rsget("LimitYn")
					FItemList(i).FLimitNo						= rsget("LimitNo")
					FItemList(i).FLimitSold						= rsget("LimitSold")
					FItemList(i).FNvstoremoonbanguRegdate		= rsget("nvstoremoonbanguRegdate")
					FItemList(i).FNvstoremoonbanguLastUpdate	= rsget("nvstoremoonbanguLastUpdate")
					FItemList(i).FNvstoremoonbanguGoodNo		= rsget("nvstoremoonbanguGoodNo")
					FItemList(i).FNvstoremoonbanguPrice			= rsget("nvstoremoonbanguPrice")
					FItemList(i).FNvstoremoonbanguSellYn		= rsget("nvstoremoonbanguSellYn")
					FItemList(i).FRegUserid						= rsget("regUserid")
					FItemList(i).FNvstoremoonbanguStatCd		= rsget("nvstoremoonbanguStatCd")
					FItemList(i).FCateMapCnt					= rsget("mapCnt")
	                FItemList(i).Fdeliverytype					= rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype			= rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit		= rsget("defaultfreeBeasongLimit")

					If Not(FItemList(i).FsmallImage = "" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FinfoDiv						= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt					= rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType				= rsget("optAddPrcRegType")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	public function GetImageByIdx(byval iGUBUN)
		Dim i
		For i=0 To FResultCount-1
			if (Not FItemList(i) is Nothing) then
				if (FItemList(i).FGubun = iGUBUN) then
					GetImageByIdx = FItemList(i).FImagename
					Exit Function
				end if
			end if
		next
    end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
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
End Class

'// ��ǰ�̹��� ���翩�� �˻�
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function rpTxt(checkvalue)
	Dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&", "&amp;")
    v = Replace(v, """", "&quot;")
    v = Replace(v, "'", "&apos;")
    v = replace(v, "<", "&lt;")
    v = replace(v, ">", "&gt;")
    v = replace(v, ":", "��")
    rpTxt = v
End Function
%>