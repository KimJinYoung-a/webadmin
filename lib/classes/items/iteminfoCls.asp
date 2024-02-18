<%
'#######################################################
'	History	: ������ ����
'			  2008.03.18 ������ ���� - Ŭ���� �и�
'			  2008.04.13 �ѿ�� �߰�
'             2008.08.27 ������ ��ü ���� ��� ���� �߰�
'	Description :��ǰ���� �Լ� ����
'#######################################################

'#=========================================#
'# ī�װ� ��ǰ ������                    #
'#=========================================#

CLASS CCategoryPrdItem

	'// �ʼ� ����  //

	dim FItemID
	dim FItemName
	dim FSellcash
	dim FOrgPrice
	dim fEval_excludeyn
	dim FNewitem

	dim FMakerID
	dim FBrandName
	dim FBrandName_kor
	dim FBrandLogo
	dim FBrandUsing
	dim FisBestBrand
	dim FUserDiv

	dim FItemDiv
	dim FMakerName
	dim FOrgMakerID

	dim FMileage
	dim FSourceArea
	dim FDeliverytype

	dim FcdL
	dim FcdM
	dim FcdS
	dim FcateCode
	dim FCateName
	dim FcateCd1
	dim FcateCd2
	dim FcateCd3
	dim FcateDepth
	dim FarrCateCd

	dim Freviewcnt


	dim FcolorCode
	dim FcolorName

	dim FLimitNo
	dim FLimitSold
	dim fsailprice
	dim FImageBasic
	dim FImageBasic600		'600px�̹���
	dim FImageBasic1000		'1000px�̹���
	dim FImageMask
	dim FImageMask1000		'1000px�̹���
	dim FImageList
	dim FImageList120
	dim FImageSmall
	dim FImageBasicIcon
	dim FImageMaskIcon
	dim FImageIcon1	'�Ż�ǰ����Ʈ, ���θ���Ʈ���� ���(200x200)
	dim FImageIcon2
	dim FImageIcon3
	dim FImageIcon4
	dim FImageIcon5
	dim FIcon1Image
	dim FIcon2Image

	'// ���� �⺻ �̹��� �߰�(2015.01.21 ������)
	Dim Ftentenimage
	Dim Ftentenimage50
	Dim Ftentenimage200
	Dim Ftentenimage400
	Dim Ftentenimage600
	Dim Ftentenimage1000

	'// ��ǰ�󼼼��� ������ �߰�(2016.02.17 ������)
	Dim FvideoUrl
	Dim FvideoWidth
	Dim FvideoHeight
	Dim Fvideogubun
	Dim FvideoType
	Dim FvideoFullUrl


	dim FOrderComment
	dim Fdeliverarea
	dim FItemSource
	dim FItemSize
	dim FItemWeight
	dim FdeliverOverseas

	dim Fkeywords
	dim FUsingHTML
	dim FItemContent

	dim Fisusing
	dim FStreetUsing

	dim FRegDate

	dim FReipgodate
	dim FSpecialbrand


	dim Fdgncomment
	dim FDesignerComment

	dim FLimitYn
	dim FSellYn
	dim FItemScore

	dim Fitemgubun

	dim FSaleYn
	dim FTenOnlyYn		'�ٹ����� ������ǰ����(2011.04.14)

	dim FEvalcnt
	dim FEvalcnt_Photo
	dim FfavCount
	dim FQnaCnt
	dim FOptionCnt
	dim FAvgDlvDate

	dim FAddimageGubun
	dim FAddimageSmall
	dim FAddImageType
	dim FAddimage
	dim FAddimage600
	dim FAddimage1000
	dim FIsExistAddimg

	dim Ffreeprizeyn '?

	dim FReipgoitemyn
	dim FSpecialUserItem

	dim Fitemcouponyn
	dim FItemCouponType
	dim FItemCouponValue
	dim FItemCouponExpire
	dim FCurrItemCouponIdx

	dim FAvailPayType               '���� ��� ���� 0-�Ϲ� ,1-�ǽð�(������)
	dim FDefaultFreeBeasongLimit    '��ü ������۽� ��ۺ� ���� ���밪
	dim FDefaultDeliverPay		    ' ��ü ������۽� ��ۺ�
	dim FRequireMakeDay				'�ֹ����ۻ�ǰ�� ���� �ҿ���(2011.04.14)

	Dim FsafetyYN		'�����������
	Dim FsafetyDiv		'������������ '10 ~ 50
	Dim FsafetyNum	'����������ȣ

	public FPoints
	public FPoint_fun
	public FPoint_dgn
	public FPoint_prc
	public FPoint_stf
	public Fuserid
	public Fcontents
	public FImageMain
	public FImageMain2			'��ǰ����2 �̹��� �߰�(2011.04.14)
	public FImageMain3			'��ǰ����3 �̹��� �߰�(2013.07.31)
	public FlinkURL

	public FCurrRank
	public FLastRank

	public FPojangOk			'�������� ���� ����

	public FBRWriteRegdate		'����Ʈ�����
	public FUseGood
	public FUseETC

	public FplusSalePro			''��Ʈ���� ������.
	public FisJust1day			'Just 1day ��ǰ ����

	'��Ÿ�϶�������
	public FStyleCd1
	public FStyleCd1Nm
	public FStyleCd2
	public FStyleCd2Nm
	public FStyleCd3
	public FStyleCd3Nm
	public fOrderNo

	'hotcateitem 2012-04-04
	Public Fidx
	Public Fitemseq
	Public Fcdmname
	Public Fcdsname
	Public Fsailyn

	'��ǰ�� �߰� 2012-11-01
	Public FInfoname
	Public FInfoContent
	Public FinfoCode

	Public ForderMinNum
	Public ForderMaxNum

	'2013 ������ ī�װ����ο�
	Public FDisp
	Public Ftype
	Public Fcode
	Public Ftitle
	Public Fsubcopy
	Public Fimgurl
	Public Ficon

	'2013 popular wish
	Public FInCount
	Public FRegTime
	Public FEvaluate
	Public FMyCount
	
	'/�귣�� ��������
	public fdetailidx
	public fmasteridx
	public fsortNo
	public Flastupdate
	public fregadminid
	public flastadminid
	public fevt_code

	'/2014 Gift
	public FtalkCnt
	public FdayCnt
	public FthemeCnt
	
	'/��ǰ���߰�
	public FLimitDispYn
	
	public fdevice
	public Fsdate
	public Fedate

	'/2015 �� �ֹ� ��ǰ
	public Forderserial
	public ForderDate
	public ForderOption
	public ForderOptionName
	public ForderCnt

	'�귣�� ���� �߰�2017-01-31 ���¿�
	public FBrandNoticeGubun
	public FBrandNoticeTitle
	public FBrandNoticeText

	'�÷������� �ɼ� ����
	Public FOptionTypeName
	Public FOptionName
	Public FOptionAddPrice
	Public FOptionCode
	
	'�� ��ǥ��ǰ �ڵ� �߰�
	Public FDealMasterItemID
	Public FItemOptionCnt

	'/��Ű����
	public Frecentsellcount
	
	public function IsRookieItem()
		IsRookieItem = false
		if (Not IsNewItem) then Exit function
		
		IsRookieItem = (Frecentsellcount>=20)
	end function

	public function IsStreetAvail() ' !
		IsStreetAvail = (FStreetUsing="Y") and (Fuserdiv<10)
	end function


	'// �� �Ǹ� ����  '!
	public Function getOrgPrice()
		if FOrgPrice=0 then
			getOrgPrice = FSellCash
		else
			getOrgPrice = FOrgPrice
		end if
	end Function

	'// �������� ��������  '!
	public Function getRealPrice()

		getRealPrice = FSellCash


		if (IsSpecialUserItem()) then
			getRealPrice = getSpecialShopItemPrice(FSellCash)
		end if
	end Function

	'//��ǰ�ڵ�  '!
	public Function FProductCode()
		 FProductCode = formatCode(FItemid)
	end Function

	'// ��ǰ��
	public Function getCuttingItemName()
		if Len(FItemName)>18 then
			getCuttingItemName=Left(FItemName,18) + "..."
		else
			getCuttingItemName=FItemName
		end if
	end Function

	'// ��ǰ ����  '?
	public Function GetCuttingItemContents()
		''## �̻��� �߶����.
		dim reStr
		reStr = LeftB(Fitemcontent,120)
		reStr = replace(reStr,"<P>","")
		reStr = replace(reStr,"<p>","")
		reStr = replace(reStr,"<br>",Chr(2))
		reStr = Left(reStr,100)
		reStr = replace(reStr,Chr(2),"&nbsp;")
		GetCuttingItemContents = reStr + "..."
	end Function

	'// ���ȸ���� ��ǰ ���� '!
	public Function IsSpecialUserItem()
	    dim uLevel
	  '  uLevel = GetLoginUserLevel()
		IsSpecialUserItem = (FSpecialUserItem>0) and (uLevel>1 and uLevel<>5)
	end Function

	'// �Ǹ�����  ���� '! '2008/07/07 �߰�
	public Function IsSoldOut()

		'isSoldOut = (FSellYn="N")
		IF FLimitNo<>"" and FLimitSold<>"" Then
			isSoldOut = (FSellYn<>"Y") or ((FLimitYn = "Y") and (clng(FLimitNo)-clng(FLimitSold)<1))
		Else
			isSoldOut = (FSellYn<>"Y")
		End If
	end Function

	'// �� �Ǹ�����  ���� '! '2017/11/17 �߰�
	public Function isDealSoldout() 
		isDealSoldout = (FSellYn="N")
	end Function

	'//�Ͻ�ǰ�� ���� '2008/07/07 �߰� '!
	public Function isTempSoldOut()

		isTempSoldOut = (FSellYn="S")

	end Function

	'// ���� ��ǰ ���� '!
	public Function IsSaleItem()
	    IsSaleItem = ((FSaleYn="Y") and (FOrgPrice-FSellCash>0)) or (IsSpecialUserItem)
	end Function

	'//	���� ���� '!
	public Function IsLimitItem()
			IsLimitItem= (FLimitYn="Y") and (FLimitDispYn="Y" or isNull(FLimitDispYn))
	end Function

	'//	���� ���� (ǥ�ÿ��ο� ������� ���� ��ǰ ��������)
	public Function IsLimitItemReal()
			IsLimitItemReal= (FLimitYn="Y")
	end Function

	'// �Ż�ǰ ���� '!
	public Function IsNewItem()
			IsNewItem =	(datediff("d",FRegdate,now())<= 14)
	end Function

	'// ���� ��� ���� ���� '?
	public function IsFreeBeasongCoupon()
		IsFreeBeasongCoupon = Fitemcoupontype="3"
	end function

	'// ��ǰ ���� ����  '!
	public Function IsCouponItem()
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// ����ǰ ���� ��ǰ ���� '?
	public Function IsGiftItem()
			IsGiftItem	= (FFreePrizeYN ="Y")
	end Function

	'// ���԰� ��ǰ ����
	public Function isReipgoItem()
		isReipgoItem = (datediff("d",FReIpgoDate,now())<= 14)
	end Function

	'// ���ϸ����� ������ ���� '!
	public Function IsMileShopitem()
		IsMileShopitem = (FItemDiv="82")
	end Function

	'// �ٹ����� ������ǰ ���� '!
	public Function IsTenOnlyitem()
		IsTenOnlyitem = (FTenOnlyYn="Y")
	end Function

	'// �ٹ����� ���尡�� ��ǰ ����
	public Function IsPojangitem()
		IsPojangitem = (FPojangOk="Y" and IsTenBeasong)
	end Function

	'// ���� ��ǰ ���� ���� '!
	public Function FRemainCount()
		if IsSoldOut then
			FRemainCount=0
		else
			FRemainCount=(clng(FLimitNo) - clng(FLimitSold))
		end if
	End Function

	'// ��ǰ ���� �ޱ� '!
	public Function IsSpecialBrand()
		IsSpecialBrand = FSpecialBrand="Y"
	End Function

	'// ���ΰ�
	public Function getDiscountPrice()
		dim tmp

		if (FDiscountRate<>1) then
			tmp = cstr(FSellcash * FDiscountRate)
			getDiscountPrice = round(tmp / 100) * 100
		else
			getDiscountPrice = FSellcash
		end if
	end Function

	'// ������ '!
	public Function getSalePro()
		if FOrgprice=0 then
			getSalePro = 0 & "%"
		else
			getSalePro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100) & "%"
		end if
	end Function

	'// ���� ���밡
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// ���� ���ΰ� '?
	public Function GetCouponDiscountPrice()
		Select case Fitemcoupontype
			case "1" ''% ����
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''�� ����
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''������ ����
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

	end Function

	'// ��ǰ ���� ����  '!
	public function GetCouponDiscountStr()

		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr = formatNumber(Fitemcouponvalue,0) + "�� ����"
			Case "3"
				GetCouponDiscountStr ="������"
			Case Else
				GetCouponDiscountStr = Fitemcoupontype
		End Select

	end function


	public function GetCouponDiscountStr_new()

		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr_new =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr_new = formatNumber(Fitemcouponvalue,0) + "�� ����"
			Case "3"
				GetCouponDiscountStr_new =""
			Case Else
				GetCouponDiscountStr_new = Fitemcoupontype
		End Select

	end function


	'// ���� ��� ����
	public Function IsFreeBeasong()
		if (getRealPrice()>=getFreeBeasongLimitByUserLevel()) then
			IsFreeBeasong = true
		else
			IsFreeBeasong = false
		end if

		if (FDeliverytype="2") or (FDeliverytype="4") or (FDeliverytype="5") or (FDeliverytype="6") then
			IsFreeBeasong = true
		end if

		''//���� ����� �������� �ƴ�
		if (FDeliverytype="7") then
		    IsFreeBeasong = false
		end if
	end Function

	'// �ؿ� ��� ����(�ٹ� + �ؿܿ��� + ��ǰ����)
	public Function IsAboardBeasong()
		if FdeliverOverseas="Y" and FItemWeight>0 and (FDeliverytype="1" or FDeliverytype="3" or FDeliverytype="4") then
			IsAboardBeasong = true
		else
			IsAboardBeasong = false
		end if
	end function

	'// �ٹ����� ��� ����
	public Function IsTenBeasong()
		IsTenBeasong = false
		if (FDeliverytype="1" or FDeliverytype="3" or FDeliverytype="4") then
			IsTenBeasong = true
		end if
	end function

	''// ��ü�� ��ۺ� �ΰ� ��ǰ(��ü ���� ���)
	public Function IsUpcheParticleDeliverItem()
	    IsUpcheParticleDeliverItem = (FDefaultFreeBeasongLimit>0) and (FDefaultDeliverPay>0) and (FDeliveryType="9")
	end function

	''// ��ü���� ��ۿ���
	public Function IsUpcheReceivePayDeliverItem()
	    IsUpcheReceivePayDeliverItem = (FDeliveryType="7")
	end function

	public function getDeliverNoticsStr()
	    getDeliverNoticsStr = ""
	    if (IsUpcheParticleDeliverItem) then
	        getDeliverNoticsStr = FBrandName & "(" & FBrandName_kor & ") ��ǰ���θ�" & "<br>"
	        getDeliverNoticsStr = getDeliverNoticsStr & FormatNumber(FDefaultFreeBeasongLimit,0) & "�� �̻� ���Ž� ������ �˴ϴ�."
	        getDeliverNoticsStr = getDeliverNoticsStr & "��ۺ�(" & FormatNumber(FDefaultDeliverPay,0) & "��)"
	    elseif (IsUpcheReceivePayDeliverItem) then
	        getDeliverNoticsStr = "���� ��ۺ�� ������ ���� ���̰� �ֽ��ϴ�. "
            getDeliverNoticsStr = getDeliverNoticsStr & " ��ǰ������ '��۾ȳ�'�� �� �о����." & "<br>"
	    end if
	end function

	' ����� ��޺� ���� ��� ����  '?
	public Function getFreeBeasongLimitByUserLevel()
		dim ulevel

		''���ο����� ����ڷ����� ������� 3�� / ��ü ������� 5�� ��ٱ��Ͽ����� üũ
		if (FDeliverytype="9") then
		    If (IsNumeric(FDefaultFreeBeasongLimit)) and (FDefaultFreeBeasongLimit<>0) then
		        getFreeBeasongLimitByUserLevel = FDefaultFreeBeasongLimit
		    else
		        getFreeBeasongLimitByUserLevel = 50000
		    end if
		else
		    getFreeBeasongLimitByUserLevel = 30000
		end if

	end Function

    '// �ɼ� ���翩�� �ɼ� ������ üũ
    public function IsItemOptionExists()
        IsItemOptionExists = (FOptioncnt>0)
    end function

	'// ��۱��� : �������� ���� ó��  '!
	public Function GetDeliveryName()
		Select Case FDeliverytype
			Case "1"
				GetDeliveryName="�ٹ����ٹ��"
			Case "2"
				if FMakerid="goodovening" then
					GetDeliveryName="��ü���"
				else
					GetDeliveryName="��ü������"
				end if
			'Case "3"
			'		GetDeliveryName="�ٹ����ٹ��"
			Case "4"
					GetDeliveryName="�ٹ����ٹ��"
			Case "5"
					GetDeliveryName="��ü������"
			Case "6"
					GetDeliveryName="������ɻ�ǰ"
			Case "7"
				GetDeliveryName="��ü���ҹ��"
			Case "9"
				if Not IsFreeBeasong then
					GetDeliveryName="��ü���ǹ��"
				else
					GetDeliveryName="��ü������"
				end if
			Case Else
				GetDeliveryName="�ٹ����ٹ��"
		End Select
	end Function


	'// ������ �̹��� & ���̾�  '!
	public Function getInterestFreeImg()
			if getRealPrice>=50000 then
				getInterestFreeImg="<div class='clicklayer' class='relative'>" & vbCrLf &_
									"<img class='btn img' src='http://fiximage.10x10.co.kr/web2012/product/product_desc_title03_1.png' style='cursor:pointer'/>" & vbCrLf &_
									"	<div class='layer credit-card'>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_ld.gif' /></div> <span class='black_11px_bold'>�Ե�ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_sh.gif' /></div> <span class='black_11px_bold'>����ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_hd.gif' /></div> <span class='black_11px_bold'>����ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_keb.gif' /></div> <span class='black_11px_bold'>����ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_bc.gif' /></div> <span class='black_11px_bold'>��ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_ss.gif' /></div> <span class='black_11px_bold'>�Ｚī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_kb.gif' /></div> <span class='black_11px_bold'>����ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf &_
									"	</div>" & vbCrLf &_
									"</div>"
				'// 2013�� 1�� 1�Ϻη� ��� ī�� ���������� ����
				getInterestFreeImg = ""

				'//2013�� 1,2�� ������ �ȳ�
				if date()>="2013-01-07" and date()<="2013-02-28" then
					getInterestFreeImg="<div class='clicklayer' class='relative'>" & vbCrLf
					getInterestFreeImg= getInterestFreeImg & "<img class='btn img' src='http://fiximage.10x10.co.kr/web2012/product/product_desc_title03_1.png' style='cursor:pointer'/>" & vbCrLf
					getInterestFreeImg= getInterestFreeImg & "	<div class='layer credit-card' style='border:3px solid #DDD;'>" & vbCrLf
					if date()>="2013-01-07" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_ss.gif' /></div> <span class='black_11px_bold'>�Ｚī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf
					if date()>="2013-01-09" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_sh.gif' /></div> <span class='black_11px_bold'>����ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf
					if date()>="2013-01-11" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_ld.gif' /></div> <span class='black_11px_bold'>�Ե�ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf
					if date()>="2013-01-11" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_hd.gif' /></div> <span class='black_11px_bold'>����ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf
					if date()>="2013-01-12" and date()<="2013-02-28" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_kb.gif' /></div> <span class='black_11px_bold'>����ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf
					if date()>="2013-01-12" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_keb.gif' /></div> <span class='black_11px_bold'>��ȯī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf
					if date()>="2013-01-12" and date()<="2013-02-28" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_nh.gif' /></div> <span class='black_11px_bold'>����ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf
					if date()>="2013-02-01" and date()<="2013-02-28" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_bc.gif' /></div> <span class='black_11px_bold'>��ī��</span>&nbsp;5������ / 2,3����<br/>" & vbCrLf
					getInterestFreeImg= getInterestFreeImg & "	</div>" & vbCrLf
					getInterestFreeImg= getInterestFreeImg & "</div>"
				end if
			end if
	end Function


    ''// ��Ʈ���� ���ΰ���
    public function GetPLusSalePrice()
        if (FplusSalePro>0) then
            GetPLusSalePrice = getRealPrice-CLng(getRealPrice*FplusSalePro/100)
        else
            GetPLusSalePrice = getRealPrice
        end if
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
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2009/bestaward/award_up.gif' width='7' height='4' align='absmiddle'> <font class='verdanared'><b>" & GetLevelUpCount() & "</b></font>"
		elseif (FCurrRank=FLastRank) and (FLastRank=0) then
			GetLevelUpArrow = ""
			'##���� GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2008/award/s_arrow_new.gif' width='9' height='5'>"
		elseif (FCurrRank=FLastRank) then
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2009/bestaward/award_none.gif' width='6' height='2' align='absmiddle'> <font class='eng11px00'><b>0</b></font>"
		elseif (FCurrRank>FLastRank) and (FLastRank=0) then
			GetLevelUpArrow = ""
			'##���� GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2008/award/s_arrow_new.gif' width='9' height='5'>"
		else
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2009/bestaward/award_down.gif' width='7' height='4' align='absmiddle'> <font class='verdanabk'><b>" & GetLevelUpCount() & "</b></font>"
			if FCurrRank-FLastRank>=FCurrPos then
				GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2009/bestaward/award_none.gif' width='6' height='2' align='absmiddle'> <font class='eng11px00'><b>0</b></font>"
			end if
		end if
	end Function

	public function isBestRankItem()
		isBestRankItem = false
		if not(FCurrRank="" or isNull(FCurrRank)) then
			if FCurrRank<=1000 then
				isBestRankItem = true
			end if
		end if
	end function

	'// ������������ ����
	public Function IsSafetyYN()
		if FsafetyYN="Y"  then
			IsSafetyYN = true
		else
			IsSafetyYN = false
		end if
	end Function

	'// ������������ ��ũ
	public Function IsSafetyDIV()
		if FsafetyDIV="10"  then
			IsSafetyDIV = "������������(KC��ũ)"
		ElseIf FsafetyDIV="20"  then
			IsSafetyDIV = "�����ǰ ��������"
		ElseIf FsafetyDIV="30"  then
			IsSafetyDIV = "KPS �������� ǥ��"
		ElseIf FsafetyDIV="40"  then
			IsSafetyDIV = "KPS �������� Ȯ�� ǥ��"
		ElseIf FsafetyDIV="50"  then
			IsSafetyDIV = "KPS ��� ��ȣ���� ǥ��"
		end if
		
		'### �Ǽ��� ��� �ҽ����߷��� �������ϴ�.
		'ElseIf FsafetyDIV="60"  then
		'	IsSafetyDIV = "KCC����(��MIC����)"
	end function
	
	
	public Function fnRealAllPrice()
		'####### ���� ���� ��� �� ����Ͽ� 1������ ��Ÿ��. ����&���� �� ������ ����.
		Dim vPrice
		vPrice = FSellCash
		IF FSaleyn = "Y" AND FItemcouponyn = "Y" Then
			vPrice = GetCouponAssignPrice
		Else
			If FItemcouponyn = "Y" Then
				vPrice = GetCouponAssignPrice
			End If
		End If
		fnRealAllPrice = vPrice
	End Function

    ''�����ǰ //2016/04/15 �߰�
    public function IsTravelItem()
        IsTravelItem = False
        if FItemDiv="18" then
			IsTravelItem = true
		end if
    end function

	Private Sub Class_Initialize()
        FplusSalePro = 0
        Frecentsellcount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

end CLASS
%>
