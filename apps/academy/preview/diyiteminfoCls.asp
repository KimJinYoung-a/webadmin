<%
'#######################################################
'	Description : diy 상품관련 함수 모음
'	History	: 한용민 생성
'#######################################################

' 카테고리 상품 아이템
CLASS CCategoryPrdItem
	dim FItemID
	dim FCateCode
	dim FarrCateCd
	dim Fchkfav
	dim FItemName
	dim FSellcash
	dim FOrgPrice	
	dim FNewitem	
	dim FMakerID
	dim FBrandName
	dim FBrandName_kor
	dim FBrandUsing
	dim FItemDiv
	dim FMakerName	
	dim FMileage
	dim FSourceArea
	dim FDeliverytype	
	dim FcdL
	dim FcdM
	dim FcdS
	dim FCateName
	dim FcolorCode
	dim FcolorName	
	dim FLimitNo
	dim FLimitSold
	dim fsailprice
	dim FImageMask
	dim FImageBasic
	dim FImageList
	dim FImageList120
	dim FImageSmall
	dim FImageBasicIcon
	dim FImageIcon1	'신상품리스트, 할인리스트에서 사용(200x200)
	dim FImageIcon2
	dim FIcon2Image	
	dim FOrderComment
	dim FItemSource
	dim FItemSize
	dim FItemWeight
	dim Fkeywords
	dim FUsingHTML
	dim FItemContent	
	dim Fisusing
	dim FRegDate	
	dim FLimitYn
	dim FSellYn
	dim FItemScore
	dim Fitemgubun
	dim FSaleYn
	dim FEvalcnt
	dim FQnaCnt 
	dim FOptionCnt
	dim FNewlectureimg	''작가/강사 이미지
	dim FReipgoitemyn	
	dim Fitemcouponyn
	dim FItemCouponType
	dim FItemCouponValue
	dim FItemCouponExpire
	dim FCurrItemCouponIdx	
	dim FAvailPayType               '결제 방식 지정 0-일반 ,1-실시간(선착순) 
	dim FDefaultFreeBeasongLimit    '업체 개별배송시 배송비 무료 적용값
	dim FDefaultDeliverPay		    ' 업체 개별배송시 배송비 	
	dim FEvalcnt_Photo
	dim FfavCount
	dim FtenOnlyYn
	dim Frecentsellcount
	dim FPojangOk
	dim FImgProfile
	dim FRealKeyword

	Public FIDX
	Public FIMGTYPE
	Public FGUBUN
	Public FADDIMAGE_400
	Public FADDIMAGE_Icon
	
	public FPoints
	public Fuserid
	public Fcontents
	public FImageMain
	public FlinkURL	
	public FCurrRank
	public FLastRank
	public FplusSalePro              ''세트구매 할인율.

	public FAddimageGubun			'추가이미지 구분
	public FAddImageType			'추가이미지 형태
	public FAddimage				'추가 큰이미지
	public FAddimageSmall			'추가 작은이미지
	Public FAddimgText				'상세 텍스트
	
	''상품고시용 추가 2016-07-11 이종화
	Public FInfoname
	Public FInfoContent
	Public FinfoCode

	''동영상용
	Public FvideoUrl
	Public FvideoWidth
	Public FvideoHeight
	Public Fvideogubun
	Public FvideoType
	Public FvideoFullUrl

	''상품상세 추가
	Public Fcstodr ''즉시발송 제작후 발송
	Public Frequiremakeday ''제작후 발송기간 ? 일
	Public Frequirecontents ''특이사항
	Public Frefundpolicy ''교환/환불 정책

	Public ForderMinNum
	Public ForderMaxNum

	Public Flecturer_name
	Public Flecturer_img
	Public Flecturer_best

	public function IsStreetAvail()
		IsStreetAvail = (FBrandUsing="Y")
	end function

	'// 세일 상품 여부 '! 
	public Function IsSaleItem() 
	    IsSaleItem = ((FSaleYn="Y") and (FOrgPrice-FSellCash>0))
	end Function

	'// 상품 쿠폰 여부  '!
	public Function IsCouponItem()
			IsCouponItem = (FItemCouponYN="Y")
	end Function
			

	'// 세일포함 실제가격  '!
	public Function getRealPrice()
		getRealPrice = FSellCash
	end Function	

	'// 판매종료 여부
	public Function IsSoldOut() 
		
		'isSoldOut = (FSellYn="N")
		IF FLimitNo<>"" and FLimitSold<>"" Then
			isSoldOut = (FSellYn<>"Y") or ((FLimitYn = "Y") and (clng(FLimitNo)-clng(FLimitSold)<1))
		Else
			isSoldOut = (FSellYn<>"Y")
		End If
	end Function

	'//	한정 여부
	public Function IsLimitItem() 
		IsLimitItem= (FLimitYn="Y")
	end Function

	'//	한정 여부 (표시여부와 상관없는 실제 상품 한정여부)
	public Function IsLimitItemReal()
		IsLimitItemReal= (FLimitYn="Y")
	end Function

	'// 추가전용상품 여부
	public Function IsPlusOnlyItem() 
		IF FitemDiv="20" Then
			IsPlusOnlyItem = true
		Else
			IsPlusOnlyItem = false
		End If
	end Function

	'// 신상품 여부
	public Function IsNewItem() 
		IsNewItem =	(datediff("d",FRegdate,now())<= 14)
	end Function

	'//일시품절 여부
	public Function isTempSoldOut() 
		isTempSoldOut = (FSellYn="S")
	end Function

	'// 원 판매 가격
	public Function getOrgPrice()
		if FOrgPrice=0 then
			getOrgPrice = FSellCash
		else
			getOrgPrice = FOrgPrice
		end if
	end Function
		

	'// 한정 상품 남은 수량 '
	public Function FRemainCount()	
		if IsSoldOut then
			FRemainCount=0
		else
			FRemainCount=(clng(FLimitNo) - clng(FLimitSold))
		end if
	End Function

	'// 할인가
	public Function getDiscountPrice() 
		dim tmp

		if (FDiscountRate<>1) then
			tmp = cstr(FSellcash * FDiscountRate)
			getDiscountPrice = round(tmp / 100) * 100
		else
			getDiscountPrice = FSellcash
		end if
	end Function

	'// 할인율 '!
	public Function getSalePro() 
		if FOrgprice=0 then
			getSalePro = 0 & "%"
		else
			getSalePro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100) & "%"
		end if
	end Function

	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// 쿠폰 할인가
	public Function GetCouponDiscountPrice() 
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

	end Function

	'// 상품 쿠폰 내용
	public function GetCouponDiscountStr()
		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr = formatNumber(Fitemcouponvalue,0) + "원 할인"
			Case "3"
				GetCouponDiscountStr ="무료배송"
			Case Else
				GetCouponDiscountStr = Fitemcoupontype
		End Select
	end function

	'// 무료 배송 쿠폰 여부
	public function IsFreeBeasongCoupon() 
		IsFreeBeasongCoupon = Fitemcoupontype="3"
	end function
		

	' 사용자 등급별 무료 배송 가격
	public Function getFreeBeasongLimitByUserLevel()
		dim ulevel

		'사용자레벨에 상관없이 3만 / 업체 개별배송 5만 장바구니에서만 체크
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

	'// 무료 배송 여부
	public Function IsFreeBeasong() 
		if (getRealPrice()>=getFreeBeasongLimitByUserLevel()) then
			IsFreeBeasong = true
		else
			IsFreeBeasong = false
		end if

		if (FDeliverytype="2") or (FDeliverytype="4") or (FDeliverytype="5") then
			IsFreeBeasong = true
		end if
		
		''//착불 배송은 무료배송이 아님
		if (FDeliverytype="7") then
		    IsFreeBeasong = false
		end if
	end Function

	'// 배송구분 : 무료배송은 따로 처리  '!
	public Function GetDeliveryName()
		Select Case FDeliverytype
			Case "1" 
					GetDeliveryName="<span class='colorRd'>텐바이텐배송</span>"
			Case "2"
				if FMakerid="goodovening" then
					GetDeliveryName="<span class='colorRd'>업체배송</span>"
				else
					GetDeliveryName="<span class='colorBl'>업체무료배송</span>"
				end if
			'Case "3"
			'		GetDeliveryName="텐바이텐배송"
			Case "4"
					GetDeliveryName="<span class='colorRd'>텐바이텐배송</span>"
			Case "5"
					GetDeliveryName="<span class='colorBl'>업체무료배송</span>" 
			Case "7"
				GetDeliveryName="<span class='colorRd'>업체착불배송</span>"
			Case "9"
				if Not IsFreeBeasong then
					GetDeliveryName="<span class='colorRd'>업체조건배송</span>"
				else
					GetDeliveryName="<span class='colorBl'>업체무료배송</span>" 
				end if
			Case Else
				GetDeliveryName="텐바이텐배송"
		End Select
	end Function

	''// 업체별 배송비 부과 상품(업체 조건 배송)
	public Function IsUpcheParticleDeliverItem()
	    IsUpcheParticleDeliverItem = (FDefaultFreeBeasongLimit>0) and (FDefaultDeliverPay>0) and (FDeliveryType="9")
	end function
	
	''// 업체착불 배송여부
	public Function IsUpcheReceivePayDeliverItem()
	    IsUpcheReceivePayDeliverItem = (FDeliveryType="7")
	end function
	
	public function getDeliverNoticsStr()
	    getDeliverNoticsStr = ""
	    if (IsUpcheParticleDeliverItem) then
	        getDeliverNoticsStr = FBrandName & "(" & FBrandName_kor & ") 제품으로만" & "<br>"
	        getDeliverNoticsStr = getDeliverNoticsStr & FormatNumber(FDefaultFreeBeasongLimit,0) & "원 이상 구매시 무료배송 됩니다."
	        getDeliverNoticsStr = getDeliverNoticsStr & "배송비(" & FormatNumber(FDefaultDeliverPay,0) & "원)"
	    elseif (IsUpcheReceivePayDeliverItem) then
	        getDeliverNoticsStr = "착불 배송비는 지역에 따라 차이가 있습니다. " 
            getDeliverNoticsStr = getDeliverNoticsStr & " 상품설명의 '배송안내'를 꼭 읽어보세요." & "<br>"
	    end if
	end function    

    '// 옵션 존재여부 옵션 갯수로 체크
    public function IsItemOptionExists()
        IsItemOptionExists = (FOptioncnt>0)
    end function

	'// 무이자 이미지 & 레이어  '!
	public Function getInterestFreeImg()
			if getRealPrice>=50000 then
				getInterestFreeImg="<img src=""http://image.thefingers.co.kr/academy2010/diyshop/btn_free.gif"" width=""63"" height=""17"" align=""absmiddle"" onClick=""ShowInterestFreeImg();"" style=""cursor:pointer;"">"
				'// 2013년 1월 1일부로 모든 카드 무이자혜택 제거
				getInterestFreeImg = ""
			end if
	end Function

    ''// 세트구매 할인가격
    public function GetPLusSalePrice()
        if (FplusSalePro>0) then
            GetPLusSalePrice = getRealPrice-CLng(getRealPrice*FplusSalePro/100)
        else
            GetPLusSalePrice = getRealPrice
        end if
    end function

	'// 마일리지샵 아이템 여부 '!
	public Function IsMileShopitem() 
		IsMileShopitem = (FItemDiv="82")
	end Function
	
	Private Sub Class_Initialize()
        FplusSalePro = 0

		ForderMaxNum = 100
        ForderMinNum = 1
	End Sub

	Private Sub Class_Terminate()

	End Sub

end class
%>	