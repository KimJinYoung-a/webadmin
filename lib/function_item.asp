<%
'==========================================================================
' 상품관련 함수 
'--------------------------------------------------------------------------
	'// 쿠폰 적용가
	public Function GetCouponAssignPrice(FItemCouponYN,Fitemcoupontype,Fitemcouponvalue,FSellCash)  
		if FItemCouponYN="Y" then
			GetCouponAssignPrice = FSellCash - GetCouponDiscountPrice(Fitemcoupontype,Fitemcouponvalue,FSellCash)
		else
			GetCouponAssignPrice = FSellCash
		end if
	end Function

	'// 쿠폰 할인가
	public Function GetCouponDiscountPrice(Fitemcoupontype,Fitemcouponvalue,FSellCash)  
		if Fitemcouponvalue="" then
			GetCouponDiscountPrice=0
			exit Function
		end if

		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				if FSellCash<>"" and FSellCash<>0 then
					GetCouponDiscountPrice = CLng(Fitemcouponvalue*FSellCash/100)
				else
					GetCouponDiscountPrice=0
				end if
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select 
 end Function


	'// 상품 쿠폰 내용
	public function GetCouponDiscountStr(Fitemcoupontype,Fitemcouponvalue)  
		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "원"
			Case "3"
				GetCouponDiscountStr ="무료배송"
			Case Else
				GetCouponDiscountStr = Fitemcoupontype
		End Select 
	end function
	'========================================================================
	%>