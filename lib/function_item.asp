<%
'==========================================================================
' ��ǰ���� �Լ� 
'--------------------------------------------------------------------------
	'// ���� ���밡
	public Function GetCouponAssignPrice(FItemCouponYN,Fitemcoupontype,Fitemcouponvalue,FSellCash)  
		if FItemCouponYN="Y" then
			GetCouponAssignPrice = FSellCash - GetCouponDiscountPrice(Fitemcoupontype,Fitemcouponvalue,FSellCash)
		else
			GetCouponAssignPrice = FSellCash
		end if
	end Function

	'// ���� ���ΰ�
	public Function GetCouponDiscountPrice(Fitemcoupontype,Fitemcouponvalue,FSellCash)  
		if Fitemcouponvalue="" then
			GetCouponDiscountPrice=0
			exit Function
		end if

		Select case Fitemcoupontype
			case "1" ''% ����
				if FSellCash<>"" and FSellCash<>0 then
					GetCouponDiscountPrice = CLng(Fitemcouponvalue*FSellCash/100)
				else
					GetCouponDiscountPrice=0
				end if
			case "2" ''�� ����
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''������ ����
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select 
 end Function


	'// ��ǰ ���� ����
	public function GetCouponDiscountStr(Fitemcoupontype,Fitemcouponvalue)  
		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "��"
			Case "3"
				GetCouponDiscountStr ="������"
			Case Else
				GetCouponDiscountStr = Fitemcoupontype
		End Select 
	end function
	'========================================================================
	%>