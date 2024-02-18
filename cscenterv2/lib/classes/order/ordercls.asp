<%

Class COrderDetailItemMakerGroupInfoItem
	public Fgroupid
	public Fmakerid

	public Fcompany_name
	public Fcompany_no
	public Fceoname
	public Fcompany_uptae
	public Fcompany_upjong
	public Fcompany_zipcode
	public Fcompany_address
	public Fcompany_address2
	public Fcompany_tel
	public Fcompany_fax
	public Freturn_zipcode
	public Freturn_address
	public Freturn_address2
	public Fmanager_name
	public Fmanager_phone
	public Fmanager_hp
	public Fmanager_email
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fregdate
	public Flastupdate


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class




Class COrderDetailItem
    public Fidx
	public Forderserial
	public Fitemid
	public Fitemoption
	public Fmasteridx
	public Fmakerid
	public Fitemno
	public Fitemcost
	public Fmileage
	public Fcancelyn
	public Fcurrstate
	public Fsongjangno
	public Fsongjangdiv
	public Fitemname
	public Fitemoptionname
	public Fbuycash
	public Fvatinclude
	public Fbeasongdate
	public Fisupchebeasong
	public Fissailitem
	public Fupcheconfirmdate
	public Foitemdiv
    public FListImage
    public FSmallImage
    public Frequiredetail

    public Fsongjangdivname
    public Ffindurl

    '''���� ���� ���
    public function getItemcostCouponNotApplied
'        if (FitemcostCouponNotApplied<>0) then
'            getItemcostCouponNotApplied = FitemcostCouponNotApplied
'        else
'            getItemcostCouponNotApplied = FItemCost
'        end if
        getItemcostCouponNotApplied = FItemCost
    end function

    ''�ֹ����� ��ǰ
    public function IsRequireDetailExistsItem()
        IsRequireDetailExistsItem = (Foitemdiv="06") or (Frequiredetail<>"")
    end function

    public function getRequireDetailHtml()
		getRequireDetailHtml = nl2br(Frequiredetail)

		getRequireDetailHtml = replace(getRequireDetailHtml,CAddDetailSpliter,"<br><br>")
	end function

    ''�Һ��ڰ�
    public Forgprice
    public Fbonuscouponidx
    public Fitemcouponidx
    public FreducedPrice
	public FcouponNotAsigncost

    ''���� ���� �ֹ����� üũ
    public function IsBonusCouponDiscountItem()
        IsBonusCouponDiscountItem = false
        if (Not IsNull(Fbonuscouponidx) and (Fbonuscouponidx<>0))  then
            IsBonusCouponDiscountItem = true
        end if
    end function

    public function IsItemCouponDiscountItem()
        IsItemCouponDiscountItem = false
        if (Not IsNull(Fitemcouponidx) and (Fitemcouponidx<>0)) then
            IsItemCouponDiscountItem = true
        end if
    end function

	public function CancelStateStr()
		CancelStateStr = "����"

		if Fcancelyn="Y" then
			CancelStateStr ="���"
		elseif Fcancelyn="D" then
			CancelStateStr ="����"
		elseif Fcancelyn="A" then
			CancelStateStr ="�߰�"
		end if
	end function

	public function CancelStateColor()
		if FCancelYn="D" then
			CancelStateColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelStateColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelStateColor = "#000000"
		elseif UCase(FCancelYn)="A" then
			CancelStateColor = "#0000FF"
		end if
	end function

	Public function GetStateName()
        if FCurrState="2" then
            if FIsUpchebeasong="Y" then
		        GetStateName = "��ü�뺸"
		    else
		        GetStateName = "�����뺸"
		    end if
	    elseif FCurrState="3" then
		    GetStateName = "��ǰ�غ�"
	    elseif FCurrState="7" then
		    GetStateName = "���Ϸ�"
		elseif FCurrState="0" then
		    GetStateName = ""
	    else
		    GetStateName = FCurrState
	    end if
	 end Function

	public function GetStateColor()
	    if FCurrState="2" then
			GetStateColor="#000000"
		elseif FCurrState="3" then
			GetStateColor="#CC9933"
		elseif FCurrState="7" then
			GetStateColor="#FF0000"
		else
			GetStateColor="#000000"
		end if
	end function

	'���ϻ�ǰ
	public function IsSaleItem()
        'IsSaleItem = (FIsSailItem="Y") or (FplussaleDiscount>0) or (FspecialShopDiscount>0)  '''or (FIsSailItem="P")  �÷��������� �÷��� ���ϱݾ��� ������. ���� �ٲ�. 20110401 ����
        IsSaleItem = (FIsSailItem="Y")
        'IsSaleItem = IsSaleItem and (Forgitemcost>FitemcostCouponNotApplied)
    end function
	'��ǰ����
    public function IsItemCouponAssignedItem()
        'IsItemCouponAssignedItem = (Fitemcouponidx>0) and (FitemcostCouponNotApplied>FItemCost)
        IsItemCouponAssignedItem = (Fitemcouponidx>0)
    end function
	'���ʽ�����
    public function IsSaleBonusCouponAssignedItem()
        IsSaleBonusCouponAssignedItem = (Fbonuscouponidx>0)
    end function
     ''���ϸ����� ��ǰ
    public function IsMileShopSangpum()
		IsMileShopSangpum = false

'		if Foitemdiv="82" then
'			IsMileShopSangpum = true
'		end if
	end function

	'' ������ ������¸� ���� �Ѱܾ���.
    public function GetItemDeliverStateName(CurrMasterIpkumDiv, CurrMasterCancelyn)
        if ((CurrMasterCancelyn="Y") or (CurrMasterCancelyn="D") or (Fcancelyn="Y")) then
            GetItemDeliverStateName = "���"
        else
            if (CurrMasterIpkumDiv="0") then
                GetItemDeliverStateName = "��������"
            elseif (CurrMasterIpkumDiv="1") then
                GetItemDeliverStateName = "�ֹ�����"
            elseif (CurrMasterIpkumDiv="2") or (CurrMasterIpkumDiv="3") then
                GetItemDeliverStateName = "�ֹ�����"
            elseif (CurrMasterIpkumDiv="9") then
                GetItemDeliverStateName = "��ǰ"
            else
                if (IsNull(Fcurrstate) or (Fcurrstate=0)) then
            		GetItemDeliverStateName = "�����Ϸ�"
                elseif Fcurrstate="2" then
                    GetItemDeliverStateName = "�ֹ��뺸"
            	elseif Fcurrstate="3" then
            		GetItemDeliverStateName = "��ǰ�غ���"
            	elseif Fcurrstate="7" then
            		GetItemDeliverStateName = "���Ϸ�"
            	else
            		GetItemDeliverStateName = ""
            	end if
            end if
        end if
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COrderMasterItem
	public Forderserial
	public Fidx
	public Fjumundiv
	public Fuserid
	public Faccountname
	public Faccountdiv
	public Faccountno
	public Ftotalvat
	public Ftotalcost
	public Ftotalmileage
	public Ftotalsum
	public Fipkumdiv
	public Fipkumdate
	public Fregdate
	public Fbeadaldiv
	public Fbeadaldate
	public Fcancelyn
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public Freqname
	public Freqzipcode
	public Freqaddress
	public Freqphone
	public Freqhp
	public Freqemail
	public Fcomment
	public Fdeliverno
	public Fsitename
	public Fpaygatetid
	public Fdiscountrate
	public Fsubtotalprice
	public Fresultmsg
	public Frduserid
	public Fmiletotalprice
	public Fjungsanflag
	public Freqzipaddr
	public Fauthcode
	public Fsongjangdiv
	public Frdsite
	public Ftencardspend
	public Fbeasongmemo

	public FInsureCd
	public Fcashreceiptreq
	public FcashreceiptTid
	public FcashreceiptIdx
	public Finireceipttid
	public Freferip
	public Fuserlevel
	public Flinkorderserial
	public Fspendmembership
	public Fsentenceidx
	public Fbaljudate

	public Fallatdiscountprice

	'��ۺ� ���� ���ݾ�
	Public FDeliverpriceCouponNotApplied
	Public FDeliverprice

    ''�ö���ֹ� ����
    public Freqdate
	public Freqtime
	public Fcardribbon
	public Fmessage
	public Ffromname

	''�ؿܹ�۰���
	public FDlvcountryCode

	public FcountryNameKr
	public FcountryNameEn
	public FemsAreaCode
    public FemsZipCode
    public FitemGubunName
    public FgoodNames
    public FitemWeigth
    public FitemUsDollar
    public FemsInsureYn
    public FemsInsurePrice
    public FemsDlvCost

    ''OkCashbag �߰�
    public FokcashbagSpend

	Public FspendTenCash
	Public Fspendgiftmoney
	public Forgorderserial

    '''�ְ������� �ݾ� = subtotalPrice-FsumPaymentEtc
    public function TotalMajorPaymentPrice()
        TotalMajorPaymentPrice = FsubtotalPrice-FsumPaymentEtc
    end function

    ''2016/08/18 �߰�
    public FsumPaymentEtc
    public FPgGubun

    ''������ ������� ��������
    public function IsDacomCyberAccountPay()
        IsDacomCyberAccountPay = false
        if (FAccountdiv<>"7") then Exit function

        if (FAccountNo="���� 470301-01-014754") _
            or (FAccountNo="���� 100-016-523130") _
            or (FAccountNo="�츮 092-275495-13-001") _
            or (FAccountNo="�ϳ� 146-910009-28804") _
            or (FAccountNo="��� 277-028182-01-046") _
            or (FAccountNo="���� 029-01-246118") then
                IsDacomCyberAccountPay = false
        else
            IsDacomCyberAccountPay = true
        end if
    end function

	''�ؿܹ����������
	public function IsForeignDeliver()
        IsForeignDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode<>"") and (FDlvcountryCode<>"KR") and (FDlvcountryCode<>"ZZ")
    end function

    ''���δ���
    public function IsArmiDeliver()
        IsArmiDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode="ZZ")
    end function

    public function IsErrSubtotalPrice()
        IsErrSubtotalPrice = (Fsubtotalprice <> (Ftotalsum - (Ftencardspend + Fmiletotalprice + Fspendmembership + Fallatdiscountprice)))
    end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

    ''�����ߴ��� ����
    public function IsPayedOrder()
        IsPayedOrder = (FIpkumdiv>3) and (FIpkumdiv<9)
    end function

	'�������ɿ���
    public function IsReceiveSiteOrder
        IsReceiveSiteOrder = (Fjumundiv="7")
    end Function

    public function GetMasterDeliveryName()
        GetMasterDeliveryName = ""
        if IsNULL(Fsongjangdiv) then Exit function

        if Fsongjangdiv="24" then
            GetMasterDeliveryName = "�簡��"
        elseif Fsongjangdiv="2" then
            GetMasterDeliveryName = "����"
        else
            GetMasterDeliveryName = Fsongjangdiv
        end if
    end function

	public function GetUserLevelColor()
		if Fuserlevel="1" then
			GetUserLevelColor = "#f0ca2c"   ''Green
		elseif Fuserlevel="2" then
			GetUserLevelColor = "#a3cf6c"   ''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelColor = "#6ca54e"   ''VIP
		elseif Fuserlevel="4" then
			GetUserLevelColor = "#f68d3f"   ''������
		elseif Fuserlevel="5" then
			GetUserLevelColor = "#865e25"  '' ���ο�
		elseif Fuserlevel="6" then
			GetUserLevelColor = "#B70606"  '' staff
		else
			GetUserLevelColor = "#f0ca2c"
		end if
	end function

	public function GetUserLevelName()
		if Fuserlevel="1" then
			GetUserLevelName = "Seed"
		elseif Fuserlevel="2" then
			GetUserLevelName = "Bud"
		elseif Fuserlevel="3" then
			GetUserLevelName = "Leaf"
		elseif Fuserlevel="4" then
			GetUserLevelName = "Bean"
	    elseif Fuserlevel="5" and FUserID<>"" then
			GetUserLevelName = "Tree"
		elseif Fuserlevel="6" then
			GetUserLevelName = "STAFF"
		else
			GetUserLevelName = "Seed"
		end if
	end function

	public function GetJumunDivName()
		if Fjumundiv="1" then
			GetJumunDivName = "���ֹ�"
		elseif Fjumundiv="3" then
			GetJumunDivName = "�����ֹ�"
		elseif Fjumundiv="5" then
			GetJumunDivName = "�ܺθ�"
		elseif Fjumundiv="6" then
			GetJumunDivName = "��ī����DIY��ǰ"
		elseif Fjumundiv="7" then
			GetJumunDivName = "�ö��"
		elseif Fjumundiv="8" then
			GetJumunDivName = "�����ֹ�"
		elseif Fjumundiv="9" then
			GetJumunDivName = "���̳ʽ�"
		else
			GetJumunDivName = Fjumundiv
		end if
	end function


	public function CancelYnName()
		CancelYnName = "����"

		if Fcancelyn="Y" then
			CancelYnName ="���"
		elseif Fcancelyn="D" then
			CancelYnName ="����"
		elseif Fcancelyn="A" then
			CancelYnName ="�߰�"
		end if
	end function

	public function CancelYnColor()
		CancelYnColor = "#000000"

		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		end if
	end function


	public function IpkumDivColor()
		if Fipkumdiv="0" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="1" then
			IpkumDivColor="#44BBBB"
		elseif Fipkumdiv="2" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			IpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			IpkumDivColor="#CC9933"
		elseif Fipkumdiv="6" then
			IpkumDivColor="#FF00FF"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="9" then
			IpkumDivColor="#FF0000"
		end if
	end function

	Public function JumunMethodName()
		if Faccountdiv="7" then
			JumunMethodName="������"
		elseif Faccountdiv="100" then
			JumunMethodName="�ſ�ī��"
		elseif Faccountdiv="20" then
			JumunMethodName="�ǽð���ü"
		elseif Faccountdiv="30" then
			JumunMethodName="����Ʈ"
		elseif Faccountdiv="50" then
			JumunMethodName="����������"
		elseif Faccountdiv="80" then
			JumunMethodName="All@ī��"
		elseif Faccountdiv="90" then
			JumunMethodName="��ǰ�ǰ���"
		elseif Faccountdiv="110" then
			JumunMethodName="OK+�ſ�"
		elseif Faccountdiv="400" then
			JumunMethodName="�ڵ�������"
		end if
	end function

	Public function IpkumDivName()
		if Fipkumdiv="0" then
			IpkumDivName="�ֹ����"
		elseif Fipkumdiv="1" then
			IpkumDivName="�ֹ�����"
		elseif Fipkumdiv="2" then
			IpkumDivName="�ֹ�����"
		elseif Fipkumdiv="3" then
			IpkumDivName="�ֹ�����(3)"
		elseif Fipkumdiv="4" then
			IpkumDivName="�����Ϸ�"
		elseif Fipkumdiv="5" then
			IpkumDivName="�ֹ��뺸"
		elseif Fipkumdiv="6" then
			IpkumDivName="��ǰ�غ�"
		elseif Fipkumdiv="7" then
			IpkumDivName="�Ϻ����"
	    elseif Fipkumdiv="8" then
			IpkumDivName="��ǰ���"
		else
			IpkumDivName=Fipkumdiv
		end if
	end Function

	Public function NormalUpcheDeliverState()
		 if IsNull(FCurrState) then
			 NormalUpcheDeliverState = "�����Ϸ�"
		 elseif FCurrState="3" then
			 NormalUpcheDeliverState = "��ǰ�غ�"
		 elseif FCurrState="7" then
			 NormalUpcheDeliverState = "��ǰ���"
		 else
			 NormalUpcheDeliverState = ""
		 end if
	 end Function

	public function UpCheDeliverStateColor()
		if IsNull(FCurrState) then
			UpCheDeliverStateColor="#3300CC"
		elseif FCurrState="3" then
			UpCheDeliverStateColor="#0000FF"
		elseif FCurrState="7" then
			UpCheDeliverStateColor="#FF0000"
		else
			UpCheDeliverStateColor="#000000"
		end if
	end function


	public function SiteNameColor()
		if Fsitename<>"10x10" then
			SiteNameColor = "#55AA22"
		else
			SiteNameColor = "#000000"
		end if
	end function


	public function SubTotalColor()
		if FSubtotalPrice<0 then
			SubTotalColor = "#DD3333"
		else
			SubTotalColor = "#000000"
		end if
	end function

    ''�ö�� ������ ��� �ֹ� ���翩��
    public function IsFixDeliverItemExists()
        IsFixDeliverItemExists = Not IsNULL(Freqdate)
    end function

    '' �ö�� ������ �ð�
    public function GetReqTimeText()
        if IsNULL(Freqtime) then Exit function
        GetReqTimeText = Freqtime & "~" & (Freqtime+2) & "�� ��"
    end function

	Private Sub Class_Initialize()
        FokcashbagSpend = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CUpcheBeasongPayItem

	public Fmakerid
	public Fdefaultfreebeasonglimit
	public Fdefaultdeliverpay

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COrderMaster
	public FOneItem
	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectOrderSerial
	public FRectUserID
	public FRectBuyname
	public FRectReqName
	public FRectIpkumName
	public FRectSubTotalPrice

	public FRectBuyHp
	public FRectReqHp
	public FRectBuyPhone
	public FRectReqPhone
	public FRectReqSongjangNo

	public FRectRegStart
	public FRectRegEnd

	public FRectExtSiteName
	public FRectIsMinus
	public FRectIsLecture
	public FRectIsFlower

    public FRectOldOrder
    public FRectDetailIdx
    public FRectIsForeign

	Public FTotItemNo
	public FTotItemKind

    ''detail query ��
    public function GetItemCostSum()

    end function

    public function GetImageFolderName(byval itemid)
		GetImageFolderName = "0" + CStr(Clng(itemid\10000))
	end function

	public function BeasongCD2Name(byval v)
		if v="0101" then
			BeasongCD2Name = "�Ϲ��ù�"
		elseif v="0201" then
			BeasongCD2Name = "������A"
		elseif v="0202" then
			BeasongCD2Name = "������B"
		elseif v="0203" then
			BeasongCD2Name = "������C"
		elseif v="0301" then
			BeasongCD2Name = "��������"
		elseif v="0501" then
			BeasongCD2Name = "������"
		end if
	end function

	public function BeasongPay()
		dim i
		for i=0 to FResultCount-1
			if FItemList(i).FItemID=0 then
				BeasongPay = FItemList(i).Fitemcost
				Exit For
			end if
		next
	end Function

	public function BeasongOptionStr()
		dim i
		for i=0 to FResultCount-1
			if FItemList(i).FItemID=0 then
				BeasongOptionStr = BeasongCD2Name(FItemList(i).Fitemoption)
				Exit For
			end if
		next
	end function

	public Sub QuickSearchOrderList()
		dim sqlStr, i
		''����
		sqlStr = "select count(*) as cnt "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		else
    		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
    	end if
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and sitename <> '" + CStr(EXCLUDE_SITENAME) + "'"

		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and orderserial='" + FRectOrderSerial + "'"
		end if

        if (FRectIsForeign<>"") then
            sqlStr = sqlStr + " and IsNULL(dlvcountryCode,'KR')<>'KR'"
        end if

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			sqlStr = sqlStr + " and buyname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			sqlStr = sqlStr + " and reqname = '" + FRectReqName + "'"  ''like
		end if

		if (FRectIpkumName<>"") then
			sqlStr = sqlStr + " and accountname = '" + FRectIpkumName + "'" ''like
		end if

		if (FRectSubTotalPrice<>"") then
			sqlStr = sqlStr + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			sqlStr = sqlStr + " and buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlStr = sqlStr + " and reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlStr = sqlStr + " and buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlStr = sqlStr + " and reqphone='" + FRectReqPhone + "'"
		end if

		if (FRectReqSongjangNo<>"") then
			sqlStr = sqlStr + " and deliverno='" + FRectReqSongjangNo + "'"
		end if

		if (FRectIsFlower="Y") then
			sqlStr = sqlStr + " and cardribbon is Not NULL "
		end if

		if (FRectIsLecture="Y") then
			sqlStr = sqlStr + " and ((reqzipaddr='') or (reqzipaddr is NULL)) "
		end if

		if (FRectIsMinus="Y") then
			sqlStr = sqlStr + " and jumundiv='9' "
		end if

		if (FRectExtSiteName<>"") then
			sqlStr = sqlStr + " and ((sitename='" + FRectExtSiteName + "') or (rdsite='" + FRectExtSiteName + "')) "
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close

		''����Ÿ.
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.* "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		else
		    sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		end if
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and sitename <> '" + CStr(EXCLUDE_SITENAME) + "'"

		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and orderserial='" + FRectOrderSerial + "'"
		end if

        if (FRectIsForeign<>"") then
            sqlStr = sqlStr + " and IsNULL(dlvcountryCode,'KR')<>'KR'"
        end if

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			sqlStr = sqlStr + " and buyname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			sqlStr = sqlStr + " and reqname = '" + FRectReqName + "'" ''like
		end if

		if (FRectIpkumName<>"") then
			sqlStr = sqlStr + " and accountname = '" + FRectIpkumName + "'" ''like
		end if

		if (FRectSubTotalPrice<>"") then
			sqlStr = sqlStr + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			sqlStr = sqlStr + " and buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlStr = sqlStr + " and reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlStr = sqlStr + " and buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlStr = sqlStr + " and reqphone='" + FRectReqPhone + "'"
		end if

		if (FRectReqSongjangNo<>"") then
			sqlStr = sqlStr + " and deliverno='" + FRectReqSongjangNo + "'"
		end if

		if (FRectIsFlower="Y") then
			sqlStr = sqlStr + " and cardribbon is Not NULL "
		end if

		if (FRectIsLecture="Y") then
			sqlStr = sqlStr + " and ((reqzipaddr='') or (reqzipaddr is NULL)) "
		end if

		if (FRectIsMinus="Y") then
			sqlStr = sqlStr + " and jumundiv='9' "
		end if

		if (FRectExtSiteName<>"") then
			sqlStr = sqlStr + " and ((sitename='" + FRectExtSiteName + "') or (rdsite='" + FRectExtSiteName + "')) "
		end if

        'if (FRectBuyname<>"") or (FRectReqName<>"") or (FRectIpkumName<>"") or (FRectSubTotalPrice<>"") or (FRectBuyHp<>"") or (FRectReqHp<>"") or (FRectBuyPhone<>"") or (FRectReqPhone<>"") or (FRectReqSongjangNo<>"") then
        'sqlStr = sqlStr + " order by orderserial desc"
        'else
		sqlStr = sqlStr + " order by m.idx desc"
	    'end if
		''response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderMasterItem
				FItemList(i).Forderserial       = rsget("orderserial")
				FItemList(i).Fjumundiv	        = rsget("jumundiv")
				FItemList(i).Fuserid			= rsget("userid")
				FItemList(i).Faccountname		= db2Html(rsget("accountname"))
				FItemList(i).Faccountdiv		= trim(rsget("accountdiv"))
				FItemList(i).Faccountno	        = rsget("accountno")

				FItemList(i).Ftotalmileage      = rsget("totalmileage")
				FItemList(i).Ftotalsum	        = rsget("totalsum")
				FItemList(i).Fipkumdiv	        = rsget("ipkumdiv")
				FItemList(i).Fipkumdate	        = rsget("ipkumdate")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fbeadaldate		= rsget("beadaldate")
				FItemList(i).Fcancelyn	        = rsget("cancelyn")

				FItemList(i).Fbuyname			= db2Html(rsget("buyname"))
				FItemList(i).Fbuyphone	        = rsget("buyphone")
				FItemList(i).Fbuyhp				= rsget("buyhp")
				FItemList(i).Fbuyemail	        = rsget("buyemail")
				FItemList(i).Freqname			= db2Html(rsget("reqname"))

				FItemList(i).Freqzipcode		= rsget("reqzipcode")
				FItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
				FItemList(i).Freqaddress		= db2Html(rsget("reqaddress"))
				FItemList(i).Freqphone	        = rsget("reqphone")
				FItemList(i).Freqhp				= rsget("reqhp")
				FItemList(i).Freqemail	        = rsget("reqemail")
				FItemList(i).Fcomment			= db2Html(rsget("comment"))

				FItemList(i).Fdeliverno	        = rsget("deliverno")

				FItemList(i).Fsitename	        = rsget("sitename")
				FItemList(i).Fpaygatetid		= rsget("paygatetid")
				FItemList(i).Fdiscountrate		= rsget("discountrate")
				FItemList(i).Fsubtotalprice		= rsget("subtotalprice")
				FItemList(i).Fresultmsg			= rsget("resultmsg")
				FItemList(i).Frduserid			= rsget("rduserid")
				FItemList(i).Fmiletotalprice	= rsget("miletotalprice")
				if IsNULL(FItemList(i).Fmiletotalprice) then FItemList(i).Fmiletotalprice=0

				FItemList(i).Fauthcode		        = rsget("authcode")
				FItemList(i).Ftencardspend			= rsget("tencardspend")
				FItemList(i).Fuserlevel		        = rsget("userlevel")
				FItemList(i).Fspendmembership		= rsget("spendmembership")

                FItemList(i).Fallatdiscountprice 	= rsget("allatdiscountprice")

                FItemList(i).Freqdate    = rsget("reqdate")
                FItemList(i).Freqtime    = rsget("reqtime")
                FItemList(i).Fcardribbon = rsget("cardribbon")
                FItemList(i).Fmessage    = rsget("message")
                FItemList(i).Ffromname   = rsget("fromname")

                FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")
                
                if (IsNull(FItemList(i).Fallatdiscountprice) = true) then
                	FItemList(i).Fallatdiscountprice = 0
                end if
                
                ''2016/09/09
                FItemList(i).Frdsite	= rsget("rdsite")
                FItemList(i).FsumPaymentEtc = rsget("sumPaymentEtc")
                FItemList(i).FPgGubun       = rsget("pggubun")
    
                if isNULL(FItemList(i).FsumPaymentEtc) then FItemList(i).FsumPaymentEtc=0
                if isNULL(FItemList(i).FPgGubun) then FItemList(i).FPgGubun=""
                    
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub



	public Sub QuickSearchOrderMaster()
		dim sqlStr, i

		sqlStr = "select top 1 m.* "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		else
		    sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		end if
		sqlStr = sqlStr + " where m.idx<>0"

		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			sqlStr = sqlStr + " and buyname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			sqlStr = sqlStr + " and reqname = '" + FRectReqName + "'" ''like
		end if

		if (FRectIpkumName<>"") then
			sqlStr = sqlStr + " and accountname ='" + FRectIpkumName + "'" ''like
		end if

		if (FRectSubTotalPrice<>"") then
			sqlStr = sqlStr + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			sqlStr = sqlStr + " and buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlStr = sqlStr + " and reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlStr = sqlStr + " and buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlStr = sqlStr + " and reqphone='" + FRectReqPhone + "'"
		end if

		if (FRectReqSongjangNo<>"") then
			sqlStr = sqlStr + " and deliverno='" + FRectReqSongjangNo + "'"
		end if

		sqlStr = sqlStr + " order by orderserial desc"
        ''sqlStr = sqlStr + " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if not rsget.Eof then
		        FTotalCount = 1
		end if

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

        if not rsget.Eof then
	        set FOneItem = new COrderMasterItem

			FOneItem.FspendTenCash	= 0
			FOneItem.Fspendgiftmoney	= 0
			FOneItem.Forgorderserial	= ""
			FOneItem.Forderserial           = rsget("orderserial")
			FOneItem.Fjumundiv	            = rsget("jumundiv")
			FOneItem.Fuserid		        = rsget("userid")
			FOneItem.Faccountname	        = db2Html(rsget("accountname"))
			FOneItem.Faccountdiv	        = trim(rsget("accountdiv"))
			FOneItem.Faccountno	            = rsget("accountno")

			FOneItem.Ftotalmileage          = rsget("totalmileage")
			FOneItem.Ftotalsum	            = rsget("totalsum")
			FOneItem.Fipkumdiv	            = rsget("ipkumdiv")
			FOneItem.Fipkumdate	            = rsget("ipkumdate")
			FOneItem.Fregdate		        = rsget("regdate")
			FOneItem.Fbaljudate		        = rsget("baljudate")
			FOneItem.Fbeadaldate	        = rsget("beadaldate")
			FOneItem.Fcancelyn	            = rsget("cancelyn")
			FOneItem.Fbuyname		        = db2Html(rsget("buyname"))
			FOneItem.Fbuyphone	            = rsget("buyphone")
			FOneItem.Fbuyhp		            = rsget("buyhp")
			FOneItem.Fbuyemail	            = rsget("buyemail")
			FOneItem.Freqname		        = db2Html(rsget("reqname"))
			FOneItem.Freqzipcode	        = rsget("reqzipcode")
			FOneItem.Freqaddress	        = db2Html(rsget("reqaddress"))
			FOneItem.Freqphone	            = rsget("reqphone")
			FOneItem.Freqhp		            = rsget("reqhp")
			FOneItem.Freqemail	            = rsget("reqemail")
			FOneItem.Fcomment		        = db2Html(rsget("comment"))
			FOneItem.Fdeliverno	            = rsget("deliverno")
			FOneItem.Fsitename	            = rsget("sitename")
			FOneItem.Fpaygatetid	        = rsget("paygatetid")
			FOneItem.Fdiscountrate	        = rsget("discountrate")
			FOneItem.Fsubtotalprice	        = rsget("subtotalprice")
			FOneItem.Fresultmsg		        = rsget("resultmsg")
			FOneItem.Frduserid		        = rsget("rduserid")
			FOneItem.Fmiletotalprice	    = rsget("miletotalprice")

			FOneItem.FInsureCd           	= rsget("InsureCd")

			if IsNULL(FOneItem.Fmiletotalprice) then FOneItem.Fmiletotalprice=0

			FOneItem.Fjungsanflag		    = rsget("jungsanflag")
			FOneItem.Freqzipaddr		    = db2Html(rsget("reqzipaddr"))
			FOneItem.Fauthcode		        = rsget("authcode")
			FOneItem.Fcashreceiptreq		= rsget("cashreceiptreq")

			FOneItem.Ftencardspend		    = rsget("tencardspend")

			FOneItem.Fuserlevel		        = rsget("userlevel")
			FOneItem.Fspendmembership	    = rsget("spendmembership")
			FOneItem.Fallatdiscountprice    = rsget("allatdiscountprice")

			FOneItem.Freqdate    = rsget("reqdate")
            FOneItem.Freqtime    = rsget("reqtime")
            FOneItem.Fcardribbon = rsget("cardribbon")
            FOneItem.Fmessage    = rsget("message")
            FOneItem.Ffromname   = rsget("fromname")

            FOneItem.FDlvcountryCode = rsget("DlvcountryCode")
            FOneItem.Frdsite	= rsget("rdsite")

            ''2016/08/18
            FOneItem.FsumPaymentEtc = rsget("sumPaymentEtc")
            FOneItem.FPgGubun       = rsget("pggubun")

            if isNULL(FOneItem.FsumPaymentEtc) then FOneItem.FsumPaymentEtc=0
            if isNULL(FOneItem.FPgGubun) then FOneItem.FPgGubun=""


	    end if
		rsget.Close

		if (FResultCount>0) then
    		if (FOneItem.Faccountdiv="110") then
    		    sqlStr = "select IsNULL(sum(acctamount),0) as okcashbagSpend"
    			sqlStr = sqlStr + "	from db_order.dbo.tbl_order_paymentEtc"
    			sqlStr = sqlStr + "	where orderserial='"&FRectOrderSerial&"'"
    			sqlStr = sqlStr + "	and acctdiv='110'"
    			rsget.Open sqlStr,dbget,1
    			if not rsget.Eof then
    		        FOneItem.FokcashbagSpend = rsget("okcashbagSpend")
    		    end if
    		    rsget.close
    		end if
    	end if
	end sub

	public Sub QuickSearchOrderDetail()
		dim sqlStr
		dim i

		sqlStr = "select d." & FIELD_DETAILIDX & " as idx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost,d.reducedPrice"
		sqlStr = sqlStr + " ,d.mileage,d.cancelyn "
		sqlStr = sqlStr + " ,d.itemname, d.makerid, i.listimage "
		sqlStr = sqlStr + " ,i.smallimage , i.orgprice, d.itemoptionname "
		sqlStr = sqlStr + " ,d.currstate, d.upcheconfirmdate, d.songjangdiv, d.songjangno"
		sqlStr = sqlStr + " ,d.beasongdate, d.isupchebeasong, d.issailitem, d.requiredetail  "
		sqlStr = sqlStr + " ,d.issailitem, d.bonuscouponidx, d." & FIELD_ITEMCOUPONIDX & " as itemcouponidx "
		sqlStr = sqlStr + " ,s.divname as songjangdivname, s.findurl, d.couponNotAsigncost, d.buycash"
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d "
		else
		    sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d "
		end if
		sqlStr = sqlStr + "     left join " & TABLE_ITEM & " i on d.itemid=i.itemid"
		sqlStr = sqlStr + "     left join " & TABLE_SONGJANG_DIV & " s on d.songjangdiv=s.divcd"
		sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"
        sqlStr = sqlStr + " order by d.isupchebeasong, d.makerid, d.itemid, d.itemoption"

        ''response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new COrderDetailItem

			FItemList(i).Forderserial = CStr(FRectOrderSerial)
			FItemList(i).Fidx         = rsget("idx")
			FItemList(i).Fmakerid     = rsget("makerid")
			FItemList(i).Fitemid      = rsget("itemid")
			FItemList(i).Fitemoption  = rsget("itemoption")
			FItemList(i).Fitemno      = rsget("itemno")
			FItemList(i).Fitemcost    = rsget("itemcost")
			FItemList(i).Fmileage     = rsget("mileage")
			FItemList(i).Fcancelyn    = rsget("cancelyn")

			FItemList(i).FItemName    = db2html(rsget("itemname"))
			FItemList(i).FSmallImage  = webImgUrl + DIRECTORY_IMAGE_SMALL + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

			if IsNull(rsget("itemoptionname")) then
				FItemList(i).FItemoptionName = "-"
			else
				FItemList(i).FItemoptionName = db2html(rsget("itemoptionname"))
			end if

			FItemList(i).Fcurrstate         = rsget("currstate")
			FItemList(i).Fsongjangdiv       = rsget("songjangdiv")
			FItemList(i).Fsongjangno        = rsget("songjangno")
			FItemList(i).Fbeasongdate       = rsget("beasongdate")
			FItemList(i).Fisupchebeasong    = rsget("isupchebeasong")
			FItemList(i).Fissailitem        = rsget("issailitem")
			FItemList(i).Fupcheconfirmdate    = rsget("upcheconfirmdate")

			FItemList(i).Frequiredetail    = rsget("requiredetail")
            FItemList(i).Fsongjangdivname  = db2html(rsget("songjangdivname"))
            FItemList(i).Ffindurl          = db2html(rsget("findurl"))

            FItemList(i).Forgprice          = rsget("orgprice")
            FItemList(i).Fissailitem        = rsget("issailitem")
            FItemList(i).Fbonuscouponidx    = rsget("bonuscouponidx")
            FItemList(i).Fitemcouponidx     = rsget("itemcouponidx")

            FItemList(i).FreducedPrice      	= rsget("reducedPrice")
			FItemList(i).FcouponNotAsigncost	= rsget("couponNotAsigncost")
			FItemList(i).Fbuycash      			= rsget("buycash")
            if Not IsNULL(FItemList(i).Fsongjangno) then
               FItemList(i).Fsongjangno = replace(FItemList(i).Fsongjangno,"-","")
            end if

			IF FItemList(i).Fitemid <> 0 THEN
				FTotItemNo = FTotItemNo + FItemList(i).Fitemno
				FTotItemKind = FTotItemKind + 1
			END IF

			rsget.movenext
			i=i+1
		loop
		rsget.close
	end sub

    public function GetOneOrderDetail
        dim sqlStr, i
	    dim mastertable, detailtable

	    if (FRectOldOrder<>"") then
			mastertable = "[db_log].[dbo].tbl_old_order_master_2003"
			detailtable	= "[db_log].[dbo].tbl_old_order_detail_2003"
		else
			mastertable = "" & TABLE_ORDERMASTER & ""
			detailtable	= "" & TABLE_ORDERDETAIL & ""
		end if

		sqlStr =	" SELECT d.idx, d.itemid, d.itemoption, d.itemno, d.itemoptionname, d.itemcost," &_
					" d.itemname, d.itemcost, d.makerid, d.currstate, replace(d.songjangno,'-','') as songjangno, d.songjangdiv," &_
					" d.cancelyn, d.isupchebeasong, d.mileage, d.requiredetail, d.oitemdiv, d.beasongdate, d.issailitem, d.upcheconfirmdate," &_
					" d.bonuscouponidx, d.itemcouponidx, d.reducedPrice," &_
					" i.smallimage, i.listimage, i.brandname, i.itemdiv, i.orgprice" &_
					" ,s.divname,s.findurl ,s.tel as DeliveryTel" &_
					" FROM " + detailtable + " d " &_
					" JOIN " & TABLE_ITEM & " i" &_
					"		ON d.itemid=i.itemid " &_
					" LEFT JOIN " & TABLE_SONGJANG_DIV & " s " &_
					"		ON d.songjangdiv = s.divcd " &_
					" WHERE d.orderserial='" + FRectOrderserial + "'" &_
					" and d.idx=" & FRectDetailIdx &_
					" and d.itemid<>0" &_
					" and d.cancelyn<>'Y'" &_
					" order by i.deliverytype"
		rsget.Open sqlStr,dbget,1

		FTotalcount = rsget.Recordcount
		FResultcount = FTotalcount


        if Not rsget.Eof then
			set FOneItem = new COrderDetailItem
			FOneItem.Forderserial = CStr(FRectOrderSerial)
			FOneItem.Fidx         = rsget("idx")
			FOneItem.Fmakerid     = rsget("makerid")
			FOneItem.Fitemid      = rsget("itemid")
			FOneItem.Fitemoption  = rsget("itemoption")
			FOneItem.Fitemno      = rsget("itemno")
			FOneItem.Fitemcost    = rsget("itemcost")
			FOneItem.Fmileage     = rsget("mileage")
			FOneItem.Fcancelyn    = rsget("cancelyn")

			FOneItem.FItemName    = db2html(rsget("itemname"))
			FItemList(i).FSmallImage  = webImgUrl + DIRECTORY_IMAGE_SMALL + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

			if IsNull(rsget("itemoptionname")) then
				FOneItem.FItemoptionName = "-"
			else
				FOneItem.FItemoptionName = db2html(rsget("itemoptionname"))
			end if

			FOneItem.Fcurrstate         = rsget("currstate")
			FOneItem.Fsongjangdiv       = rsget("songjangdiv")
			FOneItem.Fsongjangno        = rsget("songjangno")
			FOneItem.Fbeasongdate       = rsget("beasongdate")
			FOneItem.Fisupchebeasong    = rsget("isupchebeasong")
			FOneItem.Fissailitem        = rsget("issailitem")
			FOneItem.Fupcheconfirmdate    = rsget("upcheconfirmdate")

			FOneItem.Frequiredetail    = rsget("requiredetail")
            FOneItem.Fsongjangdivname  = db2html(rsget("divname"))
            FOneItem.Ffindurl          = db2html(rsget("findurl"))

            FOneItem.Forgprice          = rsget("orgprice")
            FOneItem.Fissailitem        = rsget("issailitem")
            FOneItem.Fbonuscouponidx    = rsget("bonuscouponidx")
            FOneItem.Fitemcouponidx     = rsget("itemcouponidx")

            FOneItem.FreducedPrice      = rsget("reducedPrice")
            if Not IsNULL(FOneItem.Fsongjangno) then
               FOneItem.Fsongjangno = replace(FOneItem.Fsongjangno,"-","")
            end if

		end if
		rsget.close
    end function

    public function getEmsOrderInfo()
        dim sqlStr
        sqlStr = " exec [db_order].[dbo].sp_Ten_OneEmsOrderInfo '" & FRectOrderserial & "'"

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
            FOneItem.FcountryNameEn   = rsget("countryNameEn")
            FOneItem.FemsAreaCode     = rsget("emsAreaCode")
            FOneItem.FemsZipCode      = rsget("emsZipCode")
            FOneItem.FitemGubunName   = rsget("itemGubunName")
            FOneItem.FgoodNames       = rsget("goodNames")
            FOneItem.FitemWeigth      = rsget("itemWeigth")
            FOneItem.FitemUsDollar    = rsget("itemUsDollar")
            FOneItem.FemsInsureYn     = rsget("InsureYn")
            FOneItem.FemsInsurePrice  = rsget("InsurePrice")

            FOneItem.FemsDlvCost       = rsget("emsDlvCost")
		end if
		rsget.Close
    end function

    '���� �ְ����ݾ�(+�ſ�ī�� ��Ұ��� ����)
	public Sub getMainPaymentInfo(byval paymethod, byref orgpayment, byref cardcancelok, byref cardcancelerrormsg, byref cardcancelcount, byref cardcancelsum, byref cardcode)
		dim sqlStr

		dim remailpayment, payetcresult
		dim jumundiv, orgorderserial, pggubun
		dim tmpArr

		orgpayment = 0
		cardcancelok = "N"
		cardcancelerrormsg = ""
		cardcancelcount = ""
		cardcode = ""

		'// ��ȯ�ֹ�( jumundiv = 6 )�̸� ���ֹ����� �������� �����´�.
		sqlStr = " select top 1 m.jumundiv, m.pggubun "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_academy.dbo.tbl_academy_order_master m "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.orderserial = '" & FRectOrderserial & "' "
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			jumundiv = rsget("jumundiv")
			pggubun  = rsget("pggubun")
		end if
		rsget.close

		if (jumundiv = "6") then
			sqlStr = " select top 1 c.orgorderserial "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_academy.dbo.tbl_academy_change_order c "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and c.chgorderserial = '" & FRectOrderserial & "' "
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				orgorderserial = rsget("orgorderserial")
			end if
			rsget.close
		else
			orgorderserial = FRectOrderserial
		end if

		sqlStr = " select top 1 IsNull(e.acctamount, 0) as orgpayment, IsNull(e.realPayedSum, 0) as remailpayment, IsNull(e.PayEtcResult, '') as payetcresult "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_academy_order_PaymentEtc e "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and e.orderserial = '" & orgorderserial & "' "
		sqlStr = sqlStr + " 	and e.acctdiv in ('7', '100', '550', '560', '20', '50', '80', '90', '400', '110') "							'OK CASH BAG �� �ְ��������̴�.

        'response.write sqlStr &"<br>"
        IF (paymethod="110") then
            sqlStr = " select sum(IsNull(e.acctamount, 0)) as orgpayment, sum(IsNull(e.realPayedSum, 0)) as remailpayment, '' as payetcresult "
    		sqlStr = sqlStr + " from "
    		sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_academy_order_PaymentEtc e "
    		sqlStr = sqlStr + " where "
    		sqlStr = sqlStr + " 	1 = 1 "
    		sqlStr = sqlStr + " 	and e.orderserial = '" & orgorderserial & "' "
    		sqlStr = sqlStr + " 	and e.acctdiv in ('100', '110') "
        END IF

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			orgpayment = rsget("orgpayment")
			remailpayment = rsget("remailpayment")
			payetcresult = rsget("payetcresult")

			if Len(payetcresult) = 9 and UBound(Split(payetcresult, "|")) = 3 then
				'// 14|26|0|1 => 14|26|00|1
				tmpArr = Split(payetcresult, "|")
				payetcresult = tmpArr(0) & "|" & tmpArr(1) & "|" & "0" & tmpArr(2) & "|" & tmpArr(3)
			end if
		end if
		rsget.close

        '' ���̹� ���� ���� �߰� (����Ʈ)
        if (pggubun="NP") then
            sqlStr = " select top 1 IsNull(e.acctamount, 0) as orgpayment, IsNull(e.realPayedSum, 0) as remailpayment, IsNull(e.PayEtcResult, '') as payetcresult "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_academy_order_PaymentEtc e "
            sqlStr = sqlStr + " where "
            sqlStr = sqlStr + " 	1 = 1 "
            sqlStr = sqlStr + " 	and e.orderserial = '" & orgorderserial & "' "
            sqlStr = sqlStr + " 	and e.acctdiv='120'"

            rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
            	orgpayment = orgpayment + rsget("orgpayment")
            	remailpayment = remailpayment + rsget("remailpayment")

            	if Len(payetcresult) = 7 and UBound(Split(payetcresult, "|")) = 3 then
            		'// 14||0|1 => 14|26|00|1
            		tmpArr = Split(payetcresult, "|")
            		payetcresult = tmpArr(0) & "|" & "XX" & "|" & "0" & tmpArr(2) & "|" & tmpArr(3)
            	end if
            end if
            rsget.close

        end if

		if (paymethod <> "100") then
			if (paymethod = "110") then
				cardcancelerrormsg = "OK+�ſ�(���� �κ���ҺҰ�)"
			elseif (paymethod = "20") and (pggubun="NP") then                              ''2016/07/21 �߰�
			    cardcancelok = "Y"
				cardcancelcount = 0
				cardcode = payetcresult
			elseif (paymethod = "20") then
			    cardcancelok = "Y"
				cardcancelcount = 0
				cardcode = payetcresult
			else
				cardcancelerrormsg = "�ſ�ī����� �ƴ�(getMainPaymentInfo)"
			end if
		else
			if (orgpayment = 0) or (payetcresult = "") then
				cardcancelerrormsg = "�ſ�ī������ ����"
			else
				cardcancelok = "Y"
				cardcancelcount = 0
				cardcode = payetcresult
			end if
		end if

        cardcancelcount = 0
        cardcancelsum   = 0
		if (cardcancelok = "Y") and (orgpayment <> remailpayment) then
			sqlStr = " select count(orderserial) as cnt, sum(cancelprice) as canceltotal "
			sqlStr = sqlStr + " from db_academy.dbo.tbl_academy_card_cancel_log "
			sqlStr = sqlStr + " where orderserial = '" & orgorderserial & "' and resultcode in ('00', '2001') "  '''0000' �ٽ� ���� 2016/07/21 eastone �ڵ� '00' ���� �ٲ�
			rsget.Open sqlStr,dbget,1

			if Not rsget.Eof then
				cardcancelcount = rsget("cnt")
				cardcancelsum   = rsget("canceltotal")
			end if
			rsget.close

			'9ȸ���� �κ���Ұ� ���������� ������ ���� 1���� ���ܳ��´�.
			if (cardcancelcount >= 8) then
				cardcancelok = "N"
				cardcancelerrormsg = "�κ���� Ƚ�� �ʰ�"
			end if
		end if

		if (cardcancelok = "Y") then
		    if (paymethod <> "100") then
		        ''�ǽð� ��ü.TEST
		    else
    		    '' cardcode �� ���ڸ��� Ȯ�� ����.
    		    if (LEN(cardcode)<10) then
    		        cardcancelok = "N"
    		        if (cardcancelerrormsg="") then cardcancelerrormsg  = "�κ���� <strong>�Ұ�</strong> �ŷ�"
    		    end if

    		    if (Right(cardcode,1)<>"Y") then
    		        cardcancelok = "N"
                    if (cardcancelerrormsg="") then cardcancelerrormsg  = "�κ���� <strong>�Ұ�</strong> �ŷ�"
    		    end if
		    end if
		end if

	end sub

	public Sub getUpcheBeasongPayList()
		dim sqlStr
		dim i

		sqlStr = " select distinct "
		sqlStr = sqlStr + " 	d.makerid, IsNull(b.defaultfreebeasonglimit, 0) as defaultfreebeasonglimit, IsNull(b.defaultdeliverypay, 0) as defaultdeliverpay "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	" & TABLE_ORDERDETAIL & " d "
		sqlStr = sqlStr + " 	join db_academy.dbo.tbl_lec_user b "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		d.makerid = b.lecturer_id "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and d.orderserial = '" & FRectOrderserial & "' "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and d.isupchebeasong <> 'N' "

        'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CUpcheBeasongPayItem

			FItemList(i).Fmakerid     					= rsget("makerid")
			FItemList(i).Fdefaultfreebeasonglimit     	= rsget("defaultfreebeasonglimit")
			FItemList(i).Fdefaultdeliverpay     		= rsget("defaultdeliverpay")

			if (FItemList(i).Fdefaultdeliverpay = 0) then
				FItemList(i).Fdefaultdeliverpay = 2500
			end if

			rsget.movenext
			i = i + 1
		loop
		rsget.close
	end sub

	Private Sub Class_Initialize()

		Redim FItemList(0)

		FCurrPage =1
		FPageSize = 20
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

end Class

public Function getOrgPayPrice(orderserial)
	dim sqlStr
	dim i, result

	getOrgPayPrice = 0

	sqlStr = " select top 1 e.acctamount "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_academy].[dbo].[tbl_academy_order_master] m "
	sqlStr = sqlStr + " 	join [db_academy].[dbo].[tbl_academy_order_PaymentEtc] e "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and m.orderserial = e.orderserial "
	sqlStr = sqlStr + " 		and m.accountdiv = e.acctdiv "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	m.orderserial = '" & orderserial & "' "

    'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1

	if Not rsget.Eof then
		getOrgPayPrice = rsget("acctamount")
	end if
	rsget.Close
end Function

%>
