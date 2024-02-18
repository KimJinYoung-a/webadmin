<%

Class CAcademyLecOrderMasterItem
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

	public Fgoodsname
	public Fusercnt

	public Fallatdiscountprice

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

    public FweClassYn           '''2012 �߰�

	public function isWeClass() ''��ü �������� ����
	    if isNULL(FweClassYn) then
	        isWeClass = FALSE
	        Exit function
	    end if

	    isWeClass = (FweClassYn="Y")
    end function

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
	    elseif Fuserlevel="5" then
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
			IpkumDivColor="#FF00FF"
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
		elseif Faccountdiv="900" then
			JumunMethodName="����"
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
			IpkumDivName="�������"
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

	public function GetRegDate()
		GetRegDate = FRegDate
		''CStr(year(FRegDate)) + "-" + CStr(Month(FRegDate)) + "-" + CStr(Day(FRegDate)) + " " + CStR(Hour(FRegDate)) + ":" + CStR(Min(FRegDate)) + ":" + CStR(second(FRegDate))
	end function

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

Class CAcademyLecOrderMaster
	public FOneItem
	public FMasterItemList()

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

    public FRectIsWeClass

	public FRectIpkumDiv2
	public FRectIpkumDiv4

	public FRectSearchtype01
	public FRectSearchtype02

	public FRectckdate

	public Sub QuickSearchOrderList()
		dim sqlStr, i
		dim addSql

		addSql = ""

		addSql = addSql + " and m.sitename <> 'diyitem'"

		if (FRectckdate<>"") then
			addSql = addSql + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			addSql = addSql + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectSearchtype01<>"") then
			addSql = addSql + " and ( accountname like '" + FRectIpkumName + "%'"
			addSql = addSql + " or buyname = '" + FRectIpkumName + "'"
			addSql = addSql + " or reqname = '" + FRectIpkumName + "')"
		end if

		if (FRectIpkumDiv2<>"") then
			addSql = addSql + " and m.ipkumdiv='2'"
		end if

		if (FRectIpkumDiv4<>"") then
			addSql = addSql + " and m.ipkumdiv='4'"
		end if

		if (FRectSearchtype02<>"") then
			addSql = addSql + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectOrderSerial<>"") then
			addSql = addSql + " and m.orderserial='" + FRectOrderSerial + "'"
		end if

        if (FRectIsForeign<>"") then
            addSql = addSql + " and IsNULL(m.dlvcountryCode,'KR')<>'KR'"
        end if

		if (FRectRegStart<>"") then
			addSql = addSql + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			addSql = addSql + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			addSql = addSql + " and m.userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			addSql = addSql + " and m.buyname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			addSql = addSql + " and m.reqname = '" + FRectReqName + "'" ''like
		end if

		if (FRectSubTotalPrice<>"") then
			addSql = addSql + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			addSql = addSql + " and m.buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			addSql = addSql + " and m.reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			addSql = addSql + " and m.buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			addSql = addSql + " and m.reqphone='" + FRectReqPhone + "'"
		end if

		if (FRectReqSongjangNo<>"") then
			addSql = addSql + " and m.deliverno='" + FRectReqSongjangNo + "'"
		end if

		if (FRectIsFlower="Y") then
			addSql = addSql + " and m.cardribbon is Not NULL "
		end if

		if (FRectIsLecture="Y") then
			addSql = addSql + " and ((m.reqzipaddr='') or (reqzipaddr is NULL)) "
		end if

		if (FRectIsMinus="Y") then
			addSql = addSql + " and m.jumundiv='9' "
		end if

		if (FRectExtSiteName<>"") then
			addSql = addSql + " and ((m.sitename='" + FRectExtSiteName + "') or (m.rdsite='" + FRectExtSiteName + "')) "
		end if



		sqlStr = "select count(*) as cnt "
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
		IF (FRectIsWeClass="Y") then
			sqlStr = sqlStr + " Join [db_academy].[dbo].tbl_academy_order_detail dd"
			sqlStr = sqlStr + " on m.orderserial=dd.orderserial"
			sqlStr = sqlStr + " and dd.weClassyn='Y'"
		END IF
		sqlStr = sqlStr + " where 1=1"

		sqlStr = sqlStr + addSql

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		''����Ÿ.
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.* "

		sqlStr = sqlStr & " ,IsNULL((select count(itemid) as cnt from [db_academy].[dbo].tbl_academy_order_detail D where D.orderserial=m.orderserial and  D.cancelyn <> 'Y' ),0) as cnt"
		sqlStr = sqlStr & " ,isNULL((select top 1 D1.weClassyn from [db_academy].[dbo].tbl_academy_order_detail D1 where D1.orderserial=m.orderserial),'N') as weClassyn"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
		IF (FRectIsWeClass="Y") then
		    sqlStr = sqlStr + " Join [db_academy].[dbo].tbl_academy_order_detail dd"
		    sqlStr = sqlStr + " on m.orderserial=dd.orderserial"
		    sqlStr = sqlStr + " and dd.weClassyn='Y'"
		END IF
		sqlStr = sqlStr + " where 1=1"

		sqlStr = sqlStr + addSql

		sqlStr = sqlStr + " order by m.idx desc"
		''response.write sqlStr

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1


		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if not rsACADEMYget.Eof then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FMasterItemList(i) = new CAcademyLecOrderMasterItem
				FMasterItemList(i).Forderserial       	= rsACADEMYget("orderserial")
				FMasterItemList(i).Fjumundiv	        = rsACADEMYget("jumundiv")
				FMasterItemList(i).Fuserid				= rsACADEMYget("userid")
				FMasterItemList(i).Faccountname			= db2Html(rsACADEMYget("accountname"))
				FMasterItemList(i).Faccountdiv			= trim(rsACADEMYget("accountdiv"))
				FMasterItemList(i).Faccountno	        = rsACADEMYget("accountno")

				FMasterItemList(i).Ftotalmileage      	= rsACADEMYget("totalmileage")
				FMasterItemList(i).Ftotalsum	        = rsACADEMYget("totalsum")
				FMasterItemList(i).Fipkumdiv	        = rsACADEMYget("ipkumdiv")
				FMasterItemList(i).Fipkumdate	        = rsACADEMYget("ipkumdate")
				FMasterItemList(i).Fregdate				= rsACADEMYget("regdate")
				FMasterItemList(i).Fbaljudate			= rsACADEMYget("baljudate")
				FMasterItemList(i).Fbeadaldate			= rsACADEMYget("beadaldate")
				FMasterItemList(i).Fcancelyn	        = rsACADEMYget("cancelyn")

				FMasterItemList(i).Fbuyname				= db2Html(rsACADEMYget("buyname"))
				FMasterItemList(i).Fbuyphone	        = rsACADEMYget("buyphone")
				FMasterItemList(i).Fbuyhp				= rsACADEMYget("buyhp")
				FMasterItemList(i).Fbuyemail	        = rsACADEMYget("buyemail")
				FMasterItemList(i).Freqname				= db2Html(rsACADEMYget("reqname"))

				FMasterItemList(i).Freqzipcode			= rsACADEMYget("reqzipcode")
				FMasterItemList(i).Freqzipaddr			= db2Html(rsACADEMYget("reqzipaddr"))
				FMasterItemList(i).Freqaddress			= db2Html(rsACADEMYget("reqaddress"))
				FMasterItemList(i).Freqphone	        = rsACADEMYget("reqphone")
				FMasterItemList(i).Freqhp				= rsACADEMYget("reqhp")
				FMasterItemList(i).Freqemail	        = rsACADEMYget("reqemail")
				FMasterItemList(i).Fcomment				= db2Html(rsACADEMYget("comment"))

				FMasterItemList(i).Fdeliverno	        = rsACADEMYget("deliverno")

				FMasterItemList(i).Fsitename	        = rsACADEMYget("sitename")
				FMasterItemList(i).Fpaygatetid			= rsACADEMYget("paygatetid")
				FMasterItemList(i).Fdiscountrate		= rsACADEMYget("discountrate")
				FMasterItemList(i).Fsubtotalprice		= rsACADEMYget("subtotalprice")
				FMasterItemList(i).Fresultmsg			= rsACADEMYget("resultmsg")
				FMasterItemList(i).Frduserid			= rsACADEMYget("rduserid")
				FMasterItemList(i).Fmiletotalprice		= rsACADEMYget("miletotalprice")
				if IsNULL(FMasterItemList(i).Fmiletotalprice) then FMasterItemList(i).Fmiletotalprice=0

				FMasterItemList(i).Fauthcode		    = rsACADEMYget("authcode")
				FMasterItemList(i).Ftencardspend		= rsACADEMYget("tencardspend")
				FMasterItemList(i).Fuserlevel		    = rsACADEMYget("userlevel")
				FMasterItemList(i).Fspendmembership		= rsACADEMYget("spendmembership")

                FMasterItemList(i).Fallatdiscountprice 	= rsACADEMYget("allatdiscountprice")

                FMasterItemList(i).Freqdate    = rsACADEMYget("reqdate")
                FMasterItemList(i).Freqtime    = rsACADEMYget("reqtime")
                FMasterItemList(i).Fcardribbon = rsACADEMYget("cardribbon")
                FMasterItemList(i).Fmessage    = rsACADEMYget("message")
                FMasterItemList(i).Ffromname   = rsACADEMYget("fromname")

                FMasterItemList(i).Fgoodsname  = db2Html(rsACADEMYget("goodsnames"))
                FMasterItemList(i).Fusercnt    = rsACADEMYget("cnt")

                FMasterItemList(i).FDlvcountryCode = rsACADEMYget("DlvcountryCode")

                if (IsNull(FMasterItemList(i).Fallatdiscountprice) = true) then
                	FMasterItemList(i).Fallatdiscountprice = 0
                end if

                FMasterItemList(i).FweClassyn = rsACADEMYget("weClassyn")
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

	Private Sub Class_Initialize()

		Redim FMasterItemList(0)

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

%>
