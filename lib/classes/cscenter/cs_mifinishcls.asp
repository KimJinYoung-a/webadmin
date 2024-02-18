<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2023.11.15 �ѿ�� ����(����Ʃ��)
'###########################################################

class CCSMifinishDetail
	public Fdivcd
	public FOrderserial
	public Fasid
	public FBuyname
	public FReqName
	public FItemID
	public FItemname
	public FItemno
	public FItemoption
	public FItemoptionname
	public FCurrstate
	public FSongjangno
	public FSongjangdiv
	public FIdx
	public Fdeleteyn
	public FMakerID
	public FOrderDate
	public FIpkumdiv
	public FDeliverytype
	public FMasterCancel
	public Fdeliverno
	public Fcsdetailidx
	public Fsongjangyn
	public FItemcnt
	public FJumunDiv
	public FBuyCash
	public FSellcash
	public Fmasteridx
	public FRegdate
    public FMifinishReason
    public FMifinishState
	public FMifinishipgodate
    public FMifinishregdate
	public Freguserid
	public Flastupdate

    public function getDivcdStr()
        select Case Fdivcd
            CASE "A000" : getDivcdStr="���"
            CASE "A100" : getDivcdStr="���"
            CASE "A004" : getDivcdStr="��ǰ"
            CASE ELSE : getDivcdStr = Fdivcd
        end Select
    end function

    public function getDivcdColor()
        select Case Fdivcd
            CASE "A000" : getDivcdColor="blue"
            CASE "A100" : getDivcdColor="blue"
            CASE "A004" : getDivcdColor="red"
            CASE ELSE : getDivcdColor = "black"
        end Select
    end function

    public function getMifinishStateText()
        select Case FMifinishState
            CASE 0 : getMifinishStateText="��ó��"
            CASE 4 : getMifinishStateText="���ȳ�"
            CASE 6 : getMifinishStateText="CSó���Ϸ�"
            CASE ELSE : getMifinishStateText = FMifinishState
        end Select
    end function

    public function getMifinishText()
        select Case FMifinishReason
            CASE "00" : getMifinishText = "�Է´��"
            CASE "01" : getMifinishText = "������"
            CASE "04" : getMifinishText = "�����ǰ"

            CASE "02" : getMifinishText = "�ֹ�����"
            CASE "52" : getMifinishText = "�ֹ�����"
            CASE "03" : getMifinishText = "�������"
            CASE "53" : getMifinishText = "�������"
            CASE "05" : getMifinishText = "ǰ�����Ұ�"
            CASE "55" : getMifinishText = "ǰ�����Ұ�"

            CASE "11" : getMifinishText = "������"
            CASE "12" : getMifinishText = "��ü����"

			CASE "25" : getMifinishText = "�����Է� �ȳ�"
			CASE "26" : getMifinishText = "��ǰ�Ұ� �ȳ�"
            CASE "21" : getMifinishText = "�� ����"
            CASE "22" : getMifinishText = "�� ��ǰ����"
            CASE "23" : getMifinishText = "CS�ù�����"
			CASE "41" : getMifinishText = "�ù�� ��������"

            CASE "31" : getMifinishText = "��ǰ ȸ������"
            CASE "32" : getMifinishText = "���ɹ�ǰ �Ұ���ǰ"
            CASE "33" : getMifinishText = "������û(�� ���Է�)"
            CASE "34" : getMifinishText = "��Ÿ"

            CASE ELSE : getMifinishText = FMifinishReason

        end Select
    end function

    public function getDPlusDateStr()
        getDPlusDateStr = ""

        getDPlusDateStr = "D+" & DateDiff("d",Fregdate,now())
    end function

	public function IsAvailJumun()
		IsAvailJumun = (Fdeleteyn = "N")
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end class

Class CCSMasterItem
	public Fasid
	public FOrderSerial
	public Fdeleteyn
    public Fbuyname
    public Fbuyhp
    public Fbuyemail

    public Fdivcd

    public function getDivcdStr()
        select Case Fdivcd
            CASE "A000" : getDivcdStr="���"
            CASE "A100" : getDivcdStr="���"
            CASE "A004" : getDivcdStr="��ǰ"
            CASE ELSE : getDivcdStr = Fdivcd
        end Select
    end function

    public function getDivcdColor()
        select Case Fdivcd
            CASE "A000" : getDivcdColor="blue"
            CASE "A100" : getDivcdColor="blue"
            CASE "A004" : getDivcdColor="red"
            CASE ELSE : getDivcdColor = "black"
        end Select
    end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COldMiFinishItem
	public Fcsdetailidx
	public FOrderSerial
	public Fasid
	public Fdivcd
	public FMakerId
	public FItemId
	public FItemName
	public FItemOptionName
	public FRegItemNo
	public FIsUpcheBeasong
	public FCurrState
	public Fitemlackno
	public FCode
	public FState
	public FRegDate
	public FIpgoDate
	public FBuyName
	public FBuyPhone
	public FBuyHP
	public FReqName
	public FDeliveryNo
	public FSiteName
	public FUserId
	public FSubTotalPrice
	public Fipkumdiv
	public FrequestString
	public FfinishString
    public Fbuyemail
    public FItemcnt
    public FItemoption
    public Fbeasongdate
    public FSongjangno
    public FSongjangdiv
    public FMifinishReason
    public FMifinishState
	public FMifinishipgodate
    public FMifinishregdate
    public FisSendSMS
    public FisSendEmail
    public FisSendCall
    public Fcompany_name
    public Fcompany_tel
    public Fsmallimage
    public Fdeleteyn
    public Fdetailidx

    public function getDivcdStr()
        select Case Fdivcd
            CASE "A000" : getDivcdStr="���"
            CASE "A100" : getDivcdStr="���"
            CASE "A004" : getDivcdStr="��ǰ"
            CASE ELSE : getDivcdStr = Fdivcd
        end Select
    end function

    public function getDivcdColor()
        select Case Fdivcd
            CASE "A000" : getDivcdColor="blue"
            CASE "A100" : getDivcdColor="blue"
            CASE "A004" : getDivcdColor="red"
            CASE ELSE : getDivcdColor = "black"
        end Select
    end function

    public function getDPlusDateStr()
        getDPlusDateStr = ""

        getDPlusDateStr = "D+" & DateDiff("d",Fregdate,now())
    end function

    public function getDPlusDate()
        getDPlusDate = ""

        getDPlusDate = DateDiff("d",Fregdate,now())
    end function

    public function getBeasongDPlusDate()
        getBeasongDPlusDate = ""

        if IsNULL(Fbaljudate) then
            exit function
        end if

        if IsNULL(Fbeasongdate) then
            getBeasongDPlusDate = DateDiff("d",Fbaljudate,now())
            exit function
        end if

        getBeasongDPlusDate = DateDiff("d",Fbaljudate,Fbeasongdate)
    end function

    public function getMifinishDPlusDate
        dim BaseDate , tmp

		BaseDate = Left(CStr(now()),10)

        tmp = DateDiff("d",BaseDate,FMifinishipgodate)
        if (tmp>=0) then
            getMifinishDPlusDate = tmp
        else
            getMifinishDPlusDate = 0
        end if
    end function

    public function getSMSText()
        dim smstext
        smstext = ""

        if (FMifinishipgodate<>"") then
            if (FMifinishReason="05") then

            elseif (FMifinishReason="02") then  ''�ֹ�����(����)
                ''��� �ҿ��ϼ� D+2�̻�
                if (getMifinishDPlusDate>1) then
                    smstext = "[�ٹ����� CS��������ȳ�]��û�Ͻ� ��ǰ�� "&DdotFormat(FItemName,32)&"("&FItemID&")��ǰ�� "&VbCrlf
                    smstext = smstext&"�ֹ�����(����) ��ǰ���� "&FMifinishipgodate&"�� �߼۵� �����Դϴ�. ������ ��� �˼��մϴ�."
                else
                ''��� �ҿ��ϼ� D+0/D+1
                    smstext = "[�ٹ����� CS������ȳ�]��û�Ͻ� ��ǰ�� "&DdotFormat(FItemName,32)&"("&FItemID&")��ǰ�� "&VbCrlf
                    smstext = smstext&FMifinishipgodate&"�� �߼۵� �����Դϴ�. �����մϴ�."
                end if
            elseif (FMifinishReason="03") then  ''�������
                ''��� �ҿ��ϼ� D+2�̻�
                if (getMifinishDPlusDate>1) then
                    smstext = "[�ٹ����� CS��������ȳ�]��û�Ͻ� ��ǰ�� "&DdotFormat(FItemName,32)&"("&FItemID&")��ǰ�� "&VbCrlf
                    smstext = smstext&FMifinishipgodate&"�� �߼۵� �����Դϴ�. ���ο� ������ ��� �˼��մϴ�."
                else
                ''��� �ҿ��ϼ� D+0/D+1
                    smstext = "[�ٹ����� CS������ȳ�]��û�Ͻ� ��ǰ�� "&DdotFormat(FItemName,32)&"("&FItemID&")��ǰ�� "&VbCrlf
                    smstext = smstext&FMifinishipgodate&"�� �߼۵� �����Դϴ�. �����մϴ�."

                end if
            elseif (FMifinishReason="04") then  ''�����ǰ
                ''��� �ҿ��ϼ� D+2�̻�
                if (getMifinishDPlusDate>1) then
                    smstext = "[�ٹ����� CS������ȳ�]��û�Ͻ� ��ǰ�� "&DdotFormat(FItemName,32)&"("&FItemID&")��ǰ�� "&VbCrlf
                    smstext = smstext&"�����ۻ�ǰ���� "&FMifinishipgodate&"�� �߼۵� �����Դϴ�. �����մϴ�."
                else
                ''��� �ҿ��ϼ� D+0/D+1
                    smstext = "[�ٹ����� CS������ȳ�]��û�Ͻ� ��ǰ�� "&DdotFormat(FItemName,32)&"("&FItemID&")��ǰ�� "&VbCrlf
                    smstext = smstext&"�����ۻ�ǰ���� "&FMifinishipgodate&"�� �߼۵� �����Դϴ�. �����մϴ�."

                end if
            elseif (FMifinishReason="07") then  ''���������
                ''��� �ҿ��ϼ� D+2�̻�
                if (getMifinishDPlusDate>1) then
                    smstext = "[�ٹ����� CS������ȳ�]��û�Ͻ� ��ǰ�� "&DdotFormat(FItemName,32)&"("&FItemID&")��ǰ�� "&VbCrlf
                    smstext = smstext&"��������ۻ�ǰ���� "&FMifinishipgodate&"�� �߼۵� �����Դϴ�. �����մϴ�."
                else
                ''��� �ҿ��ϼ� D+0/D+1
                    smstext = "[�ٹ����� CS������ȳ�]��û�Ͻ� ��ǰ�� "&DdotFormat(FItemName,32)&"("&FItemID&")��ǰ�� "&VbCrlf
                    smstext = smstext&"��������ۻ�ǰ���� "&FMifinishipgodate&"�� �߼۵� �����Դϴ�. �����մϴ�."

                end if
            end if
        end if
        getSMSText = smstext
    end function

    public function isMifinishAlreadyInputed()
        isMifinishAlreadyInputed = Not (IsNULL(FMifinishReason) or (FMifinishReason="00") or (FMifinishReason=""))
    end function

    public function getDlvCompanyName()
        if FIsUpchebeasong="Y" then
            getDlvCompanyName = Fcompany_name
        else
            getDlvCompanyName = "�ٹ�����"
        end if
    end function

    Public function getUpcheDeliverStateName()
		 if IsNull(FCurrState) then
		    if (Fipkumdiv<4) then
		        getUpcheDeliverStateName = "�ֹ�����"
		    else
			    getUpcheDeliverStateName = "�����Ϸ�"
			end if
		 elseif FCurrState="2" then
			 getUpcheDeliverStateName = "�ֹ��뺸"
		 elseif FCurrState="3" then
			 getUpcheDeliverStateName = "�ֹ�Ȯ��"
		 elseif FCurrState="7" then
			 getUpcheDeliverStateName = "���Ϸ�"
		 else
			 getUpcheDeliverStateName = ""
		 end if
	 end Function

	public function getUpCheDeliverStateColor()
		if IsNull(FCurrState) then
		    if (Fipkumdiv<4) then
		        getUpCheDeliverStateColor ="#444444"
		    else
			    getUpCheDeliverStateColor ="#3300CC"
			end if

		elseif FCurrState="2" then
			getUpCheDeliverStateColor="#336600"
		elseif FCurrState="3" then
			getUpCheDeliverStateColor="#CC9933"
		elseif FCurrState="7" then
			getUpCheDeliverStateColor="#FF0000"
		else
			getUpCheDeliverStateColor="#000000"
		end if
	end function

	public function IpkumDivColor()
		if Fipkumdiv="0" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="1" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="2" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			IpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			IpkumDivColor="#444400"
		elseif Fipkumdiv="6" then
			IpkumDivColor="#FFFF00"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#004444"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#FF00FF"
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
			IpkumDivName="�ֹ�����"
		elseif Fipkumdiv="4" then
			IpkumDivName="�����Ϸ�"
		elseif Fipkumdiv="5" then
			IpkumDivName="�ֹ��뺸"
		elseif Fipkumdiv="6" then
			IpkumDivName="��ǰ�غ�"
		elseif Fipkumdiv="7" then
			IpkumDivName="�Ϻ����"
		elseif Fipkumdiv="8" then
			IpkumDivName="���Ϸ�"
		end if
	end function

	public function getIpgoMayDay()
		if IsNULL(FIpgoDate) then
			getIpgoMayDay = "&nbsp;"
		else
			getIpgoMayDay = CStr(FIpgoDate)
		end if
	end function

    public function getMiFinishCodeColor()
		if FMifinishReason="05" then
			getMiFinishCodeColor = "#FF0000"
		else
			getMiFinishCodeColor = "#000000"
		end if
	end function

	public function getMiFinishCodeName()
		if FCode="00" then
			getMiFinishCodeName = "�Է´��"
		elseif FCode="01" then
			getMiFinishCodeName = "������" ''������
		elseif FCode="02" then
			getMiFinishCodeName = "�ֹ�����(����)"
		elseif FCode="03" then
			getMiFinishCodeName = "�������"
		elseif FCode="04" then
			getMiFinishCodeName = "�����ǰ" ''"������" ''������
		elseif FCode="05" then
			getMiFinishCodeName = "ǰ�����Ұ�"
		elseif FCode="06" then
			getMiFinishCodeName = "�Ż�ǰ�԰�����" ''������
	    elseif FCode="07" then
			getMiFinishCodeName = "���������" ''2011-05�߰�
		elseif FCode="11" then
			getMiFinishCodeName = "������"
		elseif FCode="12" then
			getMiFinishCodeName = "��ü����"
		elseif FCode="21" then
			getMiFinishCodeName = "�� ��ȭ����"
		elseif FCode="22" then
			getMiFinishCodeName = "�� ��ǰ����"
		elseif FCode="23" then
			getMiFinishCodeName = "CS�ù�����"
		elseif FCode="31" then
			getMiFinishCodeName = "��ǰ ȸ������"
		elseif FCode="32" then
			getMiFinishCodeName = "���ɹ�ǰ �Ұ���ǰ"
		elseif FCode="33" then
			getMiFinishCodeName = "������û(�� ���Է�)"
		elseif FCode="34" then
			getMiFinishCodeName = "��Ÿ"
		else
			getMiFinishCodeName = "&nbsp;"
		end if
	end function

	public Function GetOptionName()
		if IsNULL(FItemOptionName) or (FItemOptionName="") then
			GetOptionName = "&nbsp;"
		else
			GetOptionName = FItemOptionName
		end if
	end Function

	public Function GetBeagonGubunColor()
		if FIsUpcheBeasong="Y" then
			GetBeagonGubunColor = "#000000"
		else
			GetBeagonGubunColor = "#33EE33"
		end if
	end function

	public Function GetBeagonGubunName()
		if FIsUpcheBeasong="Y" then
			GetBeagonGubunName = "��ü"
		else
			GetBeagonGubunName = "10x10"
		end if
	end function

	public Function GetBeagonStateColor()
		if (IsNULL(FCurrState) or (FCurrState=0)) and FIsUpcheBeasong="Y" then
			GetBeagonStateColor = "#EE3333"
		elseif FCurrState="3" then
			GetBeagonStateColor = "#3333EE"
		else
			GetBeagonStateColor = "#000000"
		end if
	end function

	public Function GetBeagonStateName()
		if (IsNULL(FCurrState) or (FCurrState=0)) and FIsUpcheBeasong="Y" then
			GetBeagonStateName = "��Ȯ��"
		elseif FCurrState="3" then
			GetBeagonStateName = "��üȮ��"
		else
			GetBeagonStateName = "&nbsp;"
		end if
	end function

    ''2009�� ���� ���� isSendSMS, isSendEmail, isSendCall
	public Function GetStateString()
		if FState = "0" then
			GetStateString = "��ó��"
		elseif FState="1" then
			GetStateString = "SMS�Ϸ�"
		elseif FState="2" then
			GetStateString = "�ȳ�Mail�Ϸ�"
		elseif FState="3" then
			GetStateString = "��ȭ�Ϸ�"
		''elseif FState="3" then
		''	GetStateString = "��۽�ó��"
		elseif FState="4" then
			GetStateString = "���ȳ�"         '' 2009�ű�
		elseif FState="6" then
			GetStateString = "CSó���Ϸ�"
		elseif FState="7" then
			GetStateString = "��۽� ó���Ϸ�"
		else
			GetStateString = "&nbsp;"
		end if
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CCSMifinishMaster
	public FItemList()
	public FOneItem
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectRegStart
	public FRectRegEnd
    public FRectDesignerID
    public FRectIpkumdiv
	public FRectItemid
	public FRectOrderserial
	public FRectBuyName
	public FRectReqName
	public FRectAsid
	public FRectCSDetailIDx
	public FRectDivCD
    public FRectMifinishReason
    public FRectMifinishState
    public FRectdplusOver
    public FRectdplusLower
    public FRectSiteName
    public FRectSortBy
    public FRectExInMayChulgoDay
    public FRectExOldCS
    public FRectExChangeMindReturn
	public FRectExRegbyCS
	public FRectorder6MonthBefore

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	' /cscenter/mifinish/cs_mifinishlist.asp
	public Sub getUpcheMifinishList
	    dim sqlStr, addSql, i, tmpSql, tmpDate, stOrderSerial, edOrderserial

		stOrderSerial = Mid(Replace(CStr(FRectRegStart),"-",""),3,6) + "00000"
		edOrderserial = Mid(Replace(CStr(FRectRegEnd),"-",""),3,6) + "00000"

		if FRectorder6MonthBefore="Y" then
			addSql = " from db_log.dbo.tbl_old_order_master_2003 o with (nolock)"
		else
			addSql = " from db_order.dbo.tbl_order_master o with (nolock)"
		end if
		addSql = addSql + " join db_cs.dbo.tbl_new_as_list m with (nolock)"
		addSql = addSql + " 	on o.orderserial = m.orderserial "
		addSql = addSql + " join db_cs.dbo.tbl_new_as_detail d with (nolock)"
		addSql = addSql + " 	on m.id = d.masterid "
	    addSql = addSql + " left join [db_temp].dbo.tbl_csmifinish_list T with (nolock)"
	    addSql = addSql + " 	on d.id=T.csdetailidx "
	    addSql = addSql + " left join db_cs.dbo.tbl_new_as_list r with (nolock)"
	    addSql = addSql + " 	on m.id = r.refasid "
	    addSql = addSql + " 	and m.divcd in ('A000', 'A100') "
		addSql = addSql + " where "
		addSql = addSql + " 	1 = 1 "
		addSql = addSql + " 	and m.deleteyn = 'N'"
		if application("Svr_Info") <> "Dev" then
			addSql = addSql + " 	and m.id >= 1200000 "		'// �ӵ�����
		else
			addSql = addSql + " 	and m.id >= 600000 "		'// �ӵ�����
		end if
		addSql = addSql + " 	and m.currstate < 'B006' "
		addSql = addSql + " 	and d.itemid <> 0 "
		addSql = addSql + " 	and d.isupchebeasong='Y'"
		addSql = addSql + " 	and m.divcd in ('A000', 'A100', 'A004') "
		addSql = addSql + " 	and ((m.divcd not in ('A000', 'A100')) or (r.currstate >= 'B006')) "

        if (FRectDivCD <>"") then
			if (FRectDivCD = "chulgocs") then
				addSql = addSql + " and m.divcd in ('A000', 'A100') "
			elseif (FRectDivCD = "returncs") then
				addSql = addSql + " and m.divcd = 'A004' "
			end if
		end if

        if (FRectDesignerID <>"") then
			addSql = addSql + " and d.makerid='" & FRectDesignerID & "'"
		end if

		if (FRectSiteName<>"") then
			if (FRectSiteName = "extall") then
				addSql = addSql + " and o.sitename <> '10x10' "
			else
				addSql = addSql + " and o.sitename = '" & FRectSiteName & "'"
			end if
		end if

		if (FRectItemid<>"") then
		    addSql = addSql + " and d.itemid="&FRectItemid&""
		end if

		if (FRectOrderserial<>"") then
		    addSql = addSql + " and o.orderserial='"&FRectOrderserial&"' "
		end if

		if (FRectBuyName<>"") then
		    addSql = addSql + " and o.buyname='"&FRectBuyName&"' "
		end if

		if (FRectReqName<>"") then
		    addSql = addSql + " and o.reqname='"&FRectReqName&"' "
		end if

		if (FRectRegStart<>"") then
		    addSql = addSql + " and m.regdate >= '" & FRectRegStart & "'"
		end if

		if (FRectRegEnd<>"") then
		    addSql = addSql + " and m.regdate < '" & FRectRegEnd & "'"
		end if

		if (FRectdplusOver<>"") then

			if (FRectdplusOver = "below3day") then
				tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 3 " & VbCRLF
				rsget.CursorLocation = adUseClient
				rsget.Open tmpSql, dbget, adOpenForwardOnly
        		if Not rsget.Eof then
					tmpDate = rsget("minusworkday")
				end if
        		rsget.close

				'// �ٹ��ϼ� ���� D+3 �̸� ��ü
				addSql = addSql + "     and datediff(d, m.regdate, '" & tmpDate & "') < 0 "
			else
				tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', " & FRectdplusOver & " " & VbCRLF
				rsget.CursorLocation = adUseClient
				rsget.Open tmpSql, dbget, adOpenForwardOnly
        		if Not rsget.Eof then
					tmpDate = rsget("minusworkday")
				end if
        		rsget.close

				'// �ٹ��ϼ� ���� D+n ��
				addSql = addSql + "     and datediff(d, m.regdate, '" & tmpDate & "') >= 0 "
			end if

		end if

        if (FRectdplusLower<>"") then
			if (FRectdplusLower = "7") then
				tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 7 " & VbCRLF
				rsget.CursorLocation = adUseClient
				rsget.Open tmpSql, dbget, adOpenForwardOnly
        		if Not rsget.Eof then
					tmpDate = rsget("minusworkday")
				end if
        		rsget.close

				'// �ٹ��ϼ� ���� D+7 �̸�
				addSql = addSql + "     and datediff(d, m.regdate, '" & tmpDate & "') < 0 "
			else
				addSql = addSql + "     and datediff(d,m.regdate, getdate())<=" & FRectdplusLower
			end if


        end if

        if (FRectMifinishReason<>"") then
            if (FRectMifinishReason="00") then
                addSql = addSql + "     and IsNULL(T.code,'00')='" & FRectMifinishReason & "'"
            else
                addSql = addSql + "     and T.code='" & FRectMifinishReason & "'"
            end if
        end if

        if (FRectMifinishState="N") then
            addSql = addSql + "     and T.state is NULL"
        elseif (FRectMifinishState<>"") then
            addSql = addSql + "     and T.state='" & FRectMifinishState & "'"
        end if

		'// ������� ���� �ֹ� ����(������� ���� �Ǵ� ǰ�����Ұ� ���� ǥ��)
        if (FRectExInMayChulgoDay="Y") then
            addSql = addSql + "     and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
        end if

        if (FRectExOldCS="Y") then
            addSql = addSql + "     and (datediff(m, m.regdate, getdate()) < 3) "
        end if

        if (FRectExChangeMindReturn="Y") then
			addSql = addSql + " 	and ( "
			addSql = addSql + " 		(m.divcd <> 'A004') "
			addSql = addSql + " 		or "
			addSql = addSql + " 		not (m.gubun01 = 'C004' and m.gubun02 = 'CD01') "
			addSql = addSql + " 	) "
        end if

		if (FRectExRegbyCS = "Y") then
			addSql = addSql + " and m.title like '[[]��%' "
		end if

		sqlStr = "select count(o.orderserial) as cnt, CEILING(CAST(Count(o.orderserial) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr + addSql

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		if FTotalCount < 1 then exit Sub
		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		sqlStr = "select top "&FPageSize*FCurrPage&" m.divcd, m.orderserial, m.id as asid, d.regitemno, d.itemid, d.itemname"
		sqlStr = sqlStr + " ,d.itemoptionname, isNull(d.currstate,0) as detailstate"
		sqlStr = sqlStr + " ,m.deleteyn, m.regdate, o.buyname, o.reqname , d.makerid, d.id as csdetailidx "
		sqlStr = sqlStr + " ,m.regdate, T.code, T.state, T.ipgodate, T.regdate as mifinishregdate "
		sqlStr = sqlStr + " , (case when IsNull(m.songjangdiv, 0) <> 0 and IsNull(m.songjangno, '') <> '' then 'Y' else 'N' end) as songjangyn "
		sqlStr = sqlStr + " , T.reguserid, T.lastupdate "
		sqlStr = sqlStr + addSql

		if (FRectSortBy = "makerid") then
			sqlStr = sqlStr + " order by d.makerid, d.itemid, d.itemoption"
		elseif (FRectSortBy = "orderserial") then
			sqlStr = sqlStr + " order by m.orderserial, d.itemid, d.itemoption"
		else
			sqlStr = sqlStr + " order by isNull(m.regdate,getdate()+365),  d.currstate, d.makerid, d.itemid, d.itemoption"
		end if

		'response.write sqlStr & "<br>"
		rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)
				set FItemList(i) = new CCSMifinishDetail

				FItemList(i).Fdivcd = rsget("divcd")
				FItemList(i).FOrderserial = rsget("orderserial")
				FItemList(i).Fasid = rsget("asid")
				FItemList(i).FItemid 	    = rsget("itemid")
				FItemList(i).FItemname    = db2html(rsget("itemname"))
				FItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
				FItemList(i).FItemcnt     = rsget("regitemno")					'// ��������
				FItemList(i).FBuyname     = db2html(rsget("buyname"))
				FItemList(i).FReqname     = db2html(rsget("reqname"))
				FItemList(i).Fdeleteyn	 = rsget("deleteyn")
				FItemList(i).FRegdate     = rsget("regdate")
				FItemList(i).FCurrstate   = rsget("detailstate")
				FItemList(i).FMakerid     = rsget("makerid")
                FItemList(i).FMifinishReason  = rsget("code")
                FItemList(i).FMifinishState   = rsget("state")
                FItemList(i).FMifinishipgodate= rsget("ipgodate")
                FItemList(i).Fmifinishregdate = rsget("mifinishregdate")
                FItemList(i).Fcsdetailidx = rsget("csdetailidx")
				FItemList(i).Fsongjangyn = rsget("songjangyn")
				FItemList(i).Freguserid = rsget("reguserid")
				FItemList(i).Flastupdate = rsget("lastupdate")

				rsget.movenext
				i=i+1

			loop
		end if
		rsget.Close
    end Sub

	' /cscenter/mifinish/cs_mifinishmaster_main.asp
	public sub GetOneCSMaster
		dim sqlStr,i

		if FRectAsid="" or isnull(FRectAsid) then exit sub

		sqlStr = " select top 1 m.divcd, m.id as asid, o.orderserial, m.deleteyn, o.buyname, o.buyhp, o.buyemail"

		if FRectorder6MonthBefore="Y" then
			sqlStr = sqlStr & " from db_log.dbo.tbl_old_order_master_2003 o with (nolock)"
		else
			sqlStr = sqlStr & " from db_order.dbo.tbl_order_master o with (nolock)"
		end if

		sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_list m with (nolock)"
		sqlStr = sqlStr + " 	on o.orderserial = m.orderserial "
		sqlStr = sqlStr + " where m.id= " + CStr(FRectAsid) + " "

		'response.write sqlStr & "<Br>"
		rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount

		set FOneItem = new CCSMasterItem
		if Not rsget.Eof then
			FOneItem.Fdivcd 		= rsget("divcd")
			FOneItem.Fasid 			= rsget("asid")
			FOneItem.FOrderSerial 	= rsget("orderserial")
			FOneItem.Fdeleteyn    	= rsget("deleteyn")
			FOneItem.Fbuyname    	= db2Html(rsget("buyname"))
			FOneItem.Fbuyhp    		= rsget("buyhp")
			FOneItem.Fbuyemail    	= db2Html(rsget("buyemail"))
		end if

		rsget.Close
	end sub

	' /cscenter/mifinish/cs_mifinishmaster_main.asp
	public function getMiFinishCSDetailList()
        dim sqlStr, i

		sqlStr = " select "
		sqlStr = sqlStr + " 	d.id as csdetailidx, d.regitemno, m.orderserial, d.itemid, d.itemoption, d.itemname,d.itemoptionname, d.makerid, d.isupchebeasong, "
		sqlStr = sqlStr + " 	m.deleteyn, m.regdate, o.buyname, o.reqname , o.buyemail, o.buyhp, d.songjangno, d.songjangdiv, "
		sqlStr = sqlStr + " 	T.code, T.state, T.ipgodate, "
		sqlStr = sqlStr + " 	IsNULL(T.isSendSMS,'N') as isSendSMS, "
		sqlStr = sqlStr + " 	IsNULL(T.isSendEmail,'N') as isSendEmail, "
		sqlStr = sqlStr + " 	IsNULL(T.isSendCall,'N') as isSendCall, "
		sqlStr = sqlStr + " 	T.reqstr, IsNULL(T.itemlackno,0) as itemlackno, T.finishstr, "
		sqlStr = sqlStr + " 	i.smallimage, p.company_name, p.tel as company_tel "

		if FRectorder6MonthBefore="Y" then
			sqlStr = sqlStr & " from db_log.dbo.tbl_old_order_master_2003 o with (nolock)"
		else
			sqlStr = sqlStr & " from db_order.dbo.tbl_order_master o with (nolock)"
		end if

		sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_list m with (nolock)"
		sqlStr = sqlStr + " 	on o.orderserial = m.orderserial "
		sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d with (nolock)"
		sqlStr = sqlStr + " 	on m.id = d.masterid "
	    sqlStr = sqlStr + " left join [db_temp].dbo.tbl_csmifinish_list T with (nolock)"
	    sqlStr = sqlStr + " 	on d.id=T.csdetailidx "
	    sqlStr = sqlStr + " Left Join [db_item].dbo.tbl_item i with (nolock) on d.itemid=i.itemid "
	    sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_partner p with (nolock) on d.makerid=p.id "
	    sqlStr = sqlStr + " where "
	    sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.deleteyn = 'N'"
		sqlStr = sqlStr + " 	and m.currstate < 'B006' "
		sqlStr = sqlStr + " 	and d.itemid <> 0 "
		sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	    sqlStr = sqlStr + " 	and m.id= " + CStr(FRectAsid) + " "

		'response.write sqlStr & "<Br>"
		rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
		'rsget.CursorType = adOpenStatic
		'rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		i=0
		redim FItemList(FResultCount)
		if not rsget.EOF then
			do until rsget.eof
				set FItemList(i) = new COldMiFinishItem

    			FItemList(i).Fcsdetailidx		  = rsget("csdetailidx")
    			FItemList(i).FOrderserial		  = rsget("orderserial")
    			FItemList(i).FItemid 			  = rsget("itemid")
    			FItemList(i).FItemoption     	  = rsget("itemoption")
    			FItemList(i).FItemname 		      = db2html(rsget("itemname"))
    			FItemList(i).FItemoptionName      = db2html(rsget("itemoptionname"))
    			FItemList(i).FRegItemNo           = rsget("regitemno")
    			FItemList(i).FMakerid 			  = rsget("makerid")
    			FItemList(i).FBuyname             = db2html(rsget("buyname"))
    			FItemList(i).FReqname			  = db2html(rsget("reqname"))
    			FItemList(i).Fdeleteyn		      = rsget("deleteyn")
    			FItemList(i).FRegdate			  = rsget("regdate")
    			FItemList(i).FisUpcheBeasong      = rsget("isUpcheBeasong")
    			FItemList(i).FSongjangno          = rsget("songjangno")
    			FItemList(i).FSongjangdiv         = rsget("songjangdiv")
                FItemList(i).FCode                = rsget("code")
                FItemList(i).FState               = rsget("state")
                FItemList(i).Fipgodate            = rsget("ipgodate")
                FItemList(i).FMifinishReason      = rsget("code")
                FItemList(i).FMifinishState       = rsget("state")
                FItemList(i).FMifinishipgodate    = rsget("ipgodate")
                FItemList(i).FisSendSMS           = rsget("isSendSMS")
                FItemList(i).FisSendEmail         = rsget("isSendEmail")
                FItemList(i).FisSendCall          = rsget("isSendCall")
                FItemList(i).Fbuyemail            = rsget("buyemail")
                FItemList(i).FbuyHp               = rsget("buyHp")
                FItemList(i).FrequestString       = db2Html(rsget("reqstr"))
                FItemList(i).Fitemlackno          = rsget("itemlackno")
                FItemList(i).FfinishString        = db2Html(rsget("finishstr"))
                FItemList(i).Fcompany_name        = db2Html(rsget("company_name"))
                FItemList(i).Fcompany_tel         = db2Html(rsget("company_tel"))
                FItemList(i).Fsmallimage          = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("smallimage")

                i=i+1
                rsget.MoveNext
            loop

        end if
        rsget.Close
    end function

    public function getOneMifinishItem()
        dim sqlStr

		sqlStr = " select"
		sqlStr = sqlStr & " m.id as asid, m.divcd, d.id as csdetailidx, d.regitemno, m.orderserial, d.itemid, d.itemoption"
		sqlStr = sqlStr & " , d.itemname,d.itemoptionname, d.makerid, d.isupchebeasong"
		sqlStr = sqlStr & " , m.deleteyn, m.regdate, o.buyname, o.reqname , o.buyemail, o.buyhp, d.songjangno, d.songjangdiv"
		sqlStr = sqlStr & " , T.code, T.state, T.ipgodate"
		sqlStr = sqlStr & " , IsNULL(T.isSendSMS,'N') as isSendSMS"
		sqlStr = sqlStr & " , IsNULL(T.isSendEmail,'N') as isSendEmail"
		sqlStr = sqlStr & " , IsNULL(T.isSendCall,'N') as isSendCall"
		sqlStr = sqlStr & " , T.reqstr, IsNULL(T.itemlackno,0) as itemlackno, T.finishstr"
		sqlStr = sqlStr & " , i.smallimage, p.company_name, p.tel as company_tel"

		if FRectorder6MonthBefore="Y" then
			sqlStr = sqlStr & " from db_log.dbo.tbl_old_order_master_2003 o with (nolock)"
		else
			sqlStr = sqlStr & " from db_order.dbo.tbl_order_master o with (nolock)"
		end if

		sqlStr = sqlStr & " join db_cs.dbo.tbl_new_as_list m with (nolock)"
		sqlStr = sqlStr & " 	on o.orderserial = m.orderserial"
		sqlStr = sqlStr & " join db_cs.dbo.tbl_new_as_detail d with (nolock)"
		sqlStr = sqlStr & " 	on m.id = d.masterid"
	    sqlStr = sqlStr & " left join [db_temp].dbo.tbl_csmifinish_list T with (nolock)"
	    sqlStr = sqlStr & " 	on d.id=T.csdetailidx"
	    sqlStr = sqlStr & " Left Join [db_item].dbo.tbl_item i with (nolock) on d.itemid=i.itemid"
	    sqlStr = sqlStr & " Left Join [db_partner].dbo.tbl_partner p with (nolock) on d.makerid=p.id"

	    sqlStr = sqlStr + " where "
	    sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.deleteyn = 'N'"
		sqlStr = sqlStr + " 	and m.currstate < 'B006' "
		sqlStr = sqlStr + " 	and d.itemid <> 0 "
		sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	    sqlStr = sqlStr + " 	and d.id= " + CStr(FRectCSDetailIDx) + " "

		'response.write sqlStr & "<Br>"
		rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
		'rsget.CursorType = adOpenStatic
		'rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		if not rsget.EOF then
            set FOneItem = new COldMiFinishItem

			FOneItem.Fdivcd		  		  = rsget("divcd")
			FOneItem.Fasid		  		  = rsget("asid")
			FOneItem.Fcsdetailidx		  = rsget("csdetailidx")
			FOneItem.FOrderserial		  = rsget("orderserial")
			FOneItem.FItemid 			  = rsget("itemid")
			FOneItem.FItemoption     	  = rsget("itemoption")
			FOneItem.FItemname 		      = db2html(rsget("itemname"))
			FOneItem.FItemoptionName      = db2html(rsget("itemoptionname"))
			FOneItem.FRegItemNo           = rsget("regitemno")
			FOneItem.FMakerid 			  = rsget("makerid")
			FOneItem.FBuyname             = db2html(rsget("buyname"))
			FOneItem.FReqname			  = db2html(rsget("reqname"))
			FOneItem.Fdeleteyn		      = rsget("deleteyn")
			FOneItem.FRegdate			  = rsget("regdate")
			FOneItem.FisUpcheBeasong      = rsget("isUpcheBeasong")
			FOneItem.FSongjangno          = rsget("songjangno")
			FOneItem.FSongjangdiv         = rsget("songjangdiv")
            FOneItem.FCode                = rsget("code")
            FOneItem.FState               = rsget("state")
            FOneItem.Fipgodate            = rsget("ipgodate")
            FOneItem.FMifinishReason      = rsget("code")
            FOneItem.FMifinishState       = rsget("state")
            FOneItem.FMifinishipgodate    = rsget("ipgodate")
            FOneItem.FisSendSMS           = rsget("isSendSMS")
            FOneItem.FisSendEmail         = rsget("isSendEmail")
            FOneItem.FisSendCall          = rsget("isSendCall")
            FOneItem.Fbuyemail            = rsget("buyemail")
            FOneItem.FbuyHp               = rsget("buyHp")
            FOneItem.FrequestString       = db2Html(rsget("reqstr"))
            FOneItem.Fitemlackno          = rsget("itemlackno")
            FOneItem.FfinishString        = db2Html(rsget("finishstr"))
            FOneItem.Fcompany_name        = db2Html(rsget("company_name"))
            FOneItem.Fcompany_tel         = db2Html(rsget("company_tel"))
            FOneItem.Fsmallimage          = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemid) + "/" + rsget("smallimage")

        end if
        rsget.Close
    end function

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
