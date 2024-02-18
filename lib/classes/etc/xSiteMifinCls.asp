<%
Class cXsiteMifinItem
    public FSellSite
    public FOutMallOrderSerial
    public FOrgDetailKey
    public FfinFlag
    public FconfirmDt
    public FshppNo
    public FshppSeq
    public Foutorderstatus
    public FreOrderYn
    public FdelayNts
    public FsplVenItemId
    public FoutmallGoodsNo
    public FoutmalloptionNo
    public FordQty
    public FshppDivDtlNm
    public FuitemNm
    public FshppRsvtDt
    public FwhoutCritnDt
    public FautoShortgYn
    public FMatchorderserial
    public FmatchItemID
    public Fmatchitemoption
    public Fregdt
    public Flastupdt
    public Fcancelyn
    public Fdcancelyn
    public Fbeasongdate
    public Fsongjangdiv
    public Fsongjangno
    public Fdlvfinishdt
    public FjungsanFixDate
    public Fdivname
    public Fmakerid
    public Fitemno
    public Fisupchebeasong
    public Fitemname
    public Fitemoptionname
    public FCurrstate

    public FshppTypeDtlNm
    public FdelicoVenId
    public FdelicoVenNm
    public FwblNo

    public FOrgOutMallOrderSerial
    public Fasid


    public function isTenOutDiffSongjang()
        isTenOutDiffSongjang = False
        if isNULL(FwblNo) or isNULL(Fsongjangno) then Exit function

        if (FSellSite="ssg") and (Fdivname="��Ÿ") and (LEFT(FdelicoVenNm,LEN(Null2Blank(Fdivname)))=Null2Blank(Fdivname)) then Exit function  ''ssg ��Ÿ�ΰ��

        isTenOutDiffSongjang = TRIM(replace(Null2Blank(FwblNo),"-",""))<>TRIM(replace(null2blank(Fsongjangno),"-",""))
    end function

    public function isTenOutDiffDlvNm()
        isTenOutDiffDlvNm = False
        if isNULL(FdelicoVenNm) or isNULL(Fdivname) then Exit function

        isTenOutDiffDlvNm = TRIM(LEFT(null2blank(FdelicoVenNm),2))<>TRIM(LEFT(replace(null2blank(Fdivname),"(��)����","CJ����"),2))
    end function

    public function getOutDlvInputedStr
        dim ret
        if NOT isNULL(FshppTypeDtlNm) then ret = FshppTypeDtlNm
        if ret="��ü�ù���" or ret="�ù���" then ret=""

        if NOT isNULL(FdelicoVenNm) then
            if ret<>"" then
                ret = ret&"<br>"&CHKIIF(isTenOutDiffDlvNm,"<strong>"&FdelicoVenNm&"</strong>",FdelicoVenNm)
            else
                ret = CHKIIF(isTenOutDiffDlvNm,"<strong>"&FdelicoVenNm&"</strong>",FdelicoVenNm)
            end if
        end if

        if NOT isNULL(FwblNo) then
            if ret<>"" then
                ret = ret&"<br>"&CHKIIF(isTenOutDiffSongjang,"<strong>"&FwblNo&"</strong>",FwblNo)
            else
                ret = CHKIIF(isTenOutDiffSongjang,"<strong>"&FwblNo&"</strong>",FwblNo)
            end if
        end if

        getOutDlvInputedStr = ret
    end function

    public function getOutorderStatusNm()
        if isNULL(Foutorderstatus) then Exit function

        getOutorderStatusNm = Foutorderstatus
        if Foutorderstatus="�ֹ��뺸" then
            getOutorderStatusNm="<strong>"&Foutorderstatus&"</strong>"
        end if
    end function

    '' ��ۿϷ� ����
    public function isStatusSendDliverFinish()
        isStatusSendDliverFinish = False
        if FshppDivDtlNm<>"�Ϲ����" and FshppDivDtlNm<>"��ȯ���" and FshppDivDtlNm<>"�ֹ����" then exit function

        if Foutorderstatus="���Ϸ�" and ((FCurrstate="7" and NOT isNULL(Fdlvfinishdt)) or (FCurrstate="B007" and NOT isNULL(Fasid))) then
            isStatusSendDliverFinish = true
            exit function
        end if

        ''if FSellSite="hmall1010" and FshppDivDtlNm="��ȯ���" and (FCurrstate="B007" and NOT isNULL(Fasid)) then
        ''    isStatusSendDliverFinish = true
        ''    exit function
        ''end if
    end function

    public function isStatusSendCancelFinish()
        isStatusSendCancelFinish = False
        if FshppDivDtlNm<>"�ֹ����" then exit function

        if FshppDivDtlNm="�ֹ����" and Foutorderstatus="�ֹ�Ȯ��" and (FCurrstate="B007" and NOT isNULL(Fasid)) then
            isStatusSendCancelFinish = true
            exit function
        end if
    end function

    '' �����Է� ����
    public function isStatusSendReqSongjang()
        isStatusSendReqSongjang = False
        if FshppDivDtlNm<>"�Ϲ����" and FshppDivDtlNm<>"��ȯ���" then exit function

        ''�� �б��� �ʿ�
        ' if (FSellSite="coupang") then  ''�������� �ȵǾ���, ��ۿϷ�����̴�.
        '     if Foutorderstatus="�ֹ�Ȯ��" and FCurrstate="7" and NOT isNULL(Fdlvfinishdt) and Null2Void(FwblNo)="" then
        '         isStatusSendReqSongjang = true
        '         exit function
        '     end if
        ' else
            if Foutorderstatus="�ֹ�Ȯ��" and (FCurrstate="7" or FCurrstate="B007") then
                isStatusSendReqSongjang = true
                exit function
            end if
        'end if
    end function

    '' ���Ϸ����� (�Ϻ� ����Ʈ :ssg, hmall)
    public function isStatusSendReqChulgo()
        isStatusSendReqChulgo = False
        if (FSellSite<>"ssg") and (FSellSite<>"hmall1010") then Exit function

        if FshppDivDtlNm<>"�Ϲ����" and FshppDivDtlNm<>"��ȯ���" then exit function

        if Foutorderstatus="��ŷ�Ϸ�" and FCurrstate="7" then
            isStatusSendReqChulgo = true
            exit function
        end if

        if (Foutorderstatus="��ŷ�Ϸ�") and FCurrstate="B007" then
            isStatusSendReqChulgo = true
            exit function
        end if

        if (Foutorderstatus="�ֹ�Ȯ��") and FCurrstate="B007" then
            isStatusSendReqChulgo = true
            exit function
        end if
    end function

    '// �ֹ�Ȯ��
    public function isStatusSendReqOrderConfirm()
        isStatusSendReqOrderConfirm = False
        if (FSellSite<>"ssg") and (FSellSite<>"hmall1010") then Exit function

        if (FshppDivDtlNm<>"") and Not IsNull(FshppDivDtlNm) and FshppDivDtlNm<>"�Ϲ����" and FshppDivDtlNm<>"��ȯ���" and FshppDivDtlNm<>"�ֹ����" and FshppDivDtlNm<>"��ȯ���öȸ" then exit function

        if (Foutorderstatus="�ֹ��뺸") and (FshppDivDtlNm="��ȯ���öȸ") then
            isStatusSendReqOrderConfirm = true
            exit function
        end if

        if (Foutorderstatus="�ֹ��뺸") and (FCurrstate="7") then
            isStatusSendReqOrderConfirm = true
            exit function
        end if

        if (Foutorderstatus="�ֹ��뺸") and (FCurrstate="B001" or FCurrstate="B007") then
            isStatusSendReqOrderConfirm = true
            exit function
        end if
    end function

    public function getTenStatusNm
        Dim ret : ret = ""

        if ((FshppDivDtlNm = "��ȯ���") or (FshppDivDtlNm = "��ȯ���öȸ") or (FshppDivDtlNm = "�ֹ����")) and (Not IsNull(Fasid)) then
            if FCurrstate="B007" then
                getTenStatusNm = "�Ϸ�"
            else
                getTenStatusNm = "����"
            end if
            exit function
        end if

        if FshppDivDtlNm<>"�Ϲ����" then
            getTenStatusNm = "?"
            exit function
        end if

        if Fcancelyn<>"N" then
            ret = ret&"<strong>�ֹ�"&Fcancelyn&"</strong>"
        end if
        if Fdcancelyn<>"N" then
            ret = ret&"<strong>��ǰ"&Fdcancelyn&"</strong>"
        end if

        if FCurrstate="7" then
            if isNULL(Fdlvfinishdt) then
            ret = ret&"���Ϸ�"
            else
            ret = ret&"��ۿϷ�"
            end if
        elseif FCurrstate="2" then
            ret = ret&"��ü�뺸"
        elseif FCurrstate="3" then
            ret = ret&"��ǰ�غ�"
        elseif isNULL(FCurrstate) then

        else
            ret = ret&"("&FCurrstate&")"
        end if
        getTenStatusNm = ret
    end function

    public function getShppDivDtlNm
        if isNULL(FshppDivDtlNm) then Exit function

        getShppDivDtlNm = FshppDivDtlNm
        if FshppDivDtlNm<>"�Ϲ����" then getShppDivDtlNm="<strong>"&FshppDivDtlNm&"</strong>"
    end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub


End Class

Class CxSiteMifinCls

	public FItemList()
	public FOneItem


	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount

    public FrectSellsite
    public FRectSearchtype
    public FRectMatchorderserial
    public FRectOutMallOrderSerial
    public FRectExcNoOrderSerial
    public FRectshppDivDtl

    public FLastUpDt

    public function getLastUpDt()
        if isNULL(FLastUpDt) then Exit function
        getLastUpDt = LEFT(FLastUpDt,19)
    end function

    public function getXSiteMifinLIST()
        sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSite_MichulList_CNT]  '"&FrectSellsite&"',"&FRectSearchtype&", '" & FRectMatchorderserial & "', '" & FRectOutMallOrderSerial & "', '" & FRectExcNoOrderSerial & "','"&FRectshppDivDtl&"'"

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
        if NOT rsget.Eof then
            FTotalCount = rsget("cnt")
            FLastUpDt = rsget("mxupdt")
        end if
        rsget.close()


        sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSite_MichulList_LIST] "&FCurrPage&","&FPageSize&",'"&FrectSellsite&"',"&FRectSearchtype&", '" & FRectMatchorderserial & "', '" & FRectOutMallOrderSerial & "', '" & FRectExcNoOrderSerial & "','"&FRectshppDivDtl&"'"

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
            do until rsget.eof
                set FItemList(i) = new cXsiteMifinItem


                FItemList(i).FSellSite              = rsget("SellSite")
                FItemList(i).FOutMallOrderSerial    = rsget("OutMallOrderSerial")
                FItemList(i).FOrgDetailKey          = rsget("OrgDetailKey")
                FItemList(i).FfinFlag               = rsget("finFlag")
                FItemList(i).FconfirmDt             = rsget("confirmDt")

                FItemList(i).FshppNo                = rsget("shppNo")
                FItemList(i).FshppSeq               = rsget("shppSeq")

                FItemList(i).Foutorderstatus       = rsget("outorderstatus")

                FItemList(i).FreOrderYn             = rsget("reOrderYn")
                FItemList(i).FdelayNts              = rsget("delayNts")
                FItemList(i).FsplVenItemId          = rsget("splVenItemId")

                FItemList(i).FoutmallGoodsNo        = rsget("outmallGoodsNo")
                FItemList(i).FoutmalloptionNo       = rsget("outmalloptionNo")
                FItemList(i).FordQty                = rsget("ordQty")
                FItemList(i).FshppDivDtlNm          = rsget("shppDivDtlNm")
                FItemList(i).FuitemNm               = rsget("uitemNm")
                FItemList(i).FshppRsvtDt            = rsget("shppRsvtDt")
                FItemList(i).FwhoutCritnDt          = rsget("whoutCritnDt")
                FItemList(i).FautoShortgYn          = rsget("autoShortgYn")
                FItemList(i).FMatchorderserial      = rsget("Matchorderserial")
                FItemList(i).FmatchItemID           = rsget("matchItemID")
                FItemList(i).Fmatchitemoption       = rsget("matchitemoption")

                FItemList(i).Fregdt         = rsget("regdt")
                FItemList(i).Flastupdt      = rsget("lastupdt")
                FItemList(i).Fcancelyn      = rsget("cancelyn")
                FItemList(i).Fdcancelyn     = rsget("dcancelyn")
                FItemList(i).Fbeasongdate   = rsget("beasongdate")
                FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
                FItemList(i).Fsongjangno    = rsget("songjangno")
                FItemList(i).Fdlvfinishdt   = rsget("dlvfinishdt")
                FItemList(i).FjungsanFixDate= rsget("jungsanfixdate")

                FItemList(i).Fdivname       = rsget("divname")
                FItemList(i).Fmakerid       = rsget("makerid")
                FItemList(i).Fitemno        = rsget("itemno")
                FItemList(i).Fisupchebeasong= rsget("isupchebeasong")
                FItemList(i).Fitemname      = rsget("itemname")
                FItemList(i).FitemOptionname= rsget("itemOptionname")
                FItemList(i).FCurrstate     = rsget("Currstate")


                FItemList(i).FshppTypeDtlNm = rsget("shppTypeDtlNm")
                FItemList(i).FdelicoVenId   = rsget("delicoVenId")
                FItemList(i).FdelicoVenNm   = rsget("delicoVenNm")
                FItemList(i).FwblNo         = rsget("wblNo")

                FItemList(i).FOrgOutMallOrderSerial    = rsget("OrgOutMallOrderSerial")
                FItemList(i).Fasid    		= rsget("asid")

                if ((FItemList(i).FshppDivDtlNm = "��ȯ���") or (FItemList(i).FshppDivDtlNm = "�ֹ����") or (FItemList(i).FshppDivDtlNm = "��ȯ���öȸ")) then
                    '// CS
                    if (Not IsNull(FItemList(i).Fasid)) then
                        FItemList(i).FCurrstate     = rsget("CsCurrstate")
                        FItemList(i).Fsongjangdiv   = rsget("cssongjangdiv")
                        FItemList(i).Fsongjangno    = rsget("cssongjangno")
                        FItemList(i).Fbeasongdate   = Left(rsget("csbeasongdate"), 10)		'// CS�Ϸ���
                        FItemList(i).Fdlvfinishdt   = ""
                        FItemList(i).Fitemno        = rsget("csitemno")
                    else
                        FItemList(i).FCurrstate     = ""
                        FItemList(i).Fsongjangdiv   = ""
                        FItemList(i).Fsongjangno    = ""
                        FItemList(i).Fbeasongdate   = ""
                        FItemList(i).Fdlvfinishdt   = ""
                        FItemList(i).Fitemno        = 0
                    end if
                end if

				rsget.moveNext
				i=i+1
			loop
        end if
        rsget.close()
    end function


	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
%>
