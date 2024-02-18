<%
'###########################################################
' Description : ������ ȯ�� Ŭ����
' History : �̻� ����
'           2021.08.31 �ѿ�� ����
'###########################################################

function DispAcctStar(orgAcct,starno,minlen)
    if IsNULL(orgAcct) then Exit function

    Dim ret, starStr, i
    DispAcctStar = orgAcct
    starStr = ""
    ret = ""

    if (Len(orgAcct)<minlen) then Exit function

    if (Len(orgAcct)-starno)>=0 then
        ret = Left(orgAcct,Len(orgAcct)-starno)
    else
        Exit function
    end if


    for i=0 to starno-1
        starStr = starStr + "*"
    next

    DispAcctStar = ret + starStr
end function

''ȯ�� ���� ����
Class CCSASRefundInfoItem
    ''' TBL_AS_REFUND_INFO
    public Fsitegubun			''����Ʈ ����(�ٹ�����/��ī����)

    public Fasid
    public Forgsubtotalprice    ''�� �ֹ� ������
    public Forgitemcostsum      ''�� �ֹ� ��ǰ�հ�
    public Forgbeasongpay       ''�� �ֹ� ��۷�
    public Forgmileagesum       ''�� �ֹ� ��븶�ϸ���
    public Forgcouponsum        ''�� �ֹ� �������
    public Forgallatdiscountsum ''�� �ֹ� �ÿ�����

    public Frefundrequire       ''ȯ�ҿ�û��
    public Frefundresult        ''ȯ��  �ݾ�
    public Freturnmethod        ''ȯ��  ���

    public Frefundmileagesum    ''���  ���ϸ��� Frefundmileagesum
    public Frefundcouponsum     ''���  ����     Frefundcouponsum
    public Fallatsubtractsum    ''���  ī������ Fallatsubtractsum

    public Frefunditemcostsum   ''��� ��ǰ�հ�
    public Frefundbeasongpay    ''��ҽ� ��ۺ� ������
    public Frefunddeliverypay   ''��ҽ� ȸ�� ��ۺ�? -> Freturndeliverypay
    public Frefundadjustpay     ''��ҽ� ��Ÿ ������
    public Fcanceltotal         ''�� ��Ҿ�

    public Frebankname          ''ȯ�� ����
    public Frebankaccount       ''ȯ�� ����
    public Frebankownername     ''���� ��
    public FpaygateTid          ''Pg�� T id
    public Fencmethod           ''��ȣȭ���
    public Fencaccount          ''��ȣȭ���¹�ȣ
    public Fdecaccount          ''��ȣȭ���¹�ȣ

    public FpaygateresultTid
    public FpaygateresultMsg
    public Fupfiledate          ''���ε� ��¥

    public FreturnmethodName    ''ȯ�ҹ�ĸ�


    ''' TBL_NEW_AS_LIST
    public FOrderSerial         ''�����ֹ���ȣ
    public Fuserid              ''�ֹ���ID
    public Fcustomername        ''�ֹ���ID
    public Fregdate
    public Fcurrstate

    public rebankCode

    public Fconfirmregmsg
    public Fconfirmfinishmsg
    public Fconfirmfinishdate

    ''tbl_IBK_ERP_ICHE_DATA
    public FIBK_TIDX
    public FIBK_PROC_YN
    public FIBK_PROC_DATE
    public FIBK_ERR_MSG
    public FIBK_TEN_STATUS
    public FIBK_EB_USED          ''e-branch ���翩��

    public function IsIBKRefund()
        IsIBKRefund = Not IsNULL(FIBK_TIDX)
    end function

    public function IsIBKProcERR()
        IsIBKProcERR = Not IsNULL(FIBK_PROC_YN) and (FIBK_PROC_YN<>"Y")
    end function

    ''���� �ۼ� ��� ��������..
    public function IsRollBackValid()
        IsRollBackValid = false
        if (IsNULL(FIBK_TIDX)) then
            IsRollBackValid = true
            Exit function
        end if

        if (IsNULL(FIBK_EB_USED)) then
            IsRollBackValid = true
            Exit function
        end if

        if (FIBK_ERR_MSG = "�ڷᰡ������ ������") then
            IsRollBackValid = true
            Exit function
        end if
    end function

    public function getIBKstateName()
        getIBKstateName = ""

        if IsNULL(FIBK_TIDX) then Exit function

        if (FIBK_EB_USED="Y") and IsNULL(FIBK_PROC_YN) then
            getIBKstateName="��û��"
            Exit function
        end if

        if IsNULL(FIBK_PROC_YN) then
            getIBKstateName="����"
            Exit function
        end if

        Select Case FIBK_PROC_YN

            CASE "Y" : getIBKstateName="IBK��ü�Ϸ�"
            CASE "F" : getIBKstateName="������û����"
            CASE "D" : getIBKstateName="�ڷᰡ������ ����"
            CASE "C" : getIBKstateName="�����Կ��� ����"
            CASE "R" : getIBKstateName="�ݷ�"
            CASE "N" : getIBKstateName="��Ÿ�����߻�"
            CASE ELSE : getIBKstateName=FIBK_PROC_YN
        end Select

        ''20090616�߰�
        if (FIBK_PROC_YN="Y") and (FIBK_PROC_DATE="") then
            getIBKstateName = "Ȯ�ο��PROC_DATE"
        end if
    end function

    public function IsConfirmMsgExists()
        IsConfirmMsgExists = Not IsNULL(Fconfirmregmsg)
    end function

    public function IsConfirmMsgFinished()
        IsConfirmMsgFinished = Not IsNULL(Fconfirmfinishdate)
    end function

    public function GetCurrStateColor()
        if (Fcurrstate="B001") then
            GetCurrStateColor = "#CC33CC"
        elseif (Fcurrstate="B005") then
            GetCurrStateColor = "#CCCC33"
        elseif (Fcurrstate="B007") then
            GetCurrStateColor = "#000000"
        else
            GetCurrStateColor = "#000000"
        end if
    end function

    public function GetCurrStateName()
        if (Fcurrstate="B001") then
            GetCurrStateName = "����"
        elseif (Fcurrstate="B005") then
            GetCurrStateName = "Ȯ�ο�û"
        elseif (Fcurrstate="B007") then
            GetCurrStateName = "�Ϸ�"
        else
            GetCurrStateName = Fcurrstate
        end if
    end function

    public function getUpLoadStateName()
        if IsNULL(Fupfiledate) then
            getUpLoadStateName = ""
        else
            getUpLoadStateName = "�ۼ���"
        end if
    end function

    public function IsInValidRefundInfo()
        IsInValidRefundInfo = (Len(Frebankname)<2) or (Len(Frebankaccount)<8) or (Len(Frebankownername)<2)
    end function

    public function getUploadrebankownername()
        getUploadrebankownername = Frebankownername '''& " ȯ��"
    end function

    public function getUploadrebankaccount()
        getUploadrebankaccount = replace(replace(Frebankaccount,"-","")," ","")
    end function

    public function getUploadbankname()
        if (Frebankname="��Ƽ") then
            getUploadbankname = "��Ƽ"
        elseif (Frebankname="��������") then
            getUploadbankname = "����(��������)"
        elseif (Frebankname="����") or (Frebankname="�����߾�ȸ") then
            getUploadbankname = "����"
            'if (Len(Frebankaccount)=12) then
            '    getUploadbankname = "����(��������)"
            'else
            '    getUploadbankname = "�����߾�ȸ"
            'end if
        else
            getUploadbankname = Frebankname
        end if
    end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
End Class


Class CCSASRefundInfoGroupItem
    public Fupfiledate
    public FCount

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
End Class

Class CCSRefund
    public FItemList()
    public FOneItem

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectCurrstate
    public FRectReturnmethod
    public FRectSearchType
    public FRectSearchString
    public FRectUploadState
    public FRectUpfiledate

    public FRectNotInputOnly		'// �������� ���Է�

    public Sub GetRefundRequireByFileDate
        dim sqlStr, i
        sqlStr = " select convert(varchar(19),r.upfiledate,21) as cvupfiledate, count(a.id) as cnt "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list a" + VbCrlf
        sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_as_refund_info r on a.id=r.asid" + VbCrlf
        sqlStr = sqlStr + " where a.divcd='A003'" + VbCrlf
        sqlStr = sqlStr + " and a.deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and a.currstate='" + FRectCurrstate + "'" + VbCrlf
        sqlStr = sqlStr + " and r.returnmethod='" + FRectReturnmethod + "'" + VbCrlf

        if (FRectSearchString<>"") then
            sqlStr = sqlStr + " and a." + FRectSearchType + "='" + FRectSearchString + "'"
        end if

        sqlStr = sqlStr + " and r.upfiledate is Not NULL"
        sqlStr = sqlStr + " group by convert(varchar(19),r.upfiledate,21)"
        sqlStr = sqlStr + " order by cvupfiledate asc"

        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly

        FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

        If Not rsget.Eof then
            do until rsget.eof
				set FItemList(i) = new CCSASRefundInfoGroupItem
                FItemList(i).Fupfiledate    = rsget("cvupfiledate")
                FItemList(i).FCount         = rsget("cnt")

				rsget.moveNext
				i=i+1
			loop

        end IF
    end Sub

    public Sub GetRefundRequireByFileDateAcademy
        dim sqlStr, i
        sqlStr = " select convert(varchar(19),r.upfiledate,21) as cvupfiledate, count(a.id) as cnt "
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list a" + VbCrlf
        sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_as_refund_info r on a.id=r.asid" + VbCrlf
        sqlStr = sqlStr + " where a.divcd='A003'" + VbCrlf
        sqlStr = sqlStr + " and a.deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and a.currstate='" + FRectCurrstate + "'" + VbCrlf
        sqlStr = sqlStr + " and r.returnmethod='" + FRectReturnmethod + "'" + VbCrlf

        if (FRectSearchString<>"") then
            sqlStr = sqlStr + " and a." + FRectSearchType + "='" + FRectSearchString + "'"
        end if

        sqlStr = sqlStr + " and r.upfiledate is Not NULL"
        sqlStr = sqlStr + " group by convert(varchar(19),r.upfiledate,21)"
        sqlStr = sqlStr + " order by cvupfiledate asc"

        rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly

        FResultCount = rsACADEMYget.RecordCount
        redim preserve FItemList(FResultCount)

        If Not rsACADEMYget.Eof then
            do until rsACADEMYget.eof
				set FItemList(i) = new CCSASRefundInfoGroupItem
                FItemList(i).Fupfiledate    = rsACADEMYget("cvupfiledate")
                FItemList(i).FCount         = rsACADEMYget("cnt")

				rsACADEMYget.moveNext
				i=i+1
			loop
        end IF
        rsACADEMYget.Close

    end Sub

    public Sub GetRefundRequireList
        dim sqlStr,i
        sqlStr = " select count(a.id) as cnt from [db_cs].[dbo].tbl_new_as_list a with (nolock)" + VbCrlf
        sqlStr = sqlStr + " join [db_cs].[dbo].tbl_as_refund_info r with (nolock) on a.id=r.asid" + VbCrlf
        sqlStr = sqlStr + " where a.divcd='A003'" + VbCrlf
        sqlStr = sqlStr + " and a.deleteyn='N'" + VbCrlf
        if (FRectCurrstate<>"") then
            sqlStr = sqlStr + " and a.currstate='" + FRectCurrstate + "'" + VbCrlf
        end if

        sqlStr = sqlStr + " and r.returnmethod='" + FRectReturnmethod + "'" + VbCrlf

        if (FRectSearchString<>"") then
            sqlStr = sqlStr + " and a." + FRectSearchType + "='" + FRectSearchString + "'"
        end if

        if (FRectNotInputOnly = "Y") then
            sqlStr = sqlStr + " and ((IsNull(r.rebankname, '') = '') or (IsNull(r.rebankownername, '') = '')) "
        elseif (FRectNotInputOnly = "N") then
            sqlStr = sqlStr + " and ((IsNull(r.rebankname, '') <> '') and (IsNull(r.rebankownername, '') <> '') and (r.encaccount is not NULL)) "
        end if

        if FRectUploadState="notupload" then
            sqlStr = sqlStr + " and r.upfiledate is NULL"
        elseif FRectUploadState="uploaded" then
            sqlStr = sqlStr + " and r.upfiledate is Not NULL"
        end if

        if FRectUpfiledate<>"" then
            sqlStr = sqlStr + " and r.upfiledate='" + FRectUpfiledate + "'"
        end if

		'response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly
            FTotalCount = rsget("cnt")
        rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " a.OrderSerial" + VbCrlf
        sqlStr = sqlStr + " ,a.userid, a.customername, a.regdate, a.currstate" + VbCrlf
        sqlStr = sqlStr + " , r.*, f.confirmregmsg, f.confirmfinishmsg, f.confirmfinishdate," + VbCrlf
        sqlStr = sqlStr + " convert(varchar(19),r.upfiledate,21) as cvupfiledate, c.comm_name as returnmethodName " + VbCrlf
        sqlStr = sqlStr + " ,K.TIDX,K.EB_USED,K.PROC_YN,K.PROC_DATE,K.ERR_MSG,K.TEN_STATUS"
        sqlStr = sqlStr + " , IsNull(r.encmethod, '') as encmethod "
        sqlStr = sqlStr + " , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(r.encaccount), '') WHEN r.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(r.encaccount), '') ELSE '' END) as decaccount "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list a with (nolock)" + VbCrlf
        sqlStr = sqlStr + " join [db_cs].[dbo].tbl_as_refund_info r with (nolock) on a.id=r.asid" + VbCrlf
        sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_cs_comm_code c with (nolock) on r.returnmethod=c.comm_cd" + VbCrlf
        sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_confirm f with (nolock) on a.id=f.asid " + VbCrlf
        sqlStr = sqlStr + " left join db_log.dbo.tbl_IBK_ERP_ICHE_DATA K with (nolock)"
        sqlStr = sqlStr + "     on r.IBK_TIDX=K.TIDX and IsNull(K.SITEGUBUN, '10x10') = '10x10'" ''and R.IBK_TIDX is Not NULL
        sqlStr = sqlStr + " where a.divcd='A003'" + VbCrlf
        sqlStr = sqlStr + " and a.deleteyn='N'" + VbCrlf
        if (FRectCurrstate<>"") then
            sqlStr = sqlStr + " and a.currstate='" + FRectCurrstate + "'" + VbCrlf
        end if
        sqlStr = sqlStr + " and r.returnmethod='" + FRectReturnmethod + "'" + VbCrlf

        if (FRectSearchString<>"") then
            sqlStr = sqlStr + " and a." + FRectSearchType + "='" + FRectSearchString + "'"
        end if

        if (FRectNotInputOnly = "Y") then
            sqlStr = sqlStr + " and ((IsNull(r.rebankname, '') = '') or (IsNull(r.rebankownername, '') = '')) "
        elseif (FRectNotInputOnly = "N") then
            sqlStr = sqlStr + " and ((IsNull(r.rebankname, '') <> '') and (IsNull(r.rebankownername, '') <> '') and (r.encaccount is not NULL)) "
        end if

        if FRectUploadState="notupload" then
            sqlStr = sqlStr + " and r.upfiledate is NULL"
        elseif FRectUploadState="uploaded" then
            sqlStr = sqlStr + " and r.upfiledate is Not NULL"
        end if

        if FRectUpfiledate<>"" then
            sqlStr = sqlStr + " and r.upfiledate='" + FRectUpfiledate + "'"
        end if

        sqlStr = sqlStr + " order by a.id desc"

        if session("ssBctId")="tozzinet" then
            response.write sqlStr & "<br>"
		else
            'response.write sqlStr & "<br>"
        end if
        rsget.pagesize = FPageSize

        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSASRefundInfoItem
                FItemList(i).Fsitegubun           = "�ٹ�����"

                FItemList(i).Fasid                = rsget("asid")

                FItemList(i).Forgsubtotalprice    = rsget("orgsubtotalprice")
                FItemList(i).Forgitemcostsum      = rsget("orgitemcostsum")
                FItemList(i).Forgbeasongpay       = rsget("orgbeasongpay")
                FItemList(i).Forgmileagesum       = rsget("orgmileagesum")
                FItemList(i).Forgcouponsum        = rsget("orgcouponsum")
                FItemList(i).Forgallatdiscountsum = rsget("orgallatdiscountsum")

                FItemList(i).Frefundrequire       = rsget("refundrequire")
                FItemList(i).Frefundresult        = rsget("refundresult")
                FItemList(i).Freturnmethod        = rsget("returnmethod")

                FItemList(i).Frefundmileagesum    = rsget("refundmileagesum")
                FItemList(i).Frefundcouponsum     = rsget("refundcouponsum")
                FItemList(i).Fallatsubtractsum    = rsget("allatsubtractsum")

                FItemList(i).Frefunditemcostsum   = rsget("refunditemcostsum")
                FItemList(i).Frefundbeasongpay    = rsget("refundbeasongpay")
                FItemList(i).Frefunddeliverypay   = rsget("refunddeliverypay")
                FItemList(i).Frefundadjustpay     = rsget("refundadjustpay")
                FItemList(i).Fcanceltotal         = rsget("canceltotal")

                FItemList(i).Frebankname          = rsget("rebankname")
                FItemList(i).Frebankaccount       = rsget("rebankaccount")
                FItemList(i).Frebankownername     = rsget("rebankownername")
                FItemList(i).FpaygateTid          = rsget("paygateTid")
                FItemList(i).Fencmethod           = rsget("encmethod")
                FItemList(i).Fdecaccount          = rsget("decaccount")

                FItemList(i).FpaygateresultTid    = rsget("paygateresultTid")
                FItemList(i).FpaygateresultMsg    = rsget("paygateresultMsg")
                FItemList(i).Fupfiledate          = rsget("cvupfiledate")

                FItemList(i).FreturnmethodName    = rsget("returnmethodName")

                FItemList(i).FOrderSerial         = rsget("orderserial")
                FItemList(i).Fuserid              = rsget("userid")
                FItemList(i).Fcustomername        = db2html(rsget("customername"))
                FItemList(i).Fregdate             = rsget("regdate")

                FItemList(i).Fcurrstate           = rsget("currstate")
                FItemList(i).Fconfirmregmsg       = rsget("confirmregmsg")
                FItemList(i).Fconfirmfinishmsg    = rsget("confirmfinishmsg")
                FItemList(i).Fconfirmfinishdate   = rsget("confirmfinishdate")

                FItemList(i).FIBK_TIDX          = rsget("TIDX")
                FItemList(i).FIBK_EB_USED       = rsget("EB_USED")
                FItemList(i).FIBK_PROC_YN       = rsget("PROC_YN")
                FItemList(i).FIBK_PROC_DATE     = rsget("PROC_DATE")
                FItemList(i).FIBK_ERR_MSG       = rsget("ERR_MSG")
                FItemList(i).FIBK_TEN_STATUS    = rsget("TEN_STATUS")
				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close
    End Sub

    public Sub GetRefundRequireAcademyList
        dim sqlStr,i
        sqlStr = " select count(a.id) as cnt from [db_academy].[dbo].tbl_academy_as_list a with (nolock)" + VbCrlf
        sqlStr = sqlStr + " join [db_academy].[dbo].tbl_academy_as_refund_info r with (nolock) on a.id=r.asid" + VbCrlf
        sqlStr = sqlStr + " where a.divcd='A003'" + VbCrlf
        sqlStr = sqlStr + " and a.deleteyn='N'" + VbCrlf
        if (FRectCurrstate<>"") then
            sqlStr = sqlStr + " and a.currstate='" + FRectCurrstate + "'" + VbCrlf
        end if

        if (FRectNotInputOnly = "Y") then
            sqlStr = sqlStr + " and ((IsNull(r.rebankname, '') = '') or (IsNull(r.rebankownername, '') = '')) "
        end if

        sqlStr = sqlStr + " and r.returnmethod='" + FRectReturnmethod + "'" + VbCrlf

        if (FRectSearchString<>"") then
            sqlStr = sqlStr + " and a." + FRectSearchType + "='" + FRectSearchString + "'"
        end if

        if FRectUploadState="notupload" then
            sqlStr = sqlStr + " and r.upfiledate is NULL"
        elseif FRectUploadState="uploaded" then
            sqlStr = sqlStr + " and r.upfiledate is Not NULL"
       end if

        if FRectUpfiledate<>"" then
            sqlStr = sqlStr + " and r.upfiledate='" + FRectUpfiledate + "'"
        end if

        rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " a.*, r.*, f.confirmregmsg, f.confirmfinishmsg, f.confirmfinishdate," + VbCrlf
        sqlStr = sqlStr + " convert(varchar(19),r.upfiledate,21) as cvupfiledate, c.comm_name as returnmethodName " + VbCrlf
        sqlStr = sqlStr + " ,K.TIDX,K.EB_USED,K.PROC_YN,K.PROC_DATE,K.ERR_MSG,K.TEN_STATUS"
		sqlStr = sqlStr + " , IsNull(r.encmethod, '') as encmethod "
        sqlStr = sqlStr + " , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_academy.dbo.uf_DecAcctPH1(r.encaccount), '') ELSE '' END) as decaccount "
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list a with (nolock)" + VbCrlf
        sqlStr = sqlStr + " join [db_academy].[dbo].tbl_academy_as_refund_info r with (nolock) on a.id=r.asid" + VbCrlf
        sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_cs_comm_code c with (nolock) on r.returnmethod=c.comm_cd" + VbCrlf
        sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_as_confirm f with (nolock) on a.id=f.asid" + VbCrlf
        sqlStr = sqlStr + " left join [TENDB].db_log.dbo.tbl_IBK_ERP_ICHE_DATA K with (nolock)"
        sqlStr = sqlStr + "     on r.IBK_TIDX=K.TIDX and IsNull(K.SITEGUBUN, '10x10') = 'academy' "
        sqlStr = sqlStr + " where a.divcd='A003'" + VbCrlf
        sqlStr = sqlStr + " and a.deleteyn='N'" + VbCrlf
        if (FRectCurrstate<>"") then
            sqlStr = sqlStr + " and a.currstate='" + FRectCurrstate + "'" + VbCrlf
        end if
        sqlStr = sqlStr + " and r.returnmethod='" + FRectReturnmethod + "'" + VbCrlf

        if (FRectSearchString<>"") then
            sqlStr = sqlStr + " and a." + FRectSearchType + "='" + FRectSearchString + "'"
        end if

        if (FRectNotInputOnly = "Y") then
            sqlStr = sqlStr + " and ((IsNull(r.rebankname, '') = '') or (IsNull(r.rebankownername, '') = '')) "
        end if

        if FRectUploadState="notupload" then
            sqlStr = sqlStr + " and r.upfiledate is NULL"
        elseif FRectUploadState="uploaded" then
            sqlStr = sqlStr + " and r.upfiledate is Not NULL"
        end if

        if FRectUpfiledate<>"" then
            sqlStr = sqlStr + " and r.upfiledate='" + FRectUpfiledate + "'"
        end if

        sqlStr = sqlStr + " order by a.id desc"

        'response.write sqlStr & "<Br>"
        rsACADEMYget.pagesize = FPageSize

        rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CCSASRefundInfoItem
                FItemList(i).Fsitegubun           = "��ī����"

                FItemList(i).Fasid                = rsACADEMYget("asid")

                FItemList(i).Forgsubtotalprice    = rsACADEMYget("orgsubtotalprice")
                FItemList(i).Forgitemcostsum      = rsACADEMYget("orgitemcostsum")
                FItemList(i).Forgbeasongpay       = rsACADEMYget("orgbeasongpay")
                FItemList(i).Forgmileagesum       = rsACADEMYget("orgmileagesum")
                FItemList(i).Forgcouponsum        = rsACADEMYget("orgcouponsum")
                FItemList(i).Forgallatdiscountsum = rsACADEMYget("orgallatdiscountsum")

                FItemList(i).Frefundrequire       = rsACADEMYget("refundrequire")
                FItemList(i).Frefundresult        = rsACADEMYget("refundresult")
                FItemList(i).Freturnmethod        = rsACADEMYget("returnmethod")

                FItemList(i).Frefundmileagesum    = rsACADEMYget("refundmileagesum")
                FItemList(i).Frefundcouponsum     = rsACADEMYget("refundcouponsum")
                FItemList(i).Fallatsubtractsum    = rsACADEMYget("allatsubtractsum")

                FItemList(i).Frefunditemcostsum   = rsACADEMYget("refunditemcostsum")
                FItemList(i).Frefundbeasongpay    = rsACADEMYget("refundbeasongpay")
                FItemList(i).Frefunddeliverypay   = rsACADEMYget("refunddeliverypay")
                FItemList(i).Frefundadjustpay     = rsACADEMYget("refundadjustpay")
                FItemList(i).Fcanceltotal         = rsACADEMYget("canceltotal")

                FItemList(i).Frebankname          = rsACADEMYget("rebankname")
                FItemList(i).Frebankaccount       = rsACADEMYget("rebankaccount")
                FItemList(i).Frebankownername     = rsACADEMYget("rebankownername")
                FItemList(i).FpaygateTid          = rsACADEMYget("paygateTid")
                FItemList(i).Fencmethod           = rsACADEMYget("encmethod")
                FItemList(i).Fdecaccount          = rsACADEMYget("decaccount")

                FItemList(i).FpaygateresultTid    = rsACADEMYget("paygateresultTid")
                FItemList(i).FpaygateresultMsg    = rsACADEMYget("paygateresultMsg")
                FItemList(i).Fupfiledate          = rsACADEMYget("cvupfiledate")

                FItemList(i).FreturnmethodName    = rsACADEMYget("returnmethodName")

                FItemList(i).FOrderSerial         = rsACADEMYget("orderserial")
                FItemList(i).Fuserid              = rsACADEMYget("userid")
                FItemList(i).Fcustomername        = db2html(rsACADEMYget("customername"))
                FItemList(i).Fregdate             = rsACADEMYget("regdate")

                FItemList(i).Fcurrstate           = rsACADEMYget("currstate")
                FItemList(i).Fconfirmregmsg       = rsACADEMYget("confirmregmsg")
                FItemList(i).Fconfirmfinishmsg    = rsACADEMYget("confirmfinishmsg")
                FItemList(i).Fconfirmfinishdate   = rsACADEMYget("confirmfinishdate")

                FItemList(i).FIBK_TIDX          = rsACADEMYget("TIDX")
                FItemList(i).FIBK_EB_USED       = rsACADEMYget("EB_USED")
                FItemList(i).FIBK_PROC_YN       = rsACADEMYget("PROC_YN")
                FItemList(i).FIBK_PROC_DATE     = rsACADEMYget("PROC_DATE")
                FItemList(i).FIBK_ERR_MSG       = rsACADEMYget("ERR_MSG")
                FItemList(i).FIBK_TEN_STATUS    = rsACADEMYget("TEN_STATUS")
				rsACADEMYget.moveNext
				i=i+1
			loop
		end if

		rsACADEMYget.Close
    End Sub

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 10
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
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
%>
