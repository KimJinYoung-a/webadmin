<%


function drawSelectBoxCSCommCombo(selectBoxName,selectedId,groupCode,onChangefunction)
   dim tmp_str,sqlStr
   %>
     <select class="select" name="<%=selectBoxName%>" <%= onChangefunction %> >
     <option value='' <%if selectedId="" then response.write " selected" %> >����</option>
   <%
       sqlStr = " select comm_cd,comm_name "
       sqlStr = sqlStr + " from  "
       sqlStr = sqlStr + " " & TABLE_CS_COMMON_CODE  & " "
       sqlStr = sqlStr + " where comm_group='" + groupCode + "' "
       sqlStr = sqlStr + " and comm_isDel='N' "
       sqlStr = sqlStr + " order by comm_cd "

       rsget.Open sqlStr,dbget,1

       if  not rsget.EOF  then
           do until rsget.EOF
               if LCase(selectedId) = LCase(rsget("comm_cd")) then
                   tmp_str = " selected"
               end if
               response.write("<option value='" & rsget("comm_cd") & "' " & tmp_str & ">" + db2html(rsget("comm_name")) + " </option>")
               tmp_str = ""
               rsget.MoveNext
           loop
       end if
       rsget.close
   %>
       </select>
   <%
End function

function drawSelectBoxCancelTypeBox(selectBoxName,selectedId,orgPaymethod,divcd,onChangefunction)
    dim BufStr, selectStr
    BufStr = "<select class='select' name='returnmethod' " + onChangefunction + ">"
    BufStr = BufStr + "<option value=''>����</option>"

    if (selectedId="R000") then selectStr="selected"
        BufStr = BufStr + "<option value='R000' " + selectStr + ">ȯ�� ����</option>"
    selectStr = ""

    if (orgPaymethod="100") or (orgPaymethod="110") then
        if (selectedId="R100") then selectStr="selected"
        BufStr = BufStr + "<option value='R100' " + selectStr + ">�ſ�ī�� ���</option>"

		if True or application("Svr_Info") = "Dev" then
			if (orgPaymethod = "100") then
				selectStr = ""
				if (selectedId="R120") then selectStr="selected"
		        BufStr = BufStr + "<option value='R120' " + selectStr + ">�ſ�ī�� �κ����</option>"
		    end if
        end if
    elseif (orgPaymethod="20")  then
        if (selectedId="R020") then selectStr="selected"
        BufStr = BufStr + "<option value='R020' " + selectStr + ">�ǽð���ü ���</option>"

		if True or application("Svr_Info") = "Dev" then
			if (orgPaymethod = "20") then
				selectStr = ""
				if (selectedId="R022") then selectStr="selected"
		        BufStr = BufStr + "<option value='R022' " + selectStr + ">�ǽð���ü �κ����</option>"
		    end if
        end if
    elseif (orgPaymethod="400")  then
        if (selectedId="R400") then selectStr="selected"
        BufStr = BufStr + "<option value='R400' " + selectStr + ">�޴������� ���</option>"
    elseif (orgPaymethod="80")  then
        if (selectedId="R080") then selectStr="selected"
        BufStr = BufStr + "<option value='R080' " + selectStr + ">All@ī�� ���</option>"
    elseif (orgPaymethod="50") then
        if (selectedId="R050") then selectStr="selected"
        BufStr = BufStr + "<option value='R050' " + selectStr + ">���������� ���</option>"
    end if

    selectStr = ""

    if (selectedId="R007") then selectStr="selected"
    BufStr = BufStr + "<option value='R007' " + selectStr + ">������ ȯ��</option>"

    selectStr = ""

    if (selectedId="R900") then selectStr="selected"
    BufStr = BufStr + "<option value='R900' " + selectStr + ">���ϸ��� ȯ��</option>"
    BufStr = BufStr + "</select>"

    response.write BufStr
end function


''��� ���μ���
public function fnIsCancelProcess(idivcd)
    fnIsCancelProcess = (idivcd = "A008")
end function

''��ǰ ���μ���(ȸ��, �±�ȯ ȸ��)
public function fnIsReturnProcess(idivcd)
    fnIsReturnProcess = (idivcd = "A004") or (idivcd = "A010") or (idivcd = "A011")
end function

public function fnIsRefundProcess(idivcd)
    fnIsRefundProcess = (idivcd = "A003") or (idivcd = "A005")
end function

''�����߼�, ���񽺹߼�  ���μ���
public function fnIsServiceDeliverProcess(idivcd)
    fnIsServiceDeliverProcess = (idivcd = "A000") or (idivcd = "A001") or (idivcd = "A002")
end function

''Cs Detail ���� ����
Class CCSASDetailItem
    ''tbl_as_detail's
    public Fid
    public Fmasterid
    public Fgubun01
    public Fgubun02
    public Fgubun01name
    public Fgubun02name
    public Fregdetailstate
    public Fregitemno
    public Fconfirmitemno
    public Fcausediv
    public Fcausedetail
    public Fcausecontent

    ''tbl_order_detail's
    public Forderdetailidx
    public Forderserial
    public Fitemid
    public Fitemoption
    public Fmakerid
    public Fitemname
    public Fitemoptionname
    public Fitemcost
    public Fbuycash
    public Fitemno
	public Fprevreturnno
    public Fisupchebeasong
    public Fcancelyn

    public Foitemdiv
    public FodlvType
    public Fissailitem
    public Fitemcouponidx
    public Fbonuscouponidx

    public ForderDetailcurrstate
    public FdiscountAssingedCost    '' �ֹ��� ���εȰ��� ( ALL@ / %���α� �ݿ�)

    ''public FAllAtDiscountedPrice

    ''tbl_item's
    public FSmallImage

    ''��ü ������� ��ǰ ��ۺ� ���� ����
    public function IsUpcheParticleDeliverPayCodeItem
        IsUpcheParticleDeliverPayCodeItem = (Fitemid=0) and (Left(Fitemoption,2)="90")
    end function

    ''��ü ������� ��ǰ���� ����
    public function IsUpcheParticleDeliverItem
        IsUpcheParticleDeliverItem = (FodlvType=9)
    end function

    ''��ǰ�� ����ϴ� ��ǰ����(All@ ���ΰ�, %���� ���ΰ� �ݿ�)
    public function GetOrgPayedItemPrice()
        GetOrgPayedItemPrice = Fitemcost

        if (FdiscountAssingedCost=0) then
            ''�������
            GetOrgPayedItemPrice = Fitemcost-getAllAtDiscountedPrice
        else
            if (FdiscountAssingedCost<>Fitemcost) then
                GetOrgPayedItemPrice = FdiscountAssingedCost
            end if
        end if
    end function

    ''All@ ���εȰ���
    public function getAllAtDiscountedPrice()
        getAllAtDiscountedPrice =0
        ''���� ��ǰ���� ���εǴ°�� �߰����ξ���.
        ''���ϸ����� ��ǰ �߰� ���� ����.
	    ''���ϻ�ǰ �߰����� ����
	    '' 20070901�߰� : �������� ���ʽ��������� �߰����� ����.

'	    if (FdiscountAssingedCost=0) then
'	        ''�������
'            if (Fitemcouponidx<>0) or (IsMileShopSangpum) or (Fissailitem="Y") then
'    			getAllAtDiscountedPrice = 0
'    		else
'    			getAllAtDiscountedPrice = round(((1-0.94) * FItemCost / 100) * 100 ) * FItemNo
'    		end if
'    	else

'''			'�ϴ� ����.
'''    	    if (IsNULL(Fbonuscouponidx) or (Fbonuscouponidx=0)) and (Fitemcost>FdiscountAssingedCost) then
'''    	            getAllAtDiscountedPrice = Fitemcost-FdiscountAssingedCost
'''    	    else
'''    	        getAllAtDiscountedPrice = 0
'''    	    end if

'    	end if
    end function

    '' %���α� ���αݾ� or ī�� ���αݾ�
    public function getPercentBonusCouponDiscountedPrice()
        getPercentBonusCouponDiscountedPrice = 0
'        if (Fitemcost>FdiscountAssingedCost) then
'                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
'        end if

        if (FdiscountAssingedCost=0) then
	        ''�������
	        ''getPercentBonusCouponDiscountedPrice = Fitemcost*
	    else
            if (Fbonuscouponidx<>0)  and (Fitemcost>FdiscountAssingedCost) then
                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
            end if
        end if
    end function

    ''���ϸ����� ��ǰ
    public function IsMileShopSangpum()
		IsMileShopSangpum = false

		if Foitemdiv="82" then
			IsMileShopSangpum = true
		end if
	end function

    public function GetDefaultRegNo(IsRegState)
        if (IsRegState) then
            GetDefaultRegNo = Fitemno
        else
            GetDefaultRegNo = Fregitemno
        end if
    end function

    ''CsAction ������ ��ǰ ���� ���� ���ɿ���
    public function IsItemNoEditEnabled(byval idivcd)
        IsItemNoEditEnabled = false

        if (Fcancelyn="Y") then Exit function

        if (fnIsCancelProcess(idivcd)) then
            IsItemNoEditEnabled = true

            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=false
        elseif (fnIsReturnProcess(idivcd)) then
            ''��ǰ ����
            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=true
        elseif (fnIsServiceDeliverProcess(idivcd)) then
            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=true

        else

        end if
    end function


    ''CsAction ������ ��ǰ�� üũ ���ɿ���
    public function IsCheckAvailItem(byval iIpkumdiv, byval iMasterCancelYn, byval idivcd)
        IsCheckAvailItem = false

        if (Fcancelyn="Y") then Exit function
        if (iMasterCancelYn<>"N") then Exit function

        if (fnIsCancelProcess(idivcd)) then
            IsCheckAvailItem = true
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=false

        elseif (fnIsReturnProcess(idivcd)) then
            ''��ǰ ����
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=true

            if (FItemId=0) then IsCheckAvailItem=true
        elseif (idivcd="A006") then
            ''���� ���ǻ���
            IsCheckAvailItem=true

            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=false
        elseif (idivcd="A009") then
            ''��Ÿ����(�޸�) - All case Avail
            IsCheckAvailItem=true
        elseif (idivcd="A700") then
            ''��Ÿ���� - All case Avail
            IsCheckAvailItem=true
        elseif (idivcd = "A002") then
            if Fitemid=0 then
                IsCheckAvailItem=false
            else
                IsCheckAvailItem=true
            end if
        elseif (idivcd = "A001") then
            ''����, ����
            if (ForderDetailcurrstate>=7) or ((Fcancelyn="A") and (iIpkumdiv>=7)) then IsCheckAvailItem=true
        elseif (idivcd = "A000") then
            ''�±�ȯ
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=true
        else

        end if
    end function

    ''CsAction ������ ��ǰ�� ����Ʈ üũ��
    public function IsDefaultCheckedItem(byval iIpkumdiv, byval iMasterCancelYn, byval idivcd, byval ckAll)
        IsDefaultCheckedItem =false

        if (Not IsCheckAvailItem(iIpkumdiv,iMasterCancelYn,idivcd)) then Exit function

        if (fnIsCancelProcess(idivcd)) then
            if (ckAll<>"") then
                IsDefaultCheckedItem = true
            else
                IsDefaultCheckedItem = false
            end if

            if (Fcancelyn="Y") or (iMasterCancelYn<>"N") then IsDefaultCheckedItem=false

            if (ForderDetailcurrstate>=3) then IsDefaultCheckedItem=false
        elseif (fnIsReturnProcess(idivcd)) then
            ''��ǰ�����ΰ�� - No action
        elseif (idivcd="A006") then
            ''���� ���ǻ��� - No action
        elseif (idivcd="A009") then
            ''��Ÿ����(�޸�) - No action
        else

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
		CancelStateColor = "#000000"

		if Fcancelyn="Y" then
			CancelStateColor ="#FF0000"
		elseif Fcancelyn="D" then
			CancelStateColor ="#FF0000"
		elseif Fcancelyn="A" then
			CancelStateColor ="#0000FF"
		end if
	end function

	''order Detail's State Name : ������
	Public function GetStateName()
        if ForderDetailcurrstate="2" then
            if (Fisupchebeasong="Y") then
		        GetStateName = "��ü�뺸"
		    else
		        GetStateName = "�����뺸"
		    end if
	    elseif ForderDetailcurrstate="3" then
		    GetStateName = "��ǰ�غ�"
	    elseif ForderDetailcurrstate="7" then
		    GetStateName = "���Ϸ�"
	    else
		    GetStateName = ForderDetailcurrstate
	    end if
	end Function

	'' ��Ͻ� ����..
	Public function GetRegDetailStateName()
        if (Fregdetailstate="2") then
            if (Fisupchebeasong="Y") then
		        GetRegDetailStateName = "��ü�뺸"
		    else
		        GetRegDetailStateName = "�����뺸"
		    end if
	    elseif Fregdetailstate="3" then
		    GetRegDetailStateName = "��ǰ�غ�"
	    elseif Fregdetailstate="7" then
		    GetRegDetailStateName = "���Ϸ�"
	    else
		    GetRegDetailStateName = "----"
	    end if
	end Function

	''order Detail's State color
	public function GetStateColor()
	    if ForderDetailcurrstate="2" then
			GetStateColor="#000000"
		elseif ForderDetailcurrstate="3" then
			GetStateColor="#CC9933"
		elseif ForderDetailcurrstate="7" then
			GetStateColor="#FF0000"
		else
			GetStateColor="#000000"
		end if
	end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class


''ȯ�� ���� ����
Class CCSASRefundInfoItem
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

    public FencMethod           ''��ȣȭ���
    public FencAccount          ''��ȣȭ ���¹�ȣ
    public FdecAccount          ''��ȣȭ ���¹�ȣ

    public FpaygateresultTid
    public FpaygateresultMsg

    public FreturnmethodName    ''ȯ�ҹ�ĸ�

    public rebankCode

    public Fupfiledate          ''ȯ������ �ۼ���

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
End Class

''�� ȸ��, �±�ȯ.. �ּ�������
Class CCSDeliveryItem
    public Fasid
    public Freqname
    public Freqphone
    public Freqhp
    public Freqzipcode
    public Freqzipaddr
    public Freqetcaddr
    public Freqetcstr
    public Fsongjangdiv
    public Fsongjangno
    public Fregdate
    public Fsenddate


    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCSReturnAddressItem
    public Fbrandid
    public Fbrandname

    public Fstreetname_kor
    public Fstreetname_eng

    public FreturnName
    public FreturnPhone
    public Freturnhp
    public FreturnEmail

    public FreturnZipcode
    public FreturnZipaddr
    public FreturnEtcaddr

    public Fsongjangdiv	'�ù��
    public Fsongjangno

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

''��ǰ �ּ��� ����
Class CCSReturnAddress
	public FItemList()

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public Fbrandid
    public Fbrandname

    public Fstreetname_kor
    public Fstreetname_eng

    public FreturnName
    public FreturnPhone
    public Freturnhp
    public FreturnEmail

    public FreturnZipcode
    public FreturnZipaddr
    public FreturnEtcaddr

    public Fsongjangdiv
    public Fsongjangno

    public FRectMakerid
    public FRectGroupCode

    public sub GetReturnAddress()
        dim sqlStr
        sqlStr = " select company_name, deliver_phone, deliver_hp, return_zipcode, return_address, return_address2"
        sqlStr = sqlStr + " from " & TABLE_PARTNER & ""
        sqlStr = sqlStr + " where id='" + FRectMakerid + "'"

        rsget.Open sqlStr, dbget, 1

        if Not rsget.Eof then
            FreturnName      = db2html(rsget("company_name"))
            FreturnPhone     = db2html(rsget("deliver_phone"))
            Freturnhp        = db2html(rsget("deliver_hp"))
            FreturnZipcode   = rsget("return_zipcode")
            FreturnZipaddr   = db2html(rsget("return_address"))
            FreturnEtcaddr   = db2html(rsget("return_address2"))
            Fsongjangdiv     = ""
            Fsongjangno      = ""

        end if
        rsget.Close
    end sub

    public sub GetBrandReturnAddress()
    	'GetReturnAddress() ���� company_name �� FreturnName �� �����ϹǷ� ���� �Լ� ����
        dim sqlStr
        sqlStr = " select id as brandid, company_name as brandname, socname_kor as streetname_kor, socname as streetname_eng, return_zipcode, return_address, return_address2, deliver_phone, deliver_hp, deliver_name, deliver_email, defaultsongjangdiv "
        sqlStr = sqlStr + " from " & TABLE_PARTNER & " p, " & TABLE_USER_C & " c "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and p.id = c.userid "
        sqlStr = sqlStr + " and p.id='" + FRectMakerid + "'"

        rsget.Open sqlStr, dbget, 1

        if Not rsget.Eof then

			Fbrandid         = rsget("brandid")
			Fbrandname       = db2html(rsget("brandname"))

			Fstreetname_kor  = db2html(rsget("streetname_kor"))
			Fstreetname_eng  = db2html(rsget("streetname_eng"))

			FreturnName      = rsget("deliver_name")
			FreturnPhone     = rsget("deliver_phone")
			Freturnhp        = rsget("deliver_hp")
			FreturnEmail     = rsget("deliver_email")

            FreturnZipcode   = rsget("return_zipcode")
            FreturnZipaddr   = db2html(rsget("return_address"))
            FreturnEtcaddr   = db2html(rsget("return_address2"))

            Fsongjangdiv     = rsget("defaultsongjangdiv")

        end if
        rsget.Close
    end sub

    public sub GetReturnAddressList()
        dim sqlStr, i

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from " & TABLE_PARTNER & " p, " & TABLE_USER_C & " c "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and p.id = c.userid "
        sqlStr = sqlStr + " and p.groupid ='" + FRectGroupCode + "'"

        rsget.Open sqlStr, dbget, 1
            FTotalCount = rsget("cnt")
        rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " id as brandid, company_name as brandname, socname_kor as streetname_kor, socname as streetname_eng, return_zipcode, return_address, return_address2, deliver_phone, deliver_hp, deliver_name, deliver_email, defaultsongjangdiv "
        sqlStr = sqlStr + " from " & TABLE_PARTNER & " p, " & TABLE_USER_C & " c "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and p.id = c.userid "
        sqlStr = sqlStr + " and p.groupid ='" + FRectGroupCode + "'"
        sqlStr = sqlStr + " order by id "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

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
				set FItemList(i) = new CCSReturnAddressItem

				FItemList(i).Fbrandid         = rsget("brandid")
				FItemList(i).Fbrandname       = db2html(rsget("brandname"))

				FItemList(i).Fstreetname_kor  = db2html(rsget("streetname_kor"))
				FItemList(i).Fstreetname_eng  = db2html(rsget("streetname_eng"))

				FItemList(i).FreturnName      = rsget("deliver_name")
				FItemList(i).FreturnPhone     = rsget("deliver_phone")
				FItemList(i).Freturnhp        = rsget("deliver_hp")
				FItemList(i).FreturnEmail     = rsget("deliver_email")

	            FItemList(i).FreturnZipcode   = rsget("return_zipcode")
	            FItemList(i).FreturnZipaddr   = db2html(rsget("return_address"))
	            FItemList(i).FreturnEtcaddr   = db2html(rsget("return_address2"))

	            FItemList(i).Fsongjangdiv     = rsget("defaultsongjangdiv")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

    end sub

    Private Sub Class_Initialize()
        FreturnName     = "(��)�ٹ�����"
        FreturnPhone    = "1644-6030"
        Freturnhp       = ""

        FreturnZipcode  = "11154"
        FreturnZipaddr  = "��⵵ ��õ�� ������ ����������2�� 83"
        FreturnEtcaddr  = "�ٹ����� ��������"

        Fsongjangdiv    = "24"
        Fsongjangno     = ""

		FCurrPage = 1
		FPageSize = 20
		FScrollCount = 10
    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

''�귣�庰 CS �޸�
Class CCSBrandMemo
    public Fbrandid

	public Fis_return_allow

	public Fvacation_startday
	public Fvacation_endday

	public Ftel_start
	public Ftel_end

	public Fis_saturday_work

	public Fbrand_comment

	public Flast_modifyday

    public FRectMakerid

    public sub GetBrandMemo()
        dim sqlStr

        sqlStr = " select brandid, is_return_allow, vacation_startday, vacation_endday, tel_start, tel_end, is_saturday_work, brand_comment, last_modifyday "
        sqlStr = sqlStr + " from " & TABLE_CS_BRAND_MEMO & " "
        sqlStr = sqlStr + " where brandid='" + FRectMakerid + "'"
        rsget.Open sqlStr, dbget, 1

        if Not rsget.Eof then
            Fbrandid         		= rsget("brandid")
            Fis_return_allow		= rsget("is_return_allow")
            Fvacation_startday  	= rsget("vacation_startday")
            Fvacation_endday     	= rsget("vacation_endday")
            Ftel_start         		= rsget("tel_start")
            Ftel_end         		= rsget("tel_end")
            Fis_saturday_work       = rsget("is_saturday_work")
            Fbrand_comment          = db2html(rsget("brand_comment"))
            Flast_modifyday         = rsget("last_modifyday")

        end if
        rsget.Close
    end sub

    Private Sub Class_Initialize()
        '
    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCsConfirmItem
    public Fasid
    public Fcha                 '''2009�߰�
    public Fconfirmregmsg
    public Fconfirmreguserid
    public Fconfirmregdate
    public Fconfirmfinishmsg
    public Fconfirmfinishuserid
    public Fconfirmfinishdate

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCSASMasterItem

    public Fid
    public Fdivcd
    public Fgubun01
    public Fgubun02

    public FdivcdName
    public Fgubun01Name
    public Fgubun02Name

    public FdivcdColor
    public Fgubun01Color
    public Fgubun02Color

    public Forderserial
    public Fcustomername
    public Fuserid
    public Fwriteuser
    public Ffinishuser
    public Ftitle
    public Fcontents_jupsu
    public Fcontents_finish
    public Fcurrstate
    public FcurrstateName
    public FcurrstateColor
    public Fregdate
    public Ffinishdate

    public Fsongjangdiv
    public Fsongjangno
    public Fbeasongdate

    public Frequireupche
    public Fmakerid
    public Fdeleteyn
    public Fextsitename

    '' tbl_as_refund_info's
    public Frefundrequire
    public Frefundresult

    '' tbl_as_upcheAddjungsan
    public Fadd_upchejungsandeliverypay
    public Fadd_upchejungsancause

    public Frefminusorderserial  ''2017/03/27
    
    public Fopentitle           ''�� ���� Title
    public Fopencontents        ''�� ���� ����
    public Fsitegubun           '' 10x10 or theFingers

    public FErrMsg
    public FAuthcode


    public function IsAsRegAvail(byval iIpkumdiv, byval iCancelYn, byref descMsg)
        IsAsRegAvail = false
        if (iIpkumdiv<2) then
            IsAsRegAvail = false
            descMsg      = "������ �ֹ��� �Ǵ� ���� �ֹ����� �ƴմϴ�. "
            exit function
        end if

        if (IsCancelProcess) then
            IsAsRegAvail = false

            if (iCancelYn<>"N") then
                IsAsRegAvail = false
                descMsg      = "�̹� ��ҵ� �ŷ��Դϴ�. - ��� �Ұ��� "
                exit function
            end if

            IsAsRegAvail = true

        elseif (IsReturnProcess) then
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail = false
                descMsg      = "��� �Ϸ�/ �Ϻ� ��� ���°� �ƴմϴ�. - ��ǰ ���� �Ұ��� "
                exit function
            end if

            if (iCancelYn<>"N") then
                IsAsRegAvail = false
                descMsg      = "��ҵ� �ŷ��Դϴ�. - ��ǰ ���� �Ұ��� "
                exit function
            end if

            IsAsRegAvail = true
        elseif (Fdivcd = "A006") then
            '' ���� ���ǻ���
            IsAsRegAvail = true

            if (iIpkumdiv>=8) then
                IsAsRegAvail = false
                descMsg      = "��� ���� ���°� �ƴմϴ�. - ���� ���ǻ��� ���� �Ұ��� "
                exit function
            end if
        elseif (Fdivcd = "A009") then
            '' ��Ÿ����
            IsAsRegAvail = true
        elseif  (Fdivcd = "A002") then
            ''���񽺹߼� :��� �����ϰ� ����..
            IsAsRegAvail = true
        elseif (Fdivcd = "A001") then
            ''������߼�,
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail = false
                descMsg      = "��� �Ϸ�/ �Ϻ� ��� ���°� �ƴմϴ�. - ����/���� �߼� ���� �Ұ��� "
                exit function
            end if

            IsAsRegAvail = true
        elseif (Fdivcd = "A000") then
            ''�±�ȯ
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail = false
                descMsg      = "��� �Ϸ�/ �Ϻ� ��� ���°� �ƴմϴ�. - �±�ȯ ���� �Ұ��� "
                exit function
            end if

            IsAsRegAvail = true
        elseif (Fdivcd = "A003") then
            ''ȯ�ҿ�û
            IsAsRegAvail = true
        elseif (Fdivcd = "A005") then
            ''������ ����Ʈ ���� üũ
            IsAsRegAvail = true
         elseif (Fdivcd = "A700") then
            ''��ü ��Ÿ ����.
            IsAsRegAvail = true
        else
            descMsg = "���� ���� �ʾҽ��ϴ�." + Fdivcd
        end if

    end function

    ''��� ���μ���
    public function IsCancelProcess()
        IsCancelProcess = fnIsCancelProcess(Fdivcd)
    end function

    ''��ǰ ���μ���
    public function IsReturnProcess()
        IsReturnProcess = fnIsReturnProcess(Fdivcd)
    end function

    ''ȯ�� ���μ���
    public function IsRefundProcess()
        IsRefundProcess = fnIsRefundProcess(Fdivcd)
    end function

    ''���� �߼� ���μ���
    public function IsServiceDeliverProcess()
        IsServiceDeliverProcess = fnIsServiceDeliverProcess(Fdivcd)
    end function

    public function IsRefundProcessRequire(iIpkumdiv, iCancelyn)
        FErrMsg = ""
        IsRefundProcessRequire = False

        if (iCancelyn ="Y") or (iCancelyn ="D") then Exit function

		if (iIpkumdiv<4) then  Exit function

        '' ���, ��ǰ����
        IsRefundProcessRequire = (IsCancelProcess) or (IsReturnProcess)
    end function

    public function IsRefundProcessRequireBeforePay(iIpkumdiv, iCancelyn)
        FErrMsg = ""
        IsRefundProcessRequireBeforePay = False

        if (iCancelyn ="Y") or (iCancelyn ="D") then Exit function

		'�ֹ� �Ϻ�����̰� ����� ���ϸ����� ��һ�ǰ�� �ݾ׺��� ū��� ���������� ��Ұ� �ʿ��ϴ�.
		'����� ���ϸ����� �Ϻ���� �� �� ����.
		'if (iIpkumdiv<4) then  Exit function

        '' ���, ��ǰ����
        IsRefundProcessRequireBeforePay = (IsCancelProcess) or (IsReturnProcess)
    end function

    ''���� �ʵ尡 �ʿ��� ����
    public function IsRequireSongjangNO()
        IsRequireSongjangNO = false

        IsRequireSongjangNO = (Fdivcd="A000") or (Fdivcd="A001") or (Fdivcd="A002") or (Fdivcd="A004") or (Fdivcd="A010") or (Fdivcd="A011")
    end function

    public function GetAsDivCDName()
        GetAsDivCDName = FdivcdName


    end function

    public function GetAsDivCDColor()
        GetAsDivCDColor = FdivcdName


    end function


    public function GetCurrstateName()
        GetCurrstateName = FcurrstateName
    end function

     public function GetCurrstateColor()
        GetCurrstateColor = FcurrstateColor
    end function

    public function GetCauseString()
        GetCauseString = Fgubun01Name
    end function

    public function GetCauseDetailString()
        GetCauseDetailString = Fgubun02Name
    end function



    Private Sub Class_Initialize()
        Fadd_upchejungsandeliverypay = 0
    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCSASList
    public FItemList()
    public FOneItem

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectUserID
    public FRectUserName
    public FRectOrderSerial
    public FRectStartDate
    public FRectEndDate
    public FRectSearchType
    public FRectIdx
    public FRectMakerid

    public FRectDivcd
    public FRectCurrstate

    public FRectCsAsID
    public FRectNotCsID
    ''
    public FDeliverPay
    public IsUpchebeasongExists
    public IsTenbeasongExists

    public FRectOldOrder

    ''��ü���
    public FRectOnlyJupsu


	Public FRectDeleteYN	' �������ܿ���
	Public FRectWriteUser	' �����ھ��̵� �˻�


    public Sub GetHisOldRefundInfo()
        dim i,sqlStr

        sqlStr = " select count(asid) as cnt "
        sqlStr = sqlStr + " from " & TABLE_CS_REFUND & " r, "
        sqlStr = sqlStr + " " & TABLE_CSMASTER & " a"
        sqlStr = sqlStr + " where a.userid='" + FRectUserID + "'"
        sqlStr = sqlStr + " and a.id=r.asid"
        sqlStr = sqlStr + " and a.divcd='A003'"
        sqlStr = sqlStr + " and r.returnmethod='R007'"
        sqlStr = sqlStr + " and a.deleteyn='N'"


        rsget.Open sqlStr, dbget, 1
            FTotalCount = rsget("cnt")
        rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " r.refundrequire, r.rebankname, r.rebankaccount, r.rebankownername, r.encmethod, r.encaccount "
		sqlStr = sqlStr + " , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_academy.dbo.uf_DecAcctPH1(r.encaccount), '') ELSE '' END) as decaccount "
        sqlStr = sqlStr + " from " & TABLE_CS_REFUND & " r, "
        sqlStr = sqlStr + " " & TABLE_CSMASTER & " a"
        sqlStr = sqlStr + " where a.userid='" + FRectUserID + "'"
        sqlStr = sqlStr + " and a.id=r.asid"
        sqlStr = sqlStr + " and a.divcd='A003'"
        sqlStr = sqlStr + " and r.returnmethod='R007'"
        sqlStr = sqlStr + " and a.deleteyn='N'"
        sqlStr = sqlStr + " order by r.asid desc"

        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

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

                FItemList(i).Frefundrequire         = rsget("refundrequire")
				FItemList(i).Frebankname            = rsget("rebankname")
                FItemList(i).Frebankaccount         = rsget("rebankaccount")
                FItemList(i).Frebankownername       = rsget("rebankownername")

                FItemList(i).FencMethod             = rsget("encmethod")
                FItemList(i).FencAccount            = rsget("encaccount")
                FItemList(i).FdecAccount            = rsget("decAccount")

                ''FItemList(i).FrebankCode            = rsget("rebankCode")
				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close
    end Sub

    public Sub GetOneRefundInfo()
        dim i,sqlStr

        sqlStr = "select r.* "
        sqlStr = sqlStr + " ,C1.comm_name as returnmethodName"
		sqlStr = sqlStr + " , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_academy.dbo.uf_DecAcctPH1(r.encaccount), '') ELSE '' END) as decaccount "
        sqlStr = sqlStr + " from " & TABLE_CS_REFUND & " r"
        sqlStr = sqlStr + "     left join " & TABLE_CS_COMMON_CODE  & " C1"
        sqlStr = sqlStr + "     on C1.comm_group='Z090'"
        sqlStr = sqlStr + "     and r.returnmethod=C1.comm_cd"
        sqlStr = sqlStr + " where asid=" + CStr(FRectCsAsID)

        rsget.Open sqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CCSASRefundInfoItem
        if Not rsget.Eof then

            FOneItem.Fasid                  = rsget("asid")
            FOneItem.Forgsubtotalprice      = rsget("orgsubtotalprice")
            FOneItem.Forgitemcostsum        = rsget("orgitemcostsum")
            FOneItem.Forgbeasongpay         = rsget("orgbeasongpay")
            FOneItem.Forgmileagesum         = rsget("orgmileagesum")
            FOneItem.Forgcouponsum          = rsget("orgcouponsum")
            FOneItem.Forgallatdiscountsum   = rsget("orgallatdiscountsum")

            FOneItem.Frefundrequire         = rsget("refundrequire")
            FOneItem.Frefundresult          = rsget("refundresult")
            FOneItem.Freturnmethod          = rsget("returnmethod")

            FOneItem.Frefundmileagesum      = rsget("refundmileagesum")
            FOneItem.Frefundcouponsum       = rsget("refundcouponsum")
            FOneItem.Fallatsubtractsum      = rsget("allatsubtractsum")

            FOneItem.Frefunditemcostsum     = rsget("refunditemcostsum")
            FOneItem.Frefundbeasongpay      = rsget("refundbeasongpay")
            FOneItem.Frefunddeliverypay     = rsget("refunddeliverypay")
            FOneItem.Frefundadjustpay       = rsget("refundadjustpay")
            FOneItem.Fcanceltotal           = rsget("canceltotal")

            FOneItem.Frebankname            = rsget("rebankname")
            FOneItem.Frebankaccount         = rsget("rebankaccount")
            FOneItem.Frebankownername       = rsget("rebankownername")
            FOneItem.FpaygateTid            = rsget("paygateTid")

			FOneItem.FencMethod             = rsget("encmethod")
			FOneItem.FencAccount            = rsget("encaccount")
			FOneItem.FdecAccount            = rsget("decAccount")

            FOneItem.FpaygateresultTid      = rsget("paygateresultTid")
            FOneItem.FpaygateresultMsg      = rsget("paygateresultMsg")


            FOneItem.FreturnmethodName      = rsget("returnmethodName")

            FOneItem.Fupfiledate      = rsget("upfiledate")
        end if
        rsget.Close
    end Sub

    public Sub GetCSASMasterList()
        dim i,sqlStr, AddSQL
        AddSQL = ""

        sqlStr = " select count(A.id) as cnt "
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " A"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_REFUND & " r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join " & TABLE_ORDERMASTER & " m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.sitename <> '" & EXCLUDE_SITENAME & "' "

		if (FRectSearchType="") then
		    if (FRectOrderSerial<>"") then
		        AddSQL = AddSQL + " and A.orderserial='" + FRectOrderSerial + "'"
		    end if
		elseif (FRectSearchType="upcheview") then
		    ''��ü�� ������
            AddSQL = AddSQL + " and divcd not in ('A005','A007')"
            AddSQL = AddSQL + " and deleteyn='N'"
            AddSQL = AddSQL + " and requireupche='Y' "
            AddSQL = AddSQL + " and makerid='" + CStr(FRectMakerid) + "' "

            if (FRectOnlyJupsu="on") then
                AddSQL = AddSQL + " and currstate='B001'"
            end if

            if (FRectCurrstate = "notfinish") then
	                AddSQL = AddSQL + " and A.currstate < 'B007' "
	        elseif (FRectCurrstate <> "") then
	                AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

            if (FRectUserName <> "") then
	                AddSQL = AddSQL + " and A.customername='" + CStr(FRectUserName) + "' "
	        end if

	        if (FRectOrderSerial <> "") then
	                AddSQL = AddSQL + " and A.orderserial='" + CStr(FRectOrderSerial) + "' "
	        end if

	        if (FRectUserID <> "") then
	                AddSQL = AddSQL + " and A.userid='" + CStr(FRectUserID) + "' "
	        end if
		elseif (FRectSearchType = "searchfield") then

	        if (FRectUserID <> "") then
	                AddSQL = AddSQL + " and A.userid='" + CStr(FRectUserID) + "' "
	        end if

	        if (FRectUserName <> "") then
	                AddSQL = AddSQL + " and A.customername='" + CStr(FRectUserName) + "' "
	        end if

	        if (FRectOrderSerial <> "") then
	                AddSQL = AddSQL + " and A.orderserial='" + CStr(FRectOrderSerial) + "' "
	        end if

	        if (FRectMakerid<>"") then
	                AddSQL = AddSQL + " and A.requireupche='Y' "
	                AddSQL = AddSQL + " and A.makerid='" + CStr(FRectMakerid) + "' "
	        end if

	        if (FRectStartDate <> "") then
	                AddSQL = AddSQL + " and A.regdate>='" + CStr(FRectStartDate) + "' "
	        end if

	        if (FRectEndDate <> "") then
	                AddSQL = AddSQL + " and A.regdate <'" + CStr(FRectEndDate) + "' "
	        end if

	        if (FRectCurrstate = "notfinish") then
	                AddSQL = AddSQL + " and A.currstate < 'B007' "
	        elseif (FRectCurrstate <> "") then
	                AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if


	        if (FRectDivcd <> "") then
	                AddSQL = AddSQL + " and A.divcd ='" + CStr(FRectDivcd) + "' "
	        end if

			if (FRectWriteUser <> "") then
					AddSQL = AddSQL + " and A.writeUser = '" + CStr(FRectWriteUser) + "' "
			end if

			if (FRectDeleteYN <> "") then
					AddSQL = AddSQL + " and A.deleteyn = '" + CStr(FRectDeleteYN) + "' "
			end if

		elseif (FRectSearchType = "notfinish") then
                ''��ó����ü
                AddSQL = AddSQL + " and A.currstate<'B007' and A.deleteyn='N' "
        elseif (FRectSearchType = "norefund") then
                'ȯ�ҹ�ó��
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A003' "
                AddSQL = AddSQL + " and A.deleteyn='N'"
        elseif (FRectSearchType = "norefundmile") then
                '���ϸ���ȯ�ҹ�ó��
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A003' "
                AddSQL = AddSQL + " and A.deleteyn='N'"
                AddSQL = AddSQL + " and R.returnmethod='R900'"
        elseif (FRectSearchType = "norefundetc") then
                '���ϸ���ȯ�ҹ�ó��
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A005' "
                AddSQL = AddSQL + " and A.deleteyn='N'"
                ''AddSQL = AddSQL + " and R.returnmethod='R050'"
        elseif (FRectSearchType = "cardnocheck") then
                'ī����ҹ�ó��
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A007' and A.deleteyn='N' "
        elseif (FRectSearchType = "beasongnocheck") then
                '������ǻ���/���
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd in ('A008','A006') and ((A.requireupche is Null) or (A.requireupche='N')) and deleteyn='N' "
        elseif (FRectSearchType = "upchemifinish") then
                '��ü��ó��
                AddSQL = AddSQL + " and A.currstate<'B006' and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "upchefinish") then
                '��üó���Ϸ�
                AddSQL = AddSQL + " and A.currstate='B006' and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "returnmifinish") then
                'ȸ����û��ó��
                AddSQL = AddSQL + " and A.currstate<'B002' and A.divcd ='A010' and A.deleteyn='N'  "
        elseif (FRectSearchType = "confirm") then
                'Ȯ�ο�û ��ó��
                AddSQL = AddSQL + " and A.currstate='B005' and A.deleteyn='N' "
        elseif (FRectSearchType = "cancelnofinish") then
                '�ֹ���� ��ó��
                AddSQL = AddSQL + " and A.divcd='A008'"
                AddSQL = AddSQL + " and A.currstate='B001' and A.deleteyn='N' "
                AddSQL = AddSQL + " and A.regdate>'2008-04-23'"
        end If



        sqlStr = sqlStr + AddSQL

		'rw sqlStr
        rsget.Open sqlStr, dbget, 1

        if  not rsget.EOF  then
            FTotalCount = rsget("cnt")
        else
            FTotalCount = 0
        end if
        rsget.close


        sqlStr = " select      Top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr + "     A.id, A.divcd, A.gubun01, A.gubun02, A.orderserial, A.customername, A.userid, A.finishuser, A.writeuser, A.title, A.currstate"
        sqlStr = sqlStr + "     ,A.regdate, A.finishdate,A.deleteyn "
        sqlStr = sqlStr + "     , A.requireupche, A.makerid, A.songjangdiv ,A.songjangno"
        sqlStr = sqlStr + "     ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + "     ,m.sitename, m.authcode"
        sqlStr = sqlStr + "     ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename, C4.comm_color as currstatecolor"
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " A"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_REFUND & " r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join " & TABLE_ORDERMASTER & " m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and m.sitename <> '" & EXCLUDE_SITENAME & "' "

        sqlStr = sqlStr + AddSQL

        sqlStr = sqlStr + " order by id desc "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

        redim preserve FItemList(FResultCount)
        if  not rsget.EOF  then
            i = 0
			rsget.absolutepage = FCurrPage
            do until rsget.eof
                set FItemList(i) = new CCSASMasterItem

                FItemList(i).Fid                = rsget("id")
                FItemList(i).Fdivcd             = rsget("divcd")
                FItemList(i).FdivcdName         = db2html(rsget("divcdname"))

                FItemList(i).Forderserial       = rsget("orderserial")
                FItemList(i).Fcustomername      = db2html(rsget("customername"))
                FItemList(i).Fuserid            = rsget("userid")
                FItemList(i).Fwriteuser         = rsget("writeuser")
                FItemList(i).Ffinishuser        = rsget("finishuser")
                FItemList(i).Ftitle             = db2html(rsget("title"))
                FItemList(i).Fcurrstate         = rsget("currstate")
                FItemList(i).Fcurrstatename     = rsget("currstatename")
                FItemList(i).FcurrstateColor    = rsget("currstatecolor")

                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Ffinishdate        = rsget("finishdate")

                FItemList(i).Fgubun01           = rsget("gubun01")
                FItemList(i).Fgubun02           = rsget("gubun02")

                FItemList(i).Fgubun01Name       = db2html(rsget("gubun01name"))
                FItemList(i).Fgubun02Name       = db2html(rsget("gubun02name"))

                FItemList(i).Fdeleteyn          = rsget("deleteyn")

                FItemList(i).Frefundrequire     = rsget("refundrequire")
                FItemList(i).Frefundresult      = rsget("refundresult")

                FItemList(i).Fsongjangdiv       = rsget("songjangdiv")
                FItemList(i).Fsongjangno        = rsget("songjangno")

                FItemList(i).Frequireupche      = rsget("requireupche")
                FItemList(i).Fmakerid           = rsget("makerid")

                FItemList(i).FExtsitename          = rsget("sitename")
                FItemList(i).Fauthcode          = rsget("authcode")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close
    end sub



    public Sub GetCSASTotalPrevCancelCount()
        dim i,sqlStr

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " "
        sqlStr = sqlStr + " where 1 = 1 "

        if (FRectOrderSerial <> "") then
                sqlStr = sqlStr + " and orderserial='" + CStr(FRectOrderSerial) + "' "
        end if

        sqlStr = sqlStr + " and deleteyn='N' and divcd in ('A003','A005','A007') "
        rsget.Open sqlStr, dbget, 1

        if  not rsget.EOF  then
                FResultCount = rsget("cnt")
        else
                FResultCount = 0
        end if
        rsget.close
    end sub

    public Sub GetOneCSASMaster()
        dim i,sqlStr

        sqlStr = " select top 1 A.*, IsNULL(J.add_upchejungsandeliverypay,0) as add_upchejungsandeliverypay, J.add_upchejungsancause "
        sqlStr = sqlStr + " ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult, IsNULL(refminusorderserial,'') as refminusorderserial"  ''refminusorderserial �߰� 2017/03/27
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename"
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " A "
        sqlStr = sqlStr + " Left join " & TABLE_UPCHE_ADD_JUNGSAN & " J"
        sqlStr = sqlStr + "  on A.id=J.asid"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_REFUND & " r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"

        sqlStr = sqlStr + " where id= " + CStr(FRectCsAsID) + " "

        if (FRectMakerID<>"") then   ''��ü ��ȸ��.
            sqlStr = sqlStr + " and A.makerid='"&FRectMakerID&"'"
        end if
        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

        if  not rsget.EOF  then
            set FOneItem = new CCSASMasterItem

            FOneItem.Fid                  = rsget("id")
            FOneItem.Fdivcd               = rsget("divcd")
            FOneItem.Fgubun01             = rsget("gubun01")
            FOneItem.Fgubun02             = rsget("gubun02")

            FOneItem.FdivcdName           = db2html(rsget("divcdname"))
            FOneItem.Fgubun01Name         = db2html(rsget("gubun01name"))
            FOneItem.Fgubun02Name         = db2html(rsget("gubun02name"))

            FOneItem.Forderserial         = rsget("orderserial")
            FOneItem.Fcustomername        = db2html(rsget("customername"))
            FOneItem.Fuserid              = rsget("userid")
            FOneItem.Fwriteuser           = rsget("writeuser")
            FOneItem.Ffinishuser          = rsget("finishuser")
            FOneItem.Ftitle               = db2html(rsget("title"))
            FOneItem.Fcontents_jupsu      = db2html(rsget("contents_jupsu"))
            FOneItem.Fcontents_finish     = db2html(rsget("contents_finish"))
            FOneItem.Fcurrstate           = rsget("currstate")
            FOneItem.FcurrstateName       = rsget("currstatename")
            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Ffinishdate          = rsget("finishdate")

            FOneItem.Fdeleteyn            = rsget("deleteyn")
            FOneItem.Fextsitename         = rsget("extsitename")

            FOneItem.Fopentitle           = db2html(rsget("opentitle"))
            FOneItem.Fopencontents        = db2html(rsget("opencontents"))


            FOneItem.Fsitegubun           = rsget("sitegubun")

            FOneItem.Fsongjangdiv         = rsget("songjangdiv")
            FOneItem.Fsongjangno          = rsget("songjangno")

            FOneItem.Frequireupche        = rsget("requireupche")
            FOneItem.Fmakerid             = rsget("makerid")

            FOneItem.Fadd_upchejungsandeliverypay = rsget("add_upchejungsandeliverypay")
            FOneItem.Fadd_upchejungsancause       = rsget("add_upchejungsancause")
            
            FOneItem.Frefminusorderserial 	= rsget("refminusorderserial")
            
'            FOneItem.Fbeasongdate         = rsget("beasongdate")
'            FOneItem.Frefundrequire       = rsget("refundrequire")
'            FOneItem.Frefundresult        = rsget("refundresult")

        end if
        rsget.close
    end sub

    public Sub GetOrderDetailByCsDetail()
        dim SqlStr, i

		sqlStr = "select d." & FIELD_DETAILIDX & " as orderdetailidx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost, d.buycash, d.reducedprice as discountAssingedCost"
		sqlStr = sqlStr + " ,d.mileage,d.cancelyn "
		sqlStr = sqlStr + " ,d.itemname, d.makerid, d.itemoptionname "
		sqlStr = sqlStr + " ,d.currstate as orderdetailcurrstate, d.upcheconfirmdate, d.songjangdiv, d.songjangno"
		sqlStr = sqlStr + " ,d.beasongdate, d.isupchebeasong, d.issailitem , d.cancelyn "
		sqlStr = sqlStr + " ,d.oitemdiv, d.odlvType, d." & FIELD_ITEMCOUPONIDX & " as itemcouponidx, d.bonuscouponidx"
		sqlStr = sqlStr + " ,c.id, c.masterid, IsNULL(c.regitemno,0) as regitemno, IsNULL(c.confirmitemno,0) as confirmitemno"
		sqlStr = sqlStr + " ,c.gubun01, c.gubun02, c.regdetailstate"
		sqlStr = sqlStr + " ,C2.comm_name as gubun01name, C3.comm_name as gubun02name"
		sqlStr = sqlStr + " ,i.smallimage "

		sqlStr = sqlStr + " ,IsNull((select top 1 ad.confirmitemno "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		[db_academy].[dbo].[tbl_academy_as_list] a "
		sqlStr = sqlStr + " 		join [db_academy].[dbo].[tbl_academy_as_detail] ad "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			a.id = ad.masterid "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and a.divcd = 'A004' "
		sqlStr = sqlStr + " 		and a.deleteyn = 'N' "
		sqlStr = sqlStr + " 		and a.orderserial = d.orderserial "
		sqlStr = sqlStr + " 		and ad.itemid = d.itemid "
		sqlStr = sqlStr + " 		and ad.itemoption = d.itemoption),0) as prevreturnno "

		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d "
		else
		    sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d "
		end if
		sqlStr = sqlStr + " left join " & TABLE_ITEM & " i on d.itemid=i.itemid"
		sqlStr = sqlStr + " left join " & TABLE_CSDETAIL & " c "
		sqlStr = sqlStr + " on c.masterid=" + CStr(FRectCsAsID) + ""
		sqlStr = sqlStr + " and c.orderdetailidx=d." & FIELD_DETAILIDX & " "
		sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C2"
        sqlStr = sqlStr + "  on c.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C3"
        sqlStr = sqlStr + "  on c.gubun02=C3.comm_cd"

		sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"

        ''sqlStr = sqlStr + " order by d.isupchebeasong, d.makerid, d.itemid, d.itemoption"
		sqlStr = sqlStr + " order by d.itemid, d.itemoption"
		''response.write sqlStr
		''response.end

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSASDetailItem

            ''tbl_as_detail's
            FItemList(i).Fid              = rsget("id")
            FItemList(i).Fmasterid        = rsget("masterid")
            FItemList(i).Fgubun01         = rsget("gubun01")
            FItemList(i).Fgubun02         = rsget("gubun02")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")

            ''tbl_order_detail's
            FItemList(i).Forderdetailidx  = rsget("orderdetailidx")
            FItemList(i).Forderserial     = rsget("orderserial")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
            FItemList(i).Fitemcost        = rsget("itemcost")
            FItemList(i).Fbuycash         = rsget("buycash")

			FItemList(i).Fitemno          = rsget("itemno")
			FItemList(i).Fprevreturnno    = rsget("prevreturnno")


            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
            FItemList(i).FCancelyn        = rsget("cancelyn")
            FItemList(i).ForderDetailcurrstate = rsget("orderdetailcurrstate")

            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")

            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")


            ''���� ����ϰų�, ���ϸ����� ��ǰ�� ���� �ȵǾ���.
''            if (rsget("oitemdiv")="82") or (rsget("itemcouponidx")<>0) or (rsget("issailitem")="Y") then
''                FItemList(i).FAllAtDiscountedPrice = 0
''            else
''                FItemList(i).FAllAtDiscountedPrice = round(((1-0.94) * FItemList(i).Fitemcost / 100) * 100 )
''            end if


            ''tbl_item's
            FItemList(i).FSmallImage  	  = webImgUrl + DIRECTORY_IMAGE_SMALL + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget("gubun01name")
            FItemList(i).Fgubun02name   = rsget("gubun02name")
			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

    public Sub GetCsDetailList()
        dim SqlStr, i

		sqlStr = "select c.*"
		sqlStr = sqlStr + " ,d.currstate as orderdetailcurrstate"
		sqlStr = sqlStr + " ,d.reducedprice as discountAssingedCost, d.oitemdiv, d.odlvType, d.issailitem, d." & FIELD_ITEMCOUPONIDX & " as itemcouponidx, d.bonuscouponidx"
		sqlStr = sqlStr + " ,IsNULL(d.itemcost,0) as OrderItemcost"
		sqlStr = sqlStr + " ,C2.comm_name as gubun01name, C3.comm_name as gubun02name"
		sqlStr = sqlStr + " ,i.smallimage "

		sqlStr = sqlStr + " from " & TABLE_CSDETAIL & " c "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " left join [db_log].[dbo].tbl_old_order_detail_2003 d"
		    sqlStr = sqlStr + "  on c.orderdetailidx=d." & FIELD_DETAILIDX & ""
		else
		    sqlStr = sqlStr + " left join " & TABLE_ORDERDETAIL & " d"
		    sqlStr = sqlStr + "  on c.orderdetailidx=d." & FIELD_DETAILIDX & ""
		end if

		sqlStr = sqlStr + " left join " & TABLE_ITEM & " i "
		sqlStr = sqlStr + "  on c.itemid=i.itemid"
		sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C2"
        sqlStr = sqlStr + "  on c.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C3"
        sqlStr = sqlStr + "  on c.gubun02=C3.comm_cd"

		sqlStr = sqlStr + " where c.masterid=" + CStr(FRectCsAsID) + ""
        sqlStr = sqlStr + " order by c.isupchebeasong, c.makerid, c.itemid, c.itemoption"

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSASDetailItem

            FItemList(i).Fid              = rsget("id")
            FItemList(i).Fmasterid        = rsget("masterid")
            FItemList(i).Fgubun01         = rsget("gubun01")
            FItemList(i).Fgubun02         = rsget("gubun02")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")

            FItemList(i).Fregdetailstate  = rsget("regdetailstate")   ''���� ��� ���� ����
            FItemList(i).Forderdetailidx  = rsget("orderdetailidx")
            FItemList(i).Forderserial     = rsget("orderserial")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
            FItemList(i).Fitemcost        = rsget("itemcost")
            FItemList(i).Fbuycash         = rsget("buycash")
            FItemList(i).Fitemno          = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")

            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")

            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")


            FItemList(i).Forderdetailcurrstate  = rsget("orderdetailcurrstate")

            FItemList(i).FSmallImage  	  = webImgUrl + DIRECTORY_IMAGE_SMALL + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget("gubun01name")
            FItemList(i).Fgubun02name   = rsget("gubun02name")

            if (FItemList(i).Fitemcost=0) then
                FItemList(i).Fitemcost = rsget("OrderItemcost")
            end if

			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

    public Sub GetCSASTotalCount()
        dim i,sqlStr

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " "
        sqlStr = sqlStr + " where 1 = 1 "

        if (FRectNotCsID<> "") then
            sqlStr = sqlStr + " and id<>'" + CStr(FRectNotCsID) + "' "
        end if

        if (FRectUserID <> "") then
                sqlStr = sqlStr + " and userid='" + CStr(FRectUserID) + "' "
        end if

        if (FRectUserName <> "") then
                sqlStr = sqlStr + " and customername='" + CStr(FRectUserName) + "' "
        end if

        if (FRectOrderSerial <> "") then
                sqlStr = sqlStr + " and orderserial='" + CStr(FRectOrderSerial) + "' "
        end if

        if (FRectStartDate <> "") then
                sqlStr = sqlStr + " and regdate>='" + CStr(FRectStartDate) + "' "
        end if

        if (FRectEndDate <> "") then
                sqlStr = sqlStr + " and regdate <'" + CStr(FRectEndDate) + "' "
        end if

        if (FRectSearchType = "norefund") then
                'ȯ�ҹ�ó��
                sqlStr = sqlStr + " and currstate<7 and divcd in ('3','5') "
        elseif (FRectSearchType = "cardnocheck") then
                'ī����ҹ�ó��
                sqlStr = sqlStr + " and currstate<7 and divcd='7' "
        elseif (FRectSearchType = "beasongnocheck") then
                '������ǻ���/���
                sqlStr = sqlStr + " and currstate<7 and divcd in ('8','6') and ((requireupche is Null) or (requireupche='N')) "
        elseif (FRectSearchType = "upchemifinish") then
                '��ü��ó��
                sqlStr = sqlStr + " and currstate<6 and requireupche='Y' and deleteyn='N' "
        elseif (FRectSearchType = "upchefinish") then
                '��üó���Ϸ�
                sqlStr = sqlStr + " and currstate=6 and requireupche='Y' and deleteyn='N' "
        elseif (FRectSearchType = "returnmifinish") then
                'ȸ����û��ó��
                sqlStr = sqlStr + " and currstate<2 and divcd ='10' "
        end if

        rsget.Open sqlStr, dbget, 1

        if  not rsget.EOF  then
            FResultCount = rsget("cnt")
        else
            FResultCount = 0
        end if
        rsget.close
    end sub

    public Sub GetOneCsDeliveryItem()
        dim i,sqlStr

        sqlStr = " select top 1 A.* "
        sqlStr = sqlStr + " from " & TABLE_CS_DELIVERY & " A "
        sqlStr = sqlStr + " where asid= " + CStr(FRectCsAsID) + " "

        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

        if  not rsget.EOF  then
            set FOneItem = new CCSDeliveryItem
            FOneItem.Fasid              = rsget("asid")
            FOneItem.Freqname           = db2html(rsget("reqname"))
            FOneItem.Freqphone          = rsget("reqphone")
            FOneItem.Freqhp             = rsget("reqhp")
            FOneItem.Freqzipcode        = rsget("reqzipcode")
            FOneItem.Freqzipaddr        = rsget("reqzipaddr")
            FOneItem.Freqetcaddr        = db2html(rsget("reqetcaddr"))
            FOneItem.Freqetcstr          = db2html(rsget("reqetcstr"))
            FOneItem.Fsongjangdiv       = rsget("songjangdiv")
            FOneItem.Fsongjangno        = rsget("songjangno")
            FOneItem.Fregdate           = rsget("regdate")
            FOneItem.Fsenddate          = rsget("senddate")

        end if
        rsget.close

    end Sub

    public Sub GetOneCsDeliveryItemFromDefaultOrder()
        dim i,sqlStr

        sqlStr = " select m.reqname, m.reqphone, m.reqhp, m.reqzipcode, m.reqzipaddr, m.reqaddress"
        sqlStr = sqlStr + " from db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr + "     Join " & TABLE_CSMASTER & " a"
        sqlStr = sqlStr + "     on m.orderserial=a.orderserial"
        sqlStr = sqlStr + "     and a.id=" + CStr(FRectCsAsID) + " "

        rsget.Open sqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        if  not rsget.EOF  then
            set FOneItem = new CCSDeliveryItem
            FOneItem.Fasid              = FRectCsAsID
            FOneItem.Freqname           = db2html(rsget("reqname"))
            FOneItem.Freqphone          = rsget("reqphone")
            FOneItem.Freqhp             = rsget("reqhp")
            FOneItem.Freqzipcode        = rsget("reqzipcode")
            FOneItem.Freqzipaddr        = rsget("reqzipaddr")
            FOneItem.Freqetcaddr        = db2html(rsget("reqaddress"))
            'FOneItem.Freqetcstr          = db2html(rsget("reqetcstr"))
            'FOneItem.Fsongjangdiv       = rsget("songjangdiv")
            'FOneItem.Fsongjangno        = rsget("songjangno")
            'FOneItem.Fregdate           = rsget("regdate")
            'FOneItem.Fsenddate          = rsget("senddate")

        end if
        rsget.close

        if (FResultCount<1) then
            sqlStr = " select m.reqname, m.reqphone, m.reqhp, m.reqzipcode, m.reqzipaddr, m.reqaddress"
            sqlStr = sqlStr + " from db_log.dbo.tbl_old_order_master_2003 m"
            sqlStr = sqlStr + "     Join " & TABLE_CSMASTER & " a"
            sqlStr = sqlStr + "     on m.orderserial=a.orderserial"
            sqlStr = sqlStr + "     and a.id=" + CStr(FRectCsAsID) + " "

            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            if  not rsget.EOF  then
                set FOneItem = new CCSDeliveryItem
                FOneItem.Fasid              = FRectCsAsID
                FOneItem.Freqname           = db2html(rsget("reqname"))
                FOneItem.Freqphone          = rsget("reqphone")
                FOneItem.Freqhp             = rsget("reqhp")
                FOneItem.Freqzipcode        = rsget("reqzipcode")
                FOneItem.Freqzipaddr        = rsget("reqzipaddr")
                FOneItem.Freqetcaddr        = db2html(rsget("reqaddress"))
                'FOneItem.Freqetcstr          = db2html(rsget("reqetcstr"))
                'FOneItem.Fsongjangdiv       = rsget("songjangdiv")
                'FOneItem.Fsongjangno        = rsget("songjangno")
                'FOneItem.Fregdate           = rsget("regdate")
                'FOneItem.Fsenddate          = rsget("senddate")

            end if
            rsget.close
        end if
    end Sub

    public sub GetOneCsConfirmItem()
        dim sqlStr, i
        sqlStr = " select top 1 * from " & TABLE_CS_CONFIRM & ""
        sqlStr = sqlStr + " where asid=" + CStr(FRectCsAsID)



        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

        if  not rsget.EOF  then
            set FOneItem = new CCsConfirmItem

            FOneItem.Fasid                  = rsget("asid")
            FOneItem.Fconfirmregmsg         = db2html(rsget("confirmregmsg"))
            FOneItem.Fconfirmreguserid      = rsget("confirmreguserid")
            FOneItem.Fconfirmregdate        = rsget("confirmregdate")
            FOneItem.Fconfirmfinishmsg      = db2html(rsget("confirmfinishmsg"))
            FOneItem.Fconfirmfinishuserid   = rsget("confirmfinishuserid")
            FOneItem.Fconfirmfinishdate     = rsget("confirmfinishdate")

        end if
        rsget.close

    end sub

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
            HasPreScroll = StarScrollPage > 1
    end Function

    public Function HasNextScroll()
            HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
    end Function

    public Function StarScrollPage()
            StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
    end Function

end Class



%>
