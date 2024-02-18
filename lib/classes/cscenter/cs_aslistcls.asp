<%
function Fn_getRecentUpcheCSExsists(iorderserial,imakerid)
    Dim sqlStr, Mxregdt, MxFinishDt, dataExists
    Fn_getRecentUpcheCSExsists = false
    dataExists = false

    sqlStr = "select MAX(regdate) as Mxregdt, MAX(finishdate) as MxFinishDt" & vbCRLF
    sqlStr = sqlStr & " from db_cs.dbo.tbl_new_As_list" & vbCRLF
    sqlStr = sqlStr & " where orderserial='"&iorderserial&"'" & vbCRLF
    sqlStr = sqlStr & " and makerid='"&imakerid&"'" & vbCRLF
    sqlStr = sqlStr & " and deleteyn='N'" & vbCRLF
    sqlStr = sqlStr & " and requireupche='Y'" & vbCRLF
    sqlStr = sqlStr & " and divcd in ('A000','A012','A004')" & vbCRLF

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        Mxregdt    = rsget("Mxregdt")
        MxFinishDt = rsget("MxFinishDt")

    end if
    rsget.Close

    if (NOT isNULL(Mxregdt)) then dataExists = true

    if (NOT dataExists) then Exit function

    if (isNULL(MxFinishDt)) and (NOT isNull(Mxregdt)) then
        if (datediff("d",Mxregdt,now())<45) then  ''45�� �̳���ϵ� CS
            Fn_getRecentUpcheCSExsists = true
            exit function
        end if
    end if


    if (datediff("d",MxFinishDt,now())<10) then  ''�Ϸ���� 10�� �̳�CS
        Fn_getRecentUpcheCSExsists = true
    end if

end function

function GetCSCommName(groupCode, divcd)
	dim tmp_str,sqlStr

	sqlStr = " select top 1 comm_cd,comm_name "
	sqlStr = sqlStr + " from  "
	sqlStr = sqlStr + " [db_cs].[dbo].tbl_cs_comm_code "
	sqlStr = sqlStr + " where comm_group='" + groupCode + "' "
	sqlStr = sqlStr + " and comm_cd='" + CStr(divcd) + "' "
	sqlStr = sqlStr + " and comm_isDel='N' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	tmp_str = ""
	if  not rsget.EOF  then
		tmp_str = db2html(rsget("comm_name"))
	end if
	rsget.close

	GetCSCommName = tmp_str
End function

'�������̻�����
function IsCSRegularUser(userid)
	dim tmp_str,sqlStr

	IsCSRegularUser = False

	sqlStr = " select top 1 userid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_partner.dbo.tbl_user_tenbyten "
	sqlStr = sqlStr + " where posit_sn <= 8 and isusing = 1 and userid = '" + CStr(userid) + "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if  not rsget.EOF  then
		IsCSRegularUser = True
	end if
	rsget.close
End function

function drawSelectBoxCSCommCombo(selectBoxName,selectedId,groupCode,onChangefunction)
   dim tmp_str,sqlStr
   %>
     <select class="select" name="<%=selectBoxName%>" <%= onChangefunction %> >
     <option value='' <%if selectedId="" then response.write " selected" %> >����</option>
   <%
       sqlStr = " select comm_cd,comm_name "
       sqlStr = sqlStr + " from  "
       sqlStr = sqlStr + " [db_cs].[dbo].tbl_cs_comm_code "
       sqlStr = sqlStr + " where comm_group='" + groupCode + "' "
       sqlStr = sqlStr + " and comm_isDel='N' "
       sqlStr = sqlStr + " order by comm_cd "

       rsget.CursorLocation = adUseClient
	   rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
        selectStr = ""
        if (selectedId="R100") then selectStr="selected"
        BufStr = BufStr + "<option value='R100' " + selectStr + ">�ſ�ī�� ���</option>"

        '''if application("Svr_Info") = "Dev" then
			if (orgPaymethod = "100") then
				selectStr = ""
				if (selectedId="R120") then selectStr="selected"
		        BufStr = BufStr + "<option value='R120' " + selectStr + ">�ſ�ī�� �κ����</option>"
		    end if
        '''end if
    elseif (orgPaymethod="550")  then
        if (selectedId="R550") then selectStr="selected"
        BufStr = BufStr + "<option value='R550' " + selectStr + ">������ ���</option>"
    elseif (orgPaymethod="560")  then
        if (selectedId="R560") then selectStr="selected"
        BufStr = BufStr + "<option value='R560' " + selectStr + ">����Ƽ�� ���</option>"
    elseif (orgPaymethod="20")  then
        if (selectedId="R020") then selectStr="selected"
        BufStr = BufStr + "<option value='R020' " + selectStr + ">�ǽð���ü ���</option>"

        ''if (application("Svr_Info") = "Dev") then  '' ���̹����� ���� �� �׽�Ʈ
        if (iPgGubun="NP") then  ''����������.
            selectStr = ""
            if (selectedId="R022") then selectStr="selected"
            BufStr = BufStr + "<option value='R022' " + selectStr + ">�ǽð���ü �κ����(���̹�����)</option>"
        end if
        ''end if

        if (iPgGubun="KK") then  ''īī������ �Ӵϰ����� ���� (20181210 ����).
            selectStr = ""
            if (selectedId="R022") then selectStr="selected"
            BufStr = BufStr + "<option value='R022' " + selectStr + ">�ǽð���ü �κ����(īī������)</option>"
        end if

        if (iPgGubun="TS") then  ''����������.
            selectStr = ""
            if (selectedId="R022") then selectStr="selected"
            BufStr = BufStr + "<option value='R022' " + selectStr + ">�ǽð���ü �κ����(�佺)</option>"
        end if

        if (iPgGubun="CH") then  ''���������� �������� �Ӵϰ����� ���� (20200423 ����).
            selectStr = ""
            if (selectedId="R022") then selectStr="selected"
            BufStr = BufStr + "<option value='R022' " + selectStr + ">�ǽð���ü �κ����(��������)</option>"
        end if

        if (iPgGubun="PY") then  ''���������� ������
            selectStr = ""
            if (selectedId="R022") then selectStr="selected"
            BufStr = BufStr + "<option value='R022' " + selectStr + ">�ǽð���ü �κ����(������)</option>"
        end if

        if (iPgGubun="") then  ''���������� �̴Ͻý�(?)
            selectStr = ""
            if (selectedId="R022") then selectStr="selected"
            BufStr = BufStr + "<option value='R022' " + selectStr + ">�ǽð���ü �κ����</option>"
        end if

    elseif (orgPaymethod="400")  then
        if (selectedId="R400") then selectStr="selected"
        BufStr = BufStr + "<option value='R400' " + selectStr + ">�޴������� ���</option>"

        if application("Svr_Info") = "Dev" or C_ADMIN_AUTH then
			if (orgPaymethod = "400") then
				selectStr = ""
				if (selectedId="R420") then selectStr="selected"
		        BufStr = BufStr + "<option value='R420' " + selectStr + ">�޴������� �κ����</option>"
		    end if
        end if
    elseif (orgPaymethod="80")  then
        if (selectedId="R080") then selectStr="selected"
        BufStr = BufStr + "<option value='R080' " + selectStr + ">All@ī�� ���</option>"
    elseif (orgPaymethod="50") then
        if (selectedId="R050") then selectStr="selected"
        BufStr = BufStr + "<option value='R050' " + selectStr + ">���������� ���</option>"
    elseif (orgPaymethod="150")  then
        if (selectedId="R150") then selectStr="selected"
        BufStr = BufStr + "<option value='R150' " + selectStr + ">�̴Ϸ�Ż ���</option>"
    end if

    if ((IsStatusRegister) and (divcd="A003") and Not(C_CSPowerUser or C_ADMIN_AUTH)) then
        ''�븮�� ���� ���� ȯ�ҿ�û �Ұ� >> ���
        ''��ũ��Ʈ���� 1���� �ʰ� ����(skyer9)
        selectStr = ""
        if (selectedId="R007") then selectStr="selected"
        BufStr = BufStr + "<option value='R007' " + selectStr + ">������ ȯ��</option>"
    ELSE
        selectStr = ""
        if (selectedId="R007") then selectStr="selected"
        BufStr = BufStr + "<option value='R007' " + selectStr + ">������ ȯ��</option>"
    END IF

	selectStr = ""
    if (selectedId="R910") then selectStr="selected"
    BufStr = BufStr + "<option value='R910' " + selectStr + ">��ġ�� ȯ��</option>"

    selectStr = ""
    if (selectedId="R900") then selectStr="selected"
	if ((selectedId="R900") or Not ((divcd = "A008") or (divcd = "A004") or (divcd = "A010") or (divcd = "A003"))) then
		BufStr = BufStr + "<option value='R900' " + selectStr + ">���ϸ��� ȯ��</option>"
	end if

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

''��Ÿȸ�� ���μ���
public function fnIsServiceRecvProcess(idivcd)
    fnIsServiceRecvProcess = (idivcd = "A200")
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
    public Forderitemno
    public Fisupchebeasong
    public Fcancelyn

    public Foitemdiv
    public FodlvType
    public Fissailitem
    public Fitemcouponidx
    public Fbonuscouponidx

    public ForderDetailcurrstate
    public FdiscountAssingedCost    '' �ֹ��� ���εȰ��� ( ALL@ / %���α� �ݿ�)

	public FreducedPrice
    public Forgitemcost					'�Һ��ڰ�
    public FitemcostCouponNotApplied	'�ǸŰ�(���ΰ�)
    public FplusSaleDiscount			'�÷����������ξ�
    public FspecialshopDiscount			'��������ξ�
	public FetcDiscount					'��Ÿ���ξ�

    public Forgprice					'����Һ��ڰ�(+�ɼǰ�)

	public Fprevcsreturnfinishno		'���� CS��ǰ����(�����̻�)

	public Freforderdetailidx

	Public Fsongjangdiv
	Public Fsongjangno

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
		'// 2018-04-19, �ÿ����� ���̻� ����, etcDiscount �� ����
        getAllAtDiscountedPrice = 0
        ''���� ��ǰ���� ���εǴ°�� �߰����ξ���.
        ''���ϸ���SHOP ��ǰ �߰� ���� ����.
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
'    	    if (IsNULL(Fbonuscouponidx) or (Fbonuscouponidx=0)) and (Fitemcost>FdiscountAssingedCost) then
'    	            getAllAtDiscountedPrice = Fitemcost-FdiscountAssingedCost
'    	    else
'    	        getAllAtDiscountedPrice = 0
'    	    end if
'    	end if
    end function

    '' %���α� ���αݾ� or ī�� ���αݾ�
    public function getPercentBonusCouponDiscountedPrice()
        getPercentBonusCouponDiscountedPrice = 0
'        if (Fitemcost>FdiscountAssingedCost) then
'                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
'        end if

		if (Fitemid = 0) and (Fitemcost > FdiscountAssingedCost) and not IsNull(Fbonuscouponidx) then
			'// ��ۺ� ����
			getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
        ''elseif (FdiscountAssingedCost=0) then
	        ''�������
	    ''    ''getPercentBonusCouponDiscountedPrice = Fitemcost*
		else
			'// ���� �������� ����(2014-06-23, skyer9)
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

        elseif (fnIsServiceDeliverProcess(idivcd)) or (fnIsServiceRecvProcess(idivcd)) then
            '���� - �׻� ���� ���� ����
            if (idivcd = "A002") or (idivcd = "A200") then
            	IsItemNoEditEnabled=true

            elseif (ForderDetailcurrstate>=7) then
            	IsItemNoEditEnabled=true

            end if
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
        elseif (idivcd = "A002") or (idivcd = "A200") then
        	'���� - �׻� üũ����
            if Fitemid=0 then
                IsCheckAvailItem=false
            else
                IsCheckAvailItem=true
            end if
        elseif (idivcd = "A001") then
            ''����
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

	'==========================================================================
	'��ǰ���� ���� �ֹ����� üũ
    public function IsSaleDiscountItem()
        IsSaleDiscountItem = (GetSaleDiscountPrice() > 0)
    end function

	'��ǰ���� ���� �ֹ����� üũ
    public function IsItemCouponDiscountItem()
        IsItemCouponDiscountItem = false
        if (Not IsNull(Fitemcouponidx) and (Fitemcouponidx<>0)) then
            IsItemCouponDiscountItem = true
        end if
    end function

    '���ʽ����� ���� �ֹ����� üũ
    public function IsBonusCouponDiscountItem()
        IsBonusCouponDiscountItem = false
        if (Not IsNull(Fbonuscouponidx) and (Fbonuscouponidx<>0) and (GetItemCouponPrice > GetBonusCouponPrice))  then
            IsBonusCouponDiscountItem = true
        end if
    end function

	'��Ÿ���� ���� �ֹ����� üũ
    public function IsEtcDiscountItem()
        IsEtcDiscountItem = false
        if (Not IsNull(Fbonuscouponidx) and (Fbonuscouponidx<>0) and (GetBonusCouponPrice > GetEtcDiscountPrice))  then
            IsEtcDiscountItem = true
        end if
    end function

    '��������� ���� �ֹ����� üũ
    public function IsSpecialShopDiscountItem()
        if (FitemcostCouponNotApplied = 0) then
        	'���ŵ���Ÿ
        	if (Not IsItemCouponDiscountItem) and (Not IsBonusCouponDiscountItem) and (Fissailitem = "N") then
        		'TODO : �Һ��ڰ�����, �ɼǰ������� �ִ°�� ����Ȯ�� ���� �ȴ�.
        		GetItemCouponDiscountPrice = (Forgprice - Fitemcost) = 0
        		exit function
        	end if

        	GetItemCouponDiscountPrice = false
        	exit function
        end if

		if (FspecialshopDiscount > 0) then
			IsSpecialShopDiscountItem = true
		else
			IsSpecialShopDiscountItem = false
		end if
    end function

	'��ǰ�������ξ�
    public function GetItemCouponDiscountPrice()
        if (FitemcostCouponNotApplied = 0) then
        	'���ŵ���Ÿ
        	if (IsItemCouponDiscountItem = true) and (Not IsBonusCouponDiscountItem) and (Fissailitem = "N") then
        		'TODO : �Һ��ڰ�����, �ɼǰ�����, ����������� �ִ°�� ����Ȯ�� ���� �ȴ�.
        		GetItemCouponDiscountPrice = Forgprice - Fitemcost
        		exit function
        	end if

        	GetItemCouponDiscountPrice = 0
        	exit function
        end if

        GetItemCouponDiscountPrice = FitemcostCouponNotApplied - Fitemcost
    end function

	'���ʽ��������ξ�
    public function GetBonusCouponDiscountPrice()
        GetBonusCouponDiscountPrice = GetItemCouponPrice - GetBonusCouponPrice
    end function

	'��Ÿ�������ξ�
	public function GetEtcDiscountDiscountPrice()
        GetEtcDiscountDiscountPrice = GetBonusCouponPrice - GetEtcDiscountPrice
    end function

	'��ǰ���ξ�
    public function GetSaleDiscountPrice()
        if (FitemcostCouponNotApplied = 0) then
        	'���ŵ���Ÿ
        	if (Not IsBonusCouponDiscountItem) and (Not IsItemCouponDiscountItem) and (Fissailitem = "Y") then
        		'TODO : �Һ��ڰ�����, �ɼǰ�����, ����������� �ִ°�� ����Ȯ�� ���� �ȴ�.
        		GetSaleDiscountPrice = Forgprice - Fitemcost
        		exit function
        	end if

        	GetSaleDiscountPrice = 0
        	exit function
        end if

        GetSaleDiscountPrice = (Forgitemcost - (FitemcostCouponNotApplied + FplusSaleDiscount + FspecialshopDiscount))
    end function

    public function IsOldJumun()
    	'2011�� 4�� 1�� ���� �ֹ� �Ǵ� �� �ֹ��� ���� ���̳ʽ��ֹ�
    	IsOldJumun = (Forgitemcost = 0)
    end function

	public function GetOrgItemCostColor()
		if IsOldJumun then
			GetOrgItemCostColor = "gray"
		else
			GetOrgItemCostColor = "black"
		end if
	end function

	public function GetOrgItemCostPrice()
		if IsOldJumun then
			GetOrgItemCostPrice = Forgprice
		else
			GetOrgItemCostPrice = Forgitemcost
		end if
	end function

	public function GetSaleColor()
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				GetSaleColor = "red"
			else
				GetSaleColor = "black"
			end if
		else
			if (Forgitemcost <> FitemcostCouponNotApplied) then
				GetSaleColor = "red"
			else
				GetSaleColor = "black"
			end if
		end if
	end function

	public function GetSalePrice()
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				GetSalePrice = Fitemcost
			else
				GetSalePrice = Forgprice
			end if
		else
			GetSalePrice = FitemcostCouponNotApplied
		end if
	end function

	public function GetSaleText()
		dim result

		result = ""
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				if (Fissailitem = "Y") then
					if (Forgprice <= Fitemcost) then
						result = result + "���λ�ǰ + �Һ��ڰ� ����" + vbCrLf
					else
						result = result + "���λ�ǰ" + vbCrLf
					end if
				end if
				if (Fissailitem = "P") then
					result = result + "�÷�������" + vbCrLf
				end if
				if ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
					result = result + "��������� �Ǵ� �Һ��ڰ�/�ɼǰ� ����" + vbCrLf
				end if
			else
				result = "���󰡰�"
			end if
		else
			if (Forgitemcost <> FitemcostCouponNotApplied) then
				if (Fissailitem = "Y") then
					result = result + "���λ�ǰ : " + CStr(GetSaleDiscountPrice) + "��" + vbCrLf
				end if
				if (FplusSaleDiscount > 0) then
					result = result + "�÷������� : " + CStr(FplusSaleDiscount) + "��" + vbCrLf
				end if
				if (FspecialshopDiscount > 0) then
					result = result + "���ȸ������ : " + CStr(FspecialshopDiscount) + "��" + vbCrLf
				end if
			else
				result = "���󰡰�"
			end if
		end if

		GetSaleText = result
	end function

	public function GetItemCouponColor()
		if (IsItemCouponDiscountItem = true) then
			GetItemCouponColor = "green"
		else
			GetItemCouponColor = "black"
		end if
	end function

	public function GetItemCouponPrice()
		GetItemCouponPrice = Fitemcost
	end function

	public function GetItemCouponText()
		dim result

		result = ""
		if IsOldJumun then
			if (IsItemCouponDiscountItem = true) then
				if (GetSalePrice <> GetItemCouponPrice) then
					result = result + "��ǰ���������ǰ" + vbCrLf
				else
					result = result + "��ۺ����������ǰ" + vbCrLf
				end if
			else
				result = "���󰡰�"
			end if
		else
			if (IsItemCouponDiscountItem = true) then
				if (GetItemCouponDiscountPrice = 0) then
					result = result + "��ۺ����������ǰ" + vbCrLf
				else
					result = result + "��ǰ���� : " + CStr(GetItemCouponDiscountPrice) + "��" + vbCrLf
				end if
			else
				result = "���󰡰�"
			end if
		end if

		GetItemCouponText = result
	end function

	public function GetBonusCouponColor()
		if (IsBonusCouponDiscountItem = true) then
			GetBonusCouponColor = "purple"
		else
			GetBonusCouponColor = "black"
		end if
	end function

	public function GetBonusCouponPrice()
		if (FreducedPrice = "") then
			FreducedPrice = FdiscountAssingedCost
		end if
		if (FetcDiscount = "") then
			FetcDiscount = 0
		end if

		GetBonusCouponPrice = (FreducedPrice + FetcDiscount)
	end function

	public function GetBonusCouponText()
		dim result

		result = ""
		if IsOldJumun then
			if (IsBonusCouponDiscountItem = true) then
				result = result + "���ʽ�����" + vbCrLf
			else
				result = "���󰡰�"
			end if
		else
			if (IsBonusCouponDiscountItem = true) then
				result = result + "���ʽ����� : " + CStr(GetBonusCouponDiscountPrice) + "��" + vbCrLf
			else
				result = "���󰡰�"
			end if
		end if

		GetBonusCouponText = result
	end function

	public function GetEtcDiscountColor()
		if (IsEtcDiscountItem = true) then
			GetEtcDiscountColor = "red"
		else
			GetEtcDiscountColor = "black"
		end if
	end function

	public function GetEtcDiscountPrice()
		GetEtcDiscountPrice = FreducedPrice
	end function

	public function GetEtcDiscountText()
		dim result

		result = ""
		if IsOldJumun then
			if (IsEtcDiscountItem = true) then
				result = result + "��Ÿ����" + vbCrLf
			else
				result = "���󰡰�"
			end if
		else
			if (IsEtcDiscountItem = true) then
				result = result + "��Ÿ���� : " + CStr(GetEtcDiscountDiscountPrice) + "��" + vbCrLf
			else
				result = "���󰡰�"
			end if
		end if

		GetEtcDiscountText = result
	end function

	'==========================================================================
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

    public Forggiftcardsum      ''�� �ֹ� ��ǰ��
    public Forgdepositsum       ''�� �ֹ� ��ġ��

    public Frefundgiftcardsum   ''���  ��ǰ��
    public Frefunddepositsum    ''���  ��ġ��

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

	public Forgpercentcouponsum 		''�� �ֹ� ����������
	public Frefundpercentcouponsum		''��� ����������
	public Forgfixedcouponsum			''�� �ֹ� �����������
	public Frefundfixedcouponsum		''��� �����������

    public FpaygateresultTid
    public FpaygateresultMsg

    public FreturnmethodName    ''ȯ�ҹ�ĸ�

    public rebankCode

    public Fupfiledate          ''ȯ������ �ۼ���

	public Fcopycouponinfo		'' ���ʽ����� ��߱�
    public fcopyitemcouponinfo  ' ��ǰ���� ��߱�

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

''���� ó���� ���
Class CCSActUserHistoryItem
    public Fwriteuser
    public Ffinishuser
    public Fcurrstate
    public Ffinishdate
    public Fregdate

	Public function GetCurrStateName()
        if (Fcurrstate="B001") then
			GetCurrStateName = "����"
		elseif (Fcurrstate="B004") then
			GetCurrStateName = "������Է�"
		elseif (Fcurrstate="B005") then
			GetCurrStateName = "��üȮ�ο�û"
		elseif (Fcurrstate="B006") then
			GetCurrStateName = "��üó���Ϸ�"
		elseif (Fcurrstate="B007") then
			GetCurrStateName = "�Ϸ�"
		else
			GetCurrStateName = Fcurrstate
		end if
	end Function

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

    public FcsName
    public FcsPhone
    public Fcshp
    public FcsEmail
    public Fgroupid

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
    public flastInfoChgDT

    public sub GetReturnAddress()
        dim sqlStr
        sqlStr = " select company_name, deliver_phone, deliver_hp, return_zipcode, return_address, return_address2"
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner"
        sqlStr = sqlStr + " where id='" + FRectMakerid + "'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
        sqlStr = " select" & vbcrlf
        sqlStr = sqlStr & " id as brandid, company_name as brandname, socname_kor as streetname_kor, socname as streetname_eng" & vbcrlf
        sqlStr = sqlStr & " , return_zipcode, return_address, return_address2, deliver_phone, deliver_hp, deliver_name, deliver_email" & vbcrlf
        sqlStr = sqlStr & " , defaultsongjangdiv, p.lastInfoChgDT" & vbcrlf
        sqlStr = sqlStr & " from [db_partner].[dbo].tbl_partner p, [db_user].[dbo].tbl_user_c c" & vbcrlf
        sqlStr = sqlStr & " where 1 = 1" & vbcrlf
        sqlStr = sqlStr & " and p.id = c.userid" & vbcrlf
        sqlStr = sqlStr & " and p.id='" + FRectMakerid + "'" & vbcrlf

        'response.write sqlStr & "<Br>"
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
            flastInfoChgDT     = rsget("lastInfoChgDT")

        end if
        rsget.Close
    end sub

    public sub GetReturnAddressList()
        dim sqlStr, i

		sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_partner].[dbo].tbl_partner p "
        sqlStr = sqlStr + " 	join [db_user].[dbo].tbl_user_c c "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		p.id = c.userid "
        sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_cs_brand_memo m "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		p.id = m.brandid "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " and p.groupid ='" + FRectGroupCode + "'"

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " id as brandid, company_name as brandname, socname_kor as streetname_kor, socname as streetname_eng, return_zipcode, return_address, return_address2, deliver_phone, deliver_hp, deliver_name, deliver_email, defaultsongjangdiv, cs_name, cs_phone, cs_hp, cs_email, groupid "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_partner].[dbo].tbl_partner p "
        sqlStr = sqlStr + " 	join [db_user].[dbo].tbl_user_c c "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		p.id = c.userid "
        sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_cs_brand_memo m "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		p.id = m.brandid "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " and p.groupid ='" + FRectGroupCode + "'"
        sqlStr = sqlStr + " order by id "
		''response.write sqlStr

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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

				FItemList(i).FcsName       	= db2html(rsget("cs_name"))
				FItemList(i).FcsPhone       = db2html(rsget("cs_phone"))
				FItemList(i).Fcshp       	= db2html(rsget("cs_hp"))
				FItemList(i).FcsEmail       = db2html(rsget("cs_email"))
				FItemList(i).Fgroupid		= rsget("groupid")

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

	public Fbeasongneedday
	public Fbeasong_comment
	public Fbeasong_modifyday

	public Fbeasong_reguserid

	public Freturn_comment
	public Freturn_modifyday
	public Freturn_reguserid

	public FcsName
	public FcsPhone
	public Fcshp
	public FcsEmail
	public FcsModifyDay
	public FcsReguserID

	public Flunch_start
	public Flunch_end
	public Fvacation_div
    public Fcustomer_return_deny

    public FRectMakerid

    public sub GetBrandMemo()
        dim sqlStr

        sqlStr = " select brandid, is_return_allow, vacation_startday, vacation_endday, tel_start, tel_end, is_saturday_work, brand_comment, last_modifyday, beasongneedday, beasong_comment, beasong_modifyday, beasong_reguserid "
		sqlStr = sqlStr + " , return_comment, return_modifyday, return_reguserid, cs_name, cs_phone, cs_hp, cs_email, lunch_start, lunch_end, vacation_div, cs_modifyday, cs_reguserid, IsNull(customer_return_deny, 'N') as customer_return_deny "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_cs_brand_memo "
        sqlStr = sqlStr + " where brandid='" + FRectMakerid + "'"
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        Fcustomer_return_deny  	= "N"

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

			'// �������ø޸�
            Fbeasongneedday         = rsget("beasongneedday")
            Fbeasong_comment        = db2html(rsget("beasong_comment"))
            Fbeasong_modifyday      = rsget("beasong_modifyday")
            Fbeasong_reguserid      = rsget("beasong_reguserid")

			'// ��ǰ���ø޸�
            Freturn_comment     	= db2html(rsget("return_comment"))
            Freturn_modifyday   	= rsget("return_modifyday")
            Freturn_reguserid    	= rsget("return_reguserid")

			FcsName      			= rsget("cs_name")
			FcsPhone     			= rsget("cs_phone")
			Fcshp        			= rsget("cs_hp")
			FcsEmail     			= rsget("cs_email")
			FcsModifyDay   			= rsget("cs_modifyday")
			FcsReguserID   			= rsget("cs_reguserid")

			Flunch_start   			= rsget("lunch_start")
			Flunch_end     			= rsget("lunch_end")
			Fvacation_div  			= rsget("vacation_div")
            Fcustomer_return_deny  	= rsget("customer_return_deny")

        end if
        rsget.Close
    end sub

    Private Sub Class_Initialize()
        '
    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

''��ǰ�� CS ��۸޸�
Class CCSItemMemo
    public Fitemid
    public Fitemname

	public Fbeasongneedday
	public Fbeasong_comment
	public Fbeasong_modifyday

	public Fmaketoorderyn
	public Fstockshortyn
	public Freipgostartday
	public Freipgoendday

	public Freturn_changemindyn
	public Freturn_comment
	public Freturn_modifyday
	public Freturn_reguserid

	public Fbeasong_reguserid

    public FRectItemid

    public sub GetItemidMemo()
        dim sqlStr

		if (FRectItemid = "") then
			FRectItemid = -1
		end if

        sqlStr = " select m.itemid, m.beasongneedday, m.beasong_comment, m.beasong_modifyday, m.beasong_reguserid, IsNull(m.maketoorderyn, 'N') as maketoorderyn, IsNull(m.stockshortyn, 'N') as stockshortyn, IsNull(m.reipgostartday, convert(varchar(10), getdate(),21)) as reipgostartday, IsNull(m.reipgoendday, convert(varchar(10), getdate(),21)) as reipgoendday, i.itemname "
		sqlStr = sqlStr + " , return_changemindyn, return_comment, return_modifyday, return_reguserid "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_cs_item_memo m "
        sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_item i "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		m.itemid = i.itemid "
        sqlStr = sqlStr + " where m.itemid=" + CStr(FRectItemid) + " "
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        if Not rsget.Eof then
            Fitemid         		= rsget("itemid")
            Fitemname         		= rsget("itemname")

            Fbeasongneedday         = rsget("beasongneedday")
            Fbeasong_comment        = db2html(rsget("beasong_comment"))
            Fbeasong_modifyday      = rsget("beasong_modifyday")

            Fmaketoorderyn      	= rsget("maketoorderyn")
            Fstockshortyn      		= rsget("stockshortyn")
            Freipgostartday      	= rsget("reipgostartday")
            Freipgoendday      		= rsget("reipgoendday")

			Freturn_changemindyn    = rsget("return_changemindyn")
			Freturn_comment      	= db2html(rsget("return_comment"))
			Freturn_modifyday      	= rsget("return_modifyday")
			Freturn_reguserid      	= rsget("return_reguserid")

            Fbeasong_reguserid      = rsget("beasong_reguserid")

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
    public Fconfirmdate
    public Ffinishdate

    public Fsongjangdiv
    public Fsongjangno
	public FsongjangPreNo
	public FsongjangRegGubun
	public FsongjangRegUserID
    public Fbeasongdate
	public Fsongjangfindurl

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


    public Fopentitle           ''�� ���� Title
    public Fopencontents        ''�� ���� ����
    public Fsitegubun           '' 10x10 or theFingers

    public FErrMsg
    public FAuthcode

    public Frefminusorderserial
    public Frefchangeorderserial
    public Freceiveyn

    '// �� �߰���ۺ�(��ǰ, �±�ȯ)
    public Fcustomeraddmethod
    public Fcustomeraddbeasongpay
    public Fcustomerreceiveyn
    public Fcustomerrealbeasongpay

    public Frefasid
    public Freceivestate
    public Freceivefinishdate

    public Forgorderserial

	public Freturnmethod

	Public FneedChkYN
	Public Fpayorderserial
    public Fpaycancelyn
	public Fcustomeradditempay
	public Fcustomeradditembuypay
	public Fcustomerpayordertype

	public function GetCustomerPayOrderTypeName()
		select case Fcustomerpayordertype
			case "B"
				GetCustomerPayOrderTypeName = "��������"
			case "A"
				GetCustomerPayOrderTypeName = "��������"
			case "N"
				GetCustomerPayOrderTypeName = "�ֹ�����"
			case else
				GetCustomerPayOrderTypeName = Fcustomerpayordertype
		end select
	end function

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

            if (iIpkumdiv=8) then
                IsAsRegAvail = false
                descMsg      = "���Ϸ� ���Ŀ��� ȸ����û/��ǰ���� �� �����մϴ�. - ��� �Ұ��� "
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
        elseif (Fdivcd = "A060") then
            '' ��ü��޹���
            IsAsRegAvail = true
        elseif (Fdivcd = "A009") then
            '' ��Ÿ����
            IsAsRegAvail = true
        elseif  (Fdivcd = "A002") or (Fdivcd = "A200") then
            ''���񽺹߼� :��� �����ϰ� ����..
            IsAsRegAvail = true

            'if (iIpkumdiv < 4) then
            '    IsAsRegAvail = false
            '    descMsg      = "����Ϸ� ���������Դϴ�. - ���񽺹߼� ���� �Ұ��� "
            '    exit function
            'end if
        elseif (Fdivcd = "A001") then
            ''������߼�,
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail = false
                descMsg      = "��� �Ϸ�/ �Ϻ� ��� ���°� �ƴմϴ�. - ������߼� ���� �Ұ��� "
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
        elseif (Fdivcd = "A999") then
            ''���߰�����
            IsAsRegAvail = true
        else
            descMsg = "���� ���� �ʾҽ��ϴ�." + Fdivcd
        end if

    end function

    public function IsChangeAsRegAvail(byval iIpkumdiv, byval iCancelYn, byref descMsg)
        IsChangeAsRegAvail = false
        if (iIpkumdiv<2) then
            IsChangeAsRegAvail = false
            descMsg      = "������ �ֹ��� �Ǵ� ���� �ֹ����� �ƴմϴ�. "
            exit function
        end if

        if (IsCancelProcess) then
            IsChangeAsRegAvail = false

            if (iCancelYn<>"N") then
                IsChangeAsRegAvail = false
                descMsg      = "�̹� ��ҵ� �ŷ��Դϴ�. - ��� �Ұ��� "
                exit function
            end if

            if (iIpkumdiv=8) then
                IsChangeAsRegAvail = false
                descMsg      = "���Ϸ� ���Ŀ��� ȸ����û/��ǰ���� �� �����մϴ�. - ��� �Ұ��� "
                exit function
            end if

            if (iIpkumdiv < 8) then
                IsChangeAsRegAvail = false
                descMsg      = "��ȯ�ֹ��Դϴ�. ��ȯȸ�������� �����ϼ���. - ��� �Ұ��� "
                exit function
            end if

            IsChangeAsRegAvail = true

        elseif (IsReturnProcess) then

            if (iCancelYn<>"N") then
                IsChangeAsRegAvail = false
                descMsg      = "��ҵ� �ŷ��Դϴ�. - ��ǰ ���� �Ұ��� "
                exit function
            end if

            IsChangeAsRegAvail = true
        elseif (Fdivcd = "A006") then
            '' ���� ���ǻ���
            IsChangeAsRegAvail = false
            descMsg      = "��ȯ�ֹ��Դϴ�. - ���� ���ǻ��� ���� �Ұ��� "
            exit function
        elseif (Fdivcd = "A009") then
            '' ��Ÿ����
            IsChangeAsRegAvail = true
        elseif  (Fdivcd = "A002") or (Fdivcd = "A200") then
            ''���񽺹߼� :��� �����ϰ� ����..
            IsChangeAsRegAvail = true
        elseif (Fdivcd = "A001") then
            ''������߼�,
            IsChangeAsRegAvail = true
        elseif (Fdivcd = "A000") then
            ''�±�ȯ
            IsChangeAsRegAvail = true
        elseif (Fdivcd = "A003") then
            ''ȯ�ҿ�û
            IsChangeAsRegAvail = true
        elseif (Fdivcd = "A005") then
            ''������ ����Ʈ ���� üũ
            IsChangeAsRegAvail = true
         elseif (Fdivcd = "A700") then
            ''��ü ��Ÿ ����.
            IsChangeAsRegAvail = true
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

        IsRequireSongjangNO = (Fdivcd="A000") or (Fdivcd="A001") or (Fdivcd="A002") or (Fdivcd="A200") or (Fdivcd="A004") or (Fdivcd="A010") or (Fdivcd="A011") or (Fdivcd="A100") or (Fdivcd="A111") or (Fdivcd="A012") or (Fdivcd="A112") or (Fdivcd="A060")
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
    public FRectCsRefAsID
    public FRectNotCsID
    ''
    public FDeliverPay
    public IsUpchebeasongExists
    public IsTenbeasongExists

    public FRectOldOrder

    ''��ü���
    public FRectOnlyJupsu
	public FRectOnlyCustomerJupsu
	public FRectOnlyCSServiceRefund
    public FRectShowAX12
    public FRectReceiveYN
    public FRectExcludeB006YN
    public FRectExcludeA004YN
    public FRectExcludeOLDCSYN


	Public FRectDeleteYN	' �������ܿ���
	Public FRectWriteUser	' �����ھ��̵� �˻�
	Public FRectFinishUser

    Public FRectExtSitename

    Public FRectItemID

	public FRectDateType

	public FRectTplCompanyID
    public farrlist

    public Sub GetHisOldRefundInfo()
        dim i,sqlStr

        sqlStr = " select count(asid) as cnt "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_as_refund_info r, "
        sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_list a"
        sqlStr = sqlStr + " where a.userid='" + FRectUserID + "'"
        sqlStr = sqlStr + " and a.id=r.asid"
        sqlStr = sqlStr + " and a.divcd='A003'"
        sqlStr = sqlStr + " and r.returnmethod='R007'"
        sqlStr = sqlStr + " and a.deleteyn='N'"
		sqlStr = sqlStr + " and DateDiff(m, a.regdate, getdate()) <= 3 "			'// �ֱ� 3������ �˻�
        sqlStr = sqlStr + " and IsNull(r.refundhistorydispyn, 'Y')='Y' "
		'response.write sqlStr

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " r.asid, r.refundrequire, r.rebankname, r.rebankaccount, r.rebankownername, r.encmethod, r.encaccount "
        sqlStr = sqlStr + " , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(r.encaccount), '') WHEN r.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(r.encaccount), '') ELSE '' END) as decaccount "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_as_refund_info r, "
        sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_list a"
        sqlStr = sqlStr + " where a.userid='" + FRectUserID + "'"
        sqlStr = sqlStr + " and a.id=r.asid"
        sqlStr = sqlStr + " and a.divcd='A003'"
        sqlStr = sqlStr + " and r.returnmethod='R007'"
        sqlStr = sqlStr + " and a.deleteyn='N'"
		sqlStr = sqlStr + " and DateDiff(m, a.regdate, getdate()) <= 3 "			'// �ֱ� 3������ �˻�
        sqlStr = sqlStr + " and IsNull(r.refundhistorydispyn, 'Y')='Y' "
        sqlStr = sqlStr + " order by r.asid desc"
		'response.write sqlStr

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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

				FItemList(i).Fasid					= rsget("asid")
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

        sqlStr = "select r.*, IsNull(r.orggiftcardsum, 0) as orggiftcardsum, IsNull(r.orgdepositsum, 0) as orgdepositsum, IsNull(r.refundgiftcardsum, 0) as refundgiftcardsum, IsNull(r.refunddepositsum, 0) as refunddepositsum, IsNull(r.copycouponinfo, 'N') as copycouponinfo "
        sqlStr = sqlStr + " , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(r.encaccount), '') WHEN r.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(r.encaccount), '') ELSE '' END) as decaccount "
        sqlStr = sqlStr + " ,C1.comm_name as returnmethodName"
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_as_refund_info r with (nolock)"
        sqlStr = sqlStr + "     left join [db_cs].[dbo].tbl_cs_comm_code C1"
        sqlStr = sqlStr + "     on C1.comm_group='Z090'"
        sqlStr = sqlStr + "     and r.returnmethod=C1.comm_cd"
        sqlStr = sqlStr + " where asid=" + CStr(FRectCsAsID)
		''response.write sqlStr

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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

            FOneItem.FpaygateresultTid      = rsget("paygateresultTid")
            FOneItem.FpaygateresultMsg      = rsget("paygateresultMsg")


            FOneItem.FreturnmethodName      = rsget("returnmethodName")

            FOneItem.Forggiftcardsum      	= rsget("orggiftcardsum")
            FOneItem.Forgdepositsum      	= rsget("orgdepositsum")
            FOneItem.Frefundgiftcardsum     = rsget("refundgiftcardsum")
            FOneItem.Frefunddepositsum      = rsget("refunddepositsum")

            FOneItem.Fupfiledate      		= rsget("upfiledate")
            FOneItem.FdecAccount            = rsget("decAccount")

			FOneItem.Fcopycouponinfo        = rsget("copycouponinfo")
            FOneItem.fcopyitemcouponinfo        = rsget("copyitemcouponinfo")

            if IsNull(FOneItem.Forgmileagesum) then
            	FOneItem.Forgmileagesum = 0
            end if
        end if
        rsget.Close
    end Sub

	'��ȯ�Ҿ��հ�(CS��������)
    '// 2018-05-12, ī����� ����, skyer9
    public Function GetPrevRefundSum()
        dim i,sqlStr
        dim result

        sqlStr = " select "
        sqlStr = sqlStr + " 	IsNull(sum(r.refundrequire), 0) as refundrequire "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list a "
        sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_as_refund_info r "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		a.id=r.asid "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and a.orderserial='" & FRectOrderSerial & "' "
        sqlStr = sqlStr + " 	and a.deleteyn='N' "
        sqlStr = sqlStr + " 	and a.divcd in ('A003', 'A007') "
        sqlStr = sqlStr + " 	and a.refasid is not null "

        result = 0

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if Not rsget.Eof then
            result = rsget("refundrequire")
        end if
        rsget.Close

        GetPrevRefundSum = result
    end Function

	'��ۺ�CSȯ�ұݾ�(CS��������)
	'��ۺ� ��� ���� ��ۺ�ȯ���� �̷���� �ݾ�
    public Function GetPrevRefundCSDeliveryPaySum()
        dim i,sqlStr
        dim result

        sqlStr = " select "
        sqlStr = sqlStr + " 	IsNull(sum(r.refundbeasongpay),0) as refundbeasongpay "
        sqlStr = sqlStr + " 	, IsNull(sum(T.realdeliverypay),0) as realdeliverypay "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m "
        sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list refm "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		m.refasid = refm.id "
        sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_as_refund_info r "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		refm.id = r.asid "
        sqlStr = sqlStr + " 	left join ( "
        sqlStr = sqlStr + " 		select "
        sqlStr = sqlStr + " 			refm.id "
        sqlStr = sqlStr + " 			, sum(case when refd.itemid = 0 then refd.confirmitemno*refd.itemcost else 0 end) as realdeliverypay "
        sqlStr = sqlStr + " 		from "
        sqlStr = sqlStr + " 			[db_cs].[dbo].tbl_new_as_list refm "
        sqlStr = sqlStr + " 			join [db_cs].dbo.tbl_new_as_detail refd "
        sqlStr = sqlStr + " 			on "
        sqlStr = sqlStr + " 				refm.id = refd.masterid "
        sqlStr = sqlStr + " 		where "
        sqlStr = sqlStr + " 			refm.orderserial = '" & FRectOrderSerial & "' "
        sqlStr = sqlStr + " 		group by "
        sqlStr = sqlStr + " 			refm.id "
        sqlStr = sqlStr + " 	) T "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		refm.id = T.id "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and m.orderserial = '" & FRectOrderSerial & "' "
        sqlStr = sqlStr + " 	and m.deleteyn='N' "
        sqlStr = sqlStr + " 	and m.divcd='A003' "

        result = 0

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if Not rsget.Eof then
            result = rsget("refundbeasongpay") - rsget("realdeliverypay")
        end if
        rsget.Close

        GetPrevRefundCSDeliveryPaySum = result
    end Function

    public Sub GetCSASMasterList()
        dim i,sqlStr, AddSQL
        AddSQL = ""

        sqlStr = " select count(A.id) as cnt "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A"
		sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list B "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	A.id = B.refasid "
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " where 1 = 1 "

		if (FRectSearchType="") then
		    if (FRectOrderSerial<>"") then
		        AddSQL = AddSQL + " and A.orderserial='" + FRectOrderSerial + "'"
		    end if
		elseif (FRectSearchType="upcheview") then
		    ''��ü�� ������

            if (FRectDivcd <> "A012") and (FRectDivcd <> "A112") and (FRectShowAX12 = "") then
            	AddSQL = AddSQL + " and A.divcd not in ('A005','A007', 'A012', 'A112')"
            else
            	AddSQL = AddSQL + " and A.divcd not in ('A005','A007')"
            end if

            AddSQL = AddSQL + " and A.deleteyn='N'"
            AddSQL = AddSQL + " and A.requireupche='Y' "
            AddSQL = AddSQL + " and A.makerid='" + CStr(FRectMakerid) + "' "

            if (FRectOnlyJupsu="on") then
                AddSQL = AddSQL + " and A.currstate='B001'"
            end if

            if (FRectCurrstate = "notfinish") then
	                AddSQL = AddSQL + " and ((A.currstate < 'B007') or ((IsNull(B.currstate, 'B007') < 'B007') and (B.divcd in ('A012', 'A112')))) "
	        elseif (FRectCurrstate <> "") then
	                AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

            if (FRectReceiveYN <> "") then
            	AddSQL = AddSQL + " and B.currstate is not NULL and B.divcd in ('A012', 'A112') "
            	if (FRectReceiveYN = "Y") then
            		AddSQL = AddSQL + " and B.currstate >= 'B006' "
            	elseif (FRectReceiveYN = "N") then
            		AddSQL = AddSQL + " and B.currstate < 'B006' "
            	end if
	        end if

            if (FRectExcludeB006YN <> "") then
            	AddSQL = AddSQL + " and ((A.currstate < 'B006') or (B.divcd in ('A012', 'A112') and B.currstate < 'B006')) "
	        end if

            if (FRectExcludeA004YN <> "") then
            	AddSQL = AddSQL + " and (A.divcd <> 'A004') "
	        end if

            if (FRectExcludeOLDCSYN <> "") then
            	AddSQL = AddSQL + " and ((A.currstate >= 'B006') or (datediff(m, A.regdate, getdate()) <= 3)) "
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

	        if (FRectDivcd <> "") then
	                AddSQL = AddSQL + " and A.divcd ='" + CStr(FRectDivcd) + "' "
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
	                if (FRectCurrstate = "B006") and (FRectWriteUser <> "") then
	                	'CS �����ں� ��üó���Ϸῡ�� �±�ȯȸ���Ϸ� ���� ����
	                	AddSQL = AddSQL + " and A.currstate='B006' and (A.divcd not in ('A000', 'A100') or IsNull(B.currstate, 'B007') = 'B007') and A.requireupche='Y' and A.deleteyn='N' "
	                else
	                	AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	                end if
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
                '���ϸ���/��ġ�� ȯ�ҹ�ó��
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A003' "
                AddSQL = AddSQL + " and A.deleteyn='N'"
                AddSQL = AddSQL + " and R.returnmethod in ('R900', 'R910') "
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
        elseif (FRectSearchType = "upreturnmifinish") then
                '��ü��ǰ ��ó��
                AddSQL = AddSQL + " and A.divcd='A004' and A.currstate<'B006' and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "upchemifinish") then
                '��ü��ó��
                AddSQL = AddSQL + " and A.currstate<'B006' and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "upchefinish") then
                '��üó���Ϸ�
                AddSQL = AddSQL + " and A.currstate='B006' and (A.divcd not in ('A000', 'A100') or IsNull(B.currstate, 'B007') = 'B007') and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "returnmifinish") then
                'ȸ����û��ó��
                AddSQL = AddSQL + " and A.currstate<'B007' and A.requireupche<>'Y' and A.divcd in ('A010', 'A011', 'A111') and A.deleteyn='N' "
        elseif (FRectSearchType = "confirm") then
                'Ȯ�ο�û ��ó��
                AddSQL = AddSQL + " and A.currstate='B005' and A.deleteyn='N' "
        elseif (FRectSearchType = "cancelnofinish") then
                '�ֹ���� ��ó��
                AddSQL = AddSQL + " and A.divcd='A008'"
                AddSQL = AddSQL + " and A.currstate='B001' and A.deleteyn='N' "
                AddSQL = AddSQL + " and A.regdate>'2008-04-23'"
        end If

        IF (FRectExtSitename<>"") then
            AddSQL = AddSQL + " and A.ExtSitename='"&FRectExtSitename&"'"
        END IF

        sqlStr = sqlStr + AddSQL

		'rw sqlStr

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        if  not rsget.EOF  then
            FTotalCount = rsget("cnt")
        else
            FTotalCount = 0
        end if
        rsget.close


        sqlStr = " select      Top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr + "     A.id, A.divcd, A.gubun01, A.gubun02, A.orderserial, A.customername, A.userid, A.finishuser, A.writeuser, A.title, A.currstate, B.currstate as receivestate, B.finishdate as receivefinishdate "
        sqlStr = sqlStr + "     ,A.regdate, A.finishdate,A.deleteyn "
        sqlStr = sqlStr + "     , A.requireupche, A.makerid, A.songjangdiv ,A.songjangno, A.receiveyn"
        sqlStr = sqlStr + "     ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + "     ,m.sitename, m.authcode"
        sqlStr = sqlStr + "     ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename, C4.comm_color as currstatecolor"
        sqlStr = sqlStr + "     ,p.payorderserial, IsNull(p.additempay, 0) as additempay, IsNull(p.addbeasongpay,0) as addbeasongpay, IsNull(po.cancelyn, 'N') as paycancelyn "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A"
		sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list B "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	A.id = B.refasid "
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_order].[dbo].tbl_order_master m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " left join [db_cs].[dbo].[tbl_as_customer_addbeasongpay_info] p on A.id = p.asid"
        sqlStr = sqlStr + " left join [db_order].[dbo].[tbl_order_master] po on p.payorderserial = po.orderserial "
        sqlStr = sqlStr + " where 1 = 1 "

        sqlStr = sqlStr + AddSQL

        sqlStr = sqlStr + " order by A.id desc "

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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

                FItemList(i).Freceiveyn         = rsget("receiveyn")
                FItemList(i).Freceivestate		= rsget("receivestate")
                FItemList(i).Freceivefinishdate		= rsget("receivefinishdate")

                FItemList(i).Fpayorderserial			= rsget("payorderserial")
                FItemList(i).Fcustomeraddbeasongpay		= rsget("addbeasongpay")
                FItemList(i).Fcustomeradditempay		= rsget("additempay")

                FItemList(i).Fpaycancelyn		= rsget("paycancelyn")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close
    end sub

    public Sub GetCSASMasterList_3PL()
        dim i,sqlStr, AddSQL
        AddSQL = ""

        sqlStr = " select count(A.id) as cnt "
        sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_as_list] A"
		sqlStr = sqlStr + " left join [db_threepl].[dbo].[tbl_tpl_as_list] B "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	A.id = B.refasid "
        sqlStr = sqlStr + " where 1 = 1 "

		if (FRectSearchType="") then
		    if (FRectOrderSerial<>"") then
		        AddSQL = AddSQL + " and A.orderserial='" + FRectOrderSerial + "'"
		    end if
		elseif (FRectSearchType="upcheview") then
		    ''��ü�� ������

            if (FRectDivcd <> "A012") and (FRectDivcd <> "A112") and (FRectShowAX12 = "") then
            	AddSQL = AddSQL + " and A.divcd not in ('A005','A007', 'A012', 'A112')"
            else
            	AddSQL = AddSQL + " and A.divcd not in ('A005','A007')"
            end if

            AddSQL = AddSQL + " and A.deleteyn='N'"
            AddSQL = AddSQL + " and A.requireupche='Y' "
            AddSQL = AddSQL + " and A.makerid='" + CStr(FRectMakerid) + "' "

            if (FRectOnlyJupsu="on") then
                AddSQL = AddSQL + " and A.currstate='B001'"
            end if

            if (FRectCurrstate = "notfinish") then
	                AddSQL = AddSQL + " and ((A.currstate < 'B007') or ((IsNull(B.currstate, 'B007') < 'B007') and (B.divcd in ('A012', 'A112')))) "
	        elseif (FRectCurrstate <> "") then
	                AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

            if (FRectReceiveYN <> "") then
            	AddSQL = AddSQL + " and B.currstate is not NULL and B.divcd in ('A012', 'A112') "
            	if (FRectReceiveYN = "Y") then
            		AddSQL = AddSQL + " and B.currstate >= 'B006' "
            	elseif (FRectReceiveYN = "N") then
            		AddSQL = AddSQL + " and B.currstate < 'B006' "
            	end if
	        end if

            if (FRectExcludeB006YN <> "") then
            	AddSQL = AddSQL + " and ((A.currstate < 'B006') or (B.divcd in ('A012', 'A112') and B.currstate < 'B006')) "
	        end if

            if (FRectExcludeA004YN <> "") then
            	AddSQL = AddSQL + " and (A.divcd <> 'A004') "
	        end if

            if (FRectExcludeOLDCSYN <> "") then
            	AddSQL = AddSQL + " and ((A.currstate >= 'B006') or (datediff(m, A.regdate, getdate()) <= 3)) "
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

	        if (FRectDivcd <> "") then
	                AddSQL = AddSQL + " and A.divcd ='" + CStr(FRectDivcd) + "' "
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
	                if (FRectCurrstate = "B006") and (FRectWriteUser <> "") then
	                	'CS �����ں� ��üó���Ϸῡ�� �±�ȯȸ���Ϸ� ���� ����
	                	AddSQL = AddSQL + " and A.currstate='B006' and (A.divcd not in ('A000', 'A100') or IsNull(B.currstate, 'B007') = 'B007') and A.requireupche='Y' and A.deleteyn='N' "
	                else
	                	AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	                end if
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
                '���ϸ���/��ġ�� ȯ�ҹ�ó��
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A003' "
                AddSQL = AddSQL + " and A.deleteyn='N'"
                AddSQL = AddSQL + " and R.returnmethod in ('R900', 'R910') "
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
        elseif (FRectSearchType = "upreturnmifinish") then
                '��ü��ǰ ��ó��
                AddSQL = AddSQL + " and A.divcd='A004' and A.currstate<'B006' and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "upchemifinish") then
                '��ü��ó��
                AddSQL = AddSQL + " and A.currstate<'B006' and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "upchefinish") then
                '��üó���Ϸ�
                AddSQL = AddSQL + " and A.currstate='B006' and (A.divcd not in ('A000', 'A100') or IsNull(B.currstate, 'B007') = 'B007') and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "returnmifinish") then
                'ȸ����û��ó��
                AddSQL = AddSQL + " and A.currstate<'B007' and A.requireupche<>'Y' and A.divcd in ('A010', 'A011', 'A111') and A.deleteyn='N' "
        elseif (FRectSearchType = "confirm") then
                'Ȯ�ο�û ��ó��
                AddSQL = AddSQL + " and A.currstate='B005' and A.deleteyn='N' "
        elseif (FRectSearchType = "cancelnofinish") then
                '�ֹ���� ��ó��
                AddSQL = AddSQL + " and A.divcd='A008'"
                AddSQL = AddSQL + " and A.currstate='B001' and A.deleteyn='N' "
                AddSQL = AddSQL + " and A.regdate>'2008-04-23'"
        end If

        IF (FRectExtSitename<>"") then
            AddSQL = AddSQL + " and A.ExtSitename='"&FRectExtSitename&"'"
        END IF

        sqlStr = sqlStr + AddSQL

		'rw sqlStr

        rsget_TPL.Open sqlStr, dbget_TPL, 1

        if  not rsget_TPL.EOF  then
            FTotalCount = rsget_TPL("cnt")
        else
            FTotalCount = 0
        end if
        rsget_TPL.close


        sqlStr = " select      Top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr + "     A.id, A.divcd, A.gubun01, A.gubun02, A.orderserial, A.customername, A.userid, A.finishuser, A.writeuser, A.title, A.currstate, B.currstate as receivestate, B.finishdate as receivefinishdate "
        sqlStr = sqlStr + "     ,A.regdate, A.finishdate,A.deleteyn "
        sqlStr = sqlStr + "     , A.requireupche, A.makerid, A.songjangdiv ,A.songjangno, A.receiveyn"
        sqlStr = sqlStr + "     ,0 as refundrequire, 0 as refundresult"
        sqlStr = sqlStr + "     ,m.sitename, '' as authcode"
        sqlStr = sqlStr + "     ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename, C4.comm_color as currstatecolor"
        sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_as_list] A"
		sqlStr = sqlStr + " left join [db_threepl].[dbo].[tbl_tpl_as_list] B "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	A.id = B.refasid "
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_orderMaster] m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " where 1 = 1 "

        sqlStr = sqlStr + AddSQL

        sqlStr = sqlStr + " order by A.id desc "

        rsget_TPL.pagesize = FPageSize
        rsget_TPL.Open sqlStr, dbget_TPL, 1

        FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

        redim preserve FItemList(FResultCount)
        if  not rsget_TPL.EOF  then
            i = 0
			rsget_TPL.absolutepage = FCurrPage
            do until rsget_TPL.eof
                set FItemList(i) = new CCSASMasterItem

                FItemList(i).Fid                = rsget_TPL("id")
                FItemList(i).Fdivcd             = rsget_TPL("divcd")
                FItemList(i).FdivcdName         = db2html(rsget_TPL("divcdname"))

                FItemList(i).Forderserial       = rsget_TPL("orderserial")
                FItemList(i).Fcustomername      = db2html(rsget_TPL("customername"))
                FItemList(i).Fuserid            = rsget_TPL("userid")
                FItemList(i).Fwriteuser         = rsget_TPL("writeuser")
                FItemList(i).Ffinishuser        = rsget_TPL("finishuser")
                FItemList(i).Ftitle             = db2html(rsget_TPL("title"))
                FItemList(i).Fcurrstate         = rsget_TPL("currstate")
                FItemList(i).Fcurrstatename     = rsget_TPL("currstatename")
                FItemList(i).FcurrstateColor    = rsget_TPL("currstatecolor")

                FItemList(i).Fregdate           = rsget_TPL("regdate")
                FItemList(i).Ffinishdate        = rsget_TPL("finishdate")

                FItemList(i).Fgubun01           = rsget_TPL("gubun01")
                FItemList(i).Fgubun02           = rsget_TPL("gubun02")

                FItemList(i).Fgubun01Name       = db2html(rsget_TPL("gubun01name"))
                FItemList(i).Fgubun02Name       = db2html(rsget_TPL("gubun02name"))

                FItemList(i).Fdeleteyn          = rsget_TPL("deleteyn")

                FItemList(i).Frefundrequire     = rsget_TPL("refundrequire")
                FItemList(i).Frefundresult      = rsget_TPL("refundresult")

                FItemList(i).Fsongjangdiv       = rsget_TPL("songjangdiv")
                FItemList(i).Fsongjangno        = rsget_TPL("songjangno")

                FItemList(i).Frequireupche      = rsget_TPL("requireupche")
                FItemList(i).Fmakerid           = rsget_TPL("makerid")

                FItemList(i).FExtsitename          = rsget_TPL("sitename")
                FItemList(i).Fauthcode          = rsget_TPL("authcode")

                FItemList(i).Freceiveyn         = rsget_TPL("receiveyn")
                FItemList(i).Freceivestate		= rsget_TPL("receivestate")
                FItemList(i).Freceivefinishdate		= rsget_TPL("receivefinishdate")



                rsget_TPL.MoveNext
                i = i + 1
            loop
        end if
        rsget_TPL.close
    end sub

    public Sub GetCSASMasterListUpcheNew()
        dim i,sqlStr, AddSQL
        AddSQL = ""

        AddSQL = AddSQL + " and A.deleteyn='N'"
        AddSQL = AddSQL + " and A.requireupche='Y' "
        AddSQL = AddSQL + " and A.makerid='" + CStr(FRectMakerid) + "' "
		AddSQL = AddSQL + " and A.divcd not in ('A003','A005','A007','A999')"

	    if (FRectOrderSerial<>"") then
	        ''AddSQL = AddSQL + " and A.orderserial='" + FRectOrderSerial + "'"

			AddSQL = AddSQL + " 	and a.orderserial in ( "
			AddSQL = AddSQL + " 		select chgorderserial "
			AddSQL = AddSQL + " 		from "
			AddSQL = AddSQL + " 		db_order.dbo.tbl_change_order "
			AddSQL = AddSQL + " 		where orgorderserial = '" + FRectOrderSerial + "' and deldate is null  "
			AddSQL = AddSQL + " 		union all "
			AddSQL = AddSQL + " 		select '" + FRectOrderSerial + "' "
			AddSQL = AddSQL + " 	) "
		end if

        if (FRectOnlyJupsu="on") then
            AddSQL = AddSQL + " and A.currstate='B001'"
        end if

        if (FRectCurrstate = "notfinish") then
                AddSQL = AddSQL + " and ((A.currstate < 'B007') or ((IsNull(B.currstate, 'B007') < 'B007') and (B.divcd in ('A012', 'A112')))) "
        elseif (FRectCurrstate <> "") then
                AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
        end if

        if (FRectReceiveYN <> "") then
        	AddSQL = AddSQL + " and B.currstate is not NULL and B.divcd in ('A012', 'A112') "
        	if (FRectReceiveYN = "Y") then
        		AddSQL = AddSQL + " and B.currstate >= 'B006' "
        	elseif (FRectReceiveYN = "N") then
        		AddSQL = AddSQL + " and B.currstate < 'B006' "
        	end if
        end if

        if (FRectExcludeB006YN <> "") then
        	AddSQL = AddSQL + " and (A.currstate < 'B006') "
        end if

        if (FRectExcludeA004YN <> "") then
        	AddSQL = AddSQL + " and (A.divcd <> 'A004') "
        end if

        if (FRectExcludeOLDCSYN <> "") then
        	AddSQL = AddSQL + " and ((A.currstate >= 'B006') or (datediff(m, A.regdate, getdate()) <= 3)) "
        end if

        if (FRectUserName <> "") then
                AddSQL = AddSQL + " and A.customername='" + CStr(FRectUserName) + "' "
        end if

        if (FRectUserID <> "") then
                AddSQL = AddSQL + " and A.userid='" + CStr(FRectUserID) + "' "
        end if

        if (FRectDivcd <> "") then
                AddSQL = AddSQL + " and A.divcd ='" + CStr(FRectDivcd) + "' "
        end if

        IF (FRectExtSitename<>"") then
            AddSQL = AddSQL + " and A.ExtSitename='"&FRectExtSitename&"'"
        END IF

        IF (FRectItemID<>"") then
            '// ��Ÿ �˻������� �ִ� ��츸 �˻�

            if (FRectOrderSerial <> "") or (FRectUserName <> "") or (FRectUserID <> "") then
				AddSQL = AddSQL + " 	and a.orderserial in ( "
				AddSQL = AddSQL + " 		select distinct m.orderserial "
				AddSQL = AddSQL + " 		from "
				AddSQL = AddSQL + " 			[db_cs].[dbo].tbl_new_as_list m "
				AddSQL = AddSQL + " 			join [db_cs].dbo.tbl_new_as_detail d "
				AddSQL = AddSQL + " 			on "
				AddSQL = AddSQL + " 				m.id = d.masterid "
				AddSQL = AddSQL + " 		where "
				AddSQL = AddSQL + " 			1 = 1 "

				if (FRectOrderSerial <> "") then
					AddSQL = AddSQL + " 			and m.orderserial = '" + CStr(FRectOrderSerial) + "' "
				end if

				if (FRectUserName <> "") then
					AddSQL = AddSQL + " 			and m.customername = '" + CStr(FRectUserName) + "' "
				end if

				if (FRectUserID <> "") then
					AddSQL = AddSQL + " 			and m.userid = '" + CStr(FRectUserID) + "' "
				end if

				AddSQL = AddSQL + " 			and d.itemid = " + CStr(FRectItemID) + " "

				AddSQL = AddSQL + " 	) "
            end if
        END IF

        '�Ⱓ�˻�
        Select Case FRectDateType
            Case "regdate"
                AddSQL = AddSQL & " and A.regdate between '" & FRectStartDate & "' and '" & DateAdd("d",1,FRectEndDate) & "'"
            Case "finishdate"
                AddSQL = AddSQL & " and A.finishdate between '" & FRectStartDate & "' and '" & DateAdd("d",1,FRectEndDate) & "'"
        End Select


		'// ===================================================================
        sqlStr = " select count(A.id) as cnt "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A"
		sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list B "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	A.id = B.refasid "
        sqlStr = sqlStr + " left join db_order.dbo.tbl_change_order c "
        sqlStr = sqlStr + " on "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and a.orderserial = c.chgorderserial "
        sqlStr = sqlStr + " 	and c.deldate is null "
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " where 1 = 1 "

        sqlStr = sqlStr + AddSQL

		''rw sqlStr

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        if  not rsget.EOF  then
            FTotalCount = rsget("cnt")
        else
            FTotalCount = 0
        end if
        rsget.close


		'// ===================================================================
        sqlStr = " select      Top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr + "     A.id, A.divcd, A.gubun01, A.gubun02, A.orderserial, A.customername, A.userid, A.finishuser, A.writeuser, A.title, A.currstate, B.currstate as receivestate, B.finishdate as receivefinishdate "
        sqlStr = sqlStr + "     ,A.regdate, A.finishdate,A.deleteyn "
        sqlStr = sqlStr + "     , A.requireupche, A.makerid, A.songjangdiv ,A.songjangno, A.receiveyn"
        sqlStr = sqlStr + "     ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + "     ,m.sitename, m.authcode"
        sqlStr = sqlStr + "     ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename, C4.comm_color as currstatecolor"
        sqlStr = sqlStr + " 	, IsNull(c.orgorderserial, a.orderserial) as orgorderserial "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A"
		sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list B "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	A.id = B.refasid "
        sqlStr = sqlStr + " left join db_order.dbo.tbl_change_order c "
        sqlStr = sqlStr + " on "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and a.orderserial = c.chgorderserial "
        sqlStr = sqlStr + " 	and c.deldate is null "
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_order].[dbo].tbl_order_master m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " where 1 = 1 "

        sqlStr = sqlStr + AddSQL

        sqlStr = sqlStr + " order by A.id desc "

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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

                FItemList(i).Freceiveyn         = rsget("receiveyn")
                FItemList(i).Freceivestate		= rsget("receivestate")
                FItemList(i).Freceivefinishdate		= rsget("receivefinishdate")

				FItemList(i).Forgorderserial	= rsget("orgorderserial")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close
    end sub


    public Sub GetCSASMasterListNew()
        dim i,sqlStr, AddSQL
        dim orgorderserial

		if (FRectOrderSerial <> "") then
			'��ȯ�ֹ���ȣ -> ���ֹ���ȣ

			orgorderserial = FRectOrderSerial

			'// �ϴ� ����.(skyer9)
			''sqlStr = " select top 1 orgorderserial from db_order.dbo.tbl_change_order where chgorderserial = '" + FRectOrderSerial + "' and deldate is null "
	        ''rsget.CursorLocation = adUseClient
			''rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	        ''if Not rsget.Eof then
	        ''    orgorderserial = rsget("orgorderserial")
	        ''end if
	        ''rsget.Close

	        ''FRectOrderSerial = orgorderserial
		end if

        AddSQL = ""

        sqlStr = " select count(A.id) as cnt "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A"
		sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list B "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	A.id = B.refasid "
        sqlStr = sqlStr + " left join db_order.dbo.tbl_change_order c "
        sqlStr = sqlStr + " on "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and a.orderserial = c.chgorderserial "
        sqlStr = sqlStr + " 	and c.deldate is null "
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"

        sqlStr = sqlStr + " where 1 = 1 "

        if (FRectUserID <> "") then
			AddSQL = AddSQL + " and A.userid='" + CStr(FRectUserID) + "' "
        end if

        if (FRectUserName <> "") then
			AddSQL = AddSQL + " and A.customername='" + CStr(FRectUserName) + "' "
        end if

		if (FRectOrderSerial <> "") then
			'AddSQL = AddSQL + " and a.orderserial = '" + FRectOrderSerial + "' "

			AddSQL = AddSQL + " 	and a.orderserial in ( "
			AddSQL = AddSQL + " 		select chgorderserial "
			AddSQL = AddSQL + " 		from "
			AddSQL = AddSQL + " 		db_order.dbo.tbl_change_order "
			AddSQL = AddSQL + " 		where orgorderserial = '" + FRectOrderSerial + "' and deldate is null  "
			AddSQL = AddSQL + " 		union all "
			AddSQL = AddSQL + " 		select '" + FRectOrderSerial + "' "
			AddSQL = AddSQL + " 	) "
		end if

        if (FRectMakerid<>"") then
			AddSQL = AddSQL + " and A.requireupche='Y' "
			AddSQL = AddSQL + " and A.makerid='" + CStr(FRectMakerid) + "' "
        end if

		if (FRectWriteUser <> "") then
			AddSQL = AddSQL + " and A.writeUser = '" + CStr(FRectWriteUser) + "' "
		end if

		if (FRectDivcd <> "") then
			AddSQL = AddSQL + " and A.divcd ='" + CStr(FRectDivcd) + "' "
		end if

		if (FRectCurrstate = "notfinish") then
			AddSQL = AddSQL + " and (A.currstate < 'B007') "
		elseif (FRectCurrstate <> "") then
			AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
		end if

		if (FRectDeleteYN <> "") then
			AddSQL = AddSQL + " and A.deleteyn = '" + CStr(FRectDeleteYN) + "' "
		end if

		if (FRectSearchType <> "") then
			'��ó��CS

			AddSQL = AddSQL + " and A.currstate<'B007' "

			if (FRectSearchType = "notfinish") then
				''��ó����ü
			elseif (FRectSearchType = "norefund") then
				'ȯ�ҹ�ó��
				AddSQL = AddSQL + " and A.divcd='A003' "
			elseif (FRectSearchType = "norefundmile") then
				'���ϸ���/��ġ�� ȯ�ҹ�ó��
				AddSQL = AddSQL + " and A.divcd='A003' "
				AddSQL = AddSQL + " and R.returnmethod in ('R900', 'R910') "
			elseif (FRectSearchType = "norefundetc") then
				'�ܺθ�ȯ�ҹ�ó��
				AddSQL = AddSQL + " and A.divcd='A005' "
			elseif (FRectSearchType = "cardnocheck") then
				'ī����ҹ�ó��
				AddSQL = AddSQL + " and A.divcd='A007' "
			elseif (FRectSearchType = "beasongnocheck") then
				'������ǻ���/���
				AddSQL = AddSQL + " and A.divcd in ('A008','A006') and (IsNull(A.requireupche, 'N') = 'N') "
			elseif (FRectSearchType = "upreturnmifinish") then
				'��ü��ǰ ��ó��
				AddSQL = AddSQL + " and A.divcd='A004' and A.currstate<'B006' and A.requireupche='Y' "
			elseif (FRectSearchType = "upchemifinish") then
				'��ü��ó��
				AddSQL = AddSQL + " and A.currstate<'B006' and A.requireupche='Y' "
			elseif (FRectSearchType = "upchefinish") then
				'��üó���Ϸ�
				'// ��ȯȸ�� �Ϸ����� ���� ����
				AddSQL = AddSQL + " and A.currstate='B006' and (A.divcd not in ('A000', 'A100') or IsNull(B.currstate, 'B007') = 'B007') and A.requireupche='Y' "
			elseif (FRectSearchType = "chulgofinishnotreceive") then
				'��ȯ����Ĺ�ȸ��
				AddSQL = AddSQL + " and A.currstate='B006' and (A.divcd in ('A000', 'A100') or IsNull(B.currstate, 'B007') < 'B006') "
			elseif (FRectSearchType = "returnmifinish") then
				'ȸ����û��ó��
				AddSQL = AddSQL + " and A.requireupche<>'Y' and A.divcd in ('A010', 'A011', 'A111') "
			elseif (FRectSearchType = "confirm") then
				'Ȯ�ο�û ��ó��
				AddSQL = AddSQL + " and A.currstate='B005' "
			elseif (FRectSearchType = "cancelnofinish") then
				'�ֹ���� ��ó��
				AddSQL = AddSQL + " and A.divcd='A008'"
				AddSQL = AddSQL + " and A.currstate='B001' "
				AddSQL = AddSQL + " and A.regdate>'2008-04-23'"
			end If
		end if

		if (FRectStartDate <> "") then
			AddSQL = AddSQL + " and A.regdate>='" + CStr(FRectStartDate) + "' "
		end if

		if (FRectEndDate <> "") then
			AddSQL = AddSQL + " and A.regdate <'" + CStr(FRectEndDate) + "' "
		end if

        IF (FRectExtSitename<>"") then
            AddSQL = AddSQL + " and A.ExtSitename='"&FRectExtSitename&"'"
        END IF

        sqlStr = sqlStr + AddSQL

		''rw sqlStr

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        if  not rsget.EOF  then
            FTotalCount = rsget("cnt")
        else
            FTotalCount = 0
        end if
        rsget.close


        sqlStr = " select      Top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr + "     A.id, A.divcd, A.gubun01, A.gubun02, A.orderserial, A.customername, A.userid, A.finishuser, A.writeuser, A.title, A.currstate, B.currstate as receivestate, B.finishdate as receivefinishdate "
        sqlStr = sqlStr + "     ,A.regdate, A.finishdate, A.confirmdate, A.deleteyn "
        sqlStr = sqlStr + "     , A.requireupche, A.makerid, A.songjangdiv ,A.songjangno, A.receiveyn"
        sqlStr = sqlStr + "     ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + "     ,m.sitename, m.authcode"
        sqlStr = sqlStr + "     ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename, C4.comm_color as currstatecolor"
        sqlStr = sqlStr + " 	, IsNull(c.orgorderserial, a.orderserial) as orgorderserial "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A"
		sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list B "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	A.id = B.refasid "
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_order].[dbo].tbl_order_master m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " left join db_order.dbo.tbl_change_order c "
        sqlStr = sqlStr + " on "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and a.orderserial = c.chgorderserial "
        sqlStr = sqlStr + " 	and c.deldate is null "
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " where 1 = 1 "

        sqlStr = sqlStr + AddSQL

        sqlStr = sqlStr + " order by A.id desc "

'rw sqlStr
        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
                FItemList(i).Fconfirmdate       = rsget("confirmdate")
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

                FItemList(i).Freceiveyn         = rsget("receiveyn")
                FItemList(i).Freceivestate		= rsget("receivestate")
                FItemList(i).Freceivefinishdate		= rsget("receivefinishdate")

                FItemList(i).Forgorderserial	= rsget("orgorderserial")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close
    end sub

    ' /cscenter/action/cs_action_list.asp
    public Sub GetCSASMasterListByProcedure_notpaging()
        dim i,sqlStr, AddSQL, topN

		SqlStr = "exec [db_cs].[dbo].[usp_Ten_CsAsListNew_detail] " & CStr(FPageSize * FCurrPage) & ", '" & FRectDivcd & "', '" & FRectCurrstate & "', '" & FRectDeleteYN & "', '" & FRectOrderSerial & "' "
		sqlStr = sqlStr + " , '" & FRectUserID & "', '" &FRectUserName & "', '" & FRectMakerid & "', '" & FRectWriteUser & "', '" & FRectSearchType & "', '" & FRectStartDate & "', '" & FRectEndDate & "', '" & FRectExtSitename & "', '" & FRectOnlyCustomerJupsu & "', '" & FRectOnlyCSServiceRefund & "', '" + CStr(FRectCsAsID) + "', '" & FRectDateType & "', '" & FRectFinishUser & "' "

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount
        FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.close
    end sub

    ' /cscenter/action/cs_action_list.asp
    public Sub GetCSASMasterListByProcedure()
        dim i,sqlStr, AddSQL, topN

		SqlStr = "exec [db_cs].[dbo].[usp_Ten_CsAsCountNew] '" & FRectDivcd & "', '" & FRectCurrstate & "', '" & FRectDeleteYN & "', '" & FRectOrderSerial & "' "
		sqlStr = sqlStr + " , '" & FRectUserID & "', '" &FRectUserName & "', '" & FRectMakerid & "', '" & FRectWriteUser & "', '" & FRectSearchType & "', '" & FRectStartDate & "', '" & FRectEndDate & "', '" & FRectExtSitename & "', '" & FRectOnlyCustomerJupsu & "', '" & FRectOnlyCSServiceRefund & "', '" + CStr(FRectCsAsID) + "', '" & FRectDateType & "', '" & FRectFinishUser & "' "

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        if  not rsget.EOF  then
            FTotalCount = rsget("cnt")
        else
            FTotalCount = 0
        end if
        rsget.close

		'// ====================================================================
		'// �ӵ����� �ذ�
		'// 1. FTotalCount ���� topN �� Ŭ�� Ǯ��ĵ ����
		'// 2. FTotalCount �� 0 �� �� �˻���� ������ �ֹ���ȣ ����
		if (FTotalCount <= FPageSize) then
			FCurrPage = 1
		end if

		topN = FPageSize * FCurrPage
		if (FTotalCount < (FPageSize * FCurrPage)) and (FTotalCount <> 0) then
			topN = FTotalCount
		elseif (FTotalCount = 0) and (FRectOrderSerial = "") then
			FRectOrderSerial = "----------"
		end if


		'// ====================================================================
		SqlStr = "exec [db_cs].[dbo].[usp_Ten_CsAsListNew] " & CStr(topN) & ", '" & FRectDivcd & "', '" & FRectCurrstate & "', '" & FRectDeleteYN & "', '" & FRectOrderSerial & "' "
		sqlStr = sqlStr + " , '" & FRectUserID & "', '" &FRectUserName & "', '" & FRectMakerid & "', '" & FRectWriteUser & "', '" & FRectSearchType & "', '" & FRectStartDate & "', '" & FRectEndDate & "', '" & FRectExtSitename & "', '" & FRectOnlyCustomerJupsu & "', '" & FRectOnlyCSServiceRefund & "', '" + CStr(FRectCsAsID) + "', '" & FRectDateType & "', '" & FRectFinishUser & "' "

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
                FItemList(i).Fconfirmdate       = rsget("confirmdate")
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

                FItemList(i).Freceiveyn         = rsget("receiveyn")
                FItemList(i).Freceivestate		= rsget("receivestate")
                FItemList(i).Freceivefinishdate		= rsget("receivefinishdate")

                FItemList(i).Forgorderserial	= rsget("orgorderserial")

				FItemList(i).Freturnmethod		= rsget("returnmethod")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close
    end sub

    public Sub GetCSASMasterListByProcedure_3PL()
        dim i,sqlStr, AddSQL
		dim topN


		'// ====================================================================
		SqlStr = "exec [db_threepl].[dbo].[usp_Ten_CsAsCountNew_ADMIN] '" & FRectTplCompanyID & "', '" & FRectDivcd & "', '" & FRectCurrstate & "', '" & FRectDeleteYN & "', '" & FRectOrderSerial & "' "
		sqlStr = sqlStr + " , '" & FRectUserID & "', '" &FRectUserName & "', '" & FRectMakerid & "', '" & FRectWriteUser & "', '" & FRectSearchType & "', '" & FRectStartDate & "', '" & FRectEndDate & "', '" & FRectExtSitename & "', '" & FRectOnlyCustomerJupsu & "', '" & FRectOnlyCSServiceRefund & "', '" + CStr(FRectCsAsID) + "', '" & FRectDateType & "' "
		''rw "<!--" & sqlStr & "-->"

		rsget_TPL.CursorLocation = 3
		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr, dbget_TPL, 3, 1

        if  not rsget_TPL.EOF  then
            FTotalCount = rsget_TPL("cnt")
        else
            FTotalCount = 0
        end if
        rsget_TPL.close


		'// ====================================================================
		'// �ӵ����� �ذ�
		'// 1. FTotalCount ���� topN �� Ŭ�� Ǯ��ĵ ����
		'// 2. FTotalCount �� 0 �� �� �˻���� ������ �ֹ���ȣ ����
		if (FTotalCount <= FPageSize) then
			FCurrPage = 1
		end if

		topN = FPageSize * FCurrPage
		if (FTotalCount < (FPageSize * FCurrPage)) and (FTotalCount <> 0) then
			topN = FTotalCount
		elseif (FTotalCount = 0) and (FRectOrderSerial = "") then
			FRectOrderSerial = "----------"
		end if


		'// ====================================================================
		SqlStr = "exec [db_threepl].[dbo].[usp_Ten_CsAsListNew_ADMIN] '" & FRectTplCompanyID & "', " & CStr(topN) & ", '" & FRectDivcd & "', '" & FRectCurrstate & "', '" & FRectDeleteYN & "', '" & FRectOrderSerial & "' "
		sqlStr = sqlStr + " , '" & FRectUserID & "', '" &FRectUserName & "', '" & FRectMakerid & "', '" & FRectWriteUser & "', '" & FRectSearchType & "', '" & FRectStartDate & "', '" & FRectEndDate & "', '" & FRectExtSitename & "', '" & FRectOnlyCustomerJupsu & "', '" & FRectOnlyCSServiceRefund & "', '" + CStr(FRectCsAsID) + "', '" & FRectDateType & "' "
		''rw "<!--" & sqlStr & "-->"

		rsget_TPL.CursorLocation = 3
		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr, dbget_TPL, 3, 1

        FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

        redim preserve FItemList(FResultCount)
        if  not rsget_TPL.EOF  then
            i = 0
			rsget_TPL.absolutepage = FCurrPage
            do until rsget_TPL.eof
                set FItemList(i) = new CCSASMasterItem

                FItemList(i).Fid                = rsget_TPL("id")
                FItemList(i).Fdivcd             = rsget_TPL("divcd")
                FItemList(i).FdivcdName         = db2html(rsget_TPL("divcdname"))

                FItemList(i).Forderserial       = rsget_TPL("orderserial")
                FItemList(i).Fcustomername      = db2html(rsget_TPL("customername"))
                FItemList(i).Fuserid            = rsget_TPL("userid")
                FItemList(i).Fwriteuser         = rsget_TPL("writeuser")
                FItemList(i).Ffinishuser        = rsget_TPL("finishuser")
                FItemList(i).Ftitle             = db2html(rsget_TPL("title"))
                FItemList(i).Fcurrstate         = rsget_TPL("currstate")
                FItemList(i).Fcurrstatename     = rsget_TPL("currstatename")
                FItemList(i).FcurrstateColor    = rsget_TPL("currstatecolor")

                FItemList(i).Fregdate           = rsget_TPL("regdate")
                FItemList(i).Fconfirmdate       = rsget_TPL("confirmdate")
                FItemList(i).Ffinishdate        = rsget_TPL("finishdate")

                FItemList(i).Fgubun01           = rsget_TPL("gubun01")
                FItemList(i).Fgubun02           = rsget_TPL("gubun02")

                FItemList(i).Fgubun01Name       = db2html(rsget_TPL("gubun01name"))
                FItemList(i).Fgubun02Name       = db2html(rsget_TPL("gubun02name"))

                FItemList(i).Fdeleteyn          = rsget_TPL("deleteyn")

                FItemList(i).Frefundrequire     = rsget_TPL("refundrequire")
                FItemList(i).Frefundresult      = rsget_TPL("refundresult")

                FItemList(i).Fsongjangdiv       = rsget_TPL("songjangdiv")
                FItemList(i).Fsongjangno        = rsget_TPL("songjangno")

                FItemList(i).Frequireupche      = rsget_TPL("requireupche")
                FItemList(i).Fmakerid           = rsget_TPL("makerid")

                FItemList(i).FExtsitename       = rsget_TPL("sitename")

                FItemList(i).Freceiveyn         = rsget_TPL("receiveyn")
                FItemList(i).Freceivestate		= rsget_TPL("receivestate")
                FItemList(i).Freceivefinishdate	= rsget_TPL("receivefinishdate")

                FItemList(i).Forgorderserial	= rsget_TPL("orgorderserial")

				FItemList(i).Freturnmethod		= rsget_TPL("returnmethod")

                rsget_TPL.MoveNext
                i = i + 1
            loop
        end if
        rsget_TPL.close
    end sub

    public Sub GetCSASTotalPrevCancelCount()
        dim i,sqlStr

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list "
        sqlStr = sqlStr + " where 1 = 1 "

        if (FRectOrderSerial <> "") then
                sqlStr = sqlStr + " and orderserial='" + CStr(FRectOrderSerial) + "' "
        end if

        sqlStr = sqlStr + " and deleteyn='N' and divcd in ('A003','A005','A007') "
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
        sqlStr = sqlStr + " ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult, IsNULL(refminusorderserial,'') as refminusorderserial"
        sqlStr = sqlStr + " , IsNULL(A.refchangeorderserial,'') as refchangeorderserial, IsNULL(A.receiveyn,'') as receiveyn, IsNull(A.refasid, 0) as refasid "
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename"
        sqlStr = sqlStr + " , cu.addmethod as customeraddmethod, IsNull(cu.addbeasongpay, 0) as customeraddbeasongpay, cu.receiveyn as customerreceiveyn, cu.realbeasongpay as customerrealbeasongpay, cu.payorderserial, s.findurl as songjangfindurl "
		sqlStr = sqlStr + " , IsNull(cu.additempay, 0) as customeradditempay, IsNull(cu.payordertype, 'B') as customerpayordertype, IsNull(cu.additembuypay, 0) as customeradditembuypay "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A "
        sqlStr = sqlStr + " Left join [db_cs].[dbo].tbl_as_upcheAddjungsan J"
        sqlStr = sqlStr + "  on A.id=J.asid"
        sqlStr = sqlStr + " Left join [db_cs].[dbo].tbl_as_customer_addbeasongpay_info cu"
        sqlStr = sqlStr + "  on A.id=cu.asid"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"

		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_songjang_div s "
		sqlStr = sqlStr + " on a.songjangdiv = s.divcd and s.isUsing='Y' "

		if (FRectCsRefAsID <> "") then
			sqlStr = sqlStr + " where refasid= " + CStr(FRectCsRefAsID) + " "
		elseif (FRectCsAsID <> "") then
			sqlStr = sqlStr + " where id= " + CStr(FRectCsAsID) + " "
		else
			sqlStr = sqlStr + " where 1=0 "
		end if

        if (FRectMakerID<>"") then   ''��ü ��ȸ��.
            sqlStr = sqlStr + " and A.makerid='"&FRectMakerID&"'"
        end if
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
			FOneItem.FsongjangPreNo       = rsget("songjangPreNo")
			FOneItem.FsongjangRegGubun    = rsget("songjangRegGubun")
			FOneItem.FsongjangRegUserID   = rsget("songjangRegUserID")
			if Not IsNULL(FOneItem.Fsongjangno) then
				FOneItem.Fsongjangno = Replace(FOneItem.Fsongjangno, "-", "")
			end if
			FOneItem.Fsongjangfindurl         = db2html(rsget("songjangfindurl"))

            FOneItem.Frequireupche        = rsget("requireupche")
            FOneItem.Fmakerid             = rsget("makerid")

            FOneItem.Fadd_upchejungsandeliverypay = rsget("add_upchejungsandeliverypay")
            FOneItem.Fadd_upchejungsancause       = rsget("add_upchejungsancause")

			FOneItem.Frefminusorderserial 	= rsget("refminusorderserial")
			FOneItem.Frefchangeorderserial 	= rsget("refchangeorderserial")
			FOneItem.Freceiveyn 			= rsget("receiveyn")

			FOneItem.Fcustomeraddmethod 		= rsget("customeraddmethod")
			FOneItem.Fcustomeraddbeasongpay 	= rsget("customeraddbeasongpay")
			FOneItem.Fcustomerreceiveyn 		= rsget("customerreceiveyn")
			FOneItem.Fcustomerrealbeasongpay 	= rsget("customerrealbeasongpay")

			FOneItem.Frefasid 				= rsget("refasid")
			FOneItem.Fconfirmdate 			= rsget("confirmdate")

			FOneItem.FneedChkYN 			= rsget("needChkYN")
			if IsNull(FOneItem.FneedChkYN) Then
				FOneItem.FneedChkYN = ""
			End If

			FOneItem.Fpayorderserial			= rsget("payorderserial")
			FOneItem.Fcustomeradditempay		= rsget("customeradditempay")
			FOneItem.Fcustomeradditembuypay		= rsget("customeradditembuypay")
			FOneItem.Fcustomerpayordertype		= rsget("customerpayordertype")


'            FOneItem.Fbeasongdate         = rsget("beasongdate")
'            FOneItem.Frefundrequire       = rsget("refundrequire")
'            FOneItem.Frefundresult        = rsget("refundresult")

        end if
        rsget.close
    end sub

    public Sub GetOneCSASMaster_3PL()
        dim i,sqlStr

        sqlStr = " select top 1 A.*, 0 as add_upchejungsandeliverypay, '' as add_upchejungsancause "
        sqlStr = sqlStr + " ,0 as refundrequire, 0 as refundresult, IsNULL(refminusorderserial,'') as refminusorderserial"
        sqlStr = sqlStr + " , IsNULL(A.refchangeorderserial,'') as refchangeorderserial, IsNULL(A.receiveyn,'') as receiveyn, IsNull(A.refasid, 0) as refasid "
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename"
        sqlStr = sqlStr + " , '' as customeraddmethod, 0 as customeraddbeasongpay, '' as customerreceiveyn, 0 as customerrealbeasongpay, s.findurl as songjangfindurl "
        sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_as_list] A "
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
		sqlStr = sqlStr + " left join [db_threepl].[dbo].[tbl_tpl_songjang_div] s "
		sqlStr = sqlStr + " on a.songjangdiv = s.divcd and s.isUsing='Y' "

		if (FRectCsRefAsID <> "") then
			sqlStr = sqlStr + " where refasid= " + CStr(FRectCsRefAsID) + " "
		elseif (FRectCsAsID <> "") then
			sqlStr = sqlStr + " where id= " + CStr(FRectCsAsID) + " "
		else
			sqlStr = sqlStr + " where 1=0 "
		end if

        if (FRectMakerID<>"") then   ''��ü ��ȸ��.
            sqlStr = sqlStr + " and A.makerid='"&FRectMakerID&"'"
        end if
        rsget_TPL.Open sqlStr, dbget_TPL, 1

        FResultCount = rsget_TPL.RecordCount

        if  not rsget_TPL.EOF  then
            set FOneItem = new CCSASMasterItem

            FOneItem.Fid                  = rsget_TPL("id")
            FOneItem.Fdivcd               = rsget_TPL("divcd")
            FOneItem.Fgubun01             = rsget_TPL("gubun01")
            FOneItem.Fgubun02             = rsget_TPL("gubun02")

            FOneItem.FdivcdName           = db2html(rsget_TPL("divcdname"))
            FOneItem.Fgubun01Name         = db2html(rsget_TPL("gubun01name"))
            FOneItem.Fgubun02Name         = db2html(rsget_TPL("gubun02name"))

            FOneItem.Forderserial         = rsget_TPL("orderserial")
            FOneItem.Fcustomername        = db2html(rsget_TPL("customername"))
            FOneItem.Fuserid              = rsget_TPL("userid")
            FOneItem.Fwriteuser           = rsget_TPL("writeuser")
            FOneItem.Ffinishuser          = rsget_TPL("finishuser")
            FOneItem.Ftitle               = db2html(rsget_TPL("title"))
            FOneItem.Fcontents_jupsu      = db2html(rsget_TPL("contents_jupsu"))
            FOneItem.Fcontents_finish     = db2html(rsget_TPL("contents_finish"))
            FOneItem.Fcurrstate           = rsget_TPL("currstate")
            FOneItem.FcurrstateName       = rsget_TPL("currstatename")
            FOneItem.Fregdate             = rsget_TPL("regdate")
            FOneItem.Ffinishdate          = rsget_TPL("finishdate")

            FOneItem.Fdeleteyn            = rsget_TPL("deleteyn")
            FOneItem.Fextsitename         = rsget_TPL("extsitename")

            FOneItem.Fopentitle           = db2html(rsget_TPL("opentitle"))
            FOneItem.Fopencontents        = db2html(rsget_TPL("opencontents"))


            FOneItem.Fsitegubun           = rsget_TPL("sitegubun")

            FOneItem.Fsongjangdiv         = rsget_TPL("songjangdiv")
            FOneItem.Fsongjangno          = rsget_TPL("songjangno")
			FOneItem.FsongjangPreNo       = rsget_TPL("songjangPreNo")
			FOneItem.FsongjangRegGubun    = rsget_TPL("songjangRegGubun")
			FOneItem.FsongjangRegUserID   = rsget_TPL("songjangRegUserID")
			if Not IsNULL(FOneItem.Fsongjangno) then
				FOneItem.Fsongjangno = Replace(FOneItem.Fsongjangno, "-", "")
			end if
			FOneItem.Fsongjangfindurl         = db2html(rsget_TPL("songjangfindurl"))

            FOneItem.Frequireupche        = rsget_TPL("requireupche")
            FOneItem.Fmakerid             = rsget_TPL("makerid")

            FOneItem.Fadd_upchejungsandeliverypay = rsget_TPL("add_upchejungsandeliverypay")
            FOneItem.Fadd_upchejungsancause       = rsget_TPL("add_upchejungsancause")

			FOneItem.Frefminusorderserial 	= rsget_TPL("refminusorderserial")
			FOneItem.Frefchangeorderserial 	= rsget_TPL("refchangeorderserial")
			FOneItem.Freceiveyn 			= rsget_TPL("receiveyn")

			FOneItem.Fcustomeraddmethod 		= rsget_TPL("customeraddmethod")
			FOneItem.Fcustomeraddbeasongpay 	= rsget_TPL("customeraddbeasongpay")
			FOneItem.Fcustomerreceiveyn 		= rsget_TPL("customerreceiveyn")
			FOneItem.Fcustomerrealbeasongpay 	= rsget_TPL("customerrealbeasongpay")

			FOneItem.Frefasid 				= rsget_TPL("refasid")
			FOneItem.Fconfirmdate 			= rsget_TPL("confirmdate")

			FOneItem.FneedChkYN 			= rsget_TPL("needChkYN")
			if IsNull(FOneItem.FneedChkYN) Then
				FOneItem.FneedChkYN = ""
			End If

'            FOneItem.Fbeasongdate         = rsget_TPL("beasongdate")
'            FOneItem.Frefundrequire       = rsget_TPL("refundrequire")
'            FOneItem.Frefundresult        = rsget_TPL("refundresult")

        end if
        rsget_TPL.close
    end sub

    public Sub GetOneCSASMasterAcademy()
        dim i,sqlStr

        sqlStr = " select top 1 A.*, IsNULL(J.add_upchejungsandeliverypay,0) as add_upchejungsandeliverypay, J.add_upchejungsancause "
        sqlStr = sqlStr + " ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename"
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list A "
        sqlStr = sqlStr + " Left join [db_academy].[dbo].tbl_academy_as_upcheAddjungsan J"
        sqlStr = sqlStr + "  on A.id=J.asid"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"

        sqlStr = sqlStr + " where id= " + CStr(FRectCsAsID) + " "

        if (FRectMakerID<>"") then   ''��ü ��ȸ��.
            sqlStr = sqlStr + " and A.makerid='"&FRectMakerID&"'"
        end if
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1

        FResultCount = rsACADEMYget.RecordCount

        if  not rsACADEMYget.EOF  then
            set FOneItem = new CCSASMasterItem

            FOneItem.Fid                  = rsACADEMYget("id")
            FOneItem.Fdivcd               = rsACADEMYget("divcd")
            FOneItem.Fgubun01             = rsACADEMYget("gubun01")
            FOneItem.Fgubun02             = rsACADEMYget("gubun02")

            FOneItem.FdivcdName           = db2html(rsACADEMYget("divcdname"))
            FOneItem.Fgubun01Name         = db2html(rsACADEMYget("gubun01name"))
            FOneItem.Fgubun02Name         = db2html(rsACADEMYget("gubun02name"))

            FOneItem.Forderserial         = rsACADEMYget("orderserial")
            FOneItem.Fcustomername        = db2html(rsACADEMYget("customername"))
            FOneItem.Fuserid              = rsACADEMYget("userid")
            FOneItem.Fwriteuser           = rsACADEMYget("writeuser")
            FOneItem.Ffinishuser          = rsACADEMYget("finishuser")
            FOneItem.Ftitle               = db2html(rsACADEMYget("title"))
            FOneItem.Fcontents_jupsu      = db2html(rsACADEMYget("contents_jupsu"))
            FOneItem.Fcontents_finish     = db2html(rsACADEMYget("contents_finish"))
            FOneItem.Fcurrstate           = rsACADEMYget("currstate")
            FOneItem.FcurrstateName       = rsACADEMYget("currstatename")
            FOneItem.Fregdate             = rsACADEMYget("regdate")
            FOneItem.Ffinishdate          = rsACADEMYget("finishdate")

            FOneItem.Fdeleteyn            = rsACADEMYget("deleteyn")
            FOneItem.Fextsitename         = rsACADEMYget("extsitename")

            FOneItem.Fopentitle           = db2html(rsACADEMYget("opentitle"))
            FOneItem.Fopencontents        = db2html(rsACADEMYget("opencontents"))


            FOneItem.Fsitegubun           = rsACADEMYget("sitegubun")

            FOneItem.Fsongjangdiv         = rsACADEMYget("songjangdiv")
            FOneItem.Fsongjangno          = rsACADEMYget("songjangno")

            FOneItem.Frequireupche        = rsACADEMYget("requireupche")
            FOneItem.Fmakerid             = rsACADEMYget("makerid")

            FOneItem.Fadd_upchejungsandeliverypay = rsACADEMYget("add_upchejungsandeliverypay")
            FOneItem.Fadd_upchejungsancause       = rsACADEMYget("add_upchejungsancause")

'            FOneItem.Fbeasongdate         = rsACADEMYget("beasongdate")
'            FOneItem.Frefundrequire       = rsACADEMYget("refundrequire")
'            FOneItem.Frefundresult        = rsACADEMYget("refundresult")

        end if
        rsACADEMYget.close
    end sub

    'CS��������
    public Sub GetOrderDetailByCsDetailNew_eastone()
        dim SqlStr, i

		sqlStr = " select "
		sqlStr = sqlStr + " 	d.idx as orderdetailidx "
		sqlStr = sqlStr + " 	, d.orderserial "
		sqlStr = sqlStr + " 	, d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, d.itemno "
		sqlStr = sqlStr + " 	, d.itemcost "
		sqlStr = sqlStr + " 	, d.buycash "
		sqlStr = sqlStr + " 	, d.reducedprice as discountAssingedCost "
		sqlStr = sqlStr + " 	, d.mileage "
		sqlStr = sqlStr + " 	, d.cancelyn "
		sqlStr = sqlStr + " 	, d.itemname "
		sqlStr = sqlStr + " 	, d.makerid "
		sqlStr = sqlStr + " 	, d.itemoptionname "
		sqlStr = sqlStr + " 	, d.currstate as orderdetailcurrstate "
		sqlStr = sqlStr + " 	, d.upcheconfirmdate "
		sqlStr = sqlStr + " 	, d.songjangdiv "
		sqlStr = sqlStr + " 	, d.songjangno "
		sqlStr = sqlStr + " 	, d.beasongdate "
		sqlStr = sqlStr + " 	, d.isupchebeasong "
		sqlStr = sqlStr + " 	, d.issailitem "
		sqlStr = sqlStr + " 	, d.cancelyn "
		sqlStr = sqlStr + " 	, d.oitemdiv "
		sqlStr = sqlStr + " 	, d.odlvType "
		sqlStr = sqlStr + " 	, d.itemcouponidx "
		sqlStr = sqlStr + " 	, d.bonuscouponidx "
		sqlStr = sqlStr + " 	, c.id "
		sqlStr = sqlStr + " 	, c.masterid "
		sqlStr = sqlStr + " 	, IsNULL(c.orderitemno,d.itemno) as orderitemno "			'������� �ֹ�����
		sqlStr = sqlStr + " 	, IsNULL(c.regitemno,0) as regitemno "
		sqlStr = sqlStr + " 	, IsNULL(c.confirmitemno,0) as confirmitemno "
		sqlStr = sqlStr + " 	, c.gubun01 "
		sqlStr = sqlStr + " 	, c.gubun02 "
		sqlStr = sqlStr + " 	, c.regdetailstate "				'������� ��ǰ����
		sqlStr = sqlStr + " 	, C2.comm_name as gubun01name "
		sqlStr = sqlStr + " 	, C3.comm_name as gubun02name "
		sqlStr = sqlStr + " 	, i.smallimage "
		sqlStr = sqlStr + " 	, IsNull(d.orgitemcost, 0) as orgitemcost "
		sqlStr = sqlStr + " 	, IsNull(d.itemcostCouponNotApplied, 0) as itemcostCouponNotApplied "
		sqlStr = sqlStr + " 	, IsNull(d.plusSaleDiscount, 0) as plusSaleDiscount "
		sqlStr = sqlStr + " 	, IsNull(d.specialshopDiscount, 0) as specialshopDiscount "
		sqlStr = sqlStr + " 	, (IsNull(d.orgitemcost, 0)) as orgprice "

		sqlStr = sqlStr + " from "

		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_detail_2003 d "
		else
		    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
		end if

		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		d.itemid=i.itemid "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_detail c "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.masterid=" + CStr(FRectCsAsID) + " "
		sqlStr = sqlStr + " 		and c.orderdetailidx=d.idx "
		sqlStr = sqlStr + " 	Left Join [db_cs].[dbo].tbl_cs_comm_code C2 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.gubun01=C2.comm_cd "
		sqlStr = sqlStr + " 	Left Join [db_cs].[dbo].tbl_cs_comm_code C3 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.gubun02=C3.comm_cd "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	d.orderserial='" + CStr(FRectOrderSerial) + "' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.isupchebeasong, d.makerid, d.itemid, d.itemoption "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSASDetailItem

            'CS ��������
            FItemList(i).Fid              = rsget("id")
            FItemList(i).Fmasterid        = rsget("masterid")
            FItemList(i).Fgubun01         = rsget("gubun01")
            FItemList(i).Fgubun02         = rsget("gubun02")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")

            '�ֹ���ǰ����
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
            FItemList(i).Forderitemno     = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
            FItemList(i).FCancelyn        = rsget("cancelyn")
            FItemList(i).ForderDetailcurrstate = rsget("orderdetailcurrstate")
            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")
            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")

            '��ǰ����
            FItemList(i).FSmallImage      = webImgUrl + "/image/small/" + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget("gubun01name")
            FItemList(i).Fgubun02name   = rsget("gubun02name")

            FItemList(i).Forgitemcost      			= rsget("orgitemcost")
            FItemList(i).FitemcostCouponNotApplied  = rsget("itemcostCouponNotApplied")
            FItemList(i).FplusSaleDiscount      	= rsget("plusSaleDiscount")
            FItemList(i).FspecialshopDiscount      	= rsget("specialshopDiscount")
            FItemList(i).Forgprice          		= rsget("orgprice")

            FItemList(i).Fprevcsreturnfinishno      = 0

			rsget.movenext
			i=i+1
		loop
		rsget.close

        Dim bufArr
        IF (FResultCount>0) then
            '���� CS��ǰ����(�Ϸ᳻����, ��ǰ�����������)
    		sqlStr =          "		    select d.itemid, d.itemoption, sum(confirmitemno) as Preregno " + VbCrlf ''', max(a.id) asId ??
            sqlStr = sqlStr + "		    from" + VbCrlf
            sqlStr = sqlStr + "		    	[db_cs].[dbo].tbl_new_as_list a" + VbCrlf
            sqlStr = sqlStr + "		    	Join [db_cs].[dbo].tbl_new_as_detail d" + VbCrlf
            sqlStr = sqlStr + "		        on a.id=d.masterid" + VbCrlf
            sqlStr = sqlStr + "		    where a.orderserial='" + CStr(FRectOrderserial) + "'" + VbCrlf
            sqlStr = sqlStr + "		    and a.divcd in ('A004','A010')" + VbCrlf                ''��ǰ ȸ��.
            sqlStr = sqlStr + "		    and a.deleteyn='N'" + VbCrlf
            sqlStr = sqlStr + "		    and a.id <> " + CStr(FRectCsAsID) + " " + VbCrlf		'���� CS����
            'sqlStr = sqlStr + "		    	and a.currstate='B007'" + VbCrlf					'����+�Ϸ� ��� ���
            sqlStr = sqlStr + "			group by d.itemid, d.itemoption" + VbCrlf

            rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            if not rsget.Eof then
                bufArr = rsget.getRows()
            end if
            rsget.close

            if IsArray(bufArr) then

            end if
        end IF
    end Sub

    public Sub GetOrderDetailByCsDetailNew()
        dim SqlStr, i

		sqlStr = " select "
		sqlStr = sqlStr + " 	d.idx as orderdetailidx "
		sqlStr = sqlStr + " 	, d.orderserial "
		sqlStr = sqlStr + " 	, d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, d.itemno "
		sqlStr = sqlStr + " 	, d.itemcost "
		sqlStr = sqlStr + " 	, d.reducedprice "
		sqlStr = sqlStr + " 	, d.buycash "
		sqlStr = sqlStr + " 	, d.reducedprice as discountAssingedCost "
		sqlStr = sqlStr + " 	, d.mileage "
		sqlStr = sqlStr + " 	, d.cancelyn "
		sqlStr = sqlStr + " 	, d.itemname "
		sqlStr = sqlStr + " 	, Lower(d.makerid) as makerid "
		sqlStr = sqlStr + " 	, d.itemoptionname "
		sqlStr = sqlStr + " 	, d.currstate as orderdetailcurrstate "
		sqlStr = sqlStr + " 	, d.upcheconfirmdate "
		sqlStr = sqlStr + " 	, d.songjangdiv "
		sqlStr = sqlStr + " 	, d.songjangno "
		sqlStr = sqlStr + " 	, d.beasongdate "
		sqlStr = sqlStr + " 	, d.isupchebeasong "
		sqlStr = sqlStr + " 	, d.issailitem "
		sqlStr = sqlStr + " 	, d.cancelyn "
		sqlStr = sqlStr + " 	, d.oitemdiv "
		sqlStr = sqlStr + " 	, d.odlvType "
		sqlStr = sqlStr + " 	, d.itemcouponidx "
		sqlStr = sqlStr + " 	, d.bonuscouponidx "
		sqlStr = sqlStr + " 	, c.id "
		sqlStr = sqlStr + " 	, c.masterid "
		sqlStr = sqlStr + " 	, IsNULL(c.orderitemno,d.itemno) as orderitemno "			'������� �ֹ�����
		sqlStr = sqlStr + " 	, IsNULL(c.regitemno,0) as regitemno "
		sqlStr = sqlStr + " 	, IsNULL(c.confirmitemno,0) as confirmitemno "
		sqlStr = sqlStr + " 	, c.gubun01 "
		sqlStr = sqlStr + " 	, c.gubun02 "
		sqlStr = sqlStr + " 	, c.regdetailstate "				'������� ��ǰ����
		sqlStr = sqlStr + " 	, C2.comm_name as gubun01name "
		sqlStr = sqlStr + " 	, C3.comm_name as gubun02name "
		sqlStr = sqlStr + " 	, i.smallimage "
		sqlStr = sqlStr + " 	, IsNull(d.orgitemcost, 0) as orgitemcost "
		sqlStr = sqlStr + " 	, IsNull(d.itemcostCouponNotApplied, 0) as itemcostCouponNotApplied "
		sqlStr = sqlStr + " 	, IsNull(d.plusSaleDiscount, 0) as plusSaleDiscount "
		sqlStr = sqlStr + " 	, IsNull(d.specialshopDiscount, 0) as specialshopDiscount "
		sqlStr = sqlStr + " 	, IsNull(d.etcDiscount, 0) as etcDiscount "

		sqlStr = sqlStr + " 	, (i.orgprice + IsNull(o.optaddprice, 0)) as orgprice "
		sqlStr = sqlStr + " 	, IsNull(P.regno, 0) as prevcsreturnfinishno "

		sqlStr = sqlStr + " from "

		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_detail_2003 d "
		else
		    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
		end if

		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		d.itemid=i.itemid "
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item_option o "

		sqlStr = sqlStr + "     	on "
		sqlStr = sqlStr + "     		o.itemid=d.itemid and o.itemoption=d.itemoption "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_detail c "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.masterid=" + CStr(FRectCsAsID) + " "
		sqlStr = sqlStr + " 		and c.orderdetailidx=d.idx "

		'���� CS��ǰ����(����+�Ϸ᳻��, ��ǰ�����������)
		sqlStr = sqlStr + "		LEFT JOIN (" + VbCrlf
		sqlStr = sqlStr + "		    select d.orderdetailidx, sum(confirmitemno) as regno, max(a.id) asId " + VbCrlf
        sqlStr = sqlStr + "		    from" + VbCrlf
        sqlStr = sqlStr + "		    	[db_cs].[dbo].tbl_new_as_list a" + VbCrlf
        sqlStr = sqlStr + "		    	Join [db_cs].[dbo].tbl_new_as_detail d" + VbCrlf
        sqlStr = sqlStr + "		    on a.id=d.masterid" + VbCrlf
        sqlStr = sqlStr + "		    where a.orderserial='" + CStr(FRectOrderserial) + "'" + VbCrlf
        sqlStr = sqlStr + "		    and a.divcd in ('A004','A010', 'A111', 'A112')" + VbCrlf                ''��ǰ / ȸ�� / ��ǰ���� �±�ȯȸ��(�ٹ����ٹ��) / ��ǰ���� �±�ȯ��ǰ(��ü���).
        sqlStr = sqlStr + "		    and a.deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + "		    and a.id <> " + CStr(FRectCsAsID) + " " + VbCrlf		'���� CS����
        'sqlStr = sqlStr + "		    	and a.currstate='B007'" + VbCrlf					'����+�Ϸ� ��� ���

        sqlStr = sqlStr + "			group by d.orderdetailidx" + VbCrlf
        sqlStr = sqlStr + " ) P " + VbCrlf
        sqlStr = sqlStr + "     ON d.idx=P.orderdetailidx " + VbCrlf

		sqlStr = sqlStr + " 	Left Join [db_cs].[dbo].tbl_cs_comm_code C2 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.gubun01=C2.comm_cd "
		sqlStr = sqlStr + " 	Left Join [db_cs].[dbo].tbl_cs_comm_code C3 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.gubun02=C3.comm_cd "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	d.orderserial='" + CStr(FRectOrderSerial) + "' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.isupchebeasong, d.makerid, d.itemid, d.itemoption "
		''response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSASDetailItem

            'CS ��������
            FItemList(i).Fid              = rsget("id")
            FItemList(i).Fmasterid        = rsget("masterid")
            FItemList(i).Fgubun01         = rsget("gubun01")
            FItemList(i).Fgubun02         = rsget("gubun02")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")

            '�ֹ���ǰ����
            FItemList(i).Forderdetailidx  = rsget("orderdetailidx")
            FItemList(i).Forderserial     = rsget("orderserial")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
            FItemList(i).Fitemcost        = rsget("itemcost")
			FItemList(i).FreducedPrice    = rsget("reducedPrice")
            FItemList(i).Fbuycash         = rsget("buycash")
            FItemList(i).Fitemno          = rsget("itemno")
            FItemList(i).Forderitemno     = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
            FItemList(i).FCancelyn        = rsget("cancelyn")
            FItemList(i).ForderDetailcurrstate = rsget("orderdetailcurrstate")
            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")
            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")

            '��ǰ����
            FItemList(i).FSmallImage      = webImgSSLUrl + "/image/small/" + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")
            IF application("Svr_Info")="Dev" THEN
                if Not IsNull(FItemList(i).FSmallImage) then
                    FItemList(i).FSmallImage = Replace(FItemList(i).FSmallImage, "testwebimage", "webimage")
                end if
            end if

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget("gubun01name")
            FItemList(i).Fgubun02name   = rsget("gubun02name")

            FItemList(i).Forgitemcost      			= rsget("orgitemcost")
            FItemList(i).FitemcostCouponNotApplied  = rsget("itemcostCouponNotApplied")
            FItemList(i).FplusSaleDiscount      	= rsget("plusSaleDiscount")
            FItemList(i).FspecialshopDiscount      	= rsget("specialshopDiscount")
			FItemList(i).FetcDiscount		      	= rsget("etcDiscount")
            FItemList(i).Forgprice          		= rsget("orgprice")

            FItemList(i).Fprevcsreturnfinishno      = rsget("prevcsreturnfinishno")

			'// ��������
			FItemList(i).Fsongjangdiv	= rsget("songjangdiv")
			FItemList(i).Fsongjangno	= rsget("songjangno")

			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

    public Sub GetOrderDetailByCsDetailNew_3PL()
        dim SqlStr, i

		sqlStr = " select "
		sqlStr = sqlStr + " 	d.idx as orderdetailidx "
		sqlStr = sqlStr + " 	, d.orderserial "
		sqlStr = sqlStr + " 	, d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, d.itemno "
		sqlStr = sqlStr + " 	, d.itemcost "
		sqlStr = sqlStr + " 	, d.reducedprice "
		sqlStr = sqlStr + " 	, d.buycash "
		sqlStr = sqlStr + " 	, d.reducedprice as discountAssingedCost "
		sqlStr = sqlStr + " 	, d.mileage "
		sqlStr = sqlStr + " 	, d.cancelyn "
		sqlStr = sqlStr + " 	, d.itemname "
		sqlStr = sqlStr + " 	, Lower(d.makerid) as makerid "
		sqlStr = sqlStr + " 	, d.itemoptionname "
		sqlStr = sqlStr + " 	, d.currstate as orderdetailcurrstate "
		sqlStr = sqlStr + " 	, NULL as upcheconfirmdate "
		sqlStr = sqlStr + " 	, d.songjangdiv "
		sqlStr = sqlStr + " 	, d.songjangno "
		sqlStr = sqlStr + " 	, d.beasongdate "
		sqlStr = sqlStr + " 	, 'N' as isupchebeasong "
		sqlStr = sqlStr + " 	, 'N' as issailitem "
		sqlStr = sqlStr + " 	, d.cancelyn "
		sqlStr = sqlStr + " 	, '' as oitemdiv "
		sqlStr = sqlStr + " 	, '4' as odlvType "
		sqlStr = sqlStr + " 	, NULL as itemcouponidx "
		sqlStr = sqlStr + " 	, NULL as bonuscouponidx "
		sqlStr = sqlStr + " 	, c.id "
		sqlStr = sqlStr + " 	, c.masterid "
		sqlStr = sqlStr + " 	, IsNULL(c.orderitemno,d.itemno) as orderitemno "			'������� �ֹ�����
		sqlStr = sqlStr + " 	, IsNULL(c.regitemno,0) as regitemno "
		sqlStr = sqlStr + " 	, IsNULL(c.confirmitemno,0) as confirmitemno "
		sqlStr = sqlStr + " 	, c.gubun01 "
		sqlStr = sqlStr + " 	, c.gubun02 "
		sqlStr = sqlStr + " 	, c.regdetailstate "				'������� ��ǰ����
		sqlStr = sqlStr + " 	, C2.comm_name as gubun01name "
		sqlStr = sqlStr + " 	, C3.comm_name as gubun02name "
		sqlStr = sqlStr + " 	, '' as smallimage "
		sqlStr = sqlStr + " 	, IsNull(d.itemcost, 0) as orgitemcost "
		sqlStr = sqlStr + " 	, IsNull(d.itemcost, 0) as itemcostCouponNotApplied "
		sqlStr = sqlStr + " 	, 0 as plusSaleDiscount "
		sqlStr = sqlStr + " 	, 0 as specialshopDiscount "
		sqlStr = sqlStr + " 	, 0 as etcDiscount "

		sqlStr = sqlStr + " 	, d.itemcost as orgprice "
		sqlStr = sqlStr + " 	, IsNull(P.regno, 0) as prevcsreturnfinishno "

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_threepl].[dbo].[tbl_tpl_orderDetail] d "
		sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_as_detail] c "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.masterid=" + CStr(FRectCsAsID) + " "
		sqlStr = sqlStr + " 		and c.orderdetailidx=d.idx "

		'���� CS��ǰ����(����+�Ϸ᳻��, ��ǰ�����������)
		sqlStr = sqlStr + "		LEFT JOIN (" + VbCrlf
		sqlStr = sqlStr + "		    select d.orderdetailidx, sum(confirmitemno) as regno, max(a.id) asId " + VbCrlf
        sqlStr = sqlStr + "		    from" + VbCrlf
        sqlStr = sqlStr + "		    	[db_threepl].[dbo].[tbl_tpl_as_list] a" + VbCrlf
        sqlStr = sqlStr + "		    	Join [db_threepl].[dbo].[tbl_tpl_as_detail] d" + VbCrlf
        sqlStr = sqlStr + "		    on a.id=d.masterid" + VbCrlf
        sqlStr = sqlStr + "		    where a.orderserial='" + CStr(FRectOrderserial) + "'" + VbCrlf
        sqlStr = sqlStr + "		    and a.divcd in ('A004','A010', 'A111', 'A112')" + VbCrlf                ''��ǰ / ȸ�� / ��ǰ���� �±�ȯȸ��(�ٹ����ٹ��) / ��ǰ���� �±�ȯ��ǰ(��ü���).
        sqlStr = sqlStr + "		    and a.deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + "		    and a.id <> " + CStr(FRectCsAsID) + " " + VbCrlf		'���� CS����
        'sqlStr = sqlStr + "		    	and a.currstate='B007'" + VbCrlf					'����+�Ϸ� ��� ���

        sqlStr = sqlStr + "			group by d.orderdetailidx" + VbCrlf
        sqlStr = sqlStr + " ) P " + VbCrlf
        sqlStr = sqlStr + "     ON d.idx=P.orderdetailidx " + VbCrlf

		sqlStr = sqlStr + " 	Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C2 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.gubun01=C2.comm_cd "
		sqlStr = sqlStr + " 	Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C3 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.gubun02=C3.comm_cd "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	d.orderserial='" + CStr(FRectOrderSerial) + "' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.makerid, d.itemid, d.itemoption "
		''response.write sqlStr
		rsget_TPL.Open sqlStr,dbget_TPL,1

		FTotalCount = rsget_TPL.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget_TPL.eof
			set FItemList(i) = new CCSASDetailItem

            'CS ��������
            FItemList(i).Fid              = rsget_TPL("id")
            FItemList(i).Fmasterid        = rsget_TPL("masterid")
            FItemList(i).Fgubun01         = rsget_TPL("gubun01")
            FItemList(i).Fgubun02         = rsget_TPL("gubun02")
            FItemList(i).Fregitemno       = rsget_TPL("regitemno")
            FItemList(i).Fconfirmitemno   = rsget_TPL("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget_TPL("regdetailstate")

            '�ֹ���ǰ����
            FItemList(i).Forderdetailidx  = rsget_TPL("orderdetailidx")
            FItemList(i).Forderserial     = rsget_TPL("orderserial")
            FItemList(i).Fitemid          = rsget_TPL("itemid")
            FItemList(i).Fitemoption      = rsget_TPL("itemoption")
            FItemList(i).Fmakerid         = rsget_TPL("makerid")
            FItemList(i).Fitemname        = db2html(rsget_TPL("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget_TPL("itemoptionname"))
            FItemList(i).Fitemcost        = rsget_TPL("itemcost")
			FItemList(i).FreducedPrice    = rsget_TPL("reducedPrice")
            FItemList(i).Fbuycash         = rsget_TPL("buycash")
            FItemList(i).Fitemno          = rsget_TPL("itemno")
            FItemList(i).Forderitemno     = rsget_TPL("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget_TPL("isupchebeasong")
            FItemList(i).FCancelyn        = rsget_TPL("cancelyn")
            FItemList(i).ForderDetailcurrstate = rsget_TPL("orderdetailcurrstate")
            FItemList(i).FdiscountAssingedCost = rsget_TPL("discountAssingedCost")
            FItemList(i).Foitemdiv        = rsget_TPL("oitemdiv")
            FItemList(i).FodlvType        = rsget_TPL("odlvType")
            FItemList(i).Fissailitem      = rsget_TPL("issailitem")
            FItemList(i).Fitemcouponidx   = rsget_TPL("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget_TPL("bonuscouponidx")

            '��ǰ����
            FItemList(i).FSmallImage      = webImgSSLUrl + "/image/small/" + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget_TPL("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget_TPL("gubun01name")
            FItemList(i).Fgubun02name   = rsget_TPL("gubun02name")

            FItemList(i).Forgitemcost      			= rsget_TPL("orgitemcost")
            FItemList(i).FitemcostCouponNotApplied  = rsget_TPL("itemcostCouponNotApplied")
            FItemList(i).FplusSaleDiscount      	= rsget_TPL("plusSaleDiscount")
            FItemList(i).FspecialshopDiscount      	= rsget_TPL("specialshopDiscount")
			FItemList(i).FetcDiscount		      	= rsget_TPL("etcDiscount")
            FItemList(i).Forgprice          		= rsget_TPL("orgprice")

            FItemList(i).Fprevcsreturnfinishno      = rsget_TPL("prevcsreturnfinishno")

			'// ��������
			FItemList(i).Fsongjangdiv	= rsget_TPL("songjangdiv")
			FItemList(i).Fsongjangno	= rsget_TPL("songjangno")

			rsget_TPL.movenext
			i=i+1
		loop
		rsget_TPL.close

    end Sub

	'// �ٸ���ǰ �±�ȯ(A100, A111)
    public Sub GetChangeOrderDetailByCsDetailNew()
        dim SqlStr, i

		sqlStr = " select "
		sqlStr = sqlStr + " 	d.idx as orderdetailidx "
		sqlStr = sqlStr + " 	, c.orderserial "
		sqlStr = sqlStr + " 	, c.itemid "
		sqlStr = sqlStr + " 	, c.itemoption "
		sqlStr = sqlStr + " 	, IsNull(d.itemno, 0) as itemno "
		sqlStr = sqlStr + " 	, c.itemcost "
		sqlStr = sqlStr + " 	, c.buycash "
		sqlStr = sqlStr + " 	, IsNull(d.reducedprice, 0) as discountAssingedCost "
		sqlStr = sqlStr + " 	, IsNull(d.mileage, 0) as mileage "
		sqlStr = sqlStr + " 	, IsNull(d.cancelyn, 'N') as cancelyn "
		sqlStr = sqlStr + " 	, IsNull(d.itemname, c.itemname) as itemname "
		sqlStr = sqlStr + " 	, c.makerid "
		sqlStr = sqlStr + " 	, c.itemoptionname "
		sqlStr = sqlStr + " 	, IsNull(d.currstate, '2') as orderdetailcurrstate "
		sqlStr = sqlStr + " 	, d.upcheconfirmdate "
		sqlStr = sqlStr + " 	, d.songjangdiv "
		sqlStr = sqlStr + " 	, d.songjangno "
		sqlStr = sqlStr + " 	, d.beasongdate "
		sqlStr = sqlStr + " 	, c.isupchebeasong "
		sqlStr = sqlStr + " 	, d.issailitem "
		sqlStr = sqlStr + " 	, IsNull(d.cancelyn, 'N') as cancelyn "
		sqlStr = sqlStr + " 	, IsNull(d.oitemdiv, i.itemdiv) as oitemdiv "
		sqlStr = sqlStr + " 	, IsNull(d.odlvType, i.deliveryType) as odlvType "
		sqlStr = sqlStr + " 	, d.itemcouponidx "
		sqlStr = sqlStr + " 	, d.bonuscouponidx "
		sqlStr = sqlStr + " 	, c.id "
		sqlStr = sqlStr + " 	, c.masterid "
		sqlStr = sqlStr + " 	, IsNULL(c.orderitemno,d.itemno) as orderitemno "
		sqlStr = sqlStr + " 	, IsNULL(c.regitemno,0) as regitemno "
		sqlStr = sqlStr + " 	, IsNULL(c.confirmitemno,0) as confirmitemno "
		sqlStr = sqlStr + " 	, c.gubun01 "
		sqlStr = sqlStr + " 	, c.gubun02 "
		sqlStr = sqlStr + " 	, c.regdetailstate "
		sqlStr = sqlStr + " 	, C2.comm_name as gubun01name "
		sqlStr = sqlStr + " 	, C3.comm_name as gubun02name "
		sqlStr = sqlStr + " 	, i.smallimage "
		sqlStr = sqlStr + " 	, IsNull(d.orgitemcost, i.sellcash) as orgitemcost "
		sqlStr = sqlStr + " 	, IsNull(d.itemcostCouponNotApplied, 0) as itemcostCouponNotApplied "
		sqlStr = sqlStr + " 	, IsNull(d.plusSaleDiscount, 0) as plusSaleDiscount "
		sqlStr = sqlStr + " 	, IsNull(d.specialshopDiscount, 0) as specialshopDiscount "
		sqlStr = sqlStr + " 	, (i.orgprice + IsNull(o.optaddprice, 0)) as orgprice "
		sqlStr = sqlStr + " 	, 0 as prevcsreturnfinishno "
		sqlStr = sqlStr + " 	, c.reforderdetailidx "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_detail c "

		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " 	left join [db_log].[dbo].tbl_old_order_detail_2003 d "
		else
		    sqlStr = sqlStr + " 	left join [db_order].[dbo].tbl_order_detail d "
		end if

		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.orderdetailidx=d.idx "
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.itemid=i.itemid "
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item_option o "
		sqlStr = sqlStr + "     on "
		sqlStr = sqlStr + "     	o.itemid=c.itemid and o.itemoption=c.itemoption "
		sqlStr = sqlStr + " 	Left Join [db_cs].[dbo].tbl_cs_comm_code C2 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.gubun01=C2.comm_cd "
		sqlStr = sqlStr + " 	Left Join [db_cs].[dbo].tbl_cs_comm_code C3 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.gubun02=C3.comm_cd "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and c.masterid=" + CStr(FRectCsAsID) + " "
		sqlStr = sqlStr + " 	and c.orderserial='" + CStr(FRectOrderSerial) + "' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	c.isupchebeasong, c.makerid, c.itemid, c.itemoption "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSASDetailItem

            'CS ��������
            FItemList(i).Fid              = rsget("id")
            FItemList(i).Fmasterid        = rsget("masterid")
            FItemList(i).Fgubun01         = rsget("gubun01")
            FItemList(i).Fgubun02         = rsget("gubun02")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")

            FItemList(i).Freforderdetailidx  = rsget("reforderdetailidx")

            '�ֹ���ǰ����
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
            FItemList(i).Forderitemno     = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
            FItemList(i).FCancelyn        = rsget("cancelyn")
            FItemList(i).ForderDetailcurrstate = rsget("orderdetailcurrstate")
            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")
            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")

            '��ǰ����
            FItemList(i).FSmallImage      = webImgUrl + "/image/small/" + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget("gubun01name")
            FItemList(i).Fgubun02name   = rsget("gubun02name")

            FItemList(i).Forgitemcost      			= rsget("orgitemcost")
            FItemList(i).FitemcostCouponNotApplied  = rsget("itemcostCouponNotApplied")
            FItemList(i).FplusSaleDiscount      	= rsget("plusSaleDiscount")
            FItemList(i).FspecialshopDiscount      	= rsget("specialshopDiscount")
            FItemList(i).Forgprice          		= rsget("orgprice")

            FItemList(i).Fprevcsreturnfinishno      = rsget("prevcsreturnfinishno")

			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

    public Sub GetOrderDetailByCsDetail()
        dim SqlStr, i

		sqlStr = "select d.idx as orderdetailidx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost, d.buycash, d.reducedprice as discountAssingedCost"
		sqlStr = sqlStr + " ,d.mileage,d.cancelyn "
		sqlStr = sqlStr + " ,d.itemname, d.makerid, d.itemoptionname "
		sqlStr = sqlStr + " ,d.currstate as orderdetailcurrstate, d.upcheconfirmdate, d.songjangdiv, d.songjangno"
		sqlStr = sqlStr + " ,d.beasongdate, d.isupchebeasong, d.issailitem , d.cancelyn "
		sqlStr = sqlStr + " ,d.oitemdiv, d.odlvType, d.itemcouponidx, d.bonuscouponidx"
		sqlStr = sqlStr + " ,c.id, c.masterid, IsNULL(c.regitemno,0) as regitemno, IsNULL(c.confirmitemno,0) as confirmitemno"
		sqlStr = sqlStr + " ,c.gubun01, c.gubun02, c.regdetailstate"
		sqlStr = sqlStr + " ,C2.comm_name as gubun01name, C3.comm_name as gubun02name"
		sqlStr = sqlStr + " ,i.smallimage "

		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d "
		else
		    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d "
		end if
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_detail c "
		sqlStr = sqlStr + " on c.masterid=" + CStr(FRectCsAsID) + ""
		sqlStr = sqlStr + " and c.orderdetailidx=d.idx "
		sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C2"
        sqlStr = sqlStr + "  on c.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C3"
        sqlStr = sqlStr + "  on c.gubun02=C3.comm_cd"

		sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"

        sqlStr = sqlStr + " order by d.isupchebeasong, d.makerid, d.itemid, d.itemoption"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
            FItemList(i).FCancelyn        = rsget("cancelyn")
            FItemList(i).ForderDetailcurrstate = rsget("orderdetailcurrstate")

            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")

            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")


            ''���� ����ϰų�, ���ϸ���SHOP ��ǰ�� ���� �ȵǾ���.
''            if (rsget("oitemdiv")="82") or (rsget("itemcouponidx")<>0) or (rsget("issailitem")="Y") then
''                FItemList(i).FAllAtDiscountedPrice = 0
''            else
''                FItemList(i).FAllAtDiscountedPrice = round(((1-0.94) * FItemList(i).Fitemcost / 100) * 100 )
''            end if


            ''tbl_item's
            FItemList(i).FSmallImage      = webImgUrl + "/image/small/" + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

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
		sqlStr = sqlStr + " ,IsNull(d.currstate, '2') as orderdetailcurrstate"
		sqlStr = sqlStr + " ,IsNull(d.reducedprice, 0) as discountAssingedCost, IsNull(d.oitemdiv, i.itemdiv) as oitemdiv, IsNull(d.odlvType, i.deliveryType) as odlvType, d.issailitem, d.itemcouponidx, d.bonuscouponidx"
		sqlStr = sqlStr + " ,IsNULL(d.itemcost,0) as OrderItemcost"
		sqlStr = sqlStr + " ,C2.comm_name as gubun01name, C3.comm_name as gubun02name"
		sqlStr = sqlStr + " ,i.smallimage "

		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list m "
		sqlStr = sqlStr + " join [db_cs].[dbo].tbl_new_as_detail c "
		sqlStr = sqlStr + " on m.id = c.masterid "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " left join [db_log].[dbo].tbl_old_order_detail_2003 d"
		    sqlStr = sqlStr + "  on c.orderdetailidx=d.idx"
		else
		    sqlStr = sqlStr + " left join [db_order].[dbo].tbl_order_detail d"
		    sqlStr = sqlStr + "  on c.orderdetailidx=d.idx"
		end if

		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + "  on c.itemid=i.itemid"
		sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C2"
        sqlStr = sqlStr + "  on c.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C3"
        sqlStr = sqlStr + "  on c.gubun02=C3.comm_cd"

		if (FRectCsRefAsID <> "") then
			sqlStr = sqlStr + " where m.refasid=" + CStr(FRectCsRefAsID) + ""
		else
			sqlStr = sqlStr + " where c.masterid=" + CStr(FRectCsAsID) + ""
		end if

        sqlStr = sqlStr + " order by c.isupchebeasong, c.makerid, c.itemid, c.itemoption"
		'response.write sqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
            FItemList(i).Fitemno          = rsget("confirmitemno")
            FItemList(i).Forderitemno     = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")

            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")

            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")


            FItemList(i).Forderdetailcurrstate  = rsget("orderdetailcurrstate")

            FItemList(i).FSmallImage      = webImgUrl + "/image/small/" + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

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

    public Sub GetCsDetailList_3PL()
        dim SqlStr, i

		sqlStr = "select c.*"
		sqlStr = sqlStr + " ,IsNull(d.currstate, '2') as orderdetailcurrstate"
		sqlStr = sqlStr + " ,IsNull(d.reducedprice, 0) as discountAssingedCost, '' as oitemdiv, '4' as odlvType, '' as issailitem, NULL as itemcouponidx, NULL as bonuscouponidx"
		sqlStr = sqlStr + " ,IsNULL(d.itemcost,0) as OrderItemcost"
		sqlStr = sqlStr + " ,C2.comm_name as gubun01name, C3.comm_name as gubun02name"
		sqlStr = sqlStr + " ,NULL as smallimage "

		sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_as_list] m "
		sqlStr = sqlStr + " join [db_threepl].[dbo].[tbl_tpl_as_detail] c "
		sqlStr = sqlStr + " on m.id = c.masterid "
		sqlStr = sqlStr + " left join [db_threepl].[dbo].[tbl_tpl_orderDetail] d"
		sqlStr = sqlStr + "  on c.orderdetailidx=d.idx"
		sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C2"
        sqlStr = sqlStr + "  on c.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_threepl].[dbo].[tbl_tpl_cs_comm_code] C3"
        sqlStr = sqlStr + "  on c.gubun02=C3.comm_cd"

		if (FRectCsRefAsID <> "") then
			sqlStr = sqlStr + " where m.refasid=" + CStr(FRectCsRefAsID) + ""
		else
			sqlStr = sqlStr + " where c.masterid=" + CStr(FRectCsAsID) + ""
		end if

        sqlStr = sqlStr + " order by c.isupchebeasong, c.makerid, c.itemid, c.itemoption"
		'response.write sqlStr

		rsget_TPL.Open sqlStr,dbget_TPL,1

		FTotalCount = rsget_TPL.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget_TPL.eof
			set FItemList(i) = new CCSASDetailItem

            FItemList(i).Fid              = rsget_TPL("id")
            FItemList(i).Fmasterid        = rsget_TPL("masterid")
            FItemList(i).Fgubun01         = rsget_TPL("gubun01")
            FItemList(i).Fgubun02         = rsget_TPL("gubun02")
            FItemList(i).Fregitemno       = rsget_TPL("regitemno")
            FItemList(i).Fconfirmitemno   = rsget_TPL("confirmitemno")

            FItemList(i).Fregdetailstate  = rsget_TPL("regdetailstate")   ''���� ��� ���� ����
            FItemList(i).Forderdetailidx  = rsget_TPL("orderdetailidx")
            FItemList(i).Forderserial     = rsget_TPL("orderserial")
            FItemList(i).Fitemid          = rsget_TPL("itemid")
            FItemList(i).Fitemoption      = rsget_TPL("itemoption")
            FItemList(i).Fmakerid         = rsget_TPL("makerid")
            FItemList(i).Fitemname        = db2html(rsget_TPL("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget_TPL("itemoptionname"))
            FItemList(i).Fitemcost        = rsget_TPL("itemcost")
            FItemList(i).Fbuycash         = rsget_TPL("buycash")
            FItemList(i).Fitemno          = rsget_TPL("confirmitemno")
            FItemList(i).Forderitemno     = rsget_TPL("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget_TPL("isupchebeasong")

            FItemList(i).FdiscountAssingedCost = rsget_TPL("discountAssingedCost")

            FItemList(i).Foitemdiv        = rsget_TPL("oitemdiv")
            FItemList(i).FodlvType        = rsget_TPL("odlvType")
            FItemList(i).Fissailitem      = rsget_TPL("issailitem")
            FItemList(i).Fitemcouponidx   = rsget_TPL("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget_TPL("bonuscouponidx")


            FItemList(i).Forderdetailcurrstate  = rsget_TPL("orderdetailcurrstate")

            FItemList(i).FSmallImage      = webImgUrl + "/image/small/" + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget_TPL("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget_TPL("gubun01name")
            FItemList(i).Fgubun02name   = rsget_TPL("gubun02name")

            if (FItemList(i).Fitemcost=0) then
                FItemList(i).Fitemcost = rsget_TPL("OrderItemcost")
            end if

			rsget_TPL.movenext
			i=i+1
		loop
		rsget_TPL.close

    end Sub

    public Sub GetCsHistoryList()
        dim SqlStr, i

		sqlStr = "select h.* "
		sqlStr = sqlStr + " from db_log.dbo.tbl_new_as_list_history h "
		sqlStr = sqlStr + " where h.asid = " + CStr(FRectCsAsID) + ""
        sqlStr = sqlStr + " order by h.regdate "
		''response.write sqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSActUserHistoryItem

			FItemList(i).Fwriteuser		= rsget("writeuser")
			FItemList(i).Ffinishuser	= rsget("finishuser")
			FItemList(i).Fcurrstate		= rsget("currstate")
			FItemList(i).Ffinishdate	= rsget("finishdate")
			FItemList(i).Fregdate		= rsget("regdate")

			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

	'��ü��ǰ�ΰ�(���� ��ǰ������ �ִ�� ���� ��� ����)/�ٹ��ΰ�/�����ΰ�/��ۺ�� ���ΰ�
	'��ü��ǰ�� ���� CS������ �ջ��Ͽ� ����Ѵ�.
	public Sub GetOrderDetailRefundBeasongPay(byref isallrefund, byref makeridbeasongpay, byval isupbea, byval beasongmakerid, byval orderserial, byval checkidx)
	    dim sqlStr, i

		sqlStr =	      " SELECT " + VbCrlf
		sqlStr = sqlStr + " 	IsNull(SUM(CASE " + VbCrlf
		sqlStr = sqlStr + " 			WHEN ('" & isupbea & "' = 'Y') and (d.itemid <> 0) and (d.makerid = '" & beasongmakerid & "') and (d.idx not in (" & checkidx & ")) and ((d.itemno - IsNULL(P.regno,0)) > 0) THEN 1 " + VbCrlf
		sqlStr = sqlStr + " 			WHEN ('" & isupbea & "' <> 'Y') and (d.itemid <> 0) and (d.isupchebeasong <> 'Y') and (d.idx not in (" & checkidx & ")) and ((d.itemno - IsNULL(P.regno,0)) > 0) THEN 1 " + VbCrlf
		sqlStr = sqlStr + " 			else 0 " + VbCrlf
		sqlStr = sqlStr + " 		end " + VbCrlf
		sqlStr = sqlStr + " 	), 0) as remainitemcount " + VbCrlf
		sqlStr = sqlStr + " 	, IsNull(SUM(CASE " + VbCrlf
		sqlStr = sqlStr + " 			WHEN ('" & isupbea & "' = 'Y') and (d.itemid = 0) and (d.makerid = '" & beasongmakerid & "') THEN d.itemcost " + VbCrlf
		sqlStr = sqlStr + " 			WHEN ('" & isupbea & "' <> 'Y') and (d.itemid = 0) and (IsNull(d.makerid, '') = '') THEN d.itemcost " + VbCrlf
		sqlStr = sqlStr + " 			else 0 " + VbCrlf
		sqlStr = sqlStr + " 		end " + VbCrlf
		sqlStr = sqlStr + " 	), 0) as makeridbeasongpay " + VbCrlf
		sqlStr = sqlStr + " FROM " + VbCrlf
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d " + VbCrlf
		sqlStr = sqlStr + " 	LEFT JOIN ( " + VbCrlf
		sqlStr = sqlStr + " 		SELECT " + VbCrlf
		sqlStr = sqlStr + " 			d.itemid, d.itemoption, sum(confirmitemno) as regno, max(a.id) asId " + VbCrlf
		sqlStr = sqlStr + " 		FROM " + VbCrlf
		sqlStr = sqlStr + " 			[db_cs].[dbo].tbl_new_as_list a " + VbCrlf
		sqlStr = sqlStr + " 			, [db_cs].[dbo].tbl_new_as_detail d " + VbCrlf
		sqlStr = sqlStr + " 		WHERE " + VbCrlf
		sqlStr = sqlStr + " 			1 = 1 " + VbCrlf
		sqlStr = sqlStr + " 			and a.id = d.masterid " + VbCrlf
		sqlStr = sqlStr + " 			and a.orderserial = '" & orderserial & "' " + VbCrlf
		sqlStr = sqlStr + " 			and a.divcd in ('A004','A010') " + VbCrlf
		sqlStr = sqlStr + " 			and a.deleteyn = 'N' " + VbCrlf
		sqlStr = sqlStr + " 		group by " + VbCrlf
		sqlStr = sqlStr + " 			d.itemid, d.itemoption " + VbCrlf
		sqlStr = sqlStr + " 	) P " + VbCrlf
		sqlStr = sqlStr + " 	ON " + VbCrlf
		sqlStr = sqlStr + " 		1 = 1 " + VbCrlf
		sqlStr = sqlStr + " 		and d.itemid = P.itemid " + VbCrlf
		sqlStr = sqlStr + " 		and d.itemoption = P.itemoption " + VbCrlf
		sqlStr = sqlStr + " WHERE " + VbCrlf
		sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
		sqlStr = sqlStr + " 	and d.orderserial='" & orderserial & "' " + VbCrlf
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y' " + VbCrlf
		'response.write sqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		isallrefund = "N"

		makeridbeasongpay = getDefaultBeasongPayByDate(Left(Now, 10))       ' ��ۺ�

        if Not rsget.Eof then
        	if (rsget("remainitemcount") = 0) then
        		isallrefund = "Y"
        	end if

        	makeridbeasongpay = rsget("makeridbeasongpay")
		end if
		rsget.close
    end Sub

    public Sub GetCSASTotalCount()
        dim i,sqlStr

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list "
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

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        if  not rsget.EOF  then
            FResultCount = rsget("cnt")
        else
            FResultCount = 0
        end if
        rsget.close
    end sub

    public Sub GetCSASTotalCount_3PL()
        dim i,sqlStr

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_as_list] "
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

        rsget_TPL.Open sqlStr, dbget_TPL, 1

        if  not rsget_TPL.EOF  then
            FResultCount = rsget_TPL("cnt")
        else
            FResultCount = 0
        end if
        rsget_TPL.close
    end sub

    public Sub GetOneCsDeliveryItem()
        dim i,sqlStr

        if FRectCsAsID="" then exit Sub

        sqlStr = " select top 1 A.* "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_delivery A "
        sqlStr = sqlStr + " where asid= " + CStr(FRectCsAsID) + " "

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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

    public Sub GetOneCsDeliveryItem_3PL()
        dim i,sqlStr

        if FRectCsAsID="" then exit Sub

        sqlStr = " select top 1 A.* "
        sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_new_as_delivery] A "
        sqlStr = sqlStr + " where asid= " + CStr(FRectCsAsID) + " "

        rsget_TPL.Open sqlStr, dbget_TPL, 1

        FResultCount = rsget_TPL.RecordCount

        if  not rsget_TPL.EOF  then
            set FOneItem = new CCSDeliveryItem
            FOneItem.Fasid              = rsget_TPL("asid")
            FOneItem.Freqname           = db2html(rsget_TPL("reqname"))
            FOneItem.Freqphone          = rsget_TPL("reqphone")
            FOneItem.Freqhp             = rsget_TPL("reqhp")
            FOneItem.Freqzipcode        = rsget_TPL("reqzipcode")
            FOneItem.Freqzipaddr        = rsget_TPL("reqzipaddr")
            FOneItem.Freqetcaddr        = db2html(rsget_TPL("reqetcaddr"))
            FOneItem.Freqetcstr          = db2html(rsget_TPL("reqetcstr"))
            FOneItem.Fsongjangdiv       = rsget_TPL("songjangdiv")
            FOneItem.Fsongjangno        = rsget_TPL("songjangno")
            FOneItem.Fregdate           = rsget_TPL("regdate")
            FOneItem.Fsenddate          = rsget_TPL("senddate")

        end if
        rsget_TPL.close

    end Sub

    public Sub GetOneCsDeliveryItemFromDefaultOrder()
        dim i,sqlStr

        if FRectCsAsID="" then exit Sub

        sqlStr = " select m.reqname, m.reqphone, m.reqhp, m.reqzipcode, m.reqzipaddr, m.reqaddress"
        sqlStr = sqlStr + " from db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr + "     Join [db_cs].[dbo].tbl_new_as_list a"
        sqlStr = sqlStr + "     on m.orderserial=a.orderserial"
        sqlStr = sqlStr + "     and a.id=" + CStr(FRectCsAsID) + " "

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
            sqlStr = sqlStr + "     Join [db_cs].[dbo].tbl_new_as_list a"
            sqlStr = sqlStr + "     on m.orderserial=a.orderserial"
            sqlStr = sqlStr + "     and a.id=" + CStr(FRectCsAsID) + " "

            rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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

    public Sub GetOneCsDeliveryItemFromDefaultOrder_3PL()
        dim i,sqlStr

        if FRectCsAsID="" then exit Sub

        sqlStr = " select m.reqname, m.reqphone, m.reqhp, m.reqzipcode, m.reqzipaddr, m.reqaddress"
        sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_orderMaster] m"
        sqlStr = sqlStr + "     Join [db_threepl].[dbo].[tbl_tpl_as_list] a"
        sqlStr = sqlStr + "     on m.orderserial=a.orderserial"
        sqlStr = sqlStr + "     and a.id=" + CStr(FRectCsAsID) + " "

        rsget_TPL.Open sqlStr, dbget_TPL, 1
        FResultCount = rsget_TPL.RecordCount
        if  not rsget_TPL.EOF  then
            set FOneItem = new CCSDeliveryItem
            FOneItem.Fasid              = FRectCsAsID
            FOneItem.Freqname           = db2html(rsget_TPL("reqname"))
            FOneItem.Freqphone          = rsget_TPL("reqphone")
            FOneItem.Freqhp             = rsget_TPL("reqhp")
            FOneItem.Freqzipcode        = rsget_TPL("reqzipcode")
            FOneItem.Freqzipaddr        = rsget_TPL("reqzipaddr")
            FOneItem.Freqetcaddr        = db2html(rsget_TPL("reqaddress"))

        end if
        rsget_TPL.close

        if (FResultCount<1) then
            sqlStr = " select m.reqname, m.reqphone, m.reqhp, m.reqzipcode, m.reqzipaddr, m.reqaddress"
            sqlStr = sqlStr + " from db_log.dbo.tbl_old_order_master_2003 m"
            sqlStr = sqlStr + "     Join [db_cs].[dbo].tbl_new_as_list a"
            sqlStr = sqlStr + "     on m.orderserial=a.orderserial"
            sqlStr = sqlStr + "     and a.id=" + CStr(FRectCsAsID) + " "

            rsget_TPL.Open sqlStr, dbget_TPL, 1
            FResultCount = rsget_TPL.RecordCount
            if  not rsget_TPL.EOF  then
                set FOneItem = new CCSDeliveryItem
                FOneItem.Fasid              = FRectCsAsID
                FOneItem.Freqname           = db2html(rsget_TPL("reqname"))
                FOneItem.Freqphone          = rsget_TPL("reqphone")
                FOneItem.Freqhp             = rsget_TPL("reqhp")
                FOneItem.Freqzipcode        = rsget_TPL("reqzipcode")
                FOneItem.Freqzipaddr        = rsget_TPL("reqzipaddr")
                FOneItem.Freqetcaddr        = db2html(rsget_TPL("reqaddress"))

            end if
            rsget_TPL.close
        end if
    end Sub

    public sub GetOneCsConfirmItem()
        dim sqlStr, i
        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_new_as_confirm"
        sqlStr = sqlStr + " where asid=" + CStr(FRectCsAsID)
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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

    public sub GetOneCsConfirmItemAcademy()
        dim sqlStr, i
        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_as_confirm"
        sqlStr = sqlStr + " where asid=" + CStr(FRectCsAsID)

        rsACADEMYget.Open sqlStr, dbACADEMYget, 1

        FResultCount = rsACADEMYget.RecordCount

        if  not rsACADEMYget.EOF  then
            set FOneItem = new CCsConfirmItem

            FOneItem.Fasid                  = rsACADEMYget("asid")
            FOneItem.Fconfirmregmsg         = db2html(rsACADEMYget("confirmregmsg"))
            FOneItem.Fconfirmreguserid      = rsACADEMYget("confirmreguserid")
            FOneItem.Fconfirmregdate        = rsACADEMYget("confirmregdate")
            FOneItem.Fconfirmfinishmsg      = db2html(rsACADEMYget("confirmfinishmsg"))
            FOneItem.Fconfirmfinishuserid   = rsACADEMYget("confirmfinishuserid")
            FOneItem.Fconfirmfinishdate     = rsACADEMYget("confirmfinishdate")

        end if
        rsACADEMYget.close

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

function GetCurrStateName(currstate)
    dim CurrStateName

    if (currstate="B001") then
        CurrStateName = "����"
    elseif (currstate="B004") then
        CurrStateName = "������Է�"
    elseif (currstate="B005") then
        CurrStateName = "��üȮ�ο�û"
    elseif (currstate="B006") then
        CurrStateName = "��üó���Ϸ�"
    elseif (currstate="B007") then
        CurrStateName = "�Ϸ�"
    else
        CurrStateName = currstate
    end if
    GetCurrStateName=CurrStateName
end Function

%>
