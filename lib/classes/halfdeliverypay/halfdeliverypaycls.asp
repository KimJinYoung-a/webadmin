<%
'####################################################
' Description :  ��ۺ� �ݹ� �δ� ���� Ŭ����
' History : 2020.08.27 ������ ����
'####################################################

'// ��ۺ� �ݹ� �δ� ���� Ŭ����
Class ChalfDeliveryPay
	Public Fidx						'// idx��
	Public Fadminid					'// ����� webadmin ���̵�(�ش� ���� ���̵� �������� nickname�� �ҷ��´�.)
	Public Fisusing 				'// ��뿩�� �⺻���� N
	Public Fstartdate 				'// ������
	Public Fenddate 				'// ������
	Public Fstarttime 				'// �������� �ð�
	Public Fendtime					'// �������� �ð�
	Public Fbrandid					'// �귣�� ���̵�
	Public Fdefaultdeliverytype		'// ���ǹ�ۿ���(�ش� ��ǰ�� �귣�忡 �����Ȱ�)
	Public Fdefaultfreebeasonglimit	'// �����۱��رݾ�(�ش� ��ǰ�� �귣�忡 �����Ȱ�)
	Public Fdefaultdeliverpay		'// ��ۺ�(�ش� ��ǰ�� �귣�忡 �����Ȱ�)
	Public Fhalfdeliverypay			'// ��ۺ� �δ�ݾ�(�ٹ����ٿ��� �δ��ϴ� ��ۺ�)
	Public Fregdate 				'// �����
	Public Flastupdate 				'// ������ ������(��Ͻÿ� regdate�� ���ϰ� ��.)
	Public Flastadminid 			'// ���� ������ id
	Public FItemid					'// ��ǰ���̵�
	Public Fitemname 				'// ��ǰ��
	Public FRmainimage				'// �����̹���(�߾Ⱦ�)
	Public FRlistimage				'// 100x100�̹���
	Public FRlistimage120			'// 120x120�̹���
	Public FRbasicimage				'// 400x400�̹���
	Public FRicon1image				'// 200x200�̹���
	Public FRicon2image				'// 150x150�̹���
	Public Fsmallimage				'// �̹���
	Public FItemDeliveryType		'// �ش� ��ǰ�� ��۱��а�

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CItemBeasongpayShareMasterGrpItem
	public Fmakerid
	public FmaySum
	public Ftitle
	public Ffinishflag
	public Fjgubun
	public Fjacctcd
	public Fdifferencekey
	public Fet_cnt
	public Fdlv_totalsuplycash
	public Ftotalcommission
	public Fmaydiff
	Private Sub Class_Initialize()
        ''
	End Sub

	Private Sub Class_Terminate()
        '''
	End Sub
end Class

Class CgetHalfDeliveryPay
    public FOneItem
	public FItemList()

	Public FtotalCount
	Public FRectadminid
	public FOneUser
	Public FHalfDeliveryPayList()
	Public FOneHalfDeliveryPay
	Public FRectMaxIdx
	Public FRectpagesize
	Public FRectcurrpage
	Public FResultCount
	Public FtotalPage
	Public FRectkeyword
	Public FRectIdx
	Public FRectItemId
	Public FRectItemIds
	Public FRectStartdate
	Public FRectEnddate
	Public FRectBrandId
	Public FRectIsUsing
	Public FRectItemName
	Public FRectRegUserType
	Public FRectRegUserText
    public FRectYYYYMM

	'// �ݹ� �δ㼳�� view
	public Sub getHalfDeliveryPayview()
		dim sqlStr
		sqlstr = " SELECT p.idx, p.itemid, p.brandid, p.startdate, p.enddate, c.defaultDeliveryType  "
		sqlstr = sqlstr & " , c.defaultFreeBeasongLimit, c.defaultDeliverPay, p.halfDeliveryPay "
		sqlstr = sqlstr & " , p.isusing, p.regdate, p.lastupdate, p.adminid, p.lastupdateadminid "
		sqlstr = sqlstr & " , i.itemname, i.smallimage "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_halfdeliverypay p WITH(NOLOCK) "
		sqlstr = sqlstr & " INNER JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
		sqlstr = sqlstr & " INNER JOIN db_user.dbo.tbl_user_c c WITH(NOLOCK) ON p.brandid = c.userid "
		sqlstr = sqlstr & " Where p.idx='"&FRectIdx&"' "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOneHalfDeliveryPay = new ChalfDeliveryPay
		if Not rsget.Eof Then
			FOneHalfDeliveryPay.Fidx 						= rsget("idx")
			FOneHalfDeliveryPay.Fitemid 					= rsget("itemid")
			FOneHalfDeliveryPay.Fbrandid 					= rsget("brandid")
			FOneHalfDeliveryPay.Fstartdate 					= rsget("startdate")
			FOneHalfDeliveryPay.Fenddate 					= rsget("enddate")
			FOneHalfDeliveryPay.Fdefaultdeliverytype 		= rsget("defaultDeliveryType")
			FOneHalfDeliveryPay.Fdefaultfreebeasonglimit	= rsget("defaultFreeBeasongLimit")
			FOneHalfDeliveryPay.Fdefaultdeliverpay			= rsget("defaultDeliverPay")
			FOneHalfDeliveryPay.Fhalfdeliverypay			= rsget("halfDeliveryPay")
			FOneHalfDeliveryPay.Fisusing					= rsget("isusing")
			FOneHalfDeliveryPay.Fregdate					= rsget("regdate")
			FOneHalfDeliveryPay.Flastupdate					= rsget("lastupdate")
			FOneHalfDeliveryPay.Fadminid					= rsget("adminid")
			FOneHalfDeliveryPay.Flastadminid				= rsget("lastupdateadminid")
			FOneHalfDeliveryPay.Fitemname					= rsget("itemname")
            FOneHalfDeliveryPay.Fsmallimage        			= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FOneHalfDeliveryPay.Fitemid) + "/" + rsget("smallimage")
		end if
		rsget.Close
	End Sub

	public function SearchBeasongpayShareJungsanListGrp
		dim sqlStr

		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_JungsanTarget_BeasongpayShare] '"&FRectYYYYMM&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FRectcurrpage
            do until rsget.EOF
                set FItemList(i) = new CItemBeasongpayShareMasterGrpItem

				FItemList(i).Fmakerid				= rsget("makerid")
				FItemList(i).FmaySum				= rsget("maySum")

				FItemList(i).Ftitle					= rsget("title")
				FItemList(i).Ffinishflag			= rsget("finishflag")
				FItemList(i).Fjgubun				= rsget("jgubun")
				FItemList(i).Fjacctcd				= rsget("jacctcd")
				FItemList(i).Fdifferencekey			= rsget("differencekey")
				FItemList(i).Fet_cnt				= rsget("et_cnt")
				FItemList(i).Fdlv_totalsuplycash	= rsget("dlv_totalsuplycash")
				FItemList(i).Ftotalcommission		= rsget("totalcommission")
				FItemList(i).Fmaydiff				= rsget("maydiff")


                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close

	end function

	'// ��ۺ� �ݹ� �δ� ���� ����Ʈ
	public sub GetHalfDeliveryPayList()

		dim i, j, sqlStr

		sqlstr = " SELECT count(p.idx) "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_halfDeliveryPay p WITH(NOLOCK) "
		sqlstr = sqlstr & " INNER JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
		sqlstr = sqlstr & " LEFT JOIN db_partner.dbo.tbl_user_tenbyten t WITH(NOLOCK) ON p.adminid = t.userid "
		sqlstr = sqlstr & " WHERE p.idx IS NOT NULL "
		If Trim(FRectItemIds) <> "" Then
			sqlstr = sqlstr & " AND p.itemid in ("&FRectItemIds&") "
		End If
		If Trim(FRectBrandId) <> "" Then
			sqlstr = sqlstr & " AND p.brandid = '"&brandid&"' "
		End If
		If Trim(FRectStartdate) <> "" Then
			sqlstr = sqlstr & " AND p.startdate >= '"&FRectStartdate&"' "
		End If
		If Trim(FRectEnddate) <> "" Then
			sqlstr = sqlstr & " AND p.enddate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		End If
		If Trim(FRectIsUsing) <> "" Then
			sqlstr = sqlstr & " AND p.isusing = '"&FRectIsUsing&"' "
		End If
		If Trim(FRectItemName) <> "" Then
			sqlstr = sqlstr & " AND i.itemname like '"&FRectItemName&"%' "
		End If
		If Trim(FRectRegUserText) <> "" Then
			If Trim(FRectRegUserType) = "id" Then
				sqlstr = sqlstr & " AND t.userid like '"&FRectRegUserText&"%' "
			End If
			If Trim(FRectRegUserType) = "name" Then
				sqlstr = sqlstr & " AND t.username like '"&FRectRegUserText&"%' "
			End If
		End If
		rsget.Open sqlstr, dbget, 1
			FTotalCount = rsget(0)
		rsget.close


		sqlstr = " SELECT top " & CStr(FRectcurrpage*Frectpagesize) & " p.idx, p.itemid, i.itemname, p.brandid, p.startdate, p.enddate, p.defaultdeliveryType "
		sqlstr = sqlstr & " ,p.defaultFreeBeasongLimit, p.defaultDeliverPay, p.halfDeliveryPay, p.isusing, p.regdate, p.lastupdate, p.adminid "
		sqlstr = sqlstr & " , p.lastupdateadminid, i.deliverytype "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_halfDeliveryPay p WITH(NOLOCK) "
		sqlstr = sqlstr & " INNER JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
		sqlstr = sqlstr & " LEFT JOIN db_partner.dbo.tbl_user_tenbyten t WITH(NOLOCK) ON p.adminid = t.userid "
		sqlstr = sqlstr & " WHERE idx IS NOT NULL "
		If Trim(FRectItemIds) <> "" Then
			sqlstr = sqlstr & " AND p.itemid in ("&FRectItemIds&") "
		End If
		If Trim(FRectBrandId) <> "" Then
			sqlstr = sqlstr & " AND p.brandid = '"&brandid&"' "
		End If
		If Trim(FRectStartdate) <> "" Then
			sqlstr = sqlstr & " AND p.startdate >= '"&FRectStartdate&"' "
		End If
		If Trim(FRectEnddate) <> "" Then
			sqlstr = sqlstr & " AND p.enddate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		End If
		If Trim(FRectIsUsing) <> "" Then
			sqlstr = sqlstr & " AND p.isusing = '"&FRectIsUsing&"' "
		End If
		If Trim(FRectItemName) <> "" Then
			sqlstr = sqlstr & " AND i.itemname like '"&FRectItemName&"%' "
		End If
		If Trim(FRectRegUserText) <> "" Then
			If Trim(FRectRegUserType) = "id" Then
				sqlstr = sqlstr & " AND t.userid like '"&FRectRegUserText&"%' "
			End If
			If Trim(FRectRegUserType) = "name" Then
				sqlstr = sqlstr & " AND t.username like '"&FRectRegUserText&"%' "
			End If
		End If
		sqlstr = sqlstr & " order by p.idx desc "

		'rw sqlstr
		rsget.pagesize = FRectpagesize
		rsget.Open sqlstr, dbget, 1

		FtotalPage = CInt(FTotalCount/FRectpagesize)
		if  (FTotalCount\FRectpagesize)<>(FTotalCount/FRectpagesize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(Frectpagesize*(FRectcurrpage-1))
        if (FResultCount<1) then FResultCount=0
		redim FHalfDeliveryPayList(FResultCount)

		i=0
		if not rsget.EOF  Then
			rsget.absolutepage = FRectcurrpage
			do until rsget.eof
				set FHalfDeliveryPayList(i) = new ChalfDeliveryPay
				FHalfDeliveryPayList(i).Fidx 						= rsget("idx")
				FHalfDeliveryPayList(i).FItemId						= rsget("itemid")
				FHalfDeliveryPayList(i).Fitemname					= rsget("itemname")
				FHalfDeliveryPayList(i).Fbrandid					= rsget("brandid")
				FHalfDeliveryPayList(i).Fstartdate					= rsget("startdate")
				FHalfDeliveryPayList(i).Fenddate					= rsget("enddate")
				FHalfDeliveryPayList(i).FdefaultDeliveryType		= rsget("defaultdeliveryType")
				FHalfDeliveryPayList(i).FdefaultFreeBeasongLimit	= rsget("defaultFreeBeasongLimit")
				FHalfDeliveryPayList(i).FdefaultDeliverPay			= rsget("defaultDeliverPay")
				FHalfDeliveryPayList(i).FHalfDeliveryPay			= rsget("halfDeliveryPay")
				FHalfDeliveryPayList(i).Fisusing					= rsget("isusing")
				FHalfDeliveryPayList(i).Fregdate					= rsget("regdate")
				FHalfDeliveryPayList(i).Flastupdate					= rsget("lastupdate")
				FHalfDeliveryPayList(i).Fadminid					= rsget("adminid")
				FHalfDeliveryPayList(i).Flastadminid				= rsget("lastupdateadminid")
				FHalfDeliveryPayList(i).FItemDeliveryType			= rsget("deliverytype")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub
End Class

Function LastUpdateAdmin(adid)
	dim sqlStr
	sqlstr = " Select occupation , nickname From db_sitemaster.dbo.tbl_piece_nickname Where adminid='"&adid&"' "
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		LastUpdateAdmin = rsget("occupation") &"&nbsp;"& rsget("nickname")
	Else
		LastUpdateAdmin = ""
	End If
	rsget.close
End Function

function getBeadalDivname(BeadalDiv)
    dim BeadalDivname

    if BeadalDiv="1" then
        BeadalDivname="�ٹ����ٹ��"
    elseif BeadalDiv="2" or BeadalDiv="5" then
        BeadalDivname="��ü������"
    elseif BeadalDiv="4" then
        BeadalDivname="�ٹ����ٹ�����"
    elseif BeadalDiv="5" then
        BeadalDivname="��ü������"
    elseif BeadalDiv="6" then
        BeadalDivname="�������"
    elseif BeadalDiv="7" then
        BeadalDivname="��ü���ҹ��"
    elseif BeadalDiv="9" then
        BeadalDivname="��ü���ǹ��"
    elseif BeadalDiv="" then
        BeadalDivname="�ٹ����ٹ��"
    elseif ISNULL(BeadalDiv) then
        BeadalDivname="�ٹ����ٹ��"
    else
        BeadalDivname=""
    end if
    getBeadalDivname=BeadalDivname
end function

Function fnGetMyname(adid)
	dim sqlStr
	sqlstr = " Select top 1 username from db_partner.dbo.tbl_user_tenbyten where userid = '"&adid&"'" & vbcrlf

	' ��翹���� ó��	' 2018.10.16 �ѿ��
	sqlstr = sqlstr & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf

	'response.write sqlstr & "<Br>"
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		fnGetMyname = rsget(0)
	Else
		fnGetMyname = ""
	End If
	rsget.close
End Function
%>
