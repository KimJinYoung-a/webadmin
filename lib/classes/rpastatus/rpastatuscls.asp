<%
'####################################################
' Description :  rpa ���� ���� Ŭ����
' History : 2021.07.20 ������ ����
'####################################################

'// rpa ���� ���� ���� Ŭ����
Class CrpaStatus
	Public Fidx						'// idx��
	Public Fadminid					'// ����� webadmin ���̵�(�ش� ���� ���̵� �������� nickname�� �ҷ��´�.)
	Public Fstartdate 				'// �˻��Ⱓ ������
	Public Fenddate 				'// �˻��Ⱓ ������
	Public Fstarttime 				'// �˻��Ⱓ �������� �ð�
	Public Fendtime					'// �˻��Ⱓ �������� �ð�
	Public Ftype            		'// ����, ���� Ÿ��
    '///////////////' type ���� ////////////////
    '���̹����� ���곻�� �ٿ�ε�   - ���̹�����
    '�̼��� ���ڰ�꼭 �ٿ�ε�     - �̼���
    'KICC ���γ��� �ٿ�ε�         - KICC����
    'KICC �Աݳ��� �ٿ�ε�         - KICC�Ա�
    '���޸� ���곻�� �ٿ�ε�(����) - ���޸�����
    '���޻� ���� ���� �� ����       - ���޻����
    '�������                       - �������
    'īī�� ����Ʈ �ɼ� ��� ��Ī   - īī������Ʈ�ɼ�
    '����ī�� SCM ���ε�            - ����ī��
    '����� ���ǻ��� ����           - �����
    '���޸� �ֹ� ����               - ���޸��ֹ�
    '������� ����۾�              - ���������
	Public Ftitle               	'// rpa Ÿ��Ʋ
	Public Fcontents	        	'// rpa ���� ����
	Public FisSuccess	    		'// rpa ����/���� ����(0-����, 1-����)
	Public Fregdate 				'// rpa ����ð�

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

Class CgetRpaStatus
    public FOneItem
	public FItemList()
	Public FtotalCount
	Public FRectadminid
	public FOneUser
	Public FrpaStatusList()
	Public FOneRpaStatus
	Public FRectMaxIdx
	Public FRectpagesize
	Public FRectcurrpage
	Public FResultCount
	Public FtotalPage
	Public FRectType
	Public FRectIdx
	Public FRectIsSuccess
	Public FRectStartdate
	Public FRectEnddate
	Public FRectRegUserType
	Public FRectRegUserText
    public FRectYYYYMM

	'// rpa ���� ���� view
	public Sub getRpaStatusview()
		dim sqlStr
		sqlstr = " SELECT idx, rpatype, rpatitle, rpacontents, rpaissuccess, regdate  "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_RpaSuccessMessageReceive WITH(NOLOCK) "
		sqlstr = sqlstr & " Where idx='"&FRectIdx&"' "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOneRpaStatus = new CrpaStatus
		if Not rsget.Eof Then
			FOneRpaStatus.Fidx 						= rsget("idx")
			FOneRpaStatus.Ftype 					= rsget("rpatype")
			FOneRpaStatus.Ftitle 					= rsget("rpatitle")
			FOneRpaStatus.Fcontents 				= rsget("rpacontents")
			FOneRpaStatus.FisSuccess 				= rsget("rpaissuccess")
			FOneRpaStatus.Fregdate           		= rsget("regdate")
		end if
		rsget.Close
	End Sub

	'// ��ۺ� �ݹ� �δ� ���� ����Ʈ
	public sub GetHalfDeliveryPayList()

		dim i, j, sqlStr

		sqlstr = " SELECT count(idx) "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_RpaSuccessMessageReceive WITH(NOLOCK) "
		sqlstr = sqlstr & " WHERE idx IS NOT NULL "
		If Trim(FRectType) <> "" Then
			sqlstr = sqlstr & " AND rpatype = '"&FRectType&"' "
		End If
		If Trim(FRectStartdate) <> "" Then
			sqlstr = sqlstr & " AND regdate >= '"&FRectStartdate&"' "
		End If
		If Trim(FRectEnddate) <> "" Then
			sqlstr = sqlstr & " AND regdate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		End If
		If Trim(FRectIsSuccess) <> "" Then
			sqlstr = sqlstr & " AND rpaissuccess = '"&FRectIsSuccess&"' "
		End If
		rsget.Open sqlstr, dbget, 1
			FTotalCount = rsget(0)
		rsget.close


		sqlstr = " SELECT top " & CStr(FRectcurrpage*Frectpagesize) & " idx, rpatype, rpatitle, rpacontents, rpaissuccess, regdate "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_RpaSuccessMessageReceive WITH(NOLOCK) "
		sqlstr = sqlstr & " WHERE idx IS NOT NULL "
		If Trim(FRectType) <> "" Then
			sqlstr = sqlstr & " AND rpatype = '"&FRectType&"' "
		End If
		If Trim(FRectStartdate) <> "" Then
			sqlstr = sqlstr & " AND regdate >= '"&FRectStartdate&"' "
		End If
		If Trim(FRectEnddate) <> "" Then
			sqlstr = sqlstr & " AND regdate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		End If
		If Trim(FRectIsSuccess) <> "" Then
			sqlstr = sqlstr & " AND rpaissuccess = '"&FRectIsSuccess&"' "
		End If
		sqlstr = sqlstr & " order by idx desc "

		'rw sqlstr
		rsget.pagesize = FRectpagesize
		rsget.Open sqlstr, dbget, 1

		FtotalPage = CInt(FTotalCount/FRectpagesize)
		if  (FTotalCount\FRectpagesize)<>(FTotalCount/FRectpagesize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(Frectpagesize*(FRectcurrpage-1))
        if (FResultCount<1) then FResultCount=0
		redim FrpaStatusList(FResultCount)

		i=0
		if not rsget.EOF  Then
			rsget.absolutepage = FRectcurrpage
			do until rsget.eof
				set FrpaStatusList(i) = new CrpaStatus
				FrpaStatusList(i).Fidx 						= rsget("idx")
				FrpaStatusList(i).Ftype						= rsget("rpatype")
				FrpaStatusList(i).Ftitle					= rsget("rpatitle")
				FrpaStatusList(i).Fcontents					= rsget("rpacontents")
				FrpaStatusList(i).FisSuccess				= rsget("rpaissuccess")
				FrpaStatusList(i).Fregdate					= rsget("regdate")
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

function getRpaIsSuccessName(i)
    if i="1" then
        getRpaIsSuccessName="����"
    else
        getRpaIsSuccessName="����"
    end if
end function

function getRpaTypeName(ttype)
    Select case Trim(ttype)
        Case "���̹�����"
            getRpaTypeName = "���̹����� ���곻�� �ٿ�ε�"
        Case "�̼���"
            getRpaTypeName = "�̼��� ���ڰ�꼭 �ٿ�ε�"
        Case "KICC����"
            getRpaTypeName = "KICC ���γ��� �ٿ�ε�"
        Case "KICC�Ա�"
            getRpaTypeName = "KICC �Աݳ��� �ٿ�ε�"
        Case "���޸�����"
            getRpaTypeName = "���޸� ���곻�� �ٿ�ε�(����)"
        Case "���޻����"
            getRpaTypeName = "���޻� ���� ���� �� ����"
        Case "�������"
            getRpaTypeName = "�������"
        Case "īī������Ʈ�ɼ�"
            getRpaTypeName = "īī�� ����Ʈ �ɼ� ��� ��Ī"
        Case "����ī��"
            getRpaTypeName = "����ī�� SCM ���ε�"
        Case "�����"
            getRpaTypeName = "����� ���ǻ��� ����"
        Case "���޸��ֹ�"
            getRpaTypeName = "���޸� �ֹ� ����"
        Case "���������"
            getRpaTypeName = "������� ����۾�"
    End Select
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
