<%
'###############################################
' PageName : startupBannerCls.asp
' Discription : APP ������ ��� ���� Ŭ����
' History : 2017.03.27 ������ : ����
'###############################################

'===============================================
'// Ŭ���� ������ ����
'===============================================

Class CStartupBannerItem
    public Fidx
    public FbannerTitle
    public FstartDate
    public FexpireDate
    public FcloseType
    public FbannerType
    public FbannerImg
    public FlinkType
    public FlinkTitle
    public FlinkURL
    public FtargetOS
    public FtargetType
    public Fimportance
    public FisUsing
    public Fstatus

	Function getLinkTypeNm()
		Select Case FlinkType
			Case "event"
				getLinkTypeNm = "�̺�Ʈ"
			Case "spevt"
				getLinkTypeNm = "��ȹ��"
			Case "prd"
				getLinkTypeNm = "��ǰ"
			Case Else
				getLinkTypeNm = ""
		End Select
	end Function

	Function getImportanceNm()
		Select Case Fimportance
			Case "10"
				getImportanceNm = "����"
			Case "30"
				getImportanceNm = "����"
			Case "50"
				getImportanceNm = "����"
			Case Else
				getImportanceNm = ""
		End Select
	end Function

	Function getTargetOSNm()
		Select Case FtargetOS
			Case "ios"
				getTargetOSNm = "IOS"
			Case "android"
				getTargetOSNm = "Android"
			Case Else
				getTargetOSNm = "��ü"
		End Select
	End Function

	Function getTargetTypeNm()
		Select Case FtargetType
			Case "30"
				getTargetTypeNm = "��ȸ��"
			Case "15"
				getTargetTypeNm = "Orange"
			Case "10"
				getTargetTypeNm = "Yellow"
			Case "11"
				getTargetTypeNm = "Green"
			Case "12"
				getTargetTypeNm = "Blue"
			Case "13"
				getTargetTypeNm = "VIP Silver"
			Case "14"
				getTargetTypeNm = "VIP Gold"
			Case "16"
				getTargetTypeNm = "VVIP"
			Case "20"
				getTargetTypeNm = "VIP��ü"
			Case Else	'00
				getTargetTypeNm = "����"
		End Select
	End Function

	Function getStatusNm()
		if FisUsing="N" or FexpireDate<date then
			getStatusNm = "����"
		Else
			Select Case Fstatus
				Case "0"
					getStatusNm = "��ϴ��"
				Case "5"
					if FstartDate>now then
						getStatusNm = "���´��"
					Else
						getStatusNm = "����"
					end if
				Case Else	'����:9
					getStatusNm = "��������"
			End Select
		end if
	end Function

	function IsExpired()
		if FisUsing="N" or FexpireDate<date then
			IsExpired = false
		else
			IsExpired = true
		end if
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class 

'===============================================
'// ���� ��� Ŭ����
'===============================================
Class CStartupBanner
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectIdx
    public FRectStartDate	'�˻� �Ⱓ ������
    public FRectEndDate		'�˻� �Ⱓ ������
    public FRectTgOS		'����̽� ����
    public FRectTgType		'Ÿ�ϱ���
    public FRectTitle		'���� �˻� (liked)
    public FRectLink		'��ũ �˻� (liked)
    public FRectStatus		'���� �˻�
    public FRectIsUsing		'��뿩��


	'# ���� ���� ��� ����
	public Sub GetOneStartupBanner()
		dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_app_startupBanner"
        SqlStr = SqlStr + " where idx=" + CStr(FRectIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CStartupBannerItem
        if Not rsget.Eof then
            FOneItem.FIdx			= rsget("idx")
            FOneItem.FbannerTitle	= rsget("bannerTitle")
            FOneItem.FstartDate		= rsget("startDate")
            FOneItem.FexpireDate	= rsget("expireDate")
            FOneItem.FcloseType		= rsget("closeType")
            FOneItem.FbannerType	= rsget("bannerType")
            FOneItem.FbannerImg		= rsget("bannerImg")
            FOneItem.FlinkType		= rsget("linkType")
            FOneItem.FlinkTitle		= rsget("linkTitle")
            FOneItem.FlinkURL		= rsget("linkURL")
            FOneItem.FtargetOS		= rsget("targetOS")
            FOneItem.FtargetType	= rsget("targetType")
            FOneItem.Fimportance	= rsget("importance")
            FOneItem.FisUsing		= rsget("isUsing")
            FOneItem.Fstatus		= rsget("status")
        end if
        rsget.close
	End Sub

    '# ���������� ���
	public Sub GetStartupBannerList()
		dim sqlStr, addSql, i

		'�߰�����
		if FRectIsUsing="A" then
			addSql = " Where m.isUsing in ('Y','N')"
		else
			addSql = " Where m.isUsing='" & FRectIsUsing & "'"
		end if

		if FRectTitle<>"" then addSql = addSql & " and m.bannerTitle like '%" & FRectTitle & "%'"
		if FRectLink<>"" then addSql = addSql & " and m.linkURL like '%" & FRectLink & "%'"

		if FRectStartDate<>"" then addSql = addSql & " and m.expireDate>'" & FRectStartDate & " 00:00:00' "
		if FRectEndDate<>"" then addSql = addSql & " and m.startDate<='" & FRectEndDate & " 23:59:59' "

		if FRectTgOS<>"" then addSql = addSql & " and m.targetOS='" & FRectTgOS & "'"
		if FRectTgType<>"" then addSql = addSql & " and m.targetType='" & FRectTgType & "'"

		if FRectStatus="9" then
			'����
			addSql = addSql & " and (m.status=9 or m.expireDate<getdate())"
		elseif FRectStatus="5" then
			'���´�� & ����
			addSql = addSql & " and (m.status=5 and m.expireDate>getdate())"
		elseif FRectStatus="0" then
			'��ϴ��
			addSql = addSql & " and m.status=0"
		end if

        '��ü ī��Ʈ
        sqlStr = "select count(m.idx), CEILING(CAST(Count(m.idx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_app_startupBanner as m "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'��� ����
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " m.* "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_app_startupBanner as m "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.idx desc"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CStartupBannerItem

	            FItemList(i).FIdx			= rsget("idx")
	            FItemList(i).FbannerTitle	= rsget("bannerTitle")
	            FItemList(i).FstartDate		= rsget("startDate")
	            FItemList(i).FexpireDate	= rsget("expireDate")
	            FItemList(i).FcloseType		= rsget("closeType")
	            FItemList(i).FbannerType	= rsget("bannerType")
	            FItemList(i).FbannerImg		= rsget("bannerImg")
	            FItemList(i).FlinkType		= rsget("linkType")
	            FItemList(i).FlinkTitle		= rsget("linkTitle")
	            FItemList(i).FlinkURL		= rsget("linkURL")
	            FItemList(i).FtargetOS		= rsget("targetOS")
	            FItemList(i).FtargetType	= rsget("targetType")
	            FItemList(i).Fimportance	= rsget("importance")
	            FItemList(i).FisUsing		= rsget("isUsing")
	            FItemList(i).Fstatus		= rsget("status")

				if FItemList(i).FbannerImg="" then
					FItemList(i).FbannerImg = "http://webadmin.10x10.co.kr/images/exclam.gif"
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub


	'------------------------------------------------
	'-- Ŭ���� �⺻���� �� ��Ÿ �Լ�
	'------------------------------------------------

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
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