<%
'###########################################################
' Description :  e����Ʈī�� Ŭ����
' History : 2011.10.04 ������ ����
'###########################################################

Class cGiftCardItem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Fidx
	public FcardItemid
	public FcardItemName
	public FcardInfo
	public FcardDesc
	public FcardSellYN
	public FlastUpdate
	public FregUserid
	public Fregdate
	public FbasicImage
	public FbasicImage600
	public FsmallImage
	public FlistImage
	public FlistImage120
	public Ficon1Image
	public Ficon2Image

	public FdesignCnt

	public FcardOption
	public FcardOptionName
	public FcardSellCash
	public FcardSalePrice
	public FcardOrgPrice
	public FoptSellYn

	public FdesignId
	public FgroupDiv
	public FcardDesignName
	public FMMSThumb
	public FMMSImage
	public FMMSText
	public FemailThumb
	public FemailImage
	public FemailText
	public FisUsing
	public FsortNo

	Public FEappIdx
	Public FReqTitle
	Public FReqContent
	Public FMakeCnt
	Public FOpt
	Public FSugiPrice
	Public FMmsTitle
	Public FMmsContent
	Public FIsSend
	Public FIsSendDate

	'// �ɼ� ���̺� ���
	public function fGiftcard_optlist()
		dim sqlStr, i, strRst
		sqlStr = "Select cardOption, cardOptionName, cardSellCash, cardSalePrice, cardOrgPrice, optSellYn " & vbCrLf &_
				" From db_item.dbo.tbl_giftCard_option " & vbCrLf &_
				" Where cardItemid=" & FcardItemid & vbCrLf &_
				"	and optIsUsing='Y'"
		rsget.Open sqlStr,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
			i=0
			strRst = "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='a' bgcolor='" & adminColor("tablebg") & "'>" & vbCrLf
			strRst = strRst & "<tr align='center' bgcolor='#F0F0F0'>"
			strRst = strRst & "<td>�ڵ�</td>"
			strRst = strRst & "<td>�ɼǸ�</td>"
			strRst = strRst & "<td>�ǸŰ�</td>"
			strRst = strRst & "<td>�Ǹ�</td>"
			strRst = strRst & "</tr>" & vbCrLf

			do until rsget.EOF
				strRst = strRst & "<tr align='center' bgcolor='#FFFFFF'>"
				strRst = strRst & "<td>" & rsget("cardOption") & "</td>"
				strRst = strRst & "<td><a href=""javascript:editGiftOpt('" & FcardItemId & "','" & rsget("cardOption") & "')"" title='����Ʈī�� �ɼǼ���'>" & rsget("cardOptionName") & "</a></td>"
				strRst = strRst & "<td>" & formatNumber(rsget("cardSellCash"),0) & "��</td>"
				strRst = strRst & "<td>" & rsget("optSellYn") & "</td>"
				strRst = strRst & "</tr>" & vbCrLf
				rsget.MoveNext
			loop
			strRst = strRst & "<tr align='right' bgcolor='#F8F8F8'>"
			strRst = strRst & "<td colspan='4'><a href=""javascript:editGiftOpt('" & FcardItemId & "')"" title='����Ʈī�� �ɼ��߰�'>[+�ɼ��߰�]</a></td>"
			strRst = strRst & "</tr>" & vbCrLf
			strRst = strRst & "</table>"
		end if

		rsget.Close

		fGiftcard_optlist = strRst
	end function

	Public Function getCardOptName
		Select Case FOpt
			Case "0001"		response.write "1������"
			Case "0002"		response.write "2������"
			Case "0003"		response.write "3������"
			Case "0004"		response.write "5������"
			Case "0005"		response.write "8������"
			Case "0006"		response.write "10������"
			Case "0007"		response.write "15������"
			Case "0008"		response.write "20������"
			Case "0009"		response.write "30������"
			Case "0000"		response.write "����("& FSugiPrice &")"
		End Select
	End Function

	'// ������ �׷�� ���
	public function fgetDesignGrpName()
		Select Case FgroupDiv
			Case "1": fgetDesignGrpName = "�⺻"
			Case "2": fgetDesignGrpName = "����"
			Case "3": fgetDesignGrpName = "����"
			Case "4": fgetDesignGrpName = "����"
			Case "5": fgetDesignGrpName = "���"
		end Select
	end Function
end class

class cGiftCard
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

	public FRectIsusing
	public FRectSellYn
	public FRectCardItemid
	public FRectCardOption
	public FRectGroupDiv
	public FRectDesignId
	Public FRectIdx

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'// ����Ʈī�� ��ǰ ���
	'/admin/giftcard/giftcard_itemList.asp
	public sub fGiftcard_Itemlist()
		dim sqlStr, addSql, i

		'�߰����� �ۼ�
		addSql = ""
		if FRectSellYn<>"" then
			addSql = addSql & " and cardSellYn='" & FRectSellYn & "'"
		end if

		if FRectCardItemid<>"" then
			addSql = addSql & " and cardItemId in (" & FRectCardItemid & ")"
		end if

		'�� ���� ���ϱ�
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(cardItemId) as cnt, CEILING(CAST(Count(cardItemId) AS FLOAT)/" & FPageSize & ") as totPg" & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_giftcard_Item" & vbcrlf
		sqlStr = sqlStr & " Where 1=1 " & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount	= rsget("cnt")
			FTotalPage	= rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'������ ����Ʈ
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " cardItemid,cardItemName,cardSellYN,smallImage " & vbcrlf
		sqlStr = sqlStr & " ,(select count(designId) from db_item.dbo.tbl_giftcard_design Where cardItemId=I.cardItemId and isUsing='Y') as dsgCnt " & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_giftcard_Item as I" & vbcrlf
		sqlStr = sqlStr & " where 1=1" & addSql

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cGiftCardItem

				FItemList(i).FcardItemid	= rsget("cardItemid")
				FItemList(i).FcardItemName	= db2html(rsget("cardItemName"))
				FItemList(i).FcardSellYN	= rsget("cardSellYN")
				FItemList(i).FsmallImage	= webImgUrl & "/giftcard/small/" & GetImageSubFolderByItemid(rsget("cardItemid")) & "/" & rsget("smallImage")
				FItemList(i).FdesignCnt		= rsget("dsgCnt")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// ����Ʈī�� ��ǰ����
	'/admin/giftcard/popEditGiftcardItem.asp
	public sub fGiftcard_oneItem()
		dim sqlStr

		'��ǰ�ڵ尡 ������ ����
		if FRectCardItemid="" then
			FResultCount = 0
			exit sub
		end if

		'������ ����Ʈ
		sqlStr = "select * " & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_giftcard_Item" & vbcrlf
		sqlStr = sqlStr & " where cardItemid=" & FRectCardItemid
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordCount

		if  not rsget.EOF  then
			set FOneItem = new cGiftCardItem

			FOneItem.FcardItemid	= rsget("cardItemid")
			FOneItem.FcardItemName	= db2html(rsget("cardItemName"))
			FOneItem.FcardInfo		= db2html(rsget("cardInfo"))
			FOneItem.FcardDesc		= db2html(rsget("cardDesc"))
			FOneItem.FcardSellYN	= rsget("cardSellYN")
			FOneItem.FbasicImage	= webImgUrl & "/giftcard/basic/" & GetImageSubFolderByItemid(rsget("cardItemid")) & "/" & rsget("basicImage")
			FOneItem.FsmallImage	= webImgUrl & "/giftcard/small/" & GetImageSubFolderByItemid(rsget("cardItemid")) & "/" & rsget("smallImage")
		end if
		rsget.Close
	end sub

	'// ����Ʈī�� �ɼ� ����
	'/admin/giftcard/popEditGiftcardOption.asp
	public sub fGiftcard_oneOption()
		dim sqlStr

		'��ǰ�ڵ�,�ɼ��ڵ尡 ������ ����
		if FRectCardItemid="" or FRectCardOption="" then
			FResultCount = 0
			exit sub
		end if

		'������ ����Ʈ
		sqlStr = "select * " & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_giftcard_Option" & vbcrlf
		sqlStr = sqlStr & " where cardItemid=" & FRectCardItemid  & vbcrlf
		sqlStr = sqlStr & " and cardOption='" & FRectCardOption & "'"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordCount

		if  not rsget.EOF  then
			set FOneItem = new cGiftCardItem

			FOneItem.FcardOption		= rsget("cardOption")
			FOneItem.FcardOptionName	= db2html(rsget("cardOptionName"))
			FOneItem.FcardSellCash		= rsget("cardSellCash")
			FOneItem.FcardSalePrice		= rsget("cardSalePrice")
			FOneItem.FcardOrgPrice		= rsget("cardOrgPrice")
			FOneItem.FoptSellYn			= rsget("optSellYn")

		end if
		rsget.Close
	end sub

	'// ����Ʈī�� ������ ���
	'/admin/giftcard/popGiftcardDesignList.asp
	public sub fGiftcard_DesignList()
		dim sqlStr, addSql, i

		'�߰����� �ۼ�
		addSql = ""
		if FRectIsusing<>"" then
			addSql = addSql & " and isUsing='" & FRectIsusing & "'"
		end if

		if FRectGroupDiv<>"" then
			addSql = addSql & " and groupDiv='" & FRectGroupDiv & "'"
		end if

		'�� ���� ���ϱ�
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(designId) as cnt, CEILING(CAST(Count(designId) AS FLOAT)/" & FPageSize & ") as totPg" & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_giftcard_design" & vbcrlf
		sqlStr = sqlStr & " Where cardItemid=" & FRectCardItemid & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount	= rsget("cnt")
			FTotalPage	= rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'������ ����Ʈ
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " designId, groupDiv, cardDesignName, MMSThumb, emailThumb, isUsing " & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_giftcard_design" & vbcrlf
		sqlStr = sqlStr & " where cardItemid=" & FRectCardItemid & addSql & vbcrlf
		sqlStr = sqlStr & " order by sortNo"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cGiftCardItem

				FItemList(i).FdesignId			= rsget("designId")
				FItemList(i).FgroupDiv			= rsget("groupDiv")
				FItemList(i).FcardDesignName	= db2html(rsget("cardDesignName"))
				FItemList(i).FMMSThumb			= webImgUrl & "/giftcard/MMS/" & GetImageSubFolderByItemid(FRectCardItemid&rsget("designId")) & "/" & rsget("MMSThumb")
				FItemList(i).FisUsing			= rsget("isUsing")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// ����Ʈī�� ����������
	'/admin/giftcard/popEditGiftcardDesign.asp
	public sub fGiftcard_oneDesign()
		dim sqlStr

		'�������ڵ尡 ������ ����
		if FRectDesignId="" then
			FResultCount = 0
			exit sub
		end if

		'������ ����Ʈ
		sqlStr = "select * " & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_giftcard_Design" & vbcrlf
		sqlStr = sqlStr & " where designid=" & FRectDesignId
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordCount

		if  not rsget.EOF  then
			set FOneItem = new cGiftCardItem

			FOneItem.FcardItemid	= rsget("cardItemid")
			FOneItem.FgroupDiv		= rsget("groupDiv")
			FOneItem.FcardDesignName= db2html(rsget("cardDesignName"))
			if Not(rsget("MMSThumb")="" or isNull(rsget("MMSThumb"))) then FOneItem.FMMSThumb		= webImgUrl & "/giftcard/MMS/" & GetImageSubFolderByItemid(rsget("cardItemid")&rsget("designId")) & "/" & rsget("MMSThumb")
			if Not(rsget("MMSImage")="" or isNull(rsget("MMSImage"))) then FOneItem.FMMSImage		= webImgUrl & "/giftcard/MMS/" & GetImageSubFolderByItemid(rsget("cardItemid")&rsget("designId")) & "/" & rsget("MMSImage")
			FOneItem.FMMSText		= db2html(rsget("MMSText"))
			if Not(rsget("emailThumb")="" or isNull(rsget("emailThumb"))) then FOneItem.FemailThumb	= webImgUrl & "/giftcard/eMail/" & GetImageSubFolderByItemid(rsget("cardItemid")&rsget("designId")) & "/" & rsget("emailThumb")
			if Not(rsget("emailImage")="" or isNull(rsget("emailImage"))) then FOneItem.FemailImage	= webImgUrl & "/giftcard/eMail/" & GetImageSubFolderByItemid(rsget("cardItemid")&rsget("designId")) & "/" & rsget("emailImage")
			FOneItem.FemailText		= db2html(rsget("emailText"))
			FOneItem.FisUsing		= rsget("isUsing")
			FOneItem.FsortNo		= rsget("sortNo")

		end if
		rsget.Close
	end sub

	Public Sub getGiftCardList
		Dim sqlStr, addSql, i

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_cs.dbo.tbl_giftcard_master "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, eappIdx, reqTitle, reqContent, makeCnt, opt, sugiPrice, mmsTitle, mmsContent, regdate, isSend, regUserId, isSendDate "
		sqlStr = sqlStr & " FROM db_cs.dbo.tbl_giftcard_master "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cGiftCardItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FEappIdx		= rsget("eappIdx")
					FItemList(i).FReqTitle		= rsget("reqTitle")
					FItemList(i).FReqContent	= rsget("reqContent")
					FItemList(i).FMakeCnt		= rsget("makeCnt")
					FItemList(i).FOpt			= rsget("opt")
					FItemList(i).FSugiPrice		= rsget("sugiPrice")
					FItemList(i).FMmsTitle		= rsget("mmsTitle")
					FItemList(i).FMmsContent	= rsget("mmsContent")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FIsSend		= rsget("isSend")
					FItemList(i).FRegUserId		= rsget("regUserId")
					FItemList(i).FIsSendDate	= rsget("isSendDate")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getGiftCardOneItem
	    Dim i, sqlStr, addSql

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, eappIdx, reqTitle, reqContent, makeCnt, opt, sugiPrice, mmsTitle, mmsContent, regdate, isSend"
		sqlStr = sqlStr & " FROM db_cs.dbo.tbl_giftcard_master "
	    sqlStr = sqlStr & " WHERE idx = '" & CStr(FRectIdx) & "'"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		set FOneItem = new cGiftCardItem
		If not rsget.EOF Then
			FOneItem.FIdx			= rsget("idx")
			FOneItem.FEappIdx		= rsget("eappIdx")
			FOneItem.FReqTitle		= rsget("reqTitle")
			FOneItem.FReqContent	= rsget("reqContent")
			FOneItem.FMakeCnt		= rsget("makeCnt")
			FOneItem.FOpt			= rsget("opt")
			FOneItem.FSugiPrice		= rsget("sugiPrice")
			FOneItem.FMmsTitle		= rsget("mmsTitle")
			FOneItem.FMmsContent	= rsget("mmsContent")
			FOneItem.FRegdate		= rsget("regdate")
			FOneItem.FIsSend		= rsget("isSend")
		End If
		rsget.Close
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

Public Function getUserids(v)
	Dim sqlStr, buf
	sqlStr = ""
	sqlStr = sqlStr & " SELECT userid FROM db_cs.dbo.tbl_giftcard_detail WHERE midx = '"& v &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		buf = ""
		Do Until rsget.EOF
			buf = buf & rsget("userid") & Chr(13)
			rsget.MoveNext
		Loop
	Else
		buf = ""
	End If
	rsget.Close
	getUserids = buf
End Function
%>