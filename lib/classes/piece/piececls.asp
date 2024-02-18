<%
'####################################################
' Description :  �ǽ� Ŭ����
' History : 2017.08.31 ������ ����
'####################################################

'// PIECE ���� Ŭ����
Class Cpiece
	Public Fidx	'// �ǽ� ����Ʈ idx�� �Ǵ� nickname ���̺��� idx��
	Public Ffidx '// �ǽ� ����Ʈ ������(���ϸ��� ���ı����� �Ǿ����)
	Public Fgubun '// �ǽ� ���а�(1-����, 2-����, 3-������ǽ�(����Ʈ����), 4-�̺�Ʈ���, 5-ȸ������)
	Public Fbannergubun '// ��� ���а�(1-�ؽ�Ʈ, 2-�̹���)
	Public Fnoticeyn '// �ǽ� ������(����) �ش� �������� ��ü �ݵ�� �Ѱ��� ������.
	Public Flistimg '// �ǽ� ����Ʈ �̹���
	Public Flisttext '// �ǽ� ����
	Public Fshorttext '// ���¸�??
	Public Flisttitle '// �ǽ� ����
	Public Fadminid '// ����� scm ���� ���̵�(�ش� ���� ���̵� �������� nickname�� �ҷ��´�.)
	Public Fusertype '// 1-������, 2-����(�⺻���� 1�̸� ���� ������ ����� ���� 2�� ��ϵȴ�.)
	Public Fetclink '// ��Ÿ��ũ??
	Public Fsnsbtncnt '// ������ư Ŭ�� ī��Ʈ
	Public Fitemid '// �ش� �ǽ��� ��ϵ� ��ǰ���̵�(�迭���·� ��)
	Public Fpieceidx '// ���� ��������
	Public Fisusing '// ��뿩�� �⺻���� N
	Public Fstartdate '// �ش� �ǽ� ������
	Public Fenddate '// �ش� �ǽ� ������
	Public Fregdate '// �ش� �ǽ� �����
	Public Flastupdate '// �ش� �ǽ� ������ ������(��Ͻÿ� regdate�� ���ϰ� ��.)
	Public FDeleteYn	'// �ش� �ǽ� ��������(Y-����, N-�����ƴ�)
	Public FTagText		'// �±��Է°�
	Public FRItemid		'// ������ǰ ��ǰ���̵�
	Public FRitemname '// ������ǰ ��ǰ��
	Public FRisusing	'// �ش��ǰ ��뿩��
	Public FRlimitno	'// �ش��ǰ ��������
	Public FRlimitsold	'// �ش��ǰ �����ǸŰ���
	Public FRmainimage	'// �����̹���(�߾Ⱦ�)
	Public FRlistimage	'// 100x100�̹���
	Public FRlistimage120	'// 120x120�̹���
	Public FRbasicimage	'// 400x400�̹���
	Public FRicon1image	'// 200x200�̹���
	Public FRicon2image	'// 150x150�̹���
	Public FRsellyn	'// �Ǹſ���
	Public FRlimityn	'// ������ǰ����
	Public Fadmintext '// �۾��� ���û���
	Public Fstate '// �������
	Public Flastadminid '// ���� ������ id

	Public Foccupation '// �г��� ����
	Public Fnickname
	Public Flastoccupation
	Public Flastnickname

	Public FSellcash
	Public FOrgprice
	Public FSaleYN
	Public FItemcouponYN
	Public FItemcouponValue
	Public FitemcouponType

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


Class Cgetpiece

	Public FtotalCount
	Public FRectadminid
	public FOneUser
	Public FPieceList()
	Public FRectMaxIdx
	Public FOneOpening
	Public FOnePiece
	Public FRectpagesize
	Public FRectcurrpage
	Public FResultCount
	Public FtotalPage
	Public FRectDeal
	Public FRectOpen
	Public FRectkeyword
	Public FRectSchword
	Public FRectIdx
	Public FRectArrItemId
	Public FRelationItemlist()
	Public FRectState
	Public FRectStartdate

	public Sub adminPieceUser()
		dim sqlStr
		sqlstr = " Select * From db_sitemaster.dbo.tbl_piece_nickname Where adminid='"&FRectadminid&"' "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOneUser = new Cpiece

		if Not rsget.Eof Then
			FOneUser.Fidx			= rsget("idx")
			FOneUser.Foccupation		= rsget("occupation")
			FOneUser.Fnickname		= rsget("nickname")
			FOneUser.FRegdate	= rsget("regdate")
			FOneUser.Flastupdate	= rsget("lastupdate")
		end if

		rsget.Close

	End Sub


	'// �ǽ� ������ ������
	public Sub getPieceOpening()
		dim sqlStr
		sqlstr = " Select top 1 p.idx, p.fidx, p.gubun, p.bannergubun, p.noticeYN, p.listimg, p.listtext, p.shorttext, p.listtitle, p.adminid,  "
		sqlstr = sqlstr & " p.usertype, p.etclink, p.snsbtncnt, p.itemid, p.pieceidx, p.isusing, p.startdate, p.enddate, p.regdate, p.lastupdate, n.nickname, n.occupation, p.deleteyn, "
		sqlstr = sqlstr & "	stuff "
		sqlstr = sqlstr & "	( "
		sqlstr = sqlstr & "		( "
		sqlstr = sqlstr & "			Select ','+t.tagtext "
		sqlstr = sqlstr & "			From db_sitemaster.dbo.tbl_piece_tag t "
		sqlstr = sqlstr & "			Where t.pidx = p.idx "
		sqlstr = sqlstr & "			for XML PATH ('') "
		sqlstr = sqlstr & "		), 1, 1,'' "
		sqlstr = sqlstr & "	) as TagText , nn.occupation as lastoccupation  , nn.nickname as lastnickname  , p.state"
		sqlstr = sqlstr & " From db_sitemaster.dbo.tbl_piece p "
		sqlstr = sqlstr & " inner join db_sitemaster.dbo.tbl_piece_nickname n on p.adminid = n.adminid "
		sqlstr = sqlstr & " outer apply (select occupation , nickname from db_sitemaster.[dbo].[tbl_piece_nickname] with (nolock) where adminid = p.lastadminid) as nn "
		sqlstr = sqlstr & " Where p.deleteyn='N' And p.noticeYN='Y' "
		sqlstr = sqlstr & " order by p.idx desc "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOneOpening = new Cpiece

		if Not rsget.Eof Then
			FOneOpening.Fidx = rsget("idx")
			FOneOpening.Ffidx = rsget("fidx")
			FOneOpening.Fgubun = rsget("gubun")
			FOneOpening.Fbannergubun = rsget("bannergubun")
			FOneOpening.Fnoticeyn = rsget("noticeYN")
			FOneOpening.Flistimg = rsget("listimg")
			FOneOpening.Flisttext = rsget("listtext")
			FOneOpening.Fshorttext = rsget("shorttext")
			FOneOpening.Flisttitle = rsget("listtitle")
			FOneOpening.Fadminid = rsget("adminid")
			FOneOpening.Fusertype = rsget("usertype")
			FOneOpening.Fetclink = rsget("etclink")
			FOneOpening.Fsnsbtncnt = rsget("snsbtncnt")
			FOneOpening.Fitemid = rsget("itemid")
			FOneOpening.Fpieceidx = rsget("pieceidx")
			FOneOpening.Fisusing = rsget("isusing")
			FOneOpening.Fstartdate = rsget("startdate")
			FOneOpening.Fenddate = rsget("enddate")
			FOneOpening.Fregdate = rsget("regdate")
			FOneOpening.Flastupdate = rsget("lastupdate")
			FOneOpening.FDeleteYn = rsget("deleteyn")
			FOneOpening.Foccupation = rsget("occupation")
			FOneOpening.Fnickname = rsget("nickname")
			FOneOpening.FTagText = rsget("tagtext")
			FOneOpening.Flastoccupation = rsget("lastoccupation")
			FOneOpening.Flastnickname = rsget("lastnickname")
			FOneOpening.Fstate = rsget("state")
		end if
		rsget.Close
	End Sub

	'// �ǽ� ������ view
	public Sub getPieceview()
		dim sqlStr
		sqlstr = " Select top 1 p.idx, p.fidx, p.gubun,p.bannergubun, p.noticeYN, p.listimg, p.listtext, p.shorttext, p.listtitle, p.adminid,  "
		sqlstr = sqlstr & " p.usertype, p.etclink, p.snsbtncnt, p.itemid, p.pieceidx, p.isusing, p.startdate, p.enddate, p.regdate, p.lastupdate, n.nickname, n.occupation, p.deleteyn, "
		sqlstr = sqlstr & "	stuff "
		sqlstr = sqlstr & "	( "
		sqlstr = sqlstr & "		( "
		sqlstr = sqlstr & "			Select ','+t.tagtext "
		sqlstr = sqlstr & "			From db_sitemaster.dbo.tbl_piece_tag t "
		sqlstr = sqlstr & "			Where t.pidx = p.idx "
		sqlstr = sqlstr & "			for XML PATH ('') "
		sqlstr = sqlstr & "		), 1, 1,'' "
		sqlstr = sqlstr & "	) as TagText "
		sqlstr = sqlstr & "	, admintext , state , lastadminid"
		sqlstr = sqlstr & " From db_sitemaster.dbo.tbl_piece p "
		sqlstr = sqlstr & " inner join db_sitemaster.dbo.tbl_piece_nickname n on p.adminid = n.adminid "
		sqlstr = sqlstr & " Where p.deleteyn='N' And p.idx='"&FRectIdx&"' "
		sqlstr = sqlstr & " order by p.idx desc "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOnePiece = new Cpiece

		if Not rsget.Eof Then
			FOnePiece.Fidx = rsget("idx")
			FOnePiece.Ffidx = rsget("fidx")
			FOnePiece.Fgubun = rsget("gubun")
			FOnePiece.Fbannergubun = rsget("bannergubun")
			FOnePiece.Fnoticeyn = rsget("noticeYN")
			FOnePiece.Flistimg = rsget("listimg")
			FOnePiece.Flisttext = rsget("listtext")
			FOnePiece.Fshorttext = rsget("shorttext")
			FOnePiece.Flisttitle = rsget("listtitle")
			FOnePiece.Fadminid = rsget("adminid")
			FOnePiece.Fusertype = rsget("usertype")
			FOnePiece.Fetclink = rsget("etclink")
			FOnePiece.Fsnsbtncnt = rsget("snsbtncnt")
			FOnePiece.Fitemid = rsget("itemid")
			FOnePiece.Fpieceidx = rsget("pieceidx")
			FOnePiece.Fisusing = rsget("isusing")
			FOnePiece.Fstartdate = rsget("startdate")
			FOnePiece.Fenddate = rsget("enddate")
			FOnePiece.Fregdate = rsget("regdate")
			FOnePiece.Flastupdate = rsget("lastupdate")
			FOnePiece.FDeleteYn = rsget("deleteyn")
			FOnePiece.Foccupation = rsget("occupation")
			FOnePiece.Fnickname = rsget("nickname")
			FOnePiece.FTagText = rsget("tagtext")
			FOnePiece.Fadmintext = rsget("admintext")
			FOnePiece.Fstate = rsget("state")
			FOnePiece.Flastadminid = rsget("lastadminid")
		end if
		rsget.Close
	End Sub

	'// �ǽ� ����Ʈ
	public sub GetpieceList()

		dim i, j, sqlStr

		sqlstr = " Select count(p.idx) "
		sqlstr = sqlstr & " From db_sitemaster.[dbo].[tbl_piece] p "
		sqlstr = sqlstr & " inner join db_sitemaster.[dbo].[tbl_piece_nickname] n on p.adminid = n.adminid "
		sqlstr = sqlstr & " Where noticeYN='N' And Deleteyn='N' "
		If Trim(FRectDeal)<>"0" Then
			If Trim(FRectDeal) <> "" Then
				sqlstr = sqlstr & " And p.gubun='"&Trim(FRectDeal)&"' "
			End If
		End If
		If Trim(FRectOpen) <> "A" Then
			If Trim(FRectOpen) <> "" Then
				sqlstr = sqlstr & " And p.isusing='"&FRectOpen&"' "
			End If
		End If
		If Trim(FRectSchword) <> "" Then
			If Trim(FRectkeyword) = "snum" Then
				sqlstr = sqlstr & " And p.idx='"&FRectSchword&"' "
			End If
			If Trim(FRectkeyword) = "sname" Then
				sqlstr = sqlstr & " And n.nickname like '%"&FRectSchword&"%' "
			End If
			If Trim(FRectkeyword) = "stitle" Then
				If Trim(FRectDeal)="1" Then
					sqlstr = sqlstr & " And p.listtext like '%"&FRectSchword&"%' "
				ElseIf Trim(FRectDeal)="0" Then
					sqlstr = sqlstr & " And (p.listtext like '%"&FRectSchword&"%' or p.listtitle like '%"&FRectSchword&"%') "
				Else
					sqlstr = sqlstr & " And p.listtitle like '%"&FRectSchword&"%' "
				End If
			End If
		End If

		If FRectState <> "" Then
			sqlstr = sqlstr & " And p.state='"& FRectState &"' "
		End If

		If FRectStartdate <> "" Then
			sqlstr = sqlstr & " And convert(varchar(10),p.startdate,120) ='"& FRectStartdate &"' "
		End If

		rsget.Open sqlstr, dbget, 1
			FTotalCount = rsget(0)
		rsget.close

		sqlstr = " Select top " & CStr(FRectcurrpage*Frectpagesize) & " p.idx, p.fidx, p.gubun, p.bannergubun, p.noticeYN, p.listimg, p.listtext, p.shorttext, p.listtitle, p.adminid "
		sqlstr = sqlstr & " , p.usertype, p.etclink, p.snsbtncnt, p.itemid, p.pieceidx "
		sqlstr = sqlstr & " , p.isusing, p.startdate, p.enddate, p.regdate, p.lastupdate, n.nickname, n.occupation, p.deleteyn , nn.occupation as lastoccupation , nn.nickname as lastnickname , p.state"
		sqlstr = sqlstr & " From db_sitemaster.[dbo].[tbl_piece] p "
		sqlstr = sqlstr & " inner join db_sitemaster.[dbo].[tbl_piece_nickname] n on p.adminid = n.adminid "
		sqlstr = sqlstr & " outer apply (select occupation , nickname from db_sitemaster.[dbo].[tbl_piece_nickname] with (nolock) where adminid = p.lastadminid) as nn "
		sqlstr = sqlstr & " Where noticeYN='N' And Deleteyn='N' "
		If Trim(FRectDeal)<>"0" Then
			If Trim(FRectDeal) <> "" Then
				sqlstr = sqlstr & " And p.gubun='"&Trim(FRectDeal)&"' "
			End If
		End If
		If Trim(FRectOpen) <> "A" Then
			If Trim(FRectOpen) <> "" Then
				sqlstr = sqlstr & " And p.isusing='"&FRectOpen&"' "
			End If
		End If
		If Trim(FRectSchword) <> "" Then
			If Trim(FRectkeyword) = "snum" Then
				sqlstr = sqlstr & " And p.idx='"&FRectSchword&"' "
			End If
			If Trim(FRectkeyword) = "sname" Then
				sqlstr = sqlstr & " And n.nickname like '%"&FRectSchword&"%' "
			End If
			If Trim(FRectkeyword) = "stitle" Then
				If Trim(FRectDeal)="1" Then
					sqlstr = sqlstr & " And p.listtext like '%"&FRectSchword&"%' "
				ElseIf Trim(FRectDeal)="0" Then
					sqlstr = sqlstr & " And (p.listtext like '%"&FRectSchword&"%' or p.listtitle like '%"&FRectSchword&"%') "
				Else
					sqlstr = sqlstr & " And p.listtitle like '%"&FRectSchword&"%' "
				End If
			End If
		End If

		If FRectState <> "" Then
			sqlstr = sqlstr & " And p.state='"& FRectState &"' "
		End If

		If FRectStartdate <> "" Then
			sqlstr = sqlstr & " And convert(varchar(10),p.startdate,120) ='"& FRectStartdate &"' "
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
		redim FPieceList(FResultCount)

		i=0
		if not rsget.EOF  Then
			rsget.absolutepage = FRectcurrpage
			do until rsget.eof
				set FPieceList(i) = new Cpiece
				FPieceList(i).Fidx = rsget("idx")
				FPieceList(i).Ffidx = rsget("fidx")
				FPieceList(i).Fgubun = rsget("gubun")
				FPieceList(i).Fbannergubun = rsget("bannergubun")
				FPieceList(i).Fnoticeyn = rsget("noticeYN")
				FPieceList(i).Flistimg = rsget("listimg")
				FPieceList(i).Flisttext = rsget("listtext")
				FPieceList(i).Fshorttext = rsget("shorttext")
				FPieceList(i).Flisttitle = rsget("listtitle")
				FPieceList(i).Fadminid = rsget("adminid")
				FPieceList(i).Fusertype = rsget("usertype")
				FPieceList(i).Fetclink = rsget("etclink")
				FPieceList(i).Fsnsbtncnt = rsget("snsbtncnt")
				FPieceList(i).Fitemid = rsget("itemid")
				FPieceList(i).Fpieceidx = rsget("pieceidx")
				FPieceList(i).Fisusing = rsget("isusing")
				FPieceList(i).Fstartdate = rsget("startdate")
				FPieceList(i).Fenddate = rsget("enddate")
				FPieceList(i).Fregdate = rsget("regdate")
				FPieceList(i).Flastupdate = rsget("lastupdate")
				FPieceList(i).FDeleteYn = rsget("deleteyn")
				FPieceList(i).Foccupation = rsget("occupation")
				FPieceList(i).Fnickname = rsget("nickname")
				FPieceList(i).Flastoccupation = rsget("lastoccupation")
				FPieceList(i).Flastnickname = rsget("lastnickname")
				FPieceList(i).Fstate = rsget("state")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

	'// ������ǰ ��������
	public sub GetRelationItemList()
		dim i, j, sqlStr
		sqlstr = " Select i.itemid, i.itemname, i.isusing, i.limitno, i.limitsold, i.mainimage, i.listimage, i.listimage120, i.basicimage, i.icon1image, i.icon2image, i.sellyn, i.limityn , i.sellCash , i.orgPrice , i.sailyn , i.itemcouponYn , i.itemcouponvalue , i.itemcoupontype "
		sqlstr = sqlstr & " From db_item.[dbo].[tbl_item] i"
		sqlstr = sqlstr & " inner join db_sitemaster.dbo.tbl_piece_item p on i.itemid = p.itemid "
		sqlstr = sqlstr & " Where p.pidx='"&FRectIdx&"' "
		sqlstr = sqlstr & " order by idx "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		redim FRelationItemlist(FResultCount)
		i=0
		if not rsget.EOF  Then
			do until rsget.eof
				set FRelationItemlist(i) = new Cpiece
				FRelationItemlist(i).FRItemid = rsget("itemid")
				FRelationItemlist(i).FRitemname = rsget("itemname")
				FRelationItemlist(i).FRisusing = rsget("isusing")
				FRelationItemlist(i).FRlimitno = rsget("limitno")
				FRelationItemlist(i).FRlimitsold = rsget("limitsold")
				FRelationItemlist(i).FRmainimage = webImgUrl&"/image/main/"&GetImageSubFolderByItemid(FRelationItemlist(i).FRItemid)&"/"&rsget("mainimage")
				FRelationItemlist(i).FRlistimage = webImgUrl&"/image/list/"&GetImageSubFolderByItemid(FRelationItemlist(i).FRItemid)&"/"&rsget("listimage")
				FRelationItemlist(i).FRlistimage120 = webImgUrl&"/image/list120/"&GetImageSubFolderByItemid(FRelationItemlist(i).FRItemid)&"/"&rsget("listimage120")
				FRelationItemlist(i).FRbasicimage = webImgUrl&"/image/basic/"&GetImageSubFolderByItemid(FRelationItemlist(i).FRItemid)&"/"&rsget("basicimage")
				FRelationItemlist(i).FRicon1image = webImgUrl&"/image/icon1/"&GetImageSubFolderByItemid(FRelationItemlist(i).FRItemid)&"/"&rsget("icon1image")
				FRelationItemlist(i).FRicon2image = webImgUrl&"/image/icon2/"&GetImageSubFolderByItemid(FRelationItemlist(i).FRItemid)&"/"&rsget("icon2image")
				FRelationItemlist(i).FRsellyn = rsget("sellyn")
				FRelationItemlist(i).FRlimityn = rsget("limityn")
				FRelationItemlist(i).FSellcash        = rsget("sellCash")
				FRelationItemlist(i).FOrgprice        = rsget("orgPrice")
				FRelationItemlist(i).FSaleYN          = rsget("sailyn")
				FRelationItemlist(i).FItemcouponYN    = rsget("itemcouponYn")
				FRelationItemlist(i).FItemcouponValue = rsget("itemcouponvalue")
				FRelationItemlist(i).FitemcouponType  = rsget("itemcoupontype")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	End Sub

End Class

'// ���� ���� ���� ��������
Function pieceMyCnt(adid)
	dim sqlStr
	sqlstr = " Select count(idx) From db_sitemaster.dbo.tbl_piece Where adminid='"&adid&"' And DeleteYn = 'N' "
	rsget.Open SqlStr, dbget, 1
	pieceMyCnt = rsget(0)
	rsget.close
End Function

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

Function nowstatus(v)
	Select Case v
		Case "1"	: nowstatus = "��ϴ��"
		Case "2"	: nowstatus = "�̹���<br/>��Ͽ�û"
		Case "3"	: nowstatus = "������<br/>�۾���"
		Case "4"	: nowstatus = "���¿�û"
		Case "7"	: nowstatus = "����"
		Case "8"	: nowstatus = "����"
		Case "9"	: nowstatus = "����"
		Case Else	: nowstatus = ""
	End Select
End Function

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

'// �������� // ������ ����
Function fnGetLastPrice(sellCash , orgPrice , sailyn , itemcouponYn , itemcouponvalue , itemcoupontype)
	Dim lastprice , saleper

	If sailYN = "N" and itemcouponYn = "N" Then
		lastprice = ""&formatNumber(orgPrice,0) &""
	End If
	If sailYN = "Y" and itemcouponYn = "N" Then
		lastprice = ""&formatNumber(sellCash,0) &""
	End If
	if itemcouponYn = "Y" And itemcouponvalue>0 Then
		If itemcoupontype = "1" Then
		lastprice = ""&formatNumber(sellCash - CLng(itemcouponvalue*sellCash/100),0) &""
		ElseIf itemcoupontype = "2" Then
		lastprice = ""&formatNumber(sellCash - itemcouponvalue,0) &""
		ElseIf itemcoupontype = "3" Then
		lastprice = ""&formatNumber(sellCash,0) &""
		Else
		lastprice = ""&formatNumber(sellCash,0) &""
		End If
	End If

	If sailYN = "Y" And itemcouponYn = "Y" Then
		If itemcoupontype = "1" Then
			'//���� + %����
			saleper = ""& CLng((orgPrice-(sellCash - CLng(itemcouponvalue*sellCash/100)))/orgPrice*100)&"%"
		ElseIf itemcoupontype = "2" Then
			'//���� + ������
			saleper = ""& CLng((orgPrice-(sellCash - itemcouponvalue))/orgPrice*100)&"%"
		Else
			'//���� + ��������
			saleper = ""& CLng((orgPrice-sellCash)/orgPrice*100)&"%"
		End If
	ElseIf sailYN = "Y" and itemcouponYn = "N" Then
		If CLng((orgPrice-sellCash)/orgPrice*100)> 0 Then
			saleper = ""& CLng((orgPrice-sellCash)/orgPrice*100)&"%"
		End If
	elseif sailYN = "N" And itemcouponYn = "Y" And itemcouponvalue>0 Then
		If itemcoupontype = "1" Then
			saleper = ""&  CStr(itemcouponvalue) & "%"
		End If
	Else
		saleper = ""
	End If

	fnGetLastPrice = "<strong class='"& chkiif(saleper<>"","cRd1","")&"'>"& lastprice &"�� "&saleper&"</strong>"
End Function
%>
