<%
CLASS CCategoryPrdItem

	'// �ʼ� ����  //
	dim FItemID
	dim FItemName
	dim FSellcash
	dim FOrgPrice
	dim fEval_excludeyn
	dim FNewitem

	dim FMakerID
	dim FBrandName
	dim FBrandName_kor
	dim FBrandLogo
	dim FBrandUsing
	dim FisBestBrand
	dim FUserDiv

	dim FItemDiv
	dim FMakerName
	dim FOrgMakerID

	dim FMileage
	dim FSourceArea
	dim FDeliverytype

	dim FcdL
	dim FcdM
	dim FcdS
	dim FcateCode
	dim FCateName
	dim FcateCd1
	dim FcateCd2
	dim FcateCd3
	dim FcateDepth
	dim FarrCateCd

	dim Freviewcnt


	dim FcolorCode
	dim FcolorName

	dim FLimitNo
	dim FLimitSold
	dim fsailprice
	dim FImageBasic
	dim FImageBasic600		'600px�̹���
	dim FImageBasic1000		'1000px�̹���
	dim FImageMask
	dim FImageMask1000		'1000px�̹���
	dim FImageList
	dim FImageList120
	dim FImageSmall
	dim FImageBasicIcon
	dim FImageMaskIcon
	dim FImageIcon1	'�Ż�ǰ����Ʈ, ���θ���Ʈ���� ���(200x200)
	dim FImageIcon2
	dim FImageIcon3
	dim FImageIcon4
	dim FImageIcon5
	dim FIcon1Image
	dim FIcon2Image

	'// ���� �⺻ �̹��� �߰�(2015.01.21 ������)
	Dim Ftentenimage
	Dim Ftentenimage50
	Dim Ftentenimage200
	Dim Ftentenimage400
	Dim Ftentenimage600
	Dim Ftentenimage1000

	'// ��ǰ�󼼼��� ������ �߰�(2016.02.17 ������)
	Dim FvideoUrl
	Dim FvideoWidth
	Dim FvideoHeight
	Dim Fvideogubun
	Dim FvideoType
	Dim FvideoFullUrl


	dim FOrderComment
	dim Fdeliverarea
	dim FItemSource
	dim FItemSize
	dim FItemWeight
	dim FdeliverOverseas

	dim Fkeywords
	dim FUsingHTML
	dim FItemContent

	dim Fisusing
	dim FStreetUsing

	dim FRegDate

	dim FReipgodate
	dim FSpecialbrand


	dim Fdgncomment
	dim FDesignerComment

	dim FLimitYn
	dim FSellYn
	dim FItemScore

	dim Fitemgubun

	dim FSaleYn
	dim FTenOnlyYn		'�ٹ����� ������ǰ����(2011.04.14)

	dim FEvalcnt
	dim FEvalcnt_Photo
	dim FfavCount
	dim FQnaCnt
	dim FOptionCnt
	dim FAvgDlvDate

	dim FAddimageGubun
	dim FAddimageSmall
	dim FAddImageType
	dim FAddimage
	dim FAddimage600
	dim FAddimage1000
	dim FIsExistAddimg

	dim Ffreeprizeyn '?

	dim FReipgoitemyn
	dim FSpecialUserItem

	dim Fitemcouponyn
	dim FItemCouponType
	dim FItemCouponValue
	dim FItemCouponExpire
	dim FCurrItemCouponIdx

	dim FAvailPayType               '���� ��� ���� 0-�Ϲ� ,1-�ǽð�(������)
	dim FDefaultFreeBeasongLimit    '��ü ������۽� ��ۺ� ���� ���밪
	dim FDefaultDeliverPay		    ' ��ü ������۽� ��ۺ�
	dim FRequireMakeDay				'�ֹ����ۻ�ǰ�� ���� �ҿ���(2011.04.14)

	Dim FsafetyYN		'�����������
	Dim FsafetyDiv		'������������ '10 ~ 50
	Dim FsafetyNum	'����������ȣ

	public FPoints
	public FPoint_fun
	public FPoint_dgn
	public FPoint_prc
	public FPoint_stf
	public Fuserid
	public Fcontents
	public FImageMain
	public FImageMain2			'��ǰ����2 �̹��� �߰�(2011.04.14)
	public FImageMain3			'��ǰ����3 �̹��� �߰�(2013.07.31)
	public FlinkURL

	public FCurrRank
	public FLastRank

	public FPojangOk			'�������� ���� ����

	public FBRWriteRegdate		'����Ʈ�����
	public FUseGood
	public FUseETC

	public FplusSalePro			''��Ʈ���� ������.
	public FisJust1day			'Just 1day ��ǰ ����

	'��Ÿ�϶�������
	public FStyleCd1
	public FStyleCd1Nm
	public FStyleCd2
	public FStyleCd2Nm
	public FStyleCd3
	public FStyleCd3Nm
	public fOrderNo

	'hotcateitem 2012-04-04
	Public Fidx
	Public Fitemseq
	Public Fcdmname
	Public Fcdsname
	Public Fsailyn

	'��ǰ�� �߰� 2012-11-01
	Public FInfoname
	Public FInfoContent
	Public FinfoCode

	Public ForderMinNum
	Public ForderMaxNum

	'2013 ������ ī�װ����ο�
	Public FDisp
	Public Ftype
	Public Fcode
	Public Ftitle
	Public Fsubcopy
	Public Fimgurl
	Public Ficon

	'2013 popular wish
	Public FInCount
	Public FRegTime
	Public FEvaluate
	Public FMyCount
	
	'/�귣�� ��������
	public fdetailidx
	public fmasteridx
	public fsortNo
	public Flastupdate
	public fregadminid
	public flastadminid
	public fevt_code

	'/2014 Gift
	public FtalkCnt
	public FdayCnt
	public FthemeCnt
	
	'/��ǰ���߰�
	public FLimitDispYn
	
	public fdevice
	public Fsdate
	public Fedate

	'/2015 �� �ֹ� ��ǰ
	public Forderserial
	public ForderDate
	public ForderOption
	public ForderOptionName
	public ForderCnt

	'�귣�� ���� �߰�2017-01-31 ���¿�
	public FBrandNoticeGubun
	public FBrandNoticeTitle
	public FBrandNoticeText

	'�÷������� �ɼ� ����
	Public FOptionTypeName
	Public FOptionName
	Public FOptionAddPrice
	Public FOptionCode

	'/��Ű����
	public Frecentsellcount
	Public FAllCateName
End Class

Class ckeyItem
	Public FIdx
	Public FSubject
	Public FMode
	Public FPrekeyword
	Public FNextkeyword
	Public FEtc
	Public FRegid
	Public FRegdate
	Public FUsername

	Public FCatename
	Public FItemid
	Public FSmallimage
	Public FMakerid
	Public FItemname
	Public FKeywords
	Public FItemScore
End Class

Class cItemContent
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectSdate
	Public FRectEdate
	Public FRectMode
	Public FRectIdx
	Public FRectSearch
	Public FRectSearchstring


	Public Function fnItemcontents(byref arrItemid, byVal mxPageSize)
		Dim strSql, i
		strSql = ""
	    strSql = strSql & " SELECT TOP "&mxPageSize&" c.itemid, c.keyWords , isNULL(b.[keywordlist],'') as cateboostkeys, isNULL(k.[keywordlist],'') as addautokeys"
		strSql = strSql & " ,isNULL(nv.[nvparseKeyword],'') as nvparseKeyword"
		strSql = strSql & " , isNull(bs.sellCnt,0) as sellCnt"
		strSql = strSql & " , isNull(IA.attrNmArr,'') attrNmArr "
		strSql = strSql & " , isNull(atopt.opt_attriblist,' ') as attriblist "
		strSql = strSql & " , isNull(v.keyWords,' ') as searchkeywordslist "		
	    strSql = strSql & " FROM db_item.dbo.tbl_item_contents c WITH (NOLOCK)"
	    strSql = strSql & "     left join db_const.[dbo].[tbl_const_keyword_cate_boost_item_summary] b WITH (NOLOCK) on c.itemid=b.itemid"
	    strSql = strSql & "     left join db_const.[dbo].[tbl_const_keyword_item_summary] k WITH (NOLOCK) on c.itemid=k.itemid"
		strSql = strSql & "     left join db_const.[dbo].[tbl_const_keyword_NvMap_parse] nv WITH (NOLOCK) on c.itemid=nv.itemid"
		strSql = strSql & "     Left Join db_temp.dbo.tbl_ksearch_attrCd IA WITH (NOLOCK) on c.itemid=IA.itemid"
		strSql = strSql & "     left join db_const.[dbo].[tbl_best_sell_item] bs WITH (NOLOCK) on c.itemid =bs.itemid"
		strSql = strSql & "     left join db_temp.[dbo].[tbl_ksearch_attrListByOption] atopt  WITH (NOLOCK) on c.itemid =atopt.itemid"
		strSql = strSql & "     left join db_item.dbo.vw_item_DispCate2015 v with(nolock) on c.itemid=v.itemid"
	    strSql = strSql & " WHERE c.itemid in ("& arrItemid &") "
	    strSql = strSql & " ORDER BY c.itemid Desc "

		if (arrItemid<>"") then
			rsget.CursorLocation = adUseClient
			rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.EOF Then
				fnItemcontents = rsget.getRows()
			End If
			rsget.Close
		end if
	End Function

	Public Function fnkeywordMaster(iidx)
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 "
		sqlStr = sqlStr & "	m.idx, m.mode, m.subject, isnull(m.prekeyword, '') as prekeyword, m.nextkeyword, m.etc, m.regid, m.regdate, u.username "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_keyword_master as m "
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_user_tenbyten as u on m.regid = u.userid "
		sqlStr = sqlStr & " WHERE m.idx = '"&iidx&"' "
	    rsget.Open sqlStr,dbget,1
	    If not rsget.EOF Then
	        fnkeywordMaster = rsget.getRows()
	    End If
	    rsget.Close
	End Function

	Public Sub getKeyWordLogList
		Dim sqlStr,i, addSql
		If FRectSdate <> "" AND FRectEdate <> "" Then
			addSql = addSql & " and convert(varchar(10), m.regdate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), m.regdate, 120) <= '" & (FRectEdate) & "' "
		End If

		If FRectMode <> "" Then
			addSql = addSql & " and m.mode = '"&FRectMode&"' "
		End If

		If FRectSearch <> "" AND FRectSearchstring <> "" Then
			Select Case FRectSearch
				Case "nextkeyword"		addSql = addSql & " and m.nextkeyword = '"&FRectSearchstring&"' "
				Case "subject"			addSql = addSql & " and m.subject like '%"&FRectSearchstring&"%' "
				Case "username"			addSql = addSql & " and u.username = '"&FRectSearchstring&"' "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(m.idx) as cnt, CEILING(CAST(Count(m.idx) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_keyword_master as m "
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_user_tenbyten as u on m.regid = u.userid "
		sqlStr = sqlStr & " where 1 = 1"
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & "	m.idx, m.mode, m.subject, m.prekeyword, m.nextkeyword, m.etc, m.regid, m.regdate, u.username "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_keyword_master as m "
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_user_tenbyten as u on m.regid = u.userid "
		sqlStr = sqlStr & " where 1 = 1"
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY m.idx DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new ckeyItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FSubject		= db2html(rsget("subject"))
					FItemList(i).FMode			= rsget("mode")
					FItemList(i).FPrekeyword	= db2html(rsget("prekeyword"))
					FItemList(i).FNextkeyword	= db2html(rsget("nextkeyword"))
					FItemList(i).FEtc			= db2html(rsget("etc"))
					FItemList(i).FRegid			= rsget("regid")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FUsername		= rsget("username")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getKeyWordLogDetailList
		Dim sqlStr,i, addSql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT "
		sqlStr = sqlStr & " db_item.[dbo].[getDisplayCateName] (i.dispcate1) as catename "
		sqlStr = sqlStr & " , d.itemid, i.smallimage, i.makerid, i.itemname, c.keywords "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_keyword_detail as d "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on d.itemid = i.itemid "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_Contents as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE d.midx = "&FRectIdx&" "
		sqlStr = sqlStr & " ORDER BY d.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If (FResultCount < 1) Then FResultCount=0
		Redim preserve FItemList(FResultCount)
		i = 0
		if not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new ckeyItem
					FItemList(i).FCatename		= rsget("catename")
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FSmallimage	= rsget("smallimage")
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FItemname		= rsget("itemname")
					FItemList(i).FKeywords		= rsget("keywords")

					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If

				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount - 1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount + 1
	end Function
End Class
%>