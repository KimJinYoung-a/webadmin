<%

class CItemQnaSubItem

	public FID
	public FUserid
	public Fmakerid
	public Fcdl
	public Fusername
	public FTitle
	public FContents
	public FReplytitle
	public FReplycontents
	public FReplyuser
	public Fregdate
	public FBrandName
	public Freplydate
	public FQadiv

	public FItemID
	public Flistimage
	public FItemName
	public FSellcash

	public FUserLevel
	public Fusermail
	public Fdeliverytype

	public FItemDiv

	public function GetDeliveryTypeName()
		if Fdeliverytype="1" or Fdeliverytype="3" or Fdeliverytype="4" then
			GetDeliveryTypeName = "10x10"
		else
			GetDeliveryTypeName = "업체"
		end if
	end function

	public function GetDeliveryTypeColor()
		if Fdeliverytype="1" or Fdeliverytype="3" or Fdeliverytype="4" then
			GetDeliveryTypeColor = "#000000"
		else
			GetDeliveryTypeColor = "#CC3333"
		end if
	end function

	public function GetItemDivNameName()
		if FItemDiv="90" then
			GetItemDivNameName = "강좌"
		else
			GetItemDivNameName = " "
		end if
	end function

	public function GetQaName()
		if FQadiv="01" then
			GetQaName="강좌문의"
		elseif FQadiv="02" then
			GetQaName="재료문의"
		elseif FQadiv="04" then
			GetQaName="강좌 대기자 요청"
		elseif FQadiv="07" then
			GetQaName="DIY재료 판매문의"
		elseif FQadiv="20" then
			GetQaName="기타 문의"
		elseif FQadiv="99" then
			GetQaName="DIY 상품 문의"
		end if
	end function

	public function GetUserLevelStr()

		if Fuserlevel="1" then
			GetUserLevelStr = "<font color=#f0ca2c>Green</font>"
		elseif Fuserlevel="2" then
			GetUserLevelStr = "<font color=#a3cf6c>BLUE</font>"
		elseif Fuserlevel="3" then
			GetUserLevelStr = "<font color=#6ca54e>VIP</font>"
		elseif Fuserlevel="4" then
			GetUserLevelStr = "<font color=#f68d3f>오렌지</font>"
		elseif Fuserlevel="5" then
			GetUserLevelStr = "<font color=#865e25>옐로우</font>"
		elseif Fuserlevel="6" then
			GetUserLevelStr = "<font color=#B70606>STAFF</font>"
		else
			GetUserLevelStr = CStr(Fuserlevel)
		end if

	end function

	public function IsReplyOk()
		if IsNULL(Freplydate) then
			IsReplyOk = false
		else
			IsReplyOk = true
		end if
	end function

	public function ReplyYN()
		if IsNULL(Freplydate) then
			ReplyYN = "단변대기"
		else
			ReplyYN = "답변완료"
		end if
	end function

	public function ReplyColor()
		if IsNULL(Freplydate) then
			ReplyColor = "#0066FF"
		else
			ReplyColor = "#C80708"
		end if
	end function

	Private Sub Class_Terminate()

	End Sub

	public sub Class_Initialize()

	end sub

end Class

Class CItemQna
	public FItemList()
	public FResultCount
	public FPageSize
	public FCurrpage
	public FTotalCount
	public FTotalPage
	public FScrollCount

	public FRectItemID
	public FRectCDL

	public FRectId
	public FOneItem

	public FReckMiFinish
	public FRectOnlyTenBeasong
	public FRectMakerid
	public frectcdm
	public frectcds
	public frectuserid
	public frectstartdate
	public frectenddate

	public CItemDiv

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FResultCount  = 0
		FTotalCount = 0
		FPageSize = 20
		FCurrpage = 1
		FScrollCount = 10
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub CategoryMainItemQna()
		dim sql,i

		sql = "select top " + CStr(FPageSize) + " q.id,q.userid,q.itemid,q.makerid,'' as cdl,q.contents,"
		sql = sql + " q.replyuser,q.replycontents,q.regdate, '' as brandname, q.replydate, "
		sql = sql + " i.listimage"
		sql = sql + " from db_academy.dbo.tbl_diy_item_qna q" + vbcrlf
		sql = sql + " left join db_academy.dbo.tbl_diy_item i on q.itemid=i.itemid"
		'sql = sql + " where q.cdl = '" + Cstr(FRectCDL) + "'" + vbcrlf
		sql = sql + " where 1 = 1 " + vbcrlf
		sql = sql + " and q.isusing ='Y'" + vbcrlf
		sql = sql + " and q.replydate is not null" + vbcrlf
		sql = sql + " order by replydate desc"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(i) = new CItemQnaSubItem

				FItemList(i).FID = rsACADEMYget("id")
				FItemList(i).FUserid = rsACADEMYget("userid")
				FItemList(i).Fmakerid = rsACADEMYget("makerid")
				FItemList(i).FCdl = rsACADEMYget("cdl")
				FItemList(i).FContents = db2html(rsACADEMYget("contents"))
				FItemList(i).FTitle = DdotFormat(FItemList(i).FContents,40)

				FItemList(i).FReplyuser = rsACADEMYget("replyuser")

				FItemList(i).FReplycontents = db2html(rsACADEMYget("replycontents"))
				FItemList(i).FReplytitle = DdotFormat(FItemList(i).FReplycontents,40)
				FItemList(i).Fregdate = rsACADEMYget("regdate")
				FItemList(i).FBrandName = db2html(rsACADEMYget("brandname"))
				FItemList(i).Freplydate = rsACADEMYget("replydate")
				FItemList(i).Fitemid = rsACADEMYget("itemid")

				FItemList(i).Flistimage = "http://image.thefingers.co.kr/diyitem/webimage/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end sub

	public sub ItemQnaList()
		dim sql,i
		sql = "select count(*) as cnt from db_academy.dbo.tbl_diy_item_qna p" + vbcrlf
		sql = sql + " left join db_academy.dbo.tbl_diy_item i on p.itemid=i.itemid"
		sql = sql + " where p.itemid<>0"

		if FRectCDL<>"" then
			sql = sql + " and i.cate_large = '" + Cstr(FRectCDL) + "'" + vbcrlf
		end if

		if FRectCDm<>"" then
			sql = sql + " and i.cate_mid = '" + Cstr(FRectCDm) + "'" + vbcrlf
		end if

		if FRectCDs<>"" then
			sql = sql + " and i.cate_small = '" + Cstr(FRectCDs) + "'" + vbcrlf
		end if

		if FRectuserid<>"" then
			sql = sql + " and p.userid = '" + Cstr(FRectuserid) + "'" + vbcrlf
		end if

		if frectstartdate <> "" and frectenddate <> "" then
			sql = sql + " and p.regdate between '"+ Cstr(frectstartdate) +"' and '"+ Cstr(dateadd("d",1,frectenddate)) +"'" + vbcrlf
		end if

		if FRectItemID<>"" then
			sql = sql + " and p.itemid = " + Cstr(FRectItemID) + "" + vbcrlf
		end if

		if FRectMakerid<>"" then
			sql = sql + " and i.makerid = '" + Cstr(FRectMakerid) + "'" + vbcrlf
		end if

		if FRectOnlyTenBeasong<>"" then
			sql = sql + " and i.deliverytype in ('2','5')" + vbcrlf
		end if

		if FReckMiFinish<>"" then
			sql = sql + " and p.replydate is null" + vbcrlf
		end if

		if CItemDiv<>"" then
			sql= sql + " and i.itemdiv='" + CStr(CItemDiv) + "'" + vbcrlf
		end if

		sql = sql + " and p.isusing ='Y'" + vbcrlf

		rsACADEMYget.Open sql, dbACADEMYget, 1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sql = "select top " + CStr(FPageSize*FCurrpage) + " p.idx as id,p.userid,p.itemid,i.makerid,p.username,'' as cdl,p.contents,"
		sql = sql + " '99' as qadiv, p.replyuser,p.replycontents,p.regdate, '' as brandname, p.replydate, i.deliverytype, i.itemdiv "
		sql = sql + " from db_academy.dbo.tbl_diy_item_qna p" + vbcrlf
		sql = sql + " left join db_academy.dbo.tbl_diy_item i on p.itemid=i.itemid"

		sql = sql + " where p.itemid<>0"

		if FRectCDL<>"" then
			sql = sql + " and i.cate_large = '" + Cstr(FRectCDL) + "'" + vbcrlf
		end if

		if FRectCDm<>"" then
			sql = sql + " and i.cate_mid = '" + Cstr(FRectCDm) + "'" + vbcrlf
		end if

		if FRectCDs<>"" then
			sql = sql + " and i.cate_small = '" + Cstr(FRectCDs) + "'" + vbcrlf
		end if

		if FRectuserid<>"" then
			sql = sql + " and p.userid = '" + Cstr(FRectuserid) + "'" + vbcrlf
		end if

		if frectstartdate <> "" and frectenddate <> "" then
			sql = sql + " and p.regdate between '"+ Cstr(frectstartdate) +"' and '"+ Cstr(dateadd("d",1,frectenddate)) +"'" + vbcrlf
		end if

		if FRectItemID<>"" then
			sql = sql + " and p.itemid = " + Cstr(FRectItemID) + "" + vbcrlf
		end if

		if FRectMakerid<>"" then
			sql = sql + " and i.makerid = '" + Cstr(FRectMakerid) + "'" + vbcrlf
		end if

		if FRectOnlyTenBeasong<>"" then
			sql = sql + " and i.deliverytype in ('2','5')" + vbcrlf
		end if

		if FReckMiFinish<>"" then
			sql = sql + " and p.replydate is null" + vbcrlf
		end if

		if CItemDiv<>"" then
			sql= sql + " and i.itemdiv='" + CStr(CItemDiv) + "'" + vbcrlf
		end if

		sql = sql + " and p.isusing ='Y'" + vbcrlf
		sql = sql + " order by p.regdate desc"
		'response.write sql

		rsACADEMYget.pagesize = FPageSize

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CItemQnaSubItem

				FItemList(i).FID = rsACADEMYget("id")
				FItemList(i).FItemID = rsACADEMYget("itemid")
				FItemList(i).FUserid = rsACADEMYget("userid")
				FItemList(i).Fmakerid = rsACADEMYget("makerid")
				'FItemList(i).FCdl = rsACADEMYget("cdl")
				FItemList(i).Fusername = db2html(rsACADEMYget("username"))
				FItemList(i).FContents = db2html(rsACADEMYget("contents"))
				FItemList(i).FTitle = DdotFormat(FItemList(i).FContents,35)

				FitemList(i).FQadiv=rsACADEMYget("qadiv")
				FItemList(i).FReplyuser = rsACADEMYget("replyuser")

				FItemList(i).FReplycontents = db2html(rsACADEMYget("replycontents"))
				FItemList(i).FReplytitle = DdotFormat(FItemList(i).FReplycontents,40)

				FItemList(i).Fregdate = rsACADEMYget("regdate")
				FItemList(i).FBrandName = db2html(rsACADEMYget("brandname"))
				FItemList(i).FReplydate = rsACADEMYget("replydate")
				FItemList(i).Fdeliverytype = rsACADEMYget("deliverytype")
				FItemList(i).FItemDiv = rsACADEMYget("itemdiv")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end sub

	public sub getOneItemQna()
		dim sql,i
		sql = "select top 1 q.idx as id,q.userid,q.itemid,i.makerid,q.username,'' as cdl,q.contents,q.usermail,"
		sql = sql + " q.replyuser,q.replycontents,q.regdate, '' as brandname, q.replydate, i.listimage,i.itemname,IsNULL(i.sellcash,0) as sellcash, q.userlevel "
		sql = sql + " from db_academy.dbo.tbl_diy_item_qna q" + vbcrlf
		sql = sql + " left join db_academy.dbo.tbl_diy_item i on q.itemid=i.itemid"
		sql = sql + " where q.idx=" + CStr(FRectID)
		if FRectMakerid<>"" then
			sql = sql + " and i.makerid = '" + Cstr(FRectMakerid) + "'" + vbcrlf
		end if

		rsACADEMYget.Open sql, dbACADEMYget, 1

		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FOneItem = new CItemQnaSubItem

				FOneItem.FID = rsACADEMYget("id")
				FOneItem.FItemID = rsACADEMYget("itemid")
				FOneItem.FUserid = rsACADEMYget("userid")
				FOneItem.Fmakerid = rsACADEMYget("makerid")
				FOneItem.FCdl = rsACADEMYget("cdl")
				FOneItem.Fusername = db2html(rsACADEMYget("username"))
				FOneItem.FContents = db2html(rsACADEMYget("contents"))
				FOneItem.FTitle = DdotFormat(FOneItem.FContents,40)

				FOneItem.FReplyuser = rsACADEMYget("replyuser")

				FOneItem.FReplycontents = db2html(rsACADEMYget("replycontents"))
				FOneItem.FReplytitle = DdotFormat(FOneItem.FReplycontents,40)

				FOneItem.Fregdate = rsACADEMYget("regdate")
				FOneItem.FBrandName = db2html(rsACADEMYget("brandname"))
				FOneItem.FReplydate = rsACADEMYget("replydate")

				FOneItem.Flistimage = "http://image.thefingers.co.kr/diyitem/webimage/list/" + GetImageSubFolderByItemid(FOneItem.Fitemid) + "/" + rsACADEMYget("listimage")
				FOneItem.FItemName = db2html(rsACADEMYget("itemname"))
				FOneItem.FSellcash = rsACADEMYget("sellcash")
				FOneItem.FUserLevel = rsACADEMYget("userlevel")
				FOneItem.Fusermail = db2html(rsACADEMYget("usermail"))
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end sub

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