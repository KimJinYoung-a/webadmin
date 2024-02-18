<%
'###########################################################
' Description : 게시판
' Hieditor : 2009.04.17 이상구 생성
'			 2016.07.20 한용민 수정
'			 2017-04-25 이종화 추가
'###########################################################

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
    public FEmailok  ''2017/04/10 추가
    public Fisusing
	Public FSecretYN '//비밀글 추가

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
		end if
	end function

	'/사용금지		'/공용펑션에 공용함수 쓸것.		'/2016.07.20 한용민
	public function GetUserLevelStr()
    	if IsNULL(Fuserlevel) then
    		GetUserLevelStr = "&nbsp;"
    	elseif CStr(Fuserlevel)="0" then
    		GetUserLevelStr = "&nbsp;"
    	elseif CStr(Fuserlevel)="1" then
    		GetUserLevelStr = "<font color=#33ff66>Green</font>"
    	elseif CStr(Fuserlevel)="2" then
    		GetUserLevelStr = "<font color=#3366ff>Blue</font>"
    	elseif CStr(Fuserlevel)="3" then
    		GetUserLevelStr = "<font color=#ff3366>Vip</font>"
    	elseif CStr(Fuserlevel)="9" then
    		GetUserLevelStr = "<font color=#ff33ff>Mania</font>"
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

	public FRectsecretYN
	public FRectItemID
	public FRectCDL
	public FRectId
	public FOneItem
	public FReckMiFinish
	public FRectOnlyTenBeasong
	public FRectMakerid
	public frectcdm
	public frectcds
	public FRectCateCode
	public frectuserid
	public frectstartdate
	public frectenddate
	public FRectDPlusDay
	public CItemDiv
	Public FSecretYN '//비밀글
	public FRectissoldout
	public FRectcontents

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
		dim sqlStr,i

		sqlStr = "select top " + CStr(FPageSize) + " q.id,q.userid,q.itemid,q.makerid,q.cdl,q.contents,"
		sqlStr = sqlStr + " q.replyuser,q.replycontents,q.regdate, q.brandname, q.replydate, "
		sqlStr = sqlStr + " i.listimage"
		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_my_item_qna q" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on q.itemid=i.itemid"
		'sqlStr = sqlStr + " where q.cdl = '" + Cstr(FRectCDL) + "'" + vbcrlf
		sqlStr = sqlStr + " where q.cdl <>''" + vbcrlf
		sqlStr = sqlStr + " and q.isusing ='Y'" + vbcrlf
		sqlStr = sqlStr + " and q.replydate is not null" + vbcrlf
		sqlStr = sqlStr + " order by replydate desc"

		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CItemQnaSubItem

				FItemList(i).FID = rsget("id")
				FItemList(i).FUserid = rsget("userid")
				FItemList(i).Fmakerid = rsget("makerid")
				FItemList(i).FCdl = rsget("cdl")
				FItemList(i).FContents = db2html(rsget("contents"))
				FItemList(i).FTitle = DdotFormat(FItemList(i).FContents,40)
				FItemList(i).FReplyuser = rsget("replyuser")
				FItemList(i).FReplycontents = db2html(rsget("replycontents"))
				FItemList(i).FReplytitle = DdotFormat(FItemList(i).FReplycontents,40)
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				FItemList(i).Freplydate = rsget("replydate")
				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Flistimage = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	public sub ItemQnaList()
		dim sqlStr, i

		sqlStr = "select count(*) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg  from [db_cs].[dbo].tbl_my_item_qna p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock) on p.itemid=i.itemid"
		sqlStr = sqlStr + " where p.itemid<>0"

		if application("Svr_Info") <> "Dev" and FRectItemID = "" then
			sqlStr = sqlStr + " and p.id >= 400000 "
		end if

		if FRectCDL<>"" then
			sqlStr = sqlStr + " and i.cate_large = '" + Cstr(FRectCDL) + "'" + vbcrlf
		end if

		if FRectCDm<>"" then
			sqlStr = sqlStr + " and i.cate_mid = '" + Cstr(FRectCDm) + "'" + vbcrlf
		end if

		if FRectCDs<>"" then
			sqlStr = sqlStr + " and i.cate_small = '" + Cstr(FRectCDs) + "'" + vbcrlf
		end if

		if FRectCateCode<>"" then
			sqlStr = sqlStr + " and exists(" + vbcrlf
			sqlStr = sqlStr + "     select 1" + vbcrlf
			sqlStr = sqlStr + "     from db_item.dbo.tbl_display_cate_item as c with (nolock)" + vbcrlf
			sqlStr = sqlStr + "     where c.isDefault='y'" + vbcrlf
			sqlStr = sqlStr + " 		and c.catecode like '" + Cstr(FRectCateCode) + "%'" + vbcrlf
			sqlStr = sqlStr + " 		and c.itemid=i.itemid" + vbcrlf
			sqlStr = sqlStr + " )"
		end if

		if FRectuserid<>"" then
			sqlStr = sqlStr + " and p.userid = '" + Cstr(FRectuserid) + "'" + vbcrlf
		end if

		if frectstartdate <> "" and frectenddate <> "" then
			sqlStr = sqlStr + " and p.regdate between '"+ Cstr(frectstartdate) +"' and '"+ Cstr(dateadd("d",1,frectenddate)) +"'" + vbcrlf
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr + " and p.itemid = " + Cstr(FRectItemID) + "" + vbcrlf
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid = '" + Cstr(FRectMakerid) + "'" + vbcrlf
		end if

		if FRectOnlyTenBeasong<>"" then
			if (FRectOnlyTenBeasong = "Y") then
				sqlStr = sqlStr + " and i.deliverytype in ('1', '4')" + vbcrlf
			elseif (FRectOnlyTenBeasong = "N") then
				sqlStr = sqlStr + " and i.deliverytype not in ('1', '4')" + vbcrlf
			end if
		end if

		if (FRectDPlusDay <> "") then
			'// D+3 일 초과만
			sqlStr = sqlStr + " and DateDiff(d, p.regdate, getdate()) >= 3 " + vbcrlf
		end if

		if FReckMiFinish<>"" then
			sqlStr = sqlStr + " and p.replydate is null" + vbcrlf
		end if

		if CItemDiv<>"" then
			sqlStr = sqlStr + " and i.itemdiv='" + CStr(CItemDiv) + "'" + vbcrlf
		end If

		If FRectsecretYN <> "" Then
			sqlStr = sqlStr + " and p.secretYN='" + CStr(FRectsecretYN) + "'" + vbcrlf
		End If

		If FRectissoldout = "on" Then
			sqlStr = sqlStr & " and i.sellyn not in ('S','N')" + vbcrlf
		End If

		If FRectcontents<>"" Then
			sqlStr = sqlStr & " and p.contents like '%"& FRectcontents &"%'" + vbcrlf
		End If

		sqlStr = sqlStr + " and p.isusing ='Y'" + vbcrlf

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		FTotalPage = rsget("totPg")
		rsget.Close

		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrpage) + " p.id,p.userid,p.itemid,i.makerid,p.username,p.cdl,p.contents,"
		sqlStr = sqlStr + " p.qadiv, p.replyuser,p.replycontents,p.regdate, p.brandname, p.replydate, i.deliverytype, i.itemdiv , p.secretYN "
		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_my_item_qna p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock) on p.itemid=i.itemid"

		sqlStr = sqlStr + " where p.itemid<>0"

		if application("Svr_Info") <> "Dev" and FRectItemID = "" then
			sqlStr = sqlStr + " and p.id >= 400000 "
		end if

		if FRectCDL<>"" then
			sqlStr = sqlStr + " and i.cate_large = '" + Cstr(FRectCDL) + "'" + vbcrlf
		end if

		if FRectCDm<>"" then
			sqlStr = sqlStr + " and i.cate_mid = '" + Cstr(FRectCDm) + "'" + vbcrlf
		end if

		if FRectCDs<>"" then
			sqlStr = sqlStr + " and i.cate_small = '" + Cstr(FRectCDs) + "'" + vbcrlf
		end if

		if FRectCateCode<>"" then
			sqlStr = sqlStr + " and exists(" + vbcrlf
			sqlStr = sqlStr + "     select 1" + vbcrlf
			sqlStr = sqlStr + "     from db_item.dbo.tbl_display_cate_item as c with (nolock)" + vbcrlf
			sqlStr = sqlStr + "     where c.isDefault='y'" + vbcrlf
			sqlStr = sqlStr + " 		and c.catecode like '" + Cstr(FRectCateCode) + "%'" + vbcrlf
			sqlStr = sqlStr + " 		and c.itemid=i.itemid" + vbcrlf
			sqlStr = sqlStr + " )"
		end if

		if FRectuserid<>"" then
			sqlStr = sqlStr + " and p.userid = '" + Cstr(FRectuserid) + "'" + vbcrlf
		end if

		if frectstartdate <> "" and frectenddate <> "" then
			sqlStr = sqlStr + " and p.regdate between '"+ Cstr(frectstartdate) +"' and '"+ Cstr(dateadd("d",1,frectenddate)) +"'" + vbcrlf
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr + " and p.itemid = " + Cstr(FRectItemID) + "" + vbcrlf
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid = '" + Cstr(FRectMakerid) + "'" + vbcrlf
		end if

		if FRectOnlyTenBeasong<>"" then
			if (FRectOnlyTenBeasong = "Y") then
				sqlStr = sqlStr + " and i.deliverytype in ('1', '4')" + vbcrlf
			elseif (FRectOnlyTenBeasong = "N") then
				sqlStr = sqlStr + " and i.deliverytype not in ('1', '4')" + vbcrlf
			end if
		end if

		if (FRectDPlusDay <> "") then
			'// D+3 일 초과만
			sqlStr = sqlStr + " and DateDiff(d, p.regdate, getdate()) >= 3 " + vbcrlf
		end if

		if FReckMiFinish<>"" then
			sqlStr = sqlStr + " and p.replydate is null" + vbcrlf
		end if

		if CItemDiv<>"" then
			sqlStr = sqlStr + " and i.itemdiv='" + CStr(CItemDiv) + "'" + vbcrlf
		end If

		If FRectsecretYN <> "" Then
			sqlStr = sqlStr + " and p.secretYN='" + CStr(FRectsecretYN) + "'" + vbcrlf
		End If

		If FRectissoldout = "on" Then
			sqlStr = sqlStr & " and i.sellyn not in ('S','N')" + vbcrlf
		End If

		If FRectcontents<>"" Then
			sqlStr = sqlStr & " and p.contents like '%"& FRectcontents &"%'" + vbcrlf
		End If

		sqlStr = sqlStr + " and p.isusing ='Y'" + vbcrlf
		sqlStr = sqlStr + " order by p.regdate desc"

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize

		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemQnaSubItem
				FItemList(i).FID = rsget("id")
				FItemList(i).FItemID = rsget("itemid")
				FItemList(i).FUserid = rsget("userid")
				FItemList(i).Fmakerid = rsget("makerid")
				FItemList(i).FCdl = rsget("cdl")
				FItemList(i).Fusername = db2html(rsget("username"))
				FItemList(i).FContents = db2html(rsget("contents"))
				FItemList(i).FTitle = DdotFormat(FItemList(i).FContents,35)
				FitemList(i).FQadiv=rsget("qadiv")
				FItemList(i).FReplyuser = rsget("replyuser")
				FItemList(i).FReplycontents = db2html(rsget("replycontents"))
				FItemList(i).FReplytitle = DdotFormat(FItemList(i).FReplycontents,40)
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				FItemList(i).FReplydate = rsget("replydate")
				FItemList(i).Fdeliverytype = rsget("deliverytype")
				FItemList(i).FItemDiv = rsget("itemdiv")
				FItemList(i).FSecretYN = rsget("secretYN")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	public sub getOneItemQna()
		dim sqlStr,i

		sqlStr = "select top 1 q.id,q.userid,q.itemid,i.makerid,q.username,q.cdl,q.contents,q.usermail,"&vbCRLF
		sqlStr = sqlStr + " q.replyuser,q.replycontents,q.regdate, i.brandname, q.replydate, i.listimage,i.itemname,IsNULL(i.sellcash,0) as sellcash, q.userlevel "&vbCRLF
		sqlStr = sqlStr + " ,isnull(replace(emailok,' ','N'),'N') as emailok, q.isusing , q.secretYN"&vbCRLF ''2017/04/10 추가
		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_my_item_qna q" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on q.itemid=i.itemid"
		sqlStr = sqlStr + " where q.id=" + CStr(FRectID)

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid = '" + Cstr(FRectMakerid) + "'" + vbcrlf
		end if

		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FOneItem = new CItemQnaSubItem
				FOneItem.FID = rsget("id")
				FOneItem.FItemID = rsget("itemid")
				FOneItem.FUserid = rsget("userid")
				FOneItem.Fmakerid = rsget("makerid")
				FOneItem.FCdl = rsget("cdl")
				FOneItem.Fusername = db2html(rsget("username"))
				FOneItem.FContents = db2html(rsget("contents"))
				FOneItem.FTitle = DdotFormat(FOneItem.FContents,40)
				FOneItem.FReplyuser = rsget("replyuser")
				FOneItem.FReplycontents = db2html(rsget("replycontents"))
				FOneItem.FReplytitle = DdotFormat(FOneItem.FReplycontents,40)
				FOneItem.Fregdate = rsget("regdate")
				FOneItem.FBrandName = db2html(rsget("brandname"))
				FOneItem.FReplydate = rsget("replydate")
				FOneItem.Flistimage = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FOneItem.Fitemid) + "/" + rsget("listimage")
				FOneItem.FItemName = db2html(rsget("itemname"))
				FOneItem.FSellcash = rsget("sellcash")
				FOneItem.FUserLevel = rsget("userlevel")
				FOneItem.Fusermail = db2html(rsget("usermail"))
				FOneItem.Femailok   = rsget("emailok")
				FOneItem.Fisusing = rsget("isusing")
				FOneItem.FSecretYN = rsget("secretYN")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
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
