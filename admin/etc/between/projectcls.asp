<%
Class cProjectOneItem
	Public FPjt_code
	Public FPjt_kind
	Public FPjt_name
	Public FPjt_topImgurl
	Public FPjt_gender
	Public FPjt_state
	Public FPjt_sortType
	Public FPjt_using

	Public FPjtgroup_code
	Public FPjtgroup_desc
	Public FPjtgroup_sort
	Public FPjtgroup_pcode
	Public FPjtgroup_depth
	Public FPjtgroup_pdesc
	Public FPjtgroup_BGColor
	Public FPjtgroup_FontColor

	Public FIsTop

	Public FItemid
	Public FPjtitem_sort
	Public FMakerid
	Public FItemname
	Public FSellcash
	Public FBuycash
	Public FOrgprice
	Public FOrgsuplycash
	Public FSailprice
	Public FSailsuplycash
	Public FMileage
	Public FSmallimage
	Public FListimage
	Public FSellyn
	Public FDeliverytype
	Public FLimityn
	Public FDanjongyn
	Public FSailyn
	Public FIsusing
	Public FLimitno
	Public FLimitsold
	Public FItemcouponyn
	Public FItemcoupontype
	Public FItemcouponvalue
	Public FMwdiv
	Public FChgItemname

	Public Fcate_large
	Public Fcate_mid
	Public Fcate_small
	Public Fitemdiv
	Public Fitemgubun
	Public Fregdate
	Public Flastupdate
	Public FsellEndDate
	Public Fisextusing
	Public Fspecialuseritem
	Public Fvatinclude
	Public Fdeliverarea
	Public Fdeliverfixday
	Public Fismobileitem
	Public Fpojangok
	Public Fevalcnt
	Public Foptioncnt
	Public Fitemrackcode
	Public Fupchemanagecode
	Public Fbrandname
	Public Flistimage120
	Public Fcurritemcouponidx
	Public FinfoimageExists
	Public FdefaultFreeBeasongLimit
	Public FdefaultDeliverPay
	Public FdefaultDeliveryType
	Public Fitemscore

	Public Function IsUpcheBeasong()
		If Fdeliverytype = "2" or Fdeliverytype = "5" or Fdeliverytype = "9" or Fdeliverytype = "7" Then
			IsUpcheBeasong = true
		Else
			IsUpcheBeasong = false
		End If
	End Function

    Public Function IsSoldOut()
		IsSoldOut = (FSellYn <> "Y") or ((FLimitYn = "Y") and (GetLimitEa() < 1))
	End Function

    Public Function GetLimitEa()
		If FLimitNo - FLimitSold < 0 Then
			GetLimitEa = 0
		Else
			GetLimitEa = FLimitNo - FLimitSold
		End if
	End Function
End Class

Class cProject
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	Public FRegdate

	Public FRectPjt_code
	Public FRectPjtgroup_code

	Public FRectMakerid
	Public FRectItemid
	Public FRectItemName
	Public FRectKeyword

	Public FRectSellYN
	Public FRectIsUsing
	Public FRectDanjongyn
	Public FRectLimityn
	Public FRectMWDiv
	Public FRectSailYn
	Public FRectDeliveryType

	Public FRectCate_Large
	Public FRectCate_Mid
	Public FRectCate_Small
	Public FRectDispCate
	Public FRectSortDiv
	Public FRectItemDiv
	Public FRectCouponYn
	Public FRectVatYn

	Public FRectSGroup
	Public FRectSort

	Public FRectPjt_kind
	Public FRectSelPjt
	Public FRectSPtxt
	Public FRectPjt_state
	Public FRectPjt_gender

	Public Sub getProjectList()
		Dim sqlStr, addSql, i

		If FRectPjt_kind <> "" Then
			addSql = addSql & " and pjt_kind='" & FRectPjt_kind & "'"
		End If

		If FRectSelPjt <> "" AND FRectSPtxt <> "" Then
			If FRectSelPjt = "pjt_code" Then
				addSql = addSql & " and pjt_code='" & FRectSPtxt & "'"
			ElseIf FRectSelPjt = "pjt_name" Then
				addSql = addSql & " and pjt_name='" & FRectSPtxt & "'"
			End If
		End If

		If FRectPjt_state <> "" Then
			addSql = addSql & " and pjt_state='" & FRectPjt_state & "'"
		End If

		If FRectPjt_gender <> "" Then
			addSql = addSql & " and pjt_gender='" & FRectPjt_gender & "'"
		End If

		If FRectIsusing <> "" Then
			addSql = addSql & " and pjt_using='" & FRectIsusing & "'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_project "
		sqlStr = sqlStr & " WHERE 1 = 1  " & addSql
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " pjt_code, pjt_name, pjt_kind, pjt_topImgurl, pjt_gender, pjt_state, pjt_using "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_project "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY pjt_code DESC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cProjectOneItem
					FItemList(i).FPjt_code			= rsCTget("pjt_code")
					FItemList(i).FPjt_name			= rsCTget("pjt_name")
					FItemList(i).FPjt_kind			= rsCTget("pjt_kind")
					FItemList(i).FPjt_topImgurl		= rsCTget("pjt_topImgurl")
					FItemList(i).FPjt_gender		= rsCTget("pjt_gender")
					FItemList(i).FPjt_state			= rsCTget("pjt_state")
					FItemList(i).FPjt_using			= rsCTget("pjt_using")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	'## getProjectItemGroup : 그룹내용가져오기 ##
	Public Sub getProjectItemGroup()
		IF FRectPjt_code = "" THEN Exit Sub
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT pjtgroup_code, pjtgroup_desc, pjtgroup_sort, pjtgroup_pcode, pjtgroup_depth, pjtgroup_BGColor, pjtgroup_FontColor, "
		sqlStr = sqlStr & " (SELECT pjtgroup_desc FROM [db_outmall].[dbo].[tbl_between_project_group] WHERE pjtgroup_code = a.pjtgroup_pcode) as isTop "
		sqlStr = sqlStr & " FROM [db_outmall].[dbo].[tbl_between_project_group] as a "
		sqlStr = sqlStr & " WHERE pjt_code = "&FRectPjt_code&" and pjtgroup_using ='Y' ORDER BY pjtgroup_depth, pjtgroup_sort ASC "
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			Do until rsCTget.EOF
				Set FItemList(i) = new cProjectOneItem
					FItemList(i).FPjtgroup_code			= rsCTget("pjtgroup_code")
					FItemList(i).FPjtgroup_desc			= rsCTget("pjtgroup_desc")
					FItemList(i).FPjtgroup_sort			= rsCTget("pjtgroup_sort")
					FItemList(i).FPjtgroup_pcode		= rsCTget("pjtgroup_pcode")
					FItemList(i).FPjtgroup_depth		= rsCTget("pjtgroup_depth")
					FItemList(i).FPjtgroup_BGColor		= rsCTget("pjtgroup_BGColor")
					FItemList(i).FPjtgroup_FontColor	= rsCTget("pjtgroup_FontColor")
					FItemList(i).FIsTop					= rsCTget("isTop")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Function fnGetRootGroup()
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT pjtgroup_code, pjtgroup_desc FROM [db_outmall].[dbo].[tbl_between_project_group] "
		sqlStr = sqlStr & " WHERE pjt_code = "&FRectPjt_code&" and pjtgroup_pcode = 0 and pjtgroup_using ='Y' "
		rsCTget.Open sqlStr,dbCTget
		IF not rsCTget.EOF THEN
			fnGetRootGroup = rsCTget.getRows()
		End IF
		rsCTget.Close
	End Function

	Public Sub getProjectItemGroupCont()
		Dim sqlStr
		IF FRectPjtgroup_code = "" THEN Exit Sub
		sqlStr = ""
		sqlStr = sqlStr & " SELECT pjtgroup_code, pjtgroup_desc, pjtgroup_sort, pjtgroup_pcode, pjtgroup_depth, pjtgroup_BGColor, pjtgroup_FontColor, "
		sqlStr = sqlStr & "		isnull((select pjtgroup_desc from [db_outmall].[dbo].[tbl_between_project_group] where pjtgroup_code = a.pjtgroup_pcode),'최상위') as pjtgroup_pdesc"
		sqlStr = sqlStr & "	FROM  [db_outmall].[dbo].[tbl_between_project_group] as a "
		sqlStr = sqlStr & "	WHERE pjt_code = "&FRectpjt_code&" and pjtgroup_code="&FRectPjtgroup_code&" and pjtgroup_using ='Y' "
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		If not rsCTget.EOF Then
			Set FItemList(0) = new cProjectOneItem
				FItemList(0).FPjtgroup_code		= rsCTget("pjtgroup_code")
				FItemList(0).FPjtgroup_desc		= rsCTget("pjtgroup_desc")
				FItemList(0).FPjtgroup_sort		= rsCTget("pjtgroup_sort")
				FItemList(0).FPjtgroup_pcode	= rsCTget("pjtgroup_pcode")
				FItemList(0).FPjtgroup_depth	= rsCTget("pjtgroup_depth")
				FItemList(0).FPjtgroup_pdesc	= rsCTget("pjtgroup_pdesc")
				FItemList(0).FPjtgroup_BGColor	= rsCTget("pjtgroup_BGColor")
				FItemList(0).FPjtgroup_FontColor= rsCTget("pjtgroup_FontColor")
		End If
		rsCTget.Close
	End Sub

	Public Sub getProjectCont()
		Dim sqlStr
		IF FRectPjt_code = "" THEN Exit Sub
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 pjt_code, pjt_name, pjt_kind, pjt_gender, pjt_state, pjt_sortType, pjt_using, isnull(pjt_topImgUrl, '') as pjt_topImgUrl "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_project "
		sqlStr = sqlStr & " WHERE pjt_code = "&FRectPjt_code
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		If not rsCTget.EOF Then
			Set FItemList(0) = new cProjectOneItem
				FItemList(0).FPjt_code 		= rsCTget("pjt_code")
			 	FItemList(0).FPjt_name		= rsCTget("pjt_name")
				FItemList(0).FPjt_kind 		= rsCTget("pjt_kind")
				FItemList(0).FPjt_gender 	= rsCTget("pjt_gender")
				FItemList(0).FPjt_state 	= rsCTget("pjt_state")
				FItemList(0).FPjt_sortType 	= rsCTget("pjt_sortType")
				FItemList(0).FPjt_using 	= rsCTget("pjt_using")
				FItemList(0).FPjt_topImgUrl	= rsCTget("pjt_topImgUrl")
		End If
		rsCTget.Close
	End Sub


	'## getProjectItem :기획전상품 가져오기 ##
	Public Sub getProjectItem()
		Dim sqlStr, addSql, i, strSort

		'그룹 검색
		If FRectSGroup <> "" Then
			IF FRectSGroup = 0 Then
				addSql = " AND (G.pjtgroup_code  is null OR G.pjtgroup_code = 0 )"
			Else
				addSql = " AND G.pjtgroup_code = '"&FRectSGroup&"'"
			End If
		End If

		'정렬 검색
		If FRectSort = "slsell" Then
			strSort = " ORDER BY i.sellcash ASC"
		ElseIf FRectSort = "shsell" Then
			strSort = " ORDER BY i.sellcash DESC"
		ElseIf FRectSort = "sbest" Then
			strSort = " ORDER BY c.recentsellcount DESC, i.sellcash DESC"
		ElseIf FRectSort = "sevtitem" Then
			strSort = " ORDER BY G.pjtitem_sort ASC, i.itemid DESC"
		ElseIf FRectSort = "sevtgroup" Then
			strSort = " ORDER BY G.pjtgroup_code ASC"
		ElseIf FRectSort = "sbrand" Then
			strSort = " ORDER BY i.makerid ASC"
		Else
			strSort = " ORDER BY i.itemid DESC "
		END IF

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_project_groupItem AS G "
		sqlStr = sqlStr & " JOIN [db_AppWish].[dbo].tbl_item AS i ON G.itemid = i.itemid  "
		sqlStr = sqlStr & " LEFT JOIN [db_AppWish].[dbo].[tbl_item_contents] AS c ON G.itemid = c.itemid "
		sqlStr = sqlStr & "	WHERE G.pjt_code = '"&FRectpjt_code&"'" & addSql
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " G.itemid, G.pjtgroup_code, G.pjtitem_sort,  i.makerid, i.itemname, i.sellcash "
		sqlStr = sqlStr & " ,i.buycash,i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash, i.mileage, i.smallimage, i.listimage,   i.sellyn, i.deliverytype "
		sqlStr = sqlStr & " ,i.limityn, i.danjongyn, i.sailyn, i.isusing, i.limitno , i.limitsold, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue "
		sqlStr = sqlStr & " ,i.mwdiv "
		sqlStr = sqlStr & " ,(SELECT TOP 1 isnull(chgItemname, '') FROM db_outmall.dbo.tbl_between_cate_item AS CI WHERE G.itemid = CI.itemid) as chgItemname "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_project_groupItem AS G "
		sqlStr = sqlStr & " JOIN [db_AppWish].[dbo].tbl_item AS i ON G.itemid = i.itemid  "
		sqlStr = sqlStr & " LEFT JOIN [db_AppWish].[dbo].[tbl_item_contents] AS c ON G.itemid = c.itemid "
		sqlStr = sqlStr & "	WHERE G.pjt_code = '"&FRectpjt_code&"'" & addSql & strSort
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cProjectOneItem
					FItemList(i).FItemid			= rsCTget("itemid")
					FItemList(i).FPjtgroup_code		= rsCTget("pjtgroup_code")
					FItemList(i).FPjtitem_sort		= rsCTget("pjtitem_sort")
					FItemList(i).FMakerid			= rsCTget("makerid")
					FItemList(i).FItemname			= rsCTget("itemname")
					FItemList(i).FSellcash			= rsCTget("sellcash")
					FItemList(i).FBuycash			= rsCTget("buycash")
					FItemList(i).FOrgprice			= rsCTget("orgprice")
					FItemList(i).FOrgsuplycash		= rsCTget("orgsuplycash")
					FItemList(i).FSailprice			= rsCTget("sailprice")
					FItemList(i).FSailsuplycash		= rsCTget("sailsuplycash")
					FItemList(i).FMileage			= rsCTget("mileage")
					FItemList(i).FSmallimage		= rsCTget("smallimage")
					FItemList(i).FListimage			= rsCTget("listimage")
					FItemList(i).FSellyn			= rsCTget("sellyn")
					FItemList(i).FDeliverytype		= rsCTget("deliverytype")
					FItemList(i).FLimityn			= rsCTget("limityn")
					FItemList(i).FDanjongyn			= rsCTget("danjongyn")
					FItemList(i).FSailyn			= rsCTget("sailyn")
					FItemList(i).FIsusing			= rsCTget("isusing")
					FItemList(i).FLimitno			= rsCTget("limitno")
					FItemList(i).FLimitsold			= rsCTget("limitsold")
					FItemList(i).FItemcouponyn		= rsCTget("itemcouponyn")
					FItemList(i).FItemcoupontype	= rsCTget("itemcoupontype")
					FItemList(i).FItemcouponvalue	= rsCTget("itemcouponvalue")
					FItemList(i).FMwdiv				= rsCTget("mwdiv")
					FItemList(i).FChgItemname		= rsCTget("chgItemname")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

    public Function IsSoldOut(FSellYn,FLimitYn,FLimitNo,FLimitSold)
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa(FLimitNo,FLimitSold)<1))
	end function

    public function GetLimitEa(FLimitNo,FLimitSold)
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public Function IsUpcheBeasong(Fdeliverytype)
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" or Fdeliverytype="7" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FScrollCount = 10
	End Sub

	Private Sub Class_Terminate()
    End Sub
End Class

Function getDBcodeByName(sval)
	Select Case sval
		Case "A"	response.write "생일"
		Case "B"	response.write "100일"
		Case "C"	response.write "1주년"
		Case "D"	response.write "결혼기념일"
		Case "E"	response.write "발렌타인데이"
		Case "F"	response.write "화이트데이"
		Case "G"	response.write "빼빼로데이"
		Case "H"	response.write "크리스마스"
		Case "I"	response.write "MD PICK"
		Case "J"	response.write "ETC 기획전"
		Case "K"	response.write "200일"
		Case "L"	response.write "500일"
		Case "M"	response.write "1000일"

		Case "0"	response.write "등록대기"
		Case "7"	response.write "오픈"
		Case "9"	response.write "종료"

	End Select
End Function

Sub sbGetOptProjectCodeValue(ByVal sType, ByVal selValue, ByVal sScript)
	Dim arrList, intLoop, iValue
%>
	<select class="select" name="<%= sType %>" <%= sScript %> >
		<option value="" >- Choice -</option>
		<option value="A" <%= CHKIIF(CStr(selValue)="A", "selected","") %> >생일</option>
		<option value="B" <%= CHKIIF(CStr(selValue)="B", "selected","") %> >100일</option>
		<option value="K" <%= CHKIIF(CStr(selValue)="K", "selected","") %> >200일</option>
		<option value="L" <%= CHKIIF(CStr(selValue)="L", "selected","") %> >500일</option>
		<option value="M" <%= CHKIIF(CStr(selValue)="M", "selected","") %> >1000일</option>
		<option value="C" <%= CHKIIF(CStr(selValue)="C", "selected","") %> >1주년</option>
		<option value="D" <%= CHKIIF(CStr(selValue)="D", "selected","") %> >결혼기념일</option>
		<option value="E" <%= CHKIIF(CStr(selValue)="E", "selected","") %> >발렌타인데이</option>
		<option value="F" <%= CHKIIF(CStr(selValue)="F", "selected","") %> >화이트데이</option>
		<option value="G" <%= CHKIIF(CStr(selValue)="G", "selected","") %> >빼빼로데이</option>
		<option value="H" <%= CHKIIF(CStr(selValue)="H", "selected","") %> >크리스마스</option>
		<option value="I" <%= CHKIIF(CStr(selValue)="I", "selected","") %> >MD PICK</option>
		<option value="J" <%= CHKIIF(CStr(selValue)="J", "selected","") %> >ETC 기획전</option>
	</select>
<%
End Sub

Sub sbAlertMsg(byVal strMsg, ByVal strUrl, ByVal strTarget)
	Dim strLink
	IF strUrl = "close" THEN	'팝업 창 닫을경우
		strLink = strTarget & ".close();"
	ELSEIF strUrl ="back" THEN	'이전 페이지로 이동
		strLink = "history.back(-1);"
	ELSE
		strLink = strTarget & ".location.href='" & strUrl & "';"
	END IF
%>
<script language="javascript">
	alert("<%=strMsg%>");
	<%=strLink%>;
</script>
<%		
	response.End
End Sub
%>