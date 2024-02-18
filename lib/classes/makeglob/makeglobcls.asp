<%
'####################################################
' Description :  메이크 글로비 관리 클래스
' History : 2015.10.27 원승현 생성
'####################################################

'// 상품detail 클래스
Class CMakeGlobItemDetail

	Public Fitemid '// 상품코드
	Public Fmakerid '// 판매자아이디
	Public FBrandName '// 브랜드명(영문)
	Public FBrandNameKr '// 브랜드명(한글)
	Public FitemName '// 상품명

	'------ 각 번호별로 한쌍으로 움직인다고 보면 됨--------
	Public FsellCash '// 현재 판매가(1)
	Public FbuyCash '// 현재 매입가(1)
	Public Forgprice '// 원래 판매가(2)
	Public Forgsuplycash '// 원래매입가(2)
	Public Fsailprice '// 할인시 판매가(3)
	Public Fsailsuplycash '// 할인시 매입가(3)
	'------ 각 번호별로 한쌍으로 움직인다고 보면 됨--------

	Public Fmileage '// 마일리지
	Public Fregdate '// 등록일
	Public Flastupdate '// 최종수정일
	Public FsellStdate '// 판매개시일
	Public Fsellyn '// 판매여부
	Public Flimityn '// 한정상품여부
	Public Fsailyn '// 세일여부
	Public Fisusing '// 상품사용여부
	
	'// 최종 한정갯수=Flimitno - Flimitsold
	Public Flimitno '// 한정갯수
	Public Flimitsold '// 한정판매갯수


	Public Fmainimage '// 상품 이미지 중에 하나인데 거의 안씀
	Public Fsmallimage '// 50x50 이미지
	Public Flistimage '// 100x100 이미지
	Public Flistimage120 '// 120x120 이미지
	Public Fbasicimage '// 400x400 이미지
	Public Ficon1image '// 200x200 이미지
	Public Ficon2image '// 150x150 이미지
	Public Fitemcouponyn '// 쿠폰사용여부
	Public Fbasic600image '// 600x600 이미지
	Public Fbasic1000image '// 1000x1000 이미지
	Public Fitemscore '// 상품점수(best 등 판단할때)
	Public Fitemweight '// 상품무게(글로비 표시할땐 /1000 해서 그람수로 줘야됨)
	Public FdeliverOverseas '// 해외배송여부(글로비는 Y만 불러옴)
	Public Ftenonlyyn '// 텐바이텐 독점상품여부
	Public Fcatecode '// 카테코리 코드
	Public Fdepth '// 카테고리 뎁스
	Public FisDefault '// 카테고리 표시 중 현재 기본값으로 사용되는 값(y면 해당 카테고리가 디폴트로 사용되어지는 값임)
	Public Fkeywords '// 해당 상품 키워드
	Public Fsourcearea '// 원산지
	Public Fmakername '// 제조사
	Public Fitemsource '// 상품 재료
	Public Fitemsize '// 상품크기값
	Public Fitemcontent '// 상품상세설명

	Public FMakeGlobChkEN '// 메이크 글로비쪽으로 상품이 넘겨 졌는지 여부(영문)
	Public FMakeGlobChkZH '// 메이크 글로비쪽으로 상품이 넘겨 졌는지 여부(중문)

	Public FMakeGlobHidden '// 메이크 글로비 숨김여부
	Public FMakeGlobSoldout '// 메이크 글로비 품절여부
	Public FMakeGlobProductKey '// 메이크 글로비 상품코드
	Public FMakeGlobupdate '// 메이크 글로비 업데이트 여부
	Public FMakeGlobupdateTime '// 메이크 글로비 업데이트 일자

	Public FBaesongGubun '// 배송구분(M,W - 텐베, U-업배)

	'// 솔드아웃 여부
    public Function IsSoldOut()
		IsSoldOut = (Fsellyn<>"Y") or ((Flimityn="Y") and (GetLimitEa()<1))
	end function

	'// 한정 상품일 경우 남은 한정갯수
    public function GetLimitEa()
		if Flimitno-Flimitsold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = Flimitno-Flimitsold
		end if
	end function

	'// 세일포함 실제가격
	public Function getRealPrice() '!
		getRealPrice = FsellCash
	End Function
End Class


Class CMakeGlobItem
	public FItemList()
	public FTotalCount '// 총갯수
	public FCurrPage '// 현재페이지 번호
	public FTotalPage '// 총 페이지 갯수
	public FPageSize '// 페이지 사이즈
	public FResultCount '// 결과값 갯수
	Public FScrollCount '// 스크롤카운트
	Public FRectBrandName '// 브랜드명
	Public FRectCateCode '// 카테고리 코드
	Public FRectItemName '// 상품명
	Public FRectItemId '// 상품코드
	Public FRectSellyn '// 텐바이텐 판매여부(N-품절, S-일시품절, Y-판매중)
	Public FRectLimityn '// 텐바이텐 한정판매여부
	Public FRectIsUsing '// 텐바이텐 사용여부
	Public FRectGIsHidden '// 글로비 숨김여부
	Public FRectGIssoldout '// 글로비 품절여부
	Public FRectGProductKey '// 글로비 상품코드
	Public FRectGIscheck '// 글로비 상품등록여부
	Public FRectMakeGlobChkEN '// 영문 입력여부
	Public FRectMakeGlobChkZH '// 중문 입력여부
	Public FRectMarginSt	'// 마진율검색시작값
	Public FRectMarginEd	'// 마진율검색종료값
	Public FRectSorgpriceSt	'// 판매가검색시작값
	Public FRectSorgpriceEd	'// 판매가검색종료값
	Public FRectBaesongGubun '// 배송구분(MW-텐배, U-업배)
	Public FRectMakerID '// 메이커id

	public function GetMakeGlobItemWaitingList()
		Dim strsql, addSql, i


'        if (FRectBrandName <> "") Then
'			addSql = addSql & " And c.userid = '"&FRectMakerID&"' "
'		End If

        if (FRectMakerID <> "") Then
			addSql = addSql & " And c.userid = '"&FRectMakerID&"' "
		End If

		If (FRectCateCode <> "") Then
			addSql = addSql & " And ci.catecode like '"&FRectCateCode&"%' "
		End If

'		If (FRectItemName <> "") Then
'			addSql = addSql & " And i.itemname like '%"&FRectItemName&"%' "
'		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		If (FRectSellyn <> "") Then
			addSql = addSql & " And i.sellyn = '"&FRectSellyn&"' "
		End If

		if FRectLimityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end if

		If (FRectIsUsing <> "") Then
			addSql = addSql & " And i.isusing = '"&FRectIsUsing&"' "
		End If

		If (FRectGIsHidden <> "") Then
			addSql = addSql & " And g.hidden = '"&FRectGIsHidden&"' "
		End If

		If (FRectGIssoldout <> "") Then
			addSql = addSql & " And g.soldout = '"&FRectGIssoldout&"' "
		End If

'		If (FRectGProductKey <> "") Then
'			addSql = addSql & " And g.product_key in ("&FRectGProductKey&") "
'		End If

		'글로비 상품번호 검색
        If (FRectGProductKey <> "") then
            If Right(Trim(FRectGProductKey) ,1) = "," Then
            	FRectItemid = Replace(FRectGProductKey,",,",",")
            	addSql = addSql & " and g.product_key in (" + Left(FRectGProductKey,Len(FRectGProductKey)-1) + ")"
            Else
				FRectGProductKey = Replace(FRectGProductKey,",,",",")
            	addSql = addSql & " and g.product_key in (" + FRectGProductKey + ")"
            End If
        End If

		If FRectGIscheck <> "" Then
			If (FRectGIscheck = "Y") Then
				addSql = addSql & " And g.product_key is not null "
			ElseIf (FRectGIscheck = "N") Then
				addSql = addSql & " And g.product_key is null "
			End If
		End If

		If FRectMarginSt <> "" And FRectMarginEd <> "" Then
			If isnumeric(FRectMarginSt) And isNumeric(FRectMarginEd) Then
				addSql = addSql & " And round((1-(i.orgsuplycash/i.orgprice))*100, 1) >= "&FRectMarginSt&" And round((1-(i.orgsuplycash/i.orgprice))*100, 1) <= "&FRectMarginEd&" "
			End If
		End If

		If FRectSorgpriceSt <> "" And FRectSorgpriceEd <> "" Then
			If isnumeric(FRectSorgpriceSt) And isNumeric(FRectSorgpriceEd) Then
				addSql = addSql & " And i.orgprice >= "&FRectSorgpriceSt&" And i.orgprice <= "&FRectSorgpriceEd&" "
			End If
		End If

		If FRectBaesongGubun <> "" Then
			If Trim(FRectBaesongGubun)="tenbae" Then
				addSql = addSql & " And i.mwdiv in ('M','W') "
			Else
				addSql = addSql & " And i.mwdiv in ('U') "
			End If
		End If

		strsql = ""
		strsql = strsql & " SELECT COUNT(i.itemid) "
		strsql = strsql & " FROM db_item.dbo.tbl_item i "
		strsql = strsql & " LEFT JOIN db_item.dbo.tbl_display_cate_item ci on i.itemid = ci.itemid And ci.isDefault='y' "
		If (FRectGIscheck = "Y") Then			'글로비 등록여부 Y일 때는 JOIN, 그 외는 LEFT JOIN
			strsql = strsql & " JOIN db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid "
			strsql = strsql & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		Else
			strsql = strsql & " LEFT JOIN db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid "
			strsql = strsql & " LEFT JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		End If
		strsql = strsql & " LEFT JOIN db_item.dbo.tbl_makeglob_product g on i.itemid = g.product_code "
		strsql = strsql & " LEFT JOIN db_item.[dbo].[tbl_const_OptAddPrice_Exists] as x on i.itemid = x.itemid "
'		strsql = strsql & " Where  i.deliverOverseas='Y' And i.itemweight<>0 And i.mwdiv in ('m','w') "&addSql '// 텐배 상품만 표시
		strsql = strsql & " WHERE 1 = 1 "
		If (FRectGIscheck = "N") Then			'글로비에 미등록
			strsql = strsql & " and i.deliverOverseas='Y' And i.itemweight<>0 "
			strsql = strsql & " and isnull(x.itemid, '') = '' " 
		End If
		strsql = strsql & addsql
		'strsql = strsql & " And i.itemid not in ( Select itemid From db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid ) " &addSql '// 업체배송 상품도 표시
        rsget.Open strsql,dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget(0)
        rsget.Close

		If FTotalCount < 1 Then Exit Function
		strsql = ""
		strSql = strSql & " SELECT TOP "&Cstr(FPageSize * FCurrPage)
		strSql = strSql & "		i.itemid, i.makerid, c.socname, c.socname_kor, i.itemname, i.sellcash,i.buycash, i.orgprice, "
		strSql = strSql & "		i.orgsuplycash, i.sailprice, i.sailsuplycash, i.mileage, i.regdate, i.lastupdate, i.sellStdate, i.sellyn, i.limityn, "
		strSql = strSql & "		i.sailyn, i.isusing, i.limitno, i.limitsold, i.mainimage, i.smallimage, i.listimage, i.listimage120, "
		strSql = strSql & "		i.basicimage, i.icon1image, i.icon2image, i.basicimage600, i.basicimage1000, i.itemcouponyn, i.itemscore, i.itemweight, i.deliverOverseas, i.tenonlyyn, "
		strSql = strSql & "		ci.catecode, ci.depth, ci.isDefault, ic.keywords, ic.sourcearea, ic.makername, ic.itemsource, ic.itemsize, ic.itemcontent, g.product_key, g.product_code, g.hidden, g.soldout, "
		strSql = strSql & "		g.makeGlobYN, g.makeupdate, i.mwdiv "
		strSql = strSql & "	FROM db_item.dbo.tbl_item i "
		strsql = strsql & " LEFT JOIN db_item.dbo.tbl_display_cate_item ci on i.itemid = ci.itemid And ci.isDefault='y' "
		If (FRectGIscheck = "Y") Then			'글로비 등록여부 Y일 때는 JOIN, 그 외는 LEFT JOIN
			strsql = strsql & " JOIN db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid "
			strsql = strsql & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		Else
			strsql = strsql & " LEFT JOIN db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid "
			strsql = strsql & " LEFT JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		End If
		strsql = strsql & " LEFT JOIN db_item.dbo.tbl_makeglob_product g on i.itemid = g.product_code "
		strsql = strsql & " LEFT JOIN db_item.[dbo].[tbl_const_OptAddPrice_Exists] as x on i.itemid = x.itemid "
'		strSql = strSql & "	Where  i.deliverOverseas='Y' And i.itemweight<>0 And i.mwdiv in ('m','w') "&addSql '// 텐배 상품만 표시
		strsql = strsql & " WHERE 1 = 1 "
		If (FRectGIscheck = "N") Then			'글로비에 미등록
			strsql = strsql & " and i.deliverOverseas='Y' And i.itemweight<>0 "
			strsql = strsql & " and isnull(x.itemid, '') = '' " 
		End If
		strsql = strsql & addsql
'		strsql = strsql & " And i.itemid not in ( Select itemid From db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid ) " &addSql '// 업체배송 상품도 표시
		strSql = strSql & "	order by itemid desc "
        rsget.pagesize = FPageSize
        rsget.Open strsql,dbget, 1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CMakeGlobItemDetail
				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Fmakerid = rsget("makerid")
				FItemList(i).FBrandName = rsget("socname")
				FItemList(i).FBrandNameKr = rsget("socname_kor")
				FItemList(i).FitemName = rsget("itemname")

				FItemList(i).FsellCash = rsget("sellcash")
				FItemList(i).FbuyCash = rsget("buycash")
				FItemList(i).Forgprice = rsget("orgprice")
				FItemList(i).Forgsuplycash = rsget("orgsuplycash")
				FItemList(i).Fsailprice = rsget("sailprice")
				FItemList(i).Fsailsuplycash = rsget("sailsuplycash")

				FItemList(i).Fmileage = rsget("mileage")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Flastupdate = rsget("lastupdate")
				FItemList(i).FsellStdate = rsget("sellStdate")
				FItemList(i).Fsellyn = rsget("sellyn")
				FItemList(i).Flimityn = rsget("limityn")
				FItemList(i).Fsailyn = rsget("sailyn")
				FItemList(i).Fisusing = rsget("isusing")
	
				FItemList(i).Flimitno = rsget("limitno")
				FItemList(i).Flimitsold = rsget("limitsold")

				FItemList(i).Fmainimage = webImgUrl&"/image/main/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("mainimage")
				FItemList(i).Fsmallimage = webImgUrl&"/image/small/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("smallimage")
				FItemList(i).Flistimage = webImgUrl&"/image/list/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("listimage")
				FItemList(i).Flistimage120 = webImgUrl&"/image/list120/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("listimage120")
				FItemList(i).Fbasicimage = webImgUrl&"/image/basic/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("basicimage")
				FItemList(i).Ficon1image = webImgUrl&"/image/icon1/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("icon1image")
				FItemList(i).Ficon2image = webImgUrl&"/image/icon2/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("icon2image")
				FItemList(i).Fbasic600image = webImgUrl&"/image/basic600/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("basicimage600")
				FItemList(i).Fbasic1000image = webImgUrl&"/image/basic1000/"&GetImageSubFolderByItemid(FItemList(i).Fitemid)&"/"&rsget("basicimage1000")
				FItemList(i).Fitemcouponyn = rsget("itemcouponyn")
				FItemList(i).Fitemscore = rsget("itemscore")
				FItemList(i).Fitemweight = rsget("itemweight")
				FItemList(i).FdeliverOverseas = rsget("deliverOverseas")
				FItemList(i).Ftenonlyyn = rsget("tenonlyyn")
				FItemList(i).Fcatecode = rsget("catecode")
				FItemList(i).Fdepth = rsget("depth")
				FItemList(i).FisDefault = rsget("isDefault")
				FItemList(i).Fkeywords = rsget("keywords")
				FItemList(i).Fsourcearea = rsget("sourcearea")
				FItemList(i).Fmakername = rsget("makername")
				FItemList(i).Fitemsource = rsget("itemsource")
				FItemList(i).Fitemsize = rsget("itemsize")
				FItemList(i).Fitemcontent = rsget("itemcontent")

				FItemList(i).FMakeGlobHidden = rsget("hidden")
				FItemList(i).FMakeGlobSoldout = rsget("soldout")
				FItemList(i).FMakeGlobProductKey = rsget("product_key")
				FItemList(i).FMakeGlobupdate = rsget("makeGlobYN")
				FItemList(i).FMakeGlobupdateTime = rsget("makeupdate")

				FItemList(i).FBaesongGubun = rsget("mwdiv")


                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close

	End Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
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
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class



Sub drawSelectBoxGHiddenYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >Y</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >N</option>
   </select>
   <%
End Sub

Sub drawSelectBoxGsoldoutYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >Y</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >N</option>
   </select>
   <%
End Sub

Sub drawSelectBoxGcheckYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >Y</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >N</option>
   </select>
   <%
End Sub


Function fnPercent(oup,inp,pnt)
	'' if oup=0 or isNull(oup) then exit function ''주석처리 2014/01/16
	if inp=0 or isNull(inp) then exit function
	fnPercent = FormatNumber((1-(clng(oup)/clng(inp)))*100,pnt) & "%"
End Function

%>