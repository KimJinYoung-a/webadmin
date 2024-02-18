<%
'####################################################
' Description :  ����ũ �۷κ� ���� Ŭ����
' History : 2015.10.27 ������ ����
'####################################################

'// ��ǰdetail Ŭ����
Class CMakeGlobItemDetail

	Public Fitemid '// ��ǰ�ڵ�
	Public Fmakerid '// �Ǹ��ھ��̵�
	Public FBrandName '// �귣���(����)
	Public FBrandNameKr '// �귣���(�ѱ�)
	Public FitemName '// ��ǰ��

	'------ �� ��ȣ���� �ѽ����� �����δٰ� ���� ��--------
	Public FsellCash '// ���� �ǸŰ�(1)
	Public FbuyCash '// ���� ���԰�(1)
	Public Forgprice '// ���� �ǸŰ�(2)
	Public Forgsuplycash '// �������԰�(2)
	Public Fsailprice '// ���ν� �ǸŰ�(3)
	Public Fsailsuplycash '// ���ν� ���԰�(3)
	'------ �� ��ȣ���� �ѽ����� �����δٰ� ���� ��--------

	Public Fmileage '// ���ϸ���
	Public Fregdate '// �����
	Public Flastupdate '// ����������
	Public FsellStdate '// �ǸŰ�����
	Public Fsellyn '// �Ǹſ���
	Public Flimityn '// ������ǰ����
	Public Fsailyn '// ���Ͽ���
	Public Fisusing '// ��ǰ��뿩��
	
	'// ���� ��������=Flimitno - Flimitsold
	Public Flimitno '// ��������
	Public Flimitsold '// �����ǸŰ���


	Public Fmainimage '// ��ǰ �̹��� �߿� �ϳ��ε� ���� �Ⱦ�
	Public Fsmallimage '// 50x50 �̹���
	Public Flistimage '// 100x100 �̹���
	Public Flistimage120 '// 120x120 �̹���
	Public Fbasicimage '// 400x400 �̹���
	Public Ficon1image '// 200x200 �̹���
	Public Ficon2image '// 150x150 �̹���
	Public Fitemcouponyn '// ������뿩��
	Public Fbasic600image '// 600x600 �̹���
	Public Fbasic1000image '// 1000x1000 �̹���
	Public Fitemscore '// ��ǰ����(best �� �Ǵ��Ҷ�)
	Public Fitemweight '// ��ǰ����(�۷κ� ǥ���Ҷ� /1000 �ؼ� �׶����� ��ߵ�)
	Public FdeliverOverseas '// �ؿܹ�ۿ���(�۷κ�� Y�� �ҷ���)
	Public Ftenonlyyn '// �ٹ����� ������ǰ����
	Public Fcatecode '// ī���ڸ� �ڵ�
	Public Fdepth '// ī�װ� ����
	Public FisDefault '// ī�װ� ǥ�� �� ���� �⺻������ ���Ǵ� ��(y�� �ش� ī�װ��� ����Ʈ�� ���Ǿ����� ����)
	Public Fkeywords '// �ش� ��ǰ Ű����
	Public Fsourcearea '// ������
	Public Fmakername '// ������
	Public Fitemsource '// ��ǰ ���
	Public Fitemsize '// ��ǰũ�Ⱚ
	Public Fitemcontent '// ��ǰ�󼼼���

	Public FMakeGlobChkEN '// ����ũ �۷κ������� ��ǰ�� �Ѱ� ������ ����(����)
	Public FMakeGlobChkZH '// ����ũ �۷κ������� ��ǰ�� �Ѱ� ������ ����(�߹�)

	Public FMakeGlobHidden '// ����ũ �۷κ� ���迩��
	Public FMakeGlobSoldout '// ����ũ �۷κ� ǰ������
	Public FMakeGlobProductKey '// ����ũ �۷κ� ��ǰ�ڵ�
	Public FMakeGlobupdate '// ����ũ �۷κ� ������Ʈ ����
	Public FMakeGlobupdateTime '// ����ũ �۷κ� ������Ʈ ����

	Public FBaesongGubun '// ��۱���(M,W - �ٺ�, U-����)

	'// �ֵ�ƿ� ����
    public Function IsSoldOut()
		IsSoldOut = (Fsellyn<>"Y") or ((Flimityn="Y") and (GetLimitEa()<1))
	end function

	'// ���� ��ǰ�� ��� ���� ��������
    public function GetLimitEa()
		if Flimitno-Flimitsold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = Flimitno-Flimitsold
		end if
	end function

	'// �������� ��������
	public Function getRealPrice() '!
		getRealPrice = FsellCash
	End Function
End Class


Class CMakeGlobItem
	public FItemList()
	public FTotalCount '// �Ѱ���
	public FCurrPage '// ���������� ��ȣ
	public FTotalPage '// �� ������ ����
	public FPageSize '// ������ ������
	public FResultCount '// ����� ����
	Public FScrollCount '// ��ũ��ī��Ʈ
	Public FRectBrandName '// �귣���
	Public FRectCateCode '// ī�װ� �ڵ�
	Public FRectItemName '// ��ǰ��
	Public FRectItemId '// ��ǰ�ڵ�
	Public FRectSellyn '// �ٹ����� �Ǹſ���(N-ǰ��, S-�Ͻ�ǰ��, Y-�Ǹ���)
	Public FRectLimityn '// �ٹ����� �����Ǹſ���
	Public FRectIsUsing '// �ٹ����� ��뿩��
	Public FRectGIsHidden '// �۷κ� ���迩��
	Public FRectGIssoldout '// �۷κ� ǰ������
	Public FRectGProductKey '// �۷κ� ��ǰ�ڵ�
	Public FRectGIscheck '// �۷κ� ��ǰ��Ͽ���
	Public FRectMakeGlobChkEN '// ���� �Է¿���
	Public FRectMakeGlobChkZH '// �߹� �Է¿���
	Public FRectMarginSt	'// �������˻����۰�
	Public FRectMarginEd	'// �������˻����ᰪ
	Public FRectSorgpriceSt	'// �ǸŰ��˻����۰�
	Public FRectSorgpriceEd	'// �ǸŰ��˻����ᰪ
	Public FRectBaesongGubun '// ��۱���(MW-�ٹ�, U-����)
	Public FRectMakerID '// ����Ŀid

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

		'��ǰ�ڵ� �˻�
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

		'�۷κ� ��ǰ��ȣ �˻�
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
		If (FRectGIscheck = "Y") Then			'�۷κ� ��Ͽ��� Y�� ���� JOIN, �� �ܴ� LEFT JOIN
			strsql = strsql & " JOIN db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid "
			strsql = strsql & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		Else
			strsql = strsql & " LEFT JOIN db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid "
			strsql = strsql & " LEFT JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		End If
		strsql = strsql & " LEFT JOIN db_item.dbo.tbl_makeglob_product g on i.itemid = g.product_code "
		strsql = strsql & " LEFT JOIN db_item.[dbo].[tbl_const_OptAddPrice_Exists] as x on i.itemid = x.itemid "
'		strsql = strsql & " Where  i.deliverOverseas='Y' And i.itemweight<>0 And i.mwdiv in ('m','w') "&addSql '// �ٹ� ��ǰ�� ǥ��
		strsql = strsql & " WHERE 1 = 1 "
		If (FRectGIscheck = "N") Then			'�۷κ� �̵��
			strsql = strsql & " and i.deliverOverseas='Y' And i.itemweight<>0 "
			strsql = strsql & " and isnull(x.itemid, '') = '' " 
		End If
		strsql = strsql & addsql
		'strsql = strsql & " And i.itemid not in ( Select itemid From db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid ) " &addSql '// ��ü��� ��ǰ�� ǥ��
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
		If (FRectGIscheck = "Y") Then			'�۷κ� ��Ͽ��� Y�� ���� JOIN, �� �ܴ� LEFT JOIN
			strsql = strsql & " JOIN db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid "
			strsql = strsql & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		Else
			strsql = strsql & " LEFT JOIN db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid "
			strsql = strsql & " LEFT JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		End If
		strsql = strsql & " LEFT JOIN db_item.dbo.tbl_makeglob_product g on i.itemid = g.product_code "
		strsql = strsql & " LEFT JOIN db_item.[dbo].[tbl_const_OptAddPrice_Exists] as x on i.itemid = x.itemid "
'		strSql = strSql & "	Where  i.deliverOverseas='Y' And i.itemweight<>0 And i.mwdiv in ('m','w') "&addSql '// �ٹ� ��ǰ�� ǥ��
		strsql = strsql & " WHERE 1 = 1 "
		If (FRectGIscheck = "N") Then			'�۷κ� �̵��
			strsql = strsql & " and i.deliverOverseas='Y' And i.itemweight<>0 "
			strsql = strsql & " and isnull(x.itemid, '') = '' " 
		End If
		strsql = strsql & addsql
'		strsql = strsql & " And i.itemid not in ( Select itemid From db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid ) " &addSql '// ��ü��� ��ǰ�� ǥ��
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
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >Y</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >N</option>
   </select>
   <%
End Sub

Sub drawSelectBoxGsoldoutYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >Y</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >N</option>
   </select>
   <%
End Sub

Sub drawSelectBoxGcheckYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >Y</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >N</option>
   </select>
   <%
End Sub


Function fnPercent(oup,inp,pnt)
	'' if oup=0 or isNull(oup) then exit function ''�ּ�ó�� 2014/01/16
	if inp=0 or isNull(inp) then exit function
	fnPercent = FormatNumber((1-(clng(oup)/clng(inp)))*100,pnt) & "%"
End Function

%>