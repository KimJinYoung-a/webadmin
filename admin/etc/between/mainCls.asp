<%
Class cMainOneItem
	Public FIdx
	Public FGender
	Public FImgurl
	Public FImglink
	Public FSortno		
	Public FStartdate		
	Public FEnddate		
	Public FRegdate		
	Public FLastupdate	
	Public FAdminid		
	Public FLastadminid	
	Public FIsusing		
	Public FXmlregdate

	Public FItemid
	Public FsmallImage
	Public FSellyn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FItemname
	Public FMakerid
	Public FdefaultfreeBeasongLimit
	Public Fdeliverytype
	Public FSellcash
	Public FOrgprice
	Public FsaleYn
	Public FMainMdpickSortNo
	Public FMainMdpickXMLRegdate

	Public FPjt_kind
	Public FLinkURL
	Public FBanBGColor
	Public FPartnerNmColor
	Public FBanTxtColor
	Public FBantext1
	Public FBantext2

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[업체착불]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[업체]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function
End Class

Class cMain
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount

	Public FRectIdx
	Public FRectIsusing
	Public FRectGender

	Public Sub Get3BannerList()
		Dim sqlStr, i, addSql

		If FRectIsusing <> "" Then
			addsql = addsql & " and isusing = '"&FRectIsusing&"' " 
		End If

		If FRectGender <> "" Then
			addsql = addsql & " and gender = '"&FRectGender&"' " 
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_main_3banner "
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " idx, imgurl, gender, imglink, sortno, startdate, enddate, regdate, lastupdate, adminid, lastadminid, isusing, xmlregdate "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_main_3banner "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cMainOneItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FImgurl		= rsCTget("imgurl")
					FItemList(i).FGender		= rsCTget("gender")
					FItemList(i).FImglink		= rsCTget("imglink")
					FItemList(i).FSortno		= rsCTget("sortno")
					FItemList(i).FStartdate		= rsCTget("startdate")
					FItemList(i).FEnddate		= rsCTget("enddate")
					FItemList(i).FRegdate		= rsCTget("regdate")
					FItemList(i).FLastupdate	= rsCTget("lastupdate")
					FItemList(i).FAdminid		= rsCTget("adminid")
					FItemList(i).FLastadminid	= rsCTget("lastadminid")
					FItemList(i).FIsusing		= rsCTget("isusing")
					FItemList(i).FXmlregdate	= rsCTget("xmlregdate")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub GetOne3Banner()
		Dim sqlStr
		IF FRectIdx = "" THEN Exit Sub
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, gender, imgurl, imglink, sortno, startdate, enddate, regdate, lastupdate, adminid, lastadminid, isusing, xmlregdate "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_main_3banner "
		sqlStr = sqlStr & " WHERE idx = " & FRectIdx
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		If not rsCTget.EOF Then
			Set FItemList(0) = new cMainOneItem
				FItemList(0).FIdx			= rsCTget("idx")
				FItemList(0).FGender		= rsCTget("gender")
				FItemList(0).FImgurl		= rsCTget("imgurl")
				FItemList(0).FImglink		= rsCTget("imglink")
				FItemList(0).FSortno		= rsCTget("sortno")
				FItemList(0).FStartdate		= rsCTget("startdate")
				FItemList(0).FEnddate		= rsCTget("enddate")
				FItemList(0).FRegdate		= rsCTget("regdate")
				FItemList(0).FLastupdate	= rsCTget("lastupdate")
				FItemList(0).FAdminid		= rsCTget("adminid")
				FItemList(0).FLastadminid	= rsCTget("lastadminid")
				FItemList(0).FIsusing		= rsCTget("isusing")
				FItemList(0).FXmlregdate	= rsCTget("xmlregdate")
		End If
		rsCTget.Close
	End Sub

	Public Sub GetMdpickList()
		Dim sqlStr, i, addSql
		If FRectGender <> "" Then
			addsql = addsql & " and p.pjt_gender = '"&FRectGender&"' " 
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as CNT, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.[dbo].tbl_between_project as p "
		sqlStr = sqlStr & " JOIN db_outmall.dbo.tbl_between_project_groupItem as pg on p.pjt_code = pg.pjt_code "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on pg.itemid = i.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_AppWish.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and p.pjt_kind = 'I'  " & addSql
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " i.*, "
		sqlStr = sqlStr & " p.pjt_gender, uc.defaultfreeBeasongLimit, pg.idx, isnull(pg.MainMdpickSortNo, 0) as MainMdpickSortNo, pg.MainMdpickXMLRegdate "
		sqlStr = sqlStr & " FROM db_outmall.[dbo].tbl_between_project as p "
		sqlStr = sqlStr & " JOIN db_outmall.dbo.tbl_between_project_groupItem as pg on p.pjt_code = pg.pjt_code "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on pg.itemid = i.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_AppWish.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and p.pjt_kind = 'I'  " & addSql
		sqlStr = sqlStr & " order by pg.mdpickIsUsing desc, pg.MainMdpickSortNo ASC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cMainOneItem
					FItemList(i).FItemid					= rsCTget("itemid")
					FItemList(i).FGender					= rsCTget("pjt_gender")
					FItemList(i).FsmallImage				= rsCTget("smallImage")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsCTget("itemid")) & "/" & rsCTget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
					FItemList(i).FSellyn					= rsCTget("sellyn")
					FItemList(i).FLimitYn					= rsCTget("limityn")
					FItemList(i).FLimitNo					= rsCTget("limitNo")
					FItemList(i).FLimitSold					= rsCTget("limitSold")
					FItemList(i).Fdeliverytype				= rsCTget("deliverytype")
					FItemList(i).FItemname					= rsCTget("itemname")
					FItemList(i).FMakerid					= rsCTget("makerid")
					FItemList(i).FdefaultfreeBeasongLimit	= rsCTget("defaultfreeBeasongLimit")
					FItemList(i).FSellcash					= rsCTget("sellcash")
					FItemList(i).FOrgprice					= rsCTget("orgprice")
					FItemList(i).FsaleYn					= rsCTget("sailyn")
					FItemList(i).Fidx						= rsCTget("idx")
					FItemList(i).FMainMdpickSortNo			= rsCTget("MainMdpickSortNo")
					FItemList(i).FMainMdpickXMLRegdate		= rsCTget("MainMdpickXMLRegdate")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getTopBannerList
		Dim sqlStr, i, addSql

		If FRectIsusing <> "" Then
			addsql = addsql & " and isusing = '"&FRectIsusing&"' " 
		End If

		If FRectGender <> "" Then
			addsql = addsql & " and gender = '"&FRectGender&"' " 
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_main_topbanner "
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " idx, gender, pjt_kind, imgurl, linkURL, BanBGColor, partnerNmColor, BanTxtColor, bantext1, bantext2, regdate, lastupdate, adminid, lastadminid, isusing, xmlregdate "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_main_topbanner "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cMainOneItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FGender		= rsCTget("gender")
					FItemList(i).FPjt_kind		= rsCTget("pjt_kind")
					FItemList(i).FImgurl		= rsCTget("imgurl")
					FItemList(i).FLinkURL		= rsCTget("linkURL")
					FItemList(i).FBanBGColor	= rsCTget("BanBGColor")
					FItemList(i).FPartnerNmColor= rsCTget("partnerNmColor")
					FItemList(i).FBanTxtColor	= rsCTget("BanTxtColor")
					FItemList(i).FBantext1		= rsCTget("bantext1")
					FItemList(i).FBantext2		= rsCTget("bantext2")
					FItemList(i).FRegdate		= rsCTget("regdate")
					FItemList(i).FLastupdate	= rsCTget("lastupdate")
					FItemList(i).FAdminid		= rsCTget("adminid")
					FItemList(i).FLastadminid	= rsCTget("lastadminid")
					FItemList(i).FIsusing		= rsCTget("isusing")
					FItemList(i).FXmlregdate	= rsCTget("xmlregdate")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getOneTopBanner()
		Dim sqlStr
		IF FRectIdx = "" THEN Exit Sub
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, gender, pjt_kind, imgurl, linkURL, BanBGColor, partnerNmColor, BanTxtColor, bantext1, bantext2, regdate, lastupdate, adminid, lastadminid, isusing, xmlregdate "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_main_topbanner "
		sqlStr = sqlStr & " WHERE idx = " & FRectIdx
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		If not rsCTget.EOF Then
			Set FItemList(0) = new cMainOneItem
				FItemList(0).FIdx			= rsCTget("idx")
				FItemList(0).FGender		= rsCTget("gender")
				FItemList(0).FPjt_kind		= rsCTget("pjt_kind")
				FItemList(0).FImgurl		= rsCTget("imgurl")
				FItemList(0).FLinkURL		= rsCTget("linkURL")
				FItemList(0).FBanBGColor	= rsCTget("BanBGColor")
				FItemList(0).FPartnerNmColor= rsCTget("partnerNmColor")
				FItemList(0).FBanTxtColor	= rsCTget("BanTxtColor")
				FItemList(0).FBantext1		= rsCTget("bantext1")
				FItemList(0).FBantext2		= rsCTget("bantext2")
				FItemList(0).FRegdate		= rsCTget("regdate")
				FItemList(0).FLastupdate	= rsCTget("lastupdate")
				FItemList(0).FAdminid		= rsCTget("adminid")
				FItemList(0).FLastadminid	= rsCTget("lastadminid")
				FItemList(0).FIsusing		= rsCTget("isusing")
				FItemList(0).FXmlregdate	= rsCTget("xmlregdate")
		End If
		rsCTget.Close
	End Sub

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

'// STAFF 이름 접수
public Function getStaffUserName(uid)
	if uid="" or isNull(uid) then
		exit Function
	end if

	Dim strSql
	strSql = "Select top 1 username From db_partner.dbo.tbl_user_tenbyten Where userid='" & uid & "'"
	rsget.Open strSql, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		getStaffUserName = rsget("username")
	end if
	rsget.Close
End Function

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
		Case "K"	response.write "200일"
		Case "L"	response.write "500일"
		Case "M"	response.write "1000일"
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
	</select>
<%
End Sub
%>
