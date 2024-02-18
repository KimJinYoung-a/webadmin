<%
Class CItemColorItem
    public FcolorCode
    public FcolorName
    public FcolorIcon
    public FsortNo
    public FisUsing
    public FitemId
    public FitemName
    public FmakerId
    public Fregdate
    public FsmallImage
    public FlistImage
    public Fsellyn
    public Flimityn
    public Fmwdiv

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CItemColor
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectColorCD
	public FRectItemId
	public FRectMakerId
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectUsing

	public function GetColorList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectColorCD <> "") then addSql = addSql & " and ColorCode =" & FRectColorCD
        if (FRectUsing <> "") then addSql = addSql & " and isUsing ='" + FRectUsing + "'"

		'// 결과수 카운트
		sqlStr = "select Count(colorCode), CEILING(CAST(Count(colorCode) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips "
        sqlStr = sqlStr & " where 1=1 " & addSql

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " colorCode, colorName, colorIcon, sortNo, isUsing "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips "
        sqlStr = sqlStr & " where 1 = 1 " & addSql
		sqlStr = sqlStr & " Order by sortNo "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemColorItem

                FItemList(i).FcolorCode	= rsget("colorCode")
                FItemList(i).FcolorName	= rsget("colorName")
                FItemList(i).FcolorIcon	= webImgUrl & "/color/colorchip/" & rsget("colorIcon")
                FItemList(i).FsortNo	= rsget("sortNo")
                FItemList(i).FisUsing	= rsget("isUsing")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public function GetColorItemList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectColorCD <> "") then addSql = addSql & " and C.ColorCode =" + FRectColorCD
        if (FRectItemId <> "") then addSql = addSql & " and O.itemid =" + FRectItemId
        if (FRectMakerId <> "") then addSql = addSql & " and I.makerid ='" + FRectMakerId + "'"
        if (FRectCDL <> "") then addSql = addSql & " and I.cate_large ='" + FRectCDL + "'"
        if (FRectCDM <> "") then addSql = addSql & " and I.cate_mid ='" + FRectCDM + "'"
        if (FRectCDS <> "") then addSql = addSql & " and I.cate_small ='" + FRectCDS+ "'"
        if (FRectUsing <> "") then addSql = addSql & " and C.isUsing ='" + FRectUsing + "'"

		'// 결과수 카운트
		sqlStr = "select Count(C.colorCode), CEILING(CAST(Count(C.colorCode) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips as C "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item_colorOption as O "
        sqlStr = sqlStr & " 		on C.colorCode=O.colorCode "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item as I "
        sqlStr = sqlStr & " 		on O.itemid=I.itemid "
        sqlStr = sqlStr & " where 1=1 " & addSql

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & "		C.colorCode, C.colorName, C.colorIcon "
        sqlStr = sqlStr & "		,O.itemid, O.smallimage, O.listimage, O.regdate "
        sqlStr = sqlStr & "		,I.itemname, I.makerid, I.sellyn, I.limityn, I.mwdiv "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips as C "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item_colorOption as O "
        sqlStr = sqlStr & " 		on C.colorCode=O.colorCode "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item as I "
        sqlStr = sqlStr & " 		on O.itemid=I.itemid "
        sqlStr = sqlStr & " where 1 = 1 " & addSql
		sqlStr = sqlStr & " Order by O.regdate desc "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemColorItem

                FItemList(i).FcolorCode	= rsget("colorCode")
                FItemList(i).FcolorName	= rsget("colorName")
                FItemList(i).FcolorIcon	= webImgUrl & "/color/colorchip/" & rsget("colorIcon")
                FItemList(i).FitemId	= rsget("itemid")
                FItemList(i).FitemName	= rsget("itemname")
                FItemList(i).FmakerId	= rsget("makerid")
                FItemList(i).FsmallImage= rsget("smallimage")
                FItemList(i).FlistImage	= rsget("listimage")
                FItemList(i).Fsellyn	= rsget("sellyn")
                FItemList(i).Flimityn	= rsget("limityn")
                FItemList(i).Fmwdiv		= rsget("mwdiv")
                FItemList(i).Fregdate	= rsget("regdate")

				if ((Not IsNULL(FItemList(i).Fsmallimage)) and (FItemList(i).Fsmallimage<>"")) then FItemList(i).Fsmallimage    = webImgUrl & "/color/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Fsmallimage
				if ((Not IsNULL(FItemList(i).Flistimage)) and (FItemList(i).Flistimage<>"")) then FItemList(i).Flistimage    = webImgUrl & "/color/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Flistimage

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
    End Sub
end Class

'// 컬러칩 선택바 생성함수
Function FnSelectColorBar(icd,colSize)
	Dim oClr, tmpStr, lineCr, lp
	set oClr = new CItemColor
	oClr.FPageSize = 31
	oClr.FRectUsing = "Y"
	oClr.GetColorList

	if icd="" then lineCr = "#DD3300": else lineCr = "#dddddd": end if
	tmpStr = "<table class='a'>" &_
			"<tr>" &_
			"	<td rowspan='" & (oClr.FResultCount\colSize)+1 & "'></td>" &_
			"<td>" &_
			"	<table id='cline0' border='0' cellpadding='0' cellspacing='1' bgcolor='" & lineCr & "'>" &_
			"	<tr>" &_
			"		<td bgcolor='#FFFFFF'><a href=""javascript:selColorChip('')"" onfocus='this.blur()'><img src='" & fixImgUrl & "/web2009/common/color01_n00.gif' alt='전체' width='12' height='12' hspace='2' vspace='2' border='0'></a></td>" &_
			"	</tr>" &_
			"	</table>" &_
			"</td>"
	if oClr.FResultCount>0 then
		for lp=0 to oClr.FResultCount-1
			if icd=cStr(oClr.FItemList(lp).FcolorCode) then lineCr = "#DD3300": else lineCr = "#dddddd": end if
			tmpStr = tmpStr &_
				"<td>" &_
				"	<table id='cline" & oClr.FItemList(lp).FcolorCode & "' border='0' cellpadding='0' cellspacing='1' bgcolor='" & lineCr & "'>" &_
				"	<tr>" &_
				"		<td bgcolor='#FFFFFF'><a href='javascript:selColorChip(" & oClr.FItemList(lp).FcolorCode & ")' onfocus='this.blur()'><img src='" & oClr.FItemList(lp).FcolorIcon & "' alt='" & oClr.FItemList(lp).FcolorName & "' width='12' height='12' hspace='2' vspace='2' border='0'></a></td>" &_
				"	</tr>" &_
				"	</table>" &_
				"</td>"
			'//행구분
			if ((lp+1) mod colSize)=(colSize-1) and (lp+1)<oClr.FResultCount then
				tmpStr = tmpStr & "</tr><tr>"
			end if
		next
	end if
	tmpStr = tmpStr & "</tr></table>"
	set oClr = Nothing

	FnSelectColorBar = tmpStr
End Function

Function FnSelectColorBarMo(icd,colSize)
	Dim oClr, tmpStr, lineCr, lp
	set oClr = new CItemColor
	oClr.FPageSize = 31
	oClr.FRectUsing = "Y"
	oClr.GetColorList

	if icd="" then lineCr = "#DD3300": else lineCr = "#dddddd": end if
	tmpStr = "<table class='a'>" &_
			"<tr>" &_
			"	<td rowspan='" & (oClr.FResultCount\colSize)+1 & "'></td>" &_
			"<td>" &_
			"	<table id='mocline0' border='0' cellpadding='0' cellspacing='1' bgcolor='" & lineCr & "'>" &_
			"	<tr>" &_
			"		<td bgcolor='#FFFFFF'><a href=""javascript:selMoColorChip('')"" onfocus='this.blur()'><img src='" & fixImgUrl & "/web2009/common/color01_n00.gif' alt='전체' width='12' height='12' hspace='2' vspace='2' border='0'></a></td>" &_
			"	</tr>" &_
			"	</table>" &_
			"</td>"
	if oClr.FResultCount>0 then
		for lp=0 to oClr.FResultCount-1
			if icd=cStr(oClr.FItemList(lp).FcolorCode) then lineCr = "#DD3300": else lineCr = "#dddddd": end if
			tmpStr = tmpStr &_
				"<td>" &_
				"	<table id='mocline" & oClr.FItemList(lp).FcolorCode & "' border='0' cellpadding='0' cellspacing='1' bgcolor='" & lineCr & "'>" &_
				"	<tr>" &_
				"		<td bgcolor='#FFFFFF'><a href='javascript:selMoColorChip(" & oClr.FItemList(lp).FcolorCode & ")' onfocus='this.blur()'><img src='" & oClr.FItemList(lp).FcolorIcon & "' alt='" & oClr.FItemList(lp).FcolorName & "' width='12' height='12' hspace='2' vspace='2' border='0'></a></td>" &_
				"	</tr>" &_
				"	</table>" &_
				"</td>"
			'//행구분
			if ((lp+1) mod colSize)=(colSize-1) and (lp+1)<oClr.FResultCount then
				tmpStr = tmpStr & "</tr><tr>"
			end if
		next
	end if
	tmpStr = tmpStr & "</tr></table>"
	set oClr = Nothing

	FnSelectColorBarMo = tmpStr
End Function
%>