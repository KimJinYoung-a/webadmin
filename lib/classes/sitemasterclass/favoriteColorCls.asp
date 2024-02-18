<%
'#######################################################
'	History	:  2010.04.07 허진원 생성
'	Description : Favorite Color 관리
'#######################################################
%>
<%
Class CfavoriteColorItem

	public Fidx
	Public Fcategory
	public FcolorCD
	public Fitemid
	public Fisusing
	public FsortNo
	public FitemName
	public FImageSmall

	public FSellyn
	public FLimityn
	public FLimitno
	public FLimitsold

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CfavoriteColor
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	Public FRectCategory
	Public FRectColorCD
	Public FRectIsUsing
	public FRectStyleSerail

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function GetImageFolerName(byval i)
		'GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
		GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
	end function

	public Function GetfavoriteColor()
		dim sqlStr,addSql, i

		'추가 쿼리
		if FReCtcategory<>"" then
			addSql = addSql + " and c.category = '" + FRectCategory + "'" + vbcrlf
		end if
		if FReCtcategory<>"" then
			'addSql = addSql + " and c.colorCD = '" + FRectColorCD + "'" + vbcrlf
			addSql = addSql + " and d.colorcode = '" + FRectColorCD + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			addSql = addSql + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		'총수 접수
		sqlStr = "select count(c.idx), CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_favoriteColor c," + vbcrlf
		sqlStr = sqlStr + " db_sitemaster.dbo.tbl_favoriteColorCode d," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid and c.colorCD = d.colorcode " + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.category, d.coloricon as colorCD , c.itemid, c.isusing, i.itemname, i.smallimage, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_favoriteColor c," + vbcrlf
		sqlStr = sqlStr + " db_sitemaster.dbo.tbl_favoriteColorCode d," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid and c.colorCD = d.colorcode " + addSql + vbcrlf
		sqlStr = sqlStr + " order by c.sortNo, c.itemid desc"
		
		'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CfavoriteColorItem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fcategory		= rsget("category")
				FItemList(i).FcolorCD		= rsget("colorCD")
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).FitemName		= db2html(rsget("itemname"))
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).FsortNo		= rsget("sortNo")

				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

Class CItemColor '컬러코드 관리 2011-04-13 이종화
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

	public FcolorIcon
	

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function GetColorList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectColorCD <> "") then addSql = addSql & " and ColorCode =" + FRectColorCD
        if (FRectUsing <> "") then addSql = addSql & " and isUsing ='" + FRectUsing + "'"
        
		'// 결과수 카운트
		sqlStr = "select Count(colorCode), CEILING(CAST(Count(colorCode) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_favoriteColorCode "
        sqlStr = sqlStr & " where 1=1 " & addSql

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " colorCode, colorName, colorIcon, sortNo, isUsing "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_favoriteColorCode "
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
                FItemList(i).FcolorIcon	= "http://fiximage.10x10.co.kr/web2011/common/color/" & rsget("colorIcon")
                FItemList(i).FsortNo	= rsget("sortNo")
                FItemList(i).FisUsing	= rsget("isUsing")
                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function
	
end Class

Class CItemColorItem '컬러코드 관리 2011-04-13 이종화
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

Sub DrawSelectBoxCateTab(boxname,stats)
%>
<select name='<%=boxname%>'>
	<option>선택하세요</option>
	<option value=1 <% if stats = "1" then response.write " selected" %>>Stationary&Persnal</option>
	<option value=2 <% if stats = "2" then response.write " selected" %>>Home&Living</option>
	<option value=3 <% if stats = "3" then response.write " selected" %>>Fashion&Beauty</option>
	<option value=4 <% if stats = "4" then response.write " selected" %>>Kidult&Hobby</option>
	<option value=5 <% if stats = "5" then response.write " selected" %>>Kids&Baby</option>
</select>
<%
end Sub

Sub DrawSelectBoxColoreBar(boxname,stats)
	Dim i
%>
<script language="javascript">
function chgColorChip(ccd) {
	document.Listfrm.<%=boxname%>.value=ccd;
	for(var i=1;i<=20;i++) {
		if(i==ccd) {
			document.getElementById("tbColor"+i).style.backgroundColor="#000000";
		} else {
			document.getElementById("tbColor"+i).style.backgroundColor="#EDEDED";
		}
	}
}
</script>
<table border="0" cellspacing="3" cellpadding="0">
<tr>
<%
	For i=1 to 20
%>
	<td onClick="chgColorChip(<%=i%>)" style="cursor:pointer">
		<table id="tbColor<%=i%>" border="0" cellpadding="0" cellspacing="1" bgcolor="<% if cstr(stats)=cstr(i) then %>#000000<% else %>#EDEDED<% end if %>">
		<tr>
			<td bgcolor="#FFFFFF"><img src="http://fiximage.10x10.co.kr/web2010/favoritecolor/color_n<%=Num2Str(i,2,"0","R")%>.gif" width="15" height="15" hspace="1" vspace="1" border="0"></td>
		</tr>
		</table>
	</td>
<%
	Next
%>
<input type="hidden" name="<%=boxname%>" value="<%=stats%>">
</tr>
</table>
<%
End Sub
%>
