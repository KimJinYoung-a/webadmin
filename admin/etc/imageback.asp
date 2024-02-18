<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Class ImageItem
	public FItemID
	public FImageUrl

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class ImageBack
	public FItemList()

	public FResultCount
	public FTotalCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectImageType
	public FRectStartNo
	public FRectEndNo

	public Sub GetImageItem()
		dim sqlStr,i,splited
		sqlStr = " select count(i.itemid) as cnt from"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item_image i"
		sqlStr = sqlStr + " where i.itemid>=" + CStr(FRectStartNo)
		sqlStr = sqlStr + " and i.itemid<=" + CStr(FRectEndNo)

		if FRectImageType="small" then
			sqlStr = sqlStr + " and i.imgsmall is Not NULL"
			sqlStr = sqlStr + " and i.imgsmall<>''"
		elseif FRectImageType="list" then
			sqlStr = sqlStr + " and i.imglist is Not NULL"
			sqlStr = sqlStr + " and i.imglist<>''"
		elseif FRectImageType="main" then
			sqlStr = sqlStr + " and i.imgmain is Not NULL"
			sqlStr = sqlStr + " and i.imgmain<>''"
		elseif FRectImageType="title" then
			sqlStr = sqlStr + " and i.imgtitle is Not NULL"
			sqlStr = sqlStr + " and i.imgtitle<>''"
		elseif Left(FRectImageType,3)="add" then
			sqlStr = sqlStr + " and i.imgadd is Not NULL"
			sqlStr = sqlStr + " and i.imgadd<>''"
			sqlStr = sqlStr + " and i.imgadd<>',,,,'"
		elseif Left(FRectImageType,5)="story" then
			sqlStr = sqlStr + " and i.imgstory is Not NULL"
			sqlStr = sqlStr + " and i.imgstory<>''"
			sqlStr = sqlStr + " and i.imgstory<>',,,,'"
		else
			sqlStr = sqlStr + " and i.imgsmall is Not NULL"
			sqlStr = sqlStr + " and i.imgsmall<>''"
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)

		if FRectImageType="small" then
			sqlStr = sqlStr + " i.imgsmall as imgurl"
		elseif FRectImageType="list" then
			sqlStr = sqlStr + " i.imglist as imgurl"
		elseif FRectImageType="main" then
			sqlStr = sqlStr + " i.imgmain as imgurl"
		elseif FRectImageType="title" then
			sqlStr = sqlStr + " i.imgtitle as imgurl"
		elseif Left(FRectImageType,3)="add" then
			sqlStr = sqlStr + " i.imgadd as imgurl"
		elseif Left(FRectImageType,5)="story" then
			sqlStr = sqlStr + " i.imgstory as imgurl"
		else
			sqlStr = sqlStr + " i.imgsmall as imgurl"
		end if


		sqlStr = sqlStr + " , i.itemid from [db_item].[dbo].tbl_item_image i"
		sqlStr = sqlStr + " where i.itemid>=" + CStr(FRectStartNo)
		sqlStr = sqlStr + " and i.itemid<=" + CStr(FRectEndNo)

		if FRectImageType="small" then
			sqlStr = sqlStr + " and i.imgsmall is Not NULL"
			sqlStr = sqlStr + " and i.imgsmall<>''"
		elseif FRectImageType="list" then
			sqlStr = sqlStr + " and i.imglist is Not NULL"
			sqlStr = sqlStr + " and i.imglist<>''"
		elseif FRectImageType="main" then
			sqlStr = sqlStr + " and i.imgmain is Not NULL"
			sqlStr = sqlStr + " and i.imgmain<>''"
		elseif FRectImageType="title" then
			sqlStr = sqlStr + " and i.imgtitle is Not NULL"
			sqlStr = sqlStr + " and i.imgtitle<>''"
		elseif Left(FRectImageType,3)="add" then
			sqlStr = sqlStr + " and i.imgadd is Not NULL"
			sqlStr = sqlStr + " and i.imgadd<>''"
			sqlStr = sqlStr + " and i.imgadd<>',,,,'"
		elseif Left(FRectImageType,5)="story" then
			sqlStr = sqlStr + " and i.imgstory is Not NULL"
			sqlStr = sqlStr + " and i.imgstory<>''"
			sqlStr = sqlStr + " and i.imgstory<>',,,,'"
		else
			sqlStr = sqlStr + " and i.imgsmall is Not NULL"
			sqlStr = sqlStr + " and i.imgsmall<>''"
		end if

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new ImageItem
				FItemList(i).FItemID = rsget("itemid")
				if FRectImageType="small" then
					FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgurl")
				elseif FRectImageType="list" then
					FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgurl")
				elseif FRectImageType="main" then
					FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgurl")
				elseif FRectImageType="title" then
					FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/title/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgurl")
				elseif Left(FRectImageType,3)="add" then
					splited = rsget("imgurl")
					splited = split(splited,",")

					if FRectImageType="add1" then
						if splited(0)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/add1/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(0)
						end if
					elseif FRectImageType="add2" then
						if splited(1)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/add2/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(1)
						end if
					elseif FRectImageType="add3" then
						if  splited(2)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/add3/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(2)
						end if
					elseif FRectImageType="add4" then
						if  splited(3)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/add4/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(3)
						end if
					elseif FRectImageType="add5" then
						if  splited(4)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/add5/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(4)
						end if
					end if

				elseif Left(FRectImageType,5)="story" then
					splited = rsget("imgurl")
					splited = split(splited,",")

					if FRectImageType="story1" then
						if splited(0)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/story1/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(0)
						end if
					elseif FRectImageType="story2" then
						if splited(0)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/story2/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(1)
						end if
					elseif FRectImageType="story3" then
						if splited(0)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/story3/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(2)
						end if
					elseif FRectImageType="story4" then
						if splited(0)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/story4/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(3)
						end if
					elseif FRectImageType="story5" then
						if splited(0)<>"" then
							FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/story5/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + splited(4)
						end if
					end if
				else
					FItemList(i).FImageUrl = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgurl")
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub


	Private Sub Class_Initialize()
'		redim preserve FItemList(0)
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

	End Sub

	Private Sub Class_Terminate()

	End Sub

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


dim imgtype, startno, endno, page
imgtype = request("imgtype")
startno = request("startno")
endno = request("endno")
page = request("page")

if page="" then page=1

dim i, obackImage

set obackImage = new ImageBack
obackImage.FCurrPage = page
obackImage.FPageSize = 1001
obackImage.FRectImageType   = imgtype
obackImage.FRectStartNo     = startno
obackImage.FRectEndNo       = endno

if (imgtype<>"") and (startno<>"") and (endno<>"") then
	obackImage.GetImageItem
end if
%>
<script language='javascript'>
var errlist = "";
function AddErrLisr(v){
	errlist = errlist + v + "\r\n";
}

function writelist(){
txt.value = errlist;
}
</script>
<body onload="writelist()">
<table width="760" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="170">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		이미지구분 :
	<select name='imgtype'>
	<option value=small <% if imgtype="small" then response.write "selected" %> >small</option>
	<option value=list <% if imgtype="list" then response.write "selected" %> >list</option>
	<option value=main <% if imgtype="main" then response.write "selected" %> >main</option>
	<option value=add1 <% if imgtype="add1" then response.write "selected" %> >add1</option>
	<option value=add2 <% if imgtype="add2" then response.write "selected" %> >add2</option>
	<option value=add3 <% if imgtype="add3" then response.write "selected" %> >add3</option>
	<option value=add4 <% if imgtype="add4" then response.write "selected" %> >add4</option>
	<option value=add5 <% if imgtype="add5" then response.write "selected" %> >add5</option>
	<option value=story1 <% if imgtype="story1" then response.write "selected" %> >story1</option>
	<option value=story2 <% if imgtype="story2" then response.write "selected" %> >story2</option>
	<option value=story3 <% if imgtype="story3" then response.write "selected" %> >story3</option>
	<option value=story4 <% if imgtype="story4" then response.write "selected" %> >story4</option>
	<option value=story5 <% if imgtype="story5" then response.write "selected" %> >story5</option>
	<option value=title <% if imgtype="title" then response.write "selected" %> >title</option>
	</select>
	아이템번호
	<input type=text name=startno value="<%= startno %>" size=6 >
	~
	<input type=text name=endno value="<%= endno %>" size=6>

		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="760" border="0" cellpadding="5" cellspacing="0" >
<tr><td>
	<textarea name=txt cols=100 rows=10></textarea>
</td></tr>
</table>

<table width="760" border="0" cellpadding="5" cellspacing="0" >
<% for i=0 to obackImage.FResultCount-1 %>
<% if obackImage.FItemList(i).FImageUrl<>"" then %>
<tr>
	<td><img src="<%= obackImage.FItemList(i).FImageUrl %>" width=50 height=50 onerror="AddErrLisr('<%= obackImage.FItemList(i).FImageUrl %>')" ></td>
</tr>
<% end if %>
<% next %>
</table>
</body>
<%
set obackImage = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->