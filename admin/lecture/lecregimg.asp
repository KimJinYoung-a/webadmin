<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecturecls.asp"-->
<%
dim itemid,ix
itemid=request("itemid")
dim sqlStr



dim oitem
set oitem = new iteminfo
oitem.Citemid=itemid
oitem.GetItemData

Class itemInfo

	private FImageIndex
	private FImageMain
	private FImageAdd, FImageBasic
	private FImageAddStr, FImageAddContent
	
	private FImageList, FImageAddContentStr
	
	public Citemid
	
	public Sub GetItemData()
	
	sqlStr = "select top 1 a.imgmain, a.imglist, a.imgbasic, a.imgadd, a.imgtitle" + vbcrlf
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_image a, [db_item].[dbo].tbl_item b, [db_user].[dbo].tbl_user_c c"  + vbcrlf
	sqlStr = sqlStr + " where a.itemid = '" + CStr(Citemid) + "'" + vbcrlf
	sqlStr = sqlStr + " and a.itemid = b.itemid"
	sqlStr = sqlStr + " and b.makerid = c.userid"
		
	rsget.Open sqlStr,dbget,1
	
		if  not rsget.EOF  then
	        rsget.Movefirst

			FImageMain = "http://webimage.10x10.co.kr/image/main/" + GetImageFolerName + "/" + rsget("imgmain")
			FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageFolerName + "/" + rsget("imglist")
			FImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageFolerName + "/" + rsget("imgbasic")

	        
			FImageAddStr = rsget("imgadd")
			
			
			if not ISNull(FImageAddContentStr) then
				FImageAddContentStr = replace(FImageAddContentStr,vbcrlf," ")
				FImageAddContentStr = replace(FImageAddContentStr,"'","")
			end if


	        if FImageIndex>0 then
	        	FImageMain = ""
	        	FImageAdd = "http://webimage.10x10.co.kr/image/add/" + trim((split(FImageAddStr,","))(FImageIndex-1))

              	if not ISNull(FImageAddContentStr) then
					FImageAddContentStr = replace(FImageAddContentStr,vbcrlf," ")
              		FImageAddContentStr = replace(FImageAddContentStr,vbcr," ")
              		FImageAddContentStr = replace(FImageAddContentStr,vblf," ")
               		'FImageAddContent = trim((split(FImageAddContentStr,"|"))(FImageIndex-1))
              	end if
	        end if
		end if

	rsget.close
	End Sub


	Public Property Get ImageAddArray(byval v)
		dim tmp
		tmp = Trim(split(FImageAddStr,",")(v))
		if tmp<>"" then
			ImageAddArray = "http://webimage.10x10.co.kr/image/add" + Cstr(v + 1) + "/" + GetImageFolerName + "/" + tmp
		end if
	End Property

	Public Function AddImageCount()
		dim buf,i
		buf = split(FimageAddStr,",")
		i=0
		do until i > ubound(buf)
			if (len(trim(buf(i))) = 0) then
                 exit do
            end if
		 	i = i+1
		loop
		AddImageCount = i
	end Function
	
	public function GetImageFolerName()
		GetImageFolerName = "0" + CStr(Clng(ItemID\10000))
	end function
	
	Public Property Get ImageAddContentArray(byval v)
		if IsNull(FImageAddContentStr) then
			ImageAddContentArray = ""
			Exit Property
		end if

		if UBound(Split(FImageAddContentStr,"|")) < v then
			ImageAddContentArray = ""
		else
			ImageAddContentArray = Split(FImageAddContentStr,"|")(v)
		end if
	End Property
	
end Class


%>

<table width="600" border="0" cellpadding="0" cellspacing=0" class="a" topmargin="0">
	<form name="imgfrm" method="post" action="/admin/lecture/lecregimg.asp">
	<tr>
		<td align="center"><input type="text" name="itemid" value="<%= itemid %>">
		<br>
		(이미지를 불러올 강좌의 상품 코드 입력)
		</td>
	</tr>
	<tr>
		<td align="center"><input type="submit" value="적용"></td>
	</tr>
</table>
<% if itemid<>"" then %>
<table>
	<tr>
		<td align="center"><img src="<%'= oitem.ImageMain %>" width="300" height="250"></td>
	</tr>
	<% for ix=0 to oitem.AddImageCount - 1 %>
	<tr>
		<td align="center"><img src="<%= oitem.ImageAddArray(ix) %>" border="0" width="600"><br>
			<%= oitem.ImageAddContentArray(ix) %><br></td>
		</tr>
	<% next %>
</table>
<% end if %>
</div>
<%
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->