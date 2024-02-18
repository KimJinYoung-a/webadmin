<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<link rel="stylesheet" href="/bct.css" type="text/css">
<%
dim FItemid,FUserid,FContents,FRegDate,FTPoint,FPoint1,FPoint2,FPoint3,FPoint4,FListimage1,FListimage2,FSmallImage,FTopYn
dim strSQL,i,j,FResultCount,pointcnt

dim FRectItemid
FRectItemid= request("itemid")

				public function GetImageFolerName(byval i)
					GetImageFolerName = "0" + CStr(Clng(FItemid(i)\10000))
				end function


				'// 이벤트 진행 중인상품명 찾기
				public function getItemSelect(byval itemid)

				  	dim selSQL ,selStr

				  	selSQL = 	" select itemid ,itemname from db_contents.dbo.tbl_diary_event_evaluate " &_
				  						" group by itemid,itemname order by itemname"

						  	rsget.open selSQL,dbget,1

						  	if not rsget.eof then

						  		selStr = "<select name='itemid' onchange=TnSearch(this.value)>"
						  		selStr = selStr + "<option value=''>------------- 상품명으로 찾기 -------------</option>"

						  		do until rsget.eof
						  				if CStr(itemid)=CStr(rsget("itemid")) then
						  					selStr = selStr + "<option value='" & rsget("itemid") & "' selected >" & db2html(left(rsget("itemname"),40)) & "</option>"
						  				else
						  					selStr = selStr + "<option value='" & rsget("itemid") & "'>" & db2html(left(rsget("itemname"),40)) & "</option>"
						  				end if
						  		rsget.movenext
						  		selStr = selStr + "</selec>"

						  		loop

						  	end if

				  	rsget.close

				  	getItemSelect=selStr

				end function

				strSQL =" SELECT  e.itemid,e.userid,e.contents,e.point,e.point1,e.point2,e.point3,e.point4,e.regdate,i.smallimage,e.linkimg1,linkimg2,topYn " &_
								" FROM db_contents.dbo.tbl_diary_event_evaluate e join db_item.[10x10].tbl_item i on i.itemid=e.itemid"
				if FRectItemid<>"" then
					strSQL = strSQL + " where e.itemid=" & FRectItemid
				end if
				strSQL = strSQL + " Order By e.itemid , e.idx  "


				rsget.open strSQL,dbget,1

				if not rsget.eof then
					i=0
					FResultCount = rsget.recordcount
					redim FItemId(FResultCount)
					redim FUserid(FResultCount)
					redim FContents(FResultCount)
					redim FRegDate(FResultCount)
					redim FTPoint(FResultCount)
					redim FPoint1(FResultCount)
					redim FPoint2(FResultCount)
					redim FPoint3(FResultCount)
					redim FPoint4(FResultCount)
					redim FSmallImage(FResultCount)
					redim FListimage1(FResultCount)
					redim FListimage2(FResultCount)
					redim FTopYn(FResultCount)

					do until rsget.eof


							FItemId(i) 		= rsget("ItemId")
							FUserid(i) 		= rsget("userid")
							FContents(i)	= db2html(rsget("contents"))
							FRegDate(i)		= db2html(rsget("regdate"))
							FTPoint(i)		= rsget("point")
							FPoint1(i)		= rsget("point1")
							FPoint2(i)		= rsget("point2")
							FPoint3(i)		= rsget("point3")
							FPoint4(i)		= rsget("point4")

							if rsget("smallimage") <>"" then
							FSmallImage(i)		= "http://webimage.10x10.co.kr/image/small/" & GetImageFolerName(i) & "/" & db2html(rsget("smallimage"))
							end if

							if rsget("linkimg1") <>"" then
							FListimage1(i)	= "http://imgstatic.10x10.co.kr/goodsimage/" & GetImageFolerName(i) & "/" & db2html(rsget("linkimg1"))
							end if

							if rsget("linkimg2") <>"" then
							FListimage2(i)	= "http://imgstatic.10x10.co.kr/goodsimage/" & GetImageFolerName(i) & "/" & db2html(rsget("linkimg2"))
							end if
							FTopYn(i)			= rsget("topYn")
						rsget.movenext
						i=i+1
					loop
				end if

				rsget.close


%>
<script language="javascript" type="text/javascript">
function TnSearch(){
	document.searchfrm.submit();
}

function selBest(itemid,userid){
	document.selfrm.mode.value='best';
	document.selfrm.itemid.value=itemid;
	document.selfrm.userid.value=userid;
	document.selfrm.submit();
}

function del(itemid,userid){
	document.selfrm.mode.value='del';
	document.selfrm.itemid.value=itemid;
	document.selfrm.userid.value=userid;
	document.selfrm.submit();
}
</script>
<iframe src="" name="doFrame" frameborder="0" width="0" height="0"></iframe>

<form name="selfrm" method="post" target="doFrame" action="do_diary_event_eval.asp">
<input type="hidden" name="mode" value="" />
<input type="hidden" name="userid" value="" />
<input type="hidden" name="itemid" value="" />
</form>

<table width="680" border="0" cellspacing="0" cellpadding="0" class="a">
<form name="searchfrm" method="post" action="">
<tr>
	<td><%= getItemSelect(FRectItemid) %></td>
</tr>

</form>

<% if FResultCount < 0 then %>
<% else %>
<% for i = 0 to FResultCount -1 %>

<% if i=0 then %>
				<tr bgcolor="#CDCDCD"><td><img src="<%= FSmallImage(i) %>" width="50" height="50" border="0" />(<%= FItemId(i) %>)</td></tr>
<% else %>
		<% if Fitemid(i)<>Fitemid(i-1) then %>
				<tr bgcolor="#CDCDCD"><td><img src="<%= FSmallImage(i) %>" width="50" height="50" border="0" /><%= FItemid(i) %></td></tr>
		<% end if %>
<% end if %>
<tr>
	<td colspan="2" align="center" valign="top"  style="border-bottom:1px solid #E6E6E6; padding:3 10 3 10">
		<table width="680" border="0" cellspacing="0" cellpadding="0" class="a">


			<tr <% if FtopYn(i)="Y" then  response.write "bgcolor='#EDEDED'" %>>
				<td rowspan="2" align="center" width="120" valign=top>
					<table width="120" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td height="25" width="40"  class="coupon_kor_gray11">총평</td>
							<td height="25"><% for pointcnt=0 to  FTpoint(i)-1 %><img src="http://www.10x10.co.kr/images/category/px_02.gif" width="10" height="9"><% next %></td>
						</tr>
						<tr>
						  <td>기능</td>
						  <td><% for j=0 to  FPoint1(i)-1 %><img src="http://www.10x10.co.kr/images/category/px_01.gif" width="10" height="9"><% next %></td>
						</tr>
						<tr>
						  <td>디자인</td>
							<td><% for j=0 to  FPoint2(i)-1 %><img src="http://www.10x10.co.kr/images/category/px_01.gif" width="10" height="9"><% next %></td>
						</tr>
						<tr>
						  <td>가격</td>
							<td><% for j=0 to  FPoint3(i)-1 %><img src="http://www.10x10.co.kr/images/category/px_01.gif" width="10" height="9"><% next %></td>
						</tr>
						<tr>
						  <td>만족도</td>
							<td><% for j=0 to  FPoint4(i)-1 %><img src="http://www.10x10.co.kr/images/category/px_01.gif" width="10" height="9"><% next %></td>
						</tr>
					</table>
				</td>
				<td>
				  <span class="coupon_kor_gray11">[<% = FUserID(i) %>]</span>
				  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span onclick="selBest('<%= FItemId(i) %>','<%= FUserid(i) %>');" style="cursor:pointer">[베스트선택]</span>
				  &nbsp;&nbsp;<span onclick="del('<%= FItemId(i) %>','<%= FUserid(i) %>');" style="cursor:pointer">[삭제]</span>
				</td>
			</tr>
			<tr <% if FtopYn(i)="Y" then  response.write "bgcolor='#EDEDED'" %>>
				<td valign="top" class="coupon_kor_gray11"><% = FContents(i) %><br>
				<% if FListimage1(i)<>"" then %>
				<a href="javascript:NewWindow('<%= FListimage1(i) %>')"><img src="<% = FListimage1(i) %>" id="file1<% = i %>"></a><br>
				<% end if %>
				<% if FListimage2(i)<>"" then %>
				<a href="javascript:NewWindow('<%= FListimage2(i) %>')"><img src="<% = FListimage2(i) %>" id="file2<% = i %>"></a>
				<% end if %>
				</td>
			</tr>
		</table>
	</td>
</tr>
<% next %>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

