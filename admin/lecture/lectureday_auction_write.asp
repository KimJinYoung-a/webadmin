<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lectureday_auctioncls.asp"-->
<%
dim mode,idx

mode = request("mode")
idx = request("idx")

dim editAuction
set editAuction = New CBoardAuction
if idx="" then idx =0
editAuction.GetOneAuction idx

dim i
%>
<script language="javascript">
function AddAuction(frm){
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="text")) {
			if ((e.name=="itemid") || (e.name=="auctionname") || (e.name=="limitno") || (e.name=="startdate")
				|| (e.name=="finishdate") || (e.name=="supplyer")|| (e.name=="pricestart")|| (e.name=="priceend")
				|| (e.name=="pricefix") ){
				if (e.value.length<1){
					alert('필수 입력 사항입니다.');
					e.focus();
					return;
				}
			}

			if ((e.name=="itemid") || (e.name=="limitno") || (e.name=="pricestart") || (e.name=="priceend") || (e.name=="pricefix")){
				if (!IsDigit(e.value)){
					alert('숫자만 가능합니다.');
					e.focus();
					return;
				}
			}
		}
	}
	<% if mode="add" then %>
	var ret = confirm('추가 하시겠습니까?');
	<% else %>
	var ret = confirm('수정 하시겠습니까?');
	<% end if %>
	if (ret) { frm.submit();}
}
</script>
<form name="addfrm" method="post" action="http://partner.10x10.co.kr/admin/lectureday/donewauction.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<table width="600" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td width="100">번호</td>
	<td><%= editAuction.FAuctionList(0).Fidx %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>경매이름</td>
	<td><input type="text" name="auctionname" value="<%= editAuction.FAuctionList(0).Fauctionname %>" size="70" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>시작일<br>(2005-02-05)</td>
	<% if mode="add" then %>
	<td><input type="text" name="startdate" value="2005-02-05" size="70" class="input_b"></td>
	<% else %>
	<td><input type="text" name="startdate" value="<%= editAuction.FAuctionList(0).Fstartdate %>" size="70" class="input_b"></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>종료일<br>(2002-10-05 23:00:00)</td>
	<% if mode="add" then %>
	<td><input type="text" name="finishdate" value="2005-02-05 00:00:00" size="70" class="input_b"></td>
	<% else %>
	<td><input type="text" name="finishdate" value="<%= editAuction.FAuctionList(0).Ffinishdate %>" size="70" class="input_b"></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>판매갯수</td>
	<td><input type="text" name="itemea" value="<%= editAuction.FAuctionList(0).Fitemea %>" size="30" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>메인이미지</td>
	<td><input type="file" name="mainimg" size="50" class="input_b">(가로:550이하)
	</td>
	<% else %>
	<td>메인이미지</td>
	<td><input type="file" name="mainimg" size="50" class="input_b">(가로:550이하)<br>
		<input type="checkbox" name="dl_mainimg">삭제 (<%= editAuction.FAuctionList(0).Fmainimg %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>추가이미지1</td>
	<td><input type="file" name="img1" size="50" class="input_b">(가로:550이하)
	</td>
	<% else %>
	<td>추가이미지1</td>
	<td><input type="file" name="img1" size="50" class="input_b">(가로:550이하)<br>
		<input type="checkbox" name="dl_img1">삭제 (<%= editAuction.FAuctionList(0).Fimg1 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>추가이미지2</td>
	<td><input type="file" name="img2" size="50" class="input_b">(가로:550이하)
	</td>
	<% else %>
	<td>추가이미지2</td>
	<td><input type="file" name="img2" size="50" class="input_b">(가로:550이하)<br>
		<input type="checkbox" name="dl_img2">삭제 (<%= editAuction.FAuctionList(0).Fimg2 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>추가이미지3</td>
	<td><input type="file" name="img3" size="50" class="input_b">(가로:550이하)
	</td>
	<% else %>
	<td>추가이미지3</td>
	<td><input type="file" name="img3" size="50" class="input_b">(가로:550이하)<br>
		<input type="checkbox" name="dl_img3">삭제 (<%= editAuction.FAuctionList(0).Fimg3 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>추가이미지4</td>
	<td><input type="file" name="img4" size="50" class="input_b">(가로:550이하)
	</td>
	<% else %>
	<td>추가이미지4</td>
	<td><input type="file" name="img4" size="50" class="input_b">(가로:550이하)<br>
		<input type="checkbox" name="dl_img4">삭제 (<%= editAuction.FAuctionList(0).Fimg4 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>추가이미지5</td>
	<td><input type="file" name="img5" size="50" class="input_b">(가로:550이하)
	</td>
	<% else %>
	<td>추가이미지5</td>
	<td><input type="file" name="img5" size="50" class="input_b">(가로:550이하)<br>
		<input type="checkbox" name="dl_img5">삭제 (<%= editAuction.FAuctionList(0).Fimg5 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>리스트이미지50</td>
	<td><input type="file" name="icon1" size="50" class="input_b">(50*50)
	</td>
	<% else %>
	<td>리스트이미지50</td>
	<td><input type="file" name="icon1" size="50" class="input_b">(50*50)<br>
		<input type="checkbox" name="dl_icon1">삭제 (<%= editAuction.FAuctionList(0).Ficon1 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>리스트이미지100</td>
	<td><input type="file" name="icon2" size="50" class="input_b">(80*80)
	</td>
	<% else %>
	<td>리스트이미지100</td>
	<td><input type="file" name="icon2" size="50" class="input_b">(80*80)<br>
		<input type="checkbox" name="dl_icon2">삭제 (<%= editAuction.FAuctionList(0).Ficon2 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>상품설명</td>
	<% if mode="add" then %>
	<td><textarea name="itemcontents" rows="10" cols="70" class="input_b"></textarea></td>
	<% else %>
	<td><textarea name="itemcontents" rows="10" cols="70" class="input_b"><%= editAuction.FAuctionList(0).Fitemcontents %></textarea></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>주의사항</td>
	<% if mode="add" then %>
	<td><textarea name="etc" rows="10" cols="70" class="input_b"></textarea></td>
	<% else %>
	<td><textarea name="etc" rows="10" cols="70" class="input_b"><%= editAuction.FAuctionList(0).Fetc %></textarea></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>경매안내</td>
	<% if mode="add" then %>
	<td><textarea name="info" rows="10" cols="70" class="input_b"></textarea></td>
	<% else %>
	<td><textarea name="info" rows="10" cols="70" class="input_b"><%= editAuction.FAuctionList(0).Finfo %></textarea></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>시작가</td>
	<% if mode="add" then %>
	<td><input type="text" name="startprice" value="0" size="30" class="input_b"></td>
	<% else %>
	<td><input type="text" name="startprice" value="<%= editAuction.FAuctionList(0).Fstartprice %>" size="30" class="input_b"></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>당첨자</td>
	<td><input type="text" name="nakchaluser" value="<%= editAuction.FAuctionList(0).Fnakchaluser %>" size="30" class="input_b">
	<font color=red>(아이디 뒤에 공백 없이 할것!)</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>당첨가</td>
	<% if mode="add" then %>
	<td><input type="text" name="nakchalprice" value="0" size="30" class="input_b"></td>
	<% else %>
	<td><input type="text" name="nakchalprice" value="<%= editAuction.FAuctionList(0).Fnakchalprice %>" size="30" class="input_b"></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>사용여부</td>
	<td>
		<input type="radio" name="isusing" value="Y" <% if editAuction.FAuctionList(0).FIsUsing="Y" then response.write "checked" %> >Y
		<input type="radio" name="isusing" value="N" <% if editAuction.FAuctionList(0).FIsUsing<>"Y" then response.write "checked" %> >N
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">
		<% if mode="add" then %>
		<input type="button" value="추가" onClick="AddAuction(addfrm)">
		<% elseif  mode="edit" then %>
		<input type="button" value="수정" onClick="AddAuction(addfrm)">
		<% end if %>
	</td>
</tr>
</table>
</form>

<%
set editAuction = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->