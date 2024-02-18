<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/bloodcls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim puzzleid,itemid
puzzleid = request("puzzleid")
itemid = request("itemid")

dim mode,page
mode = request("mode")
page = request("page")

dim eid
eid = request("eid")

if page="" then
	page =1 
end if

dim Oblood
set Oblood = New CBloodMaster
Oblood.FCurrPage = page
Oblood.FPageSize = 10
Oblood.GetAllBlood 

dim EditBlood
set EditBlood = New CBloodMaster
if eid="" then eid =0
EditBlood.GetOneBlood eid

dim i
%>
<script language="javascript">
function Addpuzzle(frm){
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		
		if ((e.type=="text")) {
			if ((e.name=="itemid") || (e.name=="puzzlename") || (e.name=="limitno") || (e.name=="startdate") 
				|| (e.name=="finishdate")){
				if (e.value.length<1){
					alert('필수 입력 사항입니다.');
					e.focus();
					return;
				}
			}

			if ((e.name=="itemid") || (e.name=="limitno")){
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

function ShowImage(src, imgname, hid){
	var imgcomp;
	imgcomp = eval("document." + imgname);
	imgcomp.src = src;
	hid.checked=false;

}

function DelImage(src, imgname, hid){
	var imgcomp;
	imgcomp = eval("document." + imgname);
	if (hid.checked){
		imgcomp.src = '/images/space.gif';
	}
}
</script>

<table width="700" border="0" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td colspan="15" align="right" height="30"><a href="?mode=add"><font color="red">new</font></a></td>
	</tr>
	<tr>
		<td width="100" align="center">
			IDX
		</td>
		<td  align="center">
			Title
		</td>
		<td width="100"  align="center">
			남여구분
		</td>
		<td width="100"  align="center">
			사용여부
		</td>
		<td width="100"  align="center">
			등록일
		</td>
		<td width="100"  align="center">
			상품등록버튼
		</td>
	</tr>
	<tr>
		<td colspan="15" height="1" bgcolor="#AAAAAA"></td>
	</tr>
	<% for i=0 to Oblood.FResultcount -1 %>
	<tr>
		<td align="center" height="25">
			<%= Oblood.FbloodList(i).Fidx %>
		</td>
		<td>
			<a href="?mode=edit&eid=<%= Oblood.FbloodList(i).Fidx %>"><%= Oblood.FbloodList(i).Ftitle %></a>
		</td>
		<td align="center">
			<%= Oblood.FbloodList(i).FSexName %>
		</td>
		<td align="center">
			<% if Oblood.FbloodList(i).Fisusing="Y" then %>
			Y
			<% else %>
			<font color="red">N</font>
			<% end if %>
		</td>
		<td align="center">
			<%= FormatDateTime(Oblood.FbloodList(i).FRegDate,0) %>
		</td>
		<td align="center">
			<input type="button" value="상품선택" onclick="location.href='valentine_item_list.asp?masterid=<% = Oblood.FbloodList(i).Fidx %>'">
		</td>
	</tr>
	<% next %>
	<tr>
		<td colspan="14" height="1" bgcolor="#AAAAAA"></td>
	</tr>
	<tr>
		<td colspan="14" align="center">
			<% if Oblood.HasPreScroll then %>
				<a href="?page=<%= Oblood.StarScrollPage-1 %>">[pre]</a>
			<% else %> 
				[pre]
			<% end if %>
			
			<% for i=0 + Oblood.StarScrollPage to Oblood.FScrollCount + Oblood.StarScrollPage - 1 %>
				<% if i>Oblood.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>">[<%= i %>]</a>
				<% end if %>
			<% next %>
			
			<% if Oblood.HasNextScroll then %>
				<a href="?page=<%= i %>">[next]</a>
			<% else %> 
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<br>
<% if (mode="add") or (mode="edit") then %>
<form name="addfrm" method="post" action="http://partner.10x10.co.kr/admin/etc/blood_upload_ok.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="eid" value="<%= eid %>">
<table width="500" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="2" >
		<% if (mode="add") then %>
		- 게시판 경매 추가
		<% else %>
		- 게시판 경매 수정
		<% end if %>
	</td>
</tr>
<% if mode="add" then %>
<% else %>
<tr>
	<td>IDX</td>
	<td><%= EditBlood.FbloodList(0).Fidx %></td>
</tr>
<% end if %>
<tr>
	<td>경매이름</td>
	<td><input type="text" name="title" value="<%= EditBlood.FbloodList(0).Ftitle %>"></td>
</tr>
<tr>
	<td>남녀구분</td>
	<% if mode="add" then %>
	<td><input type="radio" name="sex" value="1"> 남자 <input type="radio" name="sex" value="2"> 여자 <input type="radio" name="sex" value="3"> 혼합</td>
	<% else %>
	<td><input type="radio" name="sex" value="1" <% if EditBlood.FbloodList(0).Fsex = 1 then response.write "checked" %>> 남자 <input type="radio" name="sex" value="2" <% if EditBlood.FbloodList(0).Fsex = 2 then response.write "checked" %>> 여자 <input type="radio" name="sex" value="3" <% if EditBlood.FbloodList(0).Fsex = 3 then response.write "checked" %>> 혼합</td>
	<% end if %>
</tr>
<tr>
	<td>사용여부</td>
	<td>
		<input type="radio" name="isusing" value="Y" <% if EditBlood.FbloodList(0).FIsUsing="Y" then response.write "checked" %> >Y
		<input type="radio" name="isusing" value="N" <% if EditBlood.FbloodList(0).FIsUsing<>"Y" then response.write "checked" %> >N
	</td>
</tr>
<tr>
	<td width="120">A형List이미지</td>
	<td>
		<% if mode="add" then %>
		<img name="iimageAmain" src="" width="300" height="200"><br>
		<% else %>
		<img name="iimageAmain" src="<%= EditBlood.FbloodList(0).FAmain %>" width="300" height="200" ><br>
		<% end if %>
		<div align="right">
		<input type="file" name="imageAmain" size="30" onchange="ShowImage(this.form.imageAmain.value,'iimageAmain',this.form.A_main)"><br>
		<input type="checkbox" name="A_main" onclick="DelImage(this.form.imageAmain,'iimageAmain',this.form.A_main)">삭제
		</div>
	</td>
</tr>
<tr>
	<td width="120">A형Main이미지</td>
	<td>
		<% if mode="add" then %>
		<img name="iimageAlist" src="" width="300" height="200"><br>
		<% else %>
		<img name="iimageAlist" src="<%= EditBlood.FbloodList(0).FAlist %>" width="300" height="200" ><br>
		<% end if %>
		<div align="right">
		<input type="file" name="imageAlist" size="30" onchange="ShowImage(this.form.imageAlist.value,'iimageAlist',this.form.A_list)"><br>
		<input type="checkbox" name="A_list" onclick="DelImage(this.form.imageAlist,'iimageAlist',this.form.A_list)">삭제
		</div>
	</td>
</tr>
<tr>
	<td width="120">B형List이미지</td>
	<td>
		<% if mode="add" then %>
		<img name="iimageBmain" src="" width="300" height="200"><br>
		<% else %>
		<img name="iimageBmain" src="<%= EditBlood.FbloodList(0).FBmain %>" width="300" height="200" ><br>
		<% end if %>
		<div align="right">
		<input type="file" name="imageBmain" size="30" onchange="ShowImage(this.form.imageBmain.value,'iimageBmain',this.form.B_main)"><br>
		<input type="checkbox" name="B_main" onclick="DelImage(this.form.imageBmain,'iimageBmain',this.form.B_main)">삭제
		</div>
	</td>
</tr>
<tr>
	<td width="120">B형Main이미지</td>
	<td>
		<% if mode="add" then %>
		<img name="iimageBlist" src="" width="300" height="200"><br>
		<% else %>
		<img name="iimageBlist" src="<%= EditBlood.FbloodList(0).FBlist %>" width="300" height="200" ><br>
		<% end if %>
		<div align="right">
		<input type="file" name="imageBlist" size="30" onchange="ShowImage(this.form.imageBlist.value,'iimageBlist',this.form.B_list)"><br>
		<input type="checkbox" name="B_list" onclick="DelImage(this.form.imageBlist,'iimageBlist',this.form.B_list)">삭제
		</div>
	</td>
</tr>
<tr>
	<td width="120">O형List이미지</td>
	<td>
		<% if mode="add" then %>
		<img name="iimageOmain" src="" width="300" height="200"><br>
		<% else %>
		<img name="iimageOmain" src="<%= EditBlood.FbloodList(0).FOmain %>" width="300" height="200" ><br>
		<% end if %>
		<div align="right">
		<input type="file" name="imageOmain" size="30" onchange="ShowImage(this.form.imageOmain.value,'iimageOmain',this.form.O_main)"><br>
		<input type="checkbox" name="O_main" onclick="DelImage(this.form.imageOmain,'iimageOmain',this.form.O_main)">삭제
		</div>
	</td>
</tr>
<tr>
	<td width="120">O형Main이미지</td>
	<td>
		<% if mode="add" then %>
		<img name="iimageOlist" src="" width="300" height="200"><br>
		<% else %>
		<img name="iimageOlist" src="<%= EditBlood.FbloodList(0).FOlist %>" width="300" height="200" ><br>
		<% end if %>
		<div align="right">
		<input type="file" name="imageOlist" size="30" onchange="ShowImage(this.form.imageOlist.value,'iimageOlist',this.form.O_list)"><br>
		<input type="checkbox" name="O_list" onclick="DelImage(this.form.imageOlist,'iimageOlist',this.form.O_list)">삭제
		</div>
	</td>
</tr>
<tr>
	<td width="120">AB형List이미지</td>
	<td>
		<% if mode="add" then %>
		<img name="iimageABmain" src="" width="300" height="200"><br>
		<% else %>
		<img name="iimageABmain" src="<%= EditBlood.FbloodList(0).FABmain %>" width="300" height="200" ><br>
		<% end if %>
		<div align="right">
		<input type="file" name="imageABmain" size="30" onchange="ShowImage(this.form.imageABmain.value,'iimageABmain',this.form.AB_main)"><br>
		<input type="checkbox" name="AB_main" onclick="DelImage(this.form.imageABmain,'iimageABmain',this.form.AB_main)">삭제
		</div>
	</td>
</tr>
<tr>
	<td width="120">AB형Main이미지</td>
	<td>
		<% if mode="add" then %>
		<img name="iimageABlist" src="" width="300" height="200"><br>
		<% else %>
		<img name="iimageABlist" src="<%= EditBlood.FbloodList(0).FABlist %>" width="300" height="200" ><br>
		<% end if %>
		<div align="right">
		<input type="file" name="imageABlist" size="30" onchange="ShowImage(this.form.imageABlist.value,'iimageABlist',this.form.AB_list)"><br>
		<input type="checkbox" name="AB_list" onclick="DelImage(this.form.imageABlist,'iimageABlist',this.form.AB_list)">삭제
		</div>
	</td>
</tr>
<tr>
	<td colspan="2" align="center">
		<% if mode="add" then %>
		<input type="button" value="추가" onClick="Addpuzzle(addfrm)">
		<% elseif  mode="edit" then %>
		<input type="button" value="수정" onClick="Addpuzzle(addfrm)">
		<% end if %>
	</td>
</tr>
</table>
</form>
<% end if %>
<%
set Oblood = Nothing
set EditBlood = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->