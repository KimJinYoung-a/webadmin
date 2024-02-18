<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/love_house_main_cls.asp" -->
<%

'// 사용안함, 2015-07-07, skyer9
dbget.close()
response.end

dim objMon
Set objMon = Server.CreateObject("DEXT.FileUploadMonitor")
objMon.UseMonitor(True)
Set objMon = Nothing

dim page,ix,i
page=request("page")
if page="" then page="1"

dim loveitem
set loveitem = new LoveHouse
loveitem.FCurrPage=page
loveitem.FPageSize=10
loveitem.FScrollcount=5
loveitem.GetLoveMainList
 %>

 <script>
function popwin(id){
	window.open("http://uploadmain.10x10.co.kr/chtml/make_love_house_winner.asp?idx="+ id,"wind","width=10,height=10,status=no,scrollbars=no,resizable=yes")
}
function InputImg(){
 	var frm;

	frm=document.getElementById("newtbl").style;

	if (frm.display=="none"){
		frm.display="block";
	} else {
		frm.display="none";
	}
}

function subInput(){

	var frm;
	frm=document.inputfrm;

	strAppVersion = navigator.appVersion;
	if (frm.winimage.value != "") {
		if (strAppVersion.indexOf('MSIE')!=-1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4) {
			winstyle = "dialogWidth=385px; dialogHeight:150px; center:yes";
			window.showModelessDialog("show_progress.asp?nav=ie", null, winstyle);
		} else {
			winpos = "left=" + ((window.screen.width-380)/2) + ",top=((window.screen.height-110)/2)";
			winstyle = "width=380,height=110,status=no,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=no,copyhistory=no," + winpos;
			window.open("show_progress.asp",null,winstyle);
		}
	}

	if (frm.inputimg.value.length<1){
		alert('이미지를 입력하세요.');
		frm.inputimg.focus();
		return false;
	}

	frm.submit();
}

function showImg(v){

	v.imgpan.src=v.inputimg.value;
}
function showImg2(v){
	v.winimgpan.src=v.winimage.value;
}
function PopEdit(id){
	window.open("love_house_edit_pop.asp?idx="+ id,"editwin","width=690, height=500, scrollbars=yes,status=no,resizable=yes")
}

 </script>
<table width="700" border="0" cellpadding="0" cellspacing="0" class="a" >
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
		<tr height="40" valign="bottom">
		    <td align="left">
				<input type="button" class="button" value="새로등록" onClick="javascript:InputImg();">
			</td>
			<td align="right">
				<input type="button" class="button" value="새로고침" onClick="javascript:document.location.reload()">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<div style="display:none" id="newtbl" align="left" >
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  bgcolor="<%= adminColor("tablebg") %>">
			<form name="inputfrm" method="post" action="<%=uploadUrl%>/linkweb/dolovehousewinner.asp" enctype="multipart/form-data">
			<input type="hidden" name="idx" value="">
			<input type="hidden" name="mode" value="add">
			<tr class="a">
				<td align="center" bgcolor="<%= adminColor("gray") %>">당첨자 아이디</td>
				<td bgcolor="#FFFFFF"><input type="text" name="userid" size="20"></td>
			</tr>
			<tr class="a">
				<td align="center" bgcolor="<%= adminColor("gray") %>">당첨월</td>
				<td bgcolor="#FFFFFF"><input type="text" name="windate" size="20">(ex:2006-03-01)</td>
			</tr>
			<tr class="a">
				<td align="center" bgcolor="<%= adminColor("gray") %>">Title 이미지</td>
				<td bgcolor="#FFFFFF">
					<input type="file" name="inputimg" size="35" Onchange="showImg(this.form)">
					<br>
					<img name="imgpan" src="">
				</td>
			</tr>
			<tr class="a">
				<td align="center" bgcolor="<%= adminColor("gray") %>">메인 이미지</td>
				<td bgcolor="#FFFFFF">
					<input type="file" name="winimage" size="35" Onchange="showImg2(this.form)">
					<br>
					<img name="winimgpan" src=""></td>
			</tr>
			<tr class="a">
				<td align="center" bgcolor="<%= adminColor("gray") %>">이미지 맵</td>
				<td bgcolor="#FFFFFF"> 맵 이름 꼭 <font color="blue">"lovemap"</font>으로 통일해주세요<br>
					&lt;map name="lovemap"&gt;<br>
					&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoProduct('<font color="blue">상품번호</font>');" onfocus="this.blur();"&gt;<br>
					&lt;/map&gt; <br>
					<textarea name="lovemap" cols="70" rows="20"></textarea></td>
			</tr>
				<tr class="a">
				<td align="center" bgcolor="<%= adminColor("gray") %>">WishList 상품번호</td>
				<td bgcolor="#FFFFFF">
					상품1: <input type="text" name="itemid1" size="10" maxlength="10"><br>
					상품2: <input type="text" name="itemid2" size="10" maxlength="10"><br>
					상품3: <input type="text" name="itemid3" size="10" maxlength="10"><br>
					상품4: <input type="text" name="itemid4" size="10" maxlength="10"><br>
				</td>
			</tr>
			<tr class="a" bgcolor="#FFFFFF">
				<td colspan="2" align="center"><input type="button" Onclick="javascript:subInput();" value="저장" class="button"></td>
			</tr>
			</form>
		</table>
		</div>
	</td>
</tr>
<tr>
	<td style="padding-top:10px;" align="left">
	<table width="100%"  cellpadding="3" cellspacing="1" class="a"  bgcolor="<%= adminColor("tablebg") %>">
	<tr class="a" bgcolor="<%= adminColor("gray") %>">
		<td align="center">번호</td>
		<td align="center">이미지</td>
		<td>&nbsp;</td>
	</tr>
	<% if loveitem.FResultCount<1 then %>
	<tr class="a" bgcolor="#FFFFFF">
		<td colspan="3" align="center">[검색결과가 없습니다.]</td>
	</tr>
	<% else %>
	<% for ix =0 to loveitem.FResultCount -1 %>
	<form name="edit_frm_<%= ix %>" method="post" action="<%=uploadUrl%>/linkweb/dolovehousewinner.asp" enctype="multipart/form-data">
	<input type="hidden" name="idx" value="<% = loveitem.Fidx(ix) %>">
	<input type="hidden" name="mode" value="edit">

	<% if loveitem.FViewYn(ix)="Y" then %>
		<tr class="a" bgcolor="#99FFFF">
	<% else %>
		<tr class="a" bgcolor="#FFFFFF">
	<% end if %>
		<td align="center"><% = loveitem.Fidx(ix) %></td>
		<td align="center">
			<a href="javascript:PopEdit('<% = loveitem.Fidx(ix) %>');"><img src="<% =loveitem.FImage(ix) %>" border="0" name="editpan "></a>
		</td>
		<td align="center">
			<!--<input type="button" value="적용" Onclick="javascript:popwin('<%= loveitem.Fidx(ix) %>')" class="button">-->
		</td>
	</tr>
	</form>
	<% next %>
	<% end if %>
	<tr class="a" bgcolor="#FFFFFF">
		<td colspan="5" align="center" height="30">
				<% if loveitem.HasPreScroll then %>
					<a href="?page=<%= loveitem.StartScrollPage-1 %>">[pre]</a>
				<% else %>
					[pre]
				<% end if %>

				<% for i=0 + loveitem.StartScrollPage to loveitem.FScrollCount + loveitem.StartScrollPage - 1 %>
					<% if i>loveitem.FTotalpage then Exit for %>
					<% if CStr(page)=CStr(i) then %>
					<font color="red">[<%= i %>]</font>
					<% else %>
					<a href="?page=<%= i %>">[<%= i %>]</a>
					<% end if %>
				<% next %>

				<% if loveitem.HasNextScroll then %>
					<a href="?page=<%= i %>">[next]</a>
				<% else %>
					[next]
				<% end if %>
		</td>
	</tr>
	</table>
	</td>
</tr>
</table>
<%
set loveitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
