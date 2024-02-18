<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<!-- #include virtual="/admin/organizer/Lib/include_event_code.asp"-->
<% 
dim Diaryid,mode
Diaryid = request("id")
Mode = request("mode")
dim oDiary

IF Mode = "" THEN Mode = "add"

dim CateCode,ItemID,RegDate,isUsing,ImageName,BasicImgUrl , commentyn ,commentImgName , commentImgUrl
dim event_code, eventgroup_code , event_start , event_end ,weight , organizer_order
dim comStat , color
IF Mode= "edit" THEN
	set oDiary = new organizerCls
	oDiary.frectorganizerid = Diaryid
	oDiary.getDiary()
	
	Diaryid = oDiary.FItem.forganizerid
	CateCode = oDiary.FItem.FCateCode
	ItemID = oDiary.FItem.Fitemid
	RegDate = oDiary.FItem.FRegDate
	isUsing = oDiary.FItem.FisUsing
	ImageName	= oDiary.FItem.FImg
	BasicImgUrl = oDiary.FItem.ImgBasic
	commentyn = oDiary.FItem.fcommentyn
	commentImgName = oDiary.FItem.fcomment_img
	commentImgUrl = oDiary.FItem.Imgcomment
	eventgroup_code = oDiary.FItem.feventgroup_code
	event_code = oDiary.FItem.fevent_code
	event_start = oDiary.FItem.fevent_start
	event_end = oDiary.FItem.fevent_end
	weight = oDiary.Fitem.Fweight
	color = oDiary.Fitem.fcolor
	organizer_order	= oDiary.Fitem.forganizer_order
	'response.write oDiary.Fitem.forganizer_order
	set oDiary = nothing 
	
End IF
IF isUsing="" THEN isUsing="N"
IF commentyn="" THEN commentyn="N"	

IF commentyn="Y" and eventgroup_code<>"" and event_start<=datevalue(now) and datevalue(now) <= event_end Then
	comStat="오픈"
ELSEIF commentyn="Y" and eventgroup_code<>"" and datevalue(now) > event_end Then
	comStat ="종료"
'ELSEIF commentyn="Y" and eventgroup_code<>"" and datevalue(now) < event_start Then
'	comStat ="준비중"
ELSE 
	comStat ="준비중"
End IF

%>


<!-- 리스트 시작 -->
<script language="javascript">
<!--
// 새상품 추가 팝업
function addnewItem(){
		var popwin;
		popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?acURL=<%'=acURL%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		popwin.focus();
}

function showimage(img){
	var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
}
function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

function jsComShow(v){
	
	var tmp = document.getElementById("comconf");
	
	if (v=='Y'){
		tmp.style.display="block";
	}else {
		tmp.style.display="none";
	}
}	
document.domain = "10x10.co.kr";

-->
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmreg" method="post" action="/admin/organizer/Lib/organizerRegProc.asp">
<input type="hidden" name="mode" value="<%= Mode %>">
<input type="hidden" name="did" value="<%= Diaryid %>">
<input type="hidden" name="event_code" value="<%=vEventCode%>">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" height="25">
			<td colspan="2" align="center"><b>오거나이저 등록</b></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> 구분</td>
			<td bgcolor="#FFFFFF" align="left">
				<% SelectList "cate",CateCode %>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> 오거나이저</td>
			<td bgcolor="#FFFFFF" align="left"><%= Diaryid %></td>
			
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> 상품코드</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" name="iid" value="<%=ItemID%>">
				<!--<input type="button" class="button" value="상품찾기" onClick="addnewItem();">-->
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> 이미지 (400x400)</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','basicimgName','basic','2000','400','true');"/>
				(<b><font color="red">400x400</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="basicimgName" value="<%= ImageName %>">
					<div align="right" id="imgdiv"><% IF ImageName<>"" THEN %><img src="<%= BasicImgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= BasicImgUrl %>');"><% End IF %></div>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap>무게</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" name="wt" value="<%= weight %>" disabled>(g)</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> 사용여부</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="radio" name="ius" value="Y" <% IF isUsing="Y" THEN %>checked<% END IF %>>사용
				<input type="radio" name="ius" value="N" <% IF isUsing="N" THEN %>checked<% END IF %> >사용안함
			</td>
			
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap>정렬순위</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" size=10 name="organizer_order" value="<%=organizer_order%>">
			</td>			
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap>color</td>
			<td bgcolor="#FFFFFF" align="left">
				<select name="color">
					<option value="">선택</option>
					<option value="10" <% if color = "10" then response.write " selected" %>>red</option>
					<option value="20" <% if color = "20" then response.write " selected" %>>orange</option>
					<option value="30" <% if color = "30" then response.write " selected" %>>pattern</option>					
					<option value="40" <% if color = "40" then response.write " selected" %>>green</option>
					<option value="50" <% if color = "50" then response.write " selected" %>>navy(짙은녹색)</option>
					<option value="60" <% if color = "60" then response.write " selected" %>>brown</option>

					<option value="70" <% if color = "70" then response.write " selected" %>>black</option>					
					<option value="80" <% if color = "80" then response.write " selected" %>>pink</option>
					<option value="90" <% if color = "90" then response.write " selected" %>>wine</option>
					<option value="100" <% if color = "100" then response.write " selected" %>>blue</option>
				</select>
			</td>
		</tr>		
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap>코맨트사용여부</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="radio" name="commentyn" disabled value="Y" <% IF commentyn="Y" THEN %>checked<% END IF %> onClick="jsComShow(this.value);">사용
				<input type="radio" name="commentyn" disabled value="N" <% IF commentyn="N" THEN %>checked<% END IF %> onClick="jsComShow(this.value);">사용안함
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" align="left">
		<% IF commentyn="Y" Then %>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"  id="comconf" style="display:block;">
		<% ELSE %>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"  id="comconf" style="display:none;">
		<% End IF %>	
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="152">코멘트 진행상태</td>
			<td bgcolor="#FFFFFF" align="left"  ><%=comStat %></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="152">코맨트그룹코드</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" name="eventgroup_code" value = "<%=eventgroup_code%>">
		
			</td>
		</tr>
		
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="152">코맨트 이미지</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv2','commentimgName','comment','2000','800','true');"/>
					<input type="hidden" name="commentimgName" value="<%= commentImgName %>">
					<div align="right" id="imgdiv2"><% IF commentImgName<>"" THEN %><img src="<%= commentImgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= commentImgUrl %>');"><% End IF %></div>
			</td>
		</tr>	
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="152">기간</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" name="event_start" size=10 value="<%= event_start %>">			
				<a href="javascript:calendarOpen3(frmreg.event_start,'시작일',frmreg.event_start.value)">
				<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
				~<input type="text" name="event_end" size=10  value="<%= event_end %>">
				<a href="javascript:calendarOpen3(frmreg.event_end,'마지막일',frmreg.event_end.value)">   
				<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
			</td>
		</tr>	
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center" bgcolor="<%=adminColor("green")%>"><br>
		<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="frmreg.submit();" style="cursor:pointer">
		<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="frmreg.reset();" style="cursor:pointer">
		<img src="http://testwebadmin.10x10.co.kr/images/icon_new_registration.gif" border="0" onClick="location.href='/admin/organizer/organizerReg.asp';" style="cursor:pointer">
	</td>
</tr>
</form>
</table>


<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="2009">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- 리스트 끝 -->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->