<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성놀이
' Hieditor : 2010.12.28 허진원 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oplay, i,page, playSn
Dim startdate, enddate, playLinkType, evt_code
	menupos = request("menupos")
	page = request("page")
	playSn = request("playSn")
	if page = "" then page = 1

'// 놀이정보 접수
set oplay = new cplayList
	oplay.frectplaySn = playSn
	oplay.FPageSize = 1
	oplay.FCurrPage = 1
	oplay.fplay_list()
	
	if oplay.ftotalcount > 0 then
		startdate = oplay.FItemList(0).fstartdate
		enddate = oplay.FItemList(0).fenddate
		playLinkType = oplay.FItemList(0).fplayLinkType
		evt_code = oplay.FItemList(0).fevtCode
	end if
set oplay = Nothing

'// 리스트
set oplay = new cPlayList
	oplay.FPageSize = 20
	oplay.FCurrPage = page
	oplay.frectplaySn = playSn	
	oplay.fitem_list()
%>

<script language="javascript">
	function AddIttems(){
		var ret = confirm(arrFrm.itemid.value + '아이템을 추가하시겠습니까?');
		if (ret){
			arrFrm.itemid.value = arrFrm.itemid.value;
			arrFrm.mode.value="itemAdd";
			arrFrm.submit();
		}
	}
	
	function AddIttems2(){
		if (document.arrFrm.itemidarr.value == ""){
			alert("아이템번호를  적어주세요!");
			document.arrFrm.itemidarr.focus();
		}
		else if (confirm(arrFrm.itemidarr.value + '아이템을 추가하시겠습니까?')){
			arrFrm.itemid.value = arrFrm.itemidarr.value;
			arrFrm.mode.value="itemAdd";
			arrFrm.submit();
		}
	}

	function delitem(upfrm){
		if (!CheckSelected()){
			alert('선택아이템이 없습니다.');
			return;
		}
	
		var ret = confirm('선택 아이템을 삭제하시겠습니까?');
	
		if (ret){
			var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.plyItemSn.value = upfrm.plyItemSn.value + frm.plyItemSn.value + "," ;
					}
				}
			}
			upfrm.mode.value="itemDel";
			upfrm.submit();
	
		}
	}

	function popItemWindow(tgf){
		var popup_item = window.open("/common/pop_CateItemList.asp?target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
		popup_item.focus();
	}

	function addFromEvt() {
		if(!document.arrFrm.evt_code.value) {
			alert("이벤트가 지정안되었습니다.\n감성놀이 관리에서 이벤트를 지정해주세요!");
		} else if(confirm('이벤트에 등록된 상품을 가져오시겠습니까?\n※가져오기 실행시 기존에 입력된 상품은 모두 삭제됩니다.')){
			arrFrm.mode.value="evtItemAdd";
			arrFrm.submit();
		}
	}

	function CheckSelected(){
		var pass=false;
		var frm;
	
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
	
			if (frm.name.indexOf('frmBuyPrc')!= -1) {
	
				pass = ((pass)||(frm.cksel.checked));
			}
	
		}
	
		if (!pass) {
			return false;
		}
		return true;
	}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">번호</td>
	<td width="160" align="left"><%=playSn%></td>
	<td width="80" bgcolor="<%= adminColor("gray") %>">진행형식</td>
	<td align="left"><%=chkIIF(playLinkType="E","이벤트 [" & evt_code & "]","직접설정")%></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">기간</td>
	<td colspan="3" align="left"><%=startdate & " ~ " & enddate%></td>
</tr>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="arrFrm" method="post" action="play_Process.asp">
<input type="hidden" name="playSn" value="<%=playSn%>">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
<input type="hidden" name="plyItemSn">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
	<tr>
		<td colspan="2" align="right" style="padding:10px 0 3px 0;">
			<input type="text" name="itemidarr" value="" size="70" class="input">
			<input type="button" value="상품 직접추가" onclick="AddIttems2()" class="button">
		</td>
	</tr>
	<tr>
		<td align="left" style="padding-bottom:5px">
			<input type="button" onclick="delitem(arrFrm);" value="선택상품삭제" class="button">
		</td>
		<td align="right" style="padding-bottom:5px;">
			<% if playLinkType="E" then %><input type="button" onclick="addFromEvt()" value="이벤트상품추가" class="button"><% end if %>
			<input type="button" onclick="popItemWindow('arrFrm.itemid');" value="상품추가" class="button">
		</td>
	</tr>
</form>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oplay.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7">
		검색결과 : <b><%= oplay.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oplay.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">상품번호</td>
	<td align="center">이미지</td>	
	<td align="center">브랜드</td>
	<td align="center">상품명</td>
	<td align="center">판매가</td>
	<td align="center">판매여부</td>
</tr>
<% for i=0 to oplay.FresultCount-1 %>			
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="plyItemSn" value="<%= oplay.FItemList(i).fplyItemSn %>">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= oplay.FItemList(i).fitemid %></td>
	<td align="center"><img src="<%= oplay.FItemList(i).FImageSmall %>"></td>
	<td align="center"><%= oplay.FItemList(i).fmakerid %></td>
	<td align="center"><%= oplay.FItemList(i).fitemname %></td>
	<td align="center"><%= FormatNumber(oplay.FItemList(i).fsellcash,0) %></td>
	<td align="center"><%=chkIIF(oplay.FItemList(i).fisusing="Y" and oplay.FItemList(i).fsellyn="Y","판매","<font color=red>품절</font>")%></td>
</tr>   
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" height="50" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7" align="center">
       	<% if oplay.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oplay.StartScrollPage-1 %>&playSn=<%=playSn%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oplay.StartScrollPage to oplay.StartScrollPage + oplay.FScrollCount - 1 %>
			<% if (i > oplay.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oplay.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&playSn=<%=playSn%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oplay.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&playSn=<%=playSn%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oplay = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->