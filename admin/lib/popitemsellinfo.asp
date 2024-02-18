<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<!-- include virtual="/lib/classes/realjaegocls.asp"-->

<%
dim itemid
itemid = request("itemid")


response.redirect "/common/pop_simpleitemedit.asp?itemid=" + CStr(itemid)
''사용안함

dim ojaego
set ojaego = new CRealJaeGo
ojaego.FRectItemID = itemid

if itemid<>"" then
	ojaego.GetItemInfoWithDailyRealJaeGo
end if

dim i
%>
<script language='javascript'>

function EnabledCheck(comp){
	var frm = document.frm2;

	if (comp.value=="Y"){
		frm.limitno.disabled = false;
		frm.limitsold.disabled = false;
	}else{
		frm.limitno.disabled = true;
		frm.limitsold.disabled = true;
	}
}

function SaveItem(frm){
	if ((frm.itemrackcode.value.length>0)&&(frm.itemrackcode.value.length!=6)){
		alert('상품 랙코드는 6자리로 고정되어있습니다.');
		frm.itemrackcode.focus();
		return;
	}

	var ret = confirm('저장 하시겠습니까?');

	if(ret){
		frm.submit();
	}
}

function popoptionEdit(iid){
	var popwin = window.open('/admin/shopmaster/popitemoptionedit.asp?menupos=239&itemid=' + iid,'popitemoptionedit','width=440 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	상품코드 : <input type="text" name="itemid" value="<%= itemid %>" Maxlength="12" size="12">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<% if ojaego.FResultCount>0 then %>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frm2 method=post action="doitemsellinfo.asp">
<input type=hidden name=itemid value="<%= itemid %>">
	<tr bgcolor="#FFFFFF">
		<td width=90 bgcolor="#DDDDFF">상품코드</td>
		<td><%= itemid %></td>
		<td width=100 rowspan=5><img src="<%= ojaego.FITemList(0).FImageList %>" width=100></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=80 bgcolor="#DDDDFF">브랜드</td>
		<td><%= ojaego.FITemList(0).Fmakerid %> (<%= ojaego.FITemList(0).FBrandName %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=80 bgcolor="#DDDDFF">판매가/매입가</td>
		<td>
		    <%= FormatNumber(ojaego.FITemList(0).FSellcash,0) %> / <%= FormatNumber(ojaego.FITemList(0).FBuycash,0) %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=80 bgcolor="#DDDDFF">매입구분</td>
		<td>
		<font color="<%= ojaego.FITemList(0).getMwDivColor %>"><%= ojaego.FITemList(0).getMwDivName %></font>
		&nbsp;
		<% if ojaego.FITemList(0).FSellcash<>0 then %>
		<%= CLng((1- ojaego.FITemList(0).FBuycash/ojaego.FITemList(0).FSellcash)*100) %> %
		<% end if %>
    	</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=80 bgcolor="#DDDDFF">랙코드</td>
		<td>
		<input type="text" name="itemrackcode" value="<%= ojaego.FITemList(0).FitemRackCode %>" size="6" maxlength="6" > (6자리 Fix)
    	</td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr>
		<td bgcolor="#FFFFFF" colspan="2">
		<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr bgcolor="#FFFFFF">
				<td colspan="4" align="right">
				최종업데이트 (<%= ojaego.FITemList(0).FLastupdate %>)
				</td>
			</tr>
			<tr bgcolor="#FFDDDD">
				<td>상품명</td>
				<td>옵션명</td>
				<td>OLD SYS</td>
				<td>NEW SYS</td>
			</tr>
		<% for i=0 to ojaego.FResultCount - 1 %>
			<% if ojaego.FITemList(i).FOptionUsing="N" then %>
			<tr bgcolor="#DDDDDD">
			<% else %>
			<tr bgcolor="#FFFFFF">
			<% end if %>
				<td><%= ojaego.FITemList(i).FItemName %></td>
				<td><%= ojaego.FITemList(i).FItemOptionName %></td>
				<td><%= ojaego.FITemList(i).Foldstockcurrno %></td>
				<td><%= ojaego.FITemList(i).GetCheckStockNo %></td>
			</tr>
		<% next %>
		</table>
		</td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr>
		<td width=90 bgcolor="#DDDDFF">옵션</td>
		<td bgcolor="#FFFFFF">
		(<%= ojaego.FITemList(0).FOptionCnt %>)
		<input type=button value="옵션수정" onclick="popoptionEdit('<%= itemid %>');">
		</td>
	</tr>

	<tr>
		<td width=80 bgcolor="#DDDDFF">배송구분</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).IsUpcheBeasong then %>
		<b>업체</b>배송
		<% else %>
		텐바이텐배송
		<% end if %>
		</td>
	</tr>

	<tr>
		<td width=80 bgcolor="#DDDDFF">SoldOut</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).IsSoldOut then %>
		<font color=red><b>Sold Out</b></font>
		<% end if %>
		</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">한정판매여부</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FLimitYn="Y" then %>
		<input type="radio" name="limityn" value="Y" checked onclick="EnabledCheck(this)">한정판매
		<input type="radio" name="limityn" value="N" onclick="EnabledCheck(this)">비한정판매
		<% else %>
		<input type="radio" name="limityn" value="Y" onclick="EnabledCheck(this)">한정판매
		<input type="radio" name="limityn" value="N" checked onclick="EnabledCheck(this)">비한정판매
		<% end if %>
		</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">총한정수량</td>
		<td bgcolor="#FFFFFF"><input type="text" name="limitno" value="<%= ojaego.FITemList(0).FLimitNo %>" size="5" maxlength=5 <% if ojaego.FITemList(0).FLimitYn="N" then response.write "disabled" %> >개</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">판매된수량</td>
		<td bgcolor="#FFFFFF"><input type="text" name="limitsold" value="<%= ojaego.FITemList(0).FLimitSold %>" size="5" maxlength=5 <% if ojaego.FITemList(0).FLimitYn="N" then response.write "disabled" %> >개</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">남은수량</td>
		<td bgcolor="#FFFFFF"><input type="text" name="remainno" value="<%= ojaego.FITemList(0).GetLimitEa %>" size="5" maxlength=5 disabled >개</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">전시여부</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FDispYn="Y" then %>
		<input type="radio" name="dispyn" value="Y" checked >전시함
		<input type="radio" name="dispyn" value="N" >전시안함
		<% else %>
		<input type="radio" name="dispyn" value="Y" >전시함
		<input type="radio" name="dispyn" value="N" checked ><font color="red">전시안함</font>
		<% end if %>
		</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">판매여부</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FSellYn="Y" then %>
		<input type="radio" name="sellyn" value="Y" checked >판매함
		<input type="radio" name="sellyn" value="N" >판매안함
		<% else %>
		<input type="radio" name="sellyn" value="Y" >판매함
		<input type="radio" name="sellyn" value="N" checked ><font color="red">판매안함</font>
		<% end if %>
		</td>
	</tr>
	<tr>
		<td width=80 bgcolor="#DDDDFF">사용여부</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FIsUsing="Y" then %>
		<input type="radio" name="isusing" value="Y" checked >사용함
		<input type="radio" name="isusing" value="N" >사용안함
		<% else %>
		<input type="radio" name="isusing" value="Y" >사용함
		<input type="radio" name="isusing" value="N" checked ><font color="red">사용안함</font>
		<% end if %>
		</td>
	</tr>
	<input type=hidden name="pojangok" value="<%= ojaego.FITemList(0).FPojangOK %>">
<!--
	<tr>
		<td width=80 bgcolor="#DDDDFF">포장여부</td>
		<td bgcolor="#FFFFFF">
		<% if ojaego.FITemList(0).FPojangOK="Y" then %>
		<input type="radio" name="pojangok" value="Y" checked >포장가능
		<input type="radio" name="pojangok" value="N" >포장불가
		<% else %>
		<input type="radio" name="pojangok" value="Y" >포장가능
		<input type="radio" name="pojangok" value="N" checked ><font color="red">포장불가</font>
		<% end if %>
		</td>
	</tr>
-->
</form>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4"><input type="button" value="저장" onclick="SaveItem(frm2)"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<% end if %>
<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->