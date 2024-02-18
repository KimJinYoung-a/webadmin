<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 옵션관리
' Hieditor : 서동석 생성
'			 2022.07.06 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/optionmanagecls.asp"-->
<%
dim cdl, cdm
cdl = request("cdl")
cdm = request("cdm")

dim onlyusing, ordertype, research
onlyusing = request("onlyusing")
ordertype = request("ordertype")
research = request("research")

if (onlyusing="") and (research="") then onlyusing="on"
if ordertype="" then ordertype="c"

dim ooption
set ooption = new COptionManager
ooption.FRectOnlyUsing = onlyusing
ooption.FRectOrderType= ordertype
ooption.GetOption01List


dim subooption
set subooption = new COptionManager
subooption.FRectOnlyUsing = onlyusing
subooption.FRectOrderType= ordertype
if cdl<>"" then
	subooption.GetOption02List cdl
end if
dim i
%>
<script type='text/javascript'>

function AddCode(cdl,cdm){
	var popwin = window.open('editoptioncode.asp?pmode=add&cdl=' + cdl + '&cdm=' + cdm,'editoptioncode','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditCode(cdl,cdm){
	var popwin = window.open('editoptioncode.asp?cdl=' + cdl + '&cdm=' + cdm,'editoptioncode','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function DelCode(cdl,cdm){
	alert('현재 삭제 불가능 합니다.');
	return;

	if (confirm('옵션 코드를 삭제 하시겠습니까?')){
		frmdel.cdl.value = cdl;
		frmdel.cdm.value = cdm;
		frmdel.submit();
	}
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<input type="radio" name="ordertype" value="c" <% if ordertype="c" then response.write "checked" %> >코드
			<input type="radio" name="ordertype" value="d" <% if ordertype="d" then response.write "checked" %> >순서					
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type=checkbox name="onlyusing" <% if onlyusing="on" then response.write "checked" %> >사용하는옵션만보기			
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<form name=frmdel method=post action="" style="margin:0px;">
<input type="hidden" name=mode value="delcode">
<input type="hidden" name=cdl value="">
<input type="hidden" name=cdm value="">
</form>
<table width=700 class="a">
<tr>
	<td><input type=button value="대분류추가" onclick="AddCode('','');"></td>
	<td></td>
	<% if cdl<>"" then %>
	<td><input type=button value="중분류추가" onclick="AddCode('<%= cdl %>','');"></td>
	<% else %>
	<td></td>
	<% end if %>
</tr>
<tr>
	<td valign=top>
		<table width=390 class="a" border=0 cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
		<tr>
			<td>코드명</td>
			<td>사용</td>
			<td>순서</td>
			<td>수정</td>
			<td>삭제</td>
		</tr>
		<% for i=0 to ooption.FResultCount-1 %>
		<tr bgcolor="#FFFFFF">
		<% if cdl=ooption.FItemList(i).Foptioncode01 then %>
			<td>
				<b><a href="?cdl=<%= ooption.FItemList(i).Foptioncode01 %>&onlyusing=<%= onlyusing %>&ordertype=<%= ordertype %>&research=<%= research %>">
				[<%= ooption.FItemList(i).Foptioncode01 %>]<%= ReplaceBracket(ooption.FItemList(i).Fcodename) %></a></b>
			</td>
		<% else %>
			<td>
				<a href="?cdl=<%= ooption.FItemList(i).Foptioncode01 %>&onlyusing=<%= onlyusing %>&ordertype=<%= ordertype %>&research=<%= research %>">
				[<%= ooption.FItemList(i).Foptioncode01 %>]<%= ReplaceBracket(ooption.FItemList(i).Fcodename) %></a>
			</td>
		<% end if %>
			<td width=40><%= ooption.FItemList(i).Foptiondispyn01 %></td>
			<td width=40><%= ooption.FItemList(i).Fdisporder01 %></td>
			<% if ooption.FItemList(i).Foptioncode01="00" then %>
			<td width=30 align=center>&nbsp;</td>
			<td width=30 align=center>&nbsp;</td>
			<% else %>
			<td width=30 align=center><a href="javascript:EditCode('<%= ooption.FItemList(i).Foptioncode01 %>','');">수정</a></td>
			<td width=30 align=center><a href="javascript:DelCode('<%= ooption.FItemList(i).Foptioncode01 %>','');">x</a></td>
			<% end if %>
		</tr>
		<% next %>
		</table>
	</td>
	<td width=20></td>
	<td valign=top>
		<table width=390 class="a" border=0 cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
		<tr>
			<td>코드명</td>
			<td>사용</td>
			<td>순서</td>
			<td>수정</td>
			<td>삭제</td>
		</tr>
		<% for i=0 to subooption.FResultCount-1 %>
		<tr bgcolor="#FFFFFF">
		<% if cdm=subooption.FItemList(i).Foptioncode02 then %>
			<td>
				<b><a href="?cdl=<%= subooption.FItemList(i).Foptioncode01 %>&cdm=<%= subooption.FItemList(i).Foptioncode02 %>&onlyusing=<%= onlyusing %>&ordertype=<%= ordertype %>&research=<%= research %>">
				[<%= subooption.FItemList(i).Foptioncode02 %>]<%= ReplaceBracket(subooption.FItemList(i).Fcodeview) %></a>
			</td>
		<% else %>
			<td>
				<a href="?cdl=<%= subooption.FItemList(i).Foptioncode01 %>&cdm=<%= subooption.FItemList(i).Foptioncode02 %>&onlyusing=<%= onlyusing %>&ordertype=<%= ordertype %>&research=<%= research %>">
				[<%= subooption.FItemList(i).Foptioncode02 %>]<%= ReplaceBracket(subooption.FItemList(i).Fcodeview) %></a>
			</td>
		<% end if %>
			<td width=40><%= subooption.FItemList(i).Foptiondispyn02 %></td>
			<td width=40><%= subooption.FItemList(i).Fdisporder02 %></td>
			<td width=30 align=center><a href="javascript:EditCode('<%= subooption.FItemList(i).Foptioncode01 %>','<%= subooption.FItemList(i).Foptioncode02 %>');">수정</a></td>
			<td width=30 align=center><a href="javascript:DelCode('<%= subooption.FItemList(i).Foptioncode01 %>','<%= subooption.FItemList(i).Foptioncode02 %>');">x</a></td>
		</tr>
		<% next %>
		</table>
	</td>
</tr>
</table>
<%
set ooption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->