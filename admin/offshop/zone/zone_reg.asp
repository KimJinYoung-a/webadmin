<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 삽별구역설정
' Hieditor : 2010.12.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->
<%
Dim ozone,idx , i , shopid,zonename,racktype,unit,orderno,regdate ,isusing , zonegroup
dim menupos
	idx = requestCheckVar(request("idx"),10)
	menupos = requestCheckVar(request("menupos"),10)

set ozone = new czone_list
	ozone.frectidx = idx
	
	'//수정시에만 쿼리
	if idx <> "" then		
		ozone.fzone_oneitem()
		
		if ozone.ftotalcount >0 then			
			shopid = ozone.FOneItem.fshopid
			zonegroup = ozone.FOneItem.fzonegroup
			racktype = ozone.FOneItem.fracktype			
			zonename = ozone.FOneItem.fzonename
			unit = ozone.FOneItem.funit			
			regdate = ozone.FOneItem.fregdate
			isusing = ozone.FOneItem.fisusing						
		end if
	end if
	
%>

<script language="javascript">
	
	function reg(){
		if (frm.shopid.value=='') {
			alert('샵을 선택해 주세요');
			frm.zonename.focus();
			return;
		}

		if (frm.zonegroup.value=='') {
			alert('그룹 선택해 주세요');
			frm.zonegroup.focus();			
			return;
		}

		if (frm.racktype.value=='') {
			alert('매대 타입을 선택해 주세요');
			frm.racktype.focus();			
			return;
		}
		
		if (frm.zonename.value=='') {
			alert('구역명을 입력해 주세요');
			frm.zonename.focus();
			return;
		}
		
		if (frm.unit.value=='') {
			alert('해당지역의 평수를 입력해 주세요');
			frm.unit.focus();			
			return;
		}
		
		if(frm.unit.value!=''){
			if (!IsDouble(frm.unit.value)){
				alert('해당지역의 평수는 숫자만 가능합니다.');
				frm.unit.focus();
				return;
			}
		}	

		if (frm.isusing.value=='') {
			alert('사용여부를 선택해 주세요');
			frm.isusing.focus();			
			return;
		}
		
		frm.action='zone_process.asp';
		frm.mode.value = "zonereg";
		frm.submit();
	}
	
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ <font color="red">[중요] </font>매장내 구역이 변경되거나 없어지면,
		<br>기존 구역을 수정하지 마시고, 사용안함 돌리신후, 새로 등록하세요.
		<br>기존 구역을 현재 변경될 구역으로 수정후 사용 하실경우,
		<br>기존 구역으로 등록되어진 상품들이 모두 현재 구역으로 변경되는 문제가 발생됩니다
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td align="center">번호<br></td>
	<td>
		<%=idx%><input type="hidden" name="idx" value="<%=idx%>">
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center">SHOP</td>
	<td>
		<% drawSelectBoxOffShop "shopid",shopid %>
	</td>
</tr>
	
<tr bgcolor="#FFFFFF">
	<td align="center">그룹</td>
	<td>
		<% drawSelectBoxOffShopzonegroup "zonegroup",zonegroup,"" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">매대타입</td>
	<td>
		<% drawSelectBoxOffShopracktype "racktype",racktype,"" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">상세구역</td>
	<td>
		<input type="text" name="zonename" value="<%=zonename%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">UNIT</td>
	<td>
		<input type="text" name="unit" value="<%=unit%>" size=5 maxlength=5> ex)1
		<Br>※ 해당지역의 평수로 사용하시거나, 유동적으로 편하신대로 지정해서 사용하시면 됩니다
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">사용여부<br></td>
	<td>
		<select name="isusing">
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=2>
		<input type="button" value="저장" class="button" onclick="reg();">
	</td>
</tr>
</form>
</table>	

<% set ozone = nothing %>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
