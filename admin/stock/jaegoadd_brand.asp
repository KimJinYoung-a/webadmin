<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  재고파악 브랜드별 재고저장
' History : 2007년 10월 31일 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim makerid , i
	makerid = request("makerid")		'브랜드명 검색을 위한 변수
	if makerid = "" then
		makerid = "없음"
	end if 	
	
dim oip						'클래스선언
	set oip = new Cfitemlist		'변수에 토탈을 넣구
	oip.frectmakerid = makerid
	oip.fbrandinsert()	
%>

<script language="javascript">
	function AnSelectAllFrame(bool){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.disabled!=true){
					frm.cksel.checked = bool;
					AnCheckClick(frm.cksel);
				}
			}
		}
	}	
		
	function AnCheckClick(e){
		if (e.checked)
			hL(e);
		else
			dL(e);
	}	
	
	function ckAll(icomp){
		var bool = icomp.checked;
		AnSelectAllFrame(bool);
	}
	
	function CheckSelected(){
		var pass=false;
		var frm;
	
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				pass = ((pass)||(frm.cksel.checked));
			}
		}
	
		if (!pass) {
			return false;
		}
		return true;
	}
	
	function reg(upfrm){
	if (!CheckSelected()){
			alert('선택아이템이 없습니다.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.drawitemid.value = upfrm.drawitemid.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.drawitemid.value;
			var aa;
			aa = window.open("jaegoadd_brand_process.asp?drawitemid=" +tot, "reg","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
</script>	
<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get">
	<input type="hidden" name="drawitemid">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>브랜드별 재고파악 수량등록</strong> / 동일상품이 재고파악중일경우 등록되지 않습니다. </font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td><br>브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>&nbsp;
			<input type=button value="검색" onclick="document.frm.submit();">
			<br><br>
			<% if oip.ftotalcount > 0 then %>
				<input type="button" value="등록" onclick="javascript:reg(frm);">
			<% end if %>
			<font color="red">한상품의 옵션이 여러개 존재 할경우 그중 한개만 선택 하세요.</font>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
	</tr>
	</form>
</table>
<!--표 헤드끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<% if oip.ftotalcount > 0 then %>	 <!--레코드 수가 0보다 크면 -->
    <tr align="center" bgcolor="#DDDDFF">
   		<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td>이미지</td>
		<td>상품코드</td>
		<td>브랜드id</td>
		<td>상품명</td>
		<td>옵션코드</td>
		<td>옵션명</td>	
		<td>재고파악용재고</td>
		</tr>

	<% for i=0 to oip.ftotalcount - 1 %>
		<form action="jaegoadd_brand_process.asp" name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->
		<tr bgcolor="#FFFFFF">
			<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>	
			<td><img src="<%= oip.flist(i).fsmallimage %>" width=50 height=50><input type="hidden" name="smallimage" value="<%= oip.flist(i).fsmallimage %>"></td>	<!--'이미지 -->
			<td><%= oip.flist(i).fitemid %><input type="hidden" name="itemid" value="<%= oip.flist(i).fitemid %>"></td>				 					<!--'상품번호	 -->
			<td><%= oip.flist(i).fmakerid %><input type="hidden" name="makerid" value="<%= oip.flist(i).fmakerid %>"></td>									 <!--'브랜드id -->
			<td><%= oip.flist(i).fitemname %><input type="hidden" name="itemname" value="<%= oip.flist(i).fitemname %>"></td>									 <!--'상품명 -->
			<td><%= oip.flist(i).fitemoption %><input type="hidden" name="itemoption" value="<%= oip.flist(i).fitemoption %>"></td>							 <!--'옵션코드 -->
			<td><%= oip.flist(i).fitemoptionname %><input type="hidden" name="itemoptionname" value="<%= oip.flist(i).fitemoptionname %>"></td>				 <!--'옵션명 -->													
			<td><%= oip.flist(i).fbasicstock %><input type="hidden" name="basicstock" value="<%= oip.flist(i).fbasicstock %>"></td>								 <!--'재고파악사항 -->
		</tr>
	    </form>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
	<td colspan=15 align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="left">
        <% if oip.ftotalcount > 0 then %>
			<input type="button" value="등록" onclick="javascript:reg(frm);">
		<% end if %>
        <input type="button" value="닫기" onclick="javascript:window.close();">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/lib/db/dbclose.asp" -->