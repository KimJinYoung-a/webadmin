<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfighterCls.asp"-->
<%
dim idx,mode
idx = request("idx")
If idx = "" Then idx=0
mode = request("mode")


dim oevent
dim cd1
set oevent = new CDesignFighterDetail
oevent.FRectidx = idx
oevent.GetDesignFighterItem
%>

<script language='javascript'>
function SubmitForm(){
	if (document.SubmitFrm.itemid1.value.length < 1){
		alert('아이템번호를 선택하세요');
		document.SubmitFrm.itemid1.focus();
		return;
	}
	else if (document.SubmitFrm.itemid2.value.length < 1){
		alert('아이템번호를 선택하세요');
		document.SubmitFrm.itemid2.focus();
		return;
	}
	else if (confirm('저장 하시겠습니까?')) {
		document.SubmitFrm.submit();
	}
}

function selcorner(num){
	for(var i=1;i<5;i++){
	
		var frm=(eval('corner'+i+'.style'));
		
		if (num==i){
			if (frm.display=="none"){
				frm.display="block";
			}else{
				frm.display="none";
			}
			
		}else {
			frm.display="none";
		}
	}
	
}

function MakeCommonUpdate(fighterid){
	if (confirm("업데이트를 하시겠습니까??????")){
	var popwin=window.open('/admin/sitemaster/lib/dofighterupdate.asp?idx='+ fighterid ,'fighterfresh','width=100,height=100');
	popwin.focus();
	}
}


</script>
<style>
.DFbutton1 {font-family: "Verdana", "돋움";	font-size: 9pt;	background-color: #DCDCDC; border: 1px outset #BABABA; color: #000000; height: 20px; cursor:pointer;}
.DFbutton2 {font-family: "Verdana", "돋움";	font-size: 9pt;	background-color: #FFB0D9; border: 1px outset #BABABA; color: #000000; height: 20px; cursor:pointer;}
.DFbutton3 {font-family: "Verdana", "돋움";	font-size: 9pt;	background-color: #CEBEE1; border: 1px outset #BABABA; color: #000000; height: 20px; cursor:pointer;}
.DFbutton4 {font-family: "Verdana", "돋움";	font-size: 9pt;	background-color: #ACFFEF; border: 1px outset #BABABA; color: #000000; height: 20px; cursor:pointer;}

.DFCorner1 {font:9pt/135% "굴림";color:#000000;	background-color: #DCDCDC; }
.DFCorner2 {font:9pt/135% "굴림";color:#000000;	background-color: #FFB0D9; }
.DFCorner3 {font:9pt/135% "굴림";color:#000000;	background-color: #CEBEE1; }
.DFCorner4 {font:9pt/135% "굴림";color:#000000;	background-color: #ACFFEF; }

</style>
<br>
<table width="800"   cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="SubmitFrm" method="post" action="<%=uploadUrl%>/linkweb/dodesignfighter.asp" onsubmit="return false;" enctype="multipart/form-data">
	<input type="hidden" name="mode" value="<% = mode %>">
	<input type="hidden" name="idx" value="<% = idx %>">
	
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<a href="http://www.10x10.co.kr/designfighter/design_fighter_preview.asp?page=1&idx=<%= idx %>" target="_blank"><b>미리보기</b></a>
		</td>
		<td colspan="3">
			<input type="button" class="button" value="오픈 전 업데이트" onclick="MakeCommonUpdate('<%= idx %>')">&nbsp;(너무 자주 누르지 마세욧)
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="110"   align="center">구분</td>
		<td align="center"  >아이템 1</td>
		<td align="center"  >아이템 2</td>
		<td width="60"  align="center" >비고 </td>
	</tr>
	
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">상품 코드</td>
		<td bgcolor="#FFFFFF"><input type="text" name="itemid1" value="<% = oevent.Fitemid1 %>" size="20" class="input_b"></td>
		<td bgcolor="#FFFFFF"><input type="text" name="itemid2" value="<% = oevent.Fitemid2 %>" size="20" class="input_b"></td>
		<td bgcolor="#FFFFFF" align="center">&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">상품명</td>
		<td bgcolor="#FFFFFF"><input type="text" name="itemname1" value="<% = oevent.Fitemname1 %>" size="32" class="input_b" maxlength="32"></td>
		<td bgcolor="#FFFFFF"><input type="text" name="itemname2" value="<% = oevent.Fitemname2 %>" size="32" class="input_b" maxlength="32"></td>
		<td bgcolor="#FFFFFF" align="center">&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">전체 타이틀</td>
		<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="title" value="<% = oevent.Ftitle %>" size="60" class="input_b" maxlength="60"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">타이틀이미지</td>
		<td bgcolor="#FFFFFF" colspan="3"><input type="file" name="titleimg" size="32" maxlength="100" class="file">
			(<% = oevent.Ftitleimg %>)
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">메인이미지</td>
		<td bgcolor="#FFFFFF"><input type="file" name="mainimg1" size="32" maxlength="100" class="file"><br>(<% = oevent.Fmainimg1 %>)</td>
		<td bgcolor="#FFFFFF"><input type="file" name="mainimg2" size="32" maxlength="100" class="file"><br>(<% = oevent.Fmainimg2 %>)</td>
		<td bgcolor="#FFFFFF" align="center">300 x 300</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">아이콘이미지</td>
		<td bgcolor="#FFFFFF"><input type="file" name="icon1" size="32" maxlength="100" class="file"><br>(<% = oevent.Ficon1 %>)</td>
		<td bgcolor="#FFFFFF"><input type="file" name="icon2" size="32" maxlength="100" class="file"><br>(<% = oevent.Ficon2 %>)</td>
		<td bgcolor="#FFFFFF" align="center">80 x 80 </td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">small아이콘이미지</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sicon1" size="32" maxlength="100" class="file"><br>(<% = oevent.Fsicon1 %>)</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sicon2" size="32" maxlength="100" class="file"><br>(<% = oevent.Fsicon2 %>)</td>
		<td bgcolor="#FFFFFF" align="center"> 50 x 50 </td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">리스트 배너 </td>
		<td bgcolor="#FFFFFF" colspan="2"><input type="file" name="banimg" size="32" maxlength="32" class="file">
		(<% = oevent.Fbanimg %>)
		</td>
		<td bgcolor="#FFFFFF" align="center"> 180 x 98</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">승리선택</td>
		<td  bgcolor="#FFFFFF" colspan="3">
			<input type="checkbox" name="winY" <% if oevent.Fwinyn="Y" then response.write "checked" %> >1번상품승리
			<input type="checkbox" name="winN" <% if oevent.Fwinyn="N" then response.write "checked" %> >2번상품승리
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">사용유무</td>
		<td  bgcolor="#FFFFFF" colspan="3">
			<input type="radio" name="isusing" value="Y" <% if oevent.FIsUsing="Y" then response.write "checked" %> >Y
			<input type="radio" name="isusing" value="N" <% if oevent.FIsUsing="N" Or oevent.FIsUsing="" then response.write "checked" %> >N
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td colspan="4" height="10">&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"align="center">코너 선택</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="button" value="코너1" onclick="selcorner('1');" class="DFbutton1">
			<input type="button" value="코너2" onclick="selcorner('2');" class="DFbutton2">
			<input type="button" value="코너3" onclick="selcorner('3');" class="DFbutton3">
			<input type="button" value="코너4" onclick="selcorner('4');" class="DFbutton4">
		</td>
	</tr>
	<tr>
		<td width="100%" colspan="4"  bgcolor="#FFFFFF">
			<!-- 코너 1 -->
			<div id="corner1" style="display:none">
			<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" class="a">
				<tr>
					<td width="110"></td>
					<td></td>
					<td></td>
					<td width="60"></td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1">타이틀</td>
					<td colspan="3">
					<font color="red">+  타이틀을 입력하신 코너만 프론트 화면에 보여집니다.</font><br>
					<input type="text" name="title1" value="<% = oevent.Ftitle1 %>" size="70" class="input_b"></td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 이미지 1</td>
					<td><input type="file" name="img1_1" size="32" class="file"><br>(<% = oevent.Fimg1_1 %>)</td>
					<td><input type="file" name="img2_1" size="32"  class="file"><br>(<% = oevent.Fimg2_1 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 내용 1</td>
					<td><textarea name="contents1_1" rows="10" cols="45" class="input_b"><% = oevent.Fcontents1_1 %></textarea></td>
					<td><textarea name="contents2_1" rows="10" cols="45" class="input_b"><% = oevent.Fcontents2_1 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 이미지 2</td>
					<td><input type="file" name="img1_2" size="32" class="file"><br>(<% = oevent.Fimg1_2 %>)</td>
					<td><input type="file" name="img2_2" size="32"  class="file"><br>(<% = oevent.Fimg2_2 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 내용 2</td>
					<td><textarea name="contents1_2" rows="10" cols="45" class="input_b"><% = oevent.Fcontents1_2 %></textarea></td>
					<td><textarea name="contents2_2" rows="10" cols="45" class="input_b"><% = oevent.Fcontents2_2 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 이미지 3</td>
					<td><input type="file" name="img1_3" size="32" class="file"><br>(<% = oevent.Fimg1_3 %>)</td>
					<td><input type="file" name="img2_3" size="32"  class="file"><br>(<% = oevent.Fimg2_3 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 내용 3</td>
					<td><textarea name="contents1_3" rows="10" cols="45" class="input_b"><% = oevent.Fcontents1_3 %></textarea></td>
					<td><textarea name="contents2_3" rows="10" cols="45" class="input_b"><% = oevent.Fcontents2_3 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 이미지 4</td>
					<td><input type="file" name="img1_4" size="32" class="file"><br>(<% = oevent.Fimg1_4 %>)</td>
					<td><input type="file" name="img2_4" size="32"  class="file"><br>(<% = oevent.Fimg2_4 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 내용 4</td>
					<td><textarea name="contents1_4" rows="10" cols="45" class="input_b"><% = oevent.Fcontents1_4 %></textarea></td>
					<td><textarea name="contents2_4" rows="10" cols="45" class="input_b"><% = oevent.Fcontents2_4 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 이미지 5</td>
					<td><input type="file" name="img1_5" size="32" class="file"><br>(<% = oevent.Fimg1_5 %>)</td>
					<td><input type="file" name="img2_5" size="32"  class="file"><br>(<% = oevent.Fimg2_5 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner1"> 내용5</td>
					<td><textarea name="contents1_5" rows="10" cols="45" class="input_b"><% = oevent.Fcontents1_5 %></textarea></td>
					<td><textarea name="contents2_5" rows="10" cols="45" class="input_b"><% = oevent.Fcontents2_5 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
			</table>
			</div>
			<!-- 코너 1 끝 -->
			<!-- 코너 2 -->
			<div id="corner2" style="display:none">
			<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" class="a">
				<tr>
					<td width="110"></td>
					<td></td>
					<td></td>
					<td width="60"></td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 타이틀</td>
					<td colspan="3"><font color="red">+  타이틀을 입력하신 코너만 프론트 화면에 보여집니다.</font><br><input type="text" name="title2" value="<% = oevent.Ftitle2 %>" size="70" class="input_b"></td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 이미지 1</td>
					<td><input type="file" name="img3_1" size="32" class="file"><br>(<% = oevent.Fimg3_1 %>)</td>
					<td><input type="file" name="img4_1" size="32"  class="file"><br>(<% = oevent.Fimg4_1 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 내용 1</td>
					<td><textarea name="contents3_1" rows="10" cols="45" class="input_b"><% = oevent.Fcontents3_1 %></textarea></td>
					<td><textarea name="contents4_1" rows="10" cols="45" class="input_b"><% = oevent.Fcontents4_1 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 이미지 2</td>
					<td><input type="file" name="img3_2" size="32" class="file"><br>(<% = oevent.Fimg3_2 %>)</td>
					<td><input type="file" name="img4_2" size="32"  class="file"><br>(<% = oevent.Fimg4_2 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 내용 2</td>
					<td><textarea name="contents3_2" rows="10" cols="45" class="input_b"><% = oevent.Fcontents3_2 %></textarea></td>
					<td><textarea name="contents4_2" rows="10" cols="45" class="input_b"><% = oevent.Fcontents4_2 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 이미지 3</td>
					<td><input type="file" name="img3_3" size="32" class="file"><br>(<% = oevent.Fimg3_3 %>)</td>
					<td><input type="file" name="img4_3" size="32"  class="file"><br>(<% = oevent.Fimg4_3 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 내용 3</td>
					<td><textarea name="contents3_3" rows="10" cols="45" class="input_b"><% = oevent.Fcontents3_3 %></textarea></td>
					<td><textarea name="contents4_3" rows="10" cols="45" class="input_b"><% = oevent.Fcontents4_3 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 이미지 4</td>
					<td><input type="file" name="img3_4" size="32" class="file"><br>(<% = oevent.Fimg3_4 %>)</td>
					<td><input type="file" name="img4_4" size="32"  class="file"><br>(<% = oevent.Fimg4_4 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 내용 4</td>
					<td><textarea name="contents3_4" rows="10" cols="45" class="input_b"><% = oevent.Fcontents3_4 %></textarea></td>
					<td><textarea name="contents4_4" rows="10" cols="45" class="input_b"><% = oevent.Fcontents4_4 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 이미지 5</td>
					<td><input type="file" name="img3_5" size="32" class="file"><br>(<% = oevent.Fimg3_5 %>)</td>
					<td><input type="file" name="img4_5" size="32"  class="file"><br>(<% = oevent.Fimg4_5 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner2"> 내용 5</td>
					<td><textarea name="contents3_5" rows="10" cols="45" class="input_b"><% = oevent.Fcontents3_5 %></textarea></td>
					<td><textarea name="contents4_5" rows="10" cols="45" class="input_b"><% = oevent.Fcontents4_5 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
			</table>
			</div>
			<!-- 코너 2 끝 -->
			
			<!-- 코너 3 -->
			<div id="corner3" style="display:none">
			<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" class="a">
				<tr>
					<td width="110"></td>
					<td></td>
					<td></td>
					<td width="60"></td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 타이틀</td>
					<td colspan="3"><font color="red">+  타이틀을 입력하신 코너만 프론트 화면에 보여집니다.</font><br><input type="text" name="title3" value="<% = oevent.Ftitle3 %>" size="70" class="input_b"></td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 이미지 1</td>
					<td><input type="file" name="img5_1" size="32" class="file"><br>(<% = oevent.Fimg5_1 %>)</td>
					<td><input type="file" name="img6_1" size="32"  class="file"><br>(<% = oevent.Fimg6_1 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 내용 1</td>
					<td><textarea name="contents5_1" rows="10" cols="45" class="input_b"><% = oevent.Fcontents5_1 %></textarea></td>
					<td><textarea name="contents6_1" rows="10" cols="45" class="input_b"><% = oevent.Fcontents6_1 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 이미지 2</td>
					<td><input type="file" name="img5_2" size="32" class="file"><br>(<% = oevent.Fimg5_2 %>)</td>
					<td><input type="file" name="img6_2" size="32"  class="file"><br>(<% = oevent.Fimg6_2 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 내용 2</td>
					<td><textarea name="contents5_2" rows="10" cols="45" class="input_b"><% = oevent.Fcontents5_2 %></textarea></td>
					<td><textarea name="contents6_2" rows="10" cols="45" class="input_b"><% = oevent.Fcontents6_2 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 이미지 3</td>
					<td><input type="file" name="img5_3" size="32" class="file"><br>(<% = oevent.Fimg5_3 %>)</td>
					<td><input type="file" name="img6_3" size="32"  class="file"><br>(<% = oevent.Fimg6_3 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 내용 3</td>
					<td><textarea name="contents5_3" rows="10" cols="45" class="input_b"><% = oevent.Fcontents5_3 %></textarea></td>
					<td><textarea name="contents6_3" rows="10" cols="45" class="input_b"><% = oevent.Fcontents6_3 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 이미지 4</td>
					<td><input type="file" name="img5_4" size="32" class="file"><br>(<% = oevent.Fimg5_4 %>)</td>
					<td><input type="file" name="img6_4" size="32"  class="file"><br>(<% = oevent.Fimg6_4 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 내용 4</td>
					<td><textarea name="contents5_4" rows="10" cols="45" class="input_b"><% = oevent.Fcontents5_4 %></textarea></td>
					<td><textarea name="contents6_4" rows="10" cols="45" class="input_b"><% = oevent.Fcontents6_4 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 이미지 5</td>
					<td><input type="file" name="img5_5" size="32" class="file"><br>(<% = oevent.Fimg5_5 %>)</td>
					<td><input type="file" name="img6_5" size="32"  class="file"><br>(<% = oevent.Fimg6_5 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner3"> 내용 5</td>
					<td><textarea name="contents5_5" rows="10" cols="45" class="input_b"><% = oevent.Fcontents5_5 %></textarea></td>
					<td><textarea name="contents6_5" rows="10" cols="45" class="input_b"><% = oevent.Fcontents6_5 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
			</table>
			</div>
			<!-- 코너 3 끝 -->
			
			<!-- 코너 4 -->
			<div id="corner4" style="display:none">
			<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" class="a">
				<tr>
					<td width="110"></td>
					<td></td>
					<td></td>
					<td width="60"></td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 타이틀</td>
					<td colspan="3"><font color="red">+  타이틀을 입력하신 코너만 프론트 화면에 보여집니다.</font><br><input type="text" name="title4" value="<% = oevent.Ftitle4 %>" size="70" class="input_b"></td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 이미지 1</td>
					<td><input type="file" name="img7_1" size="32" class="file"><br>(<% = oevent.Fimg7_1 %>)</td>
					<td><input type="file" name="img8_1" size="32"  class="file"><br>(<% = oevent.Fimg8_1 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 내용 1</td>
					<td><textarea name="contents7_1" rows="10" cols="45" class="input_b"><% = oevent.Fcontents7_1 %></textarea></td>
					<td><textarea name="contents8_1" rows="10" cols="45" class="input_b"><% = oevent.Fcontents8_1 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 이미지 2</td>
					<td><input type="file" name="img7_2" size="32" class="file"><br>(<% = oevent.Fimg7_2 %>)</td>
					<td><input type="file" name="img8_2" size="32"  class="file"><br>(<% = oevent.Fimg8_2 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 내용 2</td>
					<td><textarea name="contents7_2" rows="10" cols="45" class="input_b"><% = oevent.Fcontents7_2 %></textarea></td>
					<td><textarea name="contents8_2" rows="10" cols="45" class="input_b"><% = oevent.Fcontents8_2 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 이미지 3</td>
					<td><input type="file" name="img7_3" size="32" class="file"><br>(<% = oevent.Fimg7_3 %>)</td>
					<td><input type="file" name="img8_3" size="32"  class="file"><br>(<% = oevent.Fimg8_3 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 내용 3</td>
					<td><textarea name="contents7_3" rows="10" cols="45" class="input_b"><% = oevent.Fcontents7_3 %></textarea></td>
					<td><textarea name="contents8_3" rows="10" cols="45" class="input_b"><% = oevent.Fcontents8_3 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 이미지 4</td>
					<td><input type="file" name="img7_4" size="32" class="file"><br>(<% = oevent.Fimg7_4 %>)</td>
					<td><input type="file" name="img8_4" size="32"  class="file"><br>(<% = oevent.Fimg8_4 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 내용 4</td>
					<td><textarea name="contents7_4" rows="10" cols="45" class="input_b"><% = oevent.Fcontents7_4 %></textarea></td>
					<td><textarea name="contents8_4" rows="10" cols="45" class="input_b"><% = oevent.Fcontents8_4 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 이미지 5</td>
					<td><input type="file" name="img7_5" size="32" class="file"><br>(<% = oevent.Fimg7_5 %>)</td>
					<td><input type="file" name="img8_5" size="32"  class="file"><br>(<% = oevent.Fimg8_5 %>)</td>
					<td align="center">350 x 350</td>
				</tr>
				<tr>
					<td align="center" class="DFcorner4"> 내용 5</td>
					<td><textarea name="contents7_5" rows="10" cols="45" class="input_b"><% = oevent.Fcontents7_5 %></textarea></td>
					<td><textarea name="contents8_5" rows="10" cols="45" class="input_b"><% = oevent.Fcontents8_5 %></textarea></td>
					<td align="center">&nbsp;</td>
				</tr>
			</table>
			</div>
			<!-- 코너 4 끝 -->
		</td>
	</tr>
	
	<tr  bgcolor="#FFFFFF">
		<td colspan="4" align="center">
		<input type="image" src="/images/icon_save.gif" onClick="SubmitForm();">
		<a href="design_fighter.asp?menupos=<%=menupos%>"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>
	
	
</form>
</table>
<%
set oevent = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->