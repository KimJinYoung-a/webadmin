<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/deal_reg.asp
' Description :  딜 이벤트 등록
' History : 2017.08.23 정태훈
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
Dim idx, k, j, itemid
idx = requestCheckVar(Request("idx"),10)
If idx="" Then
Response.write "<script>alert('딜 정보가 없습니다.');history.back();</script>"
Response.End
End If
Dim oDeal, oitem, oitemimg, arrIMG
set oDeal = new CDealView
oDeal.FRectMasterIDX = idx
oDeal.GetDealView

Dim oDealitem, arrList, iTotCnt, intLoop
set oDealitem = new CDealItem
oDealitem.FRectMasterIDX = idx
arrList = oDealitem.fnGetDealEventItem	
iTotCnt = oDealitem.FTotCnt	'전체 데이터  수
Set oDealitem=Nothing
itemid=oDeal.Fdealitemid
set oitem = new CItem
oitem.FRectItemID = oDeal.Fdealitemid
oitem.GetOneItem
Dim vArr
set oitemimg = new CItemAddImage
oitemimg.FRectItemID = oDeal.Fdealitemid
vArr = oitemimg.GetAddImageListIMGTYPE1

Function FormatDatePart(div,vdate)
	If div = "h" Then
		FormatDatePart=DatePart("h",vdate)
		If FormatDatePart<10 Then FormatDatePart="0"&FormatDatePart
	Else
		FormatDatePart=DatePart("n",vdate)
		If FormatDatePart<10 Then FormatDatePart="0"&FormatDatePart
	End If
End Function
%>
<script type="text/javascript">
<!--
	function TnViewDivSelect(viewdiv){
		if(viewdiv==1){
			$("#datearea").css("display","none");
		}else{
			$("#datearea").css("display","");
		}
	}

	function TnSearchObjOpenWin(){
		var winpop = window.open('/admin/itemmaster/deal/pop_deal_addItems.asp?idx=<%=idx%>&stype=w','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
	}

	function SubmitSave(frm){
		if(frm.itemname.value=="")
		{
			alert("상품명을 입력해주세요.");
			frm.itemname.focus();
			return false;
		}
		else if(frm.itemname.value.length>50)
		{
			alert("50자 이내로 상품명을 입력해주세요.");
			frm.itemname.focus();
			return false;
		}
		else if(!frm.viewdiv[0].checked && !frm.viewdiv[1].checked)
		{
			alert("노출 기간을 선택해주세요.");
			return false;
		}
		else if(frm.viewdiv[1].checked && (frm.startdate.value=="" || frm.enddate.value==""))
		{
			alert("노출 기간을 설정해주세요.");
			return false;
		}
		else if(!frm.isusing[0].checked && !frm.isusing[1].checked)
		{
			alert("사용 여부를 선택해주세요.");
			return false;
		}
		else if(!frm.sellyn[0].checked && !frm.sellyn[1].checked && !frm.sellyn[2].checked)
		{
			alert("판매 여부를 선택해주세요.");
			return false;
		}
		else if(frm.itemid.value=="" && frm.isusing[0].checked)
		{
			alert("대표상품을 선택해주세요.");
			frm.itemid.focus();
			return false;
		}
		else if(frm.mastersellcash.value=="")
		{
			alert("대표 가격을 입력해주세요.");
			frm.mastersellcash.focus();
			return false;
		}
		else if(frm.masterdiscountrate.value=="")
		{
			alert("대표 할인율을 입력해주세요.");
			frm.masterdiscountrate.focus();
			return false;
		}
		else if($("#tbl_DispCate tr").length<1)
		{
			alert("전시 카테고리를 추가해주세요.");
			frm.catecode.focus();
			return false;
		}
		else
		{
			if(confirm("입력하신 정보로 딜상품을 수정하시겠습니까?"))
			{
				frm.target="FrameCKP";
				frm.action="<%= ItemUploadUrl %>/linkweb/items/deal_itemeditWithImage_process.asp";
				frm.submit();
			}
		}		
	}

	function ClearImage2(img,fsize,wd,ht) {
		img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +");\" class='text' size='"+ fsize +"'>";
		$("#divaddimgname").hide();
		document.frm.addimgdel.value = "del";
	}

	function TnMasterItemSelect(itemid){
		if(document.frm.itemname.value=="")
		{
			document.frm.itemname.value=$("#itemcode option:selected").text();
		}
		$.ajax({
			url: "selectdealitemkeywords.asp?itemid="+itemid,
			cache: false,
			async: false,
			success: function(message) {
				//alert(message);
				if(message!="") {
					$('#keywords').val(message);
				} else {
					alert("제공 할 정보가 없습니다.");
				}
			}
		});
	}

	// 기본정보 수정
	function editItemBasicInfo(itemid) {
		var param = "itemid=" + itemid + "&menupos=<%= menupos %>";
		popwin = window.open('/admin/itemmaster/pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=1100,height=600,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	// 기본정보 수정
	function fnSaleInfo() {
		popwin = window.open('/admin/shopmaster/sale/saleList.asp?menupos=290' ,'saleinfo','width=1100,height=600,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function onlyNumerSet(text){
		if(window.event.keyCode < 48 || window.event.keyCode > 57) {
			return false;
		}
	}

	function fnPaste() {
		var regex = /\D/ig;
		if (regex.test(window.clipboardData.getData("text"))) {
			return false;
		} else {
			return true;
		}
	}
	function fnCancel(){
		if(confirm("입력하신 정보를 저장하지 않고 취소하시겠습니까?")){
			location.href="/admin/itemmaster/deal/index.asp";
		}
	}
	function editItemImage(itemid) {
		var param = "itemid=" + itemid;

		//if(makerid =="ithinkso"){
			//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
		//}else{
			popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=1000,height=900,scrollbars=yes,resizable=yes');
		//}
		popwin.focus();
	}	
//-->
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<form name="frm" method="post" onsubmit="return false;" style="margin:0px;"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="realitemid" value="<%=oDeal.Fdealitemid%>">
<input type="hidden" name="masteritemid" value="<%=oDeal.Fmasteritemcode%>">
<input type="hidden" name="auser" value="<%=session("ssBctId")%>">
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>딜 기본 정보</B></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">상품명<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="itemname" id="itemname" size="80" maxlength="120" value="<%=oitem.FOneItem.Fitemname%>" class="text">
		</td>
	</tr>
	<tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">노출 기간<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			 <input type="radio" name="viewdiv" id="viewdiv" value="1" onClick="TnViewDivSelect(1)"<% If oDeal.Fviewdiv="1" Then Response.write " checked" %>>상시딜 <input type="radio" name="viewdiv" id="viewdiv" value="2" onClick="TnViewDivSelect(2)"<% If oDeal.Fviewdiv="2" Then Response.write " checked" %>>기간딜
			 <span id="datearea" style="display:<% If oDeal.Fviewdiv<>"2" Then Response.write "none" %>">
				<input id="startdate" name="startdate" value="<%=FormatDate(oDeal.Fstartdate,"0000-00-00")%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> <input type="text" name="shour" size="2" class="text" value="<% =FormatDatePart("h",oDeal.Fstartdate)%>"  maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">:<input type="text" name="sminute" size="2" class="text" value="<%= FormatDatePart("n",oDeal.Fstartdate)%>"  maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');"> ~
				<input id="enddate" name="enddate" value="<%=FormatDate(oDeal.Fenddate,"0000-00-00")%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" /> <input type="text" name="ehour" size="2" class="text"  value="<%=FormatDatePart("h",oDeal.Fenddate)%>"  maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">:<input type="text" name="eminute" size="2" class="text" value="<%=FormatDatePart("n",oDeal.Fenddate)%>"  maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "startdate", trigger    : "iSD_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "enddate", trigger    : "iED_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			 </span>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">사용여부<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" id="isusing" value="Y"<% If oDeal.Fisusing="Y" Then Response.write " checked" %>>사용 <input type="radio" name="isusing" id="isusing" value="N"<% If oDeal.Fisusing="N" Then Response.write " checked" %>>사용 안함
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">판매여부<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="sellyn" id="sellyn" value="Y"<% If oDeal.Fsellyn="Y" Then Response.write " checked" %>>사용 <input type="radio" name="sellyn" id="sellyn" value="S"<% If oDeal.Fsellyn="S" Then Response.write " checked" %>>일시 품절 <input type="radio" name="sellyn" id="sellyn" value="N"<% If oDeal.Fsellyn="N" Then Response.write " checked" %>>판매 안함
		</td>
	</tr>	
</table>
<p>&nbsp;</p>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>노출 상품 정보</B></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100">상품목록<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="button" class="button" style="width:105;" value="검색" onclick="TnSearchObjOpenWin('Just1Day_list.asp');">&nbsp;<b style="color:red">*</b>딜 상품을 검색하여 추가해 주세요.
			<div id="divForm">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
			<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
				<td>순서</td>
				<td>상품코드</td>
				<td>상품명</td>
				<td>판매가</td>
				<td>매입가</td>
				<td>할인율</td>
			</tr>
			<% If isArray(arrList) Then %>
			<% For intLoop = 0 To UBound(arrList,2) %>
			<tr bgcolor="#FFFFFF" align="center">
				<td><%=arrList(0,intLoop)%></td>
				<td><a href="javascript:editItemBasicInfo('<%=arrList(1,intLoop)%>')"><%=arrList(1,intLoop)%></a></td>
				<td><%=arrList(2,intLoop)%></td>
				<td>
					<%
						Response.Write FormatNumber(arrList(5,intLoop),0)
						'할인가
						if arrList(9,intLoop)="Y" then
							Response.Write "<br><font color=#F08050>(할)" & FormatNumber(arrList(7,intLoop),0) & "</font>"
						end if
						'쿠폰가
						if arrList(10,intLoop)="Y" then
							Select Case arrList(11,intLoop)
								Case "1"
									Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(4,intLoop)*((100-arrList(12,intLoop))/100),0) & "</font>"
								Case "2"
									Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(4,intLoop)-arrList(12,intLoop),0) & "</font>"
							end Select
						end if
					%>
				</td>
				<td>
					<%
						Response.Write FormatNumber(arrList(6,intLoop),0)
						'할인가
						if arrList(9,intLoop)="Y" then
							Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
						end if
						'쿠폰가
						if arrList(10,intLoop)="Y" then
							if arrList(12,intLoop)="1" or arrList(12,intLoop)="2" then
								if arrList(13,intLoop)=0 or isNull(arrList(13,intLoop)) then
									Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(6,intLoop),0) & "</font>"
								else
									Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(13,intLoop),0) & "</font>"
								end if
							end if
						end if
					%>
				</td>
				<td>
					<a href="javascript:fnSaleInfo();"><%if arrList(9,intLoop)="Y" then%>
					<font color="#F08050"><%=CLng(((arrList(5,intLoop)-arrList(7,intLoop))/arrList(5,intLoop))*100)%>%</font>		
					<%end if%>
					<%if arrList(10,intLoop)="Y" then 
					if arrList(12,intLoop)="1" or arrList(12,intLoop)="2" then
						if arrList(13,intLoop)=0 or isNull(arrList(13,intLoop)) then
							 Response.Write "<br><font color=#5080F0>" & FormatNumber( arrList(6,intLoop),0) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(12,intLoop),0) 
							 if arrList(12,intLoop)="1" then 
							 Response.Write "%"
							else
							 Response.Write "원"
							end if
							 Response.Write "</font>"
						end if
					end if
					end if%></a>
				</td>
			</tr>
			<% Next %>
			<% End If %>
			</table>
			</div>
			<div id="divFrm3"></div>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표 상품<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<% If isArray(arrList) Then %>
			<select name="itemid" id="itemcode" onChange="TnMasterItemSelect(this.value);">
				<option value="" selected>상품을 선택해 주세요.</option>
				<% For intLoop = 0 To UBound(arrList,2) %>
				<option value="<%=arrList(1,intLoop)%>"<% If arrList(1,intLoop) = oDeal.Fmasteritemcode Then  Response.write " selected"%>><%=arrList(2,intLoop)%></option>
				<% Next %>
			</select>
			<% Else %>
			<select name="itemid"  id="itemcode" disabled  onChange="TnMasterItemSelect();">
				<option value="" selected>상품을 추가해 주세요.</option>
			</select>
			<% End If %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">딜상품 이미지관리</td>
		<td bgcolor="#FFFFFF">
			<input type="button" value="이미지관리" class="button" onClick="editItemImage('<%= itemid %>')">
		</td>
	</tr>	
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표 가격, 할인<br>아이템 코드</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="saleitemid" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">
			<input type="text" name="discountitemid" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">
			(수기로 아이템 코드를 입력해도 가격,할인정보를 수정 할 수 있습니다.)
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표 가격<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="mastersellcash" size="10" class="text" value="<%=oDeal.Fmastersellcash%>" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">&nbsp;<input type="button" value="가져오기" class="button" onClick="fnGetMinPricevalue()" id="saleper1" name="saleper1" style="display:<% If Not isArray(arrList) Then %>none<% End If %>"><!-- &nbsp;<input type="checkbox" name="pricesdash" value="Y"<% If oDeal.Fpricesdash ="Y" Then Response.write " checked" %>>"~"노출 선택 (예: 19,900원~) -->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표 할인율<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="masterdiscountrate" id="masterdiscountrate" value="<%=oDeal.Fmasterdiscountrate%>" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">&nbsp;<input type="button" value="가져오기" class="button" onClick="fnGetMaxSalevalue()" id="saleper2" name="saleper2" style="display:<% If Not isArray(arrList) Then %>none<% End If %>"><!-- &nbsp;<input type="checkbox" name="sailsdash" value="Y"<% If oDeal.Fsailsdash ="Y" Then Response.write " checked" %>>"~"노출 선택 (예: ~77%) -->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">전시 카테고리<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<table class="a">
			<tr>
				<td id="lyrDispList"><%=getDispCategory(oDeal.Fdealitemid)%></td>
				<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
			</tr>
			</table>
			<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
		</td>
	</tr>
	<tr align="left">
	<td height="30" width="15%" bgcolor="<%= adminColor("tabletop") %>">구매 가능 연령 </td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="adultType" value="0" <%=chkIIF(oitem.FOneItem.FadultType=0,"checked","")%>>전체연령</label>
		<label><input type="radio" name="adultType" value="1" <%=chkIIF(oitem.FOneItem.FadultType=1,"checked","")%>>미성년 조회 불가</label>
		<label><input type="radio" name="adultType" value="2" <%=chkIIF(oitem.FOneItem.FadultType=2,"checked","")%>>구매시성인인증</label>
	</td>
</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">검색 키워드<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="keywords" id="keywords" size="80" maxlength="250" value="<%=oitem.FOneItem.Fkeywords%>" class="text"> (콤마로구분 ex: 커플,티셔츠,조명)
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">요약 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40);" class="text" size="40">
			<input type="button" value="#1 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname,40, 800, 1600)"> (선택,800X1600, Max 800KB,jpg,gif)<br>
			<input type="hidden" name="addimggubun" value="1">
			<input type="hidden" name="addimgdel" value="">
			<%
			If isArray(vArr) Then
					If vArr(3,UBound(vArr,2)) > 0 Then
					For k = 1 To vArr(3,UBound(vArr,2))
		    			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
							Response.Write "<div id=""divaddimgname"" style=""display:block;""><img src=""" & webImgUrl & "/item/contentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """ height=""250""></div>"
							Exit For
		    			End If
					Next
					End If
			End If
			%>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">비고</td>
		<td bgcolor="#FFFFFF">
			<textarea name="itemcontent" rows="18" class="textarea" style="width:99%" id="[on,off,off,off][상품설명]"><%=oDeal.Fwork_notice%></textarea>
		</td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" height="25">
        <td valign="bottom" align="center">
			<input type="button" value="수정" class="button" onClick="SubmitSave(this.form)">
			<input type="button" value="취소" class="button" onClick="fnCancel()">
        </td>
    </tr>
</table>
</form>
<% if (application("Svr_Info")	= "Dev") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>
<script type="text/javascript">
<!--
	//상품 최대 할인율 접수
	function fnGetMaxSalevalue() {
		var idx = document.frm.idx.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetDealMaxItemSalePer.asp",
			data: "idx="+idx,
			cache: false,
			success: function(message) {
				var splitmessage = message.split("|")
				if(message) {
					document.frm.masterdiscountrate.value=splitmessage[0];
					document.frm.discountitemid.value=splitmessage[1];
				} else {
					alert("상품이 없거나 할인중인 상품이 없습니다.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	//상품 최저가 접수
	function fnGetMinPricevalue() {
		var idx = document.frm.idx.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetDealMinItemPrice.asp",
			data: "idx="+idx,
			cache: false,
			success: function(message) {
				var splitmessage = message.split("|")
				if(message) {
					document.frm.mastersellcash.value=splitmessage[0];
					document.frm.saleitemid.value=splitmessage[1];
				} else {
					alert("상품이 없습니다.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// 전시카테고리 선택 팝업
	function popDispCateSelect(){
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// 레이어에서 전시카테고리 추가
	function addDispCateItem(dcd,cnm,div,dpt) {
		// 기존에 값에 중복 카테고리 여부 검사
		if(tbl_DispCate.rows.length>=2)	{
			alert("전시 카테고리는 최대 2개까지 입력가능합니다.");
			return false;
		}
		else
		{
			if(tbl_DispCate.rows.length>0)	{
				if(tbl_DispCate.rows.length>1)	{
					for(l=0;l<document.all.isDefault.length;l++)	{
						if((document.all.catecode[l].value==dcd)) {
							alert("이미 지정된 같은 카테고리가 있습니다..");
							return;
						}
					}
				}
				else {
					if((document.all.catecode.value==dcd)) {
						alert("이미 지정된 같은 카테고리가 있습니다..");
						return;
					}
				}
			}
		}
		
		// 행추가
		var oRow = tbl_DispCate.insertRow();
		oRow.onmouseover=function(){tbl_DispCate.clickedRowIndex=this.rowIndex};

		// 셀추가 (구분,카테고리,삭제버튼)
		var oCell1 = oRow.insertCell();		
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="y") {
			oCell1.innerHTML = "<font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'>";
		}
		$(cnm).each(function(i){
			if(dpt>i) {
				if(i>0) oCell2.innerHTML += " >> ";
				oCell2.innerHTML += $(this).text();
			}
		});
		oCell2.innerHTML += "<input type='hidden' name='catecode' value='" + dcd + "'>";
		oCell2.innerHTML += "<input type='hidden' name='catedepth' value='" + dpt + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle>";
		$("#lyrDispCateAdd").fadeOut();

		//상품속성 출력
		printItemAttribute();
	}

	// 선택 전시카테고리 삭제
	function delDispCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//상품속성 출력
			printItemAttribute();
		}
	}

	function printItemAttribute() {
		var arrDispCd="";
		$("input[name='catecode']").each(function(i){
			if(i>0) arrDispCd += ",";
			arrDispCd += $(this).val();
		});
		$.ajax({
			url: "/common/module/act_ItemAttribSelect.asp?itemid=0&arrDispCate="+arrDispCd,
			cache: false,
			success: function(message) {
				$("#lyrItemAttribAdd").empty().append(message);
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	function CheckImage(img, filesize, imagewidth, imageheight, extname, fsize)
	{
		var ext;
		var filename;

		filename = img.value;
		if (img.value == "") { return false; }

		if (CheckExtension(filename, extname) != true) {
			alert("이미지화일은 다음의 화일만 사용하세요.[" + extname + "]");
			ClearImage(img,fsize,imagewidth,imageheight);
			return false;
		}

		return true;
	}
//-->
</script>
<%
Set oitem = Nothing
Set oDeal = Nothing
Set oitemimg = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->