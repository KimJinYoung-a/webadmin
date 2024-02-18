<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : 전시카테고리 상품
' History	:  서동석 생성
'              2021.07.12 한용민 수정(엑셀다운로드 추가)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<BODY LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->

<%
	Dim cDisp, i, vDepth, vCateCode, vCurrpage, vPageSize, vIsThisCate, vParam, vSearch, vNotCateReg, dispCate, vOnlyBasic
	vCurrPage	= NullFillWith(Request("cpg"), "1")
	vDepth 		= NullFillWith(Request("depth_s"), "1")
	vCateCode 	= Request("catecode_s")
	vIsThisCate	= Request("isthiscate")
	vPageSize	= NullFillWith(Request("pagesize"), 20)
	vSearch		= Request("search")
	vNotCateReg	= Request("notcatereg")
	vOnlyBasic	= request("onlybasic")
	dispCate	= Request("disp")

	Dim makerid, cdl, cdm, cds, itemid_s, itemname, keyword, sellyn, usingyn, danjongyn, limityn, sailyn, deliverytype, sortDiv, mustCate
	makerid		= request("makerid")
	cdl 		= request("cdl")
	cdm 		= request("cdm")
	cds 		= request("cds")
	itemid_s	= requestCheckvar(request("itemid_s"),1500)
	itemname	= request("itemname")
	keyword		= request("keyword")
	sellyn      = request("sellyn")
	usingyn     = request("usingyn")
	danjongyn   = request("danjongyn")
	limityn     = request("limityn")
	sailyn      = request("sailyn")
	deliverytype = request("deliverytype")
	sortDiv		= request("sortDiv")
	mustCate	= request("mustCate")

	if sortDiv = "" then sortDiv = "new"

	if itemid_s<>"" then
	dim iA ,arrTemp,arrItemid
	itemid_s = replace(itemid_s,",",chr(10))
	itemid_s = replace(itemid_s,chr(13),"")
	arrTemp = Split(itemid_s,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if trim(arrTemp(iA))<>"" then
			'상품코드 유효성 검사(2008.08.05;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	itemid_s = left(arrItemid,len(arrItemid)-1)
end if

	SET cDisp = New cDispCate
	cDisp.FCurrPage = vCurrpage
	cDisp.FPageSize = vPageSize
	cDisp.FRectDepth = vDepth
	If vIsThisCate <> "" Then
		cDisp.FRectCateCode = vCateCode
	End IF
	cDisp.FRectMakerId 		= makerid
	cDisp.FRectItemID 		= itemid_s
	cDisp.FRectCDL 			= cdl
	cDisp.FRectCDM 			= cdm
	cDisp.FRectCDS 			= cds
	cDisp.FRectItemName 	= itemname
	cDisp.FRectKeyword 		= keyword
	cDisp.FRectSellYN		= sellyn
	cDisp.FRectIsUsing		= usingyn
	cDisp.FRectDanjongyn	= danjongyn
	cDisp.FRectLimityn		= limityn
	cDisp.FRectSailYn		= sailyn
	cDisp.FRectDeliveryType	= deliverytype
	cDisp.FRectSortDiv = SortDiv
	cDisp.FRectNotCateReg	= vNotCateReg
	cDisp.FRectOnlyBasic	= vOnlyBasic
	cDisp.FSearchDispCate	= dispCate
	cDisp.FRectMustCate		= mustCate
	cDisp.GetDispCateItemList()

%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
function searchFrm(p){
	$('input[name="cpg"]').val(p);

	if($('#notcatereg').prop('checked') == true) {
		parent.$('input[name="notcatereg"]').val('o');
	}else{
		parent.$('input[name="notcatereg"]').val('');
	}
	parent.$('input[name="makerid"]').val(frmitem.makerid.value);
	parent.$('input[name="cdl"]').val(frmitem.cdl.value);
	parent.$('input[name="cdm"]').val(frmitem.cdm.value);
	parent.$('input[name="cds"]').val(frmitem.cds.value);
	parent.$('input[name="itemid_s"]').val(frmitem.itemid_s.value);
	parent.$('input[name="itemname"]').val(frmitem.itemname.value);
	parent.$('input[name="keyword"]').val(frmitem.keyword.value);
	parent.$('input[name="sellyn"]').val(frmitem.sellyn.value);
	parent.$('input[name="usingyn"]').val(frmitem.usingyn.value);
	parent.$('input[name="danjongyn"]').val(frmitem.danjongyn.value);
	parent.$('input[name="limityn"]').val(frmitem.limityn.value);
	parent.$('input[name="sailyn"]').val(frmitem.sailyn.value);
	parent.$('input[name="deliverytype"]').val(frmitem.deliverytype.value);
	parent.$('input[name="sortDiv"]').val(frmitem.sortDiv.value);
	parent.$('input[name="pagesize"]').val(frmitem.pagesize.value);

	frmitem.action="";
	frmitem.target="";
	frmitem.submit();
}

function download_excel(){
	$('input[name="cpg"]').val(1);

	if($('#notcatereg').prop('checked') == true) {
		parent.$('input[name="notcatereg"]').val('o');
	}else{
		parent.$('input[name="notcatereg"]').val('');
	}
	parent.$('input[name="makerid"]').val(frmitem.makerid.value);
	parent.$('input[name="cdl"]').val(frmitem.cdl.value);
	parent.$('input[name="cdm"]').val(frmitem.cdm.value);
	parent.$('input[name="cds"]').val(frmitem.cds.value);
	parent.$('input[name="itemid_s"]').val(frmitem.itemid_s.value);
	parent.$('input[name="itemname"]').val(frmitem.itemname.value);
	parent.$('input[name="keyword"]').val(frmitem.keyword.value);
	parent.$('input[name="sellyn"]').val(frmitem.sellyn.value);
	parent.$('input[name="usingyn"]').val(frmitem.usingyn.value);
	parent.$('input[name="danjongyn"]').val(frmitem.danjongyn.value);
	parent.$('input[name="limityn"]').val(frmitem.limityn.value);
	parent.$('input[name="sailyn"]').val(frmitem.sailyn.value);
	parent.$('input[name="deliverytype"]').val(frmitem.deliverytype.value);
	parent.$('input[name="sortDiv"]').val(frmitem.sortDiv.value);
	parent.$('input[name="pagesize"]').val(frmitem.pagesize.value);

	frmitem.action="/admin/CategoryMaster/displaycate/display_cate_item_excel.asp";
	frmitem.target="exceldown";
	frmitem.submit();
	frmitem.target = "";
	frmitem.action = "";
}

function jsRegItem(itemid,spanid){
	$.ajax({
			url: "display_cate_item_proc.asp?itemid="+itemid+"&catecode=<%=vCateCode%>&depth=<%=vDepth-1%>",
			cache: false,
			success: function(message)
			{
				$("#"+spanid+"").empty().append(message);
			}
	});
}
function jsEditItem(itemid,catecode){
	$.ajax({
			url: "display_cate_item_ajax.asp?itemid="+itemid+"&catecode="+catecode+"&depth=<%=vDepth-1%>",
			cache: false,
			success: function(message)
			{
				$("#editarea").empty().append(message);
				$("#editarea").show();
				parent.jsEditLink();
			}
	});
}
function Check_All()
{
	var chk = document.frmitem.itemid;
	var cnt = 0;
	var ischecked = ""
	if(document.getElementById("chkall").checked){
		ischecked = "checked"
	}else{
		ischecked = ""
	}
	if(cnt == 0 && chk.length != 0){
		for(i = 0; i < chk.length; i++){ chk.item(i).checked = ischecked; }
		cnt++;
	}
}
function jsCheckAllReg(){
	var i = "";
	$("input:checkbox[name='itemid']").each(
		function(){
			if (this.checked)
			{
				i = i + this.value + ",";
			}
		}
	)

	if(i == ""){
		alert("선택된 상품이 없습니다.");
		return;
	}else{
		$("#DivLoadingBar").show();
		$('input[name="allitemid"]').val(i);
		document.frmallitem.submit();
	}
}

function delCateItem(){
	var i = "";
	$("input:checkbox[name='itemid']").each(
		function(){
			if (this.checked)
			{
				i = i + this.value + ",";
			}
		}
	)

	if(i == ""){
		alert("선택된 상품이 없습니다.");
		return;
	}else{
		if(confirm("선택하신 상품들을 삭제하시겠습니까?\n\n※ [필독]부분을 반드시 확인하세요.") == true) {
			$("#DivLoadingBar").show();
			$('input[name="allitemid"]').val(i);
			$('input[name="action"]').val('delete');
			document.frmallitem.submit();
		}else{
			return;
		}
	}
}
</script>

<input type="text" id="nowcatename" name="nowcatename" value="" size="150" style="border:solid 1px #ffffff;height:25px;padding-top:5px;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a" bgcolor="#FFFFFF">
<tr>
	<td width="75%">
		<form name="frmitem" method="get" action="<%=CurrURL()%>" style="margin:0px;">
		<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
		<input type="hidden" name="search" value="o">
		<input type="hidden" name="cpg" value="1">
		<input type="hidden" name="catecode_s" value="<%=vCateCode%>">
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr>
			<td bgcolor="#FFFFFF">
				<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<tr>
					<td bgcolor="#FFFFFF">
						<table class="a">
							<tr>
								<td>브랜드 : <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
								<td><span style="padding-left:10px">상품코드 :</span> </td>
								<td rowspan="2" bgcolor="#FFFFFF">
									<textarea rows="3" cols="10" name="itemid_s" id="itemid_s"><%=replace(itemid_s,",",chr(10))%></textarea>
									<span style="font-size:11px; color:gray;padding-left:5px;">(엔터로 복수입력가능)</span>
								</td>
							</tr>
							<tr>
								<td>
									상품명 :
									<input type="text" class="text" name="itemname" value="<%= itemname %>" size="25" maxlength="20">
								</td>
								</tr>
							</table>
					</td>
				</tr>
					<tr>
						<td  bgcolor="#FFFFFF">
						검색키워드 : <input type="text" class="text" name="keyword" value="<%=keyword%>" size="40">
						<span style="font-size:11px; color:gray;padding-left:5px;">(주의:느릴수있습니다.)</span>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF">
						<!-- #include virtual="/admin/CategoryMaster/displaycate/categoryselectbox.asp"-->
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF">
						전시카테고리 :
						<script type="text/javascript">
						$(function(){
							chgDispCate('<%=dispCate%>');
						});

						function chgDispCate(dc) {
							$.ajax({
								url: "/common/module/dispCateSelectBox_response.asp?disp="+dc,
								cache: false,
								async: false,
								success: function(message) {
						       		// 내용 넣기
						       		$("#lyrDispCtBox").empty().html(message);
						       		$("#oDispCate").val(dc);
								}
							});
						}
						</script>
						<span id="lyrDispCtBox"></span>
						<input type="hidden" name="disp" id="oDispCate" value="<%=dispCate%>">
						<label><input type = "checkbox" name="mustCate" value="Y" <%= Chkiif(mustCate="Y", "checked", "") %>  >해당 전시카테고리만</label>
					</td>
				</tr>

				<tr>
					<td bgcolor="#FFFFFF"   colspan="3">
					판매:<% drawSelectBoxSellYN "sellyn", sellyn %>&nbsp;
					사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>&nbsp;
					단종:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>&nbsp;
					한정:<% drawSelectBoxLimitYN "limityn", limityn %>&nbsp;
					할인 <% drawSelectBoxSailYN "sailyn", sailyn %>&nbsp;
					배송:<% drawBeadalDiv "deliverytype",deliverytype %>
					</td>
				</tr>
				<tr>
					<td bgcolor="#D4FFFF"   colspan="3">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
						<tr>
							<td>
								<strong>
								Total : <%=FormatNumber(cDisp.FTotalCount,0)%>&nbsp;&nbsp;&nbsp;
								<select name="IsThisCate" class="select" onChange="searchFrm('1');">
									<option value="" <%=CHKIIF(vIsThisCate="","selected","")%>>전체상품</option>
									<option value="o" <%=CHKIIF(vIsThisCate="o","selected","")%>>선택된카테고리상품만</option>
								</select>
								&nbsp;&nbsp;&nbsp;
								<select name="pagesize" class="select" onChange="searchFrm('1');">
									<option value="20" <%=CHKIIF(vPageSize="20","selected","")%>>20개씩보기</option>
									<option value="50" <%=CHKIIF(vPageSize="50","selected","")%>>50개씩보기</option>
									<option value="100" <%=CHKIIF(vPageSize="100","selected","")%>>100개씩보기</option>
									<option value="150" <%=CHKIIF(vPageSize="150","selected","")%>>150개씩보기</option>
									<option value="500" <%=CHKIIF(vPageSize="500","selected","")%>>500개씩보기</option>
								</select>
								&nbsp;&nbsp;&nbsp;
								<label><input type="checkbox" name="notcatereg" id="notcatereg" value="o" onClick="searchFrm('1');" <%=CHKIIF(vNotCateReg="o","checked","")%>>지정안된 카테고리만</label>
								<br>
								<label>
								<input type="radio" name="onlybasic" id="onlybasic" value="" onClick="searchFrm('1');" <%=CHKIIF(vOnlyBasic="","checked","")%>>선택안함
								<input type="radio" name="onlybasic" id="onlybasic" value="o" onClick="searchFrm('1');" <%=CHKIIF(vOnlyBasic="o","checked","")%>>기본 카테고리만
								<input type="radio" name="onlybasic" id="onlybasic" value="N" onClick="searchFrm('1');" <%=CHKIIF(vOnlyBasic="N","checked","")%>>기본 카테고리가 없는
								</label>
								</strong>
							</td>
							<td align="right">
								<strong>
								정렬 :
								<select name="sortDiv" class="select" onchange="searchFrm('1');">
									<option value="new" <% IF sortDiv="new" Then response.write "selected" %> >신상품순</option>
									<option value="cashH" <% IF sortDiv="cashH" Then response.write "selected" %>>높은가격순</option>
									<option value="cashL" <% IF sortDiv="cashL" Then response.write "selected" %>>낮은가격순</option>
									<option value="best" <% IF sortDiv="best" Then response.write "selected" %>>베스트순</option>
								</select>
								</strong>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
			<td bgcolor="#FFFFFF" width="10%" align="center">
				<table class="a">
				<tr>
					<td align="center"><input type="button" value="검 색" onClick="searchFrm('1');" style="width:60px;height:60px;"></td>
				</tr>
				<tr>
					<td align="center" style="padding-top:15px;"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" bgcolor="#FFFFFF">
				<input type="button" class="button_s" value="엑셀다운로드(10만개제한)" onclick="download_excel();">
				<input type="button" value="선택한 상품에 등록된 카테고리 모두 삭제" onClick="delCateItem()">
				<br><b>[필독] -> 삭제하면 복구 절대 안됩니다.
				<br>카테고리리스트, 브랜드상품리스트 등등 검색엔진이 사용되는 곳엔 적용되는데 시간이 많이 걸림.
			</td>
		</tr>
		</table>
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<% If vCateCode <> "" Then %>
		<tr>
			<td colspan="7" bgcolor="#FFFFFF" height="30" align="right">
				<input type="button" value="선택한것 모두 등록" onClick="jsCheckAllReg()">
			</td>
		</tr>
		<% End If %>
		<tr align="center" bgcolor="#F3F3FF" height="30">
			<td width="4%"><input type="checkbox" name="chkall" id="chkall" value="" onClick="Check_All()"></td>
			<td width="6%"></td>
			<td width="10%">Maker ID</td>
			<td>상품코드</td>
			<td>상품명</td>
			<td width="35%">지정된카테고리</td>
			<% If vCateCode <> "" Then %>
				<td width="15%">카테고리지정</td>
			<% End If %>
		</tr>
		<%
		If cDisp.FResultCount = 0 Then
		%>
			<tr>
				<td colspan="6" height="30" bgcolor="#FFFFFF" align="center">검색된 상품이 없습니다.</td>
			</tr>
		<%
		Else
			For i=0 To cDisp.FResultCount-1
		%>
			<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
				<td align="center"><input type="checkbox" name="itemid" value="<%=cDisp.FItemList(i).FItemID%>"></td>
				<td align="center"><img src="<%=cDisp.FItemList(i).FSmallImage%>"></td>
				<td align="center"><%=cDisp.FItemList(i).FMakerID%></td>
				<td  align="center"><%=cDisp.FItemList(i).FItemID%></td>
				<td><%=cDisp.FItemList(i).FItemName%></td>
				<td><span id="catenamespan<%=cDisp.FItemList(i).FItemID%>" style="font-size:0.9em">
					<%=fnCateCodeNameSplit(cDisp.FItemList(i).FCateName,cDisp.FItemList(i).FItemID)%></span></td>
				<% If vCateCode <> "" Then %>
					<td align="center" style="cursor:pointer" onClick="jsRegItem('<%=cDisp.FItemList(i).FItemID%>','catenamespan<%=cDisp.FItemList(i).FItemID%>');"><font color="blue" size="2"><b>등록하기</b></font></td>
					<!--<td align="center" style="cursor:pointer" onClick="jsEditItem('<%=cDisp.FItemList(i).FItemID%>','');"><font color="green" size="2"><b>수정,삭제</b></font></td>//-->
				<% End If %>
			</tr>
		<%
			Next
		%>
			<tr height="50" bgcolor="FFFFFF">
				<td colspan="20" align="center">
					<% if cDisp.HasPreScroll then %>
					<a href="javascript:searchFrm('<%= cDisp.StartScrollPage-1 %>')">[pre]</a>
		    		<% else %>
		    			[pre]
		    		<% end if %>

		    		<% for i=0 + cDisp.StartScrollPage to cDisp.FScrollCount + cDisp.StartScrollPage - 1 %>
		    			<% if i>cDisp.FTotalpage then Exit for %>
		    			<% if CStr(vCurrpage)=CStr(i) then %>
		    			<font color="red">[<%= i %>]</font>
		    			<% else %>
		    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
		    			<% end if %>
		    		<% next %>

		    		<% if cDisp.HasNextScroll then %>
		    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
		    		<% else %>
		    			[next]
		    		<% end if %>
				</td>
			</tr>
		<%
		End If
		%>
		</table>
		</form>
	</td>
	<td width="25%" style="padding:0 0 0 10px;vertical-align:top;">
		<div id="editarea" style="display:none;">
		</div>
	</td>
</tr>
</table>

<% SET cDisp = Nothing %>
<script>
$("#nowcatename").val(parent.$("#nowcatename").val());
<% If vSearch = "o" Then %>
	parent.jsEditLink();
<% End If %>
</script>
<form name="frmallitem" method="post" action="display_cate_item_allproc.asp" target="cateitemproc">
<input type="hidden" name="action" value="">
<input type="hidden" name="allitemid" value="">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="depth" value="<%=vDepth-1%>">
</form>
<iframe src="" id="cateitemproc" name="cateitemproc" width="0" height="0" frameborder="0"></iframe>
<!-- Loading Message Layer Start -->
<div id="DivLoadingBar" style="position:absolute; left:0px; top:0px; height:100%; width:100%; background-color:#FFFFFF; display:none;">
<table width=100% height=100% align=center border=0 cellpadding=0 cellspacing=0>
<tr>
	<td align="center" valign="top" style="padding-top:200px;">
		<table width=300 border=0 cellpadding=3 cellspacing=1 bgcolor="#CCCCCC">
		<tr>
			<td align=center bgcolor=#FFFFFF>
				<table width=100% border=0 cellpadding=0  cellspacing=1 bgcolor="#CCCCCC">
				<tr height=90>
					<td align=center style="color:#5F5F5F;font-family:vernada;font-size:9pt;font-weight:bold" bgcolor=#FFFFFF>
						저장 중입니다.<BR>잠시만 기다려주세요.
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</div>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="300"></iframe>
<% else %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="0"></iframe>
<% end if %>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->