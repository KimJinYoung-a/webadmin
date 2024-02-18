<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/jikbang_just1DayCls.asp"-->
<%
'###############################################
' PageName : Just1Day_write.asp
' Discription : 저스트 원데이 등록/수정
' History : 2008.04.08 허진원 생성
'           2012.02.15 허진원 : 미니달력 교체
'           2014.09.12 원승현 : 직방 원데이용으로 아주 쪼오끔 수정
'###############################################

dim justDate,mode,i
mode=request("mode")
justDate=request("justDate")

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--

document.domain = "10x10.co.kr";

function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb,orgImgName){

	window.open("Just1Day_PopImgInput.asp?divName="+divnm+"&inputname="+iptNm+"&ImagePath="+vPath+"&maxFileSize="+Fsize+"&maxFileWidth="+Fwidth+"&makeThumbYn="+thumb+"&orgImgName="+orgImgName,'imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
}

function editcont(){	
    //오픈된후 설명만 수정할 경우;;
    var frm=document.inputfrm;
    
    if (confirm('수정 하시겠습니까?')){
        frm.sale_code.value="";
        frm.submit();
    }
    
}

function subcheck(){
	var frm=document.inputfrm;

	if(!frm.justDate.value) {
		alert("지정할 날짜를 선택해주세요!");
		return;
	} else {
		if(frm.justDate.value<='<%=date%>') {
			//alert("상품의 수정/등록은 오늘 이후의 날짜만 가능합니다.");
			//return;
		}
	}

	if(!frm.itemid.value) {
		alert("등록할 상품을 선택해주세요!");
		return;
	}

	if(!frm.salePrice.value) {
		alert("상품의 할인금액을 입력해주세요!");
		frm.salePrice.focus();
		return;
	} else {
		if(parseInt(frm.salePrice.value)>=parseInt(frm.orgPrice.value)) {
			alert("판매가보다 할인액이 크거나 같을 수는 없습니다.\n할인액을 확인해주세요.");
			return;
		}
	}

    // 할인액0,매입액0 입력가능하게 수정
	if ((!frm.saleSuplyCash.value||frm.saleSuplyCash.value=="0")&&(frm.salePrice.value!="0")) {
		alert("상품의 매입금액을 입력해주세요!\n※매입급액은 반드시 실제매입가를 입력해야됩니다.");
		frm.saleSuplyCash.focus();
		return;
	}
    
    // 매입가가 할인판매가 보다 클 수 없음
    if (frm.saleSuplyCash.value*1>frm.salePrice.value*1) {
		alert("상품의 매입금액을 입력해주세요!\n※매입급액이 판매 금액 보다 클 수 없습니다.");
		frm.saleSuplyCash.focus();
		return;
	}
	
	if(!frm.limitNo.value) {
		alert("한정으로 판매할 개수를 입력해주세요.\n\n※ 한정판매가 아니라면 0을 입력해주세요.");
		frm.limitNo.focus();
		return;
	}

	if(frm.justDesc.value.length<=0||frm.justDesc.value.length>=240) {
		alert("상품의 간략 설명을 240자이내(4줄 이내)로 작성해주세요.\n\n");
		frm.justDesc.focus();
		return;
	}
    
    //eastone 추가 판매가0,매입가0 할인등록 안함.
    if ((frm.salePrice.value=="0")&&(frm.saleSuplyCash.value=="0")){
        if (!confirm('할인판매가 0, 할인매입가 0으로 등록시 할인 되지 않습니다. 계속하시겠습니까?')){
            return;
        }
    }
    
	if(frm.mode.value=="add"&&frm.itemOptCnt.value>0&&frm.limitNo.value>0) {
		if(confirm("옵션이 존재하는 상품일 입니다.\n입력하신 한정개수가 옵션에 자동으로 반영되지 않으므로, 시작 전에 상품정보에서 옵션 한정정보를 직접 입력하셔야 정상적으로 판매가 가능합니다.")) {
			frm.submit();
		} else {
			return;
		}
	} else {
		frm.submit();
	}
}

function popItemWindow(tgf){
	var popup_item = window.open("/common/pop_singleItemSelect.asp?target=" + tgf + "&ptype=just1day", "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

function putPercent(){
	var pct, frm = document.inputfrm;
	if(frm.orgPrice.value==0||frm.salePrice.value==0) {
		pct = "0%";
	}
	else {
		pct = 1 - (frm.salePrice.value / frm.orgPrice.value);
		pct = pct * 100;
		pct = Math.round(pct*100) / 100
		pct = pct + "%";
	}
	frm.saleRate.value= pct;
}

function delitems(){
	var frm = document.inputfrm;
	if (confirm('본 아이템을 삭제하시겠습니까?\n\n※할인예약 정보도 함께 삭제됩니다.')) {
		frm.mode.value="delete";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="doJust1Day_Process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<input type="hidden" name="itemOptCnt" value="0">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>직방 저스트 원데이 등록/수정</b></font>
	</td>
</tr>
<% if mode="add" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">날짜</td>
	<td bgcolor="#FFFFFF">
		<input id="justDate" name="justDate" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="justDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "justDate", trigger    : "justDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상품</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="itemid" value="" size="10" readonly>
		<input type="button" class="button" value="찾기" onClick="popItemWindow('inputfrm')">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">할인정보</td>
	<td bgcolor="#FFFFFF">
		할인금액 <input type="text" class="text" name="salePrice" value="" size="10" style="text-align:right" onkeyup="putPercent()">원
		/ 판매가 <input type="text" class="text_ro" name="orgPrice" value="0" size="8" readonly style="text-align:right">원,
		할인율 <input type="text" class="text_ro" name="saleRate" value="0%" size="4" readonly style="text-align:center">
		<br>매입금액 <input type="text" class="text" name="saleSuplyCash" value="" size="8" style="text-align:right">원
		<br>(할인 안할경우 할인금액0, 매입금액 0 입력)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">한정개수</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="limitNo" value="0" size="4" style="text-align:right">
		<input type="hidden" name="limitYn" value="">
		(한정갯수 0으로 설정시 비한정 으로 판매됩니다.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">간략설명</td>
	<td bgcolor="#FFFFFF">
		<textarea name="justDesc" class="textarea" cols="80" rows="3"></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>배너 이미지1</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image1div','image1','i1','250','100','false','');"/>
		<input type="hidden" name="image1" value="">
		<div align="right" id="image1div"></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>배너 이미지2</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image2div','image2','i2','600','450','true','');"/>		
		<input type="hidden" name="image2" value="">
		<div align="right" id="image2div"></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>상세이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image3div','image3','i2','600','450','true','');"/>		
		<input type="hidden" name="image3" value="">
		<div align="right" id="image3div"></div>
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New Cjust1Day
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectDate=justDate
	fmainitem.Getjust1Daymodify
%>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">날짜</td>
	<td bgcolor="#FFFFFF">
		<b><%=fmainitem.FItemList(0).FjustDate%></b>
		<input type="hidden" name="justDate" value="<%=fmainitem.FItemList(0).FjustDate%>">
		<input type="hidden" name="sale_code" value="<%=fmainitem.FItemList(0).Fsale_code%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상품</td>
	<td bgcolor="#FFFFFF">
		<%= "[" & fmainitem.FItemList(0).Fitemid & "] " & fmainitem.FItemList(0).Fitemname %>
		<input type="hidden" name="itemid" value="<%=fmainitem.FItemList(0).Fitemid%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">할인정보</td>
	<td bgcolor="#FFFFFF">
		할인금액 <input type="text" class="text" name="salePrice" value="<%= fmainitem.FItemList(0).FjustSalePrice %>" size="10" style="text-align:right" onkeyup="putPercent()">원
		/ 판매가 <input type="text" class="text_ro" name="orgPrice" value="<%= fmainitem.FItemList(0).ForgPrice %>" size="8" readonly style="text-align:right">원,
		할인율 <input type="text" class="text_ro" name="saleRate" value="<%= FormatPercent(1-(fmainitem.FItemList(0).FjustSalePrice/fmainitem.FItemList(0).ForgPrice),2) %>" size="5" readonly style="text-align:center">
		<br>매입급액 <input type="text" class="text" name="saleSuplyCash" value="<%= fmainitem.FItemList(0).FsaleSuplyCash %>" size="8" style="text-align:right">원
		<br>(할인 안할경우 할인금액0, 매입금액 0 입력)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">한정개수</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="limitNo" value="<%= fmainitem.FItemList(0).FlimitNo %>" size="4" style="text-align:right">
		<input type="hidden" name="limitYn" value="">
		(한정갯수 0으로 설정시 비한정 으로 판매됩니다.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">간략설명</td>
	<td bgcolor="#FFFFFF">
		<textarea name="justDesc" class="textarea" cols="80" rows="3"><%= fmainitem.FItemList(0).FjustDesc %></textarea>
		<input type="button" value=" 설명 수정 " class="button" onclick="editcont();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>배너 이미지 1</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image1div','image1','i1','250','100','false','<%= fmainitem.FItemList(0).Fimg1 %>');"/>
		<input type="hidden" name="image1" value="<%= fmainitem.FItemList(0).Fimg1 %>">
		<div align="right" id="image1div"><% IF fmainitem.FItemList(0).Fimg1<>"" THEN %><img src="<%= fmainitem.FItemList(0).Fimg1 %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>배너 이미지 2</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image2div','image2','i2','600','450','true','<%= fmainitem.FItemList(0).Fimg2 %>');"/>
		<input type="hidden" name="image2" value="<%= fmainitem.FItemList(0).Fimg2 %>">
		<div align="right" id="image2div"><% IF fmainitem.FItemList(0).Fimg2<>"" THEN %><img src="<%= fmainitem.FItemList(0).Fimg2 %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>상세 이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image3div','image3','i2','600','450','true','<%= fmainitem.FItemList(0).Fimg3 %>');"/>
		<input type="hidden" name="image3" value="<%= fmainitem.FItemList(0).Fimg3 %>">
		<div align="right" id="image3div"><% IF fmainitem.FItemList(0).Fimg3<>"" THEN %><img src="<%= fmainitem.FItemList(0).Fimg3 %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><input type="button" value=" 삭제 " class="button" onclick="delitems();"> &nbsp;&nbsp;<% end if %>
		<input type="button" value=" 취소 " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>

<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
