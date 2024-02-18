<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/just1DayClsTemp.asp" -->
<%
'// 변수 선언
Dim listidx , subIdx , Itemid, title, pclinkUrl, mobilelinkUrl, pcimage, mobileimage, price, saleper, sortnum, isusing
Dim mode, regdate

'// 파라메터 접수
listidx = request("listidx")
subidx = request("subidx")

If subidx = "" Then
	sortnum = "0"
	mode = "subadd"
Else
	mode = "submodify"
End If 

If subidx <> "" then
	dim just1dayItem
	set just1dayItem = new Cjust1Day
	just1dayItem.FRectSubIdx = subidx
	just1dayItem.GetOneSubItem()

	subIdx			=	just1dayItem.FOneItem.FsubIdx
	listidx			=	just1dayItem.FOneItem.Flistidx
	Itemid			=	just1dayItem.FOneItem.FItemid
	sortnum			=	just1dayItem.FOneItem.Fsortnum
	isusing			=	just1dayItem.FOneItem.Fisusing
	title			=	just1dayItem.FOneItem.FTitle
	pclinkurl		=	just1dayItem.FOneItem.FPcLinkUrl
	mobilelinkurl	=	just1dayItem.FOneItem.FMobileLinkUrl
	pcimage			=	just1dayItem.FOneItem.FPcImage
	mobileimage		=	just1dayItem.FOneItem.FMobileImage
	price			=	just1dayItem.FOneItem.FPrice
	saleper			=	just1dayItem.FOneItem.FSaleper
	regdate			=	just1dayItem.FOneItem.Fregdate

	set just1dayItem = Nothing
End If 
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
$(function(){
	//라디오버튼
	$("#rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
	//컬러피커
	$("input[name='subBGColor']").colorpicker();
});

// 폼검사
function SaveForm(frm) {
	//
	if (confirm('저장 하시겠습니까?')){
		//frm.target = "parent";
		frm.submit();
	}
}

// 상품정보 접수
function fnGetItemInfo(iid) {
	$.ajax({
		type: "GET",
		url: "/admin/sitemaster/wcms/act_iteminfo.asp?itemid="+iid,
		dataType: "xml",
		cache: false,
		async: false,
		timeout: 5000,
		beforeSend: function(x) {
			if(x && x.overrideMimeType) {
				x.overrideMimeType("text/xml;charset=euc-kr");
			}
		},
		success: function(xml) {
			if($(xml).find("itemInfo").find("item").length>0) {
				var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='50' />"
					rst += $(xml).find("itemInfo").find("item").find("itemname").text();
				$("#lyItemInfo").fadeIn();
				$("#lyItemInfo").html(rst);
			} else {
				$("#lyItemInfo").fadeOut();
			}
		},
		error: function(xhr, status, error) {
			$("#lyItemInfo").fadeOut();
			/*alert(xhr + '\n' + status + '\n' + error);*/
		}
	});
}

//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}


function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=450,height=300');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<center>
<form name="frm" method="post" action="dojust1day.asp">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="<%=mode%>" />
<input type="hidden" name="listidx" value="<%=listidx%>" />
<input type="hidden" name="subIdx" value="<%=subIdx%>" />
<input type="hidden" name="pcimage" value="<%=pcimage%>">
<input type="hidden" name="mobileimage" value="<%=mobileimage%>">
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>Just1Day 상품 정보 등록/수정</b></td>
</tr>
<colgroup>
	<col width="150" />
	<col width="*" />
</colgroup>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">상품코드</td>
    <td colspan="3">
        <input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="8" class="text" title="상품코드" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">노출명</td>
    <td colspan="3">
        <input type="text" name="title" value="<%= title %>" size="40" maxlength="60" class="text" title="노출명" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">PCIMAGE</td>
	<td colspan="3"><input type="button" name="wimg" value="PC 이미지 등록" onClick="jsSetImg('today','<%=pcimage%>','pcimage','pcimg')" class="button">
		<div id="pcimg" style="padding: 5 5 5 5">
			<%IF pcimage <> "" THEN %>
			<a href="javascript:jsImgView('<%=pcimage%>')"><img  src="<%=pcimage%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('pcimage','pcimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=pcimage%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">MOBILEIMAGE</td>
	<td colspan="3"><input type="button" name="mimg" value="MOBILE 이미지 등록" onClick="jsSetImg('today2','<%=mobileimage%>','mobileimage','mobileimg')" class="button">
		<div id="mobileimg" style="padding: 5 5 5 5">
			<%IF mobileimage <> "" THEN %>
			<a href="javascript:jsImgView('<%=mobileimage%>')"><img  src="<%=mobileimage%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('mobileimage','mobileimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=mobileimage%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">PC 링크구분</td>
	<td  colspan="3">
		<input type="radio" name="pclinkurl" id="pcrdoUsing1" value="/shopping/category_prd.asp" <%=chkIIF(pclinkurl="/shopping/category_prd.asp" or pclinkurl="","checked","")%> /><label for="pcrdoUsing1">일반상품</label>
		<input type="radio" name="pclinkurl" id="pcrdoUsing2" value="/deal/deal.asp" <%=chkIIF(pclinkurl="/deal/deal.asp","checked","")%> /><label for="pcrdoUsing2">딜상품</label>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">MOBILE 링크구분</td>
	<td  colspan="3">
		<input type="radio" name="mobilelinkurl" id="mobilerdoUsing1" value="/category/category_itemPrd.asp" <%=chkIIF(mobilelinkurl="/category/category_itemPrd.asp" or mobilelinkurl="","checked","")%> /><label for="mobilerdoUsing1">일반상품</label>
		<input type="radio" name="mobilelinkurl" id="mobilerdoUsing2" value="/deal/deal.asp" <%=chkIIF(mobilelinkurl="/deal/deal.asp","checked","")%> /><label for="mobilerdoUsing2">딜상품</label>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">가격</td>
    <td colspan="3">
        <input type="text" name="price" value="<%= price %>" size="20" maxlength="60" class="text" title="가격" /> ex ) 14,000원 or 9,900원~
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">할인율</td>
    <td colspan="3">
        <input type="text" name="saleper" value="<%= saleper %>" size="20" maxlength="60" class="text" title="가격" /> ex ) 19% or ~51%
    </td>
</tr>



<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">정렬순서</td>
    <td colspan="3">
        <input type="text" name="sortnum" class="text" size="4" value="<%=sortnum%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">사용여부</td>
    <td colspan="3">
		<span id="rdoUsing">
		<input type="radio" name="isusing" id="rdoUsing1" value="Y" <%=chkIIF(isusing="Y" or isusing="","checked","")%> /><label for="rdoUsing1">사용</label>
		<input type="radio" name="isusing" id="rdoUsing2" value="N" <%=chkIIF(isusing="N","checked","")%> /><label for="rdoUsing2">사용안함</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<% If subidx<>"" Then %>
	    <td colspan="4" align="center"><input type="button" value=" 수 정 " onClick="SaveForm(this.form);"></td>
	<% Else %>
	    <td colspan="4" align="center"><input type="button" value=" 저 장 " onClick="SaveForm(this.form);"></td>
	<% End If %>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->