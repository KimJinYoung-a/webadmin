<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/onlyBrandCls.asp" -->
<%
'// 변수 선언
Dim listidx , subIdx , subItemid , itemName , sortnum , isusing
Dim mode , smallImage , subImage1 , extraurl, orderby

'// 파라메터 접수
listidx = request("listidx")
subidx = request("subidx")

If subidx = "" Then
	orderby = "0"
	mode = "subadd"
Else
	mode = "submodify"
End If 

If subidx <> "" then
	dim onlybrandItem
	set onlybrandItem = new Conlybrand
	onlybrandItem.FRectSubIdx = subidx
	onlybrandItem.GetOneSubItem()

	subIdx			=	onlybrandItem.FOneItem.FsubIdx
	listidx			=	onlybrandItem.FOneItem.Flistidx
	subItemid		=	onlybrandItem.FOneItem.FItemid
	orderby		=	onlybrandItem.FOneItem.Forderby
	isusing			=	onlybrandItem.FOneItem.Fisusing
	itemName		=	onlybrandItem.FOneItem.FitemName
	smallImage	=	onlybrandItem.FOneItem.FsmallImage

	set onlybrandItem = Nothing
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
</script>
<center>
<form name="frm" method="post" action="doonlybrand.asp">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="<%=mode%>" />
<input type="hidden" name="listidx" value="<%=listidx%>" />
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>소재 정보 등록/수정</b></td>
</tr>
<colgroup>
	<col width="100" />
	<col width="*" />
	<col width="100" />
	<col width="*" />
</colgroup>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">소재 번호</td>
    <td colspan="3">
        <%=subIdx %>
        <input type="hidden" name="subIdx" value="<%=subIdx%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">상품코드</td>
    <td colspan="3">
        <input type="text" name="subItemid" value="<%= subItemid %>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value)" title="상품코드" />
        <div id="lyItemInfo" style="display:<%=chkIIF(subItemid="","none","")%>;">
        <%
        	if Not(itemName="" or isNull(itemName)) then
        		Response.Write "<img src='" & smallImage & "' height='50' />"
        		Response.Write itemName
        	end if
        %>
        </div>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">정렬순서</td>
    <td>
        <input type="text" name="orderby" class="text" size="4" value="<%=orderby%>" />
    </td>
    <td bgcolor="#DDDDFF">사용여부</td>
    <td>
		<span id="rdoUsing">
		<input type="radio" name="isusing" id="rdoUsing1" value="Y" <%=chkIIF(isusing="Y" or isusing="","checked","")%> /><label for="rdoUsing1">사용</label>
		<input type="radio" name="isusing" id="rdoUsing2" value="N" <%=chkIIF(isusing="N","checked","")%> /><label for="rdoUsing2">삭제</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" 저 장 " onClick="SaveForm(this.form);"></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->