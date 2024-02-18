<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->

<%
	Dim cDisp, vWidth, vDepth, vCateCode, vCateName, vUseYN, vSortNo, vCurrpage
	vCurrpage = NullFillWith(Request("cpg"), "1")
	vDepth = NullFillWith(Request("depth_s"), "1")
	vCateCode = Request("catecode_s")
	vCateName = Request("catename_s")
	vUseYN = Request("useyn_s")
	vSortNo = Request("sortno_s")
	
	'vWidth = CInt((100/vDepth))
	
	SET cDisp = New cDispCate
	cDisp.FCurrPage = vCurrpage
	cDisp.FPageSize = 1000
	cDisp.FRectDepth = vDepth
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateList()
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
function jsSaveDispCate(){
	if($('input[name="catename"]').val() == ""){
		alert("카테고리명을 입력하세요.");
		$('input[name="catename"]').focus();
		return;
	}
	if($('input[name="sortno"]').val() == ""){
		alert("정렬번호를 입력하세요.");
		$('input[name="sortno"]').focus();
		return;
	}
	if($.isNumeric($('input[name="sortno"]').val()) == false){
		alert("정렬번호는 숫자만 가능합니다.");
		$('input[name="sortno"]').val('');
		$('input[name="sortno"]').focus();
		return;
	}
	frmDispCate.submit();
}

function jsWriteCateCode(c,d,p){
	$.ajax({
			url: "display_cate_ajax.asp?catecode_s="+c+"&depth="+d+"&parentcatecode="+p+"",
			cache: false,
			success: function(message)
			{
				$("#catecodewritebox").empty().append(message);
				$("#catecodewritebox").show();
			}
	});
}

function jsWriteFormClose(){
	$("#catecodewritebox").empty().append("");
	$("#catecodewritebox").hide();
}
</script>

<table border="0" class=a>
<tr>
	<td style="padding:10px 0 10px 0;">* 카테고리 수정하려면 [카테고리코드] 를 클릭하세요.</td>
</tr>
</table>

<table border="0" class=a>
<tr>
	<td><input type="button" value="<%=vDepth%> Depth 카테고리생성" onClick="jsWriteCateCode('','<%=vDepth%>','<%=vCateCode%>');"></td>
</tr>
<tr>
	<td>
		<form name="frmDispCate" method="post" action="display_cate_proc.asp" target="cateproc">
		<div id="catecodewritebox" style="display:none;">
		</div>
		</form>
	</td>
</tr>
</table>

<table width="<%=(271*vDepth)%>" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td width="<%=(271*vDepth)%>">
		<table border="0" class="a">
		<tr>
			<% For i=1 To vDepth %>
			<td width="271">▽ <b><%=i%></b></td>
			<% Next %>
		</tr>
		</table>
	</td>
</tr>
</table>
<table border="0" cellspacing="0" cellpadding="0">
<tr>
	<td valign=top>
		<table border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<%
			Dim i, vTempDepth, vIsThisLine
			For i=0 To cDisp.FResultCount-1
				vIsThisLine = fnIsThisLine(cDisp.FItemList(i).FDepth,cDisp.FItemList(i).FCateCode,vCateCode)
				If i = 0 Then vTempDepth = cDisp.FItemList(i).FDepth End If
				
				If vTempDepth <> cDisp.FItemList(i).FDepth Then
					Response.Write "</table></td><td valign=top><table border=0 align=center cellpadding=2 cellspacing=1 class=a bgcolor=#CCCCCC>"
				End If
				
				If i = 0 Then
		%>
				<tr>
					<td bgcolor="#FFFFFF" width="250">
						<table border=0 class=a>
						<tr>
							<td width="245"><a href="/admin/CategoryMaster/displaycate/display_cate_list.asp?menupos=<%=Request("menupos")%>">Go 1 Depth</a></td>
							<td width="5" align="right"></td>
						</tr>
						</table>
					</td>
				</tr>
				<% End If %>
				<tr>
					<td bgcolor="<%=CHKIIF(vIsThisLine="o","#F1F1F1","#FFFFFF")%>" width="260">
						<table border=0 class=a>
						<tr>
							<td width="255">
								<span onClick="jsWriteCateCode('<%=cDisp.FItemList(i).FCateCode%>','<%=cDisp.FItemList(i).FDepth%>','');" style="cursor:pointer;">[<%=Right(cDisp.FItemList(i).FCateCode,3)%>]</span>
								<a href="<%=CurrURL()%>?menupos=<%=Request("menupos")%>&depth_s=<%=cDisp.FItemList(i).FDepth+1%>&catecode_s=<%=cDisp.FItemList(i).FCateCode%>"><%=cDisp.FItemList(i).FCateName%></a>
								
							</td>
							<td width="5" align="right"><%=CHKIIF(vIsThisLine="o","▶","")%></td>
						</tr>
						</table>
					</td>
				</tr>
		<%	
				vTempDepth = cDisp.FItemList(i).FDepth
			Next
		%>
		</table>
	</td>
</tr>
</table>
<br>-- 상품리스트 --
<iframe src="" id="cateproc" name="cateproc" width="0" height="0"></iframe>

<% SET cDisp = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->