<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/kaffa/kaffaCls.asp"-->

<%
dim notmatch, research, page, cdl
notmatch = request("notmatch")
research = request("research")
page     = request("page")
cdl      = RequestCheckVar(request("cdl"),3)

if ((research = "") and (notmatch = "")) then notmatch = "on"
if (page = "") then page = 1
'if (cdl = "") then cdl = "010"


dim cKaffa, i, vSelectOption, aa, vTenCate, vCate2Arr, vCate3Arr, vOptBody2, vOptBody3, vTotalCount
vSelectOption = KaffaCate1SelectBox()
set cKaffa = new cKaffaItem
vCate2Arr = cKaffa.GetKaffaCate2List
vCate3Arr = cKaffa.GetKaffaCate3List
set cKaffa = nothing


set cKaffa = new cKaffaItem
cKaffa.FRectNotMatchCategory = notmatch
cKaffa.FRectCate_large = cdl

'if (cdl<>"") then
    cKaffa.GetKaffaCategoryMachingList
    vTotalCount = cKaffa.FTotalCount
'end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
function goSearch(){
	searchfrm.submit();
}

function goCate2(area,v){
	$.ajax({
			url: "cate_option_ajax.asp?categubun=2&area="+area+"&cate1="+v,
			cache: false,
			success: function(message)
			{
				if(message != "x"){
					$("#"+area+"").empty().append(message);
				}else{
					$("#"+area+"").empty().append("없음");
				}
				$("#"+area.replace('cate2','cate3')+"").empty();
			}
	});
}

function goCate3(area,v1,v2){
	$.ajax({
			url: "cate_option_ajax.asp?categubun=3&area="+area+"&cate1="+v1+"&cate2="+v2,
			cache: false,
			success: function(message)
			{
				if(message != "x"){
					$("#"+area+"").empty().append(message);
				}else{
					$("#"+area+"").empty().append("없음");
				}
			}
	});
}

function procCate(tencode){
	var cate1 = "";
	var cate2 = "";
	var cate3 = "";

	var cate1 = $("select[name=kaffacate1-"+tencode+"]").val();
	if($("select[name=kaffacate2-"+tencode+"]").length > 0){
		cate2 = $("select[name=kaffacate2-"+tencode+"]").val();
	}else{
		cate2 = "0";
	}

	if($("select[name=kaffacate3-"+tencode+"]").length > 0){
		cate3 = $("select[name=kaffacate3-"+tencode+"]").val();
	}else{
		cate3 = "0";
	}

	frm1.tencode.value = tencode;
	frm1.cate1.value = cate1;
	frm1.cate2.value = cate2;
	frm1.cate3.value = cate3;

	frm1.submit();
}
</script>
<form name="searchfrm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="research" value="o">
<table class="a"  width="100%">
<tr>
	<td style="padding:5 0 10 0;">
		Total : <%=vTotalCount%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<select name="cdl" onChange="goSearch();">
		<option value="">-10x10카테고리선택-</option>
		<%=CateLargeSelectBox(cdl)%>
		</select>
	</td>
	<td align="right">
		<input type="checkbox" name="notmatch" onClick="goSearch();" <%=CHKIIF(notmatch="on","checked","")%>> 매칭안된것만보기
	</td>
</tr>
</table>
</form>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="100">Ten 카테코드</td>
	<td width="150">대분류</td>
	<td width="150">중분류</td>
	<td width="150">소분류</td>
	<td width="150">kaffa 카테1</td>
	<td width="150">kaffa 카테2</td>
	<td width="150">kaffa 카테3</td>
	<td></td>
</tr>
<% for i=0 to cKaffa.FResultCount-1
	vTenCate = cKaffa.FItemList(i).FCate_Large & cKaffa.FItemList(i).FCate_Mid & cKaffa.FItemList(i).FCate_Small
	aa = "3"
%>
<tr align="center" bgcolor="#FFFFFF">
    <td><%= vTenCate %></td>
    <td><%= cKaffa.FItemList(i).Fnmlarge %></td>
    <td><%= cKaffa.FItemList(i).FnmMid %></td>
    <td><%= cKaffa.FItemList(i).FnmSmall %></td>
    <td>
    	<select name="kaffacate1-<%=vTenCate%>" onChange="goCate2('cate2-<%=vTenCate%>',this.value);">
    	<option value="x">-</option>
    	<%=Replace(vSelectOption,"="""&cKaffa.FItemList(i).FKaffacate1&""">","="""&cKaffa.FItemList(i).FKaffacate1&""" selected>")%>
    	</select>
    </td>
    <td>
    	<div id="cate2-<%=vTenCate%>">
    	<%
    		IF cKaffa.FItemList(i).FKaffacate2 = "" OR isNull(cKaffa.FItemList(i).FKaffacate2) = True Then

    		Else
    			vOptBody2 = KaffaCate2SelectBox(vCate2Arr, cKaffa.FItemList(i).FKaffacate1, cKaffa.FItemList(i).FKaffacate2)
    			If vOptBody2 <> "" Then
	    			Response.Write "<select name=""kaffacate2-"&vTenCate&""" onChange=""goCate3('cate3-"&vTenCate&"',$('select[name=kaffacate1-"&vTenCate&"]').val(),this.value);"">"
	    			Response.Write KaffaCate2SelectBox(vCate2Arr, cKaffa.FItemList(i).FKaffacate1, cKaffa.FItemList(i).FKaffacate2)
	    			Response.Write "</select>"
	    		Else
	    			Response.Write "없음"
	    		End If
	    		vOptBody2 = ""
    		End If
    	%>
    	</div>
    </td>
    <td>
    	<div id="cate3-<%=vTenCate%>">
    	<%
    		IF cKaffa.FItemList(i).FKaffacate3 = "" OR isNull(cKaffa.FItemList(i).FKaffacate3) = True Then

    		Else
    			vOptBody3 = KaffaCate3SelectBox(vCate3Arr, cKaffa.FItemList(i).FKaffacate1, cKaffa.FItemList(i).FKaffacate2, cKaffa.FItemList(i).FKaffacate3)
				If vOptBody3 <> "" Then
	    			Response.Write "<select name=""kaffacate3-"&vTenCate&""">"
	    			Response.Write vOptBody3
	    			Response.Write "</select>"
	    		Else
	    			Response.Write "없음"
	    		End If
	    		vOptBody3 = ""
    		End If
    	%>
    	</div>
    </td>
    <td><input type="button" value="저장" onClick="procCate('<%= vTenCate %>');"><span id="result-<%=vTenCate%>"></span></td>
</tr>
<% next %>
</table>

<%
set cKaffa = Nothing
%>

<form name="frm1" method="post" action="cate_proc.asp" target="codeproc">
<input type="hidden" name="tencode" value="">
<input type="hidden" name="cate1" value="0">
<input type="hidden" name="cate2" value="0">
<input type="hidden" name="cate3" value="0">
</form>
<iframe src="" id="codeproc" name="codeproc" width="0" height="0"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->