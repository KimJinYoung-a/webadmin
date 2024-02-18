<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_hot_managecls.asp" -->
<%
dim idx, poscode,reload, cdl, cdm , cds
idx = request("idx")
poscode = request("poscode")
reload = request("reload")
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds") '2012 추가 : 이종화

if idx="" then idx=0

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

dim oCateContents
set oCateContents = new CCateContents
oCateContents.FRectIdx = idx
oCateContents.GetOneCateContents

If cdl = "" Then
	cdl = oCateContents.FOneItem.Fcdl
End IF

If cdm = "" AND cdl = oCateContents.FOneItem.Fcdl Then
	cdm = oCateContents.FOneItem.Fcdm
End If

'2012-04-03 이종화 추가
If cds = "" AND cdl = oCateContents.FOneItem.Fcdl AND cdm = oCateContents.FOneItem.Fcdm Then
	cds = oCateContents.FOneItem.Fcds
End If

dim oposcode, defaultMapStr

    defaultMapStr = "<map name='Map_hot_" & cdl & cdm & "'>" + VbCrlf
    defaultMapStr = defaultMapStr + VbCrlf
    defaultMapStr = defaultMapStr + "</map>"

%>

<script language='javascript'>
function SaveCateContents(frm){
    if (frm.cdl.value == ""){
       alert('대카테고리를 입력 하세요.');
        frm.cdl.focus();
        return;
    }
    
    if (frm.cdm.value == ""){
       alert('중카테고리를 입력 하세요.');
        frm.cdm.focus();
        return;
    }
    
//    if (frm.linktype.value == ""){
//       alert('링크 구분을 입력 하세요.');
//        frm.linktype.focus();
//        return;
//    }
//    
//    if (frm.linktype.value == "F" || frm.linktype.value == "X"){
//        alert('링크 구분은\n\n링크 (a href)\n맵 (#Map)\n\n만 선택하세요.');
//        frm.linktype.focus();
//        return;
//    }
//    
//    if (frm.linkurl.value.length<1){
//        alert('링크 값을 입력 하세요.');
//        frm.linkurl.focus();
//        return;
//    }
    
    if (frm.startdate.value.length!=10){
        alert('시작일을 입력  하세요.');
        frm.startdate.focus();
        return;
    }
    
    if (frm.enddate.value.length!=10){
        alert('종료일을 입력  하세요.');
        frm.enddate.focus();
        return;
    }
    
    var vstartdate = new Date(frm.startdate.value.substr(0,4), frm.startdate.value.substr(5,2), frm.startdate.value.substr(8,2));
    var venddate = new Date(frm.enddate.value.substr(0,4), frm.enddate.value.substr(5,2), frm.enddate.value.substr(8,2));
    
    if (vstartdate>venddate){
        alert('종료일이 시작일보다 클 수 없습니다.');
        frm.enddate.focus();
        return;
    }

    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}

// 카테고리 변경시 명령
function changecontent(){
	<% If oCateContents.FOneItem.Fidx <> "" Then %>
		alert("카테고리를 변경할 시 맵 name에 Map_hot_ 뒤 코드값(대중카테고리순)을 수기로 변경해야 합니다. ");
		document.getElementById("categorylist").style.display = "block";
	<% Else %>
		location.href = "?cdl=" + frmcontents.cdl.value + "<%=CHKIIF(cdl<>"","&cdm="&chr(34)&" + frmcontents.cdm.value + "&chr(34)&"","")%><%=CHKIIF(cdm<>"","&cds="&chr(34)&" + frmcontents.cds.value + "&chr(34)&"","")%>&idx=<%=idx%>";
	<% End If %>
}
</script>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/Category/doCateHotReg.asp" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF">
    <td width="20%" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        <%= oCateContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oCateContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">카테고리</td>
    <td>
        <%
        	if oCateContents.FOneItem.Fidx<>"" then
        		call DrawSelectBoxCategoryLarge("cdl", cdl)
        		Response.Write "&nbsp;"
        		if cdl <> "" then
        			call DrawSelectBoxCategoryMid("cdm",cdl, cdm)
					Response.Write "&nbsp;"
					If cdm <> "" Then
						call DrawSelectBoxCategorySmall("cds",cdl, cdm, cds )
					End If 
        		end if
        	else
    			call DrawSelectBoxCategoryLarge("cdl", cdl)
    			Response.Write "&nbsp;"
    			if cdl <> "" then
    				call DrawSelectBoxCategoryMid("cdm",cdl, cdm)
					Response.Write "&nbsp;"
					If cdm <> "" Then
						call DrawSelectBoxCategorySmall("cds",cdl, cdm, cds )
					End If 
    			end if
        	end if
        %>
        <br>
        <div id="categorylist" style="display:none;">
        <%
			   dim tmp_str,query1, tt
			   tt = 1
			
			   query1 = " select code_mid, code_nm from [db_item].[dbo].tbl_Cate_mid"
			   query1 = query1 & " where display_yn = 'Y'"
			   query1 = query1 & " and code_large = '" & cdl & "'"
			   query1 = query1 & " and code_mid<>0"
			   query1 = query1 & " order by code_mid Asc"
			
			   rsget.Open query1,dbget,1
			
			   if  not rsget.EOF  then
			       rsget.Movefirst
			
			       do until rsget.EOF
			           response.write("["&rsget("code_mid")&"]"& db2html(rsget("code_nm")) &",")
			           if tt = 5 then response.write "<br>" end if
			           rsget.MoveNext
			           tt = tt + 1
			       loop
			   end if
			   rsget.close
        %>
        </div>
    </td>
</tr>
<!-- <tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크구분</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        	<%= DrawLinktypeCombo("",oCateContents.FOneItem.Flinktype,"") %>
        <% else %>
            <%= DrawLinktypeCombo("","","") %>
        <% end if %>
        &nbsp;&nbsp;링크구분은 링크 (a href), 맵 (#Map) 만 선택하세요.
    </td>
</tr> -->

<!-- <tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지</td>
  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
  <% if oCateContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oCateContents.FOneItem.GetImageUrl %>" >
  <br> <%= oCateContents.FOneItem.GetImageUrl %>
  <% end if %>
  </td>
</tr> -->
<!-- <tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크값<br><br><b><font color="red">URL인 경우<br>맵 삭제 후 입력</font></b></td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
            <textarea name="linkurl" cols="60" rows="6"><%= oCateContents.FOneItem.Flinkurl %></textarea>
            <br><b>(이미지맵일 경우 name값 절대 변경 금지)</b>
        <% else %>
                <textarea name="linkurl" cols="60" rows="6"><%= defaultMapStr %></textarea>
                <br><b>(이미지맵일 경우 name값 절대 변경 금지)</b>
        <% end if %>
        <br><b>맵인경우 map name='Map_hot_ 에 반드시 대카테코드와 중카테코드가 들어가야함!<br>모든 맵소스는 칸내림없이 다 붙여야함!<br>링크에 자바스크립트 코드 넣으면 안됨!</b>
    </td>
</tr> -->
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영시작일</td>
    <td>
        <input type="text" name="startdate" value="<%= oCateContents.FOneItem.Fstartdate %>" maxlength="10" size="10"> (<a href="#" onClick="frmcontents.startdate.value='<%= Left(CStr(now()),10) %>'"><%= Left(CStr(now()),10) %></a>)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영종료일</td>
    <td>
        <input type="text" name="enddate" value="<%= Left(oCateContents.FOneItem.Fenddate,10) %>" maxlength="10" size="10">
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly style="background:'#EEEEEE'">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">등록일</td>
    <td>
        <%= oCateContents.FOneItem.Fregdate %> (<%= oCateContents.FOneItem.Freguserid %>)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용여부</td>
    <td>
        <% if oCateContents.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">사용함
        <input type="radio" name="isusing" value="N" checked >사용안함
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >사용함
        <input type="radio" name="isusing" value="N">사용안함
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center" style="padding:5 0 5 0">
    <table cellpadding="0" cellspacing="0" border="0">
	<tr><td style="padding-bottom:5px;">※ <b><font color="red">페이지뷰가 높은 카테고리(메인,리스트)여서 실시간 적용하면 서버가 죽습니다!!</font></b></td></tr>
    <tr><td style="padding-bottom:5px;">※ <b><font color="blue">시작일은 오늘기준으로 하루전까지 해주세요.!! 시작일을 오늘로 하면 적용이 안됨.</font></b></td></tr>
    <tr><td style="padding-bottom:5px;">※ <b><font color="red">단 <font color="blue">오늘날짜로 새벽 6시 이전까지</font> 올리는것은 적용됩니다.</font></b></td></tr>
    <tr><td style="padding-bottom:15px;">※ <b><font color="red">링크 구분을 확실히 해주세요!!</font> 그냥 <font color="blue">링크일때는 링크</font>를, <font color="blue">맵 일경우 맵</font>으로 선택!!!</b></td></tr>
    </table>
    <input type="button" value=" 저 장 " onClick="SaveCateContents(frmcontents);">
    </td>
</tr>
</form>
</table>

<script language="JavaScript">
<!--
var speed = 100 //깜빡이는 속도 - 1000은 1초

function doBlink(){
var blink = document.all.tags("blink")
for (var i=0; i < blink.length; i++)
blink[i].style.visibility = blink[i].style.visibility == "" ? "hidden" : ""
} 

function startBlink() { 
setInterval("doBlink()",speed)
} 
window.onload = startBlink; 
//-->
</script>

<%
set oposcode = Nothing
set oCateContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
