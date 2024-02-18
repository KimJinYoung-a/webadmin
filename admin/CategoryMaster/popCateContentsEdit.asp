<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_contents_managecls.asp" -->
<%
dim idx, poscode,reload, cdl, vCateCode

dim  isusing, fixtype  , validdate , prevDate 
dim strParm

isusing = request("isusing")  
fixtype = request("fixtype") 
validdate= request("validdate")
prevDate = request("prevDate")   
idx = request("idx")
poscode = request("poscode")
reload = request("reload")
cdl = request("cdl")
vCateCode = Request("catecode")

if idx="" then idx=0


if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

dim oCateContents
set oCateContents = new CCateContents
oCateContents.FRectIdx = idx
oCateContents.GetOneCateContents

dim oposcode, defaultMapStr
set oposcode = new CCateContentsCode
oposcode.FRectPosCode = poscode
if poscode<>"" then
    oposcode.GetOneContentsCode
    
    defaultMapStr = "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>" + VbCrlf
    defaultMapStr = defaultMapStr + VbCrlf
    defaultMapStr = defaultMapStr + "</map>"
end if
%>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function SaveCateContents(frm){
    if (frm.poscode.value.length<1){
        alert('구분을 먼저 선택 하세요.');
        frm.poscode.focus();
        return;
    }
    
    <% If poscode <> "370" or poscode <> "367" Then %>
    if (frm.linkurl.value.length<1){
        alert('링크 값을 입력 하세요.');
        frm.linkurl.focus();
        return;
    }
	<% end if %>

	<% If poscode = "367" Then %>
    if (frm.evtcode.value=="" || frm.evtcode.value==0){
        alert('이벤트코드를 입력 하세요.');
        frm.evtcode.focus();
        return;
    }
	<% end if %>
    
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

function ChangeLinktype(comp){
    if (comp.value=="M"){
       document.all.link_M.style.display = "";
       document.all.link_L.style.display = "none";
    }else{
       document.all.link_M.style.display = "none";
       document.all.link_L.style.display = "";
    }
}

//function getOnLoad(){
//    ChangeLinktype(frmcontents.linktype.value);
//}

//window.onload = getOnLoad;

function ChangeGubun(comp){
    location.href = "?poscode=" + comp.value;
    // nothing;
}

function ImgDelJs()
{
	alert("저장을 눌러야 이미지가 삭제 됩니다.");
	document.getElementById("ImgAreaVal").innerHTML="";
	document.frmcontents.imgDelProcType.value="1";
}

</script>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/Category/doCateContentsReg.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="imgDelProcType">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        <%= oCateContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oCateContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분명</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        <%= oCateContents.FOneItem.Fposname %> (<%= oCateContents.FOneItem.Fposcode %>)
        <input type="hidden" name="poscode" value="<%= oCateContents.FOneItem.Fposcode %>">
        <% else %>
        <% call DrawCatePosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'") %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">전시카테고리</td>
    <td>
        <%
        	Dim cDisp, i
        	
        	if oCateContents.FOneItem.Fidx<>"" then
				SET cDisp = New cDispCate
				cDisp.FCurrPage = 1
				cDisp.FPageSize = 2000
				cDisp.FRectDepth = 1
				'cDisp.FRectUseYN = "Y"
				cDisp.GetDispCateList()
				
				If cDisp.FResultCount > 0 Then
					Response.Write "<select name=""catecode"" class=""select"">" & vbCrLf
					Response.Write "<option value="""">선택</option>" & vbCrLf
					For i=0 To cDisp.FResultCount-1
						Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(oCateContents.FOneItem.Fdisp1)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
					Next
					Response.Write "</select>&nbsp;&nbsp;&nbsp;"
				End If
				Set cDisp = Nothing
        	else
        		if poscode<>"" then
					SET cDisp = New cDispCate
					cDisp.FCurrPage = 1
					cDisp.FPageSize = 2000
					cDisp.FRectDepth = 1
					'cDisp.FRectUseYN = "Y"
					cDisp.GetDispCateList()
					
					If cDisp.FResultCount > 0 Then
						Response.Write "<select name=""catecode"" class=""select"">" & vbCrLf
						Response.Write "<option value="""">선택</option>" & vbCrLf
						For i=0 To cDisp.FResultCount-1
							Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
						Next
						Response.Write "</select>&nbsp;&nbsp;&nbsp;"
					End If
					Set cDisp = Nothing
        		else
        %>
            <font color="red">구분을 먼저 선택하세요</font>
        <%
        		end if
        	end if
        %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크구분</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        <%= oCateContents.FOneItem.getlinktypeName %>
        <% else %>
            <% if poscode<>"" then %>
            <%= oposcode.FOneItem.getlinktypeName %>
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">적용구분(반영주기)</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        <%= oCateContents.FOneItem.getfixtypeName %>
        <input type="hidden" name="fixtype" value="<%= oCateContents.FOneItem.Ffixtype %>">
        <% else %>
            <% if poscode<>"" then %>
            <%= oposcode.FOneItem.getfixtypeName %>
            <input type="hidden" name="fixtype" value="<%= oposcode.FOneItem.Ffixtype %>">
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
        
    </td>
</tr>

<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지</td>
  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
  <% if oCateContents.FOneItem.Fidx<>"" then %>
	  <% If oCateContents.FOneItem.GetImageUrl <>"" Then %>
		<% If poscode="367" Then %>
			  <br>
			  <div id="ImgAreaVal">
				  <img src="<%= oCateContents.FOneItem.GetImageUrl %>" >
				  <br> <%= oCateContents.FOneItem.GetImageUrl %>
				  <br> <a href="" onclick="ImgDelJs();return false;">[이미지 삭제]</a>
			  </div>
		 <% End If %>
	  <% End If %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지Width</td>
  <td>
  <% if oCateContents.FOneItem.Fidx<>"" then %>
  <input type="text" name="imagewidth" value="<%= oCateContents.FOneItem.Fimagewidth %>" size="8" maxlength="16"> 
  <% else %>
        <% if poscode<>"" then %>
        <%= oposcode.FOneItem.Fimagewidth %>
        <% else %>
        <font color="red">구분을 먼저 선택하세요</font>
        <% end if %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지Height</td>
  <td>
  <% if oCateContents.FOneItem.Fidx<>"" then %>
  <input type="text" name="imageheight" value="<%= oCateContents.FOneItem.Fimageheight %>" size="8" maxlength="16"> 
  <% else %>
        <% if poscode<>"" then %>
        <%= oposcode.FOneItem.Fimageheight %>
        <% else %>
        <font color="red">구분을 먼저 선택하세요</font>
        <% end if %>
  <% end if %>
  </td>
</tr>
<% If poscode = "370" Then %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">브랜드아이디</td>
    <td>
    	<input type="text" name="makerid" id="[off,off,off,off][브랜드ID]" value="<%= oCateContents.FOneItem.Fmakerid %>" size="20">
    	<input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'makerid');" >
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">브랜드픽 카피(60자 이하)</td>
    <td><textarea name="brandcopy" cols="70" rows="3"><%= oCateContents.FOneItem.Fbrandcopy %></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크값</td>
    <td><textarea name="linkurl" cols="70" rows="3"><%= oCateContents.FOneItem.Flinkurl %></textarea></td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크값</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
            <% if oCateContents.FOneItem.FLinkType="M" then %>
            <textarea name="linkurl" cols="60" rows="6"><%= oCateContents.FOneItem.Flinkurl %></textarea>
            <% else %>
            <input type="text" name="linkurl" value="<%= oCateContents.FOneItem.Flinkurl %>" maxlength="128" size="70">
            <% end if %>
        <% else %>
            <% if poscode<>"" then %>
                <% if oposcode.FOneItem.FLinkType="M" then %>
                    <textarea name="linkurl" cols="60" rows="6"><%= defaultMapStr %></textarea>
                    <br>(이미지맵 변수값 변경 금지)
                <% else %>
                    <input type="text" name="linkurl" value="" maxlength="128" size="70">
                    <br>(상대경로로 표시해 주세요  ex: /event/eventmain.asp?eventid=6263)
                <% end if %>
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<% End If %>
<% If poscode = "367" Then %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이벤트코드</td>
    <td>
    	<input type="text" name="evtcode" value="<%= oCateContents.FOneItem.FevtCode %>" size="20">
    </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영시작일</td>
    <td>
        <input id="startdate" type="text" name="startdate" value="<%= oCateContents.FOneItem.Fstartdate %>" maxlength="10" size="10">
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
        <input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
	    <script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "startdate",
			trigger    : "startdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영종료일</td>
    <td>
        <input id="enddate" type="text" name="enddate" value="<%= Left(oCateContents.FOneItem.Fenddate,10) %>" maxlength="10" size="10">
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
	    <script type="text/javascript">
		var CAL_End = new Calendar({
			inputField : "enddate",
			trigger    : "enddate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">등록일</td>
    <td>
        <%= oCateContents.FOneItem.Fregdate %> (<%= oCateContents.FOneItem.Fregname %>)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">작업자</td>
    <td>
    	<% If idx <> "" AND idx <> "0" Then %>
    	최종 작업자 : <%=oCateContents.FOneItem.Fworkername%><input type="hidden" name="selDId" value="<%=session("ssBctId")%>">
    	<% Else %>
    		<input type="hidden" name="selDId" value="">
    	<% End If %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">정렬번호</td>
    <td>
        <input type="text" name="sortNo" value="<% if oCateContents.FOneItem.FsortNo<>"" then Response.Write oCateContents.FOneItem.FsortNo: Else Response.Write "0": End if %>" maxlength="4" size="5" style="text-align:right">
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
    <td width="150" bgcolor="#DDDDFF">작업 코멘트</td>
    <td>
        <textarea name="desc" class="textarea" style="width:98%;height:100px;"><%=oCateContents.FOneItem.Fdesc%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveCateContents(frmcontents);"></td>
</tr>
</form>
</table>


<%
set oposcode = Nothing
set oCateContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
