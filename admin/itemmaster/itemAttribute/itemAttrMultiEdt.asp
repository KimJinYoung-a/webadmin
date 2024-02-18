<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/items/itemAttribMultiCls.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i, j
Dim dispCate : dispCate = requestCheckvar(request("disp"),15)
Dim subcate : subcate = requestCheckvar(request("subcate"),5000)
dim page : page = requestCheckvar(request("page"),10)

if (page="") then page=1

Dim catecodelist 
Dim oAttrItemList

if (dispCate<>"") then catecodelist=dispCate

dim splitsubcate : splitsubcate= split(subcate,",")
dim onesubcate 
if isArray(splitsubcate) then
    for i=LBound(splitsubcate) to Ubound(splitsubcate)
        onesubcate = Trim(splitsubcate(i))
        onesubcate = MID(onesubcate,5,30)
        if (catecodelist="") then
            catecodelist = onesubcate
        else
            catecodelist = catecodelist&","&onesubcate
        end if
    next
end if

if catecodelist="" then 
   ' catecodelist="101102101101"
  '  subcate = "100-101102101101"
end if

SET oAttrItemList = new CAttribMulti
    oAttrItemList.FPageSize = 50
    oAttrItemList.FCurrPage = page

    oAttrItemList.FRectcateCodeList = catecodelist
    if (catecodelist<>"") then
        oAttrItemList.GetAttribMultiItemList
    end if

dim pgrptitle, grpcolcnt : grpcolcnt=0
%>
<style>
.stickyhead {
  position: sticky;
  top: 0;
  
}
.stickyhead2 {
  position: sticky;
  top: 34;
  
}

</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function researchfrm(){
    document.frmsearch.page.value = 1;
    document.frmsearch.submit();
}

function goPage(page){
    document.frmsearch.page.value = page;
    document.frmsearch.submit();
}

function switchCheckBox(comp){
    var frm = comp.form;

    if(frm.chkix.length>1){
        for(i=0;i<frm.chkix.length;i++){
            if (!frm.chkix[i].disabled){
                frm.chkix[i].checked = comp.checked;
                AnCheckClick(frm.chkix[i]);
            }
        }
    }else{
        if (!frm.chkix.disabled){
            frm.chkix.checked = comp.checked;
            AnCheckClick(frm.chkix);
        }
    }
}

function chgCheckAttr(comp,ix){
    var frm = comp.form;
    
    if (comp.value*1>=1){
        if (frm.chkix.length>1){
            frm.chkix[ix].checked=true;
            AnCheckClick(frm.chkix[ix]);
        }else{
            frm.chkix.checked=true;
            AnCheckClick(frm.chkix);
        }
    }
}

function CheckNChangAttr(comp){
    var frm = comp.form;
    var pass = false;

    if (!frm.chkix){
        alert("선택 내역이 없습니다.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            pass = (pass||frm.chkix[i].checked);
        }
    }else{
        pass = frm.chkix.checked;
    }

    if (!pass) {
        alert("선택 내역이 없습니다.");
        return;
    }

    if (confirm("선택 내역을 수정 하시겠습니까?")){
        frm.mode.value="itemattrmultiset";
        frm.submit();
    }
}

</script>
<body>
<form name="frmsearch" method="get" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
    전시카테고리: 
    <!-- #include virtual="/common/module/dispCateSelectBox.asp"--> 
    OR<br>
    전문샵전시카테고리 : 
    <%= drawSubshopcateCheckBox("subcate",subcate)%>

    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="researchfrm();">
	</td>
</tr>
<tr align="center" bgcolor="#F4F4F4">
    <td align="left">
    
    </td>
</tr>
</table>
</form>
<p>
   
</p>
<form name="frmAttr" method="POST" action="itemAttrMultiEdt_process.asp" style="margin:0;">
<input type="hidden" name="mode">
<input type="hidden" name="catecodelist" value="<%=catecodelist%>">

<table cellpadding="3" cellspacing="2" border="0" align="center" width="100%" bgcolor="#CCCCCC" class="a">
<thead class="stickyhead">
<tr height="25" bgcolor="FFFFFF" >
	<th colspan="6" align="left" bgcolor="#FFFFFF" >
		검색결과 : <b><%=oAttrItemList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oAttrItemList.FtotalPage%></b>
	</th>
    <% if oAttrItemList.FResultCount>0 then %>
    <th colspan="<%=oAttrItemList.FAttrResultCount%>" align="right" bgcolor="#FFFFFF">
        <input type="button" value="선택상품 속성 수정" onClick="CheckNChangAttr(this)";>
    </th>
    <% end if %>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" >
    <th rowspan="2" class="stickyhead" bgcolor="<%= adminColor("tabletop") %>"><input type="checkbox" name="chkALL" onClick="switchCheckBox(this);"></th>
    <th rowspan="2" class="stickyhead" bgcolor="<%= adminColor("tabletop") %>" width="50">이미지</th>
    <th rowspan="2" class="stickyhead" bgcolor="<%= adminColor("tabletop") %>" width="60">상품코드</th>
    <th rowspan="2" class="stickyhead" bgcolor="<%= adminColor("tabletop") %>" >상품명</th>
    <th rowspan="2" class="stickyhead" bgcolor="<%= adminColor("tabletop") %>" width="70">브랜드ID</th>
    <th rowspan="2" class="stickyhead" bgcolor="<%= adminColor("tabletop") %>" width="70">카테고리</th>
    <% if (oAttrItemList.FAttrResultCount>0) then %>
        <% for j=0 to oAttrItemList.FAttrResultCount-1 %>
            <% if (pgrptitle<>"") and (pgrptitle<>oAttrItemList.FAttrGbnList(j).FattribDivName) then %>
            <th colspan="<%=grpcolcnt%>" class="stickyhead" bgcolor="<%= adminColor("tabletop") %>" ><%=pgrptitle%></th>
            <% grpcolcnt=0 %>
            <% end if %>
            <%
            pgrptitle=oAttrItemList.FAttrGbnList(j).FattribDivName
            grpcolcnt=grpcolcnt+1
            %>
        <% next %>
        <th colspan="<%=grpcolcnt%>" class="stickyhead" bgcolor="<%= adminColor("tabletop") %>" ><%=pgrptitle%></th>
    <% end if %>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if (oAttrItemList.FAttrResultCount>0) then %>
        <% for j=0 to oAttrItemList.FAttrResultCount-1 %>
            <th width="34" class="stickyhead2" bgcolor="<%= adminColor("tabletop") %>" ><%=oAttrItemList.FAttrGbnList(j).FattribName%></th>
        <% next %>
    <% end if %>
</tr>
</thead>
</div>
<tbody >
<% if oAttrItemList.FResultCount>0 then %>
<%	for i=0 to oAttrItemList.FResultCount - 1 %>
<input type="hidden" name="itemid" value="<%=oAttrItemList.FItemList(i).FItemID%>">
<tr bgcolor="#FFFFFF">
    <td align="center"><input type="checkbox" name="chkix" value="<%=i%>" onClick="AnCheckClick(this);" ></td>
    <td align="center"><img width="50" src="<%=oAttrItemList.FItemList(i).Fsmallimage%>"></td>
    <td align="center"><%=oAttrItemList.FItemList(i).FItemID%></td>
    <td><a href="http://www.10x10.co.kr/<%=oAttrItemList.FItemList(i).FItemID%>" target="_front"><%=oAttrItemList.FItemList(i).Fitemname%></a></td>
    <td align="center"><%=oAttrItemList.FItemList(i).Fmakerid%></td>
    <td align="center"><%=oAttrItemList.FItemList(i).Fcatecode%></td>
    <% if (oAttrItemList.FAttrResultCount>0) then %>
        <% for j=0 to oAttrItemList.FAttrResultCount-1 %>
            <td align="center" >
            <input type="checkbox" name="chkAttrCd<%=i%>" onClick="chgCheckAttr(this,<%=i%>)" value="<%=oAttrItemList.FItemList(i).FAttrValList(j).FattribCd%>" <%=CHKIIF(oAttrItemList.FItemList(i).FAttrValList(j).FisChecked=1,"checked","")%>>
            <%=oAttrItemList.FItemList(i).FAttrValList(j).FisChecked%>
            </td>
        <% next %>
    <% end if %>
</tr>
</tbody>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="<%=6+oAttrItemList.FAttrResultCount%>" align="center">
    <% if oAttrItemList.HasPreScroll then %>
		<a href="javascript:goPage('<%= oAttrItemList.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oAttrItemList.StartScrollPage to oAttrItemList.FScrollCount + oAttrItemList.StartScrollPage - 1 %>
		<% if i>oAttrItemList.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oAttrItemList.HasNextScroll then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
<tr height="25" bgcolor="FFFFFF">
    <td colspan="<%=6+oAttrItemList.FAttrResultCount%>" align="right">
        <input type="button" value="선택상품 속성 수정" onClick="CheckNChangAttr(this)";>
    </td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
    <td colspan="<%=6+oAttrItemList.FAttrResultCount%>" align="center">
    검색결과가 없습니다.
    </td>
</tr>
<% end if %>
</table>

</form>
<br>
</p>
<%
SET oAttrItemList = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
