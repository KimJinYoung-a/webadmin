<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/items/itemAttribMultiCls.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i, j, k
dim cdl, cdm, cds

Dim dispCate : dispCate		= requestCheckvar(request("disp"),15)
Dim subcate : subcate		= requestCheckvar(request("subcate"),5000)
dim page : page 			= requestCheckvar(request("page"),10)
dim itemid : itemid      	= requestCheckvar(request("itemid"),1500)
dim itemname : itemname    	= requestCheckvar(request("itemname"),64)
dim makerid : makerid     	= requestCheckvar(request("makerid"),32)
dim attribDiv : attribDiv  	= requestCheckvar(request("attribDiv"),32)
dim attribDivs : attribDivs	= requestCheckvar(request("attribDivs"),32)
dim attribDivSearch : attribDivSearch  	= requestCheckvar(request("attribDivSearch"),32)
dim groupid : groupid  		= requestCheckvar(request("groupid"),32)

dim itemidArr

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)

if (page="") then page=1

Dim catecodelist
Dim oAttrItemList

if (dispCate<>"") then catecodelist=dispCate

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

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

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 50
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectGroupID		= groupid

oitem.GetItemList


'==============================================================================
dim oAttrib

set oAttrib = new CAttrib

oAttrib.FRectAttribDiv		= attribDiv

oAttrib.GetAttribCdList_V2


'==============================================================================
dim oAttribItem

set oAttribItem = new CAttrib

itemidArr = "-1"
for i=0 to oitem.FresultCount-1
    itemidArr = itemidArr & "," & oitem.FItemList(i).Fitemid
next

oAttribItem.FRectAttribDiv		= attribDiv
oAttribItem.FRectItemid		= itemidArr

oAttribItem.GetAttribCdConnectList_V2

dim isconnected
dim pgrptitle, grpcolcnt : grpcolcnt=0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function researchfrm(){
    document.frm.attribDiv.value = getSelectValues(document.frm.attribDivs);
    document.frm.page.value = 1;
    document.frm.submit();
}

function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

// 변경된 속성 목록
var changeList = '';
function jsSaveChanged() {
    var frm = document.frm;
    var frmchk = document.frmchk;
    var frmAttr = document.frmAttr;
    var changeListArr;
    var changeListStr = '';

    if (changeList == '') {
        alert('변경내역이 없습니다.');
        return;
    }

    changeListArr = changeList.split('|');
    for (var i = 0; i < changeListArr.length; i++) {
        if (changeListArr[i] == '') { continue; }

        if (frmchk.attribCd.length == undefined) {
            // 한개
            if (frmchk.attribCd.value == changeListArr[i]) {
                changeListStr = changeListStr + '|' + changeListArr[i] + ',' + frmchk.attribCd.checked;
            }
        } else {
            // 여러개
            for (var j = 0; j < frmchk.attribCd.length; j++) {
                if (frmchk.attribCd[j].value == changeListArr[i]) {
                    changeListStr = changeListStr + '|' + changeListArr[i] + ',' + frmchk.attribCd[j].checked;
                }
            }
        }
    }

    if (confirm('저장하시겠습니까?')) {
        frmAttr.changeList.value = changeListStr;
        frmAttr.mode.value = 'savechanged';
        frmAttr.submit();
    }
}

function jsSaveChecked(v) {
    changeList = changeList + "|" + v;
}

function getSelectValues(select) {
    var result = [];
    var options = select && select.options;
    var opt;

    for (var i=0, iLen=options.length; i<iLen; i++) {
        opt = options[i];

        if (opt.selected) {
            if (opt.value != '') {
                result.push(opt.value);
            }
        }
    }
    return result;
}
</script>
<form name="frm" method="get" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td height="25" width="80" rowspan="2" bgcolor="#EEEEEE">상품검색</td>
	<td align="left">
		<table border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td style="white-space:nowrap;">브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %> </td>
				<td style="white-space:nowrap;padding-left:5px;">상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"> </td>
				<td style="white-space:nowrap;padding-left:5px;">상품코드:</td>
				<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
			</tr>
			<tr>
				<td  style="white-space:nowrap;">관리<!-- #include virtual="/common/module/categoryselectbox.asp"--> </td>
				<td  style="white-space:nowrap;"  colspan="2">&nbsp;&nbsp;전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"--> </td>
				<td ></td>
			</tr>
		</table>
    </td>
    <td width="50" rowspan="3" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="researchfrm();">
	</td>
</tr>
<tr align="left" bgcolor="#F4F4F4">
    <td>
        업체(그룹코드) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >
    </td>
</tr>
<tr align="center" bgcolor="#F4F4F4">
    <td height="80" width="80" bgcolor="#EEEEEE">속성검색</td>
    <td align="left">
        <% Call drawSelectAttributeDiv("attribDivs", attribDiv, attribDivSearch, "") %>
        <input type="hidden" name="attribDiv" value="<%= attribDiv %>">
        또는
        <input type="text" name="attribDivSearch" value="<%= attribDivSearch %>">
        * 컨트롤키를 누르면 <font color="red">복수의 속성을 선택</font>할 수 있습니다.
    </td>
</tr>
</table>
</form>

<p />

<div align="right">
    <form name="frmAttr" method="POST" action="itemAttrMultiEdt_process.asp" style="margin:0;">
        <input type="hidden" name="mode">
        <input type="hidden" name="changeList" value="">
        <input type="button" class="button" value="변경사항 저장하기" onClick="jsSaveChanged()">
    </form>
</div>

<p />

<form name="frmchk" method="get" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%= oitem.FTotalCount%></b>
			&nbsp;
			페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" class="sticky_top">
		<td width="60">No.</td>
		<td width=50> 이미지</td>
		<td width="100">브랜드ID</td>
		<td>상품명</td>
        <td>속성명</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="16" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="left">
            <%= oitem.FItemList(i).Fmakerid %>
            <% if (oitem.FItemList(i).FfrontMakerid <> "") then %>
            [<%= oitem.FItemList(i).FfrontMakerid %>]
            <% end if %>
        </td>
		<td align="left">
			<% =oitem.FItemList(i).Fitemname %>
		</td>
        <td align="left">
            <% if oAttrib.FresultCount > 0 then %>
			<table align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
                <tr class="a" height="25" bgcolor="#FFFFFF" align="center">
                    <% for j = 0 to oAttrib.FresultCount-1 %>
                    <td width="80"><%= oAttrib.FItemList(j).FattribName %> <%= oAttrib.FItemList(j).FattribNameAdd %></td>
                    <% next %>
                </tr>
                <tr class="a" height="25" bgcolor="#FFFFFF" align="center">
                    <% for j = 0 to oAttrib.FresultCount-1 %>
                    <td >
                        <%
                        isconnected = False
                        for k = 0 to oAttribItem.FresultCount-1
                            if (oAttrib.FItemList(j).FattribCd = oAttribItem.FItemList(k).FattribCd) and (oitem.FItemList(i).Fitemid = oAttribItem.FItemList(k).Fitemid) then
                                isconnected = True
                                exit for
                            end if
                        next
                        %>
                        <input type="checkbox" name="attribCd" onClick="jsSaveChecked(this.value)" value="<%= oitem.FItemList(i).Fitemid %>,<%= oAttrib.FItemList(j).FattribCd %>" <%= CHKIIF(isconnected, "checked", "") %>>
                    </td>
                    <% next %>
                </tr>
            </table>
            <% end if %>
		</td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
			<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if i>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:goPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>
<% end if %>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
