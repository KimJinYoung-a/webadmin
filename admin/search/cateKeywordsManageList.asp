<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/search/dispCateKeywordManageCls.asp" -->
<%
dim dispCate : dispCate = requestCheckvar(request("disp"),18)
''dim catecode, searchKeyword
dim i, page
dim research : research         = request("research")
dim cateusing : cateusing       = request("cateusing")
dim mi_metakey : mi_metakey     = requestCheckvar(request("mi_metakey"),10)
dim mi_searchkey : mi_searchkey = requestCheckvar(request("mi_searchkey"),10)
dim searchKeyword : searchKeyword = requestCheckvar(Trim(request("searchKeyword")),32)
dim metaKeyword   : metaKeyword = requestCheckvar(Trim(request("metaKeyword")),32)

''catecode  = Trim(requestCheckvar(request("catecode"),30))

page = request("page")
if (page="") then page=1
    

'// ============================================================================
dim ocateKeyword

set ocateKeyword = new CDispCateKeywordsMng
ocateKeyword.FPageSize=50
ocateKeyword.FCurrPage = page
ocateKeyword.FRectDispCate = dispCate
ocateKeyword.FRectCateUsing = cateusing
ocateKeyword.FRectMi_metakey = mi_metakey
ocateKeyword.FRectMi_searchkey = mi_searchkey
ocateKeyword.FRectSearchKeyword = searchKeyword
ocateKeyword.FRectMetaKeyword = metaKeyword
ocateKeyword.getDispCateKeywords_CurrentSellitem

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function NextPage(i){
    document.frm.page.value=i;
    document.frm.submit();
}

function fncheckThis(comp,i){
    var valexists = (comp.value.length>0);
    var chkcomp;
    if (valexists){
        if (document.frmSubmit.cksel.length){
            chkcomp = document.frmSubmit.cksel[i];
        }else{
            chkcomp = document.frmSubmit.cksel;
        }
        chkcomp.checked=true;
        AnCheckClick(chkcomp);
    }
}

function SaveCateKeywords(){
    var frmS = document.frmSubmit;
    chkexists = false;
    if (frmS.cksel.length){
        for (var i=0;i<frmS.cksel.length;i++){
            if (frmS.cksel[i].checked){
                chkexists = true;
                break;
            }
        }
    }else{
        chkexists = frmS.cksel.checked;
    }
    
    if (!chkexists){
        alert('������ ī�װ��� ���� �ϼ���.');
        return;
    }
    
    if (confirm('�����Ͻðڽ��ϱ�?')){
        frmS.submit();
    }
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left" height="30" >
			����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
			
			&nbsp;&nbsp;
			ī�װ� ��뿩�� : 
			<select name="cateusing">
			    <option value="">��ü
			    <option value="Y" <%=CHKIIF(cateusing="Y","selected","")%> >���
			    <option value="N" <%=CHKIIF(cateusing="N","selected","")%> >�̻��    
			</select>
			
			<input type="checkbox" name="mi_metakey" value="on" <%=CHKIIF(mi_metakey="on","checked","") %>>��ŸŰ���� ��������
			<input type="checkbox" name="mi_searchkey" value="on" <%=CHKIIF(mi_searchkey="on","checked","") %>>�˻������߰�Ű���� ��������
			
			&nbsp;
			��ŸŰ����/ī�װ���/�˻������߰�Ű����/ī��BoostŰ���� : <input type="text" class="text" name="searchKeyword" value="<%=searchKeyword%>" size="20">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value=" �� �� " onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<p>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			* ���� �Ǹ����� ��ǰ�� ī�װ���
		</td>
		<td align="right">
		    <input type="button" class="button" value="���û�ǰ����" onClick="SaveCateKeywords()">
			&nbsp;
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<p>

<!-- ����Ʈ ���� -->
<form name="frmSubmit" method="post" action="cateKeywords_Process.asp">
<table width="100%" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
	    <td colspan="20">
    		�˻���� : <b><%= ocateKeyword.FTotalcount %></b>
    		&nbsp;
    		������ : <b><%= page %> / <%= ocateKeyword.FTotalPage %></b>
    	</td>
    </tr>
    
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td align="center" height="22" width="20"><input type="checkbox" name="ckall" onClick="fnCheckAll(this.checked,frmSubmit.cksel);"></td>
    	<td width="80" >ī�װ��ڵ�</td>
    	<td width="20">depth</td>
		<td width="250">ī�װ�Full��</td>
		<td width="50">�ǸŻ�ǰ��</td>
		<td width="100">ī�װ���</td>
		<td width="30">��뿩��</td>
		<td width="100">��ŸŰ���� (?)<br>(�޸��� ����)</td>
		<td width="100">�˻������߰�Ű����<br>(�޸��� ����)</td>
		<td width="100">ī�װ�BoostŰ����</td>
	</tr>
	<%
	for i = 0 To ocateKeyword.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
	    <td align="center" height="22" >
	        <input type="checkbox" name="cksel" value="<%= i %>" onClick="AnCheckClick(this);">
	        <input type="hidden" name="catecode" value="<%=ocateKeyword.FItemList(i).FCateCode%>">
	    </td>
		<td align="center" >
			<%= ocateKeyword.FItemList(i).FCateCode %>
		</td>
		<td align="center"><%= ocateKeyword.FItemList(i).FDepth %></td>
		<td align="left">
			<%= replace(ocateKeyword.FItemList(i).FCateFullName,"^^","&gt;&gt;") %>
		</td>
		<td align="center"><%= formatNumber(ocateKeyword.FItemList(i).FSellItemCnt,0) %></td>
			
		
		<td align="center">
			<%= ocateKeyword.FItemList(i).FCateName %>
		</td>
		<td align="center">
		    <%= ocateKeyword.FItemList(i).FUseYN %>
		</td>
		<td align="center">
			<input type="text" name="metakeywords" value="<%= ocateKeyword.FItemList(i).FMetaKeywords %>" size="36" onKeyUp="fncheckThis(this,<%=i%>)">
		</td>
		<td align="center">
			<input type="text" name="searchkeywords" value="<%= ocateKeyword.FItemList(i).FsearchKeywords %>" size="36" onKeyUp="fncheckThis(this,<%=i%>)">
		</td>
		<td align="center">
		    <strong><%= ocateKeyword.FItemList(i).FCateBoostKeyword %></strong>
		</td>   
	</tr>
	<%
	next
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="10">
	<% if (ocateKeyword.FTotalCount <1) then %>
			�˻������ �����ϴ�.
    <% else %>
        <% if ocateKeyword.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ocateKeyword.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocateKeyword.StartScrollPage to ocateKeyword.FScrollCount + ocateKeyword.StartScrollPage - 1 %>
			<% if i>ocateKeyword.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocateKeyword.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	<% end if %>
	    </td>
	</tr>
</table>
</form>
<%
set ocateKeyword = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
