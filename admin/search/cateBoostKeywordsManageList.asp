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
dim boostcateusing : boostcateusing       = request("boostcateusing")
dim searchKeyword : searchKeyword = requestCheckvar(Trim(request("searchKeyword")),32)

''catecode  = Trim(requestCheckvar(request("catecode"),30))

page = request("page")
if (page="") then page=1
    

'// ============================================================================
dim ocateKeyword

set ocateKeyword = new CDispCateKeywordsMng
ocateKeyword.FPageSize=50
ocateKeyword.FCurrPage = page
ocateKeyword.FRectDispCate = dispCate
ocateKeyword.FRectBoostKateUsing = boostcateusing
ocateKeyword.FRectSearchKeyword = searchKeyword

ocateKeyword.getDispCateBoostKeywordsList

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

function AddCateBoostKeywords(){
    var frm = document.frmaddkey;
    if (frm.addkeyword.value.length<1){
        alert('Ű���带 �Է����ּ���.');
        frm.addkeyword.focus();
        return;
    }
    
    if ((frm.addcatecode.value.length<1)||(!IsDigit(frm.addcatecode.value))){
        alert('ī�װ��ڵ带 ���ڷ� �Է����ּ���.)');
        frm.addcatecode.focus();
        return;
    }
    
    if (confirm('�߰��Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}


function chgState(addkeyword,addcatecode,edtcateusing){
    var frm = document.frmedtkey;
    frm.addkeyword.value=addkeyword;
    frm.addcatecode.value=addcatecode;
    frm.edtcateusing.value=edtcateusing;
 
    
    if (confirm('�����Ͻðڽ��ϱ�?')){
        frm.submit();
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
			ī�װ�Boost ��뿩�� : 
			<select name="boostcateusing">
			    <option value="">��ü
			    <option value="Y" <%=CHKIIF(boostcateusing="Y","selected","")%> >���
			    <option value="N" <%=CHKIIF(boostcateusing="N","selected","")%> >�̻��    
			</select>
			
			
			&nbsp;
			ī�װ�BoostŰ���� : <input type="text" class="text" name="searchKeyword" value="<%=searchKeyword%>" size="20">
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
<form name="frmaddkey" method="post" action="cateKeywords_Process.asp">
    <input type="hidden" name="mode" value="addcateboostkey">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			* ���� �Ǹ����� ��ǰ�� ī�װ���
		</td>
		<td align="right">
		    Ű����:<input type="text" name="addkeyword" value="" size="10" maxlength="20">
		     | ī�װ��ڵ�:<input type="text" name="addcatecode" value="" size="20" maxlength="18">
		    <input type="button" class="button" value="ī�װ�BoostŰ���� �߰�" onClick="AddCateBoostKeywords()">
			&nbsp;
		</td>
	</tr>
</table>
</form>
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
    	<td align="center" height="22" width="100">Ű����</td>
    	<td width="80" >ī�װ��ڵ�</td>
    	<td width="20">depth</td>
		<td width="250">ī�װ�Full��</td>
		<td width="50">�ǸŻ�ǰ��<br>(��������)</td>
		<td width="100">ī�װ���</td>
		<td width="30">��뿩��<br>(Boost)</td>
		<td width="100">��ŸŰ����<br>(�޸��� ����)</td>
		<td width="100">�˻������߰�Ű����<br>(�޸��� ����)</td>
		<td width="100">�����</td>
		<td width="100">�����</td>
		<td width="50"></td>
	</tr>
	<%
	for i = 0 To ocateKeyword.FResultCount - 1
	%>
	<tr align="center" bgcolor="<%=CHKIIF(ocateKeyword.FItemList(i).Fcateboostkeyusing="N","#CCCCCC","#FFFFFF")%>">
	    <td align="center" height="22" >
	        <%= ocateKeyword.FItemList(i).FCateBoostKeyword %>
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
		    <%= ocateKeyword.FItemList(i).Fcateboostkeyusing %>
		</td>
		<td align="center">
			<%= ocateKeyword.FItemList(i).FMetaKeywords %>
		</td>
		<td align="center">
			<%= ocateKeyword.FItemList(i).FsearchKeywords %>
		</td>
		<td align="center">
			<%= ocateKeyword.FItemList(i).FcateboostkeyRegdate %>
		</td>
		<td align="center">
			<%= ocateKeyword.FItemList(i).Freguserid %>
		</td>
		<td align="center">
		    <% if (ocateKeyword.FItemList(i).Fcateboostkeyusing="N") then %>
		    <input type="button" value="��� ��ȯ" class="button" onClick="chgState('<%=ocateKeyword.FItemList(i).FCateBoostKeyword%>','<%= ocateKeyword.FItemList(i).FCateCode %>','Y')">    
		    <% else %>
		    <input type="button" value="������ ��ȯ" class="button" onClick="chgState('<%=ocateKeyword.FItemList(i).FCateBoostKeyword%>','<%= ocateKeyword.FItemList(i).FCateCode %>','N')">    
		    <% end if %>
		</td>
	</tr>
	<%
	next
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="12">
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

<form name="frmedtkey" method="post" action="cateKeywords_Process.asp">
<input type="hidden" name="mode" value="cateboostkeychg">
<input type="hidden" name="addkeyword" value="">
<input type="hidden" name="addcatecode" value="">
<input type="hidden" name="edtcateusing" value="">
</form>

<%
set ocateKeyword = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
