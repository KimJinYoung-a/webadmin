<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ũ�۷κ� �ǸŴ���ǰ
' History : 2015.10.28 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/makeglob/makeglobCls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->

<%

	Dim currpage '// ���� ������
	Dim pagesize '// ������������
	Dim brandname '// �귣���
	Dim itemname '// ��ǰ��
	Dim itemid '// �������ڵ�
	Dim sellyn '// ��ǰ�Ǹſ���
	Dim limityn '// �����Ǹſ���
	Dim isusing '// ��뿩��
	Dim MakeGlobChkEN '// �����Է¿���
	Dim MakeGlobChkZH '// �߹��Է¿���
	Dim ghidden '// �۷κ� ���迩��
	Dim gsoldout '// �۷κ� ǰ������
	Dim gproductkey '// �۷κ� ��ǰ�ڵ�
	Dim gcheck '// �۷κ� ��Ͽ���
	Dim marginSt '// ������ ���۰�
	Dim marginEd '// ������ ���ᰪ
	Dim sOrgpriceSt '// �ǸŰ� ���۰�
	Dim sOrgpriceEd '// �ǸŰ� ���ᰪ
	Dim baesonggubun '// ��۱���(����, �ٹ�)
	Dim i, dispCate, paramvalue



	currpage = request("page")
	pagesize = 30
	brandname = request("brandname")
	itemname = request("itemname")
	itemid = request("itemid")
	sellyn = request("sellyn")
	limityn = request("limityn")
	isusing = request("isusing")
	ghidden = request("globHiddenYN")
	gsoldout = request("globSoldoutYN")
	gproductkey = request("gproductkey")
	gcheck = request("globCheckYN")
	marginSt = request("marginSt")
	marginEd = request("marginEd")
	sOrgpriceSt = request("sOrgpriceSt")
	sOrgpriceEd = request("sOrgpriceEd")
	MakeGlobChkEN = request("MakeGlobChkEN")
	MakeGlobChkZH = request("MakeGlobChkZH")
	baesonggubun = request("baesonggubun")
	dispCate = request("disp")



	'// �⺻��
	If currpage = "" Then currpage = 1
	If sellyn = "" Then sellyn = "Y"
	If isusing = "" Then isusing = "Y"

	if itemid<>"" then
		dim iA ,arrTemp,arrItemid

		itemid = replace(itemid,chr(10),",")

		arrTemp = Split(itemid,",")

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

	if gproductkey<>"" then
		dim iB ,arrTemp2,arrgproductkey

		arrTemp2 = Split(gproductkey,",")

		iB = 0
		do while iB <= ubound(arrTemp2)
			if Trim(arrTemp2(iB))<>"" and isNumeric(Trim(arrTemp2(iB))) then
				arrgproductkey = arrgproductkey & Trim(arrTemp2(iB)) & ","
			end if
			iB = iB + 1
		loop

		if len(arrgproductkey)>0 then
			gproductkey = left(arrgproductkey,len(arrgproductkey)-1)
		else
			if Not(isNumeric(gproductkey)) then
				gproductkey = ""
			end if
		end if
	end if


	Dim oitem
	set oitem = new CMakeGlobItem
	oitem.Fpagesize = pagesize
	oitem.Fcurrpage = currpage
	oitem.FRectBrandName = brandname
	oitem.FRectCateCode = dispCate
	oitem.FRectItemName = itemname
	oitem.FRectItemId = itemid
	oitem.FRectSellyn = sellyn
	oitem.FRectLimityn = limityn
	oitem.FRectIsUsing = isusing
	oitem.FRectGIsHidden = ghidden
	oitem.FRectGIssoldout = gsoldout
	oitem.FRectGProductKey = gproductkey
	oitem.FRectGIscheck = gcheck
	oitem.FRectMarginSt = marginSt
	oitem.FRectMarginEd = marginEd
	oitem.FRectSorgpriceSt = sOrgpriceSt
	oitem.FRectSorgpriceEd = sOrgpriceEd
	oitem.FRectBaesongGubun = baesonggubun
	oitem.GetMakeGlobItemWaitingList()

	paramvalue = "menupos=3751&page="&currpage&"&reload=ON&disp="&dispcate&"&itemname="&itemname&"&itemid="&itemid&"&sellyn="&sellyn&"&isusing="&isusing&"&limityn="&limityn&"&gproductkey="&gproductkey&"&globHiddenYN="&ghidden&"&globSoldoutYN="&gsoldout&"&globCheckYN="&gcheck&"&brandname="&brandname&"&baesonggubun="&baesonggubun

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

$(document).ready(function(){
    $("#checkall").click(function(){
        if($("#checkall").prop("checked")){
            $("input[name=productcode]").prop("checked",true);
        }else{
            $("input[name=productcode]").prop("checked",false);
        }
    })
})


function fnHiddenProc(val)
{
	var hiddenarrlist='';
	var hiddenalertText='';

	if (val=="Y")
	{
		hiddenalertText = "���õ� ��ǰ�� ����ó�� �Ͻðڽ��ϱ�?";
	}
	else
	{
		hiddenalertText = "���õ� ��ǰ�� ���� �Ͻðڽ��ϱ�?";
	}

	if (!$('input:checkbox[name=productcode]').is(':checked'))
	{
		alert("��ǰ�� �������ּ���.");
		return false;
	}
	else
	{
		if (confirm(hiddenalertText))
		{
			document.globFrm.mode.value="hidden";
			document.globFrm.hiddenvalue.value=val;
			$("input:checkbox[name=productcode]:checked").each(function(){
				if (hiddenarrlist=="")
				{
					hiddenarrlist=$(this).val();
				}
				else
				{
					hiddenarrlist+=','+$(this).val();
				}
			});
			document.globFrm.arrproductcode.value=hiddenarrlist;
			document.globFrm.submit();
		}
		else
		{
			return false;
		}
	}


}

function fnSoldoutProc(val)
{
	var soldarrlist='';
	var soldalertText='';

	if (val=="Y")
	{
		soldalertText = "���õ� ��ǰ�� ǰ��ó�� �Ͻðڽ��ϱ�?";
	}
	else
	{
		soldalertText = "���õ� ��ǰ�� �ǸŰ��� ���·� �����Ͻðڽ��ϱ�?";
	}

	if (!$('input:checkbox[name=productcode]').is(':checked'))
	{
		alert("��ǰ�� �������ּ���.");
		return false;
	}
	else
	{
		if (confirm(soldalertText))
		{
			document.globFrm.mode.value="soldout";
			document.globFrm.soldoutvalue.value=val;
			$("input:checkbox[name=productcode]:checked").each(function(){
				if (soldarrlist=="")
				{
					soldarrlist=$(this).val();
				}
				else
				{
					soldarrlist+=','+$(this).val();
				}
			});
			document.globFrm.arrproductcode.value=soldarrlist;
			document.globFrm.submit();
		}
		else
		{
			return false;
		}
	}


}

function fnProductInsert()
{
	var productarrlist='';
	if (!$('input:checkbox[name=productcode]').is(':checked'))
	{
		alert("��ǰ�� �������ּ���.");
		return false;
	}
	else
	{
		if (confirm('�����Ͻ� ��ǰ�� ���/���� �Ͻðڽ��ϱ�?'))
		{
			document.globFrm.mode.value="product";
			$("input:checkbox[name=productcode]:checked").each(function(){
				if (productarrlist=="")
				{
					productarrlist=$(this).val();
				}
				else
				{
					productarrlist+=','+$(this).val();
				}
			});
			document.globFrm.arrproductcode.value=productarrlist;
			document.globFrm.submit();
		}
		else
		{
			return false;
		}

	}

}


</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �귣�� : 	<input type="text" class="text" name="brandname" value="<%= brandname %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;&nbsp;
		����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		&nbsp;&nbsp;
		* ��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick='NextPage("");'>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* �ٹ����� �Ǹſ���:<% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;&nbsp;
     	* �ٹ����� ��뿩��:<% drawSelectBoxUsingYN "isusing", isusing %>
		&nbsp;&nbsp;
     	* �ٹ����� ��������:<% drawSelectBoxLimitYN "limityn", limityn %>
		&nbsp;&nbsp;
     	* ��۱���: 
		<select name="baesonggubun" class="select" >
			<option value="">��ü</option>
			<option value="tenbae" <% If baesonggubun="tenbae" Then %> selected <% End If %>>�ٹ����ٹ��</option>
			<option value="upbae" <% If baesonggubun="upbae" Then %> selected <% End If %>>��ü���</option>
		</select>
		&nbsp;&nbsp;
		<p/>
     	* ������ : <input type="text" class="text" name="marginSt" value="<%= marginSt %>" size="10" maxlength="4"> ~ <input type="text" class="text" name="marginEd" value="<%= marginEd %>" size="10" maxlength="4">
		&nbsp;&nbsp;
     	* �ǸŰ� : <input type="text" class="text" name="sOrgPriceSt" value="<%= sOrgPriceSt %>" size="10" maxlength="10"> ~ <input type="text" class="text" name="sOrgPriceEd" value="<%= sOrgPriceEd %>" size="10" maxlength="10">
		&nbsp;&nbsp;
		* �ٹ����� ��ǰ�ڵ� :
		<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* �۷κ� ��ǰ�ڵ� :
		<input type="text" class="text" name="gproductkey" value="<%= gproductkey %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
		&nbsp;&nbsp;
		* �۷κ� ���迩��:<% drawSelectBoxGHiddenYN "globHiddenYN", ghidden %>
		&nbsp;&nbsp;
     	* �۷κ� ǰ������:<% drawSelectBoxGsoldoutYN "globSoldoutYN", gsoldout %>
		&nbsp;&nbsp;
     	* �۷κ� ��Ͽ���:<% drawSelectBoxGcheckYN "globCheckYN", gcheck %>
		&nbsp;&nbsp;
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="3" width="50" align="center"><strong>����</strong></td>
		<td><input type="button" value="��ǰ����" onclick="fnHiddenProc('Y');return false;">&nbsp;&nbsp;<input type="button" value="��ǰ����" onclick="fnHiddenProc('N');return false;"></td>
	</tr>
	<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
		<td><input type="button" value="ǰ��ó��" onclick="fnSoldoutProc('Y');return false;">&nbsp;&nbsp;<input type="button" value="�ǸŰ���" onclick="fnSoldoutProc('N');return false;"></td>
	</tr>
	<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
		<td><input type="button" value="��ǰ���/����" onclick="fnProductInsert();return false;"> (������ �̹� ��ϵǾ� �ִ� ��ǰ�� �ֽ������� ����, ���� ��ǰ�� �űԷ� �߰� �˴ϴ�.)</td>
	</tr>
</table>
<br>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				�˻���� : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				������ : <b><%= currpage %> /<%=  oitem.FTotalpage %></b>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" rowspan="2"><input type="checkbox" id="checkall"></td>
	<td width="50" rowspan="2">�̹���</td>
	<td width="100" rowspan="2">�귣���</td>
	<td rowspan="2">��ǰ��</td>
	<td width="60" rowspan="2">��ǰ<br>����</td>
	<td width="60" rowspan="2">���<br>����</td>
	<td colspan="7" width="300"><strong>�ٹ�����</strong></td>
	<td colspan="7" width="120"><strong>����ũ�۷κ�</strong></td>

</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

	<td width="60">��ǰ�ڵ�</td>
	<td width="60">�ǸŰ�</td>
	<td width="60">���԰�</td>
	<td width="60">������</td>
	<td width="30">�Ǹ�<br>����</td>
	<td width="30">ǰ��<br>����</td>
	<td width="30">���<br>����</td>
	<td width="30">����<br>����</td>
	<td width="60">��ǰ�ڵ�</td>
	<td width="30">����<br>����</td>
	<td width="30">ǰ��<br>����</td>
	<td width="60">������Ʈ<br>����</td>
	<td width="60">������Ʈ<br>����</td>
</tr>


<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="19" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" <% If oitem.FItemList(i).FMakeGlobProductKey="" Or isnull(oitem.FItemList(i).FMakeGlobProductKey) Then %> bgcolor="#FFFFA5" <% Else %> bgcolor="#FFFFFF" <% End If %>align="center">

	<td align="center"><input type="checkbox" name="productcode" value="<%= oitem.FItemList(i).Fitemid %>"></td>
	<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
	<td align="left"><%= oitem.FItemList(i).FBrandName %></td>
	<td align="left"><% =oitem.FItemList(i).FitemName %></td>
	<td align="center"><%= FormatNumber((oitem.FItemList(i).FitemWeight/1000),2) %>kg</td>
	<td align="center">
		<%
			If oitem.FItemList(i).FBaesongGubun="M" Or oitem.FItemList(i).FBaesongGubun="W" Then
				Response.write "�ٹ�"
			Else
				Response.write "����"
			End If
		%>
	</td>
	<td>
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">
		<%= oitem.FItemList(i).Fitemid %></a>
	</td>
	<td align="right">
	<%
		Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
		'���ΰ�
'		if oitem.FItemList(i).Fsailyn="Y" then
'			Response.Write "<br><font color=#F08050>(���ǸŰ�)" & FormatNumber(oitem.FItemList(i).FsellCash,0) & "</font>"
'		end if

	%>
	</td>
	<td align="right">
	<%
		Response.Write "" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & ""
	%>
	</td>
	<td align="right">
	<%
		Response.Write "" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice,1) & ""
	%>
	</td>
	<!--td align="center"><%= FormatNumber(oitem.FItemList(i).FbuyCash,0) %></td-->
	<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
	<td align="center">
		<%
			If oitem.FItemList(i).isSoldout Then
				Response.write fnColor("Y", "yn")
			Else
				Response.write fnColor("N", "yn")
			End If
		%>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td><%= oitem.FItemList(i).FMakeGlobProductKey %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).FMakeGlobHidden,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).FMakeGlobSoldout,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).FMakeGlobupdate,"yn") %></td>
	<td align="center">
		<%
			If oitem.FItemList(i).FMakeGlobupdateTime ="1900-01-01" Then
				Response.write ""
			Else
				Response.write oitem.FItemList(i).FMakeGlobupdateTime
			End If
		%>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="19" align="center">
		<% if oitem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(currpage)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

</table>
<form method="post" action="/admin/makeglob/proc.asp" name="globFrm">
	<input type="hidden" name="mode">
	<input type="hidden" name="hiddenvalue">
	<input type="hidden" name="soldoutvalue">
	<input type="hidden" name="arrproductcode">
	<input type="hidden" name="paramvalue" value="<%=tenEnc(paramvalue)%>">
</form>
<% end if %>

<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->