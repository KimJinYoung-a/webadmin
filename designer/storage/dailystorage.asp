<%@ language=vbscript %>
<% option explicit %>
<%
'########################################################
' 2008�� 01�� 23�� �ѿ�� ����
'########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%
dim makerid, onoffgubun, isusing, research, imageon ,soldout_gubun, danjongyn, sellyn, mwdiv, limityn
makerid = session("ssBctID")
onoffgubun  = requestCheckVar(request("onoffgubun"),20)
isusing = requestCheckVar(request("isusing"),20)
research = requestCheckVar(request("research"),20)
imageon = requestCheckVar(request("imageon"),20)
soldout_gubun = requestCheckVar(request("soldout_gubun"),20)
danjongyn = requestCheckVar(request("danjongyn"),20)
sellyn = requestCheckVar(request("sellyn"),20)
mwdiv = requestCheckVar(request("mwdiv"),20)
limityn = requestCheckVar(request("limityn"),20)

dim page
page = requestCheckVar(request("page"),20)
IF page="" Then Page =1

dim i
if onoffgubun="" then onoffgubun="on"
if soldout_gubun="" then soldout_gubun="A"
isusing="on"


dim itemgubun, itemid, itemoption, BasicMonth
itemgubun = requestCheckVar(request("itemgubun"),20)
itemid = requestCheckVar(request("itemid"),20)
itemoption = requestCheckVar(request("itemoption"),20)
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"


%>

<%
dim osummarystockbrand
set osummarystockbrand = new CSummaryItemStock

osummarystockbrand.FRectItemGubun = "10"
osummarystockbrand.FRectMakerid = makerid
osummarystockbrand.FRectOnlyIsUsing = "Y"
osummarystockbrand.frectsoldout_gubun = soldout_gubun
osummarystockbrand.FRectDanjongyn = danjongyn
osummarystockbrand.FRectOnlySellyn = sellyn
osummarystockbrand.FRectMwDiv = mwdiv
osummarystockbrand.FPageSize = 50
osummarystockbrand.FCurrPage = page

osummarystockbrand.FRectLimityn = limityn


if (onoffgubun = "on") then
        osummarystockbrand.GetCurrentStockByOnlineBrandByDesigner  ''2016/01/07 �и�
elseif (onoffgubun = "off") then
        osummarystockbrand.GetCurrentStockByOfflineBrand
end if


dim totsysstock, totavailstock, totrealstock, totjeagosheetstock, totmaystock

%>

<script language='javascript'>

function NextPage(p){
	document.frm.page.value=p;
	document.frm.submit();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<% if (FALSE) then %>
	<!--
	<tr align="center"  bgcolor="#FFFFFF">
		<td align="left">
			��ǰ�ڵ� :
			<input type="text" class="text" name="" value="" size="32"> (��ǥ�� �����Է°���)
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="" value="" size="32" maxlength="32">
			<br>
		</td>
	</tr>
	-->
    <% end if %>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�Ǹ�
			<select class="select" name="sellyn">
		     	<option value="" selected>��ü</option>
		     	<option value="Y" <% if sellyn = "Y" then response.write "selected" %>>�ǸŻ�ǰ</option>
		     	<option value="S" <% if sellyn = "S" then response.write "selected" %>>�Ͻ�ǰ��</option>
		     	<option value="N" <% if sellyn = "N" then response.write "selected" %>>ǰ����ǰ</option>
	     	</select>
	     	&nbsp;
			ǰ��
			<select class="select" name="soldout_gubun">
	        	<option value="">��ü</option>
	        	<option value="Y" <% if soldout_gubun = "Y" then response.write "selected" %>>ǰ����ǰ</option>
	        	<option value="N" <% if soldout_gubun = "N" then response.write "selected" %>>�ǸŻ�ǰ</option>
	        </select>
	        &nbsp;
	        ����
	       <select class="select" name="danjongyn">
	            <option value="">��ü</option>
	            <option value="SN" <% if (danjongyn = "SN") then response.write "selected" end if %>>������</option>
	            <option value="Y" <% if (danjongyn = "Y") then response.write "selected" end if %>>����</option>
	            <option value="M" <% if (danjongyn = "M") then response.write "selected" end if %>>MDǰ��</option>
	            <option value="S" <% if (danjongyn = "S") then response.write "selected" end if %>>�Ͻ�ǰ��</option>
            </select>
            &nbsp;
            ����
			<select class="select" name="limityn">
		     	<option value="" selected>��ü</option>
		     	<option value="N" <% if (limityn = "N") then response.write "selected" end if %>>������</option>
		     	<option value="Y" <% if (limityn = "Y") then response.write "selected" end if %>>����</option>
		     	<option value="Y0" <% if (limityn = "Y0") then response.write "selected" end if %>>����(0)</option>
	     	</select>
	     	&nbsp;
            �ŷ�����
			<select class="select" name="mwdiv">
		     	<option value="" selected>��ü</option>
		     	<option value="MW" <% if (mwdiv = "MW") then response.write "selected" end if %>>����+Ư��</option>
		     	<option value="M" <% if (mwdiv = "M") then response.write "selected" end if %>>����</option>
		     	<option value="W" <% if (mwdiv = "W") then response.write "selected" end if %>>Ư��</option>
		     	<option value="U" <% if (mwdiv = "U") then response.write "selected" end if %>>��ü���</option>
	     	</select>
	     	&nbsp;
	        <input type=checkbox name=imageon <% if imageon="on" then response.write "checked" %> >�̹���ǥ��
	        <% if (FALSE) then %>
	        <!--
        	<input type=radio name="onoffgubun" value="on" <% if onoffgubun="on" then response.write "checked" %> >ON��ǰ
        	<input type=radio name="onoffgubun" value="off" <% if onoffgubun="off" then response.write "checked" %> >OFF��ǰ
        	<input type=checkbox name="isusing" value="on" <% if isusing="on" then response.write "checked" %> >����ǰ��
			-->
		    <% end if %>
	     </td>
	     <td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	 </tr>
	 </form>
</table>

<p>



<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= FormatNumber(osummarystockbrand.FTotalCount,0) %></b>
			&nbsp;
			������ : <b><%= page %> / <%= osummarystockbrand.FTotalpage %></b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="40">��ǰID</td>
    	<% if imageon="on" then %>
    	<td width="50">�̹���</td>
    	<% end if %>
    	<td>��ǰ��<font color="blue">[�ɼ�]</font></td>
    	<td width="30">�ŷ�<br>����</td>
    	<td width="35">��<br>�԰�/<br>��ǰ</td>
    	<td width="35">ON��<br>�Ǹ�/<br>��ǰ</td>
        <td width="35">OFF��<br>���/<br>��ǰ</td>
        <td width="30">��Ÿ<br>���/<br>��ǰ</td>
        <td width="30">CS<br>���<br>��ǰ</td>
        <td width="40" bgcolor="F4F4F4">�ý���<br>�����</td>
        <td width="30">��<br>�ҷ�</td>        
        <td width="35">��<br>�ǻ�<br>����</td>
        <td width="40" bgcolor="F4F4F4">�ǻ�<br>���</td>
        <td width="40">����<br>�����<br>����</td>
        <td width="40" bgcolor="F4F4F4">����<br>���</td>
        
        <td width="30">�Ǹ�<br>����</td>
        <td width="40">����<br>����</td>
        <% if (FALSE) then %>
        <!--
        <td width="30">���<br>����</td>

        <td width="40" bgcolor="F4F4F4">�ý���<br>��ȿ<br>���</td>
        <td width="34">ON<br>��ǰ<br>�غ�</td>
        <td width="34">OFF<br>��ǰ<br>�غ�</td>
        <td width="34" bgcolor="F4F4F4">���<br>�ľ�<br>���</td>
        <td width="34">ON<br>����<br>�Ϸ�</td>
        <td width="34">ON<br>�ֹ�<br>����</td>
        <td width="34">OFF<br>�ֹ�<br>����</td>
        <td width="34" bgcolor="F4F4F4">����<br>���</td>
        -->
        <% end if %>
		<td width="30">ǰ��<br>����</td>
		<td width="30">����<br>����</td>

    </tr>
<% for i=0 to osummarystockbrand.FResultCount - 1 %>
<%
totsysstock	= totsysstock + osummarystockbrand.FItemList(i).Ftotsysstock
totavailstock = totavailstock + osummarystockbrand.FItemList(i).Favailsysstock
totrealstock = totrealstock + osummarystockbrand.FItemList(i).Frealstock
totjeagosheetstock = totjeagosheetstock + osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno
totmaystock = totmaystock + osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno + osummarystockbrand.FItemList(i).Fipkumdiv4 + osummarystockbrand.FItemList(i).Fipkumdiv2 + osummarystockbrand.FItemList(i).Foffjupno

%>
	<% if osummarystockbrand.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align=center>
    <% else %>
    <tr bgcolor="#EEEEEE" align=center>
    <% end if %>
    	<td><%= osummarystockbrand.FItemList(i).Fitemid %></td>
    	<% if imageon="on" then %>
    	<td><img src="<%= osummarystockbrand.FItemList(i).Fimgsmall %>" width=50 height=50> </td>
    	<% end if %>
    	<td align="left">
    		<%= osummarystockbrand.FItemList(i).Fitemname %>
    		<% if osummarystockbrand.FItemList(i).FitemoptionName <>"" then %>
    		<font color="blue">[<%= osummarystockbrand.FItemList(i).FitemoptionName %>]<font color="blue">
    		<% end if %>
    	</td>
        <td><font color="<%= mwdivColor(osummarystockbrand.FItemList(i).Fmwdiv) %>"><%= osummarystockbrand.FItemList(i).GetMwDivName %></font></td>
    	<td><%= osummarystockbrand.FItemList(i).Ftotipgono %></td>
    	<td><%= -1*osummarystockbrand.FItemList(i).Ftotsellno %></td>
    	<td><%= osummarystockbrand.FItemList(i).Foffchulgono + osummarystockbrand.FItemList(i).Foffrechulgono %></td>
    	<td><%= osummarystockbrand.FItemList(i).Fetcchulgono + osummarystockbrand.FItemList(i).Fetcrechulgono %></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Ftotsysstock %></b></td>
    	<td><%= osummarystockbrand.FItemList(i).Ferrbaditemno %></td>
    	<td><%= osummarystockbrand.FItemList(i).Ferrrealcheckno %></td>
    	<td bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Frealstock %></td>
    	<td><font color="#AAAAAA"><%= osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno + osummarystockbrand.FItemList(i).Fipkumdiv4 + osummarystockbrand.FItemList(i).Fipkumdiv2 + osummarystockbrand.FItemList(i).Foffjupno %></font>
    	<td bgcolor="F4F4F4"><b><font color="#AAAAAA"><%= osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno + osummarystockbrand.FItemList(i).Fipkumdiv4 + osummarystockbrand.FItemList(i).Fipkumdiv2 + osummarystockbrand.FItemList(i).Foffjupno %></font></b></td>

        <td><font color="<%= ynColor(osummarystockbrand.FItemList(i).Fsellyn) %>"><%= osummarystockbrand.FItemList(i).Fsellyn %></font></td>
        <td>
        	<% if osummarystockbrand.FItemList(i).Flimityn="Y" then %>
        	<font color="<%= ynColor(osummarystockbrand.FItemList(i).Flimityn) %>"><%= osummarystockbrand.FItemList(i).Flimityn %></font>
        	(<%= osummarystockbrand.FItemList(i).GetLimitStr %>)
        	<% end if %>  	
        </td>
    	<% if (FALSE) then %>
    	<!--
    	<td><font color="<%= ynColor(osummarystockbrand.FItemList(i).FIsUsing) %>"><%= osummarystockbrand.FItemList(i).FIsUsing %></font></td>

    	<td bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Favailsysstock %></b></td>
    	<td><%= osummarystockbrand.FItemList(i).Fipkumdiv5 %></td>
    	<td><%= osummarystockbrand.FItemList(i).Foffconfirmno %></td>
    	<td bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno %></b></td>
    	<td><%= osummarystockbrand.FItemList(i).Fipkumdiv4 %></td>
    	<td><%= osummarystockbrand.FItemList(i).Fipkumdiv2 %></td>
    	<td><%= osummarystockbrand.FItemList(i).Foffjupno %></td>
    	-->
        <% end if %>
    	<td>
    		<% if osummarystockbrand.FItemList(i).IsSoldOut then %>
    		<font color=red>ǰ��</font>
    		<% end if %>
    	</td>
    	<td>
    		<% if osummarystockbrand.FItemList(i).Fdanjongyn="Y" then %>
			<font color="#33CC33">����</font><br>
			<% elseif osummarystockbrand.FItemList(i).Fdanjongyn="M" then %>
			<font color="#CC3333">MD<br>ǰ��</font>
			<% elseif osummarystockbrand.FItemList(i).Fdanjongyn="S" then %>
			<font color="#3333CC">�Ͻ�<br>ǰ��</font><br>
			<% end if %>
    	</td>
    	
    	
    </tr>
<% next %>
<% if (FALSE) then %>
<!--
	<tr align=center bgcolor="#FFFFFF">
		<td colspan=8></td>
		<td><%= totsysstock %></td>
		<td ></td>
		<td><%= totavailstock %></td>
		<td ></td>
		<td><%= totrealstock %></td>
		<td colspan=2></td>
		<td><%= totjeagosheetstock %></td>
		<td colspan=3></td>
		<td><%= totmaystock %></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
	</tr>
-->
<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if osummarystockbrand.HasPreScroll then %>
			<a href="javascript:NextPage('<%= osummarystockbrand.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + osummarystockbrand.StartScrollPage to osummarystockbrand.FScrollCount + osummarystockbrand.StartScrollPage - 1 %>
    			<% if i>osummarystockbrand.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if osummarystockbrand.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>



<%
set osummarystockbrand = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->