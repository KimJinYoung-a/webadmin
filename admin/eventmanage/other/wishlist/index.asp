<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventOtherCls_wishlist.asp"-->

<% 
 Dim clsWonderday
 Dim iTotCnt, arrList,intLoop
 Dim iPageSize, iCurrentpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 	
'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
iCurrentpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
IF iCurrentpage = "" THEN		iCurrentpage = 1
iPageSize = 20		'�� �������� �������� ���� ��, front�� �����ϰ�
iPerCnt = 10		'�������� ������ ����
	
dim price_s, price_e, incnt, grade, allcate
price_s = NullFillWith(request("price_s"),"")
price_e = NullFillWith(request("price_e"),"")
incnt = NullFillWith(request("incnt"),"")
grade = NullFillWith(request("grade"),"")
allcate = NullFillWith(request("allcate"),"")

dim oeventuserlist , i

	set oeventuserlist = new CWishList
 	oeventuserlist.FCPage = iCurrentpage	'����������
	oeventuserlist.FPSize = iPageSize '���������� ���̴� ���ڵ尹��
	oeventuserlist.FPriceS = price_s
	oeventuserlist.FPriceE = price_e
	oeventuserlist.FInCnt = incnt
	oeventuserlist.FGrade = grade
	oeventuserlist.FAllCate = allcate
	arrList = oeventuserlist.fnGetWishList
	iTotCnt = oeventuserlist.FTotCnt	'��ü ������  ��

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
%>

<script language="javascript">
function excel(seachbox,eventbox)
{
	location.href = 'event_user_list_excel.asp';
}

function jsGoPage(iP){
	document.frmpage.iC.value = iP;
	document.frmpage.submit();
}
	
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			&nbsp;���� ����: <input type="text" name="price_s" value="<%=price_s%>" size="7"> �� �̻� &nbsp;~&nbsp;  <input type="text" name="price_e" value="<%=price_e%>" size="7"> �� ����
			&nbsp;&nbsp;�� <b>���ڷθ� �Է��ϼ���.</b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			ȸ�����: 
			<select name="grade">
				<option value="">-����-</option>
				<option value="5" <% If grade = "5" Then %>selected<% End IF %>>Orange</option>
				<option value="0" <% If grade = "0" Then %>selected<% End IF %>>Yellow</option>
				<option value="1" <% If grade = "1" Then %>selected<% End IF %>>Green</option>
				<option value="2" <% If grade = "2" Then %>selected<% End IF %>>Blue</option>
				<option value="3" <% If grade = "3" Then %>selected<% End IF %>>VIP</option>
				<option value="4" <% If grade = "4" Then %>selected<% End IF %>>Mania</option>
				<option value="7" <% If grade = "7" Then %>selected<% End IF %>>Staff</option>
				<option value="8" <% If grade = "8" Then %>selected<% End IF %>>Friends</option>
			</select>
		</td>
		
		<td rowspan="2" width="70" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" style="height:30;" value="��  ��" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td align="left">
			&nbsp;���� ��ǰ ����: <input type="text" name="incnt" value="<%=incnt%>" size="2"> �� �̻�
			&nbsp;&nbsp;&nbsp;�� <b>���ڷθ� �Է��ϼ���.</b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="allcate" value="o" <% If allcate = "o" Then %>checked<% End IF %>>
			��� ī�װ��� ����
		</td>
	</tr>
</form>	
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">
			<input type="button" name="excelbox" value="��ü�����Ϳ�������" class="button" onclick="excel('');">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% IF isArray(arrList) THEN %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="6">
			�˻���� : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td align="center">���̵�</td>
		<td align="center">��������</td>
		<td align="center">����</td>
		<td align="center">ȸ�����</td>
		<!--<td align="center">�����</td>//-->
    </tr>
    
	<% For intLoop =0 To UBound(arrList,2) %>
    	<tr align="center" bgcolor="#FFFFFF" height="30">
			<td align="center"><%=arrList(0,intLoop)%></td>
			<td align="center"><%=arrList(1,intLoop)%> ��&nbsp;&nbsp;&nbsp;
				<a href="#" onCLick="javascript:window.open('pop_item.asp?userid=<%=arrList(0,intLoop)%>&fidx=<%=arrList(4,intLoop)%>','wishpop','width=700,height=527, scrollbars=yes');">[Ȯ���ϱ�]</a>
			</td>
			<td align="center"><%=FormatNumber(arrList(2,intLoop),0)%> ��</td>
			<td align="center"><%=UserGrade(arrList(3,intLoop))%></td>
    	</tr>   
	<% next %>
	
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="6" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>

<tr>
	<td colspan="6">
	<!-- ����¡ó�� -->
	<%
	iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1
	
	If (iCurrentpage mod iPerCnt) = 0 Then
		iEndPage = iCurrentpage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frmpage" method="post" action="index.asp">
	<input type="hidden" name="iC" value="<%=iCurrentpage%>">
	<input type="hidden" name="price_s" value="<%=price_s%>">
	<input type="hidden" name="price_e" value="<%=price_e%>">
	<input type="hidden" name="incnt" value="<%=incnt%>">
	<input type="hidden" name="grade" value="<%=grade%>">
	<input type="hidden" name="allcate" value="<%=allcate%>">
	    <tr valign="bottom" height="25">        
	        <td valign="bottom" align="center">
	         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrentpage) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
	        </td>        
	    </tr>    
	    </form>
	</table>
	</td>
</tr>
			
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

