<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
dim id, makerid
makerid = session("ssBctID")
id      = requestCheckVar(request("id"),20)

dim ibeasong
set ibeasong = new CEventsBeasong
ibeasong.FRectId = id
ibeasong.FRectDeliverMakerid = makerid

if (makerid<>"") and (id<>"") then
    ibeasong.GetOneWinnerItem
end if

if ibeasong.FResultCount<1 then
	response.write "<script>alert('�˻��� ������ �����ϴ�.');</script>"
	response.write "<script>history.back();</script>"
	dbget.close()	:	response.End
end if

dim i
dim hpArr,hp1,hp2,hp3
dim phoneArr,phone1,phone2,phone3

if IsNULL(ibeasong.FOneItem.Freqphone) then ibeasong.FOneItem.Freqphone=""
if IsNULL(ibeasong.FOneItem.Freqhp) then ibeasong.FOneItem.Freqhp=""
if IsNULL(ibeasong.FOneItem.Freqzipcode) then ibeasong.FOneItem.Freqzipcode=""

phoneArr = split(ibeasong.FOneItem.Freqphone,"-")
hpArr = split(ibeasong.FOneItem.Freqhp,"-")

if UBound(hpArr)>=0 then hp1 = hpArr(0)
if UBound(hpArr)>=1 then hp2 = hpArr(1)
if UBound(hpArr)>=2 then hp3 = hpArr(2)

if UBound(phoneArr)>=0 then phone1 = phoneArr(0)
if UBound(phoneArr)>=1 then phone2 = phoneArr(1)
if UBound(phoneArr)>=2 then phone3 = phoneArr(2)
%>

<table width="100%" border="0" cellpadding="0" cellspacing=0 class="a">
  <tr>
    <td align="center">
  	<table width="90%" border="0" cellpadding="0" cellspacing="0" class="a">
  	  <tr height="30">
  	    <td height="2" colspan="2" >* �̺�Ʈ �� ��Ÿ��� ������� </td>
      </tr>
  	  <tr height="2">
  	    <td height="2" colspan="2" bgcolor="#AAAAAA"></td>
      </tr>
  	  <tr>
		<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�̺�Ʈ�� </td>
		<td style="padding-left:7">
		    <%= ibeasong.FOneItem.Fgubunname %>
		</td>
	  </tr>
  	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
		<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��÷��ǰ</td>
		<td style="padding-left:7"><%= ibeasong.FOneItem.FPrizeTitle %></td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
		<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��÷�ڼ���</td>
		<td style="padding-left:7">
		  <%= ibeasong.FOneItem.Fusername %></td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
		<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�����μ���</td>
		<td style="padding-left:7">
		  <%= ibeasong.FOneItem.Freqname %></td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
		<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����ó</td>
		<td class="verdana_s" style="padding-left:7">
		  <%= phone1 %>
		  -
		  <%= phone2 %>
		  -
		  <%= phone3 %>
		</td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
		<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�ڵ���</td>
		<td class="verdana_s" style="padding-left:7">
		  <%= hp1 %>
		  -
		  <%= hp2 %>
		  -
		  <%= hp3 %>
		</td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
		<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">������ �ּ�</td>
		<td class="verdana_s" style="padding:5 0 5 7">
			<%= ibeasong.FOneItem.Freqzipcode %>
			<br>
			<%= ibeasong.FOneItem.Freqaddress1 %>
			&nbsp;<%= ibeasong.FOneItem.Freqaddress2 %>
		</td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
		<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��Ÿ��û����</td>
		<td class="verdana_s" style="padding:5 0 5 7">
		    <%= nl2br(ibeasong.FOneItem.Freqetc) %>
		</td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
      <tr>
		<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����û��</td>
		<td class="verdana_s" style="padding:5 0 5 7">
		<%= ibeasong.FOneItem.FreqDeliverDate %>
		</td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
		<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����</td>
		<td style="padding-left:7">
		    <% if IsNULL(ibeasong.FOneItem.Fsongjangdiv) or (ibeasong.FOneItem.Fsongjangdiv="") then %>
		    
		    <% else %>
		    <% drawSelectBoxDeliverCompany "songjangdiv",ibeasong.FOneItem.Fsongjangdiv %>
		    <% end if %>
		    <%= ibeasong.FOneItem.Fsongjangno %>
		</td>
	  </tr>
	  <tr height="1">
  	    <td height="1" colspan="2" bgcolor="#DDDDDD"></td>
      </tr>
	  <tr>
	    <td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�߼ۻ��� / �����</td>
	    <td style="padding-left:7">
	        <% if ibeasong.FOneItem.Fissended="Y" then %>
	        �߼ۿϷ� <%= ibeasong.FOneItem.Fsenddate %>
	        <% else %>
	        �̹߼�
	        <% end if %>
	        </select>
	    </td>
	  </tr>
	  <tr height="2">
  	    <td height="2" colspan="2" bgcolor="#AAAAAA"></td>
      </tr>
	  <tr height="30">
	    <td colspan="2" align="center"><input type="button" class="button" value=" �� �� " onclick="window.close();"></td>
	  </tr>
    </table>
	 </td>
  </tr>
  </form>
</table>

<%
set ibeasong = Nothing
%>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->