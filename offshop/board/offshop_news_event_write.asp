<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<%
dim shopid
shopid = session("ssBctID")

%>
<script>
function SubmitForm()
{
//alert('�������Դϴ�.');
//return;
        if (document.f.gubun.value == "") {
                alert("�� ������ �����ϼ���.");
                return;
        }
        
		if (document.f.title.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        
        if (document.f.enddate.value == "") {
                alert("�������� �Է��ϼ���.");
                return;
        }
        
        if (confirm('���� �Ͻðڽ��ϱ�?')){
            document.f.submit();
        }
}
</script>
<table  border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0" width="650" class="a">
<form method="post" name="f" action="<%= uploadImgUrl %>/linkweb/offshop/OffshopNewsEvent_process.asp" onsubmit="return false" enctype="multipart/form-data">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<input type="hidden" name="returnUrl" value="http://webadmin.10x10.co.kr/offshop/board/offshop_news_event_list.asp?menupos=567">
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">����</td>
		<td bgcolor="white" style="padding:0">
		    <input type="text" name="shopid" value="<%= shopid %>" size="16" readOnly style="border:1 solid" >
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">������</td>
		<td bgcolor="white" style="padding:0">
			<select name="gubun">
				<option value="">����</option>
				<%=fnOptCommonCode("noticegubun","")%>
			</select>
		</td>
	</tr>	
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">����</td>
		<td bgcolor="white" style="padding:0">
				<input name="title" style="width:450" maxlength="40" style="border:1 solid" value="">
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">����</td>
		<td bgcolor="white" style="padding:0">
				<textarea name="contents" cols="50" rows="15" style="border:1 solid"></textarea>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">÷�λ���</td>
		<td bgcolor="white" style="padding:0">
				<input type="file" name="file1" size="50" class="input_b">
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">������</td>
		<td bgcolor="white" style="padding:0">
				<input type="text" name="enddate" size="10" maxlength="10" style="border:1 solid" value="">
				<a href="javascript:calendarOpen(f.enddate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
				(<%= Left(now(),10) %>)
		</td>
	</tr>
	<tr>
		<td style="padding:0" colspan="2" align="right" bgcolor="white">
			<input type="button" value="Save" onclick="SubmitForm()" style="background-color:#dddddd; height:25; border:1 solid buttonface">&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->