<%@ language=vbscript %>
<% option explicit %>

<%
dim idx,orderserial
idx=request("idx")
if idx="" then idx=0

orderserial=request("orderserial")
%>

<script language="javascript">
function CheckStrLen(maxlen)
	{
	var temp; //들어오는 문자값...
	var msglen;
	msglen = 0;
	var value= document.msgfrm.msg.value;

	L = document.msgfrm.msg.value.length;
	tmpstr = "" ;

	if (L == 0)	{
		value = 0;
		nbytes.innerHTML=msglen + "/80Bytes";
	}
	else	{
		for(k=0;k<L;k++){
			temp =value.charAt(k);

			if (escape(temp).length > 4)
				msglen += 2;
			else
				msglen++;

			if(msglen > 80)	{
				alert("총 80Byte까지 보내실수 있습니다.");
				document.msgfrm.msg.value= tmpstr;
				bytes.innerHTML=msglen + "/80Bytes";
				break;
		    } 
		    else{
				tmpstr += temp;
				nbytes.innerHTML=msglen + "/80Bytes";
			}
		}
	}
}

</script>

<body onload="document.msgfrm.msg.focus();">
<div align="center">
<table border="1" cellspacing="0" cellpadding="1">
<tr>
<td>
<table border="0" cellpadding="0" cellspacing="0">
<form name="msgfrm" method="post" action="lecture_sendmsg.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="orderserial" value="<%=orderserial %>">
	<tr>
		<td colspan="2" align="center">메시지를 입력하세요.</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><textarea name="msg" cols="14" rows="7" onChange="CheckStrLen('40')" onKeyUp="CheckStrLen('40')" style="overflow:auto" style="OVERFLOW-Y: hidden"></textarea></td>
	</tr>
	<tr>
		<td valign="top"><input type="submit" value="확인"></td><td><div id="nbytes" align="right">0/80Bytes</div><br></td>
	</tr>
</form>
</table>
</td>
</tr>
</table>
</div>
</body>

