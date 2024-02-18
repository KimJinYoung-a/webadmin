<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/menu/nanishowhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim filename,fso,ofile, buf
filename = "C:/home/cube1010/www" & "/ext/nanishow/choiceofmonth.inc"
set fso = server.CreateObject("scripting.filesystemobject")
set ofile = fso.OpenTextFile (filename, 1, false )

buf = ofile.readAll

ofile.close
set ofile = nothing
set fso = nothing

dim splited
dim imgname, imgname2, eventlink
dim explain
splited = split(buf,vbcrlf)

if UBound(splited)>0 then
	imgname = splited(0)
end if

if UBound(splited)>1 then
	imgname2 = splited(1)
end if

if UBound(splited)>2 then
	eventlink = splited(2)
end if

explain = Trim(replace(buf,imgname + vbcrlf + imgname2+ vbcrlf + eventlink,""))
%>
<script language='javascript'>
function CkNSubmit(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		frm.submit();
	}
}
</script>
	<form name="frm_choicemonth" method="post" action="dochoiceofmonth.asp" onsubmit="CkNSubmit(this); return false;" enctype="multipart/form-data">
	<input type="hidden" name="oldfilename" value="<%= imgname %>">
	<input type="hidden" name="oldfilename2" value="<%= imgname2 %>">
	<table border="1" width="700" bordercolorlight="#808080" cellspacing="0" bordercolordark="#FFFFFF" valign="top">
	  <tr>
	    <td class="a" colspan="2">Choice Of This Month 
	    </td>
	  </tr>
	  <tr>
	  	<td class="a">이미지:</td>
	    <td class="a">
	    	<img src="http://www.10x10.co.kr<%= imgname %>"><br>
	    	<input type="file" name="file1" size="50" value="">
	    </td>
	  </tr>
	  <tr>
	  	<td class="a">설명:</td>
	    <td class="a"><textarea name="explain" rows="4" wrap="hard" cols="30" ><%= explain %></textarea></td>
	  </tr>
	  <tr>
	    <td class="a" colspan="2">Event For You</td>
	  </tr>
	  <tr>
	  	<td class="a">이미지:</td>
	    <td class="a">
	    	<img src="http://www.10x10.co.kr<%= imgname2 %>"><br>
	    	<input type="file" name="file2" size="50" value="">
	    </td>
	  </tr>
	  <tr>
	  	<td class="a">링크:</td>
	    <td class="a"><input type="text" name="eventlink" size="60" value="<%= eventlink %>"></td>
	  </tr>
	  <tr>
	  	<td colspan="2" align="center"><input type="button" value="저장" onclick="CkNSubmit(frm_choicemonth)"></td>
	  </tr>  	
	</form>  
	</table>

<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->