<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이미지 등록처리
' History : 2011.03.16 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	Dim sName,sSpan,strImgUrl, sWidth , sOpt
	sName	= requestCheckVar(Request("sName"),50) 
	sSpan	= requestCheckVar(Request("sSpan"),50)  
	sWidth  = requestCheckVar(Request("sWidth"),10)  
	strImgUrl	= requestCheckVar(Request("sImgUrl"),100) 
	sOpt	= requestCheckVar(Request("sOpt"),1) 

%>
<% If sOpt = "B" Then %>
<script language="javascript">
<!--	 
	var sName, sSpan;
	sName = "<%=sName%>";	
	sSpan = "<%=sSpan%>";
	
	alert("이미지가 등록되었습니다.\n\n이미지 등록후 저장버튼을 눌러야 처리완료됩니다.");
	opener.eval("document.all."+sName).value = "<%=strImgUrl%>";
	opener.eval("document.all."+sSpan+"_bg").style.background = "url(<%=strImgUrl%>)";
	opener.eval("document.all."+sSpan).innerHTML ="<%=strImgUrl%>"+												
		   	"<a href=javascript:jsDelImg('"+sName+"','"+sSpan+"');><img src='/images/icon_delete2.gif' border='0' class='delImg'></a> ";		   	
	opener.eval("document.all."+sSpan).style.display = "";   	
	window.close();
//-->
</script>
<% Else %>
<script language="javascript">
<!--	 
	var sName, sSpan;
	sName = "<%=sName%>";	
	sSpan = "<%=sSpan%>";
	
	alert("이미지가 등록되었습니다.\n\n이미지 등록후 저장버튼을 눌러야 처리완료됩니다.");
	opener.eval("document.all."+sName).value = "<%=strImgUrl%>";		
	opener.eval("document.all."+sSpan).innerHTML ="<img src='<%=strImgUrl%>'"+
			" <%IF sWidth > 400 THEN%>width='400'<%END IF%> >";		   	
	opener.eval("document.all."+sSpan).style.display = "";		   	
	window.close();
//-->
</script>
<% End If %>