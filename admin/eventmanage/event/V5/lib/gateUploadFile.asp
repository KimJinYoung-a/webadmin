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
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<% If sOpt = "B" Then %>
<script language="javascript">
	window.document.domain = "10x10.co.kr";
	alert("이미지가 등록되었습니다.\n\n이미지 등록후 저장버튼을 눌러야 처리완료됩니다.");
	$("input[name='<%=sName%>']",opener.document).val("<%=strImgUrl%>");
	$("#<%=sSpan%>",opener.document).css("background":"url(<%=strImgUrl%>)");
	$("#<%=sSpan%>",opener.document).html("<%=strImgUrl%><button class='btn4 btnGrey1 lMar05' onClick=\"jsDelImg('<%=sName%>','<%=sSpan%>');return false;\">삭제</button>");
	$("#<%=sSpan%>",opener.document).show();
	window.close();
</script>
<% ElseIf sOpt = "P" Then '// 가격연동 PC 수기이미지 등록 %>
<script language="javascript">
	window.document.domain = "10x10.co.kr";
	alert("이미지가 등록되었습니다.\n\n이미지 등록후 저장버튼을 눌러야 처리완료됩니다!");
	$("input[name='<%=sName%>']",opener.document).val("<%=strImgUrl%>");
	$("#<%=sSpan%>",opener.document).html("<img src='<%=strImgUrl%>'<%IF sWidth > 400 THEN%> width='400'<%END IF%> onclick=\"jsPcSetImg('<%=sSpan%>','','')\"><button class='btn4 btnGrey1 lMar05' onClick=\"jsDelImg('<%=sName%>','<%=sSpan%>');return false;\">삭제</button>");
	$("#<%=sSpan%>",opener.document).show();
	window.close();
</script>
<% ElseIf sOpt = "Q" Then '// 가격연동 배경이미지 등록 %>
<script language="javascript">
	window.document.domain = "10x10.co.kr";
	alert("이미지가 등록되었습니다.\n\n이미지 등록후 저장버튼을 눌러야 처리완료됩니다!");
	const parent = window.opener.document;
	const parent_sspan = parent.getElementById('<%=sSpan%>');
	parent.querySelector('input[name=<%=sName%>]').value = '<%=strImgUrl%>';
	parent_sspan.innerHTML = `<img src="<%=strImgUrl%>" <%IF sWidth > 400 THEN%> width="400"<%END IF%> width="30%">`;
	parent_sspan.style.display = 'block';

	if( parent_sspan.parentElement.querySelector('.deleteBtn') === null ) {
		$(parent_sspan).before(`<button type="button" class="btn4 btnGrey1 lMar05 deleteBtn" onClick="jsItemDelImg();return false;">삭제</button>`);
	}
	window.close();
</script>
<% Else %>
<script language="javascript">
	window.document.domain = "10x10.co.kr";
	alert("이미지가 등록되었습니다.\n\n이미지 등록후 저장버튼을 눌러야 처리완료됩니다!!");
	$("input[name='<%=sName%>']",opener.document).val("<%=strImgUrl%>");
	$("#<%=sSpan%>",opener.document).html("<img src='<%=strImgUrl%>'<%IF sWidth > 400 THEN%> width='400'<%END IF%>><button class='btn4 btnGrey1 lMar05' onClick=\"jsDelImg('<%=sName%>','<%=sSpan%>');return false;\">삭제</button>");
	$("#<%=sSpan%>",opener.document).show();
	window.close();
</script>
<% End If %>