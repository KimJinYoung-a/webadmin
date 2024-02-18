<%
function requestCheckVar(orgval,maxlen)
	requestCheckVar = trim(orgval)
	requestCheckVar = replace(requestCheckVar,"'","")
	requestCheckVar = replace(requestCheckVar,"declare","")
	requestCheckVar = replace(requestCheckVar,"DECLARE","")
	requestCheckVar = replace(requestCheckVar,"Declare","")
	requestCheckVar = Left(requestCheckVar,maxlen)
end function
%>
<script>
  function go(){
    window.frames[0].document.body.innerHTML='<form target="_parent" action="<%=requestCheckVar(Request("urll"),200)%>"></form>';
    window.frames[0].document.forms[0].submit()
  }    
</script>
<iframe onload="window.setTimeout('go()', 99)" src="about:blank" style="visibility:hidden"></iframe>