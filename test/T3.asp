<%
''411782


dim a : a = 850*5/100
''a = 16150*5/100

response.write CLNG(41.5)
response.write "<br>"
response.write CLNG(42.5)
response.write "<br>"
response.write CLNG(43.5)
response.write "<br>"
response.write CLNG(27.5)
response.write "<br>"
response.write CLNG(a)
response.write "<br>"
response.write a
response.write "<br>"
response.write 850-a
response.write "<br>"
response.write 850-CLNG(a)
response.write "<br>"
response.write CLNG(850-a)

%>

<script language='javascript'>
var a=850*5/100;
alert(a);
alert(Math.round(a));
</script>