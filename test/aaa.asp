<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

'// 한글파일임
%>
<html>
	<head>
		<style>
body {
    height: 100%;
}

#page_container {
    bottom: 0;
    left: 0;
    overflow: auto;
    position: absolute;
    right: 0;
    top: 0;
}

#menu {
    position: fixed;
    top: 30px;
    left: 10px;
}
		</style>
	</head>
<body>
<ul id="menu">
    <li><a href="#">Item 1</a></li>
    <li><a href="#">Item 2</a></li>
    <li><a href="#">Item 3</a></li>
</ul>
<div id="page_container">
    My content is here<br />
	My content is here<br />
	My content is here<br />
	My content is here
</div>
</body>
</html>
