
<%
'-------------------------����clearHTMLCode��������������ݿ�������html��¼-------------
function clearHTMLCode(art_content) 
dim reg 
set reg = new RegExp 
reg.Pattern = "<[^>]*>" 
reg.Global = true 
clearHTMLCode = reg.Replace(art_content, "") 
end Function 
%> 

<!--------------------------------- ˢ��ҳ�滺��  --------------------------------->
<%    Response.Expires = -1   
      Response.ExpiresAbsolute   =   Now()-1   
      Response.cachecontrol   =   "no-cache"%>
	  
<html>
<!-- #include file="conn.asp" --><head>


<!--------------------------------- ����  --------------------------------->
  <title>���԰�</title>    
  <meta http-equiv="Content-Type" content="text/html; charset=gb2312" /><style type="text/css">
	
<!--------------------------------- CSS��ʽ��ʼ  --------------------------------->
<!--
body,td,th {
	font-size: 15px;
	color: #000000;
}
.STYLE2 {color: #000000}
.STYLE4 {color: #FF0000}
.STYLE6 {color: #237DCF}
body {
	background-color: #F4FAFF;
}
.STYLE9 {font-size: 15px}
a:link {
	color: #237DCF;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #237DCF;
}
a:hover {
	text-decoration: none;
	color: #237DCF;
}
a:active {
	text-decoration: none;
	color: #237DCF;
}
body,td,a {  font-size: 12pt; color: #330000; text-decoration: none}
.aa {filter:alpha(opacity=90)}
.STYLE10 {color: #00FFFF; }
.STYLE12 {color: #237DCF; font-weight: bold; }
.STYLE13 {color: #FF0000; font-weight: bold; }
--> 

<!--------------------------------- CSS��ʽ����  --------------------------------->




<!---------------------------------  ��񱳾���Ч��js  --------------------------------->
  </style>
  <SCRIPT language=JavaScript>
function Cbg(obj, objColor)
{
obj.style.backgroundColor=objColor
}
</SCRIPT>
</head> 

<!---------------------------------  HTML������  --------------------------------->
<body bgcolor="#fef4d9"  onmousemove="move()" onmouseup="down=false">
<div style="position:absolute; left:378px; top:226px; z-index:1; solid;cursor:move; width: 204px; height: 80px; visibility: hidden;" id=plane1 onMouseDown="down1(this)" class="aa">
<table cellpadding="0" border="0" cellspacing="1" width="210" height="85" bgcolor="#237DCF" >
<tr><td height="18" background="bg.gif" >
<div align="right">�����½> <a href="javascript:" class="STYLE10" onClick="clase(1)">====</a></div>
</td></tr><tr>
<td bgcolor="#E6F1FB"><form id="form1" name="form1" method="post" action="login.asp">
  <table width="201" height="57" border="0" align="left">
    <tr>
      <td width="10" rowspan="2">&nbsp;</td>
      <td height="18" colspan="2"><label>Pass:
          <input name="pass" type="password" size="15" />
  </label></td>
      </tr>
    <tr>
      <td width="137"><div align="right"><input type="submit" name="Submit" value="��½" />
        
      </div>        </td>
      <td width="40">&nbsp;</td>
    </tr>
  </table>
  
 </form>
</td>
</tr></table></div>

<!--------------------------------- ��½��,��Ĳ��� --------------------------------->
<script>
var over=false,down=false,divleft,divtop,n;
function clase(x){document.all['plane'+x].style.visibility='hidden'}
function down1(m){
n=m;down=true;divleft=event.clientX-parseInt(m.style.left);divtop=event.clientY-parseInt(m.style.top)}
function move(){if(down){n.style.left=event.clientX-divleft;n.style.top=event.clientY-divtop;}}
</script>


  <table border="1" align="center" cellspacing="1"  bordercolor="#E6F1FB" bgcolor="#E6F1FB">
    <tr>
      <td height="16" colspan="5" bgcolor="#F4FAFF" ><div align="right">


<!-----------------------------  ��֤����Ա�Ƿ��½,�����½�������"�˳�����",�������"����Ա��½"  ------------------------------>
	  	  <%
if session("pass")<>""then %>
 <a href="quit.asp"> <strong>�˳�����</strong></a>
	  <% else %>
	 
	  
	  	  <a href="javascript:" onClick="plane1.style.visibility='visible'">
	  <strong>����Ա��½</strong>
	  </a>
<% end if %>
	  </div>
<!---------------------------------  ��֤����Ա��½��֤���� --------------------------------->
	  
	  </td>
    </tr>
    <tr>
      <td height="20" colspan="5" background="bg.gif" bgcolor="#B9D8F3" ><span class="STYLE6"><strong>����Ҫ����</strong></span></td>
    </tr>
    <tr width=300px>
      <td width="55" ><span class="STYLE6">&nbsp;�ǳ�:</span></td>
      <td  width="292" height="14" >
	          <p>
	  	 <!-- �����ύ�� -->
		<form id="form1" name="form1" method="post" action="add.asp">
      <input name="name" type="text" id="name" /></td>
      <td width="166" >        <div align="right">
          <input name="submit" type="submit" value="   ��������   "/>
          </div></td>
    </tr>
    <tr  width=300px>
      <td valign="top" ><span class="STYLE6">��д����:</span></td>
      <td height="15" colspan="2" bgcolor="#F4FAFF" ><textarea name="ly" cols="63" rows="5" class="inputinput" id="ly"></textarea></td>
    </tr>
</table>
</form>

<!---------------------------------  �������ݲ�ѯ��ʼ  --------------------------------->
<%
       set rs=server.CreateObject("adodb.recordset")
       sql="select * from data order by id DESC"
       rs.open sql,conn,1,1
	   
'----------------------------------------ҳ��--------------------------------------------
  page=request.QueryString("page")
   if IsNumeric(page) then
            page=cint(page)
            if page<1 then page=1 
         else 
            page=1 
         end if
  everypage=5
  rs.pagesize=everypage
  if rs.bof and rs.eof then
  
response.write "<BR> <p align='center' class='STYLE3'>���ݿ����޼�¼..." & allrows & "</p>"
response.end
  else
  page_count=rs.pagecount
  rs.AbsolutePage=page
  do while not rs.eof and j<rs.pagesize
  
'-----------------------------------------------------------------------------------------------
%>
	 	 <% ly=rs("ly")  %>
		 <% hf=rs("hf")  %>
		  <% id=rs("id")  
'-----------------------------------------------------------------------------------------------
%>

<!---------------------------------  �������ݱ��ʼ  --------------------------------->
  <table border="0" cellpadding="3" cellspacing="1" width="542" align="center" style="background-color: #b9d8f3;">
    <tr>
      <td background="bg.gif" bgcolor="#B9D8F3" ><span class="STYLE6"><strong><strong><%=rs("id")%></strong> ¥</strong>:�� <span class="STYLE2"><%=clearHTMLCode(rs("name"))%></span> ��</span></td>
	  
      <td background="bg.gif" bgcolor="#B9D8F3" >
	  
	  <%
if session("pass")<>""then %>
 <div align="right"><span class="STYLE12"><a href="del.asp?id=<%=rs("id")%>">ɾ��</a></span></div>
<% else %>
	 
	  <% end if %>	  </td>
    </tr>
    <tr bgcolor='#F4FAFF'>
      <td width="487"onmouseover="Cbg(this, 'ffffff')" onMouseOut="Cbg(this, '#F4FAFF')" ><div align="left"><span class="STYLE2"><%=clearHTMLCode(rs("ly"))%></span></div>
          <br />
          <div align="right"><span class="STYLE6"><%=rs("time")%> ����</span></div></td>
		  
 <!--------------------------------- �����ظ��ύ�������ύ��ť��JS����  --------------------------------->
		  <SCRIPT type=text/javascript 
      src="jquery.js"></SCRIPT>
	   <SCRIPT type=text/javascript>

					
					function Reply<%=rs("id")%>()
					{

						$("#Reply<%=rs("id")%>").slideToggle('slow', function() {
							window.scrollBy(0,0);
						});
					}
					</SCRIPT>
					
					
      <td width="40" height="30" valign="bottom" bgcolor="#F4FAFF"onmouseover="Cbg(this, 'ffffff')" onMouseOut="Cbg(this, '#F4FAFF')" >
	  
<!---------------------------------  ��֤����Ա�Ƿ��½,�����½����ʾ"�ظ�"  --------------------------------->
	  <%
if session("pass")<>""then %>
	  <div align="right"><span class="STYLE12"><a href="javascript:Reply<%=rs("id")%>()">�ظ�</a></span></div>
	  
	  <% else %>
	 
	  <% end if %>
	  
	  </td>
    </tr>

    <tr>
<!---------------------------------  �ظ��ύ��  --------------------------------->
	<form name="form2" method="post" action="Reply.asp?id=<%=rs("id")%>">
      <td style="DISPLAY: none" id=Reply<%=rs("id")%> height="20" colspan="2"  bgcolor='#E6F1FB' >
        <textarea name="hf" cols="66" rows="3" class="STYLE4" id="hf"></textarea>
            <input type="submit" name="Submit2" value="�ύ">
     
      </td>
	   </form>
    <tr>
<!---------------------------------  ��֤"yd"�Ƿ�������,���������ʾ"վ���ظ�"  --------------------------------->
		<%
	if rs("yd")<>""then %>
      <td  bgcolor='#E6F1FB' ><span class="STYLE4">վ���ظ�:<%=clearHTMLCode(rs("hf"))%></span></td>
      <td height="10"  bgcolor='#E6F1FB' >&nbsp;</td>
	 <% 
	 end if
	  %>
      </table>
  <div align="center"><br>
    <span class="STYLE9">
    <strong>

    <%
  j=j+1
  rs.movenext
  loop
  end if
%>
<!--  ������ѯ  -->
	

<!--------------------------------  ��ҳ��ʼ -------------------------------->
    <%
if page=>8 then
 Response.Write"<a href=index.asp?page=1>��һҳ</a>"

   else
 Response.Write" "
 end if
 %>
 
    <%for j=page-4 to page-1%>
    <%if j>0 then%>
    <a href="index.asp?page=<%=j%>"><%=j%></a>
    <%end if%>
    <%next%>
 
    <%
 for j=page to page+4
%>
    <% if j<=page_count then%>
    <%if j=page then%>
    <%=j%>
    <%else%>
    <a href="index.asp?page=<%=j%>"><%=j%></a>
    <%end if%> 
    <%end if%>
    <% next 
    %>
    <%if page<page_count then%>
    </strong><a href="index.asp?page=<%=page+1%>">��һҳ</a>
    <%else%>
    <span class="STYLE6">��һҳ</span>
<%end if%>
    </span></div>
	
<!--------------------------------  ��ҳ����  -------------------------------->
</body>
</html>