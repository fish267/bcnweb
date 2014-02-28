

<!--#include file="conn.asp" -->
<!--#include file="md5.asp" -->
<%
if request.Form("pass")="" then                                    '如果获取的pass为空,则输出以下内容.     
    response.Write"<script language=javascript>alert('未输入验证密码');history.back(-1);</script>"
 response.End()                                                    '并且结束整个过程
 end if                                                            '结束if语句
 pass=md5(trim(request.form("pass")))                              '将pass去处空格,并且与md5验证
 sql="select * from admin where pass='"&pass&"'"                   '查询数据库
 set rs=server.CreateObject("adodb.recordset")                     '给予rs 操作权限
 rs.open sql,conn,1,1                                              '操作数据库
 if rs.eof then                                                    '如果到最后还没有和输入的密码相同的则转到err.html
      response.Redirect"err.html" 
	  response.End()                                               '结束一切操作
 else                                                              '否则
      session("pass")=rs("pass")                                   'pass等于数据库里的pass则转到index.asp 
      response.Redirect"index.asp"
 end if                                                            '结束if 语句
 rs.close                                                          '结束一切rs操作
 set rs=nothing                                                     '清空rs
%> 