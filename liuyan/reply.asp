<!--
---------------------文    件:reply.asp
---------------------程序名字:留言板程序(附全注释)
---------------------笔者网名:Dmlt
---------------------笔者博客:Http://HI.Baidu.com/Int_yi
---------------------笔者附言:如果有BUG,或者漏写部分,请与本人联系;
---------------------联系方式:dmlt@vip.qq.com
-->

<!--#include file="conn.asp" -->
<!--#include file="sqlin.asp" -->
<%
    if session("pass")="" then                                               '如果pass为空'则输出一下内容
	response.write "还没有登陆验证呢!"
	response.end
    end if                                                                   '结束if 语句
%>


<% 
dim rs3                                                                      '定义变量
dim sql3                                                                     '定义变量

id=request.form("id")                                                        '获取ID

hf=request("hf")                                                             '获取hf
yd="1"                                                                       '赋值给yd

set rs3=server.CreateObject("adodb.recordset")                               '给rs3数据库操作权利
sql3="select * from data where id="&cint(request.querystring("id"))          '查询数据
rs3.open sql3,conn,3,3                                                       '操作数据库
rs3("hf")=hf                                                                 '赋值给hf,更新给数据库
rs3("yd")=yd                                                                 '赋值给yd,更新给数据库
rs3.update                                                                   'rs3操作为update (更新操作)
rs3.close                                                                    '结束rs3一切操作
response.write "<script>history.go(-1)</script>"                             '返回到上一页

%>