

<%
      dim db,conn,ccc
      db="db1.mdb"                                                 '���ݿ�·��!
      Set Conn = Server.CreateObject("ADODB.ConNECtion")
      set ccc = server.createobject("adodb.recordset")
      rs="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"") 
      conn.Open rs
%>