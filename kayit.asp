<center>
<% @language="javascript" %>
<h2>Ziyaret�i Bilgileri<hr></h2>
<p>Ad�....:<%=Request.Form("ad")%>
<p>Soyad�....:<%=Request.Form("soyad")%>
<p>E-Posta....:<%=Request.Form("eposta")%>
<p>D���nceler:<%=Request.Form("mesaj")%>
<%
fso=Server.CreateObject("scripting.FileSystemObject")
if(fso.FileExists(Server.Mappath("uye.txt"))==true)
dosya=fso.OpenTextFile(Server.Mappath("uye.txt"),8)
else
dosya=fso.CreateTextFile(Server.Mappath("uye.txt"),2)
dosya.WriteLine(Request.Form("ad"))
dosya.WriteLine(Request.Form("soyad"))
dosya.WriteLine(Request.Form("eposta"))
dosya.WriteLine(Request.Form("mesaj"))
dosya.Close()
%>
<h3>Ziyaret�i Defterine Kaydedildiniz.</h3>
<a href="liste.asp">Ziyaret�i Defteri...</a>
<a href="index.asp">Ana Sayfa...</a>
</center>