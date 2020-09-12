Nueva versión del código de la DLL<br>
La versión 1.0.0.3 no existe, ya que he pasado directamente de la 2 a la 4 (son dos cambios internos en el código antes de publicarlo).
Uno en la definición del paquete para publicar en NuGet y el otro en el código para colorear desde RTF.
En el que en algunos textos RTF pegados directamente (desde Visual Studio 2019 Preview Version 16.8.0 Preview 2.1) 
asigna la definición de la tabla de colores en una línea diferente, mientras que normalmente está en la misma línea del inicio del código RTF 
o en la línea de viewkind.<br>
Este es el código<br>
<pre><span style='color:#008000'>' Puede que también esto esté en línea diferente        (12/Sep/20)
' {\colortbl
</span><span style='color:#0000FF'>If</span><span style='color:#000000'> lineas(i).TrimStart.Contains(</span><span style='color:#A31515'>"{\colortbl"</span><span style='color:#000000'>) </span><span style='color:#0000FF'>Then
</span><span style='color:#000000'>    </span><span style='color:#0000FF'>Continue</span><span style='color:#000000'> </span><span style='color:#0000FF'>For
End</span><span style='color:#000000'> </span><span style='color:#0000FF'>If
</span></pre>
<br>

