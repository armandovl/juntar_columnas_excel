Attribute VB_Name = "M�dulo1"
'Macro unir columnas
Sub unir_columnas()
'Por Armando Vald�s
ci = Columns("A").Column 'columna inicial a unir
cf = Columns("M").Column 'columna final a unir
cd = Columns("n").Column 'columna uni�n
f = 1 'fila inicial de datos (sin encabezados es "+2")
For i = ci To cf
    uf = Cells(Rows.Count, i).End(xlUp).Row
    ud = Cells(Rows.Count, cd).End(xlUp).Row + 1 'fila inicial donde se pega
    Range(Cells(f, i), Cells(uf, i)).Copy Cells(ud, cd)
Next
End Sub
