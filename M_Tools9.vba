Option Explicit

'Exemplos
#If Win64 Then
   Declare PtrSafe Function MyMathFunc Lib "User32" (ByVal N As LongLong) As LongLong
#Else
   Declare Function MyMathFunc Lib "User32" (ByVal N As Long) As Long
#End If
#If VBA7 Then
   Declare PtrSafe Sub MessageBeep Lib "User32" (ByVal N As Long)
#Else
   Declare Sub MessageBeep Lib "User32" (ByVal N As Long)
#End If


Function shuffle(n As Long) As Variant
    'Coluna 1: index embaralhado
    'Coluna 2: index original
    Dim i As Long
    Dim out As Variant
    
    ReDim out(1 To n, 1 To 2)
    
    
    For i = 1 To n
        out(i, 1) = i
        out(i, 2) = Math.Rnd * 10
    Next i
    
    QuickSortArray out, 1, n, 2, "Crescente"
    
    
    For i = 1 To n
        out(i, 2) = i
    Next i
    
    shuffle = out

End Function



Function arrayOperacoes(arr1 As Variant, ncol As Integer, operacao As String) As Double
'Recebe um array como variant (array de duas dimensões, uma tabela), e a coluna referência
'Retorna a operacao descrita sobre a coluna
'Opcoes:
'"soma", "mult", "maior", "menor"
'"gcd", "media", "desvpad"

    Dim i As Long
    Dim pivot As Double
    Dim media As Double
    
    'Checa se recebeu uma array
    If Not Information.IsArray(arr1) Then
        MsgBox "Verificar a array"
        Exit Function
    End If
    
    
    Select Case operacao
        Case "soma"
            pivot = 0
            For i = LBound(arr1, 1) To UBound(arr1, 1)
                pivot = pivot + arr1(i, ncol)
            Next i
        
        Case "mult"
            pivot = 1
            For i = LBound(arr1, 1) To UBound(arr1, 1)
                pivot = pivot * arr1(i, ncol)
            Next i
        
        Case "maior"
            pivot = arr1(LBound(arr1, 1), ncol)
            For i = LBound(arr1, 1) + 1 To UBound(arr1, 1)
                pivot = Application.WorksheetFunction.max(pivot, arr1(i, ncol))
            Next i
        
        Case "menor"
            pivot = arr1(LBound(arr1, 1), ncol)
            For i = LBound(arr1, 1) + 1 To UBound(arr1, 1)
                pivot = Application.WorksheetFunction.Min(pivot, arr1(i, ncol))
            Next i
                
        Case "gcd"
            pivot = arr1(LBound(arr1, 1), ncol)
            For i = LBound(arr1, 1) + 1 To UBound(arr1, 1)
                pivot = Application.WorksheetFunction.gcd(pivot, arr1(i, ncol))
            Next i
        
        Case "media"
            media = 0
            For i = LBound(arr1, 1) To UBound(arr1, 1)
                media = media + arr1(i, ncol)
            Next i
            pivot = media / (UBound(arr1, 1) - LBound(arr1, 1) + 1)

        Case "desvpad"
            media = 0
            For i = LBound(arr1, 1) To UBound(arr1, 1)
                media = media + arr1(i, ncol)
            Next i
            media = media / (UBound(arr1, 1) - LBound(arr1, 1) + 1)

            pivot = 0
            For i = LBound(arr1, 1) To UBound(arr1, 1)
                pivot = pivot + (arr1(i, ncol) - media) ^ 2
            Next i
            
            pivot = Math.Sqr(pivot / (UBound(arr1, 1) - LBound(arr1, 1)))

    End Select
     
    arrayOperacoes = pivot
    
    
End Function


Sub transpor(ByVal tabref As Variant, ByRef out As Variant)
'Recebe um array como variant (array de duas dimensões, uma tabela), e um array de saída como referência
'Retorna a matriz "out" transposta de "tabref"

    Dim i0 As Long
    Dim imax As Long
    
    Dim i As Long, j As Long, k As Long
    
    
    
    ReDim out(LBound(tabref, 2) To UBound(tabref, 2), LBound(tabref, 1) To UBound(tabref, 1))
    
    For i = LBound(tabref, 1) To UBound(tabref, 1)
        For j = LBound(tabref, 2) To UBound(tabref, 2)
            out(j, i) = tabref(i, j)
        Next j
    Next i
    

End Sub




Sub unpivot(ByVal edados As Variant, colref As Long, ByRef oDados As Variant)
'Faz unpivot, considerando colunas até colref (inclusive) como colunas
'Primeira linha é o título

Dim i As Long, j As Long
Dim nl As Long
Dim nc As Long
Dim p As Long
Dim count As Long
nl = UBound(edados, 1)
nc = UBound(edados, 2)

ReDim oDados(1 To nl * (nc - colref), 1 To nc)

count = 0
For i = 2 To nl 'A primeira linha é do título

    For j = colref + 1 To nc
        If edados(i, j) > 0 Then
            count = count + 1
            
            'Copia colunas até colref
            For p = 1 To colref
                oDados(count, p) = edados(i, p)
            Next p
            
            'Copia título das colunas pivot
            
            oDados(count, colref + 1) = edados(1, j)
            
            'Copia dados
            oDados(count, colref + 2) = edados(i, j)
        End If
    Next j
    
Next i


redimNaoVazio oDados, 1, oDados

End Sub



Sub ExportNumChart(wbname As String, rngRef As String, imgWidth As Long, imgHeight As Long, nomeOut As String)
'wbname: nome do workbook
'rngRef: range a ser exportado como imagem
'nameout: Number.jpg
'imgWidth: 800
'imgHeight:400

Dim pic_rng As Range
Dim ShTemp As Worksheet
Dim ChTemp As Chart
Dim PicTemp As Picture

Application.ScreenUpdating = False

Set pic_rng = Worksheets(wbname).Range(rngRef)

Set ShTemp = Worksheets.Add
Charts.Add
ActiveChart.Location Where:=xlLocationAsObject, Name:=ShTemp.Name
Set ChTemp = ActiveChart
With ChTemp.Parent
.Width = imgWidth
.Height = imgHeight
End With

pic_rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
ChTemp.Paste
Set PicTemp = Selection

ChTemp.Export Filename:=ThisWorkbook.Path & "\" & nomeOut, FilterName:="jpg"
'UserForm1.Image1.Picture = LoadPicture(FName)
'Kill FName
Application.DisplayAlerts = False
ShTemp.Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub


Sub msgtempo(t0 As Double, unidade As String)

Select Case unidade
    Case "s"
        MsgBox "Concluído em " & Math.Round((DateTime.Now - t0) * 24 * 60 * 60, 3) & " s"
    Case "min"
        MsgBox "Concluído em " & Math.Round((DateTime.Now - t0) * 24 * 60, 1) & " min"
    Case "h"
        MsgBox "Concluído em " & Math.Round((DateTime.Now - t0) * 24, 1) & "  h"
End Select

End Sub

Sub EscondeGridLines()

Dim sht As Worksheet

For Each sht In ActiveWorkbook.Worksheets
    sht.Activate
    ActiveWindow.DisplayGridlines = False
    
Next
    
End Sub


Sub timedelay(segundos As Double)
'Para fracoes de segundos

Dim x As Single
Dim i As Long


x = DateTime.Timer

While DateTime.Timer - x < segundos
    
Wend


End Sub

Sub esperar(segundos As Integer)
Dim newHour
Dim newMinute
Dim newSecond
Dim waittime


newHour = Hour(Now())
newMinute = Minute(Now())
newSecond = Second(Now()) + segundos

waittime = TimeSerial(newHour, newMinute, newSecond)


Application.Wait (waittime)

'Ou algo do tipo
'Application.Wait (datetime.now+ timeserial("00:00:01")

End Sub

Function Is64bit() As Boolean
    Is64bit = Len(Environ("ProgramW6432")) > 0
End Function


Sub refreshPivots()
'Just to do not forget
ThisWorkbook.RefreshAll

End Sub


Public Sub ImportaTexto(FName As String, ByRef varOut As String)
'Le arquivo e devolve uma string com tudo
'Ex.
'ImportaTexto "C:\\JScript\\Exemplo1.txt", ","

Application.ScreenUpdating = False
'On Error GoTo EndMacro:
Dim varSwap  As String

Open FName For Input Access Read As #1

While Not EOF(1)
    Line Input #1, varSwap
    varOut = varOut & varSwap & ","
Wend

EndMacro:
On Error GoTo 0
Application.ScreenUpdating = True
Close #1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END ImportTextFile
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub



Private Sub exportaTexto(ByVal strPath As String, varArray As Variant)
'Exporta Formulação no formato de leitura
'Entradas:
Dim strSwap As String, tamstrswap As Integer
Dim fs As Variant
Dim afs As Variant, i As Long, j As Long
Dim ndim As Long


Application.ScreenUpdating = False



'Inicializa funcao de manipulação de arquivos
Set fs = CreateObject("Scripting.FileSystemObject")
Set afs = fs.CreateTextFile(strPath, 1, 1)


If Information.IsArray(varArray) Then
    ndim = nDimensions(varArray)
    For i = 1 To UBound(varArray, 1)
        strSwap = ""
        
        If ndim = 2 Then
            For j = 1 To UBound(varArray, 2)
                strSwap = strSwap & Strings.Replace(varArray(i, j), ",", ".") & " "
            Next j
        Else
            strSwap = strSwap & Strings.Replace(varArray(i, 1), ",", ".")
        End If
        
        afs.WriteLine (strSwap)
    Next i
Else
     afs.WriteLine (varArray)
End If

afs.Close

Set afs = Nothing
Set fs = Nothing

End Sub



Function vba_to_jsChart(arrRef As Variant, ArrName As String, strPath As String)
'Makes a json object from an array and saves it locally

Dim strArr As String
Dim i As Long, j As Long


strArr = "var " & ArrName & " = new Array("

For i = 1 To UBound(arrRef, 1)
    strArr = strArr + "["
    
    j = 1
    strArr = strArr & "'" & arrRef(i, j) & "', "
    
    For j = 2 To UBound(arrRef, 2) - 1
        strArr = strArr & arrRef(i, j) & ", "
    Next j
    
    j = UBound(arrRef, 2)
    strArr = strArr & arrRef(i, j)
    
    strArr = strArr & "]"
    
    If i < UBound(arrRef, 1) Then
        strArr = strArr & ", "
    End If
    
Next i
strArr = strArr & ");"


exportaTexto strPath & "/" & ArrName & ".js", strArr


End Function
Function nDimensions(a As Variant) As Long
    Dim c As New Collection
    Dim v As Variant, i As Long

    On Error GoTo exit_function
    i = 1
    Do While True
        v = Array(LBound(a, i), UBound(a, i))
        c.Add v
        i = i + 1
    Loop
exit_function:
    'Set nBounds = c
    
    nDimensions = c.count
End Function

Function vba_to_json(arrRef As Variant, colref As Variant, objectName As String, strPath As String)
'Makes a json object from an array and saves it locally

Dim strJson As String
Dim i As Long, j As Long
Dim strExport As String


strJson = "var " & objectName & " = [" & Strings.Chr(10)

For i = 1 To UBound(arrRef, 1)
    strJson = strJson & "{"
    For j = 1 To UBound(arrRef, 2)
        strExport = Strings.Replace(arrRef(i, j), ",", ".")
        
        strJson = strJson & """" & colref(1, j) & """" & ": " & """" & strExport & """"

        If j < UBound(arrRef, 2) Then
            strJson = strJson & ", "
        End If
    Next j
    
    If i < UBound(arrRef, 1) Then
        strJson = strJson & "}," & Strings.Chr(10)
    Else
        strJson = strJson & "}" & Strings.Chr(10)
    End If
Next i

strJson = strJson & "];"


exportaTexto strPath & "/" & objectName & ".js", strJson


End Function

Sub idx2arr(vecIdx As Variant, maxlin As Long, maxcol As Long, outarr As Variant)
'Transforma vecIdx [idxVolume, tRef, Volume]
'em outarr (nos formatos de linha, coluna)
Dim i As Long
Dim lin As Long, col As Long

ReDim outarr(1 To maxlin, 1 To maxcol)

For i = 1 To UBound(vecIdx, 1)
    If vecIdx(i, 1) <> "" Then
        lin = vecIdx(i, 1)
        col = vecIdx(i, 2)
        outarr(lin, col) = vecIdx(i, 3)
    End If
Next i

End Sub

Sub arr2idx(Arr As Variant, vecIdx As Variant)
'Transforma array em vecIdx [idxVolume, tRef, Volume]

Dim i As Long, j As Long, co As Long
Dim maxlin As Long, maxcol As Long
maxlin = UBound(Arr, 1)
maxcol = UBound(Arr, 2)

ReDim vecIdx(1 To maxlin * maxcol, 1 To 3)

For i = 1 To maxlin
    For j = 1 To maxcol
        If Arr(i, j) <> "" Then
            co = co + 1
            vecIdx(co, 1) = i
            vecIdx(co, 2) = j
            vecIdx(co, 3) = Arr(i, j)
        End If
    Next j
Next i


redimNaoVazio vecIdx, 1, vecIdx

End Sub

Sub contaCorrente(arrVolumes As Variant, arrProdut As Variant, vecIdxVol As Variant)
'ArrVolumes em linha, arrProdutividade em coluna
'vecIdxVol [idxVolume, tRef, Volume]

Dim i As Long, j As Long, k As Long
Dim volRef As Double
Dim tref As Long, co As Long
Dim resProdut As Double

'vecIdxVol [idx, tref, vol]
ReDim vecIdxVol(1 To 1000, 1 To 3)

'Calcula volumes dadas as produtividades


tref = 0
resProdut = 0
co = 0
For i = 1 To UBound(arrVolumes, 1)
    volRef = arrVolumes(i, 1)
    
    'Distribui o volume até acabar os slots da produtividade
    While volRef > 0.0001
        If volRef < resProdut Then
            'Atribui
            co = co + 1
            vecIdxVol(co, 1) = i 'idx
            vecIdxVol(co, 2) = tref 'tref
            vecIdxVol(co, 3) = volRef 'vol
            
            resProdut = resProdut - volRef
            volRef = 0
            
        Else
            
            'Atribui o que resta de producao
            If resProdut > 0 Then
                co = co + 1
                vecIdxVol(co, 1) = i 'idx
                vecIdxVol(co, 2) = tref 'tref
                vecIdxVol(co, 3) = resProdut 'vol
                volRef = volRef - resProdut
            End If
            
            
            'Passa para o próximo periodo
            tref = tref + 1
            
            
            If tref > UBound(arrProdut, 2) Then
                'Processamento concluido
                Exit Sub
            Else
                resProdut = arrProdut(1, tref)
                
            End If
        End If
    Wend

Next i


'Redim não vazio
'redimNaoVazio vecIdxVol, 1, vecIdxVol


End Sub



Sub testeconcatColunas()
Dim arr1
Dim out


copiaDados 2, 1, 44, arr1, "Plan1"

out = concatColunas(arr1, 2, 3, 4)


End Sub

Function concatColunas(arrRef As Variant, ParamArray colunas() As Variant) As Variant
'Monta uma chave com as colunas concatenadas de arrRef indicadas pelo array colunas,
'Chave separada por "_"
'Retorna um vetor de dimensão [ubound(arrref),1] com a chave

Dim j As Long
Dim i As Long
Dim vecChave As Variant

ReDim vecChave(1 To UBound(arrRef, 1), 1 To 1)

For i = 1 To UBound(arrRef, 1)
    For j = LBound(colunas) To UBound(colunas)
        vecChave(i, 1) = vecChave(i, 1) & "_" & arrRef(i, colunas(j))
    Next j
Next i

concatColunas = vecChave



End Function


Sub testeBuscaChaveScript()
Dim chv As Object
Dim arr1 As Variant
Dim arrfind As Variant
Dim outIdx As Variant

copiaDados 2, 1, 44, arr1, "Plan1"
copiaDados 1, 1, 1, arrfind, "Plan3"


'Gera a chave de scrips
Set chv = geraChaveScript(arr1, 1)

'Procura os valores de arrFind (colFind) na chave de scripts e retorna o índice
outIdx = findChaveScript(chv, arrfind, 1)

colaDados 1, 2, 1, outIdx, "Plan3"

End Sub

Function geraChaveScriptBusca(arrRef As Variant, colref As Long) As Object
'Retorna o objeto

Dim i As Long, j As Long
Dim chaveref As Object

If Not Information.IsArray(arrRef) Then
    MsgBox "Array vazio"
    Exit Function
End If


Set chaveref = CreateObject("scripting.dictionary")

'Cria chave
For i = 1 To UBound(arrRef, 1)
    If arrRef(i, colref) <> "" Then
        If Not chaveref.exists(arrRef(i, colref)) Then
            chaveref.Add arrRef(i, colref), i
        End If
    End If
Next i

Set geraChaveScript = chaveref


End Function


Function findChaveScript(chaveref As Object, arrfind As Variant, colFind As Long) As Variant
'retorna vetor de indices vecidx
Dim i As Long


If Not Information.IsArray(arrfind) Then
    MsgBox "Array vazio"
    Exit Function
End If


ReDim vecIdx(1 To UBound(arrfind, 1), 1 To 1)

'Acha indices
For i = 1 To UBound(arrfind, 1)
    vecIdx(i, 1) = chaveref.Item(arrfind(i, colFind))
Next i

findChaveScript = vecIdx

End Function

Sub buscaEficiente(arrRef As Variant, colref As Long, arrfind As Variant, colFind As Long, vecIdx As Variant)
'buscaEficiente arrRef, colRef , arrFind , colFind , vecIdx
'Busca valores de arrFind e colFind em ArrRef-colRef
'Retorna indexes   vecIdx(i, 1)

Dim chaveref As Object
Dim i As Long, j As Long


If Not Information.IsArray(arrRef) Or Not Information.IsArray(arrfind) Then
    MsgBox "Arrays vazios"
    Exit Sub
End If

ReDim vecIdx(1 To UBound(arrfind, 1), 1 To 1)


Set chaveref = CreateObject("scripting.dictionary")

'Cria chave
For i = 1 To UBound(arrRef, 1)
    If arrRef(i, colref) <> "" Then
        If Not chaveref.exists(arrRef(i, colref)) Then
            chaveref.Add arrRef(i, colref), i
        End If
    End If
Next i


'Acha indices
For i = 1 To UBound(arrfind, 1)
    vecIdx(i, 1) = chaveref.Item(arrfind(i, colFind))
Next i

End Sub



Sub buscaEficiente1val(arrRef As Variant, colref As Long, valFind As String, vecIdx As Variant)
'buscaEficiente arrRef, colRef , arrFind , colFind , vecIdx

Dim chaveref As Object
Dim i As Long, j As Long


If Not Information.IsArray(arrRef) Then
    MsgBox "Arrays vazios"
    Exit Sub
End If


Set chaveref = CreateObject("scripting.dictionary")

'Cria chave
For i = 1 To UBound(arrRef, 1)
    If arrRef(i, colref) <> "" Then
        If Not chaveref.exists(arrRef(i, colref)) Then
            chaveref.Add arrRef(i, colref), i
        End If
    End If
Next i


'Acha indices
vecIdx = chaveref.Item(valFind)

End Sub
 
 
Function corcelula(celref As Range)

 corcelula = celref.Interior.Color
 
End Function


Function corcelulaHX(celref As Range)
Dim colorVal
Dim cR, cG, cB

colorVal = celref.Interior.Color


cR = colorVal Mod 256
cG = (colorVal \ 256) Mod 256
cB = colorVal \ 65536

corcelulaHX = "#" & Strings.Format(Hex(cR), "00") & Strings.Format(Hex(cG), "00") & Strings.Format(Hex(cB), "00")
 
End Function

Sub delNames()

Dim nme As Name

On Error Resume Next
For Each nme In ActiveWorkbook.Names
    nme.Delete
Next

End Sub

Sub desativaautofiltros()

Dim sht As Worksheet

For Each sht In ActiveWorkbook.Worksheets
    If sht.AutoFilterMode Then
        sht.AutoFilter.ShowAllData
    End If
Next

End Sub

Sub PseudoSelect(ByRef varRef As Variant, nomeSht As String, Optional strPath As String, Optional maxlin As Long = 10 ^ 6)
'Copia dados da base excel
'Tipo um select * from

Dim wbname As String
Dim nl As Long, nc As Long
Dim linIni As Long, colini As Long

linIni = 2
colini = 1

maxlin = 50000

Application.ScreenUpdating = False

If strPath <> "" Then
    Workbooks.Open strPath
    wbname = ActiveWorkbook.Name


    If nomeSht <> "" Then
         Sheets(nomeSht).Activate
    End If
    
    nl = Application.WorksheetFunction.CountA(Range(Cells(linIni, 1), Cells(linIni + maxlin, 1)))
    nc = Application.WorksheetFunction.CountA(Range(Cells(1, 1), Cells(1, 1000))) 'Max 1000 col
    
    varRef = Range(Cells(linIni, colini), Cells(linIni + nl - 1, colini + nc - 1))
    Workbooks(wbname).Close
    

End If

End Sub


Sub PseudoInsert(linIni As Integer, colini As Integer, ncols As Long, ByRef varRef As Variant, Optional nomeSht As String = "", Optional strPath As String, Optional maxlin As Long = 10 ^ 6)
'Cola dados numa planilha externa
'Tipo um insert table com delete
Dim wbname As String

Application.ScreenUpdating = False

If strPath <> "" Then
    Workbooks.Open strPath
    wbname = ActiveWorkbook.Name
End If

If nomeSht <> "" Then
     Sheets(nomeSht).Activate
End If

Range(Cells(linIni, colini), Cells(linIni + 500000, colini + ncols - 1)).ClearContents
Range(Cells(linIni, colini), Cells(linIni, colini)).Resize(UBound(varRef, 1), UBound(varRef, 2)) = varRef

Workbooks(wbname).Save
Workbooks(wbname).Close

End Sub



Sub colaDados(linIni As Integer, colini As Integer, ncols As Long, ByRef varRef As Variant, Optional nomeSht As String = "", Optional maxlin As Long = 10 ^ 6)
'Cola dados

If nomeSht <> "" Then
        Sheets(nomeSht).Activate
End If

Range(Cells(linIni, colini), Cells(linIni + 500000, colini + ncols - 1)).ClearContents
Range(Cells(linIni, colini), Cells(linIni, colini)).Resize(UBound(varRef, 1), ncols) = varRef

End Sub



Sub copiaDados(linIni As Integer, colini As Integer, ncols As Long, ByRef varRef As Variant, Optional nomeSht As String = "", Optional maxlin As Long = 10 ^ 6)
'Copia dados da planilha nomeSht, a comecas da linIni e colIni, para varRef

Dim nl As Long
Dim nc As Long

Application.ScreenUpdating = False


If nomeSht <> "" Then
        Sheets(nomeSht).Activate
End If


nl = Application.WorksheetFunction.CountA(Range(Cells(linIni, colini), Cells(linIni + maxlin, colini)))

Cells(linIni, colini).Select
nc = ncols

If nl > 0 And nc > 0 Then
    varRef = Range(Cells(linIni, colini), Cells(linIni + nl - 1, colini + nc - 1))
End If



End Sub

Sub copiaDadosLin(linIni As Integer, colini As Integer, ncols As Long, nlins As Long, ByRef varRef As Variant, Optional nomeSht As String = "", Optional maxlin As Long = 10 ^ 6)
'Copia dados da planilha nomeSht, a comecas da linIni e colIni, para varRef

Dim nc As Long

Application.ScreenUpdating = False


If nomeSht <> "" Then
        Sheets(nomeSht).Activate
End If



Cells(linIni, colini).Select
nc = ncols

If nlins > 0 And nc > 0 Then
    varRef = Range(Cells(linIni, colini), Cells(linIni + nlins - 1, colini + nc - 1))
End If



End Sub


Sub limpaAutofiltro()

If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilter.ShowAllData
End If

End Sub

Function buscaSimples(ByVal tabref As Variant, ByVal colref As Long, ByVal val As String) As Long
'Procura primeiro valor menor, menor igual, maior, maior igual
'Retorna o indice da tabela ou zero, se nao achou
Dim i As Long
 buscaSimples = 0

    For i = 1 To UBound(tabref, 1)
        If tabref(i, colref) = val Then
            buscaSimples = i
            Exit For
        End If
    Next i

End Function

Function buscaSimplesDois(ByVal tabref As Variant, ByVal colref1 As Long, ByVal val1 As String, ByVal colref2 As Long, ByVal val2 As Long) As Long
'Procura primeiro valor menor, menor igual, maior, maior igual
'Retorna o indice da tabela ou zero, se nao achou
Dim i As Long
 buscaSimplesDois = 0

    For i = 1 To UBound(tabref, 1)
        If tabref(i, 1) <> "" Then
            If tabref(i, colref1) = val1 And tabref(i, colref2) = val2 Then
                buscaSimplesDois = i
                Exit For
            End If
        End If
    Next i

End Function

Sub selectDistinctSimples(ByVal tabref As Variant, ByVal colref As Long, ByRef out As Variant)
'Faz select distinct da coluna de referencia
'

Dim i As Long
Dim counter As Long
Dim nargs As Integer
Dim swap As Variant
Dim swap2 As Variant
Dim strSwap As Variant


ReDim out(1 To UBound(tabref, 1), 1 To 1)
ReDim swap(1 To UBound(tabref, 1), 1 To 1)


'Monta uma tabela concatenada
For i = 1 To UBound(tabref, 1)
    swap(i, 1) = swap(i, 1) & tabref(i, colref)
Next i

'remove duplicatas da tabela concatenada
 swap2 = remove_duplicate(swap)
 counter = 0
For i = LBound(swap2, 1) To UBound(swap2, 1)
    out(i + 1, 1) = swap2(i)
Next i

End Sub
Function buscamin(ByVal tabref As Variant, ByVal colref As Long)
Dim i As Long

Dim ref As Double
Dim idxref As Long

ref = 1000000

For i = 1 To UBound(tabref, 1)
    If tabref(i, colref) < ref And tabref(i, colref) <> "" Then
        idxref = i
        ref = tabref(i, colref)
    End If
Next i

buscamin = idxref
End Function



Function buscamax(ByVal tabref As Variant, ByVal colref As Long) As Long
'Busca o maior valor da tabela
Dim i As Long
Dim max As Double
Dim idx As Long

max = 0

For i = 1 To UBound(tabref, 1)
    If tabref(i, colref) > max Then
       max = tabref(i, colref)
       idx = i
    End If
Next i

buscamax = idx
End Function

Function buscamenor(ByVal tabref As Variant, ByVal colref As Long, ByVal val As Double, ByVal sinal As String)
'Procura primeiro valor menor, menor igual, maior, maior igual
'Retorna o indice da tabela ou zero, se nao achou
Dim i As Long
Dim idx As Long

Select Case sinal
Case "<"
    For i = 1 To UBound(tabref, 1)
        If tabref(i, colref) < val Then
            idx = i
            Exit For
        End If
    Next i
Case "<="
    For i = 1 To UBound(tabref, 1)
        If tabref(i, colref) <= val Then
            idx = i
            Exit For
        End If
    Next i
Case ">"
    For i = 1 To UBound(tabref, 1)
        If tabref(i, colref) > val Then
            idx = i
            Exit For
        End If
    Next i
Case ">="
    For i = 1 To UBound(tabref, 1)
        If tabref(i, colref) >= val Then
            idx = i
            Exit For
        End If
    Next i
End Select
    

End Function


Function tempodecorrido(ByVal t0 As Double) As String

    tempodecorrido = "Concluido em " & Math.Round((Now - t0) * 24 * 60, 1) & " min"

End Function


Sub reexibir()

Dim sht As Worksheet
For Each sht In ActiveWorkbook.Worksheets
    
    sht.Visible = xlSheetVisible
Next

End Sub





Sub filtrosimplesMenorIgual(ByVal tabref As Variant, colref As Long, critfiltro As Long, out As Variant)
'filtra tabref na coluna colref, criterio critfiltro, e retorna out
Dim i As Long, j As Long
Dim counter As Long

ReDim out(1 To UBound(tabref, 1), 1 To UBound(tabref, 2))

For i = 1 To UBound(tabref, 1)
    If tabref(i, colref) <= critfiltro Then
        counter = counter + 1
        For j = 1 To UBound(tabref, 2)
            out(counter, j) = tabref(i, j)
        Next j
    End If
Next i

End Sub

Sub filtrosimplesMaiorIgual(ByVal tabref As Variant, colref As Long, critfiltro As Double, out As Variant)
'filtra tabref na coluna colref, criterio critfiltro, e retorna out
Dim i As Long, j As Long
Dim counter As Long

ReDim out(1 To UBound(tabref, 1), 1 To UBound(tabref, 2))

For i = 1 To UBound(tabref, 1)
    If tabref(i, colref) >= critfiltro Then
        counter = counter + 1
        For j = 1 To UBound(tabref, 2)
            out(counter, j) = tabref(i, j)
        Next j
    End If
Next i

End Sub

Sub filtrosimples(tabref As Variant, colref As Long, critfiltro As String, out As Variant)
'filtra tabref na coluna colref, criterio critfiltro, e retorna out
Dim i As Long, j As Long
Dim counter As Long

ReDim out(1 To UBound(tabref, 1), 1 To UBound(tabref, 2))

For i = 1 To UBound(tabref, 1)
    If Conversion.CStr(tabref(i, colref)) = critfiltro Then
        counter = counter + 1
        For j = 1 To UBound(tabref, 2)
            out(counter, j) = tabref(i, j)
        Next j
    End If
Next i

End Sub

Sub filtrosimplesdiferente(ByVal tabref As Variant, ByVal colref As Long, ByVal critfiltro As String, ByRef out As Variant)
'filtra tabref <> critfiltro na coluna colref, criterio critfiltro, e retorna out
Dim i As Long, j As Long
Dim counter As Long

ReDim out(1 To UBound(tabref, 1), 1 To UBound(tabref, 2))

For i = 1 To UBound(tabref, 1)
    If Conversion.CStr(tabref(i, colref)) <> critfiltro Then
        counter = counter + 1
        For j = 1 To UBound(tabref, 2)
            out(counter, j) = tabref(i, j)
        Next j
    End If
Next i

End Sub



Sub salvarCenario(tabref As Variant, ncen As Long, planref As String)
'salva a tabref, número de cenário ncen, na planilha de nome planref
'da espacamento de 9 colunas

Dim ncols As Long
Dim nlin As Long
Dim colini As Long, colfim As Long

Application.ScreenUpdating = False
nlin = UBound(tabref, 1)
ncols = UBound(tabref, 2)


colini = 1 + (ncen - 1) * (ncols + 9)
colfim = 1 + (ncen) * (ncols + 9) - 1

Sheets(planref).Activate

Range(Cells(2, colini), Cells(60000, colfim)).ClearContents
Range(Cells(2, colini), Cells(2, colini)).Resize(nlin, ncols) = tabref

'Salva n linhas e cols
Range("a1") = ncols
Range(Cells(1, colini + 1), Cells(1, colini + 1)) = nlin


End Sub


Function obterCenario(tabref As Variant, ncen As Long, planref As String) As Long
'obtem número de cenário ncen, na planilha de nome planref e retorna em tabref
'da espacamento de 9 colunas

Dim ncols As Long
Dim nlin As Long
Dim colini As Long, colfim As Long

Application.ScreenUpdating = False

Sheets(planref).Activate

ncols = Range("a1")


colini = 1 + (ncen - 1) * (ncols + 9)
colfim = 1 + (ncen) * (ncols + 9) - 1
nlin = Range(Cells(1, colini + 1), Cells(1, colini + 1))

If nlin > 0 Then
    tabref = Range(Cells(2, colini), Cells(nlin, colini + ncols - 1))
    obterCenario = 1
Else
    obterCenario = 0
End If


End Function




Sub selectdistinct(ByVal tabref As Variant, cols As Variant, out As Variant)
'Faz select distinct de colunas a serem referenciadas
'out tem as colunas filtradas
'cols no formato col(1, j) = 1

Dim i As Long, j As Long
Dim counter As Long
Dim nargs As Integer
Dim swap As Variant
Dim swap2 As Variant
Dim strSwap As Variant



For j = 1 To UBound(cols, 2)
    If cols(1, j) > 0 Then
        nargs = nargs + 1
    End If
Next j


ReDim out(1 To UBound(tabref, 1), 1 To nargs)
ReDim swap(1 To UBound(tabref, 1), 1 To 1)


'Monta uma tabela concatenada
For i = 1 To UBound(tabref, 1)
    
    For j = 1 To UBound(cols, 2)
        If cols(1, j) > 0 Then
            swap(i, 1) = swap(i, 1) & tabref(i, j) & "!!!"
        End If
    Next j
Next i

'remove duplicatas da tabela concatenada
 swap2 = remove_duplicate(swap)
 
For i = LBound(swap2, 1) To UBound(swap2, 1)
    strSwap = ""
    strSwap = Strings.Split(swap2(i), "!!!")
    
    For j = 0 To UBound(strSwap, 1) - 1 'Para descontar o ultimo "!!!"
        out(i + 1, j + 1) = strSwap(j)
    Next j

Next i

End Sub


 Function remove_duplicate(ByVal varArray As Variant)

' \\ Declaração de variáveis
Dim varValue As Variant

' \\ Cria o objeto dictionary
With CreateObject("scripting.dictionary")
  .CompareMode = vbTextCompare ' \\ Compara texto
  For Each varValue In varArray '\\ Para cada valor na matriz
   If Not Strings.Len(varValue) = 0 And Not .exists(varValue) Then '\\ Desconsidera valores vazios, alterar esta linha caso queira considerar
      .Add varValue, Nothing
    End If
  Next
  remove_duplicate = .keys
End With

End Function
    
Function remove_duplicateCol(ByVal varArray As Variant, ncol As Long)

' \\ Declaração de variáveis
Dim varValue As Variant
Dim i As Long

' \\ Cria o objeto dictionary
With CreateObject("scripting.dictionary")
  .CompareMode = vbTextCompare ' \\ Compara texto
  For i = 1 To UBound(varArray, 1) '\\ Para cada valor na matriz
      
     varValue = varArray(i, ncol)
     If Not Strings.Len(varValue) = 0 And Not .exists(varValue) Then '\\ Desconsidera valores vazios, alterar esta linha caso queira considerar
        .Add varValue, Nothing
      End If
  Next
  remove_duplicateCol = .keys
End With

End Function
    

Sub OrdenarMatriz(ByRef arrMatriz As Variant, kcol As Variant, strSentido As Variant)
'Ordena arrMatriz
'Kcol - Seq de ordenacao 1, 2, 3
'Sentido "Crescente" ou "Decrescente"
    Dim temp As Variant
    Dim i As Long
    Dim j As Long
Dim ncriterio As Long
    Dim Xi As Long
    Dim Xf As Long
    Dim x As Long
    Dim icrit As Long
    Dim Lbnd As Integer
    Dim colref As Long
Dim ncols As Long

'Conta quantas colunas devem ser filtradas

For i = 1 To UBound(kcol, 1)
    If kcol(i, 1) > 0 Then
        ncols = ncols + 1
    End If
Next i

    Lbnd = LBound(arrMatriz, 1)
    Xi = LBound(arrMatriz, 2)
    Xf = UBound(arrMatriz, 2)

For ncriterio = ncols To 1 Step -1
    For icrit = 1 To UBound(kcol, 1)
        If kcol(icrit, 1) = ncriterio Then
            colref = icrit
            Exit For
        End If
    Next icrit

If strSentido(colref, 1) = "Crescente" Then
    For i = UBound(arrMatriz, 1) - 1 To Lbnd Step -1
            For j = Lbnd To i
                If arrMatriz(j, colref) > arrMatriz(j + 1, colref) Then
                    For x = Xi To Xf
                        temp = arrMatriz(j + 1, x)
                        arrMatriz(j + 1, x) = arrMatriz(j, x)
                        arrMatriz(j, x) = temp
                    Next x
                End If
            Next
    Next
Else
    For i = UBound(arrMatriz, 1) - 1 To Lbnd Step -1
            For j = Lbnd To i
                If arrMatriz(j, colref) < arrMatriz(j + 1, colref) Then
                    For x = Xi To Xf
                        temp = arrMatriz(j + 1, x)
                        arrMatriz(j + 1, x) = arrMatriz(j, x)
                        arrMatriz(j, x) = temp
                    Next x
                End If
            Next
    Next i

End If
Next ncriterio

End Sub


Sub Ordenarsimples(ByRef arrMatriz As Variant, kcol As Long, strSentido As Variant)
'Ordena arrMatriz
'Kcol: coluna a ordenar
'Sentido "Crescente" ou "Decrescente"
    Dim temp As Variant
    Dim i As Long
    Dim j As Long
    Dim ncriterio As Long
    Dim Xi As Long
    Dim Xf As Long
    Dim x As Long
    Dim icrit As Long
    Dim Lbnd As Integer
    Dim colref As Long
    Dim ncols As Long


    ncols = 1
    
    Lbnd = LBound(arrMatriz, 1)
    Xi = LBound(arrMatriz, 2)
    Xf = UBound(arrMatriz, 2)


colref = kcol

If strSentido = "Crescente" Then
    For i = UBound(arrMatriz, 1) - 1 To Lbnd Step -1
            For j = Lbnd To i
                If arrMatriz(j, colref) > arrMatriz(j + 1, colref) And arrMatriz(j + 1, colref) > 0 Then
                    For x = Xi To Xf
                        temp = arrMatriz(j + 1, x)
                        arrMatriz(j + 1, x) = arrMatriz(j, x)
                        arrMatriz(j, x) = temp
                    Next x
                End If
            Next
    Next
Else
    For i = UBound(arrMatriz, 1) - 1 To Lbnd Step -1
            For j = Lbnd To i
                If arrMatriz(j, colref) < arrMatriz(j + 1, colref) And arrMatriz(j + 1, colref) > 0 Then
                    For x = Xi To Xf
                        temp = arrMatriz(j + 1, x)
                        arrMatriz(j + 1, x) = arrMatriz(j, x)
                        arrMatriz(j, x) = temp
                    Next x
                End If
            Next
    Next i

End If

End Sub


Sub filtroTab(ByVal tabref As Variant, ByVal e_colfiltro As Variant, ByVal e_colsinal As Variant, ByVal e_colvalues As Variant, ByRef out As Variant)
'Recebo uma tabela (tabref) de referencia, quais colunas serao filtradas (e_colfiltro), valores de filtro (e_colvalues
'O colFiltro tem o seguinte código: 1 =, 2 >, 3 >=, 4 <, 5 <=, 6 <>
'Condição ou na segunda dimensão do e_colfiltro
'Exemplo
'Dim e_colfiltro as variant
'Dim e_colsinal as variant
'Dim e_colvalues as variant
'   ReDim e_colfiltro(1 To 5, 1 To 1)
'    ReDim e_colsinal(1 To 5, 1 To 1)
'    ReDim e_colvalues(1 To 5, 1 To 1)
'
'    'Mesmo produto gram
'    e_colfiltro(1, 1) = 3
'    e_colsinal(1, 1) = "="
'    e_colvalues(1, 1) = Strings.Left(e_estoque(i, 1), 5)
'    e_colfiltro(2, 1) = 4
'    e_colsinal(2, 1) = "="
'    e_colvalues(2, 1) = Conversion.CLng(Strings.Right(e_estoque(i, 1), 3))
'    filtroTab e_pedido, e_colfiltro, e_colsinal, e_colvalues, out

'retorno out
Dim i As Long, j As Long
Dim ouCtrl As Long
Dim isok As Boolean
Dim isOkInner As Boolean
Dim isOKor As Boolean
Dim counter As Long

ReDim out(1 To UBound(tabref, 1), 1 To UBound(tabref, 2))

counter = 1

For i = 1 To UBound(tabref, 1) 'Percorre todas as linhas da tabela
    isOKor = False
    For ouCtrl = 1 To UBound(e_colfiltro, 2) 'Para todas as condicoes "ou"
    isok = True
    
    For j = 1 To UBound(e_colfiltro, 1) 'Percorre as condicoes a serem seguidas
        If e_colsinal(j, ouCtrl) = "=" Then
            isOkInner = False
            If tabref(i, e_colfiltro(j, ouCtrl)) = e_colvalues(j, ouCtrl) And e_colvalues(j, ouCtrl) <> "" Then
                    isOkInner = True
            End If
        ElseIf e_colsinal(j, ouCtrl) = ">" Then
            isOkInner = False
                If Conversion.CDbl(tabref(i, e_colfiltro(j, ouCtrl))) > Conversion.CDbl(e_colvalues(j, ouCtrl)) And e_colvalues(j, ouCtrl) <> "" Then
                        isOkInner = True
                End If
        ElseIf e_colsinal(j, ouCtrl) = ">=" Then
            isOkInner = False
                If Conversion.CDbl(tabref(i, e_colfiltro(j, ouCtrl))) >= Conversion.CDbl(e_colvalues(j, ouCtrl)) And e_colvalues(j, ouCtrl) <> "" Then
                        isOkInner = True
                End If
        ElseIf e_colsinal(j, ouCtrl) = "<" Then
            isOkInner = False
                If Conversion.CDbl(tabref(i, e_colfiltro(j, ouCtrl))) < Conversion.CDbl(e_colvalues(j, ouCtrl)) And e_colvalues(j, ouCtrl) <> "" Then
                        isOkInner = True
                End If
        ElseIf e_colsinal(j, ouCtrl) = "<=" Then
            isOkInner = False
                If Conversion.CDbl(tabref(i, e_colfiltro(j, ouCtrl))) <= Conversion.CDbl(e_colvalues(j, ouCtrl)) And e_colvalues(j, ouCtrl) <> "" Then
                        isOkInner = True
                End If
        ElseIf e_colsinal(j, ouCtrl) = "<>" Then
            isOkInner = False
                If tabref(i, e_colfiltro(j, ouCtrl)) <> e_colvalues(j, ouCtrl) And e_colvalues(j, ouCtrl) <> "" Then
                        isOkInner = True
                End If
'        Else
'            isOkInner = False
        End If
        
        If e_colfiltro(j, ouCtrl) <> "" Then
            isok = isok * isOkInner
        End If
    Next j


    isOKor = isOKor Or isok
    Next ouCtrl
    
    If isOKor = True Then
        For j = 1 To UBound(tabref, 2)
            out(counter, j) = tabref(i, j)
        Next j
        counter = counter + 1
    End If
    
Next i



End Sub




Sub Mapeiatabela(ByVal tabref As Variant, ByVal e_depara As Variant, ByRef out As Variant)
'copia uma tabela para outra de outro formato
'de tabentrada para tabsaida
'e_depara(1,1) = 14


'retorno out
Dim i As Long, j As Long, k As Long
Dim isok As Boolean
Dim isOkInner As Boolean
Dim counter As Long
Dim ncolspara As Long

'Conta numero de colunas da saida
counter = 0
For i = 1 To UBound(e_depara, 1)
    If e_depara(i, 1) > counter Then
        counter = e_depara(i, 1)
    End If
Next i
ncolspara = counter


ReDim out(1 To UBound(tabref, 1), 1 To ncolspara)


For i = 1 To UBound(tabref, 1)
    For j = 1 To UBound(e_depara, 1)
        If e_depara(j, 1) > 0 Then
            out(i, e_depara(j, 1)) = tabref(i, j)
        End If
    Next j
Next i


End Sub



Sub somarconcatsimples(ByVal tabref As Variant, ByVal colFiltro As Long, ByRef out As Variant)
Dim e_colfiltro As Variant
Dim e_colsoma As Variant
Dim e_colconcat As Variant


ReDim e_colfiltro(1 To UBound(tabref, 2), 1 To 1)
ReDim e_colsoma(1 To UBound(tabref, 2), 1 To 1)
ReDim e_colconcat(1 To UBound(tabref, 2), 1 To 1)


e_colfiltro(colFiltro, 1) = 1

somarconcat tabref, e_colfiltro, e_colsoma, e_colconcat, out



End Sub

Function listaUnicos(tabref As Variant, colref As Long, lstUnique As Variant) As Boolean
'Identifica na tabela, na coluna de referencia, quem tem valores unicos e joga o resultado em lstUnique
'lstUnique (col1: ponteiro para original, col2: valor)

Dim i As Long, p As Long
Dim nl As Long, count As Long
Dim isFound As Boolean


nl = UBound(tabref, 1)

ReDim lstUnique(1 To nl, 1 To 2)


lstUnique(1, 1) = 1 'idx
lstUnique(1, 2) = tabref(1, colref) 'valor
count = 1

For i = 2 To nl
    isFound = False
    
    For p = 1 To count
        If tabref(i, colref) = lstUnique(p, 2) Then
            isFound = True
            Exit For
        End If
    Next p
    
    If isFound = False Then
        count = count + 1
        lstUnique(count, 1) = i 'idx
        lstUnique(count, 2) = tabref(i, colref)  'valor
    End If
Next i


If count > 1 Then
    redimNaoVazio lstUnique, 1, lstUnique
    listaUnicos = True
Else
    listaUnicos = False
End If



    
    
End Function

Sub reindexa(tabref As Variant, lstIdx As Variant, colIdx As Long, out As Variant)
'Reindexa a matriz tabref com base nos indices presentes em lstIdx(colIdx)  e retorna o resultado em out

ReDim out(1 To UBound(tabref, 1), 1 To UBound(tabref, 2))

Dim i As Long, j As Long, c As Long
Dim idx As Long


For i = 1 To UBound(lstIdx, 1)
    If lstIdx(i, colIdx) > 0 And lstIdx(i, colIdx) <> "" Then
        idx = lstIdx(i, colIdx)
        c = c + 1
        For j = 1 To UBound(tabref, 2)
            out(c, j) = tabref(idx, j)
        Next j
    End If
Next i


redimlinha out, c


End Sub


Sub somarconcat(ByVal tabref As Variant, ByVal e_colfiltro As Variant, ByVal e_colsoma As Variant, ByVal e_colconcat As Variant, ByRef out As Variant)
'Recebo uma tabela de referencia
'colunas que são o select
'colunas para somar e concatenar

'=====================================================
'Exemplo
'Dim e_colfiltro
'Dim e_colsoma
'Dim e_colconcat
'ReDim e_colfiltro(1 To 7, 1 To 1)
'ReDim e_colsoma(1 To 7, 1 To 1)
'ReDim e_colconcat(1 To 7, 1 To 1)
'e_colfiltro(2, 1) = 1 'B/F
'e_colfiltro(3, 1) = 1 'Largura
'e_colfiltro(7, 1) = 1 'Prioritario
'
'e_colsoma(6, 1) = 1 'Qde a programar
'somarconcat e_carteiraProgramar, e_colfiltro, e_colsoma, e_colconcat, outswap
'=====================================================

Dim chave() As String
Dim i As Long, j As Long, k As Long
Dim Repetido As Boolean
Dim linhaIdx As Long, chaveidx As String
Dim counter As Long

ReDim chave(1 To UBound(tabref, 1), 1 To 1)
ReDim out(1 To UBound(tabref, 1), 1 To UBound(tabref, 2))


counter = 1

For i = 1 To UBound(tabref, 1)
    Repetido = False
    chaveidx = Empty
    
    'Monta a chave
    For j = 1 To UBound(e_colfiltro, 1)
        If e_colfiltro(j, 1) = 1 Then
            chaveidx = chaveidx & "." & tabref(i, j)
        End If
    Next j
    
    'Verifica se a chave já foi utilizada
    If i = 1 Then
        chave(counter, 1) = chaveidx
        For j = 1 To UBound(tabref, 2)
            out(counter, j) = tabref(i, j)
        Next j
        counter = counter + 1
    Else
        For k = 1 To counter - 1
            If chave(k, 1) = chaveidx Then
                Repetido = True
                linhaIdx = k
                Exit For
            End If
        Next k
        
        If Repetido = True Then
            For j = 1 To UBound(tabref, 2)
                If e_colsoma(j, 1) = 1 Then
                    out(linhaIdx, j) = out(linhaIdx, j) + tabref(i, j)
                ElseIf e_colconcat(j, 1) = 1 Then
                    out(linhaIdx, j) = out(linhaIdx, j) & ", " & tabref(i, j)
                End If
            Next j
            
        Else
            chave(counter, 1) = chaveidx
            For j = 1 To UBound(tabref, 2)
                out(counter, j) = tabref(i, j)
            Next j
            counter = counter + 1
        End If
    End If
    
Next i

End Sub



Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

Sub QuickSortArray(ByRef Arr As Variant, ByVal Lo As Long, ByVal Hi As Long, ByVal colref As Long, ByVal Direcao As String)
Dim i As Long, j As Long, varPivot As Variant
Dim idxLo() As Long, idxEq() As Long, idxHi() As Long
Dim countLo As Long, countEq As Long, countHi As Long
Dim swapArr As Variant, ncols As Long

'Guarda índices das posicoes low, equal, high
ReDim idxLo(1 To Hi - Lo + 1, 1 To 1)
ReDim idxEq(1 To Hi - Lo + 1, 1 To 1)
ReDim idxHi(1 To Hi - Lo + 1, 1 To 1)

ncols = UBound(Arr, 2)
ReDim swapArr(1 To Hi - Lo + 1, 1 To ncols)


varPivot = Arr((Lo + Hi) \ 2, colref)
For i = Lo To Hi
    If varPivot > Arr(i, colref) Then
        countLo = countLo + 1
        idxLo(countLo, 1) = i
    ElseIf varPivot = Arr(i, colref) Then
        countEq = countEq + 1
        idxEq(countEq, 1) = i
    Else
        countHi = countHi + 1
        idxHi(countHi, 1) = i
    End If
Next i


'MOnta a matriz de swap
If Direcao = "Crescente" Then
    'Escreve lo
    If countLo > 0 Then
        For i = 1 To countLo
            For j = 1 To ncols
                swapArr(i, j) = Arr(idxLo(i, 1), j)
            Next j
        Next i
    End If
    'Escreve Equal
    If countEq > 0 Then
        For i = 1 To countEq
            For j = 1 To ncols
                swapArr(i + countLo, j) = Arr(idxEq(i, 1), j)
            Next j
        Next i
    End If
    'Escreve Hi
    If countHi > 0 Then
        For i = 1 To countHi
            For j = 1 To ncols
                swapArr(i + countLo + countEq, j) = Arr(idxHi(i, 1), j)
            Next j
        Next i
    End If
Else
    'Escreve Hi
    If countHi > 0 Then
        For i = 1 To countHi
            For j = 1 To ncols
                swapArr(i, j) = Arr(idxHi(i, 1), j)
            Next j
        Next i
    End If
    'Escreve Equal
    If countEq > 0 Then
        For i = 1 To countEq
            For j = 1 To ncols
                swapArr(i + countHi, j) = Arr(idxEq(i, 1), j)
            Next j
        Next i
    End If
    'Escreve lo
    If countLo > 0 Then
        For i = 1 To countLo
            For j = 1 To ncols
                swapArr(i + countEq + countHi, j) = Arr(idxLo(i, 1), j)
            Next j
        Next i
    End If

End If
'Remonta a matriz de acordo com a especificacao
For i = 1 To Hi - Lo + 1
    For j = 1 To ncols
        Arr(i + Lo - 1, j) = swapArr(i, j)
    Next j
Next i

'proxima iteracao
If Direcao = "Crescente" Then
    If countLo > 1 Then
        QuickSortArray Arr, Lo, Lo - 1 + countLo, colref, Direcao
    End If
    
    If countHi > 1 Then
        QuickSortArray Arr, Lo - 1 + countLo + countEq + 1, Lo - 1 + countLo + countEq + countHi, colref, Direcao
    End If
Else
    If countHi > 1 Then
        QuickSortArray Arr, Lo, Lo - 1 + countHi, colref, Direcao
    End If
    
    If countLo > 1 Then
        QuickSortArray Arr, Lo - 1 + countHi + countEq + 1, Lo - 1 + countLo + countEq + countHi, colref, Direcao
    End If
End If

End Sub


Sub testebuscaordenada()


Dim m1 As Variant

m1 = Range("d2:d30002")

QuickSortCol m1, 1, 31, 1


Range("n1") = buscaOrdenada(m1, 1, -998, 1, 31)

End Sub

Function buscaOrdenada(ByRef tabref As Variant, ByRef colref As Long, ByRef strvalue As Double, ByRef idxmin As Long, ByRef idxMax As Long)
'Faz busca num conjunto ordenado, em colRef, e retorna o index
'Para ordenar, fazer quicksortcol
'Algoritmo de busca quadratico

Dim idx As Long, idxvalue As Double

idx = (idxmin + idxMax) \ 2
idxvalue = tabref(idx, colref)

If idxMax - idxmin = 1 Then
    If tabref(idxmin, colref) = strvalue Then
        buscaOrdenada = idxmin
    ElseIf tabref(idxMax, colref) = strvalue Then
        buscaOrdenada = idxMax
    Else
        buscaOrdenada = Null
    End If
Else
    If idxvalue >= strvalue Then
        buscaOrdenada = buscaOrdenada(tabref, colref, strvalue, idxmin, idx)
    Else
        buscaOrdenada = buscaOrdenada(tabref, colref, strvalue, idx, idxMax)
    End If
End If
End Function



Private Sub RemoveComments()
'Remove todos os comentarios da planilha

Dim sht As Worksheet
Dim rng As Range

For Each sht In Worksheets
If sht.Comments.count > 0 Then
sht.Activate
Cells.ClearComments

End If
Next


Function comparaMin(ByVal a As Variant, ByVal b As Variant) As Variant


If a <= b Then
    comparaMin = a
Else
    comparaMin = b
End If

End Function
Function comparaMax(ByVal a As Variant, ByVal b As Variant) As Variant


If a >= b Then
    comparaMax = a
Else
    comparaMax = b
End If

End Function


Sub indexaRepetidos(ByVal tabref As Variant, ByVal colref As Long, ByRef outIdx As Variant)
'Indexa repetidos da matriz e retorna valor no formato idx | 1
'Ex. 1 | 1; 2 | 1; 3 |1
'se colref = 0, faz para matriz inteira

Dim i As Long, j As Long, k As Long
Dim chave As Variant, chaveFilt As Variant
Dim nl As Long


nl = UBound(tabref, 1)
ReDim chave(1 To nl, 1 To 1)

If colref = 0 Then 'indexa todas as colunas
    For i = 1 To nl
        For j = 1 To UBound(tabref, 2)
            chave(i, 1) = chave(i, 1) & "|" & tabref(i, j)
        Next j
    Next i
Else
    For i = 1 To nl
        chave(i, 1) = tabref(i, colref)
    Next i
End If

'Tira repetidos
selectDistinctSimples chave, 1, chaveFilt

ReDim outIdx(1 To nl, 1 To 2)
For i = 1 To nl
    outIdx(i, 1) = i
    For j = 1 To UBound(chaveFilt, 1)
        If chave(i, 1) = chaveFilt(j, 1) Then
            outIdx(i, 2) = j
            Exit For
        End If
    Next j
Next i

End Sub



Sub redimNaoVazio(ByVal tabref As Variant, ByVal colref As Long, ByRef outtab As Variant)
'Redimensiona a matriz tirando quem é vazio

Dim i As Long, count As Long, j As Long
Dim c2 As Long

count = 0

For i = 1 To UBound(tabref, 1)
    If tabref(i, colref) <> "" Then
        count = count + 1
    End If
Next i

c2 = 0
If count > 0 Then
    ReDim outtab(1 To count, 1 To UBound(tabref, 1))
    
    For i = 1 To UBound(tabref, 1)
        If tabref(i, colref) <> "" Then
            c2 = c2 + 1
            
            For j = 1 To UBound(tabref, 2)
                outtab(c2, j) = tabref(i, j)
            Next j
        End If
    Next i
    
Else
    outtab = Empty
End If
    

End Sub

Sub redimNaoZero(ByVal tabref As Variant, ByVal colref As Long, ByRef outtab As Variant)
'Redimensiona a matriz tirando quem é vazio

Dim i As Long, count As Long, j As Long

ReDim outtab(1 To UBound(tabref, 2), 1 To UBound(tabref, 1))

count = 0

For i = 1 To UBound(tabref, 1)
    If tabref(i, colref) <> 0 Then
        count = count + 1
        For j = 1 To UBound(tabref, 2)
            outtab(j, count) = tabref(i, j)
        Next j
    End If
Next i

ReDim Preserve outtab(1 To UBound(tabref, 2), 1 To count)

transpor outtab, outtab
    

End Sub

'Append Cols
Sub appendCols(ByVal m1 As Variant, ByVal m2 As Variant, ByRef out As Variant)
'(Cola m2 no final de m1, mesmas linhas)
Dim i As Long
Dim tam1 As Long, tam2 As Long


tam1 = UBound(m1, 2)
tam2 = UBound(m2, 2)


'Redimensiona out
ReDim out(1 To UBound(m1, 1), 1 To tam1 + tam2)


For i = 1 To UBound(m1, 1)
 For j = 1 To tam1
   out(i, j) = m1(i, j)
Next j
Next i

For i = 1 To UBound(m1, 1)
 For j = tam1 + 1 To tam1 + tam2
  out(i, j) = m2(i, j - tam1)
Next j
Next i

End Sub



'Append lins
Sub appendlins(ByVal m1 As Variant, ByVal m2 As Variant, ByRef out As Variant)
'(Cola m2 no final de m1, mesmas colunas)
Dim i As Long, j As Long
Dim tam1 As Long, tam2 As Long

If Information.IsArray(m1) Then

    tam1 = UBound(m1, 1)
    tam2 = UBound(m2, 1)
    
    'Copia m1 em out
    copia m1, out
    
    
    'Redimensiona out
     redimlinha out, tam1 + tam2
    
    'Copia o m2 no final
    For i = tam1 + 1 To tam1 + tam2
        For j = 1 To UBound(m1, 2)
            out(i, j) = m2(i - tam1, j)
        Next j
    Next i

Else
    copia m2, out
End If
End Sub


Sub copia(ByVal m1 As Variant, ByRef out As Variant)
'copia m1 em out

Dim i As Long, j As Long

ReDim out(1 To UBound(m1, 1), 1 To UBound(m1, 2))
For i = 1 To UBound(m1, 1)
For j = 1 To UBound(m1, 2)
 out(i, j) = m1(i, j)
Next j
Next i

End Sub

Sub redimlinha(ByRef m1 As Variant, ByVal nlin As Long)

transpor m1, m1
ReDim Preserve m1(1 To UBound(m1, 1), 1 To nlin)
transpor m1, m1

End Sub




Sub extraicol(ByVal m1, ByVal ncol, ByRef out)
'Extrai uma col da matriz
Dim i As Long

ReDim out(1 To UBound(m1, 1), 1 To 1)
For i = 1 To UBound(m1, 1)
 out(i, 1) = m1(i, ncol)
Next

End Sub


Sub insereLinha(ByVal m1, ByVal nlin, ByRef out, ByVal nlinout)
'Insere a linha nlin de m1 em out (linha nlinout)

Dim j As Long

For j = 1 To UBound(m1, 2)
    out(nlinout, j) = m1(nlin, j)
Next j
End Sub

Sub matrizCabecaBaixo(ByVal tabref As Variant, ByRef out As Variant)
Dim i As Long, j As Long
Dim maxlin As Long

maxlin = UBound(tabref, 1)
ReDim out(1 To maxlin, 1 To UBound(tabref, 2))



For i = 1 To maxlin
    For j = 1 To UBound(tabref, 2)
        out(maxlin - i + 1, j) = tabref(i, j)
    Next j
Next i

End Sub


Sub pivotTAB(ByVal e_dados As Variant, ByVal e_colLin As Variant, ByVal e_colCol As Variant, ByVal e_colsoma As Long, ByRef out_dados As Variant, ByRef ref_chavelin As Variant, ByRef ref_chavecol As Variant)
'dim e_colLin as variant
'dim e_colCol as variant
'dim e_colSoma as long

'redim e_colLin(1 to ubound(e_dados,2), 1 to 1)
'redim e_colCol(1 to ubound(e_dados,2), 1 to 1)
'redim e_colSoma(1 to ubound(e_dados,2), 1 to 1)
'
'e_colLin(1, 1) = 1
'e_colLin(3, 1) = 1
'e_colcol(2, 1) = 1
'e_colsoma(4, 1) = 1

'dynamicTAB e_dados , e_colLin , e_colcol , e_colsoma , out
Dim i As Long, j As Long, k As Long
Dim chaveLin As Variant
Dim chavecol As Variant
Dim nchavecol As Long, nchavelin As Long
Dim idxlin As Long, idxCol As Long
Dim c_col As Long, c_lin As Long
Dim out_chavelin, out_chavecol
Dim swap As Variant

For j = 1 To UBound(e_dados, 2)
    If e_colLin(j, 1) = 1 Then
       nchavelin = nchavelin + 1
    ElseIf e_colCol(j, 1) = 1 Then
       nchavecol = nchavecol + 1
    End If
Next j
   
'Cria chaves
ReDim chaveLin(1 To UBound(e_dados, 1), 1 To 1)
ReDim chavecol(1 To UBound(e_dados, 1), 1 To 1)

For i = 1 To UBound(e_dados, 1)
    For j = 1 To UBound(e_dados, 2)
        If e_colLin(j, 1) = 1 Then
            chaveLin(i, 1) = chaveLin(i, 1) & ".!-" & e_dados(i, j)
        ElseIf e_colCol(j, 1) = 1 Then
            chavecol(i, 1) = chavecol(i, 1) & ".!-" & e_dados(i, j)
        End If
    Next j
Next i


'faz select distict
selectDistinctSimples chaveLin, 1, out_chavelin
selectDistinctSimples chavecol, 1, out_chavecol


ReDim out_dados(1 To UBound(out_chavelin, 1), 1 To UBound(out_chavecol, 1))
ReDim ref_chavelin(1 To UBound(out_chavelin, 1), 1 To nchavelin)
ReDim ref_chavecol(1 To UBound(out_chavecol, 1), 1 To nchavecol)

'Preenche dados
For i = 1 To UBound(e_dados, 1)
    idxlin = buscaSimplesIdx(out_chavelin, 1, Conversion.CStr(chaveLin(i, 1)))
    idxCol = buscaSimplesIdx(out_chavecol, 1, Conversion.CStr(chavecol(i, 1)))
    out_dados(idxlin, idxCol) = out_dados(idxlin, idxCol) + e_dados(i, e_colsoma)
Next i

'Preenche chave col e chave lin
For i = 1 To UBound(out_chavelin, 1)
    swap = Strings.Split(out_chavelin(i, 1), ".!-")
    For j = 1 To UBound(swap, 1)
        ref_chavelin(i, j) = swap(j)
    Next j
Next i


For i = 1 To UBound(out_chavecol, 1)
    swap = Strings.Split(out_chavecol(i, 1), ".!-")
    For j = 1 To UBound(swap, 1)
        ref_chavecol(i, j) = swap(j)
    Next j
Next i

End Sub











