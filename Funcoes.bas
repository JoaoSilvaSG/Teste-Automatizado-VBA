Attribute VB_Name = "Funcoes"
' Localizar o cursor do mouse (user32) - dll
Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long
' Acessar a função de localização em user32.dll
Declare Function SetCursorPos Lib "user32" _
(ByVal x As Long, ByVal y As Long) As Long

Public Declare Sub mouse_event Lib "user32" _
(ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'Incluir a variável POINT - API excel

Type POINTAPI
    X_Pos As Long
    Y_Pos As Long
End Type

Option Explicit
 
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1
Public resolucaoX As Integer
Public resolucaoY As Integer

Declare Sub sleep Lib "kernel32" _
 Alias "Sleep" (ByVal dwMilliseconds As Long)


'Definição do localizador

Sub Get_Cursor_Pos()

' Dimensão x e y do cursor
Dim Hold As POINTAPI

' Criar uma localização varíavel base rotina de Progressão crescente/decrescente
GetCursorPos Hold

' Mostrar as coordenadas se necessário (este não será utilizado)
MsgBox "X Position is : " & Hold.X_Pos & Chr(10) & _
"Y Position is : " & Hold.Y_Pos
End Sub

Sub tempoIrrelevante()
 sleep 500   'Faz o código esperar por 0.5 segundos
End Sub

Sub tempoPequeno()
    Application.Wait (Now + TimeValue("0:00:02"))
End Sub

Sub tempoMedio()
        Application.Wait (Now + TimeValue("0:00:04"))
End Sub

Sub RightClick(x As Integer, y As Integer)
    Call CalcularXY(x, y)
    
    SetCursorPos x, y
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Sub LeftClick(x As Integer, y As Integer, Optional ByVal asSleep As Boolean = True)
    Call CalcularXY(x, y)
    
    If asSleep Then sleep 400
    SetCursorPos x, y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    If asSleep Then sleep 400
End Sub


Sub Cabecalho(x As Integer, y As Integer)
    SetCursorPos x, y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Call tempoIrrelevante
End Sub

Sub PassarFoco()
    Windows(nomePlanilha).Activate
    Sheets("Principal").Select
    sleep 200
End Sub

Sub VerificarDados(valor, posicao, tela As String)

        If (CStr(Sheets("Dados para verificação").Range(posicao).Value) <> valor) Then
        
            'Aqui irá verificar se tem zero depois da virgula
            On Error GoTo erronaconversao
            If (Round(CDec(CStr(Sheets("Dados para verificação").Range(posicao).Value)), 2) = Round(CDec(valor), 2)) Then Exit Sub
erronaconversao:
            On Error GoTo 0
        
            Call PassarFoco
            Call atualizarErro("Ocorreu algum erro com " & Sheets("Dados para verificação").Range(Left(posicao, 1) & "1").Value & " da " & tela & vbCrLf & "Valor esperado: " & valor & vbCrLf & "Valor obtido:" & Sheets("Dados para verificação").Range(posicao).Value)
            SendKeys "{NUMLOCK}"
            MsgBox "Ocorreu algum erro com " & Sheets("Dados para verificação").Range(Left(posicao, 1) & "1").Value & " da tela de " & tela & vbCrLf & "Valor esperado: " & valor & vbCrLf & "Valor obtido:" & Sheets("Dados para verificação").Range(posicao).Value
            End
        End If
End Sub

Sub VerificarFoco(valor, campo As String, Optional ByVal SelecionarTudo As Boolean = False)
    'No agro o Ctrl+C não funciona, logo é necessário utilizado o lado direito do mouse + copiar
    SendKeys "+({F10})"
    Application.Wait (Now + TimeValue("0:00:01"))
    SendKeys "c"
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Windows(nomePlanilha).Activate
    Sheets("Dados para verificação").Select
    Range("A1").Select '--selecione a célula onde deseja colar
    ActiveSheet.Paste '-- código para colar
    Call PassarFoco
    
    If (Sheets("Dados para verificação").Range("A1").Value <> valor) Then
            'Aqui irá verificar se tem zero depois da virgula
            On Error GoTo erronaconversao
            If (Round(CDec(CStr(Sheets("Dados para verificação").Range("A1").Value)), 2) = Round(CDec(valor), 2)) Then Exit Sub
erronaconversao:
            On Error GoTo 0
    
            Call PassarFoco
            Call atualizarErro("O campo " & campo & " não está com o valor esperado." & "Valor esperado: " & valor)
            SendKeys "{NUMLOCK}"
            MsgBox "O campo " & campo & " não está com o valor esperado." & "Valor esperado: " & valor
            End
    End If
    '---- limpando os dados
    Call LimparDados
End Sub

Function VerificarFocoRelatorio(valor, campo As String) As Boolean
    VerificarFocoRelatorio = True
    'No agro o Ctrl+C não funciona, logo é necessário utilizado o lado direito do mouse + copiar
    SendKeys "+({F10})"
    Application.Wait (Now + TimeValue("0:00:01"))
    SendKeys "c"
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Windows(nomePlanilha).Activate
    Sheets("Dados para verificação").Select
    Range("A1").Select '--selecione a célula onde deseja colar
    ActiveSheet.Paste '-- código para colar
    
    If (Sheets("Dados para verificação").Range("A1").Value <> valor) Then
            'Aqui irá verificar se tem zero depois da virgula
            On Error GoTo erronaconversao
            If (Round(CDec(CStr(Sheets("Dados para verificação").Range("A1").Value)), 2) = Round(CDec(valor), 2)) Then Exit Function
erronaconversao:
            On Error GoTo 0
                VerificarFocoRelatorio = False
                Exit Function
            End
    End If
    '---- limpando os dados
    VerificarFocoRelatorio = True
    Call LimparDados
End Function

Sub CopiarDados(x As Integer, y As Integer)
    Call CalcularXY(x, y)
    
    Application.CutCopyMode = False '--código para tirar a seleção de cópia (ganho na memória da máquina)
    Call RightClick(x, y)
    SendKeys ("T") 'copia os dados para verificação
    Application.Wait (Now + TimeValue("0:00:01"))
    
        Windows(nomePlanilha).Activate
        Sheets("Dados para verificação").Select
        Range("A1").Select '--selecione a célula onde deseja colar
        ActiveSheet.Paste '-- código para colar
End Sub

Sub CopiarLinha(Optional ByVal linha As Integer = 1, Optional ByVal coluna As Integer = 5)
    Dim row As Integer
    Dim col As Integer
    Dim erros As Integer
    row = 1
    
    Application.CutCopyMode = False '--código para tirar a seleção de cópia (ganho na memória da máquina)

    sleep 200
    While row <= linha
        col = 1
        
        While col <= coluna
            erros = 0
            SendKeys "^C"
            sleep 300

            Windows(nomePlanilha).Activate
            Sheets("Dados para verificação").Select
            Cells(row, col).Select '--selecione a célula onde deseja colar
            On Error Resume Next
            Do
                sleep 100
                Err.Clear
                ActiveSheet.Paste '-- código para colar
                sleep 100
                erros = erros + 1
                If erros = 40 Then Exit Sub
                If erros = 20 Then SendKeys "^C"
            Loop Until (Err.Number = 0)
            On Error GoTo 0

            col = col + 1
            SendKeys "{Right}"
            sleep 200
        Wend
        SendKeys "^{Left}"
        SendKeys "{Down}"
        row = row + 1
        sleep 200
    Wend
End Sub

Sub LimparDados()
    Windows(nomePlanilha).Activate
    Sheets("Dados para verificação").Select
    Sheets("Dados para verificação").Range("A1:AZ12").Select
    Selection.ClearContents
    Selection.NumberFormat = "@"
End Sub

Sub Atualizar(porcentagem As Double, atual As String)
    Dim MyStr
    MyStr = Format(porcentagem, "##.##")
    Windows(nomePlanilha).Activate
    Sheets("Principal").Select
    
    Sheets(1).Range("J18").Value = atual
    Sheets(1).Range("J19").Value = MyStr & "%"
End Sub

Sub AtualizarSmoke(porcentagem As Double, atual As String)
    Dim MyStr
    MyStr = Format(porcentagem, "##.##")

    Windows(nomePlanilha).Activate
    Sheets("Principal").Select
    
    Sheets(1).Range("J9").Value = atual
    Sheets(1).Range("J10").Value = MyStr & "%"
    
    If (porcentagem = 0) Then
        Sheets(1).Range("J10").Value = "0,00%"
    End If
End Sub

Sub atualizarErro(mensagem As String)
    On Error Resume Next
    
    Dim contador As Integer
    Dim rastro As String
    contador = 2
    
    While (Sheets("Relatório").Range("A" & contador) <> "")
        contador = contador + 1
    Wend
    Call LeftClick(2, 2)
    Sheets("Relatório").Range("A" & contador) = mensagem
    
    If (Sheets("Principal").Range("j10").Value <> "100,%") Then
        Sheets("Relatório").Range("B" & contador) = Sheets("Principal").Range("j9").Value
    Else
        Sheets("Relatório").Range("B" & contador) = Sheets("Principal").Range("j18").Value
    End If
    Call LeftClick(0, 0)
    Application.Wait (Now + TimeValue("0:00:01"))
    SendKeys "+(^{F6})"
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Sheets("Dados para verificação").Range("A12").Value = ""
    Sheets("Dados para verificação").Select
    Range("A12").Select
    ActiveSheet.Paste '-- código para colar
    
    rastro = Sheets("Dados para verificação").Range("A12").Value
    Sheets("Relatório").Range("c" & contador) = rastro
    Sheets("Dados para verificação").Range("A12").Value = ""
    
    Call PassarFoco
    On Error GoTo 0
End Sub

Sub VerificarRelatorio(busca As String, tela As String)

    SendKeys "{INSERT}"
    Application.Wait (Now + TimeValue("0:00:06"))
    Sheets("Banco de Dados").Range("L1").Copy
    
    If (busca = "") Then
        busca = "Sem Busca"
    Else
        Call LeftClick(50, 80, False)
        Application.Wait (Now + TimeValue("0:00:02"))
        preencher busca, 5
    End If
    
    Call LeftClick(74, 1020)
    Call tempoIrrelevante
    SendKeys "+({F10})"
    Application.Wait (Now + TimeValue("0:00:01"))
    SendKeys "s"
    Application.Wait (Now + TimeValue("0:00:01"))

    If VerificarFocoRelatorio("1", tela & "Relatório Aberto") Then
        SendKeys "{ESC}" 'Saindo do relatorio
        Call tempoPequeno
        Call PassarFoco
        Exit Sub
    End If
    
    SendKeys "{ESC}" 'Saindo do relatorio
    Call tempoPequeno
    Call PassarFoco
    Call atualizarErro("Ocorreu algum erro com o relatório de " & tela & vbCrLf & "Possivelmente o relatório não foi aberto ou a informação não foi buscada." & vbCrLf & "Valor buscado: " & busca)
    SendKeys "{ESC}" 'Caso de erro ou não encontre o valor, irá sair
    Call tempoPequeno
End Sub
    

Sub VerificaRastro(tela As String)
    Call Shell("cmd.exe /S /C" & "echo off | clip", vbHide)  'limpa a area de trasferencia
    
    Dim erros As Integer
    erros = 0
    Call tempoPequeno
    SendKeys "+(^{F6})"
    sleep 500
    
    Windows(nomePlanilha).Activate
    Sheets("Dados para verificação").Select
    Sheets("Dados para verificação").Range("A1").Value = ""
    Range("A1").Select
   
    On Error Resume Next
    Do
        Err.Clear
        ActiveSheet.Paste '-- código para colar
        sleep 300
        erros = erros + 1
        If (Sheets("Dados para verificação").Range("A1").Value = "") And (erros = 25) Then SendKeys "+(^{F6})"
    Loop Until (CStr(Sheets("Dados para verificação").Range("A1").Value) <> "") Or (erros = 50)
    Call PassarFoco
    sleep 300

    If (erros = 50) Or (CStr(Sheets("Dados para verificação").Range("A1").Value) <> tela) Then
        Call atualizarErro("Não foi encontrado a tela " & tela)
        Application.Wait (Now + TimeValue("0:00:01"))
        SendKeys "{NUMLOCK}"
        MsgBox "Aparentemente não está com a tela de rastro " & tela & " aberto, favor, verificar!"
        End
    End If
End Sub

Sub VerificaImagem(x As Integer, y As Integer, w As Integer, z As Integer, arquivo As String)
    Call CalcularXY(x, y)
    Call CalcularXY(w, z)
    
    Dim cont As Integer
    
    cont = 0
    Call tempoPequeno
    
    If Dir("C:\Realtec\Agro\Dados\SmokeTest\Teste.jpg") <> "" Then
        Kill "C:\Realtec\Agro\Dados\SmokeTest\Teste.jpg"
    End If

    Call LimparDados

    Do
        Sheets("Banco de Dados").Range("L1").Copy
        Application.SendKeys "({1068})"
        Call tempoPequeno
        SendKeys "+(^{F6})"
        Application.Wait (Now + TimeValue("0:00:01"))

            Windows(nomePlanilha).Activate
            Sheets("Dados para verificação").Select
            Range("A1").Select '--selecione a célula onde deseja colar
            ActiveSheet.Paste '-- código para colar
            
            cont = cont + 1
        If (cont > 10) Then
            MsgBox "Não foi possível abrir o Lightshot"
            End
        End If
    Loop Until (Sheets("Dados para verificação").Range("A1").Value = "zzzzzzzz")
    
    Call tempoPequeno
    Call Click_Move(x, y, w, z)
    
    Call tempoIrrelevante
    SetCursorPos w - 44, z + 13
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Call tempoPequeno
    SendKeys "        C:\Realtec\Agro\Dados\SmokeTest\Teste.jpg"
    Application.Wait (Now + TimeValue("0:00:01"))
    SendKeys "{ENTER}"
    
    Call tempoMedio
    
    Dim caminho As String
    Dim fso
    
    caminho = CStr(ThisWorkbook.Path)
    caminho = caminho & "\Imagens\" & Environ("Computername") & "\"

    If (Dir(caminho, vbDirectory) = vbNullString) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFolder CStr(ThisWorkbook.Path) & "\Imagens\Padrao", CStr(ThisWorkbook.Path) & "\Imagens\" & Environ("Computername")
    End If

    If (CompararImagens(caminho & arquivo & ".jpg", "C:\Realtec\Agro\Dados\SmokeTest\Teste.jpg") = False) Then
        Sheets("Dados para verificação").Range("A1").Value = caminho & arquivo & ".jpg"
        Call atualizarErro("As imagens não são compatíveis" & vbCrLf & "Primeira imagem se encontra em:" & vbCrLf & "C:\Realtec\Agro\Dados\SmokeTest\Teste.jpg" _
        & vbCrLf & "Segunda imagem se encontra em:" & vbCrLf & caminho & arquivo & ".jpg")
        
        Call PassarFoco
        UserFormImagen.Show
        Call PassarFoco
        SendKeys "{NUMLOCK}"
        End
    End If
    
    Call LeftClick(700, 2) 'volta o foco para o Agro
    Call tempoPequeno
    Call LimparDados
    Call PassarFoco
End Sub

Sub Click_Move(x As Integer, y As Integer, w As Integer, z As Integer)
    SetCursorPos x, y     'Local do Clique

    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
     
    SetCursorPos w, z    'Local do Destino, em X e Y.
     
    Application.Wait Now + TimeSerial(0, 0, 1)   'Aguarda um segundo antes do próximo passo.
     
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Function CompararImagens(strFilename1 As String, strFilename2 As String) As Boolean
    Dim byt1() As Byte
    Dim byt2() As Byte
    Dim f1 As Integer
    Dim f2 As Integer
    Dim lngFileLen1 As Long
    Dim lngFileLen2 As Long
    Dim i As Long
    
    Dim cont As Integer
    cont = 0

    'Test to see if we have actually been passed 2 filenames
    If LenB(strFilename1) = 0 Or LenB(strFilename2) = 0 Then Exit Function
    
    'Verifica se os arquivos existem
    If LenB(Dir(strFilename2)) = 0 Then Exit Function
    
    If LenB(Dir(strFilename1)) = 0 Then
        FileCopy strFilename2, strFilename1
        atualizarErro ("Arquivo não existia e foi salvo com sucesso. Caminho arquivo:" & strFilename1)
        CompararImagens = True
        
        Exit Function
    End If

    'OK now start looking at the file contents
    f1 = FreeFile
    Open strFilename1 For Binary Access Read As #f1
    f2 = FreeFile
    Open strFilename2 For Binary Access Read As #f2
    lngFileLen1 = LOF(f1)
    lngFileLen2 = LOF(f2)

    If lngFileLen1 = lngFileLen2 Then
      'Continue - there is a possibility they are the same
      ReDim byt1(1 To lngFileLen1) As Byte
      ReDim byt2(1 To lngFileLen2) As Byte
      Get #f1, , byt1
      Get #f2, , byt2
      For i = 1 To lngFileLen1
        If byt1(i) <> byt2(i) Then
            'cont = cont + 1
            'If (cont = ) Then
            GoTo IdenticalFiles_Exit 'The 2 files are not the same
        End If
      Next
      'We got this far so the 2 files must be the same
      CompararImagens = True
    End If

IdenticalFiles_Exit:
    Close #f1
    Close #f2
End Function

Function VerifyScreenResolution() As Boolean
    VerifyScreenResolution = False
    
    resolucaoX = GetSystemMetrics(SM_CXSCREEN)
    resolucaoY = GetSystemMetrics(SM_CYSCREEN)
    
    If (resolucaoX = 1920) And (resolucaoY = 1080) Then VerifyScreenResolution = True
End Function

Public Sub CalcularXY(ByRef x As Integer, ByRef y As Integer)
    If calcularNovaResolucao Then
        x = (resolucaoX / 2) + (x - 960) '960/540 pois é a metade de 1920/1080 que é a resolução padrão
        y = (resolucaoY / 2) + (y - 540)
    End If
End Sub

Public Sub preencher(x As String, Optional ByVal QuantidadeEnter As Integer = 1, Optional ByVal EnterAnterior As Integer = 0)
    Dim cont As Integer
    cont = 0

    While cont < EnterAnterior
        sleep 400
        SendKeys "{ENTER}"
        cont = cont + 1
    Wend
    
    cont = 0
    
    While cont > EnterAnterior
        sleep 400
        SendKeys "+{TAB}"
        cont = cont - 1
    Wend

    sleep 400
    SendKeys x
    sleep 400
    
    cont = 0

    While cont < QuantidadeEnter
        SendKeys "{ENTER}"
        sleep 400
        cont = cont + 1
    Wend
    
    cont = 0
    
    While cont > QuantidadeEnter
        SendKeys "+{TAB}"
        sleep 400
        cont = cont - 1
    Wend
End Sub


Public Sub preencherMes(x As String, Optional ByVal QuantidadeEnter As Integer = 1, Optional ByVal EnterAnterior As Integer = 0)
    If Len(Day(x)) = 1 Then
        preencher "0" & Day(x), 0, EnterAnterior
    Else
        preencher Day(x), 0, EnterAnterior
    End If
    
    If Len(Month(x)) = 1 Then
        preencher "0" & Month(x), 0
    Else
        preencher Month(x), 0
    End If
    
    preencher Year(x), QuantidadeEnter
End Sub

Sub VerificarDadosLinha(tela, texto As String, Optional ByVal row As Integer = 1)
    Dim col As Integer
    Dim valores() As String
    valores = Split(texto, ";")
    col = 0

    While col < WorksheetFunction.CountA(valores)
        If (RTrim(CStr(Worksheets("Dados para verificação").Cells(row, col + 1).Value)) <> RTrim(CStr(valores(col)))) And _
            (CStr(valores(col)) <> "Empty") Then
            'Aqui irá verificar se tem zero depois da virgula
            On Error GoTo erronaconversao
            If (Round(CDec(CStr(Worksheets("Dados para verificação").Cells(row, col + 1).Value)), 2) = Round(CDec(valores(col)), 2)) Then GoTo proximo
erronaconversao:
            On Error GoTo 0
        
            Call PassarFoco
            Call atualizarErro("Ocorreu algum erro na tela: " & tela & vbCrLf & "Valor esperado: " & valores(col) & vbCrLf & "Valor obtido: " & Worksheets("Dados para verificação").Cells(row, col + 1).Value & vbCrLf & "Posição: " & (col + 1))
            SendKeys "{NUMLOCK}"
            MsgBox "Ocorreu algum erro na tela: " & tela & vbCrLf & "Valor esperado: " & valores(col) & vbCrLf & "Valor obtido: " & Worksheets("Dados para verificação").Cells(row, col + 1).Value & vbCrLf & "Posição: " & (col + 1)
            End
        End If
proximo:
        col = col + 1
    Wend
    
    LimparDados
    Exit Sub
End Sub

Sub CopiarEVerificar(tela, texto As String)
    Call LimparDados
    Dim valores() As String
    valores = Split(texto, ";")
    
    Call CopiarLinha(1, WorksheetFunction.CountA(valores))
    Call VerificarDadosLinha(tela, texto)
End Sub

Sub AcessarGuiaComplementar(Optional ByVal x As Integer = 570, Optional ByVal y As Integer = 240)
    Call LeftClick(x, y)
End Sub

Function GetSenhaAgro() As String
    Dim dia As String
    Dim hora As String
    
    dia = Day(Date)
    hora = Hour(Now)
    
    If Len(dia) = 1 Then dia = "0" & dia
    If Len(hora) = 1 Then hora = "0" & hora
    
    GetSenhaAgro = "agrors" & dia & hora
End Function

Sub AlterarPadraoPrint(tipoPrint As String)
    Dim siglaTipoPrint As String
    siglaTipoPrint = Left(tipoPrint, 1)

    Dim cont As Integer

    cont = 0

    If Dir("C:\Realtec\Agro\Dados\SmokeTest\Padrao." & tipoPrint) <> "" Then
        Kill "C:\Realtec\Agro\Dados\SmokeTest\Padrao." & tipoPrint
    End If

    Call LimparDados

    Do
        Sheets("Banco de Dados").Range("L1").Copy
        Application.SendKeys "({1068})"
        Call tempoPequeno
        SendKeys "+(^{F6})"
        Application.Wait (Now + TimeValue("0:00:01"))

            Windows(nomePlanilha).Activate
            Sheets("Dados para verificação").Select
            Range("A1").Select '--selecione a célula onde deseja colar
            ActiveSheet.Paste '-- código para colar
            
            cont = cont + 1
        If (cont > 10) Then
            MsgBox "Não foi possível abrir o Lightshot"
            Exit Sub
        End If
    Loop Until (Sheets("Dados para verificação").Range("A1").Value = "zzzzzzzz")
    
    Application.Wait (Now + TimeValue("0:00:01"))
    Call Click_Move(0, 0, 1, 1)
    
    Call tempoPequeno
    SetCursorPos 158, 20
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Call tempoPequeno
    SendKeys "        C:\Realtec\Agro\Dados\SmokeTest\Padrao." & tipoPrint
    Application.Wait (Now + TimeValue("0:00:01"))
    SendKeys "{TAB}"
    Application.Wait (Now + TimeValue("0:00:01"))
    SendKeys siglaTipoPrint
    Application.Wait (Now + TimeValue("0:00:01"))
    SendKeys "{ENTER}"
End Sub

Sub apagarImagem()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim arquivo As String
    
    arquivo = Sheets("Banco de Dados").Range("K35").Value
    
    If arquivo = "" Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(CStr(ThisWorkbook.Path) & "\Imagens\")

    For Each file In folder.subfolders
        If Dir(file.Path & arquivo) <> "" Then Kill file.Path & arquivo
    Next file
End Sub
