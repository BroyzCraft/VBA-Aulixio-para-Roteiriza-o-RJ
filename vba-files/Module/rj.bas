Attribute VB_Name = "rj"
Sub imprimirCortes()
    
form_rj.Hide

' CONFIRMA��O
Dim confirmacao As VbMsgBoxResult
confirmacao = MsgBox("Voc� solicitou a impress�o das capas de corte, Continuar?", vbYesNo)

If confirmacao = vbYes Then
    qtd = Application.InputBox("Digite quantas capas deseja imprimir: (Padr�o: 3)")
    Sheets("rj-capa-corte").Select
    Range("A1:E48").Select
    Selection.PrintOut Copies:=qtd, Collate:=True, Preview:=True
End If

Sheets("rj-menu").Select
    
End Sub

Sub imprimirControle()
    
form_rj.Hide

Dim confirmacao As VbMsgBoxResult
confirmacao = MsgBox("Voc� solicitou a impress�o do controle, Continuar?", vbYesNo)

If confirmacao = vbYes Then

    Dim nome As String
    
    confirmacao = MsgBox("Deseja criar um novo controle? ", vbYesNo)
    Sheets("rj-menu").Select
    nome = Range("B12").Value
    
    If confirmacao = vbYes Then
        Worksheets("rj-controle").Copy After:=Worksheets(1)
        ActiveSheet.Name = nome
    End If
    
    qtd = Application.InputBox("Digite quantas capas deseja imprimir: (Padr�o: 4)")
    Sheets(nome).Select
    Range("A1:J40").Select
    Selection.PrintOut Copies:=qtd, Collate:=True

Else: Exit Sub
End If

confirmacao = MsgBox("Deseja salvar os dados?", vbYesNo)

If confirmacao = vbYes Then
        
    'Consolida os dados
    Sheets(nome).Select
    Range("A1:J40").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1:J40").Select
    
    'Gerar PDF
    strPathNome = "L:\Logistica\Transporte\2_ROUTEASY\0 - ARQUIVOS DA ROTEIRIZA��O (EXCEL)\" & "Resumo RJ - " & nome
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strPathNome, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
    
End If

gerarBackup
Sheets("rj-controle").Select
    
End Sub

Sub gerarBackup()

Dim nome As String
Dim plan As String
Dim macro As String

Sheets("rj-menu").Select
nome = Range("B12").Value
macro = ActiveWorkbook.Name
plan = "01.JANEIRO.xlsx"

Workbooks.Open ("\\Ecfs1\leo\Logistica\Transporte\4_ROTEIRIZACAO\Roteiriza��o TP  RJ\2021\" & plan)

Workbooks(macro).Activate
Sheets(nome).Select
ActiveSheet.Move Before:=Workbooks(plan).Sheets(1)

Workbooks(plan).Close True

End Sub

Sub abreForm()

form_rj.Show

End Sub
