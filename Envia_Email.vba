Sub envia_email()

' VARIÁVEIS QUE SERÃO TRABALHADAS

    Dim Wks     As Worksheet
    Dim corpo   As String
    Dim copia   As String
    Dim titulo  As String
    Dim dado1
    Dim dado2
    Dim data
    Dim mail
    Dim linha
    Dim limite

' REFERÊNCIA A PLANILHA ATIVA ATUAL

    Set Wks = ActiveSheet

' DESABILITA RECURSOS PARA MELHOR PERFORMANCE

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

' MENSAGEM DE TEXTO

    MsgBox "OS E-MAILS SERÃO ENVIADOS." & Chr(10) & "PARA CANCELAR, TECLE ESC E FIM"

' PARÂMETROS GENÉRICOS

    titulo = Range("B1")
    corpo = ActiveWorkbook.Worksheets("PARAMETROS").Range("B7").Value
    copia = ActiveWorkbook.Worksheets("PARAMETROS").Range("D4").Value
    pasta_arquivos = ActiveWorkbook.Worksheets("PARAMETROS").Range("D2").Value

' PRIMEIRA E ÚLTIMA LINHA DAS ENTRADAS (em desenvolvimento)

    linha = 7
    limite = Range("G4")
    limite = CInt(limite)

' AGE SOB CADA LINHA ESPECÍFICADA

    For linha = 7 To limite

' EM CASO DE ERRO, CONTINUAR COM PRÓXIMO LAÇO

        On Error GoTo Skip

        If Cells(linha, "f") <> "ENVIADO" Then

' PARAMÊTROS ESPECÍFICOS DO E-MAIL

            Set dado1 = Cells(linha, "b")
            Set dado2 = Cells(linha, "c")
            Set mail = Cells(linha, "d")
            Set data = Cells(linha, "e")

' SELECIONA PLANILHA

            Wks.Select

' MONTA O ASSUNTO DO E-MAIL

            assunto = "CONTROLE DE ENVIO - " & dado1 & " - BOLETO VENCIMENTO - " & data

' CONSTROI OBJETO DO OUTLOOK (PROCESSO, CONTEXTO E E-MAIL)

            Set out = CreateObject("Outlook.Application")
            Set mapi = out.GetNamespace("MAPI")
            Set Email = out.CreateItem(0)

' MONTA ESTRUTURA DO E-MAIL

            Email.To = mail
            Email.cc = copia
            Email.Body = corpo
            Email.Subject = assunto

' ADICIONA ANEXO AO E-MAIL

            Email.Attachments.Add (pasta_boleto & "arquivo.pdf")

' CASO DESEJE ADICIONAR ARQUIVOS ESPECIFICOS AO E-MAILS

            'Email.Attachments.Add (pasta_boleto & dado1 & ".pdf")

' VISUALIZA O E-MAIL ( removível )

            Email.Display

' ENVIA O E-MAIL

            'Email.send

' DELETA VARIAVEIS DO OBJETO DO OUTLOOK

            Set Email = Nothing
            Set mapi = Nothing
            Set out = Nothing

' INDICA ITEMS QUE OBTIVERAM EXÍTO NO PROCESSO

            Cells(linha, "f") = "ENVIADO"

' REATIVA RECURSOS DESABILITADOS

            With Application
                .EnableEvents = True
                .ScreenUpdating = True
            End With

        End If

' CASO OCORRA ERROR, ESSA GAMBIARRA VAI SER EXECUTADA :/

Skip:
    Resume Skip2
Skip2:
    Next

' MENSAGEM DE TEXTO

    MsgBox "Todos os e-mails foram enviados"

End Sub
