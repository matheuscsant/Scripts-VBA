Global sapGuiApp As Object
        Global oConnection As Object
        Global Connection As Object
        Global session As Object
        Global sapapplication As Object

Sub Reserva()

If sapapplication Is Nothing Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set sapapplication = SapGuiAuto.GetScriptingEngine
End If
If Connection Is Nothing Then
   Set Connection = sapapplication.Children(0)
End If
If session Is Nothing Then
   Set session = Connection.Children(0)
End If

Range("J8").Select
Selection.End(xlDown).Select
Linhas = ActiveCell.Row
Range("J8").Select

Dim ContadorG As Integer

ContadorG = 8

ContadorSAP = 0

For i = 2 To Linhas
    If Range("J" & ContadorG).Value = Range("P1").Value Then
        Cod_Mat = Range("D" & ContadorG).Text
        Qtd_Mat = Range("E" & ContadorG).Text ' definindo ranges das informações para reserva
        Cod_Dp = Range("G" & ContadorG).Text
        Cod_CC = Range("K4").Text
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/N"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB21" ' transação de criação de reserva
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtRM07M-BWART").Text = "201"
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "3003" ' setando informações na prim pag da reserva
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = Cod_CC ' setando informações para reserva
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-MATNR[" & ContadorSAP & ",7]").Text = Cod_Mat
        session.findById("wnd[0]/usr/sub:SAPMM07R:0521/txtRESB-ERFMG[" & ContadorSAP & ",26]").Text = Qtd_Mat ' setando informações para reserva
        session.findById("wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-LGORT[" & ContadorSAP & ",53]").Text = Cod_Dp
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0 ' prevenção para erro de estoque
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[11]").press ' Salvar
        session.findById("wnd[0]/sbar").DoubleClick
        numresv = Mid(session.findById("wnd[0]/sbar").Text, 12, 12)
        Range("J" & ContadorG).Value = numresv
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzmm_r002" ' transação de impressão de resv
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtS_RSNUM-LOW").Text = Range("J" & ContadorG) ' setando informações para impressão da reserva
        session.findById("wnd[0]/usr/ctxtP_WERKS").Text = "3003"
        session.findById("wnd[0]/tbar[1]/btn[8]").press ' Script roda sozinho a partir daqui
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = ""
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/tbar[1]/btn[2]").press
        session.findById("wnd[1]/tbar[0]/btn[86]").press
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/N"
        session.findById("wnd[0]").sendVKey 0 ' retorna a pag inicial
        ContadorG = ContadorG + 1 ' contador que faz a verificação de cada linha, tentando encontrar FAZER
        ContadorSAP = ContadorSAP + 1 ' contador para linhas durante a reserva
    Else:
        ContadorG = ContadorG + 1
        GoTo Prox
    End If
Prox:
    Next
End Sub
