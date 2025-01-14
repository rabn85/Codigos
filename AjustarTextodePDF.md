Sub AjustarQuebrasDeLinhaSelecionado()
    Dim rng As Range
        ' Define o intervalo selecionado
    Set rng = Selection.Range
        ' Remove quebras de linha manuais
    rng.Text = Replace(rng.Text, Chr(11), " ")
        ' Remove quebras de linha extras
    rng.Text = Replace(rng.Text, vbCr, " ")
        ' Remove espaços duplos
    Do While InStr(rng.Text, "  ") > 0
        rng.Text = Replace(rng.Text, "  ", " ")
    Loop
      End Sub
