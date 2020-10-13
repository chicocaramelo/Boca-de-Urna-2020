Attribute VB_Name = "BocaDeUrna2020"

Sub AlgarismosParaNumExtensoAtasBocaDeUrna2020()
Attribute AlgarismosParaNumExtensoAtasBocaDeUrna2020.VB_Description = "Remove números de 0 a 9 alinhados a esquerda, direita ou justificados para tratamento de atas no projeto Boca de Urna 2020."
Attribute AlgarismosParaNumExtensoAtasBocaDeUrna2020.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.RemoverNumerosExcetoCentralizados"
'
' AlgarismosParaNumExtensoAtasBocaDeUrna2020 Macro
' Altera algarismos de 01 a 99 do título das atas para números por extenso para tratamento de atas no projeto Boca de Urna 2020.
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "ATA nº 01/"
    .Replacement.Text = "ATA nº um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 02/"
    .Replacement.Text = "ATA nº dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 03/"
    .Replacement.Text = "ATA nº três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 04/"
    .Replacement.Text = "ATA nº quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 05/"
    .Replacement.Text = "ATA nº cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 06/"
    .Replacement.Text = "ATA nº seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 07/"
    .Replacement.Text = "ATA nº sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 08/"
    .Replacement.Text = "ATA nº oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 09/"
    .Replacement.Text = "ATA nº nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 10/"
    .Replacement.Text = "ATA nº dez/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 11/"
    .Replacement.Text = "ATA nº onze/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 12/"
    .Replacement.Text = "ATA nº doze/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 13/"
    .Replacement.Text = "ATA nº treze/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 14/"
    .Replacement.Text = "ATA nº quatorze/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 15/"
    .Replacement.Text = "ATA nº quinze/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 16/"
    .Replacement.Text = "ATA nº dezesseis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 17/"
    .Replacement.Text = "ATA nº dezessete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 18/"
    .Replacement.Text = "ATA nº dezoito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 19/"
    .Replacement.Text = "ATA nº dezenove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 20/"
    .Replacement.Text = "ATA nº vinte/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 21/"
    .Replacement.Text = "ATA nº vinte e um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 22/"
    .Replacement.Text = "ATA nº vinte e dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 23/"
    .Replacement.Text = "ATA nº vinte e três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 24/"
    .Replacement.Text = "ATA nº vinte e quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 25/"
    .Replacement.Text = "ATA nº vinte e cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 26/"
    .Replacement.Text = "ATA nº vinte e seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 27/"
    .Replacement.Text = "ATA nº vinte e sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 28/"
    .Replacement.Text = "ATA nº vinte e oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 29/"
    .Replacement.Text = "ATA nº vinte e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 30/"
    .Replacement.Text = "ATA nº trinta/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 31/"
    .Replacement.Text = "ATA nº trinta e um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 32/"
    .Replacement.Text = "ATA nº trinta e dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 33/"
    .Replacement.Text = "ATA nº trinta e três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 34/"
    .Replacement.Text = "ATA nº trinta e quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 35/"
    .Replacement.Text = "ATA nº trinta e cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 36/"
    .Replacement.Text = "ATA nº trinta e seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 37/"
    .Replacement.Text = "ATA nº trinta e sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 38/"
    .Replacement.Text = "ATA nº trinta e oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 39/"
    .Replacement.Text = "ATA nº trinta e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 40/"
    .Replacement.Text = "ATA nº quarenta/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 41/"
    .Replacement.Text = "ATA nº quarenta e um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 42/"
    .Replacement.Text = "ATA nº quarenta e dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 43/"
    .Replacement.Text = "ATA nº quarenta e três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 44/"
    .Replacement.Text = "ATA nº quarenta e quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 45/"
    .Replacement.Text = "ATA nº quarenta e cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 46/"
    .Replacement.Text = "ATA nº quarenta e seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 47/"
    .Replacement.Text = "ATA nº quarenta e sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 48/"
    .Replacement.Text = "ATA nº quarenta e oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 49/"
    .Replacement.Text = "ATA nº quarenta e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 50/"
    .Replacement.Text = "ATA nº cinquenta/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 51/"
    .Replacement.Text = "ATA nº cinquenta e um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 52/"
    .Replacement.Text = "ATA nº cinquenta e dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 53/"
    .Replacement.Text = "ATA nº cinquenta e três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 54/"
    .Replacement.Text = "ATA nº cinquenta e quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 55/"
    .Replacement.Text = "ATA nº cinquenta e cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 56/"
    .Replacement.Text = "ATA nº cinquenta e seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 57/"
    .Replacement.Text = "ATA nº cinquenta e sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 58/"
    .Replacement.Text = "ATA nº cinquenta e oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 59/"
    .Replacement.Text = "ATA nº cinquenta e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 60/"
    .Replacement.Text = "ATA nº sessenta/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 61/"
    .Replacement.Text = "ATA nº sessenta e um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 62/"
    .Replacement.Text = "ATA nº sessenta e dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 63/"
    .Replacement.Text = "ATA nº sessenta e três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 64/"
    .Replacement.Text = "ATA nº sessenta e quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 65/"
    .Replacement.Text = "ATA nº sessenta e cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 66/"
    .Replacement.Text = "ATA nº sessenta e seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 67/"
    .Replacement.Text = "ATA nº sessenta e sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 68/"
    .Replacement.Text = "ATA nº sessenta e oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 69/"
    .Replacement.Text = "ATA nº sessenta e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 70/"
    .Replacement.Text = "ATA nº setenta/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 71/"
    .Replacement.Text = "ATA nº setenta e um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 72/"
    .Replacement.Text = "ATA nº setenta e dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 73/"
    .Replacement.Text = "ATA nº setenta e três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 74/"
    .Replacement.Text = "ATA nº setenta e quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 75/"
    .Replacement.Text = "ATA nº setenta e cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 76/"
    .Replacement.Text = "ATA nº setenta e seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 77/"
    .Replacement.Text = "ATA nº setenta e sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 78/"
    .Replacement.Text = "ATA nº setenta e oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 79/"
    .Replacement.Text = "ATA nº setenta e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 80/"
    .Replacement.Text = "ATA nº oitenta/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 81/"
    .Replacement.Text = "ATA nº oitenta e um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 82/"
    .Replacement.Text = "ATA nº oitenta e dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 83/"
    .Replacement.Text = "ATA nº oitenta e três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 84/"
    .Replacement.Text = "ATA nº oitenta e quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 85/"
    .Replacement.Text = "ATA nº oitenta e cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 86/"
    .Replacement.Text = "ATA nº oitenta e seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 87/"
    .Replacement.Text = "ATA nº oitenta e sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 88/"
    .Replacement.Text = "ATA nº oitenta e oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 89/"
    .Replacement.Text = "ATA nº oitenta e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 90/"
    .Replacement.Text = "ATA nº noventa/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 91/"
    .Replacement.Text = "ATA nº noventa e um/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 92/"
    .Replacement.Text = "ATA nº noventa e dois/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 93/"
    .Replacement.Text = "ATA nº noventa e três/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 94/"
    .Replacement.Text = "ATA nº noventa e quatro/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 95/"
    .Replacement.Text = "ATA nº noventa e cinco/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 96/"
    .Replacement.Text = "ATA nº noventa e seis/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 97/"
    .Replacement.Text = "ATA nº noventa e sete/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 98/"
    .Replacement.Text = "ATA nº noventa e oito/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº 99/"
    .Replacement.Text = "ATA nº noventa e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

MsgBox "Algarismos nos títulos das atas substituídos por números por extenso.", vbInformation, "Show de bola!"

End Sub
Sub ApagaTodosOsAlgarismosBocaDeUrna2020()
Attribute ApagaTodosOsAlgarismosBocaDeUrna2020.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ApagaTodosOsAlgarismosBocaDeUrna2020"
'
' ApagaTodosOsAlgarismosBocaDeUrna2020 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "0"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "1"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "2"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "3"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "4"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "5"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "6"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "7"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "8"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "9"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    MsgBox "Algarismos de 0 a 9 deletados do documento todo.", vbInformation, "Show de bola!"
    
End Sub
Sub LimparBocaDeUrna2020()
Attribute LimparBocaDeUrna2020.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.LimparQuebrasBocaDeUrna2020"
'
' LimparBocaDeUrna2020 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "    "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "CÂMARA MUNICIPAL DE VEREADORES DE SANTA MARIA – RS" & vbTab _
            & "CSS/"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "   "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "    "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    MsgBox "Deletados: cabeçalhos errados, quebras de linha e de seção, espaços duplos, triplos e quádruplos.", vbInformation, "Show de bola!"
    
End Sub
Sub QuebraTítuloAtas()
Attribute QuebraTítuloAtas.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.QuebraTítuloAtas"
'
' QuebraTítuloAtas Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "ATA nº um/"
    .Replacement.Text = "^pATA nº um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº dois/"
    .Replacement.Text = "^pATA nº dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº três/"
    .Replacement.Text = "^pATA nº três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quatro/"
    .Replacement.Text = "^pATA nº quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinco/"
    .Replacement.Text = "^pATA nº cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº seis/"
    .Replacement.Text = "^pATA nº seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sete/"
    .Replacement.Text = "^pATA nº sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oito/"
    .Replacement.Text = "^pATA nº oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº nove/"
    .Replacement.Text = "^pATA nº nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº dez/"
    .Replacement.Text = "^pATA nº dez/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº onze/"
    .Replacement.Text = "^pATA nº onze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº doze/"
    .Replacement.Text = "^pATA nº doze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº treze/"
    .Replacement.Text = "^pATA nº treze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quatorze/"
    .Replacement.Text = "^pATA nº quatorze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quinze/"
    .Replacement.Text = "^pATA nº quinze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº dezesseis/"
    .Replacement.Text = "^pATA nº dezesseis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº dezessete/"
    .Replacement.Text = "^pATA nº dezessete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº dezoito/"
    .Replacement.Text = "^pATA nº dezoito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº dezenove/"
    .Replacement.Text = "^pATA nº dezenove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte/"
    .Replacement.Text = "^pATA nº vinte/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e um/"
    .Replacement.Text = "^pATA nº vinte e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e dois/"
    .Replacement.Text = "^pATA nº vinte e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e três/"
    .Replacement.Text = "^pATA nº vinte e três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e quatro/"
    .Replacement.Text = "^pATA nº vinte e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e cinco/"
    .Replacement.Text = "^pATA nº vinte e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e seis/"
    .Replacement.Text = "^pATA nº vinte e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e sete/"
    .Replacement.Text = "^pATA nº vinte e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e oito/"
    .Replacement.Text = "^pATA nº vinte e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº vinte e nove/"
    .Replacement.Text = "^pATA nº vinte e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta/"
    .Replacement.Text = "^pATA nº trinta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e um/"
    .Replacement.Text = "^pATA nº trinta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e dois/"
    .Replacement.Text = "^pATA nº trinta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e três/"
    .Replacement.Text = "^pATA nº trinta e três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e quatro/"
    .Replacement.Text = "^pATA nº trinta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e cinco/"
    .Replacement.Text = "^pATA nº trinta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e seis/"
    .Replacement.Text = "^pATA nº trinta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e sete/"
    .Replacement.Text = "^pATA nº trinta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e oito/"
    .Replacement.Text = "^pATA nº trinta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº trinta e nove/"
    .Replacement.Text = "^pATA nº trinta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta/"
    .Replacement.Text = "^pATA nº quarenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e um/"
    .Replacement.Text = "^pATA nº quarenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e dois/"
    .Replacement.Text = "^pATA nº quarenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e três/"
    .Replacement.Text = "^pATA nº quarenta e três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e quatro/"
    .Replacement.Text = "^pATA nº quarenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e cinco/"
    .Replacement.Text = "^pATA nº quarenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e seis/"
    .Replacement.Text = "^pATA nº quarenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e sete/"
    .Replacement.Text = "^pATA nº quarenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e oito/"
    .Replacement.Text = "^pATA nº quarenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº quarenta e nove/"
    .Replacement.Text = "^pATA nº quarenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta/"
    .Replacement.Text = "^pATA nº cinquenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e um/"
    .Replacement.Text = "^pATA nº cinquenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e dois/"
    .Replacement.Text = "^pATA nº cinquenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e três/"
    .Replacement.Text = "^pATA nº cinquenta e três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e quatro/"
    .Replacement.Text = "^pATA nº cinquenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e cinco/"
    .Replacement.Text = "^pATA nº cinquenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e seis/"
    .Replacement.Text = "^pATA nº cinquenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e sete/"
    .Replacement.Text = "^pATA nº cinquenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e oito/"
    .Replacement.Text = "^pATA nº cinquenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº cinquenta e nove/"
    .Replacement.Text = "^pATA nº cinquenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta/"
    .Replacement.Text = "^pATA nº sessenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e um/"
    .Replacement.Text = "^pATA nº sessenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e dois/"
    .Replacement.Text = "^pATA nº sessenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e três/"
    .Replacement.Text = "^pATA nº sessenta e três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e quatro/"
    .Replacement.Text = "^pATA nº sessenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e cinco/"
    .Replacement.Text = "^pATA nº sessenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e seis/"
    .Replacement.Text = "^pATA nº sessenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e sete/"
    .Replacement.Text = "^pATA nº sessenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e oito/"
    .Replacement.Text = "^pATA nº sessenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº sessenta e nove/"
    .Replacement.Text = "^pATA nº sessenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta/"
    .Replacement.Text = "^pATA nº setenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e um/"
    .Replacement.Text = "^pATA nº setenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e dois/"
    .Replacement.Text = "^pATA nº setenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e três/"
    .Replacement.Text = "^pATA nº setenta e três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e quatro/"
    .Replacement.Text = "^pATA nº setenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e cinco/"
    .Replacement.Text = "^pATA nº setenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e seis/"
    .Replacement.Text = "^pATA nº setenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e sete/"
    .Replacement.Text = "^pATA nº setenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e oito/"
    .Replacement.Text = "^pATA nº setenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº setenta e nove/"
    .Replacement.Text = "^pATA nº setenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta/"
    .Replacement.Text = "^pATA nº oitenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e um/"
    .Replacement.Text = "^pATA nº oitenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e dois/"
    .Replacement.Text = "^pATA nº oitenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e três/"
    .Replacement.Text = "^pATA nº oitenta e três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e quatro/"
    .Replacement.Text = "^pATA nº oitenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e cinco/"
    .Replacement.Text = "^pATA nº oitenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e seis/"
    .Replacement.Text = "^pATA nº oitenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e sete/"
    .Replacement.Text = "^pATA nº oitenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e oito/"
    .Replacement.Text = "^pATA nº oitenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº oitenta e nove/"
    .Replacement.Text = "^pATA nº oitenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa/"
    .Replacement.Text = "^pATA nº noventa/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e um/"
    .Replacement.Text = "^pATA nº noventa e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e dois/"
    .Replacement.Text = "^pATA nº noventa e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e três/"
    .Replacement.Text = "^pATA nº noventa e três/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e quatro/"
    .Replacement.Text = "^pATA nº noventa e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e cinco/"
    .Replacement.Text = "^pATA nº noventa e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e seis/"
    .Replacement.Text = "^pATA nº noventa e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e sete/"
    .Replacement.Text = "^pATA nº noventa e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e oito/"
    .Replacement.Text = "^pATA nº noventa e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA nº noventa e nove/"
    .Replacement.Text = "^pATA nº noventa e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

        MsgBox "Quebra de linha antes e depois das atas inserida.", vbInformation, "Show de bola!"
End Sub
Sub QuebraTítuloSeções()
Attribute QuebraTítuloSeções.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.QuebraTítuloSeções"
'
' QuebraTítuloSeções Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "SESSÃO PLENÁRIA ORDINÁRIA"
        .Replacement.Text = "^pSESSÃO PLENÁRIA ORDINÁRIA^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "PERÍODO DAS COMUNICAÇÕES"
        .Replacement.Text = "^pPERÍODO DAS COMUNICAÇÕES^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "PERÍODO DAS COMUNICAÇÕES"
        .Replacement.Text = "^pPERÍODO DAS COMUNICAÇÕES^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "GRANDE EXPEDIENTE"
        .Replacement.Text = "^pGRANDE EXPEDIENTE^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ESPAÇO DE LIDERANÇA"
        .Replacement.Text = "^pESPAÇO DE LIDERANÇA^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ESPAÇO DE LIDERANÇA"
        .Replacement.Text = "^pESPAÇO DE LIDERANÇA^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ESPAÇO DE LIDERANÇA"
        .Replacement.Text = "^pESPAÇO DE LIDERANÇA^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    MsgBox "Quebra de linha antes e depois das seções inserida.", vbInformation, "Show de bola!"
    
End Sub
Sub AtribuiEstiloSeções()
Attribute AtribuiEstiloSeções.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.AtribuiEstiloSeções"
'
' AtribuiEstiloSeções Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 1")
    With Selection.Find
        .Text = "ATA nº"
        .Replacement.Text = "ATA nº"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 2")
    With Selection.Find
        .Text = "SESSÃO PLENÁRIA ORDINÁRIA"
        .Replacement.Text = "SESSÃO PLENÁRIA ORDINÁRIA"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 2")
    With Selection.Find
        .Text = "PERÍODO DAS COMUNICAÇÕES"
        .Replacement.Text = "PERÍODO DAS COMUNICAÇÕES"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "GRANDE EXPEDIENTE"
        .Replacement.Text = "GRANDE EXPEDIENTE"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ESPAÇO DE LIDERANÇA"
        .Replacement.Text = "ESPAÇO DE LIDERANÇA"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    MsgBox "Títulos primários e secundários atribuidos.", vbInformation, "Show de bola!"
    
End Sub
Sub QuebraLinhasVereadores()
Attribute QuebraLinhasVereadores.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.QuebraLinhasVereadores"
'
' QuebraLinhasVereadores Macro
'
'
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Adelar Vargas"
    .Replacement.Text = "^pVereador Adelar Vargas^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Admar Pozzobom"
    .Replacement.Text = "^pVereador Admar Pozzobom^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Alexandre Vargas"
    .Replacement.Text = "^pVereador Alexandre Vargas^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereadora Celita da Silva"
    .Replacement.Text = "^pVereadora Celita da Silva^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Daniel Diniz"
    .Replacement.Text = "^pVereador Daniel Diniz^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereadora Deili Silva"
    .Replacement.Text = "^pVereadora Deili Silva^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Francisco Harrisson"
    .Replacement.Text = "^pVereador Francisco Harrisson^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador João Kaus"
    .Replacement.Text = "^pVereador João Kaus^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador João Chaves"
    .Replacement.Text = "^pVereador João Chaves^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador João Ricardo Vargas"
    .Replacement.Text = "^pVereador João Ricardo Vargas^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Jorge Trindade Soares"
    .Replacement.Text = "^pVereador Jorge Trindade Soares^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Juliano Soares"
    .Replacement.Text = "^pVereador Juliano Soares^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Leopoldo Ochulaki"
    .Replacement.Text = "^pVereador Leopoldo Ochulaki^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Leopoldo Vanderlei Ochulaki"
    .Replacement.Text = "^pVereador Leopoldo Vanderlei Ochulaki^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereadora Luci Duartes"
    .Replacement.Text = "^pVereadora Luci Duartes^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Luciano Guerra"
    .Replacement.Text = "^pVereador Luciano Guerra^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Manoel Badke"
    .Replacement.Text = "^pVereador Manoel Badke^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereadora Maria Aparecida Brizola"
    .Replacement.Text = "^pVereadora Maria Aparecida Brizola^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Marion Mortari"
    .Replacement.Text = "^pVereador Marion Mortari^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereadora Marta Zanella"
    .Replacement.Text = "^pVereadora Marta Zanella^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Ovídio Mayer"
    .Replacement.Text = "^pVereador Ovídio Mayer^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Valdir Oliveira"
    .Replacement.Text = "^pVereador Valdir Oliveira^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Vanderlei Araújo"
    .Replacement.Text = "^pVereador Vanderlei Araújo^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador André Agne Domingues"
    .Replacement.Text = "^pVereador André Agne Domingues^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Font.Bold = True
With Selection.Find
    .Text = "Vereador Cezar Gehm"
    .Replacement.Text = "^pVereador Cezar Gehm^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

    MsgBox "Quebra de linha nas menções aos vereadores.", vbInformation, "Show de bola!"

End Sub
Sub AtribuiEstiloVereadores()
Attribute AtribuiEstiloVereadores.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.AtribuiEstiloVereadores"
'
' AtribuiEstiloVereadores Macro
'
'
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Adelar Vargas"
    .Replacement.Text = "Vereador Adelar Vargas"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Admar Pozzobom"
    .Replacement.Text = "Vereador Admar Pozzobom"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Alexandre Vargas"
    .Replacement.Text = "Vereador Alexandre Vargas"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereadora Celita da Silva"
    .Replacement.Text = "Vereadora Celita da Silva"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Daniel Diniz"
    .Replacement.Text = "Vereador Daniel Diniz"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereadora Deili Silva"
    .Replacement.Text = "Vereadora Deili Silva"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Francisco Harrisson"
    .Replacement.Text = "Vereador Francisco Harrisson"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador João Kaus"
    .Replacement.Text = "Vereador João Kaus"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador João Chaves"
    .Replacement.Text = "Vereador João Chaves"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador João Ricardo Vargas"
    .Replacement.Text = "Vereador João Ricardo Vargas"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Jorge Trindade Soares"
    .Replacement.Text = "Vereador Jorge Trindade Soares"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Juliano Soares"
    .Replacement.Text = "Vereador Juliano Soares"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Leopoldo Ochulaki"
    .Replacement.Text = "Vereador Leopoldo Ochulaki"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Leopoldo Vanderlei Ochulaki"
    .Replacement.Text = "Vereador Leopoldo Vanderlei Ochulaki"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereadora Luci Duartes"
    .Replacement.Text = "Vereadora Luci Duartes"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Luciano Guerra"
    .Replacement.Text = "Vereador Luciano Guerra"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Manoel Badke"
    .Replacement.Text = "Vereador Manoel Badke"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereadora Maria Aparecida Brizola"
    .Replacement.Text = "Vereadora Maria Aparecida Brizola"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Marion Mortari"
    .Replacement.Text = "Vereador Marion Mortari"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereadora Marta Zanella"
    .Replacement.Text = "Vereadora Marta Zanella"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Ovídio Mayer"
    .Replacement.Text = "Vereador Ovídio Mayer"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Valdir Oliveira"
    .Replacement.Text = "Vereador Valdir Oliveira"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Vanderlei Araújo"
    .Replacement.Text = "Vereador Vanderlei Araújo"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador André Agne Domingues"
    .Replacement.Text = "Vereador André Agne Domingues"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Replacement.ClearFormatting
Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
With Selection.Find
    .Text = "Vereador Cezar Gehm"
    .Replacement.Text = "Vereador Cezar Gehm"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

    MsgBox "Títulos terciários atribuídos.", vbInformation, "Show de bola!"

End Sub
Sub DeletarCabecalhosERodapes()
'
' DeletarCabecalhosERodapes Macro
'
'
  Dim sec As Section
  Dim hd_ft As HeaderFooter

  For Each sec In ActiveDocument.Sections
    For Each hd_ft In sec.Headers
      hd_ft.Range.Delete
    Next
    For Each hd_ft In sec.Footers
      hd_ft.Range.Delete
    Next
  Next sec
  
End Sub
Sub B_URNA_2A_LimparBocaDeUrna2020()
'
' TratamentoDeAtas Macro
'
'


' LimparBocaDeUrna2020 Macro
' Sem tirar quebra de linha


    Selection.Find.ClearFormatting
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "CÂMARA MUNICIPAL DE VEREADORES DE SANTA MARIA – RS" & vbTab _
            & "CSS/"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "   "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "    "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll



    MsgBox "A princício, tudo certo.", vbInformation, "Show de bola!"


End Sub
Sub BocaDeUrna1()
'
' BocaDeUrna1 Macro
'
'
' LimparBocaDeUrna2020 Macro
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "    "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "CÂMARA MUNICIPAL DE VEREADORES DE SANTA MARIA – RS" & vbTab _
            & "CSS/"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "   "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "    "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' AlgarismosParaNumExtensoAtasBocaDeUrna2020 Macro
    ' Altera algarismos de 01 a 99 do título das atas para números por extenso para tratamento de atas no projeto Boca de Urna 2020.
    '
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ATA nº 01/"
        .Replacement.Text = "ATA nº um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 02/"
        .Replacement.Text = "ATA nº dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 03/"
        .Replacement.Text = "ATA nº três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 04/"
        .Replacement.Text = "ATA nº quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 05/"
        .Replacement.Text = "ATA nº cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 06/"
        .Replacement.Text = "ATA nº seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 07/"
        .Replacement.Text = "ATA nº sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 08/"
        .Replacement.Text = "ATA nº oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 09/"
        .Replacement.Text = "ATA nº nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 10/"
        .Replacement.Text = "ATA nº dez/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 11/"
        .Replacement.Text = "ATA nº onze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 12/"
        .Replacement.Text = "ATA nº doze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 13/"
        .Replacement.Text = "ATA nº treze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 14/"
        .Replacement.Text = "ATA nº quatorze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 15/"
        .Replacement.Text = "ATA nº quinze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 16/"
        .Replacement.Text = "ATA nº dezesseis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 17/"
        .Replacement.Text = "ATA nº dezessete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 18/"
        .Replacement.Text = "ATA nº dezoito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 19/"
        .Replacement.Text = "ATA nº dezenove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 20/"
        .Replacement.Text = "ATA nº vinte/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 21/"
        .Replacement.Text = "ATA nº vinte e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 22/"
        .Replacement.Text = "ATA nº vinte e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 23/"
        .Replacement.Text = "ATA nº vinte e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 24/"
        .Replacement.Text = "ATA nº vinte e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 25/"
        .Replacement.Text = "ATA nº vinte e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 26/"
        .Replacement.Text = "ATA nº vinte e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 27/"
        .Replacement.Text = "ATA nº vinte e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 28/"
        .Replacement.Text = "ATA nº vinte e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 29/"
        .Replacement.Text = "ATA nº vinte e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 30/"
        .Replacement.Text = "ATA nº trinta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 31/"
        .Replacement.Text = "ATA nº trinta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 32/"
        .Replacement.Text = "ATA nº trinta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 33/"
        .Replacement.Text = "ATA nº trinta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 34/"
        .Replacement.Text = "ATA nº trinta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 35/"
        .Replacement.Text = "ATA nº trinta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 36/"
        .Replacement.Text = "ATA nº trinta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 37/"
        .Replacement.Text = "ATA nº trinta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 38/"
        .Replacement.Text = "ATA nº trinta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 39/"
        .Replacement.Text = "ATA nº trinta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 40/"
        .Replacement.Text = "ATA nº quarenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 41/"
        .Replacement.Text = "ATA nº quarenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 42/"
        .Replacement.Text = "ATA nº quarenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 43/"
        .Replacement.Text = "ATA nº quarenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 44/"
        .Replacement.Text = "ATA nº quarenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 45/"
        .Replacement.Text = "ATA nº quarenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 46/"
        .Replacement.Text = "ATA nº quarenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 47/"
        .Replacement.Text = "ATA nº quarenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 48/"
        .Replacement.Text = "ATA nº quarenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 49/"
        .Replacement.Text = "ATA nº quarenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 50/"
        .Replacement.Text = "ATA nº cinquenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 51/"
        .Replacement.Text = "ATA nº cinquenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 52/"
        .Replacement.Text = "ATA nº cinquenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 53/"
        .Replacement.Text = "ATA nº cinquenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 54/"
        .Replacement.Text = "ATA nº cinquenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 55/"
        .Replacement.Text = "ATA nº cinquenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 56/"
        .Replacement.Text = "ATA nº cinquenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 57/"
        .Replacement.Text = "ATA nº cinquenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 58/"
        .Replacement.Text = "ATA nº cinquenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 59/"
        .Replacement.Text = "ATA nº cinquenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 60/"
        .Replacement.Text = "ATA nº sessenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 61/"
        .Replacement.Text = "ATA nº sessenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 62/"
        .Replacement.Text = "ATA nº sessenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 63/"
        .Replacement.Text = "ATA nº sessenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 64/"
        .Replacement.Text = "ATA nº sessenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 65/"
        .Replacement.Text = "ATA nº sessenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 66/"
        .Replacement.Text = "ATA nº sessenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 67/"
        .Replacement.Text = "ATA nº sessenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 68/"
        .Replacement.Text = "ATA nº sessenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 69/"
        .Replacement.Text = "ATA nº sessenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 70/"
        .Replacement.Text = "ATA nº setenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 71/"
        .Replacement.Text = "ATA nº setenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 72/"
        .Replacement.Text = "ATA nº setenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 73/"
        .Replacement.Text = "ATA nº setenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 74/"
        .Replacement.Text = "ATA nº setenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 75/"
        .Replacement.Text = "ATA nº setenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 76/"
        .Replacement.Text = "ATA nº setenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 77/"
        .Replacement.Text = "ATA nº setenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 78/"
        .Replacement.Text = "ATA nº setenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 79/"
        .Replacement.Text = "ATA nº setenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 80/"
        .Replacement.Text = "ATA nº oitenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 81/"
        .Replacement.Text = "ATA nº oitenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 82/"
        .Replacement.Text = "ATA nº oitenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 83/"
        .Replacement.Text = "ATA nº oitenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 84/"
        .Replacement.Text = "ATA nº oitenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 85/"
        .Replacement.Text = "ATA nº oitenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 86/"
        .Replacement.Text = "ATA nº oitenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 87/"
        .Replacement.Text = "ATA nº oitenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 88/"
        .Replacement.Text = "ATA nº oitenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 89/"
        .Replacement.Text = "ATA nº oitenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 90/"
        .Replacement.Text = "ATA nº noventa/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 91/"
        .Replacement.Text = "ATA nº noventa e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 92/"
        .Replacement.Text = "ATA nº noventa e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 93/"
        .Replacement.Text = "ATA nº noventa e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 94/"
        .Replacement.Text = "ATA nº noventa e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 95/"
        .Replacement.Text = "ATA nº noventa e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 96/"
        .Replacement.Text = "ATA nº noventa e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 97/"
        .Replacement.Text = "ATA nº noventa e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 98/"
        .Replacement.Text = "ATA nº noventa e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 99/"
        .Replacement.Text = "ATA nº noventa e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll



    MsgBox "Parte 1 ok", vbInformation, "Show de bola!"
    
End Sub
Sub BocaDeUrna2()
'
' BocaDeUrna2 Macro
'
'




    ' ApagaTodosOsAlgarismosBocaDeUrna2020 Macro
    '
    '
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "0"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "1"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "2"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "3"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "4"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "5"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "6"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "7"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "8"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "9"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll

        ' QuebraTítuloAtas Macro
        '
        '
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "ATA nº um/"
            .Replacement.Text = "^pATA nº um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dois/"
            .Replacement.Text = "^pATA nº dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº três/"
            .Replacement.Text = "^pATA nº três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quatro/"
            .Replacement.Text = "^pATA nº quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinco/"
            .Replacement.Text = "^pATA nº cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº seis/"
            .Replacement.Text = "^pATA nº seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sete/"
            .Replacement.Text = "^pATA nº sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oito/"
            .Replacement.Text = "^pATA nº oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº nove/"
            .Replacement.Text = "^pATA nº nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dez/"
            .Replacement.Text = "^pATA nº dez/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº onze/"
            .Replacement.Text = "^pATA nº onze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº doze/"
            .Replacement.Text = "^pATA nº doze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº treze/"
            .Replacement.Text = "^pATA nº treze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quatorze/"
            .Replacement.Text = "^pATA nº quatorze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quinze/"
            .Replacement.Text = "^pATA nº quinze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dezesseis/"
            .Replacement.Text = "^pATA nº dezesseis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dezessete/"
            .Replacement.Text = "^pATA nº dezessete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dezoito/"
            .Replacement.Text = "^pATA nº dezoito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dezenove/"
            .Replacement.Text = "^pATA nº dezenove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte/"
            .Replacement.Text = "^pATA nº vinte/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e um/"
            .Replacement.Text = "^pATA nº vinte e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e dois/"
            .Replacement.Text = "^pATA nº vinte e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e três/"
            .Replacement.Text = "^pATA nº vinte e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e quatro/"
            .Replacement.Text = "^pATA nº vinte e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e cinco/"
            .Replacement.Text = "^pATA nº vinte e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e seis/"
            .Replacement.Text = "^pATA nº vinte e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e sete/"
            .Replacement.Text = "^pATA nº vinte e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e oito/"
            .Replacement.Text = "^pATA nº vinte e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e nove/"
            .Replacement.Text = "^pATA nº vinte e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta/"
            .Replacement.Text = "^pATA nº trinta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e um/"
            .Replacement.Text = "^pATA nº trinta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e dois/"
            .Replacement.Text = "^pATA nº trinta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e três/"
            .Replacement.Text = "^pATA nº trinta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e quatro/"
            .Replacement.Text = "^pATA nº trinta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e cinco/"
            .Replacement.Text = "^pATA nº trinta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e seis/"
            .Replacement.Text = "^pATA nº trinta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e sete/"
            .Replacement.Text = "^pATA nº trinta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e oito/"
            .Replacement.Text = "^pATA nº trinta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e nove/"
            .Replacement.Text = "^pATA nº trinta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta/"
            .Replacement.Text = "^pATA nº quarenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e um/"
            .Replacement.Text = "^pATA nº quarenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e dois/"
            .Replacement.Text = "^pATA nº quarenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e três/"
            .Replacement.Text = "^pATA nº quarenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e quatro/"
            .Replacement.Text = "^pATA nº quarenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e cinco/"
            .Replacement.Text = "^pATA nº quarenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e seis/"
            .Replacement.Text = "^pATA nº quarenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e sete/"
            .Replacement.Text = "^pATA nº quarenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e oito/"
            .Replacement.Text = "^pATA nº quarenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e nove/"
            .Replacement.Text = "^pATA nº quarenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta/"
            .Replacement.Text = "^pATA nº cinquenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e um/"
            .Replacement.Text = "^pATA nº cinquenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e dois/"
            .Replacement.Text = "^pATA nº cinquenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e três/"
            .Replacement.Text = "^pATA nº cinquenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e quatro/"
            .Replacement.Text = "^pATA nº cinquenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e cinco/"
            .Replacement.Text = "^pATA nº cinquenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e seis/"
            .Replacement.Text = "^pATA nº cinquenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e sete/"
            .Replacement.Text = "^pATA nº cinquenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e oito/"
            .Replacement.Text = "^pATA nº cinquenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e nove/"
            .Replacement.Text = "^pATA nº cinquenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta/"
            .Replacement.Text = "^pATA nº sessenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e um/"
            .Replacement.Text = "^pATA nº sessenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e dois/"
            .Replacement.Text = "^pATA nº sessenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e três/"
            .Replacement.Text = "^pATA nº sessenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e quatro/"
            .Replacement.Text = "^pATA nº sessenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e cinco/"
            .Replacement.Text = "^pATA nº sessenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e seis/"
            .Replacement.Text = "^pATA nº sessenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e sete/"
            .Replacement.Text = "^pATA nº sessenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e oito/"
            .Replacement.Text = "^pATA nº sessenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e nove/"
            .Replacement.Text = "^pATA nº sessenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta/"
            .Replacement.Text = "^pATA nº setenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e um/"
            .Replacement.Text = "^pATA nº setenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e dois/"
            .Replacement.Text = "^pATA nº setenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e três/"
            .Replacement.Text = "^pATA nº setenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e quatro/"
            .Replacement.Text = "^pATA nº setenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e cinco/"
            .Replacement.Text = "^pATA nº setenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e seis/"
            .Replacement.Text = "^pATA nº setenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e sete/"
            .Replacement.Text = "^pATA nº setenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e oito/"
            .Replacement.Text = "^pATA nº setenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e nove/"
            .Replacement.Text = "^pATA nº setenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta/"
            .Replacement.Text = "^pATA nº oitenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e um/"
            .Replacement.Text = "^pATA nº oitenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e dois/"
            .Replacement.Text = "^pATA nº oitenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e três/"
            .Replacement.Text = "^pATA nº oitenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e quatro/"
            .Replacement.Text = "^pATA nº oitenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e cinco/"
            .Replacement.Text = "^pATA nº oitenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e seis/"
            .Replacement.Text = "^pATA nº oitenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e sete/"
            .Replacement.Text = "^pATA nº oitenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e oito/"
            .Replacement.Text = "^pATA nº oitenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e nove/"
            .Replacement.Text = "^pATA nº oitenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa/"
            .Replacement.Text = "^pATA nº noventa/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e um/"
            .Replacement.Text = "^pATA nº noventa e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e dois/"
            .Replacement.Text = "^pATA nº noventa e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e três/"
            .Replacement.Text = "^pATA nº noventa e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e quatro/"
            .Replacement.Text = "^pATA nº noventa e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e cinco/"
            .Replacement.Text = "^pATA nº noventa e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e seis/"
            .Replacement.Text = "^pATA nº noventa e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e sete/"
            .Replacement.Text = "^pATA nº noventa e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e oito/"
            .Replacement.Text = "^pATA nº noventa e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e nove/"
            .Replacement.Text = "^pATA nº noventa e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll


    MsgBox "Parte 2 ok", vbInformation, "Show de bola!"
    
End Sub
Sub BocaDeUrna3()
'
' BocaDeUrna3 Macro
'
'



        ' QuebraTítuloSeções Macro
        '
        '
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "SESSÃO PLENÁRIA ORDINÁRIA"
                .Replacement.Text = "^pSESSÃO PLENÁRIA ORDINÁRIA^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "PERÍODO DAS COMUNICAÇÕES"
                .Replacement.Text = "^pPERÍODO DAS COMUNICAÇÕES^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "PERÍODO DAS COMUNICAÇÕES"
                .Replacement.Text = "^pPERÍODO DAS COMUNICAÇÕES^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "GRANDE EXPEDIENTE"
                .Replacement.Text = "^pGRANDE EXPEDIENTE^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "ESPAÇO DE LIDERANÇA"
                .Replacement.Text = "^pESPAÇO DE LIDERANÇA^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "ESPAÇO DE LIDERANÇA"
                .Replacement.Text = "^pESPAÇO DE LIDERANÇA^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "ESPAÇO DE LIDERANÇA"
                .Replacement.Text = "^pESPAÇO DE LIDERANÇA^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With

            ' AtribuiEstiloSeções Macro
            '
            '
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 1")
                With Selection.Find
                    .Text = "ATA nº"
                    .Replacement.Text = "ATA nº"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 2")
                With Selection.Find
                    .Text = "SESSÃO PLENÁRIA ORDINÁRIA"
                    .Replacement.Text = "SESSÃO PLENÁRIA ORDINÁRIA"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 2")
                With Selection.Find
                    .Text = "PERÍODO DAS COMUNICAÇÕES"
                    .Replacement.Text = "PERÍODO DAS COMUNICAÇÕES"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                With Selection.Find
                    .Text = "GRANDE EXPEDIENTE"
                    .Replacement.Text = "GRANDE EXPEDIENTE"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                With Selection.Find
                    .Text = "ESPAÇO DE LIDERANÇA"
                    .Replacement.Text = "ESPAÇO DE LIDERANÇA"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll


                ' QuebraLinhasVereadores Macro
                '
                '
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Adelar Vargas"
                    .Replacement.Text = "^pVereador Adelar Vargas^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Admar Pozzobom"
                    .Replacement.Text = "^pVereador Admar Pozzobom^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Alexandre Vargas"
                    .Replacement.Text = "^pVereador Alexandre Vargas^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Celita da Silva"
                    .Replacement.Text = "^pVereadora Celita da Silva^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Daniel Diniz"
                    .Replacement.Text = "^pVereador Daniel Diniz^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Deili Silva"
                    .Replacement.Text = "^pVereadora Deili Silva^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Francisco Harrisson"
                    .Replacement.Text = "^pVereador Francisco Harrisson^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador João Kaus"
                    .Replacement.Text = "^pVereador João Kaus^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador João Chaves"
                    .Replacement.Text = "^pVereador João Chaves^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador João Ricardo Vargas"
                    .Replacement.Text = "^pVereador João Ricardo Vargas^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Jorge Trindade Soares"
                    .Replacement.Text = "^pVereador Jorge Trindade Soares^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Juliano Soares"
                    .Replacement.Text = "^pVereador Juliano Soares^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Leopoldo Ochulaki"
                    .Replacement.Text = "^pVereador Leopoldo Ochulaki^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Leopoldo Vanderlei Ochulaki"
                    .Replacement.Text = "^pVereador Leopoldo Vanderlei Ochulaki^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Luci Duartes"
                    .Replacement.Text = "^pVereadora Luci Duartes^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Luciano Guerra"
                    .Replacement.Text = "^pVereador Luciano Guerra^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Manoel Badke"
                    .Replacement.Text = "^pVereador Manoel Badke^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Maria Aparecida Brizola"
                    .Replacement.Text = "^pVereadora Maria Aparecida Brizola^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Marion Mortari"
                    .Replacement.Text = "^pVereador Marion Mortari^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Marta Zanella"
                    .Replacement.Text = "^pVereadora Marta Zanella^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Ovídio Mayer"
                    .Replacement.Text = "^pVereador Ovídio Mayer^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Valdir Oliveira"
                    .Replacement.Text = "^pVereador Valdir Oliveira^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Vanderlei Araújo"
                    .Replacement.Text = "^pVereador Vanderlei Araújo^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador André Agne Domingues"
                    .Replacement.Text = "^pVereador André Agne Domingues^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Cezar Gehm"
                    .Replacement.Text = "^pVereador Cezar Gehm^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll


                ' AtribuiEstiloVereadores Macro
                '
                '
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Adelar Vargas"
                    .Replacement.Text = "Vereador Adelar Vargas"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Admar Pozzobom"
                    .Replacement.Text = "Vereador Admar Pozzobom"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Alexandre Vargas"
                    .Replacement.Text = "Vereador Alexandre Vargas"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Celita da Silva"
                    .Replacement.Text = "Vereadora Celita da Silva"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Daniel Diniz"
                    .Replacement.Text = "Vereador Daniel Diniz"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Deili Silva"
                    .Replacement.Text = "Vereadora Deili Silva"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Francisco Harrisson"
                    .Replacement.Text = "Vereador Francisco Harrisson"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador João Kaus"
                    .Replacement.Text = "Vereador João Kaus"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador João Chaves"
                    .Replacement.Text = "Vereador João Chaves"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador João Ricardo Vargas"
                    .Replacement.Text = "Vereador João Ricardo Vargas"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Jorge Trindade Soares"
                    .Replacement.Text = "Vereador Jorge Trindade Soares"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Juliano Soares"
                    .Replacement.Text = "Vereador Juliano Soares"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Leopoldo Ochulaki"
                    .Replacement.Text = "Vereador Leopoldo Ochulaki"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Leopoldo Vanderlei Ochulaki"
                    .Replacement.Text = "Vereador Leopoldo Vanderlei Ochulaki"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Luci Duartes"
                    .Replacement.Text = "Vereadora Luci Duartes"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Luciano Guerra"
                    .Replacement.Text = "Vereador Luciano Guerra"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Manoel Badke"
                    .Replacement.Text = "Vereador Manoel Badke"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Maria Aparecida Brizola"
                    .Replacement.Text = "Vereadora Maria Aparecida Brizola"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Marion Mortari"
                    .Replacement.Text = "Vereador Marion Mortari"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Marta Zanella"
                    .Replacement.Text = "Vereadora Marta Zanella"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Ovídio Mayer"
                    .Replacement.Text = "Vereador Ovídio Mayer"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Valdir Oliveira"
                    .Replacement.Text = "Vereador Valdir Oliveira"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Vanderlei Araújo"
                    .Replacement.Text = "Vereador Vanderlei Araújo"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador André Agne Domingues"
                    .Replacement.Text = "Vereador André Agne Domingues"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Cezar Gehm"
                    .Replacement.Text = "Vereador Cezar Gehm"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll

    MsgBox "Parte 3 ok", vbInformation, "Show de bola!"

End Sub
Sub B_URNA_1()
Attribute B_URNA_1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.B_URNA_1"
'
' B_URNA_1 Macro
'
'
'
' LimparBocaDeUrna2020 Macro
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "    "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "CÂMARA MUNICIPAL DE VEREADORES DE SANTA MARIA – RS" & vbTab _
            & "CSS/"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "   "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "    "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' AlgarismosParaNumExtensoAtasBocaDeUrna2020 Macro
    ' Altera algarismos de 01 a 99 do título das atas para números por extenso para tratamento de atas no projeto Boca de Urna 2020.
    ' Adaptado para três digitos no número da ata, como as atas de 2017
    '
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ATA nº 001 /"
        .Replacement.Text = "ATA nº um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 002 /"
        .Replacement.Text = "ATA nº dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 003 /"
        .Replacement.Text = "ATA nº três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 004 /"
        .Replacement.Text = "ATA nº quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 005 /"
        .Replacement.Text = "ATA nº cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 006 /"
        .Replacement.Text = "ATA nº seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 007 /"
        .Replacement.Text = "ATA nº sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 008 /"
        .Replacement.Text = "ATA nº oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 009 /"
        .Replacement.Text = "ATA nº nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 010 /"
        .Replacement.Text = "ATA nº dez/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 011 /"
        .Replacement.Text = "ATA nº onze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 012 /"
        .Replacement.Text = "ATA nº doze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 013 /"
        .Replacement.Text = "ATA nº treze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 014 /"
        .Replacement.Text = "ATA nº quatorze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 015 /"
        .Replacement.Text = "ATA nº quinze/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 016 /"
        .Replacement.Text = "ATA nº dezesseis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 017 /"
        .Replacement.Text = "ATA nº dezessete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 018 /"
        .Replacement.Text = "ATA nº dezoito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 019 /"
        .Replacement.Text = "ATA nº dezenove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 020 /"
        .Replacement.Text = "ATA nº vinte/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 021 /"
        .Replacement.Text = "ATA nº vinte e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 022 /"
        .Replacement.Text = "ATA nº vinte e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 023 /"
        .Replacement.Text = "ATA nº vinte e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 024 /"
        .Replacement.Text = "ATA nº vinte e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 025 /"
        .Replacement.Text = "ATA nº vinte e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 026 /"
        .Replacement.Text = "ATA nº vinte e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 027 /"
        .Replacement.Text = "ATA nº vinte e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 028 /"
        .Replacement.Text = "ATA nº vinte e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 029 /"
        .Replacement.Text = "ATA nº vinte e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 030 /"
        .Replacement.Text = "ATA nº trinta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 031 /"
        .Replacement.Text = "ATA nº trinta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 032 /"
        .Replacement.Text = "ATA nº trinta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 033 /"
        .Replacement.Text = "ATA nº trinta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 034 /"
        .Replacement.Text = "ATA nº trinta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 035 /"
        .Replacement.Text = "ATA nº trinta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 036 /"
        .Replacement.Text = "ATA nº trinta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 037 /"
        .Replacement.Text = "ATA nº trinta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 038 /"
        .Replacement.Text = "ATA nº trinta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 039 /"
        .Replacement.Text = "ATA nº trinta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 040 /"
        .Replacement.Text = "ATA nº quarenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 041 /"
        .Replacement.Text = "ATA nº quarenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 042 /"
        .Replacement.Text = "ATA nº quarenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 043 /"
        .Replacement.Text = "ATA nº quarenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 044 /"
        .Replacement.Text = "ATA nº quarenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 045 /"
        .Replacement.Text = "ATA nº quarenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 046 /"
        .Replacement.Text = "ATA nº quarenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 047 /"
        .Replacement.Text = "ATA nº quarenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 048 /"
        .Replacement.Text = "ATA nº quarenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 049 /"
        .Replacement.Text = "ATA nº quarenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 050 /"
        .Replacement.Text = "ATA nº cinquenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 051 /"
        .Replacement.Text = "ATA nº cinquenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 052 /"
        .Replacement.Text = "ATA nº cinquenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 053 /"
        .Replacement.Text = "ATA nº cinquenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 054 /"
        .Replacement.Text = "ATA nº cinquenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 055 /"
        .Replacement.Text = "ATA nº cinquenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 056 /"
        .Replacement.Text = "ATA nº cinquenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 057 /"
        .Replacement.Text = "ATA nº cinquenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 058 /"
        .Replacement.Text = "ATA nº cinquenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 059 /"
        .Replacement.Text = "ATA nº cinquenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 060 /"
        .Replacement.Text = "ATA nº sessenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 061 /"
        .Replacement.Text = "ATA nº sessenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 062 /"
        .Replacement.Text = "ATA nº sessenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 063 /"
        .Replacement.Text = "ATA nº sessenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 064 /"
        .Replacement.Text = "ATA nº sessenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 065 /"
        .Replacement.Text = "ATA nº sessenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 066 /"
        .Replacement.Text = "ATA nº sessenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 067 /"
        .Replacement.Text = "ATA nº sessenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 068 /"
        .Replacement.Text = "ATA nº sessenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 069 /"
        .Replacement.Text = "ATA nº sessenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 070 /"
        .Replacement.Text = "ATA nº setenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 071 /"
        .Replacement.Text = "ATA nº setenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 072 /"
        .Replacement.Text = "ATA nº setenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 073 /"
        .Replacement.Text = "ATA nº setenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 074 /"
        .Replacement.Text = "ATA nº setenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 075 /"
        .Replacement.Text = "ATA nº setenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 076 /"
        .Replacement.Text = "ATA nº setenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 077 /"
        .Replacement.Text = "ATA nº setenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 078 /"
        .Replacement.Text = "ATA nº setenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 079 /"
        .Replacement.Text = "ATA nº setenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 080 /"
        .Replacement.Text = "ATA nº oitenta/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 081 /"
        .Replacement.Text = "ATA nº oitenta e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 082 /"
        .Replacement.Text = "ATA nº oitenta e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 083 /"
        .Replacement.Text = "ATA nº oitenta e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 084 /"
        .Replacement.Text = "ATA nº oitenta e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 085 /"
        .Replacement.Text = "ATA nº oitenta e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 086 /"
        .Replacement.Text = "ATA nº oitenta e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 087 /"
        .Replacement.Text = "ATA nº oitenta e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 088 /"
        .Replacement.Text = "ATA nº oitenta e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 089 /"
        .Replacement.Text = "ATA nº oitenta e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 090 /"
        .Replacement.Text = "ATA nº noventa/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 091 /"
        .Replacement.Text = "ATA nº noventa e um/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 092 /"
        .Replacement.Text = "ATA nº noventa e dois/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 093 /"
        .Replacement.Text = "ATA nº noventa e três/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 094 /"
        .Replacement.Text = "ATA nº noventa e quatro/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 095 /"
        .Replacement.Text = "ATA nº noventa e cinco/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 096 /"
        .Replacement.Text = "ATA nº noventa e seis/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 097 /"
        .Replacement.Text = "ATA nº noventa e sete/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 098 /"
        .Replacement.Text = "ATA nº noventa e oito/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "ATA nº 099 /"
        .Replacement.Text = "ATA nº noventa e nove/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll



    MsgBox "Parte 1 ok", vbInformation, "Show de bola!"
    
End Sub
Sub B_URNA_2()
Attribute B_URNA_2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.B_URNA_2"
'
' B_URNA_2 Macro
'
'





    ' ApagaTodosOsAlgarismosBocaDeUrna2020 Macro
    '
    '
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "0"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "1"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "2"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "3"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "4"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "5"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "6"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "7"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "8"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
            .Text = "9"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll

        ' QuebraTítuloAtas Macro
        '
        '
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "ATA nº um/"
            .Replacement.Text = "^pATA nº um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dois/"
            .Replacement.Text = "^pATA nº dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº três/"
            .Replacement.Text = "^pATA nº três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quatro/"
            .Replacement.Text = "^pATA nº quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinco/"
            .Replacement.Text = "^pATA nº cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº seis/"
            .Replacement.Text = "^pATA nº seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sete/"
            .Replacement.Text = "^pATA nº sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oito/"
            .Replacement.Text = "^pATA nº oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº nove/"
            .Replacement.Text = "^pATA nº nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dez/"
            .Replacement.Text = "^pATA nº dez/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº onze/"
            .Replacement.Text = "^pATA nº onze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº doze/"
            .Replacement.Text = "^pATA nº doze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº treze/"
            .Replacement.Text = "^pATA nº treze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quatorze/"
            .Replacement.Text = "^pATA nº quatorze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quinze/"
            .Replacement.Text = "^pATA nº quinze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dezesseis/"
            .Replacement.Text = "^pATA nº dezesseis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dezessete/"
            .Replacement.Text = "^pATA nº dezessete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dezoito/"
            .Replacement.Text = "^pATA nº dezoito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº dezenove/"
            .Replacement.Text = "^pATA nº dezenove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte/"
            .Replacement.Text = "^pATA nº vinte/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e um/"
            .Replacement.Text = "^pATA nº vinte e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e dois/"
            .Replacement.Text = "^pATA nº vinte e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e três/"
            .Replacement.Text = "^pATA nº vinte e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e quatro/"
            .Replacement.Text = "^pATA nº vinte e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e cinco/"
            .Replacement.Text = "^pATA nº vinte e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e seis/"
            .Replacement.Text = "^pATA nº vinte e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e sete/"
            .Replacement.Text = "^pATA nº vinte e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e oito/"
            .Replacement.Text = "^pATA nº vinte e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº vinte e nove/"
            .Replacement.Text = "^pATA nº vinte e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta/"
            .Replacement.Text = "^pATA nº trinta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e um/"
            .Replacement.Text = "^pATA nº trinta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e dois/"
            .Replacement.Text = "^pATA nº trinta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e três/"
            .Replacement.Text = "^pATA nº trinta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e quatro/"
            .Replacement.Text = "^pATA nº trinta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e cinco/"
            .Replacement.Text = "^pATA nº trinta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e seis/"
            .Replacement.Text = "^pATA nº trinta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e sete/"
            .Replacement.Text = "^pATA nº trinta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e oito/"
            .Replacement.Text = "^pATA nº trinta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº trinta e nove/"
            .Replacement.Text = "^pATA nº trinta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta/"
            .Replacement.Text = "^pATA nº quarenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e um/"
            .Replacement.Text = "^pATA nº quarenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e dois/"
            .Replacement.Text = "^pATA nº quarenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e três/"
            .Replacement.Text = "^pATA nº quarenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e quatro/"
            .Replacement.Text = "^pATA nº quarenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e cinco/"
            .Replacement.Text = "^pATA nº quarenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e seis/"
            .Replacement.Text = "^pATA nº quarenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e sete/"
            .Replacement.Text = "^pATA nº quarenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e oito/"
            .Replacement.Text = "^pATA nº quarenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº quarenta e nove/"
            .Replacement.Text = "^pATA nº quarenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta/"
            .Replacement.Text = "^pATA nº cinquenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e um/"
            .Replacement.Text = "^pATA nº cinquenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e dois/"
            .Replacement.Text = "^pATA nº cinquenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e três/"
            .Replacement.Text = "^pATA nº cinquenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e quatro/"
            .Replacement.Text = "^pATA nº cinquenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e cinco/"
            .Replacement.Text = "^pATA nº cinquenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e seis/"
            .Replacement.Text = "^pATA nº cinquenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e sete/"
            .Replacement.Text = "^pATA nº cinquenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e oito/"
            .Replacement.Text = "^pATA nº cinquenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº cinquenta e nove/"
            .Replacement.Text = "^pATA nº cinquenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta/"
            .Replacement.Text = "^pATA nº sessenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e um/"
            .Replacement.Text = "^pATA nº sessenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e dois/"
            .Replacement.Text = "^pATA nº sessenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e três/"
            .Replacement.Text = "^pATA nº sessenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e quatro/"
            .Replacement.Text = "^pATA nº sessenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e cinco/"
            .Replacement.Text = "^pATA nº sessenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e seis/"
            .Replacement.Text = "^pATA nº sessenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e sete/"
            .Replacement.Text = "^pATA nº sessenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e oito/"
            .Replacement.Text = "^pATA nº sessenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº sessenta e nove/"
            .Replacement.Text = "^pATA nº sessenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta/"
            .Replacement.Text = "^pATA nº setenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e um/"
            .Replacement.Text = "^pATA nº setenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e dois/"
            .Replacement.Text = "^pATA nº setenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e três/"
            .Replacement.Text = "^pATA nº setenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e quatro/"
            .Replacement.Text = "^pATA nº setenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e cinco/"
            .Replacement.Text = "^pATA nº setenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e seis/"
            .Replacement.Text = "^pATA nº setenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e sete/"
            .Replacement.Text = "^pATA nº setenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e oito/"
            .Replacement.Text = "^pATA nº setenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº setenta e nove/"
            .Replacement.Text = "^pATA nº setenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta/"
            .Replacement.Text = "^pATA nº oitenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e um/"
            .Replacement.Text = "^pATA nº oitenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e dois/"
            .Replacement.Text = "^pATA nº oitenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e três/"
            .Replacement.Text = "^pATA nº oitenta e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e quatro/"
            .Replacement.Text = "^pATA nº oitenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e cinco/"
            .Replacement.Text = "^pATA nº oitenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e seis/"
            .Replacement.Text = "^pATA nº oitenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e sete/"
            .Replacement.Text = "^pATA nº oitenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e oito/"
            .Replacement.Text = "^pATA nº oitenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº oitenta e nove/"
            .Replacement.Text = "^pATA nº oitenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa/"
            .Replacement.Text = "^pATA nº noventa/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e um/"
            .Replacement.Text = "^pATA nº noventa e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e dois/"
            .Replacement.Text = "^pATA nº noventa e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e três/"
            .Replacement.Text = "^pATA nº noventa e três/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e quatro/"
            .Replacement.Text = "^pATA nº noventa e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e cinco/"
            .Replacement.Text = "^pATA nº noventa e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e seis/"
            .Replacement.Text = "^pATA nº noventa e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e sete/"
            .Replacement.Text = "^pATA nº noventa e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e oito/"
            .Replacement.Text = "^pATA nº noventa e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA nº noventa e nove/"
            .Replacement.Text = "^pATA nº noventa e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll


    MsgBox "Parte 2 ok", vbInformation, "Show de bola!"
    
End Sub
Sub B_URNA_3()
Attribute B_URNA_3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.B_URNA_3"
'
' B_URNA_3 Macro
'
'




        ' QuebraTítuloSeções Macro
        ' Adaptado para as atas de 2017 que dizem comunicação de liderança
        '
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "SESSÃO PLENÁRIA ORDINÁRIA"
                .Replacement.Text = "^pSESSÃO PLENÁRIA ORDINÁRIA^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "PERÍODO DAS COMUNICAÇÕES"
                .Replacement.Text = "^pPERÍODO DAS COMUNICAÇÕES^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "PERÍODO DAS COMUNICAÇÕES"
                .Replacement.Text = "^pPERÍODO DAS COMUNICAÇÕES^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "GRANDE EXPEDIENTE"
                .Replacement.Text = "^pGRANDE EXPEDIENTE^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "COMUNICAÇÃO DE LIDERANÇA"
                .Replacement.Text = "^pCOMUNICAÇÃO DE LIDERANÇA^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "COMUNICAÇÃO DE LIDERANÇA"
                .Replacement.Text = "^pCOMUNICAÇÃO DE LIDERANÇA^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "COMUNICAÇÃO DE LIDERANÇA"
                .Replacement.Text = "^pCOMUNICAÇÃO DE LIDERANÇA^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll

            ' AtribuiEstiloSeções Macro
            '
            '
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 1")
                With Selection.Find
                    .Text = "ATA nº"
                    .Replacement.Text = "ATA nº"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 2")
                With Selection.Find
                    .Text = "SESSÃO PLENÁRIA ORDINÁRIA"
                    .Replacement.Text = "SESSÃO PLENÁRIA ORDINÁRIA"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 2")
                With Selection.Find
                    .Text = "PERÍODO DAS COMUNICAÇÕES"
                    .Replacement.Text = "PERÍODO DAS COMUNICAÇÕES"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                With Selection.Find
                    .Text = "GRANDE EXPEDIENTE"
                    .Replacement.Text = "GRANDE EXPEDIENTE"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                With Selection.Find
                    .Text = "ESPAÇO DE LIDERANÇA"
                    .Replacement.Text = "ESPAÇO DE LIDERANÇA"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                With Selection.Find
                    .Text = "COMUNICAÇÃO DE LIDERANÇA"
                    .Replacement.Text = "COMUNICAÇÃO DE LIDERANÇA"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll


                ' QuebraLinhasVereadores Macro
                ' Adicionada Lorena, João da Silva
                '
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador João da Silva Chaves"
                    .Replacement.Text = "^pVereadora João da Silva Chaves^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Lorena dos Santos"
                    .Replacement.Text = "^pVereadora Lorena dos Santos^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Adelar Vargas"
                    .Replacement.Text = "^pVereador Adelar Vargas^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Admar Pozzobom"
                    .Replacement.Text = "^pVereador Admar Pozzobom^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Alexandre Vargas"
                    .Replacement.Text = "^pVereador Alexandre Vargas^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Celita da Silva"
                    .Replacement.Text = "^pVereadora Celita da Silva^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Daniel Diniz"
                    .Replacement.Text = "^pVereador Daniel Diniz^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Deili Silva"
                    .Replacement.Text = "^pVereadora Deili Silva^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Francisco Harrisson"
                    .Replacement.Text = "^pVereador Francisco Harrisson^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador João Kaus"
                    .Replacement.Text = "^pVereador João Kaus^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador João Chaves"
                    .Replacement.Text = "^pVereador João Chaves^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador João Ricardo Vargas"
                    .Replacement.Text = "^pVereador João Ricardo Vargas^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Jorge Trindade Soares"
                    .Replacement.Text = "^pVereador Jorge Trindade Soares^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Juliano Soares"
                    .Replacement.Text = "^pVereador Juliano Soares^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Leopoldo Ochulaki"
                    .Replacement.Text = "^pVereador Leopoldo Ochulaki^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Leopoldo Vanderlei Ochulaki"
                    .Replacement.Text = "^pVereador Leopoldo Vanderlei Ochulaki^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Luci Duartes"
                    .Replacement.Text = "^pVereadora Luci Duartes^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Luciano Guerra"
                    .Replacement.Text = "^pVereador Luciano Guerra^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Manoel Badke"
                    .Replacement.Text = "^pVereador Manoel Badke^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Maria Aparecida Brizola"
                    .Replacement.Text = "^pVereadora Maria Aparecida Brizola^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Marion Mortari"
                    .Replacement.Text = "^pVereador Marion Mortari^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereadora Marta Zanella"
                    .Replacement.Text = "^pVereadora Marta Zanella^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Ovídio Mayer"
                    .Replacement.Text = "^pVereador Ovídio Mayer^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Valdir Oliveira"
                    .Replacement.Text = "^pVereador Valdir Oliveira^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Vanderlei Araújo"
                    .Replacement.Text = "^pVereador Vanderlei Araújo^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador André Agne Domingues"
                    .Replacement.Text = "^pVereador André Agne Domingues^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Cezar Gehm"
                    .Replacement.Text = "^pVereador Cezar Gehm^p"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll


                ' AtribuiEstiloVereadores Macro
                ' Adicionada Lorena e João da Silva
                '
                '
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador João da Silva Chaves"
                    .Replacement.Text = "Vereador João da Silva Chaves"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Lorena dos Santos"
                    .Replacement.Text = "Vereadora Lorena dos Santos"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Adelar Vargas"
                    .Replacement.Text = "Vereador Adelar Vargas"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Admar Pozzobom"
                    .Replacement.Text = "Vereador Admar Pozzobom"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Alexandre Vargas"
                    .Replacement.Text = "Vereador Alexandre Vargas"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Celita da Silva"
                    .Replacement.Text = "Vereadora Celita da Silva"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Daniel Diniz"
                    .Replacement.Text = "Vereador Daniel Diniz"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Deili Silva"
                    .Replacement.Text = "Vereadora Deili Silva"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Francisco Harrisson"
                    .Replacement.Text = "Vereador Francisco Harrisson"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador João Kaus"
                    .Replacement.Text = "Vereador João Kaus"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador João Chaves"
                    .Replacement.Text = "Vereador João Chaves"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador João Ricardo Vargas"
                    .Replacement.Text = "Vereador João Ricardo Vargas"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Jorge Trindade Soares"
                    .Replacement.Text = "Vereador Jorge Trindade Soares"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Juliano Soares"
                    .Replacement.Text = "Vereador Juliano Soares"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Leopoldo Ochulaki"
                    .Replacement.Text = "Vereador Leopoldo Ochulaki"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Leopoldo Vanderlei Ochulaki"
                    .Replacement.Text = "Vereador Leopoldo Vanderlei Ochulaki"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Luci Duartes"
                    .Replacement.Text = "Vereadora Luci Duartes"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Luciano Guerra"
                    .Replacement.Text = "Vereador Luciano Guerra"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Manoel Badke"
                    .Replacement.Text = "Vereador Manoel Badke"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Maria Aparecida Brizola"
                    .Replacement.Text = "Vereadora Maria Aparecida Brizola"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Marion Mortari"
                    .Replacement.Text = "Vereador Marion Mortari"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereadora Marta Zanella"
                    .Replacement.Text = "Vereadora Marta Zanella"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Ovídio Mayer"
                    .Replacement.Text = "Vereador Ovídio Mayer"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Valdir Oliveira"
                    .Replacement.Text = "Vereador Valdir Oliveira"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Vanderlei Araújo"
                    .Replacement.Text = "Vereador Vanderlei Araújo"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador André Agne Domingues"
                    .Replacement.Text = "Vereador André Agne Domingues"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("Título 3")
                With Selection.Find
                    .Text = "Vereador Cezar Gehm"
                    .Replacement.Text = "Vereador Cezar Gehm"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll

    MsgBox "Parte 3 ok", vbInformation, "Show de bola!"


End Sub
Sub B_URNA_2B_AdequarAtasDe2017()
Attribute B_URNA_2B_AdequarAtasDe2017.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.BOCA_AdequarAtasDe2017"
'
' B_URNA_2B_AdequarAtasDe2017 Macro
'
'
         With Selection.Find
        .Text = "Cel Vargas "
        .Replacement.Text = "João Ricardo Vargas "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Cel. Vargas "
        .Replacement.Text = "João Ricardo Vargas "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Luci "
        .Replacement.Text = "Luci Duartes "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Deili "
        .Replacement.Text = "Deili Silva "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
             With Selection.Find
        .Text = "Vereador Maneco "
        .Replacement.Text = "Vereador Manoel Badke "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
             With Selection.Find
        .Text = "Lorena "
        .Replacement.Text = "Lorena dos Santos "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
             With Selection.Find
        .Text = "Juba "
        .Replacement.Text = "Juliano Soares "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Jorjão "
        .Replacement.Text = "Jorge Trindade Soares "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
             With Selection.Find
        .Text = "Celita "
        .Replacement.Text = "Celita da Silva "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Bolinha "
        .Replacement.Text = "Adelar Vargas "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Alemão do Gás "
        .Replacement.Text = "Leopoldo Vanderlei Ochulaki "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    
    
        With Selection.Find
        .Text = "Silva Silva"
        .Replacement.Text = "Silva"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "dos Santos dos Santos"
        .Replacement.Text = "dos Santos"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "Duartes Duartes"
        .Replacement.Text = "Duartes"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "da Silva da Silva"
        .Replacement.Text = "da Silva"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "PERÍODO DA COMUNICAÇÃO"
        .Replacement.Text = "PERÍODO DAS COMUNICAÇÕES"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "Vereador Dr. Ovidio "
        .Replacement.Text = "Vereador Ovídio Mayer "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find
        .Text = "Vanderley "
        .Replacement.Text = "Vanderlei "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Ovido "
        .Replacement.Text = "Ovídio Mayer "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Ovidio "
        .Replacement.Text = "Ovídio Mayer "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
             With Selection.Find
        .Text = "Vereador Dr. Francisco "
        .Replacement.Text = "Vereador Francisco Harrisson "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Dra Cida Brizola "
        .Replacement.Text = "Vereadora Maria Aparecida Brizola "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Drª Cida Brizola "
        .Replacement.Text = "Vereadora Maria Aparecida Brizola "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Cida Brizola "
        .Replacement.Text = "Vereadora Maria Aparecida Brizola "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereador Dr Deili Silva "
        .Replacement.Text = "Vereadora Deili Silva "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Pastora Lorena "
        .Replacement.Text = "Vereadora Lorena dos Santos "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereador Coronel Vargas "
        .Replacement.Text = "Vereador João Ricardo Vargas "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereador Dr. Francisco Harrisson "
        .Replacement.Text = "Vereador Francisco Harrisson "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Professora Luci Tia da Moto "
        .Replacement.Text = "Vereadora Luci Duartes "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Professora Celita da Silva "
        .Replacement.Text = "Vereadora Celita da Silva "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Profª Luci Tia da Moto "
        .Replacement.Text = "Vereadora Luci Duartes "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Profª Celita da Silva "
        .Replacement.Text = "Vereadora Celita da Silva "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Prof Celita da Silva "
        .Replacement.Text = "Vereadora Celita da Silva "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
         With Selection.Find
        .Text = "Vereadora Prof Luci Tia da Moto "
        .Replacement.Text = "Vereadora Luci Duartes "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
    
        MsgBox "Parte 2B ok", vbInformation, "Show de bola!"
        
        
End Sub
