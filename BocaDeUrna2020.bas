Attribute VB_Name = "BocaDeUrna2020"

Sub AlgarismosParaNumExtensoAtasBocaDeUrna2020()
Attribute AlgarismosParaNumExtensoAtasBocaDeUrna2020.VB_Description = "Remove n�meros de 0 a 9 alinhados a esquerda, direita ou justificados para tratamento de atas no projeto Boca de Urna 2020."
Attribute AlgarismosParaNumExtensoAtasBocaDeUrna2020.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.RemoverNumerosExcetoCentralizados"
'
' AlgarismosParaNumExtensoAtasBocaDeUrna2020 Macro
' Altera algarismos de 01 a 99 do t�tulo das atas para n�meros por extenso para tratamento de atas no projeto Boca de Urna 2020.
' Rascunho
' Usado nas atas de 2019 e 2020
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "ATA n� 01/"
    .Replacement.Text = "ATA n� um/"
    .Forward = True
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
    .Text = "ATA n� 02/"
    .Replacement.Text = "ATA n� dois/"
    .Forward = True
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
    .Text = "ATA n� 03/"
    .Replacement.Text = "ATA n� tr�s/"
    .Forward = True
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
    .Text = "ATA n� 04/"
    .Replacement.Text = "ATA n� quatro/"
    .Forward = True
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
    .Text = "ATA n� 05/"
    .Replacement.Text = "ATA n� cinco/"
    .Forward = True
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
    .Text = "ATA n� 06/"
    .Replacement.Text = "ATA n� seis/"
    .Forward = True
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
    .Text = "ATA n� 07/"
    .Replacement.Text = "ATA n� sete/"
    .Forward = True
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
    .Text = "ATA n� 08/"
    .Replacement.Text = "ATA n� oito/"
    .Forward = True
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
    .Text = "ATA n� 09/"
    .Replacement.Text = "ATA n� nove/"
    .Forward = True
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
    .Text = "ATA n� 10/"
    .Replacement.Text = "ATA n� dez/"
    .Forward = True
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
    .Text = "ATA n� 11/"
    .Replacement.Text = "ATA n� onze/"
    .Forward = True
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
    .Text = "ATA n� 12/"
    .Replacement.Text = "ATA n� doze/"
    .Forward = True
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
    .Text = "ATA n� 13/"
    .Replacement.Text = "ATA n� treze/"
    .Forward = True
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
    .Text = "ATA n� 14/"
    .Replacement.Text = "ATA n� quatorze/"
    .Forward = True
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
    .Text = "ATA n� 15/"
    .Replacement.Text = "ATA n� quinze/"
    .Forward = True
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
    .Text = "ATA n� 16/"
    .Replacement.Text = "ATA n� dezesseis/"
    .Forward = True
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
    .Text = "ATA n� 17/"
    .Replacement.Text = "ATA n� dezessete/"
    .Forward = True
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
    .Text = "ATA n� 18/"
    .Replacement.Text = "ATA n� dezoito/"
    .Forward = True
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
    .Text = "ATA n� 19/"
    .Replacement.Text = "ATA n� dezenove/"
    .Forward = True
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
    .Text = "ATA n� 20/"
    .Replacement.Text = "ATA n� vinte/"
    .Forward = True
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
    .Text = "ATA n� 21/"
    .Replacement.Text = "ATA n� vinte e um/"
    .Forward = True
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
    .Text = "ATA n� 22/"
    .Replacement.Text = "ATA n� vinte e dois/"
    .Forward = True
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
    .Text = "ATA n� 23/"
    .Replacement.Text = "ATA n� vinte e tr�s/"
    .Forward = True
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
    .Text = "ATA n� 24/"
    .Replacement.Text = "ATA n� vinte e quatro/"
    .Forward = True
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
    .Text = "ATA n� 25/"
    .Replacement.Text = "ATA n� vinte e cinco/"
    .Forward = True
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
    .Text = "ATA n� 26/"
    .Replacement.Text = "ATA n� vinte e seis/"
    .Forward = True
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
    .Text = "ATA n� 27/"
    .Replacement.Text = "ATA n� vinte e sete/"
    .Forward = True
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
    .Text = "ATA n� 28/"
    .Replacement.Text = "ATA n� vinte e oito/"
    .Forward = True
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
    .Text = "ATA n� 29/"
    .Replacement.Text = "ATA n� vinte e nove/"
    .Forward = True
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
    .Text = "ATA n� 30/"
    .Replacement.Text = "ATA n� trinta/"
    .Forward = True
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
    .Text = "ATA n� 31/"
    .Replacement.Text = "ATA n� trinta e um/"
    .Forward = True
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
    .Text = "ATA n� 32/"
    .Replacement.Text = "ATA n� trinta e dois/"
    .Forward = True
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
    .Text = "ATA n� 33/"
    .Replacement.Text = "ATA n� trinta e tr�s/"
    .Forward = True
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
    .Text = "ATA n� 34/"
    .Replacement.Text = "ATA n� trinta e quatro/"
    .Forward = True
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
    .Text = "ATA n� 35/"
    .Replacement.Text = "ATA n� trinta e cinco/"
    .Forward = True
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
    .Text = "ATA n� 36/"
    .Replacement.Text = "ATA n� trinta e seis/"
    .Forward = True
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
    .Text = "ATA n� 37/"
    .Replacement.Text = "ATA n� trinta e sete/"
    .Forward = True
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
    .Text = "ATA n� 38/"
    .Replacement.Text = "ATA n� trinta e oito/"
    .Forward = True
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
    .Text = "ATA n� 39/"
    .Replacement.Text = "ATA n� trinta e nove/"
    .Forward = True
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
    .Text = "ATA n� 40/"
    .Replacement.Text = "ATA n� quarenta/"
    .Forward = True
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
    .Text = "ATA n� 41/"
    .Replacement.Text = "ATA n� quarenta e um/"
    .Forward = True
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
    .Text = "ATA n� 42/"
    .Replacement.Text = "ATA n� quarenta e dois/"
    .Forward = True
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
    .Text = "ATA n� 43/"
    .Replacement.Text = "ATA n� quarenta e tr�s/"
    .Forward = True
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
    .Text = "ATA n� 44/"
    .Replacement.Text = "ATA n� quarenta e quatro/"
    .Forward = True
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
    .Text = "ATA n� 45/"
    .Replacement.Text = "ATA n� quarenta e cinco/"
    .Forward = True
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
    .Text = "ATA n� 46/"
    .Replacement.Text = "ATA n� quarenta e seis/"
    .Forward = True
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
    .Text = "ATA n� 47/"
    .Replacement.Text = "ATA n� quarenta e sete/"
    .Forward = True
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
    .Text = "ATA n� 48/"
    .Replacement.Text = "ATA n� quarenta e oito/"
    .Forward = True
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
    .Text = "ATA n� 49/"
    .Replacement.Text = "ATA n� quarenta e nove/"
    .Forward = True
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
    .Text = "ATA n� 50/"
    .Replacement.Text = "ATA n� cinquenta/"
    .Forward = True
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
    .Text = "ATA n� 51/"
    .Replacement.Text = "ATA n� cinquenta e um/"
    .Forward = True
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
    .Text = "ATA n� 52/"
    .Replacement.Text = "ATA n� cinquenta e dois/"
    .Forward = True
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
    .Text = "ATA n� 53/"
    .Replacement.Text = "ATA n� cinquenta e tr�s/"
    .Forward = True
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
    .Text = "ATA n� 54/"
    .Replacement.Text = "ATA n� cinquenta e quatro/"
    .Forward = True
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
    .Text = "ATA n� 55/"
    .Replacement.Text = "ATA n� cinquenta e cinco/"
    .Forward = True
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
    .Text = "ATA n� 56/"
    .Replacement.Text = "ATA n� cinquenta e seis/"
    .Forward = True
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
    .Text = "ATA n� 57/"
    .Replacement.Text = "ATA n� cinquenta e sete/"
    .Forward = True
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
    .Text = "ATA n� 58/"
    .Replacement.Text = "ATA n� cinquenta e oito/"
    .Forward = True
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
    .Text = "ATA n� 59/"
    .Replacement.Text = "ATA n� cinquenta e nove/"
    .Forward = True
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
    .Text = "ATA n� 60/"
    .Replacement.Text = "ATA n� sessenta/"
    .Forward = True
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
    .Text = "ATA n� 61/"
    .Replacement.Text = "ATA n� sessenta e um/"
    .Forward = True
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
    .Text = "ATA n� 62/"
    .Replacement.Text = "ATA n� sessenta e dois/"
    .Forward = True
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
    .Text = "ATA n� 63/"
    .Replacement.Text = "ATA n� sessenta e tr�s/"
    .Forward = True
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
    .Text = "ATA n� 64/"
    .Replacement.Text = "ATA n� sessenta e quatro/"
    .Forward = True
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
    .Text = "ATA n� 65/"
    .Replacement.Text = "ATA n� sessenta e cinco/"
    .Forward = True
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
    .Text = "ATA n� 66/"
    .Replacement.Text = "ATA n� sessenta e seis/"
    .Forward = True
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
    .Text = "ATA n� 67/"
    .Replacement.Text = "ATA n� sessenta e sete/"
    .Forward = True
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
    .Text = "ATA n� 68/"
    .Replacement.Text = "ATA n� sessenta e oito/"
    .Forward = True
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
    .Text = "ATA n� 69/"
    .Replacement.Text = "ATA n� sessenta e nove/"
    .Forward = True
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
    .Text = "ATA n� 70/"
    .Replacement.Text = "ATA n� setenta/"
    .Forward = True
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
    .Text = "ATA n� 71/"
    .Replacement.Text = "ATA n� setenta e um/"
    .Forward = True
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
    .Text = "ATA n� 72/"
    .Replacement.Text = "ATA n� setenta e dois/"
    .Forward = True
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
    .Text = "ATA n� 73/"
    .Replacement.Text = "ATA n� setenta e tr�s/"
    .Forward = True
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
    .Text = "ATA n� 74/"
    .Replacement.Text = "ATA n� setenta e quatro/"
    .Forward = True
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
    .Text = "ATA n� 75/"
    .Replacement.Text = "ATA n� setenta e cinco/"
    .Forward = True
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
    .Text = "ATA n� 76/"
    .Replacement.Text = "ATA n� setenta e seis/"
    .Forward = True
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
    .Text = "ATA n� 77/"
    .Replacement.Text = "ATA n� setenta e sete/"
    .Forward = True
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
    .Text = "ATA n� 78/"
    .Replacement.Text = "ATA n� setenta e oito/"
    .Forward = True
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
    .Text = "ATA n� 79/"
    .Replacement.Text = "ATA n� setenta e nove/"
    .Forward = True
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
    .Text = "ATA n� 80/"
    .Replacement.Text = "ATA n� oitenta/"
    .Forward = True
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
    .Text = "ATA n� 81/"
    .Replacement.Text = "ATA n� oitenta e um/"
    .Forward = True
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
    .Text = "ATA n� 82/"
    .Replacement.Text = "ATA n� oitenta e dois/"
    .Forward = True
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
    .Text = "ATA n� 83/"
    .Replacement.Text = "ATA n� oitenta e tr�s/"
    .Forward = True
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
    .Text = "ATA n� 84/"
    .Replacement.Text = "ATA n� oitenta e quatro/"
    .Forward = True
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
    .Text = "ATA n� 85/"
    .Replacement.Text = "ATA n� oitenta e cinco/"
    .Forward = True
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
    .Text = "ATA n� 86/"
    .Replacement.Text = "ATA n� oitenta e seis/"
    .Forward = True
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
    .Text = "ATA n� 87/"
    .Replacement.Text = "ATA n� oitenta e sete/"
    .Forward = True
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
    .Text = "ATA n� 88/"
    .Replacement.Text = "ATA n� oitenta e oito/"
    .Forward = True
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
    .Text = "ATA n� 89/"
    .Replacement.Text = "ATA n� oitenta e nove/"
    .Forward = True
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
    .Text = "ATA n� 90/"
    .Replacement.Text = "ATA n� noventa/"
    .Forward = True
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
    .Text = "ATA n� 91/"
    .Replacement.Text = "ATA n� noventa e um/"
    .Forward = True
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
    .Text = "ATA n� 92/"
    .Replacement.Text = "ATA n� noventa e dois/"
    .Forward = True
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
    .Text = "ATA n� 93/"
    .Replacement.Text = "ATA n� noventa e tr�s/"
    .Forward = True
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
    .Text = "ATA n� 94/"
    .Replacement.Text = "ATA n� noventa e quatro/"
    .Forward = True
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
    .Text = "ATA n� 95/"
    .Replacement.Text = "ATA n� noventa e cinco/"
    .Forward = True
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
    .Text = "ATA n� 96/"
    .Replacement.Text = "ATA n� noventa e seis/"
    .Forward = True
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
    .Text = "ATA n� 97/"
    .Replacement.Text = "ATA n� noventa e sete/"
    .Forward = True
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
    .Text = "ATA n� 98/"
    .Replacement.Text = "ATA n� noventa e oito/"
    .Forward = True
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
    .Text = "ATA n� 99/"
    .Replacement.Text = "ATA n� noventa e nove/"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

MsgBox "Algarismos nos t�tulos das atas substitu�dos por n�meros por extenso.", vbInformation, "Show de bola!"

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
        .Text = "C�MARA MUNICIPAL DE VEREADORES DE SANTA MARIA � RS" & vbTab _
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

    MsgBox "Deletados: cabe�alhos errados, quebras de linha e de se��o, espa�os duplos, triplos e qu�druplos.", vbInformation, "Show de bola!"

End Sub


Sub QuebraT�tuloAtas()
Attribute QuebraT�tuloAtas.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.QuebraT�tuloAtas"
'
' QuebraT�tuloAtas Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "ATA n� um/"
    .Replacement.Text = "^pATA n� um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� dois/"
    .Replacement.Text = "^pATA n� dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� tr�s/"
    .Replacement.Text = "^pATA n� tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quatro/"
    .Replacement.Text = "^pATA n� quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinco/"
    .Replacement.Text = "^pATA n� cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� seis/"
    .Replacement.Text = "^pATA n� seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sete/"
    .Replacement.Text = "^pATA n� sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oito/"
    .Replacement.Text = "^pATA n� oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� nove/"
    .Replacement.Text = "^pATA n� nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� dez/"
    .Replacement.Text = "^pATA n� dez/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� onze/"
    .Replacement.Text = "^pATA n� onze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� doze/"
    .Replacement.Text = "^pATA n� doze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� treze/"
    .Replacement.Text = "^pATA n� treze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quatorze/"
    .Replacement.Text = "^pATA n� quatorze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quinze/"
    .Replacement.Text = "^pATA n� quinze/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� dezesseis/"
    .Replacement.Text = "^pATA n� dezesseis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� dezessete/"
    .Replacement.Text = "^pATA n� dezessete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� dezoito/"
    .Replacement.Text = "^pATA n� dezoito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� dezenove/"
    .Replacement.Text = "^pATA n� dezenove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte/"
    .Replacement.Text = "^pATA n� vinte/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e um/"
    .Replacement.Text = "^pATA n� vinte e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e dois/"
    .Replacement.Text = "^pATA n� vinte e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e tr�s/"
    .Replacement.Text = "^pATA n� vinte e tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e quatro/"
    .Replacement.Text = "^pATA n� vinte e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e cinco/"
    .Replacement.Text = "^pATA n� vinte e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e seis/"
    .Replacement.Text = "^pATA n� vinte e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e sete/"
    .Replacement.Text = "^pATA n� vinte e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e oito/"
    .Replacement.Text = "^pATA n� vinte e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� vinte e nove/"
    .Replacement.Text = "^pATA n� vinte e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta/"
    .Replacement.Text = "^pATA n� trinta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e um/"
    .Replacement.Text = "^pATA n� trinta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e dois/"
    .Replacement.Text = "^pATA n� trinta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e tr�s/"
    .Replacement.Text = "^pATA n� trinta e tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e quatro/"
    .Replacement.Text = "^pATA n� trinta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e cinco/"
    .Replacement.Text = "^pATA n� trinta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e seis/"
    .Replacement.Text = "^pATA n� trinta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e sete/"
    .Replacement.Text = "^pATA n� trinta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e oito/"
    .Replacement.Text = "^pATA n� trinta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� trinta e nove/"
    .Replacement.Text = "^pATA n� trinta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta/"
    .Replacement.Text = "^pATA n� quarenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e um/"
    .Replacement.Text = "^pATA n� quarenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e dois/"
    .Replacement.Text = "^pATA n� quarenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e tr�s/"
    .Replacement.Text = "^pATA n� quarenta e tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e quatro/"
    .Replacement.Text = "^pATA n� quarenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e cinco/"
    .Replacement.Text = "^pATA n� quarenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e seis/"
    .Replacement.Text = "^pATA n� quarenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e sete/"
    .Replacement.Text = "^pATA n� quarenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e oito/"
    .Replacement.Text = "^pATA n� quarenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� quarenta e nove/"
    .Replacement.Text = "^pATA n� quarenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta/"
    .Replacement.Text = "^pATA n� cinquenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e um/"
    .Replacement.Text = "^pATA n� cinquenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e dois/"
    .Replacement.Text = "^pATA n� cinquenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e tr�s/"
    .Replacement.Text = "^pATA n� cinquenta e tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e quatro/"
    .Replacement.Text = "^pATA n� cinquenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e cinco/"
    .Replacement.Text = "^pATA n� cinquenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e seis/"
    .Replacement.Text = "^pATA n� cinquenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e sete/"
    .Replacement.Text = "^pATA n� cinquenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e oito/"
    .Replacement.Text = "^pATA n� cinquenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� cinquenta e nove/"
    .Replacement.Text = "^pATA n� cinquenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta/"
    .Replacement.Text = "^pATA n� sessenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e um/"
    .Replacement.Text = "^pATA n� sessenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e dois/"
    .Replacement.Text = "^pATA n� sessenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e tr�s/"
    .Replacement.Text = "^pATA n� sessenta e tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e quatro/"
    .Replacement.Text = "^pATA n� sessenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e cinco/"
    .Replacement.Text = "^pATA n� sessenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e seis/"
    .Replacement.Text = "^pATA n� sessenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e sete/"
    .Replacement.Text = "^pATA n� sessenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e oito/"
    .Replacement.Text = "^pATA n� sessenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� sessenta e nove/"
    .Replacement.Text = "^pATA n� sessenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta/"
    .Replacement.Text = "^pATA n� setenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e um/"
    .Replacement.Text = "^pATA n� setenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e dois/"
    .Replacement.Text = "^pATA n� setenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e tr�s/"
    .Replacement.Text = "^pATA n� setenta e tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e quatro/"
    .Replacement.Text = "^pATA n� setenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e cinco/"
    .Replacement.Text = "^pATA n� setenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e seis/"
    .Replacement.Text = "^pATA n� setenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e sete/"
    .Replacement.Text = "^pATA n� setenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e oito/"
    .Replacement.Text = "^pATA n� setenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� setenta e nove/"
    .Replacement.Text = "^pATA n� setenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta/"
    .Replacement.Text = "^pATA n� oitenta/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e um/"
    .Replacement.Text = "^pATA n� oitenta e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e dois/"
    .Replacement.Text = "^pATA n� oitenta e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e tr�s/"
    .Replacement.Text = "^pATA n� oitenta e tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e quatro/"
    .Replacement.Text = "^pATA n� oitenta e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e cinco/"
    .Replacement.Text = "^pATA n� oitenta e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e seis/"
    .Replacement.Text = "^pATA n� oitenta e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e sete/"
    .Replacement.Text = "^pATA n� oitenta e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e oito/"
    .Replacement.Text = "^pATA n� oitenta e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� oitenta e nove/"
    .Replacement.Text = "^pATA n� oitenta e nove/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa/"
    .Replacement.Text = "^pATA n� noventa/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e um/"
    .Replacement.Text = "^pATA n� noventa e um/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e dois/"
    .Replacement.Text = "^pATA n� noventa e dois/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e tr�s/"
    .Replacement.Text = "^pATA n� noventa e tr�s/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e quatro/"
    .Replacement.Text = "^pATA n� noventa e quatro/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e cinco/"
    .Replacement.Text = "^pATA n� noventa e cinco/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e seis/"
    .Replacement.Text = "^pATA n� noventa e seis/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e sete/"
    .Replacement.Text = "^pATA n� noventa e sete/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e oito/"
    .Replacement.Text = "^pATA n� noventa e oito/^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
    .Text = "ATA n� noventa e nove/"
    .Replacement.Text = "^pATA n� noventa e nove/^p"
    .Forward = True
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
Sub QuebraT�tuloSe��es()
Attribute QuebraT�tuloSe��es.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.QuebraT�tuloSe��es"
'
' QuebraT�tuloSe��es Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "SESS�O PLEN�RIA ORDIN�RIA"
        .Replacement.Text = "^pSESS�O PLEN�RIA ORDIN�RIA^p"
        .Forward = True
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
        .Text = "PER�ODO DAS COMUNICA��ES"
        .Replacement.Text = "^pPER�ODO DAS COMUNICA��ES^p"
        .Forward = True
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
        .Text = "PER�ODO DAS COMUNICA��ES"
        .Replacement.Text = "^pPER�ODO DAS COMUNICA��ES^p"
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
        .Text = "ESPA�O DE LIDERAN�A"
        .Replacement.Text = "^pESPA�O DE LIDERAN�A^p"
        .Forward = True
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
        .Text = "ESPA�O DE LIDERAN�A"
        .Replacement.Text = "^pESPA�O DE LIDERAN�A^p"
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
        .Text = "ESPA�O DE LIDERAN�A"
        .Replacement.Text = "^pESPA�O DE LIDERAN�A^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    MsgBox "Quebra de linha antes e depois das se��es inserida.", vbInformation, "Show de bola!"

End Sub
Sub AtribuiEstiloSe��es()
Attribute AtribuiEstiloSe��es.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.AtribuiEstiloSe��es"
'
' AtribuiEstiloSe��es Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 1")
    With Selection.Find
        .Text = "ATA n�"
        .Replacement.Text = "ATA n�"
        .Forward = True
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
    Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 2")
    With Selection.Find
        .Text = "SESS�O PLEN�RIA ORDIN�RIA"
        .Replacement.Text = "SESS�O PLEN�RIA ORDIN�RIA"
        .Forward = True
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
    Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 2")
    With Selection.Find
        .Text = "PER�ODO DAS COMUNICA��ES"
        .Replacement.Text = "PER�ODO DAS COMUNICA��ES"
        .Forward = True
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
        .Text = "ESPA�O DE LIDERAN�A"
        .Replacement.Text = "ESPA�O DE LIDERAN�A"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    MsgBox "T�tulos prim�rios e secund�rios atribuidos.", vbInformation, "Show de bola!"

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
    .Text = "Vereador Jo�o Kaus"
    .Replacement.Text = "^pVereador Jo�o Kaus^p"
    .Forward = True
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
    .Text = "Vereador Jo�o Chaves"
    .Replacement.Text = "^pVereador Jo�o Chaves^p"
    .Forward = True
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
    .Text = "Vereador Jo�o Ricardo Vargas"
    .Replacement.Text = "^pVereador Jo�o Ricardo Vargas^p"
    .Forward = True
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
    .Text = "Vereador Ov�dio Mayer"
    .Replacement.Text = "^pVereador Ov�dio Mayer^p"
    .Forward = True
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
    .Text = "Vereador Vanderlei Ara�jo"
    .Replacement.Text = "^pVereador Vanderlei Ara�jo^p"
    .Forward = True
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
    .Text = "Vereador Andr� Agne Domingues"
    .Replacement.Text = "^pVereador Andr� Agne Domingues^p"
    .Forward = True
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

    MsgBox "Quebra de linha nas men��es aos vereadores.", vbInformation, "Show de bola!"

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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
With Selection.Find
    .Text = "Vereador Jo�o Kaus"
    .Replacement.Text = "Vereador Jo�o Kaus"
    .Forward = True
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
With Selection.Find
    .Text = "Vereador Jo�o Chaves"
    .Replacement.Text = "Vereador Jo�o Chaves"
    .Forward = True
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
With Selection.Find
    .Text = "Vereador Jo�o Ricardo Vargas"
    .Replacement.Text = "Vereador Jo�o Ricardo Vargas"
    .Forward = True
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
With Selection.Find
    .Text = "Vereador Ov�dio Mayer"
    .Replacement.Text = "Vereador Ov�dio Mayer"
    .Forward = True
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
With Selection.Find
    .Text = "Vereador Vanderlei Ara�jo"
    .Replacement.Text = "Vereador Vanderlei Ara�jo"
    .Forward = True
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
With Selection.Find
    .Text = "Vereador Andr� Agne Domingues"
    .Replacement.Text = "Vereador Andr� Agne Domingues"
    .Forward = True
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
Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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

    MsgBox "T�tulos terci�rios atribu�dos.", vbInformation, "Show de bola!"

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
        .Text = "C�MARA MUNICIPAL DE VEREADORES DE SANTA MARIA � RS" & vbTab _
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



    MsgBox "A princ�cio, tudo certo.", vbInformation, "Show de bola!"


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
        .Text = "C�MARA MUNICIPAL DE VEREADORES DE SANTA MARIA � RS" & vbTab _
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
    ' Altera algarismos de 01 a 99 do t�tulo das atas para n�meros por extenso para tratamento de atas no projeto Boca de Urna 2020.
    '
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ATA n� 01/"
        .Replacement.Text = "ATA n� um/"
        .Forward = True
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
        .Text = "ATA n� 02/"
        .Replacement.Text = "ATA n� dois/"
        .Forward = True
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
        .Text = "ATA n� 03/"
        .Replacement.Text = "ATA n� tr�s/"
        .Forward = True
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
        .Text = "ATA n� 04/"
        .Replacement.Text = "ATA n� quatro/"
        .Forward = True
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
        .Text = "ATA n� 05/"
        .Replacement.Text = "ATA n� cinco/"
        .Forward = True
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
        .Text = "ATA n� 06/"
        .Replacement.Text = "ATA n� seis/"
        .Forward = True
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
        .Text = "ATA n� 07/"
        .Replacement.Text = "ATA n� sete/"
        .Forward = True
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
        .Text = "ATA n� 08/"
        .Replacement.Text = "ATA n� oito/"
        .Forward = True
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
        .Text = "ATA n� 09/"
        .Replacement.Text = "ATA n� nove/"
        .Forward = True
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
        .Text = "ATA n� 10/"
        .Replacement.Text = "ATA n� dez/"
        .Forward = True
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
        .Text = "ATA n� 11/"
        .Replacement.Text = "ATA n� onze/"
        .Forward = True
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
        .Text = "ATA n� 12/"
        .Replacement.Text = "ATA n� doze/"
        .Forward = True
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
        .Text = "ATA n� 13/"
        .Replacement.Text = "ATA n� treze/"
        .Forward = True
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
        .Text = "ATA n� 14/"
        .Replacement.Text = "ATA n� quatorze/"
        .Forward = True
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
        .Text = "ATA n� 15/"
        .Replacement.Text = "ATA n� quinze/"
        .Forward = True
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
        .Text = "ATA n� 16/"
        .Replacement.Text = "ATA n� dezesseis/"
        .Forward = True
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
        .Text = "ATA n� 17/"
        .Replacement.Text = "ATA n� dezessete/"
        .Forward = True
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
        .Text = "ATA n� 18/"
        .Replacement.Text = "ATA n� dezoito/"
        .Forward = True
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
        .Text = "ATA n� 19/"
        .Replacement.Text = "ATA n� dezenove/"
        .Forward = True
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
        .Text = "ATA n� 20/"
        .Replacement.Text = "ATA n� vinte/"
        .Forward = True
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
        .Text = "ATA n� 21/"
        .Replacement.Text = "ATA n� vinte e um/"
        .Forward = True
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
        .Text = "ATA n� 22/"
        .Replacement.Text = "ATA n� vinte e dois/"
        .Forward = True
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
        .Text = "ATA n� 23/"
        .Replacement.Text = "ATA n� vinte e tr�s/"
        .Forward = True
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
        .Text = "ATA n� 24/"
        .Replacement.Text = "ATA n� vinte e quatro/"
        .Forward = True
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
        .Text = "ATA n� 25/"
        .Replacement.Text = "ATA n� vinte e cinco/"
        .Forward = True
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
        .Text = "ATA n� 26/"
        .Replacement.Text = "ATA n� vinte e seis/"
        .Forward = True
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
        .Text = "ATA n� 27/"
        .Replacement.Text = "ATA n� vinte e sete/"
        .Forward = True
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
        .Text = "ATA n� 28/"
        .Replacement.Text = "ATA n� vinte e oito/"
        .Forward = True
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
        .Text = "ATA n� 29/"
        .Replacement.Text = "ATA n� vinte e nove/"
        .Forward = True
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
        .Text = "ATA n� 30/"
        .Replacement.Text = "ATA n� trinta/"
        .Forward = True
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
        .Text = "ATA n� 31/"
        .Replacement.Text = "ATA n� trinta e um/"
        .Forward = True
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
        .Text = "ATA n� 32/"
        .Replacement.Text = "ATA n� trinta e dois/"
        .Forward = True
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
        .Text = "ATA n� 33/"
        .Replacement.Text = "ATA n� trinta e tr�s/"
        .Forward = True
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
        .Text = "ATA n� 34/"
        .Replacement.Text = "ATA n� trinta e quatro/"
        .Forward = True
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
        .Text = "ATA n� 35/"
        .Replacement.Text = "ATA n� trinta e cinco/"
        .Forward = True
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
        .Text = "ATA n� 36/"
        .Replacement.Text = "ATA n� trinta e seis/"
        .Forward = True
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
        .Text = "ATA n� 37/"
        .Replacement.Text = "ATA n� trinta e sete/"
        .Forward = True
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
        .Text = "ATA n� 38/"
        .Replacement.Text = "ATA n� trinta e oito/"
        .Forward = True
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
        .Text = "ATA n� 39/"
        .Replacement.Text = "ATA n� trinta e nove/"
        .Forward = True
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
        .Text = "ATA n� 40/"
        .Replacement.Text = "ATA n� quarenta/"
        .Forward = True
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
        .Text = "ATA n� 41/"
        .Replacement.Text = "ATA n� quarenta e um/"
        .Forward = True
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
        .Text = "ATA n� 42/"
        .Replacement.Text = "ATA n� quarenta e dois/"
        .Forward = True
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
        .Text = "ATA n� 43/"
        .Replacement.Text = "ATA n� quarenta e tr�s/"
        .Forward = True
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
        .Text = "ATA n� 44/"
        .Replacement.Text = "ATA n� quarenta e quatro/"
        .Forward = True
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
        .Text = "ATA n� 45/"
        .Replacement.Text = "ATA n� quarenta e cinco/"
        .Forward = True
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
        .Text = "ATA n� 46/"
        .Replacement.Text = "ATA n� quarenta e seis/"
        .Forward = True
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
        .Text = "ATA n� 47/"
        .Replacement.Text = "ATA n� quarenta e sete/"
        .Forward = True
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
        .Text = "ATA n� 48/"
        .Replacement.Text = "ATA n� quarenta e oito/"
        .Forward = True
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
        .Text = "ATA n� 49/"
        .Replacement.Text = "ATA n� quarenta e nove/"
        .Forward = True
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
        .Text = "ATA n� 50/"
        .Replacement.Text = "ATA n� cinquenta/"
        .Forward = True
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
        .Text = "ATA n� 51/"
        .Replacement.Text = "ATA n� cinquenta e um/"
        .Forward = True
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
        .Text = "ATA n� 52/"
        .Replacement.Text = "ATA n� cinquenta e dois/"
        .Forward = True
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
        .Text = "ATA n� 53/"
        .Replacement.Text = "ATA n� cinquenta e tr�s/"
        .Forward = True
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
        .Text = "ATA n� 54/"
        .Replacement.Text = "ATA n� cinquenta e quatro/"
        .Forward = True
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
        .Text = "ATA n� 55/"
        .Replacement.Text = "ATA n� cinquenta e cinco/"
        .Forward = True
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
        .Text = "ATA n� 56/"
        .Replacement.Text = "ATA n� cinquenta e seis/"
        .Forward = True
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
        .Text = "ATA n� 57/"
        .Replacement.Text = "ATA n� cinquenta e sete/"
        .Forward = True
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
        .Text = "ATA n� 58/"
        .Replacement.Text = "ATA n� cinquenta e oito/"
        .Forward = True
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
        .Text = "ATA n� 59/"
        .Replacement.Text = "ATA n� cinquenta e nove/"
        .Forward = True
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
        .Text = "ATA n� 60/"
        .Replacement.Text = "ATA n� sessenta/"
        .Forward = True
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
        .Text = "ATA n� 61/"
        .Replacement.Text = "ATA n� sessenta e um/"
        .Forward = True
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
        .Text = "ATA n� 62/"
        .Replacement.Text = "ATA n� sessenta e dois/"
        .Forward = True
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
        .Text = "ATA n� 63/"
        .Replacement.Text = "ATA n� sessenta e tr�s/"
        .Forward = True
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
        .Text = "ATA n� 64/"
        .Replacement.Text = "ATA n� sessenta e quatro/"
        .Forward = True
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
        .Text = "ATA n� 65/"
        .Replacement.Text = "ATA n� sessenta e cinco/"
        .Forward = True
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
        .Text = "ATA n� 66/"
        .Replacement.Text = "ATA n� sessenta e seis/"
        .Forward = True
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
        .Text = "ATA n� 67/"
        .Replacement.Text = "ATA n� sessenta e sete/"
        .Forward = True
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
        .Text = "ATA n� 68/"
        .Replacement.Text = "ATA n� sessenta e oito/"
        .Forward = True
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
        .Text = "ATA n� 69/"
        .Replacement.Text = "ATA n� sessenta e nove/"
        .Forward = True
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
        .Text = "ATA n� 70/"
        .Replacement.Text = "ATA n� setenta/"
        .Forward = True
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
        .Text = "ATA n� 71/"
        .Replacement.Text = "ATA n� setenta e um/"
        .Forward = True
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
        .Text = "ATA n� 72/"
        .Replacement.Text = "ATA n� setenta e dois/"
        .Forward = True
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
        .Text = "ATA n� 73/"
        .Replacement.Text = "ATA n� setenta e tr�s/"
        .Forward = True
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
        .Text = "ATA n� 74/"
        .Replacement.Text = "ATA n� setenta e quatro/"
        .Forward = True
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
        .Text = "ATA n� 75/"
        .Replacement.Text = "ATA n� setenta e cinco/"
        .Forward = True
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
        .Text = "ATA n� 76/"
        .Replacement.Text = "ATA n� setenta e seis/"
        .Forward = True
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
        .Text = "ATA n� 77/"
        .Replacement.Text = "ATA n� setenta e sete/"
        .Forward = True
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
        .Text = "ATA n� 78/"
        .Replacement.Text = "ATA n� setenta e oito/"
        .Forward = True
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
        .Text = "ATA n� 79/"
        .Replacement.Text = "ATA n� setenta e nove/"
        .Forward = True
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
        .Text = "ATA n� 80/"
        .Replacement.Text = "ATA n� oitenta/"
        .Forward = True
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
        .Text = "ATA n� 81/"
        .Replacement.Text = "ATA n� oitenta e um/"
        .Forward = True
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
        .Text = "ATA n� 82/"
        .Replacement.Text = "ATA n� oitenta e dois/"
        .Forward = True
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
        .Text = "ATA n� 83/"
        .Replacement.Text = "ATA n� oitenta e tr�s/"
        .Forward = True
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
        .Text = "ATA n� 84/"
        .Replacement.Text = "ATA n� oitenta e quatro/"
        .Forward = True
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
        .Text = "ATA n� 85/"
        .Replacement.Text = "ATA n� oitenta e cinco/"
        .Forward = True
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
        .Text = "ATA n� 86/"
        .Replacement.Text = "ATA n� oitenta e seis/"
        .Forward = True
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
        .Text = "ATA n� 87/"
        .Replacement.Text = "ATA n� oitenta e sete/"
        .Forward = True
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
        .Text = "ATA n� 88/"
        .Replacement.Text = "ATA n� oitenta e oito/"
        .Forward = True
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
        .Text = "ATA n� 89/"
        .Replacement.Text = "ATA n� oitenta e nove/"
        .Forward = True
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
        .Text = "ATA n� 90/"
        .Replacement.Text = "ATA n� noventa/"
        .Forward = True
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
        .Text = "ATA n� 91/"
        .Replacement.Text = "ATA n� noventa e um/"
        .Forward = True
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
        .Text = "ATA n� 92/"
        .Replacement.Text = "ATA n� noventa e dois/"
        .Forward = True
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
        .Text = "ATA n� 93/"
        .Replacement.Text = "ATA n� noventa e tr�s/"
        .Forward = True
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
        .Text = "ATA n� 94/"
        .Replacement.Text = "ATA n� noventa e quatro/"
        .Forward = True
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
        .Text = "ATA n� 95/"
        .Replacement.Text = "ATA n� noventa e cinco/"
        .Forward = True
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
        .Text = "ATA n� 96/"
        .Replacement.Text = "ATA n� noventa e seis/"
        .Forward = True
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
        .Text = "ATA n� 97/"
        .Replacement.Text = "ATA n� noventa e sete/"
        .Forward = True
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
        .Text = "ATA n� 98/"
        .Replacement.Text = "ATA n� noventa e oito/"
        .Forward = True
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
        .Text = "ATA n� 99/"
        .Replacement.Text = "ATA n� noventa e nove/"
        .Forward = True
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

        ' QuebraT�tuloAtas Macro
        '
        '
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "ATA n� um/"
            .Replacement.Text = "^pATA n� um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dois/"
            .Replacement.Text = "^pATA n� dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� tr�s/"
            .Replacement.Text = "^pATA n� tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quatro/"
            .Replacement.Text = "^pATA n� quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinco/"
            .Replacement.Text = "^pATA n� cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� seis/"
            .Replacement.Text = "^pATA n� seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sete/"
            .Replacement.Text = "^pATA n� sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oito/"
            .Replacement.Text = "^pATA n� oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� nove/"
            .Replacement.Text = "^pATA n� nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dez/"
            .Replacement.Text = "^pATA n� dez/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� onze/"
            .Replacement.Text = "^pATA n� onze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� doze/"
            .Replacement.Text = "^pATA n� doze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� treze/"
            .Replacement.Text = "^pATA n� treze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quatorze/"
            .Replacement.Text = "^pATA n� quatorze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quinze/"
            .Replacement.Text = "^pATA n� quinze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dezesseis/"
            .Replacement.Text = "^pATA n� dezesseis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dezessete/"
            .Replacement.Text = "^pATA n� dezessete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dezoito/"
            .Replacement.Text = "^pATA n� dezoito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dezenove/"
            .Replacement.Text = "^pATA n� dezenove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte/"
            .Replacement.Text = "^pATA n� vinte/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e um/"
            .Replacement.Text = "^pATA n� vinte e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e dois/"
            .Replacement.Text = "^pATA n� vinte e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e tr�s/"
            .Replacement.Text = "^pATA n� vinte e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e quatro/"
            .Replacement.Text = "^pATA n� vinte e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e cinco/"
            .Replacement.Text = "^pATA n� vinte e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e seis/"
            .Replacement.Text = "^pATA n� vinte e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e sete/"
            .Replacement.Text = "^pATA n� vinte e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e oito/"
            .Replacement.Text = "^pATA n� vinte e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e nove/"
            .Replacement.Text = "^pATA n� vinte e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta/"
            .Replacement.Text = "^pATA n� trinta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e um/"
            .Replacement.Text = "^pATA n� trinta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e dois/"
            .Replacement.Text = "^pATA n� trinta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e tr�s/"
            .Replacement.Text = "^pATA n� trinta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e quatro/"
            .Replacement.Text = "^pATA n� trinta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e cinco/"
            .Replacement.Text = "^pATA n� trinta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e seis/"
            .Replacement.Text = "^pATA n� trinta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e sete/"
            .Replacement.Text = "^pATA n� trinta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e oito/"
            .Replacement.Text = "^pATA n� trinta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e nove/"
            .Replacement.Text = "^pATA n� trinta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta/"
            .Replacement.Text = "^pATA n� quarenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e um/"
            .Replacement.Text = "^pATA n� quarenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e dois/"
            .Replacement.Text = "^pATA n� quarenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e tr�s/"
            .Replacement.Text = "^pATA n� quarenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e quatro/"
            .Replacement.Text = "^pATA n� quarenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e cinco/"
            .Replacement.Text = "^pATA n� quarenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e seis/"
            .Replacement.Text = "^pATA n� quarenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e sete/"
            .Replacement.Text = "^pATA n� quarenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e oito/"
            .Replacement.Text = "^pATA n� quarenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e nove/"
            .Replacement.Text = "^pATA n� quarenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta/"
            .Replacement.Text = "^pATA n� cinquenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e um/"
            .Replacement.Text = "^pATA n� cinquenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e dois/"
            .Replacement.Text = "^pATA n� cinquenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e tr�s/"
            .Replacement.Text = "^pATA n� cinquenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e quatro/"
            .Replacement.Text = "^pATA n� cinquenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e cinco/"
            .Replacement.Text = "^pATA n� cinquenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e seis/"
            .Replacement.Text = "^pATA n� cinquenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e sete/"
            .Replacement.Text = "^pATA n� cinquenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e oito/"
            .Replacement.Text = "^pATA n� cinquenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e nove/"
            .Replacement.Text = "^pATA n� cinquenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta/"
            .Replacement.Text = "^pATA n� sessenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e um/"
            .Replacement.Text = "^pATA n� sessenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e dois/"
            .Replacement.Text = "^pATA n� sessenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e tr�s/"
            .Replacement.Text = "^pATA n� sessenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e quatro/"
            .Replacement.Text = "^pATA n� sessenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e cinco/"
            .Replacement.Text = "^pATA n� sessenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e seis/"
            .Replacement.Text = "^pATA n� sessenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e sete/"
            .Replacement.Text = "^pATA n� sessenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e oito/"
            .Replacement.Text = "^pATA n� sessenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e nove/"
            .Replacement.Text = "^pATA n� sessenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta/"
            .Replacement.Text = "^pATA n� setenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e um/"
            .Replacement.Text = "^pATA n� setenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e dois/"
            .Replacement.Text = "^pATA n� setenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e tr�s/"
            .Replacement.Text = "^pATA n� setenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e quatro/"
            .Replacement.Text = "^pATA n� setenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e cinco/"
            .Replacement.Text = "^pATA n� setenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e seis/"
            .Replacement.Text = "^pATA n� setenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e sete/"
            .Replacement.Text = "^pATA n� setenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e oito/"
            .Replacement.Text = "^pATA n� setenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e nove/"
            .Replacement.Text = "^pATA n� setenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta/"
            .Replacement.Text = "^pATA n� oitenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e um/"
            .Replacement.Text = "^pATA n� oitenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e dois/"
            .Replacement.Text = "^pATA n� oitenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e tr�s/"
            .Replacement.Text = "^pATA n� oitenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e quatro/"
            .Replacement.Text = "^pATA n� oitenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e cinco/"
            .Replacement.Text = "^pATA n� oitenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e seis/"
            .Replacement.Text = "^pATA n� oitenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e sete/"
            .Replacement.Text = "^pATA n� oitenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e oito/"
            .Replacement.Text = "^pATA n� oitenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e nove/"
            .Replacement.Text = "^pATA n� oitenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa/"
            .Replacement.Text = "^pATA n� noventa/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e um/"
            .Replacement.Text = "^pATA n� noventa e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e dois/"
            .Replacement.Text = "^pATA n� noventa e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e tr�s/"
            .Replacement.Text = "^pATA n� noventa e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e quatro/"
            .Replacement.Text = "^pATA n� noventa e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e cinco/"
            .Replacement.Text = "^pATA n� noventa e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e seis/"
            .Replacement.Text = "^pATA n� noventa e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e sete/"
            .Replacement.Text = "^pATA n� noventa e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e oito/"
            .Replacement.Text = "^pATA n� noventa e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e nove/"
            .Replacement.Text = "^pATA n� noventa e nove/^p"
            .Forward = True
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



        ' QuebraT�tuloSe��es Macro
        '
        '
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "SESS�O PLEN�RIA ORDIN�RIA"
                .Replacement.Text = "^pSESS�O PLEN�RIA ORDIN�RIA^p"
                .Forward = True
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
                .Text = "PER�ODO DAS COMUNICA��ES"
                .Replacement.Text = "^pPER�ODO DAS COMUNICA��ES^p"
                .Forward = True
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
                .Text = "PER�ODO DAS COMUNICA��ES"
                .Replacement.Text = "^pPER�ODO DAS COMUNICA��ES^p"
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
                .Text = "ESPA�O DE LIDERAN�A"
                .Replacement.Text = "^pESPA�O DE LIDERAN�A^p"
                .Forward = True
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
                .Text = "ESPA�O DE LIDERAN�A"
                .Replacement.Text = "^pESPA�O DE LIDERAN�A^p"
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
                .Text = "ESPA�O DE LIDERAN�A"
                .Replacement.Text = "^pESPA�O DE LIDERAN�A^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With

            ' AtribuiEstiloSe��es Macro
            '
            '
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 1")
                With Selection.Find
                    .Text = "ATA n�"
                    .Replacement.Text = "ATA n�"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 2")
                With Selection.Find
                    .Text = "SESS�O PLEN�RIA ORDIN�RIA"
                    .Replacement.Text = "SESS�O PLEN�RIA ORDIN�RIA"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 2")
                With Selection.Find
                    .Text = "PER�ODO DAS COMUNICA��ES"
                    .Replacement.Text = "PER�ODO DAS COMUNICA��ES"
                    .Forward = True
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
                    .Text = "ESPA�O DE LIDERAN�A"
                    .Replacement.Text = "ESPA�O DE LIDERAN�A"
                    .Forward = True
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
                    .Text = "Vereador Jo�o Kaus"
                    .Replacement.Text = "^pVereador Jo�o Kaus^p"
                    .Forward = True
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
                    .Text = "Vereador Jo�o Chaves"
                    .Replacement.Text = "^pVereador Jo�o Chaves^p"
                    .Forward = True
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
                    .Text = "Vereador Jo�o Ricardo Vargas"
                    .Replacement.Text = "^pVereador Jo�o Ricardo Vargas^p"
                    .Forward = True
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
                    .Text = "Vereador Ov�dio Mayer"
                    .Replacement.Text = "^pVereador Ov�dio Mayer^p"
                    .Forward = True
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
                    .Text = "Vereador Vanderlei Ara�jo"
                    .Replacement.Text = "^pVereador Vanderlei Ara�jo^p"
                    .Forward = True
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
                    .Text = "Vereador Andr� Agne Domingues"
                    .Replacement.Text = "^pVereador Andr� Agne Domingues^p"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Jo�o Kaus"
                    .Replacement.Text = "Vereador Jo�o Kaus"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Jo�o Chaves"
                    .Replacement.Text = "Vereador Jo�o Chaves"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Jo�o Ricardo Vargas"
                    .Replacement.Text = "Vereador Jo�o Ricardo Vargas"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Ov�dio Mayer"
                    .Replacement.Text = "Vereador Ov�dio Mayer"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Vanderlei Ara�jo"
                    .Replacement.Text = "Vereador Vanderlei Ara�jo"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Andr� Agne Domingues"
                    .Replacement.Text = "Vereador Andr� Agne Domingues"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
        .Text = "C�MARA MUNICIPAL DE VEREADORES DE SANTA MARIA � RS" & vbTab _
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
    ' Altera algarismos de 01 a 99 do t�tulo das atas para n�meros por extenso para tratamento de atas no projeto Boca de Urna 2020.
    ' Adaptado para tr�s digitos no n�mero da ata, como as atas de 2017
    '
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ATA n� 001 /"
        .Replacement.Text = "ATA n� um/"
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
        .Text = "ATA n� 002 /"
        .Replacement.Text = "ATA n� dois/"
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
        .Text = "ATA n� 003 /"
        .Replacement.Text = "ATA n� tr�s/"
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
        .Text = "ATA n� 004 /"
        .Replacement.Text = "ATA n� quatro/"
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
        .Text = "ATA n� 005 /"
        .Replacement.Text = "ATA n� cinco/"
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
        .Text = "ATA n� 006 /"
        .Replacement.Text = "ATA n� seis/"
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
        .Text = "ATA n� 007 /"
        .Replacement.Text = "ATA n� sete/"
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
        .Text = "ATA n� 008 /"
        .Replacement.Text = "ATA n� oito/"
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
        .Text = "ATA n� 009 /"
        .Replacement.Text = "ATA n� nove/"
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
        .Text = "ATA n� 010 /"
        .Replacement.Text = "ATA n� dez/"
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
        .Text = "ATA n� 011 /"
        .Replacement.Text = "ATA n� onze/"
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
        .Text = "ATA n� 012 /"
        .Replacement.Text = "ATA n� doze/"
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
        .Text = "ATA n� 013 /"
        .Replacement.Text = "ATA n� treze/"
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
        .Text = "ATA n� 014 /"
        .Replacement.Text = "ATA n� quatorze/"
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
        .Text = "ATA n� 015 /"
        .Replacement.Text = "ATA n� quinze/"
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
        .Text = "ATA n� 016 /"
        .Replacement.Text = "ATA n� dezesseis/"
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
        .Text = "ATA n� 017 /"
        .Replacement.Text = "ATA n� dezessete/"
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
        .Text = "ATA n� 018 /"
        .Replacement.Text = "ATA n� dezoito/"
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
        .Text = "ATA n� 019 /"
        .Replacement.Text = "ATA n� dezenove/"
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
        .Text = "ATA n� 020 /"
        .Replacement.Text = "ATA n� vinte/"
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
        .Text = "ATA n� 021 /"
        .Replacement.Text = "ATA n� vinte e um/"
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
        .Text = "ATA n� 022 /"
        .Replacement.Text = "ATA n� vinte e dois/"
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
        .Text = "ATA n� 023 /"
        .Replacement.Text = "ATA n� vinte e tr�s/"
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
        .Text = "ATA n� 024 /"
        .Replacement.Text = "ATA n� vinte e quatro/"
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
        .Text = "ATA n� 025 /"
        .Replacement.Text = "ATA n� vinte e cinco/"
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
        .Text = "ATA n� 026 /"
        .Replacement.Text = "ATA n� vinte e seis/"
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
        .Text = "ATA n� 027 /"
        .Replacement.Text = "ATA n� vinte e sete/"
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
        .Text = "ATA n� 028 /"
        .Replacement.Text = "ATA n� vinte e oito/"
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
        .Text = "ATA n� 029 /"
        .Replacement.Text = "ATA n� vinte e nove/"
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
        .Text = "ATA n� 030 /"
        .Replacement.Text = "ATA n� trinta/"
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
        .Text = "ATA n� 031 /"
        .Replacement.Text = "ATA n� trinta e um/"
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
        .Text = "ATA n� 032 /"
        .Replacement.Text = "ATA n� trinta e dois/"
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
        .Text = "ATA n� 033 /"
        .Replacement.Text = "ATA n� trinta e tr�s/"
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
        .Text = "ATA n� 034 /"
        .Replacement.Text = "ATA n� trinta e quatro/"
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
        .Text = "ATA n� 035 /"
        .Replacement.Text = "ATA n� trinta e cinco/"
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
        .Text = "ATA n� 036 /"
        .Replacement.Text = "ATA n� trinta e seis/"
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
        .Text = "ATA n� 037 /"
        .Replacement.Text = "ATA n� trinta e sete/"
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
        .Text = "ATA n� 038 /"
        .Replacement.Text = "ATA n� trinta e oito/"
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
        .Text = "ATA n� 039 /"
        .Replacement.Text = "ATA n� trinta e nove/"
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
        .Text = "ATA n� 040 /"
        .Replacement.Text = "ATA n� quarenta/"
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
        .Text = "ATA n� 041 /"
        .Replacement.Text = "ATA n� quarenta e um/"
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
        .Text = "ATA n� 042 /"
        .Replacement.Text = "ATA n� quarenta e dois/"
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
        .Text = "ATA n� 043 /"
        .Replacement.Text = "ATA n� quarenta e tr�s/"
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
        .Text = "ATA n� 044 /"
        .Replacement.Text = "ATA n� quarenta e quatro/"
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
        .Text = "ATA n� 045 /"
        .Replacement.Text = "ATA n� quarenta e cinco/"
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
        .Text = "ATA n� 046 /"
        .Replacement.Text = "ATA n� quarenta e seis/"
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
        .Text = "ATA n� 047 /"
        .Replacement.Text = "ATA n� quarenta e sete/"
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
        .Text = "ATA n� 048 /"
        .Replacement.Text = "ATA n� quarenta e oito/"
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
        .Text = "ATA n� 049 /"
        .Replacement.Text = "ATA n� quarenta e nove/"
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
        .Text = "ATA n� 050 /"
        .Replacement.Text = "ATA n� cinquenta/"
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
        .Text = "ATA n� 051 /"
        .Replacement.Text = "ATA n� cinquenta e um/"
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
        .Text = "ATA n� 052 /"
        .Replacement.Text = "ATA n� cinquenta e dois/"
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
        .Text = "ATA n� 053 /"
        .Replacement.Text = "ATA n� cinquenta e tr�s/"
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
        .Text = "ATA n� 054 /"
        .Replacement.Text = "ATA n� cinquenta e quatro/"
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
        .Text = "ATA n� 055 /"
        .Replacement.Text = "ATA n� cinquenta e cinco/"
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
        .Text = "ATA n� 056 /"
        .Replacement.Text = "ATA n� cinquenta e seis/"
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
        .Text = "ATA n� 057 /"
        .Replacement.Text = "ATA n� cinquenta e sete/"
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
        .Text = "ATA n� 058 /"
        .Replacement.Text = "ATA n� cinquenta e oito/"
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
        .Text = "ATA n� 059 /"
        .Replacement.Text = "ATA n� cinquenta e nove/"
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
        .Text = "ATA n� 060 /"
        .Replacement.Text = "ATA n� sessenta/"
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
        .Text = "ATA n� 061 /"
        .Replacement.Text = "ATA n� sessenta e um/"
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
        .Text = "ATA n� 062 /"
        .Replacement.Text = "ATA n� sessenta e dois/"
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
        .Text = "ATA n� 063 /"
        .Replacement.Text = "ATA n� sessenta e tr�s/"
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
        .Text = "ATA n� 064 /"
        .Replacement.Text = "ATA n� sessenta e quatro/"
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
        .Text = "ATA n� 065 /"
        .Replacement.Text = "ATA n� sessenta e cinco/"
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
        .Text = "ATA n� 066 /"
        .Replacement.Text = "ATA n� sessenta e seis/"
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
        .Text = "ATA n� 067 /"
        .Replacement.Text = "ATA n� sessenta e sete/"
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
        .Text = "ATA n� 068 /"
        .Replacement.Text = "ATA n� sessenta e oito/"
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
        .Text = "ATA n� 069 /"
        .Replacement.Text = "ATA n� sessenta e nove/"
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
        .Text = "ATA n� 070 /"
        .Replacement.Text = "ATA n� setenta/"
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
        .Text = "ATA n� 071 /"
        .Replacement.Text = "ATA n� setenta e um/"
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
        .Text = "ATA n� 072 /"
        .Replacement.Text = "ATA n� setenta e dois/"
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
        .Text = "ATA n� 073 /"
        .Replacement.Text = "ATA n� setenta e tr�s/"
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
        .Text = "ATA n� 074 /"
        .Replacement.Text = "ATA n� setenta e quatro/"
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
        .Text = "ATA n� 075 /"
        .Replacement.Text = "ATA n� setenta e cinco/"
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
        .Text = "ATA n� 076 /"
        .Replacement.Text = "ATA n� setenta e seis/"
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
        .Text = "ATA n� 077 /"
        .Replacement.Text = "ATA n� setenta e sete/"
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
        .Text = "ATA n� 078 /"
        .Replacement.Text = "ATA n� setenta e oito/"
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
        .Text = "ATA n� 079 /"
        .Replacement.Text = "ATA n� setenta e nove/"
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
        .Text = "ATA n� 080 /"
        .Replacement.Text = "ATA n� oitenta/"
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
        .Text = "ATA n� 081 /"
        .Replacement.Text = "ATA n� oitenta e um/"
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
        .Text = "ATA n� 082 /"
        .Replacement.Text = "ATA n� oitenta e dois/"
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
        .Text = "ATA n� 083 /"
        .Replacement.Text = "ATA n� oitenta e tr�s/"
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
        .Text = "ATA n� 084 /"
        .Replacement.Text = "ATA n� oitenta e quatro/"
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
        .Text = "ATA n� 085 /"
        .Replacement.Text = "ATA n� oitenta e cinco/"
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
        .Text = "ATA n� 086 /"
        .Replacement.Text = "ATA n� oitenta e seis/"
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
        .Text = "ATA n� 087 /"
        .Replacement.Text = "ATA n� oitenta e sete/"
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
        .Text = "ATA n� 088 /"
        .Replacement.Text = "ATA n� oitenta e oito/"
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
        .Text = "ATA n� 089 /"
        .Replacement.Text = "ATA n� oitenta e nove/"
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
        .Text = "ATA n� 090 /"
        .Replacement.Text = "ATA n� noventa/"
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
        .Text = "ATA n� 091 /"
        .Replacement.Text = "ATA n� noventa e um/"
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
        .Text = "ATA n� 092 /"
        .Replacement.Text = "ATA n� noventa e dois/"
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
        .Text = "ATA n� 093 /"
        .Replacement.Text = "ATA n� noventa e tr�s/"
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
        .Text = "ATA n� 094 /"
        .Replacement.Text = "ATA n� noventa e quatro/"
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
        .Text = "ATA n� 095 /"
        .Replacement.Text = "ATA n� noventa e cinco/"
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
        .Text = "ATA n� 096 /"
        .Replacement.Text = "ATA n� noventa e seis/"
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
        .Text = "ATA n� 097 /"
        .Replacement.Text = "ATA n� noventa e sete/"
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
        .Text = "ATA n� 098 /"
        .Replacement.Text = "ATA n� noventa e oito/"
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
        .Text = "ATA n� 099 /"
        .Replacement.Text = "ATA n� noventa e nove/"
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

        ' QuebraT�tuloAtas Macro
        '
        '
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "ATA n� um/"
            .Replacement.Text = "^pATA n� um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dois/"
            .Replacement.Text = "^pATA n� dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� tr�s/"
            .Replacement.Text = "^pATA n� tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quatro/"
            .Replacement.Text = "^pATA n� quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinco/"
            .Replacement.Text = "^pATA n� cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� seis/"
            .Replacement.Text = "^pATA n� seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sete/"
            .Replacement.Text = "^pATA n� sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oito/"
            .Replacement.Text = "^pATA n� oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� nove/"
            .Replacement.Text = "^pATA n� nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dez/"
            .Replacement.Text = "^pATA n� dez/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� onze/"
            .Replacement.Text = "^pATA n� onze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� doze/"
            .Replacement.Text = "^pATA n� doze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� treze/"
            .Replacement.Text = "^pATA n� treze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quatorze/"
            .Replacement.Text = "^pATA n� quatorze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quinze/"
            .Replacement.Text = "^pATA n� quinze/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dezesseis/"
            .Replacement.Text = "^pATA n� dezesseis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dezessete/"
            .Replacement.Text = "^pATA n� dezessete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dezoito/"
            .Replacement.Text = "^pATA n� dezoito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� dezenove/"
            .Replacement.Text = "^pATA n� dezenove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte/"
            .Replacement.Text = "^pATA n� vinte/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e um/"
            .Replacement.Text = "^pATA n� vinte e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e dois/"
            .Replacement.Text = "^pATA n� vinte e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e tr�s/"
            .Replacement.Text = "^pATA n� vinte e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e quatro/"
            .Replacement.Text = "^pATA n� vinte e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e cinco/"
            .Replacement.Text = "^pATA n� vinte e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e seis/"
            .Replacement.Text = "^pATA n� vinte e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e sete/"
            .Replacement.Text = "^pATA n� vinte e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e oito/"
            .Replacement.Text = "^pATA n� vinte e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� vinte e nove/"
            .Replacement.Text = "^pATA n� vinte e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta/"
            .Replacement.Text = "^pATA n� trinta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e um/"
            .Replacement.Text = "^pATA n� trinta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e dois/"
            .Replacement.Text = "^pATA n� trinta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e tr�s/"
            .Replacement.Text = "^pATA n� trinta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e quatro/"
            .Replacement.Text = "^pATA n� trinta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e cinco/"
            .Replacement.Text = "^pATA n� trinta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e seis/"
            .Replacement.Text = "^pATA n� trinta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e sete/"
            .Replacement.Text = "^pATA n� trinta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e oito/"
            .Replacement.Text = "^pATA n� trinta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� trinta e nove/"
            .Replacement.Text = "^pATA n� trinta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta/"
            .Replacement.Text = "^pATA n� quarenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e um/"
            .Replacement.Text = "^pATA n� quarenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e dois/"
            .Replacement.Text = "^pATA n� quarenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e tr�s/"
            .Replacement.Text = "^pATA n� quarenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e quatro/"
            .Replacement.Text = "^pATA n� quarenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e cinco/"
            .Replacement.Text = "^pATA n� quarenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e seis/"
            .Replacement.Text = "^pATA n� quarenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e sete/"
            .Replacement.Text = "^pATA n� quarenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e oito/"
            .Replacement.Text = "^pATA n� quarenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� quarenta e nove/"
            .Replacement.Text = "^pATA n� quarenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta/"
            .Replacement.Text = "^pATA n� cinquenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e um/"
            .Replacement.Text = "^pATA n� cinquenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e dois/"
            .Replacement.Text = "^pATA n� cinquenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e tr�s/"
            .Replacement.Text = "^pATA n� cinquenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e quatro/"
            .Replacement.Text = "^pATA n� cinquenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e cinco/"
            .Replacement.Text = "^pATA n� cinquenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e seis/"
            .Replacement.Text = "^pATA n� cinquenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e sete/"
            .Replacement.Text = "^pATA n� cinquenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e oito/"
            .Replacement.Text = "^pATA n� cinquenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� cinquenta e nove/"
            .Replacement.Text = "^pATA n� cinquenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta/"
            .Replacement.Text = "^pATA n� sessenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e um/"
            .Replacement.Text = "^pATA n� sessenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e dois/"
            .Replacement.Text = "^pATA n� sessenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e tr�s/"
            .Replacement.Text = "^pATA n� sessenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e quatro/"
            .Replacement.Text = "^pATA n� sessenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e cinco/"
            .Replacement.Text = "^pATA n� sessenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e seis/"
            .Replacement.Text = "^pATA n� sessenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e sete/"
            .Replacement.Text = "^pATA n� sessenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e oito/"
            .Replacement.Text = "^pATA n� sessenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� sessenta e nove/"
            .Replacement.Text = "^pATA n� sessenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta/"
            .Replacement.Text = "^pATA n� setenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e um/"
            .Replacement.Text = "^pATA n� setenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e dois/"
            .Replacement.Text = "^pATA n� setenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e tr�s/"
            .Replacement.Text = "^pATA n� setenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e quatro/"
            .Replacement.Text = "^pATA n� setenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e cinco/"
            .Replacement.Text = "^pATA n� setenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e seis/"
            .Replacement.Text = "^pATA n� setenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e sete/"
            .Replacement.Text = "^pATA n� setenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e oito/"
            .Replacement.Text = "^pATA n� setenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� setenta e nove/"
            .Replacement.Text = "^pATA n� setenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta/"
            .Replacement.Text = "^pATA n� oitenta/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e um/"
            .Replacement.Text = "^pATA n� oitenta e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e dois/"
            .Replacement.Text = "^pATA n� oitenta e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e tr�s/"
            .Replacement.Text = "^pATA n� oitenta e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e quatro/"
            .Replacement.Text = "^pATA n� oitenta e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e cinco/"
            .Replacement.Text = "^pATA n� oitenta e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e seis/"
            .Replacement.Text = "^pATA n� oitenta e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e sete/"
            .Replacement.Text = "^pATA n� oitenta e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e oito/"
            .Replacement.Text = "^pATA n� oitenta e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� oitenta e nove/"
            .Replacement.Text = "^pATA n� oitenta e nove/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa/"
            .Replacement.Text = "^pATA n� noventa/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e um/"
            .Replacement.Text = "^pATA n� noventa e um/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e dois/"
            .Replacement.Text = "^pATA n� noventa e dois/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e tr�s/"
            .Replacement.Text = "^pATA n� noventa e tr�s/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e quatro/"
            .Replacement.Text = "^pATA n� noventa e quatro/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e cinco/"
            .Replacement.Text = "^pATA n� noventa e cinco/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e seis/"
            .Replacement.Text = "^pATA n� noventa e seis/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e sete/"
            .Replacement.Text = "^pATA n� noventa e sete/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e oito/"
            .Replacement.Text = "^pATA n� noventa e oito/^p"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "ATA n� noventa e nove/"
            .Replacement.Text = "^pATA n� noventa e nove/^p"
            .Forward = True
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




        ' QuebraT�tuloSe��es Macro
        ' Adaptado para as atas de 2017 que dizem comunica��o de lideran�a
        '
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = "SESS�O PLEN�RIA ORDIN�RIA"
                .Replacement.Text = "^pSESS�O PLEN�RIA ORDIN�RIA^p"
                .Forward = True
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
                .Text = "PER�ODO DAS COMUNICA��ES"
                .Replacement.Text = "^pPER�ODO DAS COMUNICA��ES^p"
                .Forward = True
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
                .Text = "PER�ODO DAS COMUNICA��ES"
                .Replacement.Text = "^pPER�ODO DAS COMUNICA��ES^p"
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
                .Text = "COMUNICA��O DE LIDERAN�A"
                .Replacement.Text = "^pCOMUNICA��O DE LIDERAN�A^p"
                .Forward = True
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
                .Text = "COMUNICA��O DE LIDERAN�A"
                .Replacement.Text = "^pCOMUNICA��O DE LIDERAN�A^p"
                .Forward = True
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
                .Text = "COMUNICA��O DE LIDERAN�A"
                .Replacement.Text = "^pCOMUNICA��O DE LIDERAN�A^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll

            ' AtribuiEstiloSe��es Macro
            '
            '
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 1")
                With Selection.Find
                    .Text = "ATA n�"
                    .Replacement.Text = "ATA n�"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 2")
                With Selection.Find
                    .Text = "SESS�O PLEN�RIA ORDIN�RIA"
                    .Replacement.Text = "SESS�O PLEN�RIA ORDIN�RIA"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 2")
                With Selection.Find
                    .Text = "PER�ODO DAS COMUNICA��ES"
                    .Replacement.Text = "PER�ODO DAS COMUNICA��ES"
                    .Forward = True
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
                    .Text = "ESPA�O DE LIDERAN�A"
                    .Replacement.Text = "ESPA�O DE LIDERAN�A"
                    .Forward = True
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
                    .Text = "COMUNICA��O DE LIDERAN�A"
                    .Replacement.Text = "COMUNICA��O DE LIDERAN�A"
                    .Forward = True
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
                ' Adicionada Lorena, Jo�o da Silva
                '
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Font.Bold = True
                With Selection.Find
                    .Text = "Vereador Jo�o da Silva Chaves"
                    .Replacement.Text = "^pVereadora Jo�o da Silva Chaves^p"
                    .Forward = True
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
                    .Text = "Vereador Jo�o Kaus"
                    .Replacement.Text = "^pVereador Jo�o Kaus^p"
                    .Forward = True
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
                    .Text = "Vereador Jo�o Chaves"
                    .Replacement.Text = "^pVereador Jo�o Chaves^p"
                    .Forward = True
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
                    .Text = "Vereador Jo�o Ricardo Vargas"
                    .Replacement.Text = "^pVereador Jo�o Ricardo Vargas^p"
                    .Forward = True
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
                    .Text = "Vereador Ov�dio Mayer"
                    .Replacement.Text = "^pVereador Ov�dio Mayer^p"
                    .Forward = True
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
                    .Text = "Vereador Vanderlei Ara�jo"
                    .Replacement.Text = "^pVereador Vanderlei Ara�jo^p"
                    .Forward = True
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
                    .Text = "Vereador Andr� Agne Domingues"
                    .Replacement.Text = "^pVereador Andr� Agne Domingues^p"
                    .Forward = True
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
                ' Adicionada Lorena e Jo�o da Silva
                '
                '
                Selection.Find.ClearFormatting
                Selection.Find.Font.Bold = True
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Jo�o da Silva Chaves"
                    .Replacement.Text = "Vereador Jo�o da Silva Chaves"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Jo�o Kaus"
                    .Replacement.Text = "Vereador Jo�o Kaus"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Jo�o Chaves"
                    .Replacement.Text = "Vereador Jo�o Chaves"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Jo�o Ricardo Vargas"
                    .Replacement.Text = "Vereador Jo�o Ricardo Vargas"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Ov�dio Mayer"
                    .Replacement.Text = "Vereador Ov�dio Mayer"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Vanderlei Ara�jo"
                    .Replacement.Text = "Vereador Vanderlei Ara�jo"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
                With Selection.Find
                    .Text = "Vereador Andr� Agne Domingues"
                    .Replacement.Text = "Vereador Andr� Agne Domingues"
                    .Forward = True
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
                Selection.Find.Replacement.Style = ActiveDocument.Styles("T�tulo 3")
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
        .Replacement.Text = "Jo�o Ricardo Vargas "
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
        .Replacement.Text = "Jo�o Ricardo Vargas "
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
        .Text = "Jorj�o "
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
        .Text = "Alem�o do G�s "
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
        .Text = "PER�ODO DA COMUNICA��O"
        .Replacement.Text = "PER�ODO DAS COMUNICA��ES"
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
        .Replacement.Text = "Vereador Ov�dio Mayer "
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
        .Replacement.Text = "Ov�dio Mayer "
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
        .Replacement.Text = "Ov�dio Mayer "
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
        .Text = "Vereadora Dr� Cida Brizola "
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
        .Replacement.Text = "Vereador Jo�o Ricardo Vargas "
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
        .Text = "Vereadora Prof� Luci Tia da Moto "
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
        .Text = "Vereadora Prof� Celita da Silva "
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
