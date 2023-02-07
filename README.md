# vba_excel
Sistema de cálculo de amostragem estatística em VBA
# Exposição do Template

Sub TelaCheia_On()
                
        ' Processa Macro sem evidenciar
        
Application.ScreenUpdating = False
        
        'Oculta todos os Menus (Ribbons)
        
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
 
        Application.DisplayFormulaBar = False
        ActiveWindow.DisplayHeadings = False
 
        With ActiveWindow
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayWorkbookTabs = False
            .DisplayHeadings = False
            .DisplayZeros = False
            .DisplayHeadings = False
            .DisplayGridlines = False
        End With
 
    End Sub
     Sub TelaCheia_Off()
        'Exibe todos os Menus (Ribbons)
        
Application.ScreenUpdating = False
        
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
 
        Application.DisplayFormulaBar = True
        ActiveWindow.DisplayHeadings = True
 
        With ActiveWindow
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayWorkbookTabs = True
            .DisplayHeadings = True
            .DisplayZeros = True
            .DisplayHeadings = True
            .DisplayGridlines = True
        End With
 
     End Sub

Private Sub Botao_Pense_nisso_Click()

' Aparecer a userform para definir a população

Mensagem_2_Defina_a_população.Show
End Sub
Private Sub ComboBox1_DropButtonClick()

' Lista para combobox

ComboBox1.List = Array("", "0", "1", "2", "3", "4", "5", "6", "7", "8", ">8")
End Sub
Private Sub ComboBox2_DropButtonClick()

' Lista para combobox

ComboBox2.List = Array("", "Transferências de Recursos a Terceiros", "Processos Licitatórios")
End Sub
Private Sub ComboBox3_DropButtonClick()

' Lista para combobox

ComboBox3.List = Array("", "Baixo", "Médio", "Alto")
End Sub
Private Sub ComboBox4_DropButtonClick()

' Lista para combobox

ComboBox4.List = Array("", "Sistemático", "Aleatório")
End Sub
Private Sub ComboBox5_DropButtonClick()

' Lista para combobox

ComboBox5.List = Array("", "0", "1", "2", "3", "4", "5", "6", "7", "8", ">8")
End Sub
Private Sub ComboBox6_DropButtonClick()

' Lista para combobox

ComboBox6.List = Array("", "Sim", "Não")
End Sub
Private Sub CommandButton1_Click()

' Aparecer mensagem de "Pense bem" userform

Mensagem_4.Show
End Sub
Private Sub CommandButton2_Click()

' Aparecer a mensagem 3 userform

Mensagem_3_Defina_a_exceção.Show
End Sub
Private Sub CommandButton3_Click()

' Aparacer a mensagem esclamação

Mensagem_esclamação.Show
End Sub
Private Sub TextBox1_Change()

' Caixa de texto

If (Me.TextBox1 <> "") Then Me.TextBox1 = Round(Me.TextBox1)
End Sub

Private Sub TextBox2_Change()

End Sub


