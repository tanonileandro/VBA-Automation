VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormImpNuevo 
   Caption         =   "EMBALAJES SRL"
   ClientHeight    =   8490.001
   ClientLeft      =   1080
   ClientTop       =   3465
   ClientWidth     =   8805.001
   OleObjectBlob   =   "FormImpNuevo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormImpNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' INICIALIZACION DEL FORMULARIO

Private Sub UserForm_Initialize()
    ' Enero
    TextEnero.Enabled = False
    TextEnero.BackColor = RGB(240, 240, 240)
    TextEnero.ForeColor = RGB(128, 128, 128)
    TextBox1.Enabled = False
    TextBox1.BackColor = RGB(240, 240, 240)
    TextBox1.ForeColor = RGB(128, 128, 128)
    Label11.Enabled = False
    Label11.ForeColor = RGB(128, 128, 128)
    
    ' Febrero
    TextFebrero.Enabled = False
    TextFebrero.BackColor = RGB(240, 240, 240)
    TextFebrero.ForeColor = RGB(128, 128, 128)
    TextBox2.Enabled = False
    TextBox2.BackColor = RGB(240, 240, 240)
    TextBox2.ForeColor = RGB(128, 128, 128)
    Label12.Enabled = False
    Label12.ForeColor = RGB(128, 128, 128)
    
    ' Marzo
    TextMarzo.Enabled = False
    TextMarzo.BackColor = RGB(240, 240, 240)
    TextMarzo.ForeColor = RGB(128, 128, 128)
    TextBox3.Enabled = False
    TextBox3.BackColor = RGB(240, 240, 240)
    TextBox3.ForeColor = RGB(128, 128, 128)
    Label13.Enabled = False
    Label13.ForeColor = RGB(128, 128, 128)
    
    ' Abril
    TextAbril.Enabled = False
    TextAbril.BackColor = RGB(240, 240, 240)
    TextAbril.ForeColor = RGB(128, 128, 128)
    TextBox4.Enabled = False
    TextBox4.BackColor = RGB(240, 240, 240)
    TextBox4.ForeColor = RGB(128, 128, 128)
    Label14.Enabled = False
    Label14.ForeColor = RGB(128, 128, 128)
    
    ' Mayo
    TextMayo.Enabled = False
    TextMayo.BackColor = RGB(240, 240, 240)
    TextMayo.ForeColor = RGB(128, 128, 128)
    TextBox5.Enabled = False
    TextBox5.BackColor = RGB(240, 240, 240)
    TextBox5.ForeColor = RGB(128, 128, 128)
    Label15.Enabled = False
    Label15.ForeColor = RGB(128, 128, 128)
    
    ' Junio
    TextJunio.Enabled = False
    TextJunio.BackColor = RGB(240, 240, 240)
    TextJunio.ForeColor = RGB(128, 128, 128)
    TextBox6.Enabled = False
    TextBox6.BackColor = RGB(240, 240, 240)
    TextBox6.ForeColor = RGB(128, 128, 128)
    Label16.Enabled = False
    Label16.ForeColor = RGB(128, 128, 128)
    
    ' Julio
    TextJulio.Enabled = False
    TextJulio.BackColor = RGB(240, 240, 240)
    TextJulio.ForeColor = RGB(128, 128, 128)
    TextBox7.Enabled = False
    TextBox7.BackColor = RGB(240, 240, 240)
    TextBox7.ForeColor = RGB(128, 128, 128)
    Label18.Enabled = False
    Label18.ForeColor = RGB(128, 128, 128)
    
    ' Agosto
    TextAgosto.Enabled = False
    TextAgosto.BackColor = RGB(240, 240, 240)
    TextAgosto.ForeColor = RGB(128, 128, 128)
    TextBox8.Enabled = False
    TextBox8.BackColor = RGB(240, 240, 240)
    TextBox8.ForeColor = RGB(128, 128, 128)
    Label19.Enabled = False
    Label19.ForeColor = RGB(128, 128, 128)
    
    ' Septiembre
    TextSeptiembre.Enabled = False
    TextSeptiembre.BackColor = RGB(240, 240, 240)
    TextSeptiembre.ForeColor = RGB(128, 128, 128)
    TextBox9.Enabled = False
    TextBox9.BackColor = RGB(240, 240, 240)
    TextBox9.ForeColor = RGB(128, 128, 128)
    Label20.Enabled = False
    Label20.ForeColor = RGB(128, 128, 128)
    
    ' Octubre
    TextOctubre.Enabled = False
    TextOctubre.BackColor = RGB(240, 240, 240)
    TextOctubre.ForeColor = RGB(128, 128, 128)
    TextBox10.Enabled = False
    TextBox10.BackColor = RGB(240, 240, 240)
    TextBox10.ForeColor = RGB(128, 128, 128)
    Label21.Enabled = False
    Label21.ForeColor = RGB(128, 128, 128)
    
    ' Noviembre
    TextNoviembre.Enabled = False
    TextNoviembre.BackColor = RGB(240, 240, 240)
    TextNoviembre.ForeColor = RGB(128, 128, 128)
    TextBox11.Enabled = False
    TextBox11.BackColor = RGB(240, 240, 240)
    TextBox11.ForeColor = RGB(128, 128, 128)
    Label22.Enabled = False
    Label22.ForeColor = RGB(128, 128, 128)
    
    ' Diciembre
    TextDiciembre.Enabled = False
    TextDiciembre.BackColor = RGB(240, 240, 240)
    TextDiciembre.ForeColor = RGB(128, 128, 128)
    TextBox12.Enabled = False
    TextBox12.BackColor = RGB(240, 240, 240)
    TextBox12.ForeColor = RGB(128, 128, 128)
    Label23.Enabled = False
    Label23.ForeColor = RGB(128, 128, 128)
    
    ' Tamaño del formulario
    Me.Width = 455
    Me.Height = 366
    
    ' Posición del formulario al centro de la pantalla principal de Excel
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
End Sub

' CONFIGURACION DE CHECKBOX - ACTIVO E INACTIVO

Private Sub CheckBoxEne_Change()
    ' Habilitar o deshabilitar el cuadro de texto
    TextEnero.Enabled = CheckBoxEne.Value
    TextBox1.Enabled = CheckBoxEne.Value
    If CheckBoxEne.Value Then
        TextEnero.BackColor = RGB(255, 255, 255)  ' Color de fondo blanco cuando está activo
        TextEnero.ForeColor = RGB(0, 0, 0)        ' Color de fuente negro cuando está activo
        TextBox1.BackColor = RGB(255, 255, 255)
        TextBox1.ForeColor = RGB(0, 0, 0)
    Else
        TextEnero.BackColor = RGB(240, 240, 240)  ' Color de fondo gris claro cuando está inactivo
        TextEnero.ForeColor = RGB(128, 128, 128)  ' Color de fuente gris oscuro cuando está inactivo
        TextBox1.BackColor = RGB(240, 240, 240)
        TextBox1.ForeColor = RGB(128, 128, 128)
    End If
    
    ' Habilitar o deshabilitar la etiqueta asociada
    Label11.Enabled = CheckBoxEne.Value
    If CheckBoxEne.Value Then
        Label11.ForeColor = RGB(0, 0, 0)  ' Color de fuente negro cuando está activo
    Else
        Label11.ForeColor = RGB(128, 128, 128)  ' Color de fuente gris oscuro cuando está inactivo
    End If
End Sub

Private Sub CheckBoxFeb_Change()
    TextFebrero.Enabled = CheckBoxFeb.Value
    TextBox2.Enabled = CheckBoxFeb.Value
    If CheckBoxFeb.Value Then
        TextFebrero.BackColor = RGB(255, 255, 255)
        TextFebrero.ForeColor = RGB(0, 0, 0)
        TextBox2.BackColor = RGB(255, 255, 255)
        TextBox2.ForeColor = RGB(0, 0, 0)
    Else
        TextFebrero.BackColor = RGB(240, 240, 240)
        TextFebrero.ForeColor = RGB(128, 128, 128)
        TextBox2.BackColor = RGB(240, 240, 240)
        TextBox2.ForeColor = RGB(128, 128, 128)
    End If
    
    Label12.Enabled = CheckBoxFeb.Value
    If CheckBoxFeb.Value Then
        Label12.ForeColor = RGB(0, 0, 0)
    Else
        Label12.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxMar_Change()
    TextMarzo.Enabled = CheckBoxMar.Value
    TextBox3.Enabled = CheckBoxMar.Value
    If CheckBoxMar.Value Then
        TextMarzo.BackColor = RGB(255, 255, 255)
        TextMarzo.ForeColor = RGB(0, 0, 0)
        TextBox3.BackColor = RGB(255, 255, 255)
        TextBox3.ForeColor = RGB(0, 0, 0)
    Else
        TextMarzo.BackColor = RGB(240, 240, 240)
        TextMarzo.ForeColor = RGB(128, 128, 128)
        TextBox3.BackColor = RGB(240, 240, 240)
        TextBox3.ForeColor = RGB(128, 128, 128)
    End If
    
    Label13.Enabled = CheckBoxMar.Value
    If CheckBoxMar.Value Then
        Label13.ForeColor = RGB(0, 0, 0)
    Else
        Label13.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxAbr_Change()
    TextAbril.Enabled = CheckBoxAbr.Value
    TextBox4.Enabled = CheckBoxAbr.Value
    If CheckBoxAbr.Value Then
        TextAbril.BackColor = RGB(255, 255, 255)
        TextAbril.ForeColor = RGB(0, 0, 0)
        TextBox4.BackColor = RGB(255, 255, 255)
        TextBox4.ForeColor = RGB(0, 0, 0)
    Else
        TextAbril.BackColor = RGB(240, 240, 240)
        TextAbril.ForeColor = RGB(128, 128, 128)
        TextBox4.BackColor = RGB(240, 240, 240)
        TextBox4.ForeColor = RGB(128, 128, 128)
    End If
    
    Label14.Enabled = CheckBoxAbr.Value
    If CheckBoxAbr.Value Then
        Label14.ForeColor = RGB(0, 0, 0)
    Else
        Label14.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxMay_Change()
    TextMayo.Enabled = CheckBoxMay.Value
    TextBox5.Enabled = CheckBoxMay.Value
    If CheckBoxMay.Value Then
        TextMayo.BackColor = RGB(255, 255, 255)
        TextMayo.ForeColor = RGB(0, 0, 0)
        TextBox5.BackColor = RGB(255, 255, 255)
        TextBox5.ForeColor = RGB(0, 0, 0)
    Else
        TextMayo.BackColor = RGB(240, 240, 240)
        TextMayo.ForeColor = RGB(128, 128, 128)
        TextBox5.BackColor = RGB(240, 240, 240)
        TextBox5.ForeColor = RGB(128, 128, 128)
    End If
    
    Label15.Enabled = CheckBoxMay.Value
    If CheckBoxMay.Value Then
        Label15.ForeColor = RGB(0, 0, 0)
    Else
        Label15.ForeColor = RGB(128, 128, 128)
End Sub

Private Sub CheckBoxJun_Change()
    TextJunio.Enabled = CheckBoxJun.Value
    TextBox6.Enabled = CheckBoxJun.Value
    If CheckBoxJun.Value Then
        TextJunio.BackColor = RGB(255, 255, 255)
        TextJunio.ForeColor = RGB(0, 0, 0)
        TextBox6.BackColor = RGB(255, 255, 255)
        TextBox6.ForeColor = RGB(0, 0, 0)
    Else
        TextJunio.BackColor = RGB(240, 240, 240)
        TextJunio.ForeColor = RGB(128, 128, 128)
        TextBox6.BackColor = RGB(240, 240, 240)
        TextBox6.ForeColor = RGB(128, 128, 128)
    End If
    
    Label16.Enabled = CheckBoxJun.Value
    If CheckBoxJun.Value Then
        Label16.ForeColor = RGB(0, 0, 0)
    Else
        Label16.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxJul_Change()
    TextJulio.Enabled = CheckBoxJul.Value
    TextBox7.Enabled = CheckBoxJul.Value
    If CheckBoxJul.Value Then
        TextJulio.BackColor = RGB(255, 255, 255)
        TextJulio.ForeColor = RGB(0, 0, 0)
        TextBox7.BackColor = RGB(255, 255, 255)
        TextBox7.ForeColor = RGB(0, 0, 0)
    Else
        TextJulio.BackColor = RGB(240, 240, 240)
        TextJulio.ForeColor = RGB(128, 128, 128)
        TextBox7.BackColor = RGB(240, 240, 240)
        TextBox7.ForeColor = RGB(128, 128, 128)
    End If
    
    Label18.Enabled = CheckBoxJul.Value
    If CheckBoxJul.Value Then
        Label18.ForeColor = RGB(0, 0, 0)
    Else
        Label18.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxAgo_Change()
    TextAgosto.Enabled = CheckBoxAgo.Value
    TextBox8.Enabled = CheckBoxAgo.Value
    If CheckBoxAgo.Value Then
        TextAgosto.BackColor = RGB(255, 255, 255)
        TextAgosto.ForeColor = RGB(0, 0, 0)
        TextBox8.BackColor = RGB(255, 255, 255)
        TextBox8.ForeColor = RGB(0, 0, 0)
    Else
        TextAgosto.BackColor = RGB(240, 240, 240)
        TextAgosto.ForeColor = RGB(128, 128, 128)
        TextBox8.BackColor = RGB(240, 240, 240)
        TextBox8.ForeColor = RGB(128, 128, 128)
    End If
    
    Label19.Enabled = CheckBoxAgo.Value
    If CheckBoxAgo.Value Then
        Label19.ForeColor = RGB(0, 0, 0)
    Else
        Label19.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxSep_Change()
    TextSeptiembre.Enabled = CheckBoxSep.Value
    TextBox9.Enabled = CheckBoxSep.Value
    If CheckBoxSep.Value Then
        TextSeptiembre.BackColor = RGB(255, 255, 255)
        TextSeptiembre.ForeColor = RGB(0, 0, 0)
        TextBox9.BackColor = RGB(255, 255, 255)
        TextBox9.ForeColor = RGB(0, 0, 0)
    Else
        TextSeptiembre.BackColor = RGB(240, 240, 240)
        TextSeptiembre.ForeColor = RGB(128, 128, 128)
        TextBox9.BackColor = RGB(240, 240, 240)
        TextBox9.ForeColor = RGB(128, 128, 128)
    End If

    Label20.Enabled = CheckBoxSep.Value
    If CheckBoxSep.Value Then
        Label20.ForeColor = RGB(0, 0, 0)
    Else
        Label20.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxOct_Change()
    TextOctubre.Enabled = CheckBoxOct.Value
    TextBox10.Enabled = CheckBoxOct.Value
    If CheckBoxOct.Value Then
        TextOctubre.BackColor = RGB(255, 255, 255)
        TextOctubre.ForeColor = RGB(0, 0, 0)
        TextBox10.BackColor = RGB(255, 255, 255)
        TextBox10.ForeColor = RGB(0, 0, 0)
    Else
        TextOctubre.BackColor = RGB(240, 240, 240)
        TextOctubre.ForeColor = RGB(128, 128, 128)
        TextBox10.BackColor = RGB(240, 240, 240)
        TextBox10.ForeColor = RGB(128, 128, 128)
    End If
    
    Label21.Enabled = CheckBoxOct.Value
    If CheckBoxOct.Value Then
        Label21.ForeColor = RGB(0, 0, 0)
    Else
        Label21.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxNov_Change()

    TextNoviembre.Enabled = CheckBoxNov.Value
    TextBox11.Enabled = CheckBoxNov.Value
    If CheckBoxNov.Value Then
        TextNoviembre.BackColor = RGB(255, 255, 255)
        TextNoviembre.ForeColor = RGB(0, 0, 0)
        TextBox11.BackColor = RGB(255, 255, 255)
        TextBox11.ForeColor = RGB(0, 0, 0)
    Else
        TextNoviembre.BackColor = RGB(240, 240, 240)
        TextNoviembre.ForeColor = RGB(128, 128, 128)
        TextBox11.BackColor = RGB(240, 240, 240)
        TextBox11.ForeColor = RGB(128, 128, 128)
    End If
    
    Label22.Enabled = CheckBoxNov.Value
    If CheckBoxNov.Value Then
        Label22.ForeColor = RGB(0, 0, 0)
    Else
        Label22.ForeColor = RGB(128, 128, 128)
    End If
End Sub

Private Sub CheckBoxDic_Change()
    TextDiciembre.Enabled = CheckBoxDic.Value
    TextBox12.Enabled = CheckBoxDic.Value
    If CheckBoxDic.Value Then
        TextDiciembre.BackColor = RGB(255, 255, 255)
        TextDiciembre.ForeColor = RGB(0, 0, 0)
        TextBox12.BackColor = RGB(255, 255, 255)
        TextBox12.ForeColor = RGB(0, 0, 0)
    Else
        TextDiciembre.BackColor = RGB(240, 240, 240)
        TextDiciembre.ForeColor = RGB(128, 128, 128)
        TextBox12.BackColor = RGB(240, 240, 240)
        TextBox12.ForeColor = RGB(128, 128, 128)
    End If
    
    Label23.Enabled = CheckBoxDic.Value
    If CheckBoxDic.Value Then
        Label23.ForeColor = RGB(0, 0, 0)
    Else
        Label23.ForeColor = RGB(128, 128, 128)
    End If
End Sub
Private Sub CheckBoxTodos_Click()
    Dim isChecked As Boolean
    isChecked = CheckBoxTodos.Value

    TextEnero.Enabled = isChecked
    TextFebrero.Enabled = isChecked
    TextMarzo.Enabled = isChecked
    TextAbril.Enabled = isChecked
    TextMayo.Enabled = isChecked
    TextJunio.Enabled = isChecked
    TextJulio.Enabled = isChecked
    TextAgosto.Enabled = isChecked
    TextSeptiembre.Enabled = isChecked
    TextOctubre.Enabled = isChecked
    TextNoviembre.Enabled = isChecked
    TextDiciembre.Enabled = isChecked
    
    TextBox1.Enabled = isChecked
    TextBox2.Enabled = isChecked
    TextBox3.Enabled = isChecked
    TextBox4.Enabled = isChecked
    TextBox5.Enabled = isChecked
    TextBox6.Enabled = isChecked
    TextBox7.Enabled = isChecked
    TextBox8.Enabled = isChecked
    TextBox9.Enabled = isChecked
    TextBox10.Enabled = isChecked
    TextBox11.Enabled = isChecked
    TextBox12.Enabled = isChecked
    
    ' Configurar el color de fondo y fuente de los textboxes según su estado
    If isChecked Then
        TextEnero.BackColor = RGB(255, 255, 255)  ' Color de fondo blanco cuando está activo
        TextFebrero.BackColor = RGB(255, 255, 255)
        TextMarzo.BackColor = RGB(255, 255, 255)
        TextAbril.BackColor = RGB(255, 255, 255)
        TextMayo.BackColor = RGB(255, 255, 255)
        TextJunio.BackColor = RGB(255, 255, 255)
        TextJulio.BackColor = RGB(255, 255, 255)
        TextAgosto.BackColor = RGB(255, 255, 255)
        TextSeptiembre.BackColor = RGB(255, 255, 255)
        TextOctubre.BackColor = RGB(255, 255, 255)
        TextNoviembre.BackColor = RGB(255, 255, 255)
        TextDiciembre.BackColor = RGB(255, 255, 255)
        
        TextBox1.BackColor = RGB(255, 255, 255)  ' Color de fondo blanco cuando está activo
        TextBox2.BackColor = RGB(255, 255, 255)
        TextBox3.BackColor = RGB(255, 255, 255)
        TextBox4.BackColor = RGB(255, 255, 255)
        TextBox5.BackColor = RGB(255, 255, 255)
        TextBox6.BackColor = RGB(255, 255, 255)
        TextBox7.BackColor = RGB(255, 255, 255)
        TextBox8.BackColor = RGB(255, 255, 255)
        TextBox9.BackColor = RGB(255, 255, 255)
        TextBox10.BackColor = RGB(255, 255, 255)
        TextBox11.BackColor = RGB(255, 255, 255)
        TextBox12.BackColor = RGB(255, 255, 255)
        
        TextEnero.ForeColor = RGB(0, 0, 0)  ' Color de fuente negro cuando está activo
        TextFebrero.ForeColor = RGB(0, 0, 0)
        TextMarzo.ForeColor = RGB(0, 0, 0)
        TextAbril.ForeColor = RGB(0, 0, 0)
        TextMayo.ForeColor = RGB(0, 0, 0)
        TextJunio.ForeColor = RGB(0, 0, 0)
        TextJulio.ForeColor = RGB(0, 0, 0)
        TextAgosto.ForeColor = RGB(0, 0, 0)
        TextSeptiembre.ForeColor = RGB(0, 0, 0)
        TextOctubre.ForeColor = RGB(0, 0, 0)
        TextNoviembre.ForeColor = RGB(0, 0, 0)
        TextDiciembre.ForeColor = RGB(0, 0, 0)
        
        TextBox1.ForeColor = RGB(0, 0, 0)        ' Color de fuente negro cuando está activo
        TextBox2.ForeColor = RGB(0, 0, 0)
        TextBox3.ForeColor = RGB(0, 0, 0)
        TextBox4.ForeColor = RGB(0, 0, 0)
        TextBox5.ForeColor = RGB(0, 0, 0)
        TextBox6.ForeColor = RGB(0, 0, 0)
        TextBox7.ForeColor = RGB(0, 0, 0)
        TextBox8.ForeColor = RGB(0, 0, 0)
        TextBox9.ForeColor = RGB(0, 0, 0)
        TextBox10.ForeColor = RGB(0, 0, 0)
        TextBox11.ForeColor = RGB(0, 0, 0)
        TextBox12.ForeColor = RGB(0, 0, 0)
    Else
        TextEnero.BackColor = RGB(240, 240, 240)  ' Color de fondo gris claro cuando está inactivo
        TextFebrero.BackColor = RGB(240, 240, 240)
        TextMarzo.BackColor = RGB(240, 240, 240)
        TextAbril.BackColor = RGB(240, 240, 240)
        TextMayo.BackColor = RGB(240, 240, 240)
        TextJunio.BackColor = RGB(240, 240, 240)
        TextJulio.BackColor = RGB(240, 240, 240)
        TextAgosto.BackColor = RGB(240, 240, 240)
        TextSeptiembre.BackColor = RGB(240, 240, 240)
        TextOctubre.BackColor = RGB(240, 240, 240)
        TextNoviembre.BackColor = RGB(240, 240, 240)
        TextDiciembre.BackColor = RGB(240, 240, 240)
        
        TextBox1.BackColor = RGB(240, 240, 240)  ' Color de fondo gris claro cuando está inactivo
        TextBox2.BackColor = RGB(240, 240, 240)
        TextBox3.BackColor = RGB(240, 240, 240)
        TextBox4.BackColor = RGB(240, 240, 240)
        TextBox5.BackColor = RGB(240, 240, 240)
        TextBox6.BackColor = RGB(240, 240, 240)
        TextBox7.BackColor = RGB(240, 240, 240)
        TextBox8.BackColor = RGB(240, 240, 240)
        TextBox9.BackColor = RGB(240, 240, 240)
        TextBox10.BackColor = RGB(240, 240, 240)
        TextBox11.BackColor = RGB(240, 240, 240)
        TextBox12.BackColor = RGB(240, 240, 240)
        
        TextEnero.ForeColor = RGB(128, 128, 128)  ' Color de fuente gris oscuro cuando está inactivo
        TextFebrero.ForeColor = RGB(128, 128, 128)
        TextMarzo.ForeColor = RGB(128, 128, 128)
        TextAbril.ForeColor = RGB(128, 128, 128)
        TextMayo.ForeColor = RGB(128, 128, 128)
        TextJunio.ForeColor = RGB(128, 128, 128)
        TextJulio.ForeColor = RGB(128, 128, 128)
        TextAgosto.ForeColor = RGB(128, 128, 128)
        TextSeptiembre.ForeColor = RGB(128, 128, 128)
        TextOctubre.ForeColor = RGB(128, 128, 128)
        TextNoviembre.ForeColor = RGB(128, 128, 128)
        TextDiciembre.ForeColor = RGB(128, 128, 128)
        
        TextBox1.ForeColor = RGB(128, 128, 128)  ' Color de fuente gris oscuro cuando está inactivo
        TextBox2.ForeColor = RGB(128, 128, 128)
        TextBox3.ForeColor = RGB(128, 128, 128)
        TextBox4.ForeColor = RGB(128, 128, 128)
        TextBox5.ForeColor = RGB(128, 128, 128)
        TextBox6.ForeColor = RGB(128, 128, 128)
        TextBox7.ForeColor = RGB(128, 128, 128)
        TextBox8.ForeColor = RGB(128, 128, 128)
        TextBox9.ForeColor = RGB(128, 128, 128)
        TextBox10.ForeColor = RGB(128, 128, 128)
        TextBox11.ForeColor = RGB(128, 128, 128)
        TextBox12.ForeColor = RGB(128, 128, 128)
    End If
    
    If isChecked Then
        Label11.ForeColor = RGB(0, 0, 0)  ' Color de fuente negro cuando está activo
        Label12.ForeColor = RGB(0, 0, 0)
        Label13.ForeColor = RGB(0, 0, 0)
        Label14.ForeColor = RGB(0, 0, 0)
        Label15.ForeColor = RGB(0, 0, 0)
        Label16.ForeColor = RGB(0, 0, 0)
        Label18.ForeColor = RGB(0, 0, 0)
        Label19.ForeColor = RGB(0, 0, 0)
        Label20.ForeColor = RGB(0, 0, 0)
        Label21.ForeColor = RGB(0, 0, 0)
        Label22.ForeColor = RGB(0, 0, 0)
        Label23.ForeColor = RGB(0, 0, 0)
    Else
        Label11.ForeColor = RGB(128, 128, 128)  ' Color de fuente gris oscuro cuando está inactivo
        Label12.ForeColor = RGB(128, 128, 128)
        Label13.ForeColor = RGB(128, 128, 128)
        Label14.ForeColor = RGB(128, 128, 128)
        Label15.ForeColor = RGB(128, 128, 128)
        Label16.ForeColor = RGB(128, 128, 128)
        Label18.ForeColor = RGB(128, 128, 128)
        Label19.ForeColor = RGB(128, 128, 128)
        Label20.ForeColor = RGB(128, 128, 128)
        Label21.ForeColor = RGB(128, 128, 128)
        Label22.ForeColor = RGB(128, 128, 128)
        Label23.ForeColor = RGB(128, 128, 128)
    End If
End Sub

' BOTON CARGAR FORMULARIO

Private Sub Cargar_Click()
    Dim ws As Worksheet
    Dim i As Integer
    Dim mes As String
    Dim fila As Long
    
    Set ws = ThisWorkbook.Sheets("ImpAnual")
    fila = ws.ListObjects("Tabla3").ListRows.Count + 1 ' Próxima fila disponible en la tabla
    
    ' Validar que al menos un checkbox esté seleccionado
    If Not CheckBoxEne.Value And Not CheckBoxFeb.Value And Not CheckBoxMar.Value And _
       Not CheckBoxAbr.Value And Not CheckBoxMay.Value And Not CheckBoxJun.Value And _
       Not CheckBoxJul.Value And Not CheckBoxAgo.Value And Not CheckBoxSep.Value And _
       Not CheckBoxOct.Value And Not CheckBoxNov.Value And Not CheckBoxDic.Value And _
       Not CheckBoxTodos.Value Then
        
        MsgBox "Debe seleccionar al menos un mes.", vbExclamation
        Exit Sub
    End If
    
    ' Validar campos obligatorios
    If TextBoxCUIT.Text = "" Or TextBoxTipoServicio.Text = "" Or TextBoxDetalle.Text = "" Then
        MsgBox "Los campos CUIT, Tipo de Servicio y Detalle son obligatorios.", vbExclamation
        Exit Sub
    End If
    
    ' Insertar datos según los checkboxes seleccionados
    If CheckBoxTodos.Value Then
        ' Si se selecciona Todos, se recorren todos los meses y se inserta una fila por cada uno
        For i = 1 To 12
            mes = Format(DateSerial(Year(Date), i, 1), "Mmm") ' Obtener el nombre del mes en tres letras
            InsertarFila ws, fila, mes
            fila = fila + 1
        Next i
    Else
        ' Si no se selecciona Todos, se inserta una fila por cada checkbox seleccionado
        If CheckBoxEne.Value Then InsertarFila ws, fila, "ene"
        If CheckBoxFeb.Value Then InsertarFila ws, fila, "feb"
        If CheckBoxMar.Value Then InsertarFila ws, fila, "mar"
        If CheckBoxAbr.Value Then InsertarFila ws, fila, "abr"
        If CheckBoxMay.Value Then InsertarFila ws, fila, "may"
        If CheckBoxJun.Value Then InsertarFila ws, fila, "jun"
        If CheckBoxJul.Value Then InsertarFila ws, fila, "jul"
        If CheckBoxAgo.Value Then InsertarFila ws, fila, "ago"
        If CheckBoxSep.Value Then InsertarFila ws, fila, "sep"
        If CheckBoxOct.Value Then InsertarFila ws, fila, "oct"
        If CheckBoxNov.Value Then InsertarFila ws, fila, "nov"
        If CheckBoxDic.Value Then InsertarFila ws, fila, "dic"
    End If
    
    ' Limpiar campos después de insertar los datos
    LimpiarCampos
    
    ' Informar al usuario que los datos han sido cargados correctamente
    MsgBox "Los datos han sido cargados exitosamente.", vbInformation
End Sub

' FUNCION LLAMADA PARA INSERTAR FILAS

Private Sub InsertarFila(ws As Worksheet, fila As Long, mes As String)
    
    With ws.ListObjects("Tabla3").ListRows.Add(fila).Range
        .Cells(1, 1).Value = mes ' Columna A: Mes
        .Cells(1, 3).Value = TextBoxCUIT.Text ' Columna C: CUIT
        .Cells(1, 4).Value = TextBoxTipoServicio.Text ' Columna D: Tipo de Servicio
        .Cells(1, 5).Value = TextBoxDetalle.Text ' Columna E: Detalle
        .Cells(1, 6).Value = TextBoxDireccion.Text ' Columna F: Dirección
        .Cells(1, 7).Value = TextBoxCuenta.Text ' Columna G: Cuenta
        .Cells(1, 8).Value = TextBoxNroCuenta.Text ' Columna H: NroCuenta
        .Cells(1, 9).Value = TextBoxNroIdentificacion.Text ' Columna I: NroIdentificacion
        
        ' Insertar los valores de los campos de texto adicionales según el mes seleccionado o todos los meses
        If mes = "ene" Then
            .Cells(1, 13).Value = IIf(TextEnero.Enabled, "'" & TextEnero.Text, "")
            .Cells(1, 11).Value = IIf(TextBox1.Enabled, "'" & TextBox1.Text, "")
        End If
        If mes = "feb" Then
            .Cells(1, 13).Value = IIf(TextFebrero.Enabled, "'" & TextFebrero.Text, "")
            .Cells(1, 11).Value = IIf(TextBox2.Enabled, "'" & TextBox2.Text, "")
        End If
        If mes = "mar" Then
            .Cells(1, 13).Value = IIf(TextMarzo.Enabled, "'" & TextMarzo.Text, "")
            .Cells(1, 11).Value = IIf(TextBox3.Enabled, "'" & TextBox3.Text, "")
        End If
        If mes = "abr" Then
            .Cells(1, 13).Value = IIf(TextAbril.Enabled, "'" & TextAbril.Text, "")
            .Cells(1, 11).Value = IIf(TextBox4.Enabled, "'" & TextBox4.Text, "")
        End If
        If mes = "may" Then
            .Cells(1, 13).Value = IIf(TextMayo.Enabled, "'" & TextMayo.Text, "")
            .Cells(1, 11).Value = IIf(TextBox5.Enabled, "'" & TextBox5.Text, "")
        End If
        If mes = "jun" Then
            .Cells(1, 13).Value = IIf(TextJunio.Enabled, "'" & TextJunio.Text, "")
            .Cells(1, 11).Value = IIf(TextBox6.Enabled, "'" & TextBox6.Text, "")
        End If
        If mes = "jul" Then
            .Cells(1, 13).Value = IIf(TextJulio.Enabled, "'" & TextJulio.Text, "")
            .Cells(1, 11).Value = IIf(TextBox7.Enabled, "'" & TextBox7.Text, "")
        End If
        If mes = "ago" Then
            .Cells(1, 13).Value = IIf(TextAgosto.Enabled, "'" & TextAgosto.Text, "")
            .Cells(1, 11).Value = IIf(TextBox8.Enabled, "'" & TextBox8.Text, "")
        End If
        If mes = "sep" Then
            .Cells(1, 13).Value = IIf(TextSeptiembre.Enabled, "'" & TextSeptiembre.Text, "")
            .Cells(1, 11).Value = IIf(TextBox9.Enabled, "'" & TextBox9.Text, "")
        End If
        If mes = "oct" Then
            .Cells(1, 13).Value = IIf(TextOctubre.Enabled, "'" & TextOctubre.Text, "")
            .Cells(1, 11).Value = IIf(TextBox10.Enabled, "'" & TextBox10.Text, "")
        End If
        If mes = "nov" Then
            .Cells(1, 13).Value = IIf(TextNoviembre.Enabled, "'" & TextNoviembre.Text, "")
            .Cells(1, 11).Value = IIf(TextBox11.Enabled, "'" & TextBox11.Text, "")
        End If
        If mes = "dic" Then
            .Cells(1, 13).Value = IIf(TextDiciembre.Enabled, "'" & TextDiciembre.Text, "")
            .Cells(1, 11).Value = IIf(TextBox12.Enabled, "'" & TextBox12.Text, "")
        End If
    End With
End Sub

' LIMPIAR CAMPOS LUEGO DE APRETAR EL BOTON DE CARGAR

Private Sub LimpiarCampos()
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is MSForms.CheckBox Then
            ctrl.Value = False
        End If
    Next ctrl
End Sub
Private Sub CloseForm_Click()
    FormImpNuevo.Hide
End Sub
