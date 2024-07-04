VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormImpNuevo 
   Caption         =   "EMBALAJES SRL"
   ClientHeight    =   14175
   ClientLeft      =   1830
   ClientTop       =   6420
   ClientWidth     =   13800
   OleObjectBlob   =   "FormImpNuevo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormImpNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable global para almacenar la ruta del PDF seleccionado
Dim pdfPathGlobal As String


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
    Me.Height = 397
    
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
    
    Label11.Enabled = isChecked
    Label12.Enabled = isChecked
    Label13.Enabled = isChecked
    Label14.Enabled = isChecked
    Label15.Enabled = isChecked
    Label16.Enabled = isChecked
    Label18.Enabled = isChecked
    Label19.Enabled = isChecked
    Label20.Enabled = isChecked
    Label21.Enabled = isChecked
    Label22.Enabled = isChecked
    Label23.Enabled = isChecked
    
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

' CARGAR DATOS DEL FORMULARIO EN LA TABLA
Private Sub Cargar_Click()
    Dim ws As Worksheet
    Dim i As Integer
    Dim mes As String
    Dim fila As Long
    Dim tbl As ListObject
    
    ' Configurar la hoja activa y la tabla
    Set ws = ActiveSheet
    On Error Resume Next
    Set tbl = ws.ListObjects(1) ' Intenta obtener la primera tabla en la hoja activa
    On Error GoTo 0
    
    ' Desactivar filtro si está aplicado
    If Not tbl Is Nothing Then
        If tbl.AutoFilter.FilterMode Then
            ' Desactivar filtro de forma permanente
            tbl.AutoFilter.ShowAllData
        End If
    End If
    
    ' Obtener próxima fila disponible en la tabla
    fila = tbl.ListRows.Count + 1
    
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
            InsertarFila tbl, fila, mes
            fila = fila + 1
        Next i
    Else
        ' Si no se selecciona Todos, se inserta una fila por cada checkbox seleccionado
        If CheckBoxEne.Value Then InsertarFila tbl, fila, "ene"
        If CheckBoxFeb.Value Then InsertarFila tbl, fila, "feb"
        If CheckBoxMar.Value Then InsertarFila tbl, fila, "mar"
        If CheckBoxAbr.Value Then InsertarFila tbl, fila, "abr"
        If CheckBoxMay.Value Then InsertarFila tbl, fila, "may"
        If CheckBoxJun.Value Then InsertarFila tbl, fila, "jun"
        If CheckBoxJul.Value Then InsertarFila tbl, fila, "jul"
        If CheckBoxAgo.Value Then InsertarFila tbl, fila, "ago"
        If CheckBoxSep.Value Then InsertarFila tbl, fila, "sep"
        If CheckBoxOct.Value Then InsertarFila tbl, fila, "oct"
        If CheckBoxNov.Value Then InsertarFila tbl, fila, "nov"
        If CheckBoxDic.Value Then InsertarFila tbl, fila, "dic"
    End If
        ' Llamar a la función para asignar los números de mes
    AsignarNumerosDeMes
    ' Cerrar el formulario para limpiar los campos
    Unload Me
End Sub

Private Sub InsertarFila(tbl As ListObject, ByRef fila As Long, mes As String)
    ' Añadir nueva fila al final de la tabla
    Dim nuevaFila As ListRow
    Set nuevaFila = tbl.ListRows.Add(AlwaysInsert:=True)
    
    With nuevaFila.Range
        .Cells(1, 1).Value = mes ' Columna A: Mes
        .Cells(1, 3).Value = TextBoxCUIT.Text ' Columna C: CUIT
        .Cells(1, 4).Value = TextBoxTipoServicio.Text ' Columna D: Tipo de Servicio
        .Cells(1, 5).Value = TextBoxDetalle.Text ' Columna E: Detalle
        .Cells(1, 6).Value = TextBoxDireccion.Text ' Columna F: Dirección
        .Cells(1, 7).Value = TextBoxCuenta.Text ' Columna G: Cuenta
        .Cells(1, 8).Value = TextBoxNroCuenta.Text ' Columna H: NroCuenta
        .Cells(1, 9).Value = TextBoxNroIdentificacion.Text ' Columna I: NroIdentificación
        .Cells(1, 16).Value = TextBoxObservaciones.Text
        
        ' Insertar la ruta del PDF como hipervínculo en la columna L (12) si hay un PDF seleccionado
        If pdfPathGlobal <> "" Then
            Dim rng As Range
            Set rng = .Cells(1, 12)
            rng.Hyperlinks.Add Anchor:=rng, Address:=pdfPathGlobal, TextToDisplay:="Abrir Comprobante"
        End If
        
        ' Insertar los valores de los campos de texto adicionales según el mes seleccionado o todos los meses
        Select Case mes
            Case "ene"
                .Cells(1, 13).Value = IIf(TextEnero.Enabled, TextEnero.Text, "")
                .Cells(1, 11).Value = IIf(TextBox1.Enabled, "'" & TextBox1.Text, "")
            Case "feb"
                .Cells(1, 13).Value = IIf(TextFebrero.Enabled, TextFebrero.Text, "")
                .Cells(1, 11).Value = IIf(TextBox2.Enabled, "'" & TextBox2.Text, "")
            Case "mar"
                .Cells(1, 13).Value = IIf(TextMarzo.Enabled, TextMarzo.Text, "")
                .Cells(1, 11).Value = IIf(TextBox3.Enabled, "'" & TextBox3.Text, "")
            Case "abr"
                .Cells(1, 13).Value = IIf(TextAbril.Enabled, TextAbril.Text, "")
                .Cells(1, 11).Value = IIf(TextBox4.Enabled, "'" & TextBox4.Text, "")
            Case "may"
                .Cells(1, 13).Value = IIf(TextMayo.Enabled, TextMayo.Text, "")
                .Cells(1, 11).Value = IIf(TextBox5.Enabled, "'" & TextBox5.Text, "")
            Case "jun"
                .Cells(1, 13).Value = IIf(TextJunio.Enabled, TextJunio.Text, "")
                .Cells(1, 11).Value = IIf(TextBox6.Enabled, "'" & TextBox6.Text, "")
            Case "jul"
                .Cells(1, 13).Value = IIf(TextJulio.Enabled, TextJulio.Text, "")
                .Cells(1, 11).Value = IIf(TextBox7.Enabled, "'" & TextBox7.Text, "")
            Case "ago"
                .Cells(1, 13).Value = IIf(TextAgosto.Enabled, TextAgosto.Text, "")
                .Cells(1, 11).Value = IIf(TextBox8.Enabled, "'" & TextBox8.Text, "")
            Case "sep"
                .Cells(1, 13).Value = IIf(TextSeptiembre.Enabled, TextSeptiembre.Text, "")
                .Cells(1, 11).Value = IIf(TextBox9.Enabled, "'" & TextBox9.Text, "")
            Case "oct"
                .Cells(1, 13).Value = IIf(TextOctubre.Enabled, TextOctubre.Text, "")
                .Cells(1, 11).Value = IIf(TextBox10.Enabled, "'" & TextBox10.Text, "")
            Case "nov"
                .Cells(1, 13).Value = IIf(TextNoviembre.Enabled, TextNoviembre.Text, "")
                .Cells(1, 11).Value = IIf(TextBox11.Enabled, "'" & TextBox11.Text, "")
            Case "dic"
                .Cells(1, 13).Value = IIf(TextDiciembre.Enabled, TextDiciembre.Text, "")
                .Cells(1, 11).Value = IIf(TextBox12.Enabled, "'" & TextBox12.Text, "")
        End Select
    End With
    
    ' Incrementar la variable fila para seguir controlando el número de la próxima fila
    fila = fila + 1
End Sub

' CARGAR PDF
Private Sub ButtonCargaImp_Click()
    ' Abrir el diálogo para seleccionar el PDF de Imp
    pdfImpPath = SelectPDFFile
    
    ' Verificar si se seleccionó un PDF
    If pdfImpPath <> "" Then
        ' Guardar la ruta del PDF seleccionado en la variable global
        pdfPathGlobal = pdfImpPath
        Me.MsjCargaImp.Caption = "PDF cargado OK"
    End If
End Sub
Private Function SelectPDFFile() As String
    Dim fd As FileDialog
    Dim selectedFile As String
    
    ' Configurar el diálogo de selección de archivo
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccionar archivo PDF"
        .Filters.Clear
        .Filters.Add "Archivos PDF", "*.pdf"
        .FilterIndex = 1
        .ButtonName = "Seleccionar"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        End If
    End With
    
    SelectPDFFile = selectedFile
End Function

' ASIGNAR NUMEROS EN LA COLUMNA Q SEGUN MES CARGADO
Function AsignarNumerosDeMes() As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim i As Long
    Dim mes As String
    Dim mesNumeros As Variant
    
    ' Definir los nombres de los meses y sus números correspondientes
    mesNumeros = Array("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
    
    ' Referenciar la hoja activa
    Set ws = ActiveSheet
    
    ' Buscar la primera tabla en la hoja activa
    On Error Resume Next
    Set tbl = ws.ListObjects(1)  ' Intenta obtener la primera tabla en la hoja activa
    On Error GoTo 0
    
    ' Verificar que la hoja y la tabla existen
    If ws Is Nothing Then
        MsgBox "La hoja activa no se encontró.", vbCritical
        AsignarNumerosDeMes = False
        Exit Function
    End If
    
    If tbl Is Nothing Then
        MsgBox "No se encontró ninguna tabla en la hoja activa.", vbCritical
        AsignarNumerosDeMes = False
        Exit Function
    End If
    
    ' Encontrar la última fila con datos en la columna A de la tabla
    lastRow = tbl.ListColumns("Mes").DataBodyRange.Rows.Count
    
    ' Iterar sobre las celdas de la columna A de la tabla
    For i = 1 To lastRow
        mes = Left(tbl.DataBodyRange(i, 1).Value, 3)  ' Obtener las primeras tres letras del mes en la columna A
        
        ' Buscar el mes en el array mesNumeros y asignar su posición + 1 (para que enero sea 1, febrero 2, etc.)
        For j = LBound(mesNumeros) To UBound(mesNumeros)
            If mes = mesNumeros(j) Then
                tbl.DataBodyRange(i, "Q").Value = j + 1
                Exit For
            End If
        Next j
    Next i
    
    AsignarNumerosDeMes = True
End Function




