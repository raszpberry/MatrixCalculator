Public Class Main
    Overloads Sub update(ByRef row, ByVal col, ByVal group)
        For i As Integer = row + 1 To 5
            For j As Integer = 1 To 5
                Me.Controls("grp_" & group).Controls(group & "_tb" & CStr(i) & CStr(j)).text = ""
            Next
        Next
        For i As Integer = 1 To 5
            For j As Integer = col + 1 To 5
                Me.Controls("grp_" & group).Controls(group & "_tb" & CStr(i) & CStr(j)).text = ""
            Next
        Next
    End Sub
#Region "Matrix A Keypress"
    Private Sub TextBox1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb11.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb12.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb13.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb14.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox5_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb21.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox6_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb22.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox7_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb23.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox8_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb24.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox9_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb31.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox10_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb32.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox11_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb33.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox12_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb34.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox13_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb41.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox14_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb42.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox15_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb43.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox16_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb44.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
#End Region
#Region "Matrix B Keypress"
    Private Sub TextBox17_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb11.KeyPress, b_tb11.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox18_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb12.KeyPress, b_tb12.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox19_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb13.KeyPress, b_tb13.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox20_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb14.KeyPress, b_tb14.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox21_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb21.KeyPress, b_tb21.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox22_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb22.KeyPress, b_tb22.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox23_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb23.KeyPress, b_tb23.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox24_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb24.KeyPress, b_tb24.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox25_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb31.KeyPress, b_tb31.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox26_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb32.KeyPress, b_tb32.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox27_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb33.KeyPress, b_tb33.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox28_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb34.KeyPress, b_tb34.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox29_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb41.KeyPress, b_tb41.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox30_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb42.KeyPress, b_tb42.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox31_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb43.KeyPress, b_tb43.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox32_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles a_tb44.KeyPress, b_tb44.KeyPress
        If Not (Char.IsNumber(e.KeyChar) Or Asc(e.KeyChar) = 8 OrElse e.KeyChar = "") Then
            e.Handled = True
        End If
    End Sub
#End Region
#Region "Main Buttons"
    Private Sub btn_ext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ext.Click
        Me.Close()
    End Sub
    Private Sub btn_clr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_clr.Click
        For Each Control As Control In Me.grp_a.Controls
            If TypeOf Control Is TextBox Then
                Control.Text = String.Empty
            End If
        Next
        For Each Control As Control In Me.grp_b.Controls
            If TypeOf Control Is TextBox Then
                Control.Text = String.Empty
            End If
        Next
        For Each Control As Control In Me.grp_ans.Controls
            If TypeOf Control Is TextBox Then
                Control.Text = String.Empty
            End If
        Next
    End Sub
    Private Sub btn_clra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_clra.Click
        For Each Control As Control In Me.grp_a.Controls
            If TypeOf Control Is TextBox Then
                Control.Text = String.Empty
            End If
        Next
    End Sub
    Private Sub btn_clrb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_clrb.Click
        For Each Control As Control In Me.grp_b.Controls
            If TypeOf Control Is TextBox Then
                Control.Text = String.Empty
            End If
        Next
    End Sub
#End Region
#Region "Matrix A Buttons"
    Private Sub a_addc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles a_addc.Click
        '1
        If a_tb11.Visible = True And a_tb21.Visible = False And a_tb12.Visible = False Then
            a_tb12.Visible = True
            Label3.Text = "1x2"
            Label6.Text = "2"
        ElseIf a_tb12.Visible = True And a_tb22.Visible = False And a_tb13.Visible = False Then
            a_tb13.Visible = True
            Label3.Text = "1x3"
            Label6.Text = "3"
        ElseIf a_tb13.Visible = True And a_tb23.Visible = False And a_tb14.Visible = False Then
            a_tb14.Visible = True
            Label3.Text = "1x4"
            Label6.Text = "4"
        ElseIf Label3.Text = "1x4" Then
            a_tb15.Visible = True
            Label3.Text = "1x5"
            Label6.Text = "5"
            '2
        ElseIf a_tb21.Visible = True And a_tb31.Visible = False And a_tb22.Visible = False Then
            a_tb12.Visible = True
            a_tb22.Visible = True
            Label3.Text = "2x2"
            Label6.Text = "2"
        ElseIf a_tb22.Visible = True And a_tb32.Visible = False And a_tb23.Visible = False Then
            a_tb13.Visible = True
            a_tb23.Visible = True
            Label3.Text = "2x3"
            Label6.Text = "3"
        ElseIf a_tb23.Visible = True And a_tb33.Visible = False And a_tb24.Visible = False Then
            a_tb14.Visible = True
            a_tb24.Visible = True
            Label3.Text = "2x4"
            Label6.Text = "4"
        ElseIf Label3.Text = "2x4" Then
            a_tb15.Visible = True
            a_tb25.Visible = True
            Label3.Text = "2x5"
            Label6.Text = "5"
            '3
        ElseIf a_tb31.Visible = True And a_tb41.Visible = False And a_tb32.Visible = False Then
            a_tb12.Visible = True
            a_tb22.Visible = True
            a_tb32.Visible = True
            Label3.Text = "3x2"
            Label6.Text = "2"
        ElseIf a_tb32.Visible = True And a_tb42.Visible = False And a_tb33.Visible = False Then
            a_tb13.Visible = True
            a_tb23.Visible = True
            a_tb33.Visible = True
            Label3.Text = "3x3"
            Label6.Text = "3"
        ElseIf a_tb33.Visible = True And a_tb43.Visible = False And a_tb34.Visible = False Then
            a_tb14.Visible = True
            a_tb24.Visible = True
            a_tb34.Visible = True
            Label3.Text = "3x4"
            Label6.Text = "4"
        ElseIf Label3.Text = "3x4" Then
            a_tb15.Visible = True
            a_tb25.Visible = True
            a_tb35.Visible = True
            Label3.Text = "3x5"
            Label6.Text = "5"
            '4
        ElseIf a_tb41.Visible = True And Label3.Text = "4x1" And a_tb42.Visible = False Then
            a_tb12.Visible = True
            a_tb22.Visible = True
            a_tb32.Visible = True
            a_tb42.Visible = True
            Label3.Text = "4x2"
            Label6.Text = "2"
        ElseIf a_tb42.Visible = True And Label3.Text = "4x2" And a_tb43.Visible = False Then
            a_tb13.Visible = True
            a_tb23.Visible = True
            a_tb33.Visible = True
            a_tb43.Visible = True
            Label3.Text = "4x3"
            Label6.Text = "3"
        ElseIf a_tb43.Visible = True And Label3.Text = "4x3" And a_tb44.Visible = False Then
            a_tb14.Visible = True
            a_tb24.Visible = True
            a_tb34.Visible = True
            a_tb44.Visible = True
            Label3.Text = "4x4"
            Label6.Text = "4"
        ElseIf Label3.Text = "4x4" Then
            a_tb15.Visible = True
            a_tb25.Visible = True
            a_tb35.Visible = True
            a_tb45.Visible = True
            Label3.Text = "4x5"
            Label6.Text = "5"
            '5
        ElseIf Label3.Text = "5x1" And a_tb51.Visible = True And a_tb52.Visible = False Then
            a_tb12.Visible = True
            a_tb22.Visible = True
            a_tb32.Visible = True
            a_tb42.Visible = True
            a_tb52.Visible = True
            Label3.Text = "5x2"
            Label6.Text = "2"
        ElseIf Label3.Text = "5x2" And a_tb52.Visible = True And a_tb53.Visible = False Then
            a_tb13.Visible = True
            a_tb23.Visible = True
            a_tb33.Visible = True
            a_tb43.Visible = True
            a_tb53.Visible = True
            Label3.Text = "5x3"
            Label6.Text = "3"
        ElseIf Label3.Text = "5x3" And a_tb53.Visible = True And a_tb54.Visible = False Then
            a_tb14.Visible = True
            a_tb24.Visible = True
            a_tb34.Visible = True
            a_tb44.Visible = True
            a_tb54.Visible = True
            Label3.Text = "5x4"
            Label6.Text = "4"
        ElseIf Label3.Text = "5x4" And a_tb54.Visible = True And a_tb55.Visible = False Then
            a_tb15.Visible = True
            a_tb25.Visible = True
            a_tb35.Visible = True
            a_tb45.Visible = True
            a_tb55.Visible = True
            Label3.Text = "5x5"
            Label6.Text = "5"
        End If
    End Sub
    Private Sub a_addr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles a_addr.Click
        '1
        If a_tb11.Visible = True And a_tb12.Visible = False And a_tb21.Visible = False Then
            a_tb21.Visible = True
            Label3.Text = "2x1"
            Label6.Text = "1"
        ElseIf a_tb21.Visible = True And a_tb22.Visible = False And a_tb31.Visible = False Then
            a_tb31.Visible = True
            Label3.Text = "3x1"
            Label6.Text = "1"
        ElseIf a_tb31.Visible = True And a_tb32.Visible = False And a_tb41.Visible = False Then
            a_tb41.Visible = True
            Label3.Text = "4x1"
            Label6.Text = "1"
        ElseIf Label3.Text = "4x1" Then
            a_tb51.Visible = True
            Label3.Text = "5x1"
            Label6.Text = "1"
            '2
        ElseIf a_tb12.Visible = True And a_tb13.Visible = False And a_tb21.Visible = False Then
            a_tb21.Visible = True
            a_tb22.Visible = True
            Label3.Text = "2x2"
            Label6.Text = "2"
        ElseIf a_tb22.Visible = True And a_tb23.Visible = False And a_tb31.Visible = False Then
            a_tb31.Visible = True
            a_tb32.Visible = True
            Label3.Text = "3x2"
            Label6.Text = "2"
        ElseIf a_tb32.Visible = True And a_tb33.Visible = False And a_tb41.Visible = False Then
            a_tb41.Visible = True
            a_tb42.Visible = True
            Label3.Text = "4x2"
            Label6.Text = "2"
        ElseIf Label3.Text = "4x2" Then
            a_tb51.Visible = True
            a_tb52.Visible = True
            Label3.Text = "5x2"
            Label6.Text = "2"
            '3
        ElseIf a_tb13.Visible = True And a_tb14.Visible = False And a_tb21.Visible = False Then
            a_tb21.Visible = True
            a_tb22.Visible = True
            a_tb23.Visible = True
            Label3.Text = "2x3"
            Label6.Text = "3"
        ElseIf a_tb23.Visible = True And a_tb24.Visible = False And a_tb31.Visible = False Then
            a_tb31.Visible = True
            a_tb32.Visible = True
            a_tb33.Visible = True
            Label3.Text = "3x3"
            Label6.Text = "3"
        ElseIf a_tb33.Visible = True And a_tb34.Visible = False And a_tb41.Visible = False Then
            a_tb41.Visible = True
            a_tb42.Visible = True
            a_tb43.Visible = True
            Label3.Text = "4x3"
            Label6.Text = "3"
        ElseIf Label3.Text = "4x3" Then
            a_tb51.Visible = True
            a_tb52.Visible = True
            a_tb53.Visible = True
            Label3.Text = "5x3"
            Label6.Text = "3"
            '4
        ElseIf a_tb14.Visible = True And a_tb21.Visible = False And Label3.Text = "1x4" Then
            a_tb21.Visible = True
            a_tb22.Visible = True
            a_tb23.Visible = True
            a_tb24.Visible = True
            Label3.Text = "2x4"
            Label6.Text = "4"
        ElseIf a_tb14.Visible = True And a_tb31.Visible = False And Label3.Text = "2x4" Then
            a_tb31.Visible = True
            a_tb32.Visible = True
            a_tb33.Visible = True
            a_tb34.Visible = True
            Label3.Text = "3x4"
            Label6.Text = "4"
        ElseIf a_tb31.Visible = True And a_tb41.Visible = False And Label3.Text = "3x4" Then
            a_tb41.Visible = True
            a_tb42.Visible = True
            a_tb43.Visible = True
            a_tb44.Visible = True
            Label3.Text = "4x4"
            Label6.Text = "4"
        ElseIf Label3.Text = "4x4" Then
            a_tb51.Visible = True
            a_tb52.Visible = True
            a_tb53.Visible = True
            a_tb54.Visible = True
            Label3.Text = "5x4"
            Label6.Text = "4"
            '5
        ElseIf Label3.Text = "1x5" And a_tb15.Visible = True And a_tb25.Visible = False Then
            a_tb21.Visible = True
            a_tb22.Visible = True
            a_tb23.Visible = True
            a_tb24.Visible = True
            a_tb25.Visible = True
            Label3.Text = "2x5"
            Label6.Text = "5"
        ElseIf Label3.Text = "2x5" And a_tb25.Visible = True And a_tb35.Visible = False Then
            a_tb31.Visible = True
            a_tb32.Visible = True
            a_tb33.Visible = True
            a_tb34.Visible = True
            a_tb35.Visible = True
            Label3.Text = "3x5"
            Label6.Text = "5"
        ElseIf Label3.Text = "3x5" And a_tb35.Visible = True And a_tb45.Visible = False Then
            a_tb41.Visible = True
            a_tb42.Visible = True
            a_tb43.Visible = True
            a_tb44.Visible = True
            a_tb45.Visible = True
            Label3.Text = "4x5"
            Label6.Text = "5"
        ElseIf Label3.Text = "4x5" And a_tb45.Visible = True And a_tb55.Visible = False Then
            a_tb51.Visible = True
            a_tb52.Visible = True
            a_tb53.Visible = True
            a_tb54.Visible = True
            a_tb55.Visible = True
            Label3.Text = "5x5"
            Label6.Text = "5"
        End If
    End Sub
    Private Sub a_redc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles a_redc.Click
        '1
        If a_tb12.Visible = True And a_tb13.Visible = False And a_tb22.Visible = False Then
            a_tb12.Visible = False
            Label3.Text = "1x1"
            Label6.Text = "1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb13.Visible = True And a_tb14.Visible = False And a_tb23.Visible = False Then
            a_tb13.Visible = False
            Label3.Text = "1x2"
            Label6.Text = "2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb14.Visible = True And a_tb15.Visible = False And a_tb24.Visible = False Then
            a_tb14.Visible = False
            Label3.Text = "1x3"
            Label6.Text = "3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb15.Visible = True And a_tb25.Visible = False Then
            a_tb15.Visible = False
            Label3.Text = "1x4"
            Label6.Text = "4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
            '2
        ElseIf a_tb22.Visible = True And a_tb23.Visible = False And a_tb32.Visible = False And a_tb33.Visible = False Then
            a_tb22.Visible = False
            a_tb12.Visible = False
            Label3.Text = "2x1"
            Label6.Text = "1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb23.Visible = True And a_tb24.Visible = False And a_tb33.Visible = False And a_tb34.Visible = False Then
            a_tb13.Visible = False
            a_tb23.Visible = False
            Label3.Text = "2x2"
            Label6.Text = "2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb24.Visible = True And a_tb25.Visible = False And a_tb34.Visible = False And a_tb35.Visible = False Then
            a_tb14.Visible = False
            a_tb24.Visible = False
            Label3.Text = "2x3"
            Label6.Text = "3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb25.Visible = True And a_tb35.Visible = False Then
            a_tb15.Visible = False
            a_tb25.Visible = False
            Label3.Text = "2x4"
            Label6.Text = "4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
            '3
        ElseIf a_tb31.Visible = True And a_tb32.Visible = False And a_tb41.Visible = False And a_tb42.Visible = False Then
            Label6.Text = "1"
        ElseIf a_tb32.Visible = True And a_tb33.Visible = False And a_tb42.Visible = False And a_tb43.Visible = False Then
            a_tb22.Visible = False
            a_tb12.Visible = False
            a_tb32.Visible = False
            Label3.Text = "3x1"
            Label6.Text = "1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb33.Visible = True And a_tb34.Visible = False And a_tb43.Visible = False And a_tb44.Visible = False Then
            a_tb13.Visible = False
            a_tb23.Visible = False
            a_tb33.Visible = False
            Label3.Text = "3x2"
            Label6.Text = "2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb34.Visible = True And a_tb35.Visible = False And a_tb44.Visible = False And a_tb45.Visible = False Then
            a_tb14.Visible = False
            a_tb24.Visible = False
            a_tb34.Visible = False
            Label3.Text = "3x3"
            Label6.Text = "3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb35.Visible = True And a_tb45.Visible = False Then
            a_tb15.Visible = False
            a_tb25.Visible = False
            a_tb35.Visible = False
            Label3.Text = "3x4"
            Label6.Text = "4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
            '4
        ElseIf a_tb42.Visible = True And a_tb52.Visible = False And a_tb43.Visible = False And a_tb53.Visible = False Then
            a_tb12.Visible = False
            a_tb22.Visible = False
            a_tb32.Visible = False
            a_tb42.Visible = False
            Label3.Text = "4x1"
            Label6.Text = "1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb43.Visible = True And a_tb53.Visible = False And a_tb44.Visible = False And a_tb54.Visible = False Then
            a_tb13.Visible = False
            a_tb23.Visible = False
            a_tb33.Visible = False
            a_tb43.Visible = False
            Label3.Text = "4x2"
            Label6.Text = "2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb44.Visible = True And a_tb54.Visible = False And a_tb45.Visible = False And a_tb55.Visible = False Then
            a_tb14.Visible = False
            a_tb24.Visible = False
            a_tb34.Visible = False
            a_tb44.Visible = False
            Label3.Text = "4x3"
            Label6.Text = "3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb45.Visible = True And a_tb55.Visible = False Then
            a_tb15.Visible = False
            a_tb25.Visible = False
            a_tb35.Visible = False
            a_tb45.Visible = False
            Label3.Text = "4x4"
            Label6.Text = "4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
            '5
        ElseIf Label3.Text = "5x2" And a_tb52.Visible = True Then
            a_tb12.Visible = False
            a_tb22.Visible = False
            a_tb32.Visible = False
            a_tb42.Visible = False
            a_tb52.Visible = False
            Label3.Text = "5x1"
            Label6.Text = "1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf Label3.Text = "5x3" And a_tb53.Visible = True Then
            a_tb13.Visible = False
            a_tb23.Visible = False
            a_tb33.Visible = False
            a_tb43.Visible = False
            a_tb53.Visible = False
            Label3.Text = "5x2"
            Label6.Text = "2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf Label3.Text = "5x4" And a_tb54.Visible = True Then
            a_tb14.Visible = False
            a_tb24.Visible = False
            a_tb34.Visible = False
            a_tb44.Visible = False
            a_tb54.Visible = False
            Label3.Text = "5x3"
            Label6.Text = "3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf Label3.Text = "5x5" And a_tb55.Visible = True Then
            a_tb15.Visible = False
            a_tb25.Visible = False
            a_tb35.Visible = False
            a_tb45.Visible = False
            a_tb55.Visible = False
            Label3.Text = "5x4"
            Label6.Text = "4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        End If
    End Sub
    Private Sub a_redr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles a_redr.Click
        '1 
        If a_tb21.Visible = True And a_tb12.Visible = False And a_tb31.Visible = False Then
            a_tb21.Visible = False
            Label3.Text = "1x1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb31.Visible = True And a_tb22.Visible = False And a_tb41.Visible = False Then
            a_tb31.Visible = False
            Label3.Text = "2x1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb41.Visible = True And a_tb32.Visible = False And a_tb51.Visible = False Then
            a_tb41.Visible = False
            Label3.Text = "3x1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf Label3.Text = "5x1" And a_tb51.Visible = True Then
            a_tb51.Visible = False
            Label3.Text = "4x1"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
            '2
        ElseIf a_tb22.Visible = True And a_tb23.Visible = False And a_tb32.Visible = False And a_tb33.Visible = False Then
            a_tb21.Visible = False
            a_tb22.Visible = False
            Label3.Text = "1x2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb32.Visible = True And a_tb33.Visible = False And a_tb42.Visible = False And a_tb43.Visible = False Then
            a_tb31.Visible = False
            a_tb32.Visible = False
            Label3.Text = "2x2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb42.Visible = True And a_tb43.Visible = False And a_tb52.Visible = False And a_tb53.Visible = False Then
            a_tb41.Visible = False
            a_tb42.Visible = False
            Label3.Text = "3x2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf Label3.Text = "5x2" And a_tb52.Visible = True Then
            a_tb51.Visible = False
            a_tb52.Visible = False
            Label3.Text = "4x2"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
            '3
        ElseIf a_tb23.Visible = True And a_tb24.Visible = False And a_tb33.Visible = False And a_tb34.Visible = False Then
            a_tb21.Visible = False
            a_tb22.Visible = False
            a_tb23.Visible = False
            Label3.Text = "1x3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb33.Visible = True And a_tb34.Visible = False And a_tb43.Visible = False And a_tb44.Visible = False Then
            a_tb31.Visible = False
            a_tb32.Visible = False
            a_tb33.Visible = False
            Label3.Text = "2x3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb43.Visible = True And a_tb44.Visible = False And a_tb53.Visible = False And a_tb54.Visible = False Then
            a_tb41.Visible = False
            a_tb42.Visible = False
            a_tb43.Visible = False
            Label3.Text = "3x3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb53.Visible = True And Label3.Text = "5x3" Then
            a_tb51.Visible = False
            a_tb52.Visible = False
            a_tb53.Visible = False
            Label3.Text = "4x3"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
            '4
        ElseIf a_tb24.Visible = True And a_tb25.Visible = False And a_tb34.Visible = False And a_tb35.Visible = False Then
            a_tb21.Visible = False
            a_tb22.Visible = False
            a_tb23.Visible = False
            a_tb24.Visible = False
            Label3.Text = "1x4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb34.Visible = True And a_tb35.Visible = False And a_tb44.Visible = False And a_tb45.Visible = False Then
            a_tb31.Visible = False
            a_tb32.Visible = False
            a_tb33.Visible = False
            a_tb34.Visible = False
            Label3.Text = "2x4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb44.Visible = True And a_tb45.Visible = False And a_tb54.Visible = False And a_tb55.Visible = False Then
            a_tb41.Visible = False
            a_tb42.Visible = False
            a_tb43.Visible = False
            a_tb44.Visible = False
            Label3.Text = "3x4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb54.Visible = True And Label3.Text = "5x4" Then
            a_tb51.Visible = False
            a_tb52.Visible = False
            a_tb53.Visible = False
            a_tb54.Visible = False
            Label3.Text = "4x4"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb25.Visible = True And Label3.Text = "2x5" Then
            a_tb21.Visible = False
            a_tb22.Visible = False
            a_tb23.Visible = False
            a_tb24.Visible = False
            a_tb25.Visible = False
            Label3.Text = "1x5"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb35.Visible = True And Label3.Text = "3x5" Then
            a_tb31.Visible = False
            a_tb32.Visible = False
            a_tb33.Visible = False
            a_tb34.Visible = False
            a_tb35.Visible = False
            Label3.Text = "2x5"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb35.Visible = True And Label3.Text = "4x5" Then
            a_tb41.Visible = False
            a_tb42.Visible = False
            a_tb43.Visible = False
            a_tb44.Visible = False
            a_tb45.Visible = False
            Label3.Text = "3x5"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        ElseIf a_tb45.Visible = True And Label3.Text = "5x5" Then
            a_tb51.Visible = False
            a_tb52.Visible = False
            a_tb53.Visible = False
            a_tb54.Visible = False
            a_tb55.Visible = False
            Label3.Text = "4x5"
            update(CInt(Label3.Text.Substring(0, 1)), CInt(Label3.Text.Substring(2, 1)), "a")
        End If

    End Sub
#End Region
#Region "Matrix B Buttons"
    Private Sub b_addc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_addc.Click
        '1
        If b_tb11.Visible = True And b_tb21.Visible = False And b_tb12.Visible = False Then
            b_tb12.Visible = True
            Label4.Text = "1x2"
            Label7.Text = "1"
        ElseIf b_tb12.Visible = True And b_tb22.Visible = False And b_tb13.Visible = False Then
            b_tb13.Visible = True
            Label4.Text = "1x3"
            Label7.Text = "1"
        ElseIf b_tb13.Visible = True And b_tb23.Visible = False And b_tb14.Visible = False Then
            b_tb14.Visible = True
            Label4.Text = "1x4"
            Label7.Text = "1"
        ElseIf Label4.Text = "1x4" Then
            b_tb15.Visible = True
            Label4.Text = "1x5"
            Label7.Text = "1"
            '2
        ElseIf b_tb21.Visible = True And b_tb31.Visible = False And b_tb22.Visible = False Then
            b_tb12.Visible = True
            b_tb22.Visible = True
            Label4.Text = "2x2"
            Label7.Text = "2"
        ElseIf b_tb22.Visible = True And b_tb32.Visible = False And b_tb23.Visible = False Then
            b_tb13.Visible = True
            b_tb23.Visible = True
            Label4.Text = "2x3"
            Label7.Text = "2"
        ElseIf b_tb23.Visible = True And b_tb33.Visible = False And b_tb24.Visible = False Then
            b_tb14.Visible = True
            b_tb24.Visible = True
            Label4.Text = "2x4"
            Label7.Text = "2"
        ElseIf Label4.Text = "2x4" Then
            b_tb15.Visible = True
            b_tb25.Visible = True
            Label4.Text = "2x5"
            Label7.Text = "2"
            '3
        ElseIf b_tb31.Visible = True And b_tb41.Visible = False And b_tb32.Visible = False Then
            b_tb12.Visible = True
            b_tb22.Visible = True
            b_tb32.Visible = True
            Label4.Text = "3x2"
            Label7.Text = "3"
        ElseIf b_tb32.Visible = True And b_tb42.Visible = False And b_tb33.Visible = False Then
            b_tb13.Visible = True
            b_tb23.Visible = True
            b_tb33.Visible = True
            Label4.Text = "3x3"
            Label7.Text = "3"
        ElseIf b_tb33.Visible = True And b_tb43.Visible = False And b_tb34.Visible = False Then
            b_tb14.Visible = True
            b_tb24.Visible = True
            b_tb34.Visible = True
            Label4.Text = "3x4"
            Label7.Text = "3"
        ElseIf Label4.Text = "3x4" Then
            b_tb15.Visible = True
            b_tb25.Visible = True
            b_tb35.Visible = True
            Label4.Text = "3x5"
            Label7.Text = "3"
            '4
        ElseIf b_tb41.Visible = True And Label4.Text = "4x1" And b_tb42.Visible = False Then
            b_tb12.Visible = True
            b_tb22.Visible = True
            b_tb32.Visible = True
            b_tb42.Visible = True
            Label4.Text = "4x2"
            Label7.Text = "4"
        ElseIf b_tb42.Visible = True And Label4.Text = "4x2" And b_tb43.Visible = False Then
            b_tb13.Visible = True
            b_tb23.Visible = True
            b_tb33.Visible = True
            b_tb43.Visible = True
            Label4.Text = "4x3"
            Label7.Text = "4"
        ElseIf b_tb43.Visible = True And Label4.Text = "4x3" And b_tb44.Visible = False Then
            b_tb14.Visible = True
            b_tb24.Visible = True
            b_tb34.Visible = True
            b_tb44.Visible = True
            Label4.Text = "4x4"
            Label7.Text = "4"
        ElseIf Label4.Text = "4x4" Then
            b_tb15.Visible = True
            b_tb25.Visible = True
            b_tb35.Visible = True
            b_tb45.Visible = True
            Label4.Text = "4x5"
            Label7.Text = "4"
            '5
        ElseIf Label4.Text = "5x1" And b_tb51.Visible = True And b_tb52.Visible = False Then
            b_tb12.Visible = True
            b_tb22.Visible = True
            b_tb32.Visible = True
            b_tb42.Visible = True
            b_tb52.Visible = True
            Label4.Text = "5x2"
            Label7.Text = "5"
        ElseIf Label4.Text = "5x2" And b_tb52.Visible = True And b_tb53.Visible = False Then
            b_tb13.Visible = True
            b_tb23.Visible = True
            b_tb33.Visible = True
            b_tb43.Visible = True
            b_tb53.Visible = True
            Label4.Text = "5x3"
            Label7.Text = "5"
        ElseIf Label4.Text = "5x3" And b_tb53.Visible = True And b_tb54.Visible = False Then
            b_tb14.Visible = True
            b_tb24.Visible = True
            b_tb34.Visible = True
            b_tb44.Visible = True
            b_tb54.Visible = True
            Label4.Text = "5x4"
            Label7.Text = "5"
        ElseIf Label4.Text = "5x4" And b_tb54.Visible = True And b_tb55.Visible = False Then
            b_tb15.Visible = True
            b_tb25.Visible = True
            b_tb35.Visible = True
            b_tb45.Visible = True
            b_tb55.Visible = True
            Label4.Text = "5x5"
            Label7.Text = "5"
        End If
    End Sub
    Private Sub b_addr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_addr.Click
        '1
        If b_tb11.Visible = True And b_tb12.Visible = False And b_tb21.Visible = False Then
            b_tb21.Visible = True
            Label4.Text = "2x1"
            Label7.Text = "2"
        ElseIf b_tb21.Visible = True And b_tb22.Visible = False And b_tb31.Visible = False Then
            b_tb31.Visible = True
            Label4.Text = "3x1"
            Label7.Text = "3"
        ElseIf b_tb31.Visible = True And b_tb32.Visible = False And b_tb41.Visible = False Then
            b_tb41.Visible = True
            Label4.Text = "4x1"
            Label7.Text = "4"
        ElseIf Label4.Text = "4x1" Then
            b_tb51.Visible = True
            Label4.Text = "5x1"
            Label7.Text = "5"
            '2
        ElseIf b_tb12.Visible = True And b_tb13.Visible = False And b_tb21.Visible = False Then
            b_tb21.Visible = True
            b_tb22.Visible = True
            Label4.Text = "2x2"
            Label7.Text = "2"
        ElseIf b_tb22.Visible = True And b_tb23.Visible = False And b_tb31.Visible = False Then
            b_tb31.Visible = True
            b_tb32.Visible = True
            Label4.Text = "3x2"
            Label7.Text = "3"
        ElseIf b_tb32.Visible = True And b_tb33.Visible = False And b_tb41.Visible = False Then
            b_tb41.Visible = True
            b_tb42.Visible = True
            Label4.Text = "4x2"
            Label7.Text = "4"
        ElseIf Label4.Text = "4x2" Then
            b_tb51.Visible = True
            b_tb52.Visible = True
            Label4.Text = "5x2"
            Label7.Text = "5"
            '3
        ElseIf b_tb13.Visible = True And b_tb14.Visible = False And b_tb21.Visible = False Then
            b_tb21.Visible = True
            b_tb22.Visible = True
            b_tb23.Visible = True
            Label4.Text = "2x3"
            Label7.Text = "2"
        ElseIf b_tb23.Visible = True And b_tb24.Visible = False And b_tb31.Visible = False Then
            b_tb31.Visible = True
            b_tb32.Visible = True
            b_tb33.Visible = True
            Label4.Text = "3x3"
            Label7.Text = "3"
        ElseIf b_tb33.Visible = True And b_tb34.Visible = False And b_tb41.Visible = False Then
            b_tb41.Visible = True
            b_tb42.Visible = True
            b_tb43.Visible = True
            Label4.Text = "4x3"
            Label7.Text = "4"
        ElseIf Label4.Text = "4x3" Then
            b_tb51.Visible = True
            b_tb52.Visible = True
            b_tb53.Visible = True
            Label4.Text = "5x3"
            Label7.Text = "5"
            '4
        ElseIf b_tb14.Visible = True And b_tb21.Visible = False And Label4.Text = "1x4" Then
            b_tb21.Visible = True
            b_tb22.Visible = True
            b_tb23.Visible = True
            b_tb24.Visible = True
            Label4.Text = "2x4"
            Label7.Text = "2"
        ElseIf b_tb14.Visible = True And b_tb31.Visible = False And Label4.Text = "2x4" Then
            b_tb31.Visible = True
            b_tb32.Visible = True
            b_tb33.Visible = True
            b_tb34.Visible = True
            Label4.Text = "3x4"
            Label7.Text = "3"
        ElseIf b_tb31.Visible = True And b_tb41.Visible = False And Label4.Text = "3x4" Then
            b_tb41.Visible = True
            b_tb42.Visible = True
            b_tb43.Visible = True
            b_tb44.Visible = True
            Label4.Text = "4x4"
            Label7.Text = "4"
        ElseIf Label4.Text = "4x4" Then
            b_tb51.Visible = True
            b_tb52.Visible = True
            b_tb53.Visible = True
            b_tb54.Visible = True
            Label4.Text = "5x4"
            Label7.Text = "5"
        ElseIf Label4.Text = "1x5" And b_tb15.Visible = True And b_tb25.Visible = False Then
            b_tb21.Visible = True
            b_tb22.Visible = True
            b_tb23.Visible = True
            b_tb24.Visible = True
            b_tb25.Visible = True
            Label4.Text = "2x5"
            Label7.Text = "2"
        ElseIf Label4.Text = "2x5" And b_tb25.Visible = True And b_tb35.Visible = False Then
            b_tb31.Visible = True
            b_tb32.Visible = True
            b_tb33.Visible = True
            b_tb34.Visible = True
            b_tb35.Visible = True
            Label4.Text = "3x5"
            Label7.Text = "3"
        ElseIf Label4.Text = "3x5" And b_tb35.Visible = True And b_tb45.Visible = False Then
            b_tb41.Visible = True
            b_tb42.Visible = True
            b_tb43.Visible = True
            b_tb44.Visible = True
            b_tb45.Visible = True
            Label4.Text = "4x5"
            Label7.Text = "4"
        ElseIf Label4.Text = "4x5" And b_tb45.Visible = True And b_tb55.Visible = False Then
            b_tb51.Visible = True
            b_tb52.Visible = True
            b_tb53.Visible = True
            b_tb54.Visible = True
            b_tb55.Visible = True
            Label4.Text = "5x5"
            Label7.Text = "5"
        End If
    End Sub
    Private Sub b_redc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_redc.Click
        '1
        If b_tb12.Visible = True And b_tb13.Visible = False And b_tb22.Visible = False Then
            b_tb12.Visible = False
            Label4.Text = "1x1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb13.Visible = True And b_tb14.Visible = False And b_tb23.Visible = False Then
            b_tb13.Visible = False
            Label4.Text = "1x2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb14.Visible = True And b_tb15.Visible = False And b_tb24.Visible = False Then
            b_tb14.Visible = False
            Label4.Text = "1x3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb15.Visible = True And b_tb25.Visible = False Then
            b_tb15.Visible = False
            Label4.Text = "1x4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
            '2
        ElseIf b_tb22.Visible = True And b_tb23.Visible = False And b_tb32.Visible = False And b_tb33.Visible = False Then
            b_tb22.Visible = False
            b_tb12.Visible = False
            Label4.Text = "2x1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb23.Visible = True And b_tb24.Visible = False And b_tb33.Visible = False And b_tb34.Visible = False Then
            b_tb13.Visible = False
            b_tb23.Visible = False
            Label4.Text = "2x2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb24.Visible = True And b_tb25.Visible = False And b_tb34.Visible = False And b_tb35.Visible = False Then
            b_tb14.Visible = False
            b_tb24.Visible = False
            Label4.Text = "2x3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb25.Visible = True And b_tb35.Visible = False Then
            b_tb15.Visible = False
            b_tb25.Visible = False
            Label4.Text = "2x4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
            '3
        ElseIf b_tb32.Visible = True And b_tb33.Visible = False And b_tb42.Visible = False And b_tb43.Visible = False Then
            b_tb22.Visible = False
            b_tb12.Visible = False
            b_tb32.Visible = False
            Label4.Text = "3x1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb33.Visible = True And b_tb34.Visible = False And b_tb43.Visible = False And b_tb44.Visible = False Then
            b_tb13.Visible = False
            b_tb23.Visible = False
            b_tb33.Visible = False
            Label4.Text = "3x2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb34.Visible = True And b_tb35.Visible = False And b_tb44.Visible = False And b_tb45.Visible = False Then
            b_tb14.Visible = False
            b_tb24.Visible = False
            b_tb34.Visible = False
            Label4.Text = "3x3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb35.Visible = True And b_tb45.Visible = False Then
            b_tb15.Visible = False
            b_tb25.Visible = False
            b_tb35.Visible = False
            Label4.Text = "3x4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
            '4
        ElseIf b_tb42.Visible = True And b_tb52.Visible = False And b_tb43.Visible = False And b_tb53.Visible = False Then
            b_tb12.Visible = False
            b_tb22.Visible = False
            b_tb32.Visible = False
            b_tb42.Visible = False
            Label4.Text = "4x1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb43.Visible = True And b_tb53.Visible = False And b_tb44.Visible = False And b_tb54.Visible = False Then
            b_tb13.Visible = False
            b_tb23.Visible = False
            b_tb33.Visible = False
            b_tb43.Visible = False
            Label4.Text = "4x2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb44.Visible = True And b_tb54.Visible = False And b_tb45.Visible = False And b_tb55.Visible = False Then
            b_tb14.Visible = False
            b_tb24.Visible = False
            b_tb34.Visible = False
            b_tb44.Visible = False
            Label4.Text = "4x3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb45.Visible = True And b_tb55.Visible = False Then
            b_tb15.Visible = False
            b_tb25.Visible = False
            b_tb35.Visible = False
            b_tb45.Visible = False
            Label4.Text = "4x4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
            '5
        ElseIf Label4.Text = "5x2" And b_tb52.Visible = True Then
            b_tb12.Visible = False
            b_tb22.Visible = False
            b_tb32.Visible = False
            b_tb42.Visible = False
            b_tb52.Visible = False
            Label4.Text = "5x1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf Label4.Text = "5x3" And b_tb53.Visible = True Then
            b_tb13.Visible = False
            b_tb23.Visible = False
            b_tb33.Visible = False
            b_tb43.Visible = False
            b_tb53.Visible = False
            Label4.Text = "5x2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf Label4.Text = "5x4" And b_tb54.Visible = True Then
            b_tb14.Visible = False
            b_tb24.Visible = False
            b_tb34.Visible = False
            b_tb44.Visible = False
            b_tb54.Visible = False
            Label4.Text = "5x3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf Label4.Text = "5x5" And b_tb55.Visible = True Then
            b_tb15.Visible = False
            b_tb25.Visible = False
            b_tb35.Visible = False
            b_tb45.Visible = False
            b_tb55.Visible = False
            Label4.Text = "5x4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        End If
    End Sub
    Private Sub b_redr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_redr.Click
        '1 
        If b_tb21.Visible = True And b_tb12.Visible = False And b_tb31.Visible = False Then
            b_tb21.Visible = False
            Label4.Text = "1x1"
            Label7.Text = "1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb31.Visible = True And b_tb22.Visible = False And b_tb41.Visible = False Then
            b_tb31.Visible = False
            Label4.Text = "2x1"
            Label7.Text = "2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb41.Visible = True And b_tb32.Visible = False And b_tb51.Visible = False Then
            b_tb41.Visible = False
            Label4.Text = "3x1"
            Label7.Text = "3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf Label4.Text = "5x1" And b_tb51.Visible = True Then
            b_tb51.Visible = False
            Label4.Text = "4x1"
            Label7.Text = "4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
            '2
        ElseIf b_tb22.Visible = True And b_tb23.Visible = False And b_tb32.Visible = False And b_tb33.Visible = False Then
            b_tb21.Visible = False
            b_tb22.Visible = False
            Label4.Text = "1x2"
            Label7.Text = "1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb32.Visible = True And b_tb33.Visible = False And b_tb42.Visible = False And b_tb43.Visible = False Then
            b_tb31.Visible = False
            b_tb32.Visible = False
            Label4.Text = "2x2"
            Label7.Text = "2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb42.Visible = True And b_tb43.Visible = False And b_tb52.Visible = False And b_tb53.Visible = False Then
            b_tb41.Visible = False
            b_tb42.Visible = False
            Label4.Text = "3x2"
            Label7.Text = "3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf Label4.Text = "5x2" And b_tb52.Visible = True Then
            b_tb51.Visible = False
            b_tb52.Visible = False
            Label4.Text = "4x2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
            Label7.Text = "4"
            '3
        ElseIf b_tb23.Visible = True And b_tb24.Visible = False And b_tb33.Visible = False And b_tb34.Visible = False Then
            b_tb21.Visible = False
            b_tb22.Visible = False
            b_tb23.Visible = False
            Label4.Text = "1x3"
            Label7.Text = "1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb33.Visible = True And b_tb34.Visible = False And b_tb43.Visible = False And b_tb44.Visible = False Then
            b_tb31.Visible = False
            b_tb32.Visible = False
            b_tb33.Visible = False
            Label4.Text = "2x3"
            Label7.Text = "2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb43.Visible = True And b_tb44.Visible = False And b_tb53.Visible = False And b_tb54.Visible = False Then
            b_tb41.Visible = False
            b_tb42.Visible = False
            b_tb43.Visible = False
            Label4.Text = "3x3"
            Label7.Text = "3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb53.Visible = True And Label4.Text = "5x3" Then
            b_tb51.Visible = False
            b_tb52.Visible = False
            b_tb53.Visible = False
            Label4.Text = "4x3"
            Label7.Text = "4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
            '4
        ElseIf b_tb24.Visible = True And b_tb25.Visible = False And b_tb34.Visible = False And b_tb35.Visible = False Then
            b_tb21.Visible = False
            b_tb22.Visible = False
            b_tb23.Visible = False
            b_tb24.Visible = False
            Label4.Text = "1x4"
            Label7.Text = "1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb34.Visible = True And b_tb35.Visible = False And b_tb44.Visible = False And b_tb45.Visible = False Then
            b_tb31.Visible = False
            b_tb32.Visible = False
            b_tb33.Visible = False
            b_tb34.Visible = False
            Label4.Text = "2x4"
            Label7.Text = "2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb44.Visible = True And b_tb45.Visible = False And b_tb54.Visible = False And b_tb55.Visible = False Then
            b_tb41.Visible = False
            b_tb42.Visible = False
            b_tb43.Visible = False
            b_tb44.Visible = False
            Label4.Text = "3x4"
            Label7.Text = "3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb54.Visible = True And Label4.Text = "5x4" Then
            b_tb51.Visible = False
            b_tb52.Visible = False
            b_tb53.Visible = False
            b_tb54.Visible = False
            Label4.Text = "4x4"
            Label7.Text = "4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb25.Visible = True And Label4.Text = "2x5" Then
            b_tb21.Visible = False
            b_tb22.Visible = False
            b_tb23.Visible = False
            b_tb24.Visible = False
            b_tb25.Visible = False
            Label4.Text = "1x5"
            Label7.Text = "1"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb35.Visible = True And Label4.Text = "3x5" Then
            b_tb31.Visible = False
            b_tb32.Visible = False
            b_tb33.Visible = False
            b_tb34.Visible = False
            b_tb35.Visible = False
            Label4.Text = "2x5"
            Label7.Text = "2"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb35.Visible = True And Label4.Text = "4x5" Then
            b_tb41.Visible = False
            b_tb42.Visible = False
            b_tb43.Visible = False
            b_tb44.Visible = False
            b_tb45.Visible = False
            Label4.Text = "3x5"
            Label7.Text = "3"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        ElseIf b_tb45.Visible = True And Label4.Text = "5x5" Then
            b_tb51.Visible = False
            b_tb52.Visible = False
            b_tb53.Visible = False
            b_tb54.Visible = False
            b_tb55.Visible = False
            Label4.Text = "4x5"
            Label7.Text = "4"
            update(CInt(Label4.Text.Substring(0, 1)), CInt(Label4.Text.Substring(2, 1)), "b")
        End If
    End Sub
#End Region
#Region "Operations"
#Region "Text Change"
    Private Sub Label3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.TextChanged
        Label5.Text = Label3.Text & "," & Label4.Text
    End Sub
    Private Sub Label4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.TextChanged
        Label5.Text = Label3.Text & "," & Label4.Text
    End Sub
    Private Sub Label9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.TextChanged
        If Val(Label9.Text) >= 2 Then
            For Each Control In Me.grp_ans.Controls
                If TypeOf Control Is TextBox Then
                    Control.text = String.Empty
                End If
            Next
        End If
    End Sub
#End Region
#Region "Add"
    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        If Label3.Text <> Label4.Text Then
            MsgBox("INVALID! Matrix A must equal to Matrix B in size", MsgBoxStyle.Exclamation)
        Else
            Label9.Text = Val(Label9.Text) + 1
            '1
            If a_tb11.Visible = True And a_tb21.Visible = False And a_tb12.Visible = False And b_tb11.Visible = True And b_tb21.Visible = False And b_tb12.Visible = False Then
                If a_tb11.Text = "" Or b_tb11.Text = "" Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                End If
            ElseIf a_tb21.Visible = True And a_tb31.Visible = False And a_tb22.Visible = False And b_tb21.Visible = True And b_tb31.Visible = False And b_tb22.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                End If
            ElseIf a_tb31.Visible = True And a_tb41.Visible = False And a_tb32.Visible = False And b_tb31.Visible = True And b_tb41.Visible = False And b_tb32.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                End If
            ElseIf a_tb41.Visible = True And a_tb51.Visible = False And a_tb42.Visible = False And b_tb41.Visible = True And b_tb52.Visible = False And b_tb42.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                End If
            ElseIf Label5.Text = "5x1,5x1" And a_tb51.Visible = True And b_tb51.Visible = True Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb51.Text = Val(a_tb51.Text) + Val(b_tb51.Text)
                End If
                '2
            ElseIf a_tb12.Visible = True And a_tb22.Visible = False And a_tb13.Visible = False And b_tb12.Visible = True And b_tb22.Visible = False And b_tb13.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                End If
            ElseIf a_tb22.Visible = True And a_tb32.Visible = False And a_tb23.Visible = False And b_tb22.Visible = True And b_tb32.Visible = False And b_tb23.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                End If
            ElseIf a_tb32.Visible = True And a_tb42.Visible = False And a_tb33.Visible = False And b_tb32.Visible = True And b_tb42.Visible = False And b_tb33.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                End If
            ElseIf a_tb42.Visible = True And a_tb52.Visible = False And a_tb43.Visible = False And b_tb42.Visible = True And b_tb52.Visible = False And b_tb43.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) + Val(b_tb42.Text)
                End If
            ElseIf Label5.Text = "5x2,5x2" And a_tb52.Visible = True And b_tb52.Visible = True Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) + Val(b_tb42.Text)
                    c_tb51.Text = Val(a_tb51.Text) + Val(b_tb51.Text)
                    c_tb52.Text = Val(a_tb52.Text) + Val(b_tb52.Text)
                End If
                '3
            ElseIf a_tb13.Visible = True And a_tb23.Visible = False And a_tb14.Visible = False And b_tb13.Visible = True And b_tb23.Visible = False And b_tb14.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                End If
            ElseIf a_tb23.Visible = True And a_tb33.Visible = False And a_tb24.Visible = False And b_tb23.Visible = True And b_tb33.Visible = False And b_tb24.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                End If
            ElseIf a_tb33.Visible = True And a_tb43.Visible = False And a_tb34.Visible = False And b_tb33.Visible = True And b_tb43.Visible = False And b_tb34.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                End If
            ElseIf a_tb43.Visible = True And a_tb53.Visible = False And a_tb44.Visible = False And b_tb43.Visible = True And b_tb53.Visible = False And b_tb44.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) + Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) + Val(b_tb43.Text)
                End If
            ElseIf Label5.Text = "5x3,5x3" And a_tb53.Visible = True And b_tb53.Visible = True Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "" Or b_tb53.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) + Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) + Val(b_tb43.Text)
                    c_tb51.Text = Val(a_tb51.Text) + Val(b_tb51.Text)
                    c_tb52.Text = Val(a_tb52.Text) + Val(b_tb52.Text)
                    c_tb53.Text = Val(a_tb53.Text) + Val(b_tb53.Text)
                End If
                '4
            ElseIf a_tb14.Visible = True And a_tb24.Visible = False And a_tb15.Visible = False And b_tb14.Visible = True And b_tb24.Visible = False And b_tb15.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                End If
            ElseIf a_tb24.Visible = True And a_tb34.Visible = False And a_tb25.Visible = False And b_tb24.Visible = True And b_tb34.Visible = False And b_tb25.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) + Val(b_tb24.Text)
                End If
            ElseIf a_tb34.Visible = True And a_tb44.Visible = False And a_tb35.Visible = False And b_tb34.Visible = True And b_tb44.Visible = False And b_tb35.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) + Val(b_tb24.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) + Val(b_tb34.Text)
                End If
            ElseIf a_tb44.Visible = True And a_tb54.Visible = False And a_tb45.Visible = False And b_tb44.Visible = True And b_tb54.Visible = False And b_tb45.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "" Or b_tb44.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) + Val(b_tb24.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) + Val(b_tb34.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) + Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) + Val(b_tb43.Text)
                    c_tb44.Text = Val(a_tb44.Text) + Val(b_tb44.Text)
                End If
            ElseIf Label5.Text = "5x4,5x4" And a_tb54.Visible = True And b_tb54.Visible = True Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "" Or b_tb53.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "" Or b_tb44.Text = "" Or b_tb54.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) + Val(b_tb24.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) + Val(b_tb34.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) + Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) + Val(b_tb43.Text)
                    c_tb44.Text = Val(a_tb44.Text) + Val(b_tb44.Text)
                    c_tb51.Text = Val(a_tb51.Text) + Val(b_tb51.Text)
                    c_tb52.Text = Val(a_tb52.Text) + Val(b_tb52.Text)
                    c_tb53.Text = Val(a_tb53.Text) + Val(b_tb53.Text)
                    c_tb54.Text = Val(a_tb54.Text) + Val(b_tb54.Text)
                End If
                '5
            ElseIf Label5.Text = "1x5,1x5" And a_tb15.Visible = True And b_tb15.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) + Val(b_tb15.Text)
                End If
            ElseIf Label5.Text = "2x5,2x5" And a_tb25.Visible = True And b_tb25.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (a_tb21.Text = "" Or a_tb22.Text = "" Or a_tb23.Text = "" Or a_tb24.Text = "" Or a_tb25.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) + Val(b_tb15.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) + Val(b_tb24.Text)
                    c_tb25.Text = Val(a_tb25.Text) + Val(b_tb25.Text)
                End If
            ElseIf Label5.Text = "3x5,3x5" And a_tb35.Visible = True And b_tb35.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (a_tb21.Text = "" Or a_tb22.Text = "" Or a_tb23.Text = "" Or a_tb24.Text = "" Or a_tb25.Text = "") Or (a_tb31.Text = "" Or a_tb32.Text = "" Or a_tb33.Text = "" Or a_tb34.Text = "" Or a_tb35.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) + Val(b_tb15.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) + Val(b_tb24.Text)
                    c_tb25.Text = Val(a_tb25.Text) + Val(b_tb25.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) + Val(b_tb34.Text)
                    c_tb35.Text = Val(a_tb35.Text) + Val(b_tb35.Text)
                End If
            ElseIf Label5.Text = "4x5,4x5" And a_tb45.Visible = True And b_tb45.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (a_tb21.Text = "" Or a_tb22.Text = "" Or a_tb23.Text = "" Or a_tb24.Text = "" Or a_tb25.Text = "") Or (a_tb31.Text = "" Or a_tb32.Text = "" Or a_tb33.Text = "" Or a_tb34.Text = "" Or a_tb35.Text = "") Or (a_tb41.Text = "" Or a_tb42.Text = "" Or a_tb43.Text = "" Or a_tb44.Text = "" Or a_tb45.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) + Val(b_tb15.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) + Val(b_tb24.Text)
                    c_tb25.Text = Val(a_tb25.Text) + Val(b_tb25.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) + Val(b_tb34.Text)
                    c_tb35.Text = Val(a_tb35.Text) + Val(b_tb35.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) + Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) + Val(b_tb43.Text)
                    c_tb44.Text = Val(a_tb44.Text) + Val(b_tb44.Text)
                    c_tb45.Text = Val(a_tb45.Text) + Val(b_tb45.Text)
                End If
            ElseIf Label5.Text = "5x5,5x5" And a_tb55.Visible = True And b_tb55.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (a_tb21.Text = "" Or a_tb22.Text = "" Or a_tb23.Text = "" Or a_tb24.Text = "" Or a_tb25.Text = "") Or (a_tb31.Text = "" Or a_tb32.Text = "" Or a_tb33.Text = "" Or a_tb34.Text = "" Or a_tb35.Text = "") Or (a_tb41.Text = "" Or a_tb42.Text = "" Or a_tb43.Text = "" Or a_tb44.Text = "" Or a_tb45.Text = "") Or (a_tb51.Text = "" Or a_tb52.Text = "" Or a_tb53.Text = "" Or a_tb54.Text = "" Or a_tb55.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "" Or b_tb55.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) + Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) + Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) + Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) + Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) + Val(b_tb15.Text)
                    c_tb21.Text = Val(a_tb21.Text) + Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) + Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) + Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) + Val(b_tb24.Text)
                    c_tb25.Text = Val(a_tb25.Text) + Val(b_tb25.Text)
                    c_tb31.Text = Val(a_tb31.Text) + Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) + Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) + Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) + Val(b_tb34.Text)
                    c_tb35.Text = Val(a_tb35.Text) + Val(b_tb35.Text)
                    c_tb41.Text = Val(a_tb41.Text) + Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) + Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) + Val(b_tb43.Text)
                    c_tb44.Text = Val(a_tb44.Text) + Val(b_tb44.Text)
                    c_tb45.Text = Val(a_tb45.Text) + Val(b_tb45.Text)
                    c_tb51.Text = Val(a_tb51.Text) + Val(b_tb51.Text)
                    c_tb52.Text = Val(a_tb52.Text) + Val(b_tb52.Text)
                    c_tb53.Text = Val(a_tb53.Text) + Val(b_tb53.Text)
                    c_tb54.Text = Val(a_tb54.Text) + Val(b_tb54.Text)
                    c_tb55.Text = Val(a_tb55.Text) + Val(b_tb55.Text)
                End If
            End If
        End If
        grp_ans.Text = "A + B"
    End Sub
#End Region
#Region "Subtract"
    Private Sub btn_sub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_sub.Click
        grp_ans.Text = "A - B"
        If Label3.Text <> Label4.Text Then
            MsgBox("INVALID! Matrix A must equal to Matrix B in size", MsgBoxStyle.Exclamation)
        Else
            Label9.Text = Val(Label9.Text) - 1
            '1
            If a_tb11.Visible = True And a_tb21.Visible = False And a_tb12.Visible = False And b_tb11.Visible = True And b_tb21.Visible = False And b_tb12.Visible = False Then
                If a_tb11.Text = "" Or b_tb11.Text = "" Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                End If
            ElseIf a_tb21.Visible = True And a_tb31.Visible = False And a_tb22.Visible = False And b_tb21.Visible = True And b_tb31.Visible = False And b_tb22.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                End If
            ElseIf a_tb31.Visible = True And a_tb41.Visible = False And a_tb32.Visible = False And b_tb31.Visible = True And b_tb41.Visible = False And b_tb32.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                End If
            ElseIf a_tb41.Visible = True And a_tb51.Visible = False And a_tb42.Visible = False And b_tb41.Visible = True And b_tb52.Visible = False And b_tb42.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                End If
            ElseIf Label5.Text = "5x1,5x1" And a_tb51.Visible = True And b_tb51.Visible = True Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb51.Text = Val(a_tb51.Text) - Val(b_tb51.Text)
                End If
                '2
            ElseIf a_tb12.Visible = True And a_tb22.Visible = False And a_tb13.Visible = False And b_tb12.Visible = True And b_tb22.Visible = False And b_tb13.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                End If
            ElseIf a_tb22.Visible = True And a_tb32.Visible = False And a_tb23.Visible = False And b_tb22.Visible = True And b_tb32.Visible = False And b_tb23.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                End If
            ElseIf a_tb32.Visible = True And a_tb42.Visible = False And a_tb33.Visible = False And b_tb32.Visible = True And b_tb42.Visible = False And b_tb33.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                End If
            ElseIf a_tb42.Visible = True And a_tb52.Visible = False And a_tb43.Visible = False And b_tb42.Visible = True And b_tb52.Visible = False And b_tb43.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) - Val(b_tb42.Text)
                End If
            ElseIf Label5.Text = "5x2,5x2" And a_tb52.Visible = True And b_tb52.Visible = True Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) - Val(b_tb42.Text)
                    c_tb51.Text = Val(a_tb51.Text) - Val(b_tb51.Text)
                    c_tb52.Text = Val(a_tb52.Text) - Val(b_tb52.Text)
                End If
                '3
            ElseIf a_tb13.Visible = True And a_tb23.Visible = False And a_tb14.Visible = False And b_tb13.Visible = True And b_tb23.Visible = False And b_tb14.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                End If
            ElseIf a_tb23.Visible = True And a_tb33.Visible = False And a_tb24.Visible = False And b_tb23.Visible = True And b_tb33.Visible = False And b_tb24.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                End If
            ElseIf a_tb33.Visible = True And a_tb43.Visible = False And a_tb34.Visible = False And b_tb33.Visible = True And b_tb43.Visible = False And b_tb34.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                End If
            ElseIf a_tb43.Visible = True And a_tb53.Visible = False And a_tb44.Visible = False And b_tb43.Visible = True And b_tb53.Visible = False And b_tb44.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "" Or b_tb53.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) - Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) - Val(b_tb43.Text)
                End If
            ElseIf Label5.Text = "5x3,5x3" And a_tb53.Visible = True And b_tb53.Visible = True Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "" Or b_tb53.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) - Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) - Val(b_tb43.Text)
                    c_tb51.Text = Val(a_tb51.Text) - Val(b_tb51.Text)
                    c_tb52.Text = Val(a_tb52.Text) - Val(b_tb52.Text)
                    c_tb53.Text = Val(a_tb53.Text) - Val(b_tb53.Text)
                End If
                '4
            ElseIf a_tb14.Visible = True And a_tb24.Visible = False And a_tb21.Visible = False And b_tb14.Visible = True And b_tb24.Visible = False And b_tb21.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                End If
            ElseIf a_tb24.Visible = True And a_tb34.Visible = False And a_tb31.Visible = False And b_tb24.Visible = True And b_tb34.Visible = False And b_tb31.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) - Val(b_tb24.Text)
                End If
            ElseIf a_tb34.Visible = True And a_tb44.Visible = False And a_tb41.Visible = False And b_tb34.Visible = True And b_tb44.Visible = False And b_tb41.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) - Val(b_tb24.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) - Val(b_tb34.Text)
                End If
            ElseIf a_tb44.Visible = True And a_tb54.Visible = False And a_tb45.Visible = False And b_tb44.Visible = True And b_tb54.Visible = False And b_tb45.Visible = False Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "" Or b_tb44.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) - Val(b_tb24.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) - Val(b_tb34.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) - Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) - Val(b_tb43.Text)
                    c_tb44.Text = Val(a_tb44.Text) - Val(b_tb44.Text)
                End If
            ElseIf Label5.Text = "5x4,5x4" And a_tb54.Visible = True And b_tb54.Visible = True Then
                If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "" Or b_tb53.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "" Or b_tb44.Text = "" Or b_tb54.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) - Val(b_tb24.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) - Val(b_tb34.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) - Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) - Val(b_tb43.Text)
                    c_tb44.Text = Val(a_tb44.Text) - Val(b_tb44.Text)
                    c_tb51.Text = Val(a_tb51.Text) - Val(b_tb51.Text)
                    c_tb52.Text = Val(a_tb52.Text) - Val(b_tb52.Text)
                    c_tb53.Text = Val(a_tb53.Text) - Val(b_tb53.Text)
                    c_tb54.Text = Val(a_tb54.Text) - Val(b_tb54.Text)
                End If
                '5
            ElseIf Label5.Text = "1x5,1x5" And a_tb15.Visible = True And b_tb15.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) - Val(b_tb15.Text)
                End If
            ElseIf Label5.Text = "2x5,2x5" And a_tb25.Visible = True And b_tb25.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (a_tb21.Text = "" Or a_tb22.Text = "" Or a_tb23.Text = "" Or a_tb24.Text = "" Or a_tb25.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) - Val(b_tb15.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) - Val(b_tb24.Text)
                    c_tb25.Text = Val(a_tb25.Text) - Val(b_tb25.Text)
                End If
            ElseIf Label5.Text = "3x5,3x5" And a_tb35.Visible = True And b_tb35.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (a_tb21.Text = "" Or a_tb22.Text = "" Or a_tb23.Text = "" Or a_tb24.Text = "" Or a_tb25.Text = "") Or (a_tb31.Text = "" Or a_tb32.Text = "" Or a_tb33.Text = "" Or a_tb34.Text = "" Or a_tb35.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) - Val(b_tb15.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) - Val(b_tb24.Text)
                    c_tb25.Text = Val(a_tb25.Text) - Val(b_tb25.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) - Val(b_tb34.Text)
                    c_tb35.Text = Val(a_tb35.Text) - Val(b_tb35.Text)
                End If
            ElseIf Label5.Text = "4x5,4x5" And a_tb45.Visible = True And b_tb45.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (a_tb21.Text = "" Or a_tb22.Text = "" Or a_tb23.Text = "" Or a_tb24.Text = "" Or a_tb25.Text = "") Or (a_tb31.Text = "" Or a_tb32.Text = "" Or a_tb33.Text = "" Or a_tb34.Text = "" Or a_tb35.Text = "") Or (a_tb41.Text = "" Or a_tb42.Text = "" Or a_tb43.Text = "" Or a_tb44.Text = "" Or a_tb45.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) - Val(b_tb15.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) - Val(b_tb24.Text)
                    c_tb25.Text = Val(a_tb25.Text) - Val(b_tb25.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) - Val(b_tb34.Text)
                    c_tb35.Text = Val(a_tb35.Text) - Val(b_tb35.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) - Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) - Val(b_tb43.Text)
                    c_tb44.Text = Val(a_tb44.Text) - Val(b_tb44.Text)
                    c_tb45.Text = Val(a_tb45.Text) - Val(b_tb45.Text)
                End If
            ElseIf Label5.Text = "5x5,5x5" And a_tb55.Visible = True And b_tb55.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (a_tb21.Text = "" Or a_tb22.Text = "" Or a_tb23.Text = "" Or a_tb24.Text = "" Or a_tb25.Text = "") Or (a_tb31.Text = "" Or a_tb32.Text = "" Or a_tb33.Text = "" Or a_tb34.Text = "" Or a_tb35.Text = "") Or (a_tb41.Text = "" Or a_tb42.Text = "" Or a_tb43.Text = "" Or a_tb44.Text = "" Or a_tb45.Text = "") Or (a_tb51.Text = "" Or a_tb52.Text = "" Or a_tb53.Text = "" Or a_tb54.Text = "" Or a_tb55.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "" Or b_tb55.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) - Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb12.Text) - Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb13.Text) - Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb14.Text) - Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb15.Text) - Val(b_tb15.Text)
                    c_tb21.Text = Val(a_tb21.Text) - Val(b_tb21.Text)
                    c_tb22.Text = Val(a_tb22.Text) - Val(b_tb22.Text)
                    c_tb23.Text = Val(a_tb23.Text) - Val(b_tb23.Text)
                    c_tb24.Text = Val(a_tb24.Text) - Val(b_tb24.Text)
                    c_tb25.Text = Val(a_tb25.Text) - Val(b_tb25.Text)
                    c_tb31.Text = Val(a_tb31.Text) - Val(b_tb31.Text)
                    c_tb32.Text = Val(a_tb32.Text) - Val(b_tb32.Text)
                    c_tb33.Text = Val(a_tb33.Text) - Val(b_tb33.Text)
                    c_tb34.Text = Val(a_tb34.Text) - Val(b_tb34.Text)
                    c_tb35.Text = Val(a_tb35.Text) - Val(b_tb35.Text)
                    c_tb41.Text = Val(a_tb41.Text) - Val(b_tb41.Text)
                    c_tb42.Text = Val(a_tb42.Text) - Val(b_tb42.Text)
                    c_tb43.Text = Val(a_tb43.Text) - Val(b_tb43.Text)
                    c_tb44.Text = Val(a_tb44.Text) - Val(b_tb44.Text)
                    c_tb45.Text = Val(a_tb45.Text) - Val(b_tb45.Text)
                    c_tb51.Text = Val(a_tb51.Text) - Val(b_tb51.Text)
                    c_tb52.Text = Val(a_tb52.Text) - Val(b_tb52.Text)
                    c_tb53.Text = Val(a_tb53.Text) - Val(b_tb53.Text)
                    c_tb54.Text = Val(a_tb54.Text) - Val(b_tb54.Text)
                    c_tb55.Text = Val(a_tb55.Text) - Val(b_tb55.Text)
                End If
            End If
        End If
    End Sub
#End Region
#Region "Multiply"
    Private Sub btn_mul_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_mul.Click
        grp_ans.Text = "A*B"
        If Label6.Text <> Label7.Text Or Label7.Text <> Label6.Text Then
            MsgBox("INVALID! Matrix A`s columns must equal to Matrix B`s rows", MsgBoxStyle.Exclamation)
        Else
            Label9.Text = Val(Label9.Text) + 1
            If a_tb11.Visible = True And a_tb12.Visible = False And a_tb21.Visible = False And b_tb11.Visible = True And b_tb12.Visible = False And b_tb21.Visible = False Then
                If a_tb11.Text = "" Or b_tb11.Text = "" Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                End If
                '1r 1-4c
            ElseIf a_tb11.Visible = True And a_tb12.Visible = False And a_tb21.Visible = False And b_tb12.Visible = True And b_tb13.Visible = False And b_tb21.Visible = False Then
                If a_tb11.Text = "" Or (b_tb11.Text = "" Or b_tb12.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                End If
                '1
            ElseIf a_tb11.Visible = True And a_tb12.Visible = False And a_tb21.Visible = False And b_tb13.Visible = True And b_tb14.Visible = False And b_tb21.Visible = False Then
                If a_tb11.Text = "" Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                End If
            ElseIf a_tb11.Visible = True And a_tb12.Visible = False And a_tb21.Visible = False And b_tb14.Visible = True And b_tb15.Visible = False And b_tb21.Visible = False Then
                If a_tb11.Text = "" Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                End If
            ElseIf Label5.Text = "1x1,1x5" And a_tb11.Visible = True And b_tb15.Visible = True Then
                If a_tb11.Text = "" Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                    c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                    c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                    c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                    c_tb15.Text = Val(a_tb11.Text) * Val(b_tb15.Text)
                End If
                '2
            ElseIf a_tb12.Visible = True And a_tb22.Visible = False And a_tb13.Visible = False And b_tb21.Visible = True And b_tb22.Visible = False And b_tb31.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                End If
            ElseIf a_tb12.Visible = True And a_tb22.Visible = False And a_tb13.Visible = False And b_tb22.Visible = True And b_tb23.Visible = False And b_tb31.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                End If
            ElseIf a_tb12.Visible = True And a_tb22.Visible = False And a_tb13.Visible = False And b_tb23.Visible = True And b_tb24.Visible = False And b_tb31.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                End If
            ElseIf a_tb12.Visible = True And a_tb22.Visible = False And a_tb13.Visible = False And b_tb24.Visible = True And b_tb25.Visible = False And b_tb31.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                    c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                End If
            ElseIf Label5.Text = "1x2,2x5" And a_tb12.Visible = True And b_tb25.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                    c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                    c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text))
                End If
                '3
            ElseIf a_tb13.Visible = True And a_tb23.Visible = False And a_tb14.Visible = False And b_tb31.Visible = True And b_tb32.Visible = False And b_tb41.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                End If
            ElseIf a_tb13.Visible = True And a_tb23.Visible = False And a_tb14.Visible = False And b_tb32.Visible = True And b_tb33.Visible = False And b_tb41.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                End If
            ElseIf a_tb13.Visible = True And a_tb23.Visible = False And a_tb14.Visible = False And b_tb33.Visible = True And b_tb34.Visible = False And b_tb41.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                End If
            ElseIf a_tb13.Visible = True And a_tb23.Visible = False And a_tb14.Visible = False And b_tb34.Visible = True And b_tb35.Visible = False And b_tb41.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                    c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                End If
            ElseIf Label5.Text = "1x3,3x5" And a_tb32.Visible = True And b_tb35.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                    c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                    c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text))
                End If
                '4
            ElseIf a_tb14.Visible = True And a_tb24.Visible = False And a_tb15.Visible = False And b_tb41.Visible = True And b_tb42.Visible = False And b_tb51.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                End If
            ElseIf a_tb14.Visible = True And a_tb24.Visible = False And a_tb15.Visible = False And b_tb42.Visible = True And b_tb43.Visible = False And b_tb52.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                End If
            ElseIf a_tb14.Visible = True And a_tb24.Visible = False And a_tb15.Visible = False And b_tb43.Visible = True And b_tb44.Visible = False And b_tb53.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                End If
            ElseIf a_tb14.Visible = True And a_tb24.Visible = False And a_tb15.Visible = False And b_tb44.Visible = True And b_tb45.Visible = False And b_tb54.Visible = False Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "" Or b_tb44.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                    c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text))
                End If
            ElseIf Label5.Text = "1x4,4x5" And a_tb14.Visible = True And b_tb45.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                    c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text))
                    c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text)) + (Val(a_tb14.Text) * Val(b_tb45.Text))
                End If
                '5
            ElseIf Label5.Text = "1x5,5x1" And a_tb15.Visible = True And b_tb51.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                End If
            ElseIf Label5.Text = "1x5,5x2" And a_tb15.Visible = True And b_tb52.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                End If
            ElseIf Label5.Text = "1x5,5x3" And a_tb15.Visible = True And b_tb53.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "" Or b_tb53.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                End If
            ElseIf Label5.Text = "1x5,5x4" And a_tb15.Visible = True And b_tb54.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "" Or b_tb53.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "" Or b_tb44.Text = "" Or b_tb54.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                    c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                End If
            ElseIf Label5.Text = "1x5,5x5" And a_tb15.Visible = True And b_tb55.Visible = True Then
                If (a_tb11.Text = "" Or a_tb12.Text = "" Or a_tb13.Text = "" Or a_tb14.Text = "" Or a_tb15.Text = "") Or (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "" Or b_tb51.Text = "") Or (b_tb12.Text = "" Or b_tb22.Text = "" Or b_tb32.Text = "" Or b_tb42.Text = "" Or b_tb52.Text = "") Or (b_tb13.Text = "" Or b_tb23.Text = "" Or b_tb33.Text = "" Or b_tb43.Text = "" Or b_tb53.Text = "") Or (b_tb14.Text = "" Or b_tb24.Text = "" Or b_tb34.Text = "" Or b_tb44.Text = "" Or b_tb54.Text = "") Or (b_tb15.Text = "" Or b_tb25.Text = "" Or b_tb35.Text = "" Or b_tb45.Text = "" Or b_tb55.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                    c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                    c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                    c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                    c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text)) + (Val(a_tb14.Text) * Val(b_tb45.Text)) + (Val(a_tb15.Text) * Val(b_tb55.Text))
                End If
                '1-4r 1c
            ElseIf a_tb21.Visible = True And a_tb22.Visible = False And a_tb31.Visible = False And b_tb11.Visible = True And b_tb12.Visible = False And b_tb21.Visible = False Then
                If b_tb11.Text = "" Or (a_tb11.Text = "" Or a_tb21.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                End If
            ElseIf a_tb31.Visible = True And a_tb32.Visible = False And a_tb41.Visible = False And b_tb11.Visible = True And b_tb12.Visible = False And b_tb21.Visible = False Then
                If b_tb11.Text = "" Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                    c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                End If
            ElseIf a_tb41.Visible = True And a_tb42.Visible = False And a_tb51.Visible = False And b_tb11.Visible = True And b_tb12.Visible = False And b_tb21.Visible = False Then
                If b_tb11.Text = "" Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                    c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                    c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                End If
            ElseIf Label5.Text = "5x1,1x1" And a_tb51.Visible = True And b_tb11.Visible = True Then
                If b_tb11.Text = "" Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                    c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                    c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                    c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                    c_tb51.Text = Val(a_tb51.Text) * Val(b_tb11.Text)
                End If
                '
            ElseIf a_tb22.Visible = True And a_tb23.Visible = False And a_tb32.Visible = False And b_tb21.Visible = True And b_tb22.Visible = False And b_tb31.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                End If
            ElseIf a_tb32.Visible = True And a_tb33.Visible = False And a_tb42.Visible = False And b_tb21.Visible = True And b_tb22.Visible = False And b_tb31.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                End If
            ElseIf a_tb42.Visible = True And a_tb43.Visible = False And b_tb52.Visible = True And b_tb21.Visible = True And b_tb22.Visible = False And b_tb31.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                    c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                End If
            ElseIf Label5.Text = "5x2,2x1" And a_tb52.Visible = True And b_tb21.Visible = True Then
                If (b_tb11.Text = "" Or b_tb21.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                    c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                    c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text))
                End If
                '
            ElseIf a_tb23.Visible = True And a_tb24.Visible = False And a_tb33.Visible = False And b_tb31.Visible = True And b_tb32.Visible = False And b_tb41.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                End If
            ElseIf a_tb33.Visible = True And a_tb34.Visible = False And a_tb43.Visible = False And b_tb31.Visible = True And b_tb32.Visible = False And b_tb41.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                End If
            ElseIf a_tb43.Visible = True And a_tb44.Visible = False And a_tb53.Visible = False And b_tb31.Visible = True And b_tb32.Visible = False And b_tb41.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                    c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                End If
            ElseIf Label5.Text = "5x3,3x1" And a_tb53.Visible = True And b_tb31.Visible = True Then
                If (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                    c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                    c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text))
                End If
                '
            ElseIf a_tb24.Visible = True And a_tb25.Visible = False And a_tb34.Visible = False And b_tb41.Visible = True And b_tb42.Visible = False And b_tb51.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                End If
            ElseIf a_tb34.Visible = True And a_tb35.Visible = False And a_tb44.Visible = False And b_tb41.Visible = True And b_tb42.Visible = False And b_tb51.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                End If
            ElseIf a_tb44.Visible = True And a_tb45.Visible = False And a_tb54.Visible = False And b_tb41.Visible = True And b_tb42.Visible = False And b_tb51.Visible = False Then
                If (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                    c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text))
                End If
            ElseIf Label5.Text = "5x4,4x1" And a_tb54.Visible = True And b_tb41.Visible = True Then
                If (b_tb11.Text = "" Or b_tb21.Text = "" Or b_tb31.Text = "" Or b_tb41.Text = "") Or (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Then
                    MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                Else
                    c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                    c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                    c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                    c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text))
                    c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text)) + (Val(a_tb54.Text) * Val(b_tb41.Text))
                End If
                '----
            ElseIf Label4.Text = "1x2" Then
                If Label3.Text = "2x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                    End If
                ElseIf Label3.Text = "3x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                    End If
                ElseIf Label3.Text = "4x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                        c_tb42.Text = Val(a_tb41.Text) * Val(b_tb12.Text)
                    End If
                ElseIf Label3.Text = "5x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                        c_tb42.Text = Val(a_tb41.Text) * Val(b_tb12.Text)
                        c_tb51.Text = Val(a_tb51.Text) * Val(b_tb11.Text)
                        c_tb52.Text = Val(a_tb51.Text) * Val(b_tb12.Text)
                    End If
                End If
                '
            ElseIf Label4.Text = "1x3" Then
                If Label3.Text = "2x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                    End If
                ElseIf Label3.Text = "3x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                    End If
                ElseIf Label3.Text = "4x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                        c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                        c_tb42.Text = Val(a_tb41.Text) * Val(b_tb12.Text)
                        c_tb43.Text = Val(a_tb41.Text) * Val(b_tb13.Text)
                    End If
                ElseIf Label3.Text = "5x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                        c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                        c_tb42.Text = Val(a_tb41.Text) * Val(b_tb12.Text)
                        c_tb43.Text = Val(a_tb41.Text) * Val(b_tb13.Text)
                        c_tb51.Text = Val(a_tb51.Text) * Val(b_tb11.Text)
                        c_tb52.Text = Val(a_tb51.Text) * Val(b_tb12.Text)
                        c_tb53.Text = Val(a_tb51.Text) * Val(b_tb13.Text)
                    End If
                End If
            ElseIf Label4.Text = "1x4" Then
                If Label3.Text = "2x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb24.Text = Val(a_tb21.Text) * Val(b_tb14.Text)
                    End If
                ElseIf Label3.Text = "3x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb24.Text = Val(a_tb21.Text) * Val(b_tb14.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                        c_tb34.Text = Val(a_tb31.Text) * Val(b_tb14.Text)
                    End If
                ElseIf Label3.Text = "4x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb24.Text = Val(a_tb21.Text) * Val(b_tb14.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                        c_tb34.Text = Val(a_tb31.Text) * Val(b_tb14.Text)
                        c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                        c_tb42.Text = Val(a_tb41.Text) * Val(b_tb12.Text)
                        c_tb43.Text = Val(a_tb41.Text) * Val(b_tb13.Text)
                        c_tb44.Text = Val(a_tb41.Text) * Val(b_tb14.Text)
                    End If
                ElseIf Label3.Text = "5x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb24.Text = Val(a_tb21.Text) * Val(b_tb14.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                        c_tb34.Text = Val(a_tb31.Text) * Val(b_tb14.Text)
                        c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                        c_tb42.Text = Val(a_tb41.Text) * Val(b_tb12.Text)
                        c_tb43.Text = Val(a_tb41.Text) * Val(b_tb13.Text)
                        c_tb44.Text = Val(a_tb41.Text) * Val(b_tb14.Text)
                        c_tb51.Text = Val(a_tb51.Text) * Val(b_tb11.Text)
                        c_tb52.Text = Val(a_tb51.Text) * Val(b_tb12.Text)
                        c_tb53.Text = Val(a_tb51.Text) * Val(b_tb13.Text)
                        c_tb54.Text = Val(a_tb51.Text) * Val(b_tb14.Text)
                    End If
                End If
            ElseIf Label4.Text = "1x5" Then
                If Label3.Text = "2x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                        c_tb15.Text = Val(a_tb11.Text) * Val(b_tb15.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb24.Text = Val(a_tb21.Text) * Val(b_tb14.Text)
                        c_tb25.Text = Val(a_tb21.Text) * Val(b_tb15.Text)
                    End If
                ElseIf Label3.Text = "3x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                        c_tb15.Text = Val(a_tb11.Text) * Val(b_tb15.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb24.Text = Val(a_tb21.Text) * Val(b_tb14.Text)
                        c_tb25.Text = Val(a_tb21.Text) * Val(b_tb15.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                        c_tb34.Text = Val(a_tb31.Text) * Val(b_tb14.Text)
                        c_tb35.Text = Val(a_tb31.Text) * Val(b_tb15.Text)
                    End If
                ElseIf Label3.Text = "4x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                        c_tb15.Text = Val(a_tb11.Text) * Val(b_tb15.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb24.Text = Val(a_tb21.Text) * Val(b_tb14.Text)
                        c_tb25.Text = Val(a_tb21.Text) * Val(b_tb15.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                        c_tb34.Text = Val(a_tb31.Text) * Val(b_tb14.Text)
                        c_tb35.Text = Val(a_tb31.Text) * Val(b_tb15.Text)
                        c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                        c_tb42.Text = Val(a_tb41.Text) * Val(b_tb12.Text)
                        c_tb43.Text = Val(a_tb41.Text) * Val(b_tb13.Text)
                        c_tb44.Text = Val(a_tb41.Text) * Val(b_tb14.Text)
                        c_tb45.Text = Val(a_tb41.Text) * Val(b_tb15.Text)
                    End If
                ElseIf Label3.Text = "5x1" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = Val(a_tb11.Text) * Val(b_tb11.Text)
                        c_tb12.Text = Val(a_tb11.Text) * Val(b_tb12.Text)
                        c_tb13.Text = Val(a_tb11.Text) * Val(b_tb13.Text)
                        c_tb14.Text = Val(a_tb11.Text) * Val(b_tb14.Text)
                        c_tb15.Text = Val(a_tb11.Text) * Val(b_tb15.Text)
                        c_tb21.Text = Val(a_tb21.Text) * Val(b_tb11.Text)
                        c_tb22.Text = Val(a_tb21.Text) * Val(b_tb12.Text)
                        c_tb23.Text = Val(a_tb21.Text) * Val(b_tb13.Text)
                        c_tb24.Text = Val(a_tb21.Text) * Val(b_tb14.Text)
                        c_tb25.Text = Val(a_tb21.Text) * Val(b_tb15.Text)
                        c_tb31.Text = Val(a_tb31.Text) * Val(b_tb11.Text)
                        c_tb32.Text = Val(a_tb31.Text) * Val(b_tb12.Text)
                        c_tb33.Text = Val(a_tb31.Text) * Val(b_tb13.Text)
                        c_tb34.Text = Val(a_tb31.Text) * Val(b_tb14.Text)
                        c_tb35.Text = Val(a_tb31.Text) * Val(b_tb15.Text)
                        c_tb41.Text = Val(a_tb41.Text) * Val(b_tb11.Text)
                        c_tb42.Text = Val(a_tb41.Text) * Val(b_tb12.Text)
                        c_tb43.Text = Val(a_tb41.Text) * Val(b_tb13.Text)
                        c_tb44.Text = Val(a_tb41.Text) * Val(b_tb14.Text)
                        c_tb45.Text = Val(a_tb41.Text) * Val(b_tb15.Text)
                        c_tb51.Text = Val(a_tb51.Text) * Val(b_tb11.Text)
                        c_tb52.Text = Val(a_tb51.Text) * Val(b_tb12.Text)
                        c_tb53.Text = Val(a_tb51.Text) * Val(b_tb13.Text)
                        c_tb54.Text = Val(a_tb51.Text) * Val(b_tb14.Text)
                        c_tb55.Text = Val(a_tb51.Text) * Val(b_tb15.Text)
                    End If
                End If
                '
            ElseIf Label4.Text = "2x2" Then
                If Label3.Text = "2x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                    End If
                ElseIf Label3.Text = "3x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                    End If
                ElseIf Label3.Text = "4x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text))
                    End If
                ElseIf Label3.Text = "5x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text))
                    End If
                End If
            ElseIf Label4.Text = "2x3" Then
                If Label3.Text = "2x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                    End If
                ElseIf Label3.Text = "3x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                    End If
                ElseIf Label3.Text = "4x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text))
                    End If
                ElseIf Label3.Text = "5x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text))
                    End If
                End If
            ElseIf Label4.Text = "2x4" Then
                If Label3.Text = "2x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text))
                    End If
                ElseIf Label3.Text = "3x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text))
                    End If
                ElseIf Label3.Text = "4x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text))
                    End If
                ElseIf Label3.Text = "5x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text))
                        c_tb54.Text = (Val(a_tb51.Text) * Val(b_tb14.Text)) + (Val(a_tb52.Text) * Val(b_tb24.Text))
                    End If
                End If
            ElseIf Label4.Text = "2x5" Then
                If Label3.Text = "2x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text))
                    End If
                ElseIf Label3.Text = "3x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text))
                    End If
                ElseIf Label3.Text = "4x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text))
                        c_tb45.Text = (Val(a_tb41.Text) * Val(b_tb15.Text)) + (Val(a_tb42.Text) * Val(b_tb25.Text))
                    End If
                ElseIf Label3.Text = "5x2" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text))
                        c_tb45.Text = (Val(a_tb41.Text) * Val(b_tb15.Text)) + (Val(a_tb42.Text) * Val(b_tb25.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text))
                        c_tb54.Text = (Val(a_tb51.Text) * Val(b_tb14.Text)) + (Val(a_tb52.Text) * Val(b_tb24.Text))
                        c_tb55.Text = (Val(a_tb51.Text) * Val(b_tb15.Text)) + (Val(a_tb52.Text) * Val(b_tb25.Text))
                    End If
                End If
            ElseIf Label4.Text = "3x2" Then
                If Label3.Text = "2x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                    End If
                ElseIf Label3.Text = "3x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                    End If
                ElseIf Label3.Text = "4x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text))
                    End If
                ElseIf Label3.Text = "5x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text))
                    End If
                End If
            ElseIf Label4.Text = "3x3" Then
                If Label3.Text = "2x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                    End If
                ElseIf Label3.Text = "3x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                    End If
                ElseIf Label3.Text = "4x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text))
                    End If
                ElseIf Label3.Text = "5x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text)) + (Val(a_tb53.Text) * Val(b_tb33.Text))
                    End If
                End If
            ElseIf Label4.Text = "3x4" Then
                If Label3.Text = "2x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text))
                    End If
                ElseIf Label3.Text = "3x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text))
                    End If
                ElseIf Label3.Text = "4x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text))
                    End If
                ElseIf Label3.Text = "5x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text)) + (Val(a_tb53.Text) * Val(b_tb33.Text))
                        c_tb54.Text = (Val(a_tb51.Text) * Val(b_tb14.Text)) + (Val(a_tb52.Text) * Val(b_tb24.Text)) + (Val(a_tb53.Text) * Val(b_tb34.Text))
                    End If
                End If
            ElseIf Label4.Text = "3x5" Then
                If Label3.Text = "2x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text)) + (Val(a_tb23.Text) * Val(b_tb35.Text))
                    End If
                ElseIf Label3.Text = "3x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text)) + (Val(a_tb23.Text) * Val(b_tb35.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text)) + (Val(a_tb33.Text) * Val(b_tb35.Text))
                    End If
                ElseIf Label3.Text = "4x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text)) + (Val(a_tb23.Text) * Val(b_tb35.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text)) + (Val(a_tb33.Text) * Val(b_tb35.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text))
                        c_tb45.Text = (Val(a_tb41.Text) * Val(b_tb15.Text)) + (Val(a_tb42.Text) * Val(b_tb25.Text)) + (Val(a_tb43.Text) * Val(b_tb35.Text))
                    End If
                ElseIf Label3.Text = "5x3" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text)) + (Val(a_tb23.Text) * Val(b_tb35.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text)) + (Val(a_tb33.Text) * Val(b_tb35.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text))
                        c_tb45.Text = (Val(a_tb41.Text) * Val(b_tb15.Text)) + (Val(a_tb42.Text) * Val(b_tb25.Text)) + (Val(a_tb43.Text) * Val(b_tb35.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text)) + (Val(a_tb53.Text) * Val(b_tb33.Text))
                        c_tb54.Text = (Val(a_tb51.Text) * Val(b_tb14.Text)) + (Val(a_tb52.Text) * Val(b_tb24.Text)) + (Val(a_tb53.Text) * Val(b_tb34.Text))
                        c_tb55.Text = (Val(a_tb51.Text) * Val(b_tb15.Text)) + (Val(a_tb52.Text) * Val(b_tb25.Text)) + (Val(a_tb53.Text) * Val(b_tb35.Text))
                    End If
                End If
            ElseIf Label4.Text = "4x2" Then
                If Label3.Text = "2x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                    End If
                ElseIf Label3.Text = "3x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                    End If
                ElseIf Label3.Text = "4x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text))
                    End If
                ElseIf Label3.Text = "5x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text)) + (Val(a_tb54.Text) * Val(b_tb41.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text)) + (Val(a_tb54.Text) * Val(b_tb42.Text))
                    End If
                End If
            ElseIf Label4.Text = "4x3" Then
                If Label3.Text = "2x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text))
                    End If
                ElseIf Label3.Text = "3x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text))
                    End If
                ElseIf Label3.Text = "4x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text))
                    End If
                ElseIf Label3.Text = "5x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text)) + (Val(a_tb54.Text) * Val(b_tb41.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text)) + (Val(a_tb54.Text) * Val(b_tb42.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text)) + (Val(a_tb53.Text) * Val(b_tb33.Text)) + (Val(a_tb54.Text) * Val(b_tb43.Text))
                    End If
                End If
            ElseIf Label4.Text = "4x4" Then
                If Label3.Text = "2x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text))
                    End If
                ElseIf Label3.Text = "3x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text))
                    End If
                ElseIf Label3.Text = "4x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text)) + (Val(a_tb44.Text) * Val(b_tb44.Text))
                    End If
                ElseIf Label3.Text = "5x4" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text)) + (Val(a_tb44.Text) * Val(b_tb44.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text)) + (Val(a_tb54.Text) * Val(b_tb41.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text)) + (Val(a_tb54.Text) * Val(b_tb42.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text)) + (Val(a_tb53.Text) * Val(b_tb33.Text)) + (Val(a_tb54.Text) * Val(b_tb43.Text))
                        c_tb54.Text = (Val(a_tb51.Text) * Val(b_tb14.Text)) + (Val(a_tb52.Text) * Val(b_tb24.Text)) + (Val(a_tb53.Text) * Val(b_tb34.Text)) + (Val(a_tb54.Text) * Val(b_tb44.Text))
                    End If
                End If
            ElseIf Label4.Text = "5x2" Then
                If Label3.Text = "2x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                    End If
                ElseIf Label3.Text = "3x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                    End If
                ElseIf Label3.Text = "4x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "" Or a_tb45.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text)) + (Val(a_tb45.Text) * Val(b_tb51.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text)) + (Val(a_tb45.Text) * Val(b_tb52.Text))
                    End If
                ElseIf Label3.Text = "5x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb34.Text = "" Or a_tb35.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb45.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "" Or a_tb45.Text = "" Or a_tb55.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text)) + (Val(a_tb45.Text) * Val(b_tb51.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text)) + (Val(a_tb45.Text) * Val(b_tb52.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text)) + (Val(a_tb54.Text) * Val(b_tb41.Text)) + (Val(a_tb55.Text) * Val(b_tb51.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text)) + (Val(a_tb54.Text) * Val(b_tb42.Text)) + (Val(a_tb55.Text) * Val(b_tb52.Text))
                    End If
                End If
            ElseIf Label4.Text = "5x3" Then
                If Label3.Text = "2x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                    End If
                ElseIf Label3.Text = "3x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                    End If
                ElseIf Label3.Text = "4x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "" Or a_tb45.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text)) + (Val(a_tb45.Text) * Val(b_tb51.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text)) + (Val(a_tb45.Text) * Val(b_tb52.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text)) + (Val(a_tb45.Text) * Val(b_tb53.Text))
                    End If
                ElseIf Label3.Text = "5x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "" Or a_tb45.Text = "" Or a_tb55.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text)) + (Val(a_tb45.Text) * Val(b_tb51.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text)) + (Val(a_tb45.Text) * Val(b_tb52.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text)) + (Val(a_tb45.Text) * Val(b_tb53.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text)) + (Val(a_tb54.Text) * Val(b_tb41.Text)) + (Val(a_tb55.Text) * Val(b_tb51.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text)) + (Val(a_tb54.Text) * Val(b_tb42.Text)) + (Val(a_tb55.Text) * Val(b_tb52.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text)) + (Val(a_tb53.Text) * Val(b_tb33.Text)) + (Val(a_tb54.Text) * Val(b_tb43.Text)) + (Val(a_tb55.Text) * Val(b_tb53.Text))
                    End If
                End If
            ElseIf Label4.Text = "5x4" Then
                If Label3.Text = "2x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text)) + (Val(a_tb25.Text) * Val(b_tb54.Text))
                    End If
                ElseIf Label3.Text = "3x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text)) + (Val(a_tb25.Text) * Val(b_tb54.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text)) + (Val(a_tb35.Text) * Val(b_tb54.Text))
                    End If
                ElseIf Label3.Text = "4x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "" Or a_tb45.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text)) + (Val(a_tb25.Text) * Val(b_tb54.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text)) + (Val(a_tb35.Text) * Val(b_tb54.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text)) + (Val(a_tb45.Text) * Val(b_tb51.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text)) + (Val(a_tb45.Text) * Val(b_tb52.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text)) + (Val(a_tb45.Text) * Val(b_tb53.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text)) + (Val(a_tb44.Text) * Val(b_tb44.Text)) + (Val(a_tb45.Text) * Val(b_tb54.Text))
                    End If
                ElseIf Label3.Text = "5x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "" Or a_tb45.Text = "" Or a_tb55.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text)) + (Val(a_tb25.Text) * Val(b_tb54.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text)) + (Val(a_tb35.Text) * Val(b_tb54.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text)) + (Val(a_tb45.Text) * Val(b_tb51.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text)) + (Val(a_tb45.Text) * Val(b_tb52.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text)) + (Val(a_tb45.Text) * Val(b_tb53.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text)) + (Val(a_tb44.Text) * Val(b_tb44.Text)) + (Val(a_tb45.Text) * Val(b_tb54.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text)) + (Val(a_tb54.Text) * Val(b_tb41.Text)) + (Val(a_tb55.Text) * Val(b_tb51.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text)) + (Val(a_tb54.Text) * Val(b_tb42.Text)) + (Val(a_tb55.Text) * Val(b_tb52.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text)) + (Val(a_tb53.Text) * Val(b_tb33.Text)) + (Val(a_tb54.Text) * Val(b_tb43.Text)) + (Val(a_tb55.Text) * Val(b_tb53.Text))
                        c_tb54.Text = (Val(a_tb51.Text) * Val(b_tb14.Text)) + (Val(a_tb52.Text) * Val(b_tb24.Text)) + (Val(a_tb53.Text) * Val(b_tb34.Text)) + (Val(a_tb54.Text) * Val(b_tb44.Text)) + (Val(a_tb55.Text) * Val(b_tb54.Text))
                    End If
                End If
            ElseIf Label4.Text = "5x5" Then
                If Label3.Text = "2x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "" Or b_tb55.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text)) + (Val(a_tb14.Text) * Val(b_tb45.Text)) + (Val(a_tb15.Text) * Val(b_tb55.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text)) + (Val(a_tb25.Text) * Val(b_tb54.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text)) + (Val(a_tb23.Text) * Val(b_tb35.Text)) + (Val(a_tb24.Text) * Val(b_tb45.Text)) + (Val(a_tb25.Text) * Val(b_tb55.Text))
                    End If
                ElseIf Label3.Text = "3x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "" Or b_tb55.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text)) + (Val(a_tb14.Text) * Val(b_tb45.Text)) + (Val(a_tb15.Text) * Val(b_tb55.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text)) + (Val(a_tb25.Text) * Val(b_tb54.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text)) + (Val(a_tb23.Text) * Val(b_tb35.Text)) + (Val(a_tb24.Text) * Val(b_tb45.Text)) + (Val(a_tb25.Text) * Val(b_tb55.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text)) + (Val(a_tb35.Text) * Val(b_tb54.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text)) + (Val(a_tb33.Text) * Val(b_tb35.Text)) + (Val(a_tb34.Text) * Val(b_tb45.Text)) + (Val(a_tb35.Text) * Val(b_tb55.Text))
                    End If
                ElseIf Label3.Text = "4x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "" Or a_tb45.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "" Or b_tb55.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text)) + (Val(a_tb14.Text) * Val(b_tb45.Text)) + (Val(a_tb15.Text) * Val(b_tb55.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text)) + (Val(a_tb25.Text) * Val(b_tb54.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text)) + (Val(a_tb23.Text) * Val(b_tb35.Text)) + (Val(a_tb24.Text) * Val(b_tb45.Text)) + (Val(a_tb25.Text) * Val(b_tb55.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text)) + (Val(a_tb35.Text) * Val(b_tb54.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text)) + (Val(a_tb33.Text) * Val(b_tb35.Text)) + (Val(a_tb34.Text) * Val(b_tb45.Text)) + (Val(a_tb35.Text) * Val(b_tb55.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text)) + (Val(a_tb45.Text) * Val(b_tb51.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text)) + (Val(a_tb45.Text) * Val(b_tb52.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text)) + (Val(a_tb45.Text) * Val(b_tb53.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text)) + (Val(a_tb44.Text) * Val(b_tb44.Text)) + (Val(a_tb45.Text) * Val(b_tb54.Text))
                        c_tb45.Text = (Val(a_tb41.Text) * Val(b_tb15.Text)) + (Val(a_tb42.Text) * Val(b_tb25.Text)) + (Val(a_tb43.Text) * Val(b_tb35.Text)) + (Val(a_tb44.Text) * Val(b_tb45.Text)) + (Val(a_tb45.Text) * Val(b_tb55.Text))
                    End If
                ElseIf Label3.Text = "5x5" Then
                    If (a_tb11.Text = "" Or a_tb21.Text = "" Or a_tb31.Text = "" Or a_tb41.Text = "" Or a_tb51.Text = "") Or (a_tb12.Text = "" Or a_tb22.Text = "" Or a_tb32.Text = "" Or a_tb42.Text = "" Or a_tb52.Text = "") Or (a_tb13.Text = "" Or a_tb23.Text = "" Or a_tb33.Text = "" Or a_tb43.Text = "" Or a_tb53.Text = "") Or (a_tb14.Text = "" Or a_tb24.Text = "" Or a_tb34.Text = "" Or a_tb44.Text = "" Or a_tb54.Text = "") Or (a_tb15.Text = "" Or a_tb25.Text = "" Or a_tb35.Text = "" Or a_tb45.Text = "" Or a_tb55.Text = "") Or (b_tb11.Text = "" Or b_tb12.Text = "" Or b_tb13.Text = "" Or b_tb14.Text = "" Or b_tb15.Text = "") Or (b_tb21.Text = "" Or b_tb22.Text = "" Or b_tb23.Text = "" Or b_tb24.Text = "" Or b_tb25.Text = "") Or (b_tb31.Text = "" Or b_tb32.Text = "" Or b_tb33.Text = "" Or b_tb34.Text = "" Or b_tb35.Text = "") Or (b_tb41.Text = "" Or b_tb42.Text = "" Or b_tb43.Text = "" Or b_tb44.Text = "" Or b_tb45.Text = "") Or (b_tb51.Text = "" Or b_tb52.Text = "" Or b_tb53.Text = "" Or b_tb54.Text = "" Or b_tb55.Text = "") Then
                        MsgBox("INVALID! Lack of input", MsgBoxStyle.Exclamation)
                    Else
                        c_tb11.Text = (Val(a_tb11.Text) * Val(b_tb11.Text)) + (Val(a_tb12.Text) * Val(b_tb21.Text)) + (Val(a_tb13.Text) * Val(b_tb31.Text)) + (Val(a_tb14.Text) * Val(b_tb41.Text)) + (Val(a_tb15.Text) * Val(b_tb51.Text))
                        c_tb12.Text = (Val(a_tb11.Text) * Val(b_tb12.Text)) + (Val(a_tb12.Text) * Val(b_tb22.Text)) + (Val(a_tb13.Text) * Val(b_tb32.Text)) + (Val(a_tb14.Text) * Val(b_tb42.Text)) + (Val(a_tb15.Text) * Val(b_tb52.Text))
                        c_tb13.Text = (Val(a_tb11.Text) * Val(b_tb13.Text)) + (Val(a_tb12.Text) * Val(b_tb23.Text)) + (Val(a_tb13.Text) * Val(b_tb33.Text)) + (Val(a_tb14.Text) * Val(b_tb43.Text)) + (Val(a_tb15.Text) * Val(b_tb53.Text))
                        c_tb14.Text = (Val(a_tb11.Text) * Val(b_tb14.Text)) + (Val(a_tb12.Text) * Val(b_tb24.Text)) + (Val(a_tb13.Text) * Val(b_tb34.Text)) + (Val(a_tb14.Text) * Val(b_tb44.Text)) + (Val(a_tb15.Text) * Val(b_tb54.Text))
                        c_tb15.Text = (Val(a_tb11.Text) * Val(b_tb15.Text)) + (Val(a_tb12.Text) * Val(b_tb25.Text)) + (Val(a_tb13.Text) * Val(b_tb35.Text)) + (Val(a_tb14.Text) * Val(b_tb45.Text)) + (Val(a_tb15.Text) * Val(b_tb55.Text))
                        c_tb21.Text = (Val(a_tb21.Text) * Val(b_tb11.Text)) + (Val(a_tb22.Text) * Val(b_tb21.Text)) + (Val(a_tb23.Text) * Val(b_tb31.Text)) + (Val(a_tb24.Text) * Val(b_tb41.Text)) + (Val(a_tb25.Text) * Val(b_tb51.Text))
                        c_tb22.Text = (Val(a_tb21.Text) * Val(b_tb12.Text)) + (Val(a_tb22.Text) * Val(b_tb22.Text)) + (Val(a_tb23.Text) * Val(b_tb32.Text)) + (Val(a_tb24.Text) * Val(b_tb42.Text)) + (Val(a_tb25.Text) * Val(b_tb52.Text))
                        c_tb23.Text = (Val(a_tb21.Text) * Val(b_tb13.Text)) + (Val(a_tb22.Text) * Val(b_tb23.Text)) + (Val(a_tb23.Text) * Val(b_tb33.Text)) + (Val(a_tb24.Text) * Val(b_tb43.Text)) + (Val(a_tb25.Text) * Val(b_tb53.Text))
                        c_tb24.Text = (Val(a_tb21.Text) * Val(b_tb14.Text)) + (Val(a_tb22.Text) * Val(b_tb24.Text)) + (Val(a_tb23.Text) * Val(b_tb34.Text)) + (Val(a_tb24.Text) * Val(b_tb44.Text)) + (Val(a_tb25.Text) * Val(b_tb54.Text))
                        c_tb25.Text = (Val(a_tb21.Text) * Val(b_tb15.Text)) + (Val(a_tb22.Text) * Val(b_tb25.Text)) + (Val(a_tb23.Text) * Val(b_tb35.Text)) + (Val(a_tb24.Text) * Val(b_tb45.Text)) + (Val(a_tb25.Text) * Val(b_tb55.Text))
                        c_tb31.Text = (Val(a_tb31.Text) * Val(b_tb11.Text)) + (Val(a_tb32.Text) * Val(b_tb21.Text)) + (Val(a_tb33.Text) * Val(b_tb31.Text)) + (Val(a_tb34.Text) * Val(b_tb41.Text)) + (Val(a_tb35.Text) * Val(b_tb51.Text))
                        c_tb32.Text = (Val(a_tb31.Text) * Val(b_tb12.Text)) + (Val(a_tb32.Text) * Val(b_tb22.Text)) + (Val(a_tb33.Text) * Val(b_tb32.Text)) + (Val(a_tb34.Text) * Val(b_tb42.Text)) + (Val(a_tb35.Text) * Val(b_tb52.Text))
                        c_tb33.Text = (Val(a_tb31.Text) * Val(b_tb13.Text)) + (Val(a_tb32.Text) * Val(b_tb23.Text)) + (Val(a_tb33.Text) * Val(b_tb33.Text)) + (Val(a_tb34.Text) * Val(b_tb43.Text)) + (Val(a_tb35.Text) * Val(b_tb53.Text))
                        c_tb34.Text = (Val(a_tb31.Text) * Val(b_tb14.Text)) + (Val(a_tb32.Text) * Val(b_tb24.Text)) + (Val(a_tb33.Text) * Val(b_tb34.Text)) + (Val(a_tb34.Text) * Val(b_tb44.Text)) + (Val(a_tb35.Text) * Val(b_tb54.Text))
                        c_tb35.Text = (Val(a_tb31.Text) * Val(b_tb15.Text)) + (Val(a_tb32.Text) * Val(b_tb25.Text)) + (Val(a_tb33.Text) * Val(b_tb35.Text)) + (Val(a_tb34.Text) * Val(b_tb45.Text)) + (Val(a_tb35.Text) * Val(b_tb55.Text))
                        c_tb41.Text = (Val(a_tb41.Text) * Val(b_tb11.Text)) + (Val(a_tb42.Text) * Val(b_tb21.Text)) + (Val(a_tb43.Text) * Val(b_tb31.Text)) + (Val(a_tb44.Text) * Val(b_tb41.Text)) + (Val(a_tb45.Text) * Val(b_tb51.Text))
                        c_tb42.Text = (Val(a_tb41.Text) * Val(b_tb12.Text)) + (Val(a_tb42.Text) * Val(b_tb22.Text)) + (Val(a_tb43.Text) * Val(b_tb32.Text)) + (Val(a_tb44.Text) * Val(b_tb42.Text)) + (Val(a_tb45.Text) * Val(b_tb52.Text))
                        c_tb43.Text = (Val(a_tb41.Text) * Val(b_tb13.Text)) + (Val(a_tb42.Text) * Val(b_tb23.Text)) + (Val(a_tb43.Text) * Val(b_tb33.Text)) + (Val(a_tb44.Text) * Val(b_tb43.Text)) + (Val(a_tb45.Text) * Val(b_tb53.Text))
                        c_tb44.Text = (Val(a_tb41.Text) * Val(b_tb14.Text)) + (Val(a_tb42.Text) * Val(b_tb24.Text)) + (Val(a_tb43.Text) * Val(b_tb34.Text)) + (Val(a_tb44.Text) * Val(b_tb44.Text)) + (Val(a_tb45.Text) * Val(b_tb54.Text))
                        c_tb45.Text = (Val(a_tb41.Text) * Val(b_tb15.Text)) + (Val(a_tb42.Text) * Val(b_tb25.Text)) + (Val(a_tb43.Text) * Val(b_tb35.Text)) + (Val(a_tb44.Text) * Val(b_tb45.Text)) + (Val(a_tb45.Text) * Val(b_tb55.Text))
                        c_tb51.Text = (Val(a_tb51.Text) * Val(b_tb11.Text)) + (Val(a_tb52.Text) * Val(b_tb21.Text)) + (Val(a_tb53.Text) * Val(b_tb31.Text)) + (Val(a_tb54.Text) * Val(b_tb41.Text)) + (Val(a_tb55.Text) * Val(b_tb51.Text))
                        c_tb52.Text = (Val(a_tb51.Text) * Val(b_tb12.Text)) + (Val(a_tb52.Text) * Val(b_tb22.Text)) + (Val(a_tb53.Text) * Val(b_tb32.Text)) + (Val(a_tb54.Text) * Val(b_tb42.Text)) + (Val(a_tb55.Text) * Val(b_tb52.Text))
                        c_tb53.Text = (Val(a_tb51.Text) * Val(b_tb13.Text)) + (Val(a_tb52.Text) * Val(b_tb23.Text)) + (Val(a_tb53.Text) * Val(b_tb33.Text)) + (Val(a_tb54.Text) * Val(b_tb43.Text)) + (Val(a_tb55.Text) * Val(b_tb53.Text))
                        c_tb54.Text = (Val(a_tb51.Text) * Val(b_tb14.Text)) + (Val(a_tb52.Text) * Val(b_tb24.Text)) + (Val(a_tb53.Text) * Val(b_tb34.Text)) + (Val(a_tb54.Text) * Val(b_tb44.Text)) + (Val(a_tb55.Text) * Val(b_tb54.Text))
                        c_tb55.Text = (Val(a_tb51.Text) * Val(b_tb15.Text)) + (Val(a_tb52.Text) * Val(b_tb25.Text)) + (Val(a_tb53.Text) * Val(b_tb35.Text)) + (Val(a_tb54.Text) * Val(b_tb45.Text)) + (Val(a_tb55.Text) * Val(b_tb55.Text))
                    End If
                End If
            End If
        End If
    End Sub
#End Region
#End Region
#Region "MainLoad"
    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label3.Text = "1x1"
        Label4.Text = "1x1"
        Label6.Text = "1"
        Label7.Text = "1"
        Label9.Text = "1"
    End Sub
#End Region
End Class
