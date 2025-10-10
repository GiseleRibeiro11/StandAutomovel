Imports System.Drawing.Printing
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports Microsoft.Data.SqlClient

Public Class Form1
    Dim ligacao As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\gisel\OneDrive\Documentos\NST-PROG21\3934- VisualBasic.NET\Tarefa2\bin\Debug\net9.0-windows\bdaguiar.mdf;Integrated Security=True;Connect Timeout=30")
    Dim comando As New SqlCommand
    Dim ler As SqlDataReader
    Private Sub btn_adicionar_Click(sender As Object, e As EventArgs) Handles btn_adicionar.Click
        If (txt_marca.Text.Length = 0) Then
            MessageBox.Show("Marca é um campo obrigatório",
                            "Marca", MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            txt_marca.Focus()
            Exit Sub
        End If
        If (txt_modelo.Text.Length = 0) Then
            MessageBox.Show("Modelo é um campo obrigatório",
                            "Modelo", MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            txt_modelo.Focus()
            Exit Sub
        End If
        If (btn_adicionar.Text = "Atualizar") Then
            list_marca.Items(list_marca.SelectedIndex) = txt_marca.Text
            list_modelo.Items(list_modelo.SelectedIndex) = txt_modelo.Text
            list_matricula.Items(list_matricula.SelectedIndex) = mtxt_matricula.Text
            list_kms.Items(list_kms.SelectedIndex) = mtxt_kms.Text
            limpar()
            btn_adicionar.Text = "Adicionar"
            Exit Sub
        End If
        list_marca.Items.Add(txt_marca.Text)
        list_modelo.Items.Add(txt_modelo.Text)
        list_matricula.Items.Add(mtxt_matricula.Text)
        list_kms.Items.Add(mtxt_kms.Text)
        ' Limpar os campos
    End Sub

    Sub limpar()
        txt_marca.Clear()
        txt_modelo.Clear()
        mtxt_matricula.Clear()
        mtxt_kms.Clear()
    End Sub

    Private Sub btn_sair_Click(sender As Object, e As EventArgs) Handles btn_sair.Click
        limpar()
    End Sub

    Private Sub list_marca_SelectedIndexChanged(sender As Object, e As EventArgs) Handles list_marca.SelectedIndexChanged
        list_modelo.SelectedIndex = list_marca.SelectedIndex()
        list_matricula.SelectedIndex = list_marca.SelectedIndex()
        list_kms.SelectedIndex = list_marca.SelectedIndex()
    End Sub

    Private Sub list_modelo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles list_modelo.SelectedIndexChanged
        list_marca.SelectedIndex = list_modelo.SelectedIndex()
        list_matricula.SelectedIndex = list_modelo.SelectedIndex()
        list_kms.SelectedIndex = list_modelo.SelectedIndex()
    End Sub

    Private Sub list_matricula_SelectedIndexChanged(sender As Object, e As EventArgs) Handles list_matricula.SelectedIndexChanged
        list_modelo.SelectedIndex = list_matricula.SelectedIndex()
        list_marca.SelectedIndex = list_matricula.SelectedIndex()
        list_kms.SelectedIndex = list_matricula.SelectedIndex()
    End Sub

    Private Sub list_kms_SelectedIndexChanged(sender As Object, e As EventArgs) Handles list_kms.SelectedIndexChanged
        list_marca.SelectedIndex = list_kms.SelectedIndex()
        list_modelo.SelectedIndex = list_kms.SelectedIndex()
        list_matricula.SelectedIndex = list_kms.SelectedIndex()
    End Sub

    Private Sub btn_eliminar_Click(sender As Object, e As EventArgs) Handles btn_eliminar.Click
        If (list_marca.SelectedIndex < 0) Then
            MessageBox.Show("Selecione um elemento para eliminar",
                        "Eliminar registo", MessageBoxButtons.OK,
                         MessageBoxIcon.Error)
            Exit Sub
        Else
            Dim indice As Integer = list_marca.SelectedIndex
            list_marca.Items.RemoveAt(indice)
            list_modelo.Items.RemoveAt(indice)
            list_matricula.Items.RemoveAt(indice)
            list_kms.Items.RemoveAt(indice)
        End If
    End Sub

    Private Sub btn_sair2_Click(sender As Object, e As EventArgs) Handles btn_sair2.Click
        End
    End Sub

    Private Sub btn_alterar_Click(sender As Object, e As EventArgs) Handles btn_alterar.Click
        If (list_marca.SelectedIndex < 0) Then
            MessageBox.Show("Selecione um elemento para alterar",
                        "Alterar registo", MessageBoxButtons.OK,
                         MessageBoxIcon.Error)
        Else
            txt_marca.Text = list_marca.SelectedItem.ToString()
            txt_modelo.Text = list_modelo.SelectedItem.ToString()
            mtxt_matricula.Text = list_matricula.SelectedItem.ToString()
            mtxt_kms.Text = list_kms.SelectedItem.ToString()
            btn_adicionar.Text = "Atualizar"
        End If
    End Sub

    Private Sub btn_imprimir_Click(sender As Object, e As EventArgs) Handles btn_imprimir.Click
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim fonte1 As New Font("Verdana", 14)
        Dim fonte2 As New Font("Verdana", 10)

        e.Graphics.DrawString("Stand Aguiar, Lda", fonte1, Brushes.Blue, 50, 50)
        e.Graphics.DrawString("Listagem de viaturas", fonte1, Brushes.Blue, 50, 70)
        e.Graphics.DrawString("Marca", fonte2, Brushes.Black, 50, 150)
        e.Graphics.DrawString("Modelo", fonte2, Brushes.Black, 250, 150)
        e.Graphics.DrawString("Matrícula", fonte2, Brushes.Black, 450, 150)
        e.Graphics.DrawString("Kms", fonte2, Brushes.Black, 650, 150)
        Dim linha As Integer = 200
        For i = 0 To list_marca.Items.Count - 1
            e.Graphics.DrawString(list_marca.Items(i).ToString(), fonte2, Brushes.Black, 50, linha)
            e.Graphics.DrawString(list_modelo.Items(i).ToString(), fonte2, Brushes.Black, 250, linha)
            e.Graphics.DrawString(list_matricula.Items(i).ToString(), fonte2, Brushes.Black, 450, linha)
            e.Graphics.DrawString(list_kms.Items(i).ToString(), fonte2, Brushes.Black, 600, linha)
            linha += 20
        Next
        e.Graphics.DrawString("Fim da listagem ", fonte2, Brushes.Black, 50, linha)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ligacao.Open()
            MsgBox("Ligação à base de dados efetuada com sucesso", MsgBoxStyle.Information, "Atenção!")
        Catch ex As Exception
            MsgBox("Erro ao ligar à base de dados", MsgBoxStyle.Critical, "Atenção!")
        End Try
    End Sub

    Private Sub btn_exportar_Click(sender As Object, e As EventArgs) Handles btn_exportar.Click
        If (list_marca.Items.Count = 0) Then
            MsgBox("Não existem dados para exportar", MsgBoxStyle.Critical, "Atenção!")
        Else
            For i = 0 To list_marca.Items.Count - 1
                comando = New SqlCommand("INSERT INTO viaturas (marca, modelo, matricula, kms)
                                             VALUES (@marca, @modelo, @matricula, @kms)", ligacao)
                comando.Parameters.AddWithValue("@marca", list_marca.Items(i).ToString())
                comando.Parameters.AddWithValue("@modelo", list_modelo.Items(i).ToString())
                comando.Parameters.AddWithValue("@matricula", list_matricula.Items(i).ToString())
                comando.Parameters.AddWithValue("@kms", list_kms.Items(i).ToString())
                comando.ExecuteNonQuery()
            Next
        End If
        MsgBox("Exportação efetuada", MsgBoxStyle.Information, "Atenção!")
    End Sub
End Class

