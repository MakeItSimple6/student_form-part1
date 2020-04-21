Imports System.IO
Imports System.Data.OleDb

Public Class add_student
    Dim bytImage() As Byte
    Private abyt As Byte()
    Dim con As New OleDbConnection
    Dim constr As String
    Dim command As New OleDbCommand
    Dim Imagepath As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Image Upload'
        Try
            With OpenFileDialog1
                .Filter = "Image Files(*.BMP;*.JPG;*.GIF;*.PNG)|*.BMP;*.JPG;*.GIF;*.PNG|All Files (*.*)|*.*"
                .FileName = ""
                .Title = "Choose a Picture..."
                .AddExtension = True
                .FilterIndex = 1
                .Multiselect = False
                .ValidateNames = True
                .RestoreDirectory = True

                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Dim fs As New FileStream(.FileName, FileMode.Open)
                    Dim br As New BinaryReader(fs)
                    abyt = br.ReadBytes(CInt(fs.Length))
                    br.Close()
                    Dim ms As New MemoryStream(abyt)
                    PictureBox1.Image = Image.FromStream(ms)
                End If

            End With
        Catch ex As Exception
            MessageBox.Show("Error : " + ex.Message.ToString())
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Clear'
        PictureBox1.Image = Nothing
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Close
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Reset
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
        DateTimePicker1.Value = "01-01-1990"
        PictureBox1.Image = Nothing
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Add'
        Try
            With con
                .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Kamar\Desktop\Student Management System\student.accdb"
                .Open()

                With command
                    .Connection = con
                    .CommandText = "Insert into student([ID],[student_name],[dob],[gender],[phone_no],[address],[image])
                                    Values(@id,@name,@dob,@gender,@phone,@address,@image)"
                End With

                command.Parameters.AddWithValue("@id", TextBox1.Text)
                command.Parameters.AddWithValue("@name", TextBox2.Text)
                command.Parameters.AddWithValue("@dob", DateTimePicker1.Value.Date)
                command.Parameters.AddWithValue("@gender", ComboBox1.Text)
                command.Parameters.AddWithValue("@phone", TextBox3.Text)
                command.Parameters.AddWithValue("@address", TextBox4.Text)
                command.Parameters.AddWithValue("@image", abyt)

                command.ExecuteNonQuery()
                MsgBox("Student Details Saved Successfully!")
                con.Close()
                command.Dispose()
            End With
        Catch ex As Exception
            MessageBox.Show("Error : " + ex.Message.ToString())
        End Try
    End Sub
End Class